using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry; // 包含 Point2d, Point3d
using OpenCvSharp; // 包含 Point, Point2d, Point3d 等
using OpenCvSharp.Extensions;
using Plugin_ContourMaster.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Reflection;
using System.Runtime.InteropServices;
// 定义别名以解决冲突
using AcadPoint2d = Autodesk.AutoCAD.Geometry.Point2d;
using AcadPoint3d = Autodesk.AutoCAD.Geometry.Point3d;
using CvPoint = OpenCvSharp.Point;

namespace Plugin_ContourMaster.Services
{

    public class ContourEngine
    {
        private readonly ContourSettings _settings;

        public ContourEngine(ContourSettings settings) => _settings = settings;

        // 1. 提取实体总范围
        private Extents3d GetSelectionExtents(SelectionSet ss)
        {
            Extents3d totalExt = new Extents3d();
            using (Transaction tr = Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject so in ss)
                {
                    Entity ent = (Entity)tr.GetObject(so.ObjectId, OpenMode.ForRead);
                    try { totalExt.AddExtents(ent.GeometricExtents); } catch { }
                }
                tr.Commit();
            }
            return totalExt;
        }

        // 2. 主处理方法
        public void ProcessContour()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                PromptSelectionResult psr = ed.GetSelection(new PromptSelectionOptions { MessageForAdding = "\n请选择要生成轮廓的元素: " });
                if (psr.Status != PromptStatus.OK) return;

                Extents3d totalExt = GetSelectionExtents(psr.Value);

                // 提高目标尺寸以增加初始精度
                int targetSize = 4000;
                double worldW = totalExt.MaxPoint.X - totalExt.MinPoint.X;
                double worldH = totalExt.MaxPoint.Y - totalExt.MinPoint.Y;

                if (worldW <= 0 || worldH <= 0) return;
                double scale = targetSize / Math.Max(worldW, worldH);

                using (Bitmap bmp = RasterizeSelection(doc, psr.Value, totalExt, scale))
                {
                    if (bmp == null) return;

                    // 使用 OpenCV 提取闭合区域轮廓
                    List<List<AcadPoint2d>> cadContours = ExtractCadContoursWithOpenCv(bmp, totalExt, scale);

                    if (cadContours.Count == 0)
                    {
                        ed.WriteMessage("\n[未发现轮廓] 请确保选中的线能够构成闭合区域，且不是“块”。");
                        return;
                    }

                    // 执行绘制并进行矢量吸附回正
                    DrawInCad(doc, cadContours, psr.Value, scale);

                    ed.WriteMessage($"\n[ContourMaster] 轮廓生成完毕，成功识别 {cadContours.Count} 个闭合区域。");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[异常] {ex.Message}");
            }
        }
        /// <summary>
        /// 将 OpenCV 提取的点坐标吸附到最近的原始线段上
        /// </summary>
        private AcadPoint2d SnapPointToOriginalCurves(AcadPoint2d pt, SelectionSet originalSs, Transaction tr, double pixelSize)
        {
            AcadPoint3d pt3d = new AcadPoint3d(pt.X, pt.Y, 0);
            AcadPoint3d bestPt = pt3d;

            // 吸附阈值设为 1.5 像素宽度
            double threshold = pixelSize * 1.5;
            double minDistance = threshold;

            foreach (SelectedObject so in originalSs)
            {
                Curve curve = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                if (curve == null) continue;

                try
                {
                    // 快速范围判定，提升处理 400+ 元素的效率
                    Extents3d ext = curve.GeometricExtents;
                    if (pt.X < ext.MinPoint.X - threshold || pt.X > ext.MaxPoint.X + threshold ||
                        pt.Y < ext.MinPoint.Y - threshold || pt.Y > ext.MaxPoint.Y + threshold)
                    {
                        continue;
                    }

                    // 获取原始曲线上最近的点坐标
                    AcadPoint3d closest = curve.GetClosestPointTo(pt3d, false);
                    double dist = pt3d.DistanceTo(closest);

                    if (dist < minDistance)
                    {
                        minDistance = dist;
                        bestPt = closest;
                    }
                }
                catch { continue; }
            }
            return new AcadPoint2d(bestPt.X, bestPt.Y);
        }
        // 3. 渲染位图
        // 修改后的 RasterizeSelection 方法
        private Bitmap RasterizeSelection(Document doc, SelectionSet ss, Extents3d ext, double scale)
        {
            int bmpW = (int)((ext.MaxPoint.X - ext.MinPoint.X) * scale) + 100;
            int bmpH = (int)((ext.MaxPoint.Y - ext.MinPoint.Y) * scale) + 100;

            // 防止创建过大的位图导致内存崩溃
            bmpW = Math.Min(bmpW, 8000);
            bmpH = Math.Min(bmpH, 8000);

            Bitmap bmp = new Bitmap(bmpW, bmpH);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(System.Drawing.Color.Black);
                // 开启抗锯齿，有助于 OpenCV 提取更平滑的边缘
                g.SmoothingMode = SmoothingMode.AntiAlias;

                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    using (System.Drawing.Pen thinPen = new System.Drawing.Pen(System.Drawing.Color.White, 1.2f))
                    {
                        foreach (SelectedObject so in ss)
                        {
                            Curve curve = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                            if (curve == null) continue;

                            // 修复：过滤长度几乎为0的无效线段，防止 eInvalidInput
                            double startParam = curve.StartParam;
                            double endParam = curve.EndParam;
                            if (Math.Abs(endParam - startParam) < 1e-7) continue;

                            int segments = 200;
                            List<PointF> pts = new List<PointF>();
                            double step = (endParam - startParam) / segments;

                            for (int i = 0; i <= segments; i++)
                            {
                                // 修复：强制约束参数范围，防止浮点数微弱越界导致的 eInvalidInput
                                double t = startParam + (step * i);
                                if (t < startParam) t = startParam;
                                if (t > endParam) t = endParam;

                                try
                                {
                                    AcadPoint3d p = curve.GetPointAtParameter(t);
                                    float px = (float)((p.X - ext.MinPoint.X) * scale) + 50f;
                                    float py = (float)((ext.MaxPoint.Y - p.Y) * scale) + 50f;
                                    pts.Add(new PointF(px, py));
                                }
                                catch { continue; }
                            }

                            if (pts.Count > 1)
                            {
                                g.DrawLines(thinPen, pts.ToArray());
                            }
                        }
                    }
                    tr.Commit();
                }
            }
            return bmp;
        }

        // 4. OpenCV 提取轮廓 (显式指定点类型解决冲突)
        private List<List<AcadPoint2d>> ExtractCadContoursWithOpenCv(Bitmap bmp, Extents3d ext, double scale)
        {
            List<List<AcadPoint2d>> results = new List<List<AcadPoint2d>>();

            using (Mat src = bmp.ToMat())
            using (Mat gray = new Mat())
            {
                Cv2.CvtColor(src, gray, ColorConversionCodes.BGRA2GRAY);

                using (Mat binary = new Mat())
                {
                    Cv2.Threshold(gray, binary, _settings.Threshold, 255, ThresholdTypes.Binary);

                    CvPoint[][] contours;
                    HierarchyIndex[] hierarchy;

                    // 使用 Tree 模式提取所有层级关系
                    Cv2.FindContours(binary, out contours, out hierarchy, RetrievalModes.Tree, ContourApproximationModes.ApproxSimple);

                    for (int i = 0; i < contours.Length; i++)
                    {
                        // 核心逻辑：只有拥有父级轮廓的才是“闭合孔洞”区域
                        // Parent == -1 的是外包络线，直接跳过
                        if (hierarchy[i].Parent == -1) continue;

                        var contour = contours[i];
                        if (contour.Length < 4) continue;

                        // 过滤面积过小的杂质
                        if (Cv2.ContourArea(contour) < 50) continue;

                        double epsilon = 1.0 + (_settings.SmoothLevel * 0.2);
                        CvPoint[] approx = Cv2.ApproxPolyDP(contour, epsilon, true);

                        List<AcadPoint2d> cadPath = new List<AcadPoint2d>();
                        foreach (var p in approx)
                        {
                            double wx = (p.X + 0.5 - 50.0) / scale + ext.MinPoint.X;
                            double wy = ext.MaxPoint.Y - (p.Y + 0.5 - 50.0) / scale;
                            cadPath.Add(new AcadPoint2d(wx, wy));
                        }

                        if (_settings.SmoothLevel > 3)
                        {
                            cadPath = SmoothCadPoints(cadPath);
                        }

                        results.Add(cadPath);
                    }
                }
            }
            return results;
        }

        // 5. 平滑算法 (使用别名 AcadPoint2d)
        private List<AcadPoint2d> SmoothCadPoints(List<AcadPoint2d> pts)
        {
            if (pts.Count < 5) return pts;
            List<AcadPoint2d> smoothed = new List<AcadPoint2d>(pts.Count) { pts[0] };
            for (int i = 1; i < pts.Count - 1; i++)
            {
                double sx = pts[i - 1].X * 0.15 + pts[i].X * 0.7 + pts[i + 1].X * 0.15;
                double sy = pts[i - 1].Y * 0.15 + pts[i].Y * 0.7 + pts[i + 1].Y * 0.15;
                smoothed.Add(new AcadPoint2d(sx, sy));
            }
            smoothed.Add(pts[pts.Count - 1]);
            return smoothed;
        }

        // 6. 绘制方法 (使用别名 AcadPoint2d)
        private void DrawInCad(Document doc, List<List<AcadPoint2d>> contours, SelectionSet originalLines, double scale)
        {
            double pixelSize = 1.0 / scale;

            using (var loc = doc.LockDocument())
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(doc.Database), OpenMode.ForWrite);

                string layerName = _settings.LayerName ?? "LK_XS";
                CheckAndCreateLayer(doc.Database, tr, layerName);

                foreach (var pts in contours)
                {
                    if (pts.Count < 2) continue;

                    Polyline pl = new Polyline();
                    for (int i = 0; i < pts.Count; i++)
                    {
                        // 执行吸附校正
                        AcadPoint2d snappedPt = SnapPointToOriginalCurves(pts[i], originalLines, tr, pixelSize);
                        pl.AddVertexAt(i, snappedPt, 0, 0, 0);
                    }

                    pl.Closed = true;
                    pl.Layer = layerName;
                    pl.ColorIndex = 256; // ByLayer

                    btr.AppendEntity(pl);
                    tr.AddNewlyCreatedDBObject(pl, true);
                }
                tr.Commit();
            }
        }

        private void CheckAndCreateLayer(Database db, Transaction tr, string layerName)
        {
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                LayerTableRecord ltr = new LayerTableRecord { Name = layerName };
                // 显式引用 AutoCAD 颜色命名空间
                ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2);
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }
    }
}