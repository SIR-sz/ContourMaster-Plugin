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

                int targetSize = 3000;
                double worldW = totalExt.MaxPoint.X - totalExt.MinPoint.X;
                double worldH = totalExt.MaxPoint.Y - totalExt.MinPoint.Y;
                double scale = targetSize / Math.Max(worldW, worldH);

                using (Bitmap bmp = RasterizeSelection(doc, psr.Value, totalExt, scale))
                {
                    if (bmp == null) return;

                    // 使用 OpenCV 提取轮廓
                    List<List<AcadPoint2d>> cadContours = ExtractCadContoursWithOpenCv(bmp, totalExt, scale);

                    DrawInCad(doc, cadContours);
                }

                ed.WriteMessage("\n[ContourMaster] 轮廓生成完毕。已修正命名空间冲突。");
            }
            catch (System.Exception ex) { ed.WriteMessage($"\n[异常] {ex.Message}"); }
        }

        // 3. 渲染位图
        // 修改后的 RasterizeSelection 方法
        private Bitmap RasterizeSelection(Document doc, SelectionSet ss, Extents3d ext, double scale)
        {
            // 增加：防止除以 0 导致的长宽异常
            double worldW = ext.MaxPoint.X - ext.MinPoint.X;
            double worldH = ext.MaxPoint.Y - ext.MinPoint.Y;
            if (worldW <= 0 || worldH <= 0) return null;

            int bmpW = (int)(worldW * scale) + 100;
            int bmpH = (int)(worldH * scale) + 100;

            // 限制位图最大尺寸，防止内存溢出 (400个元素可能跨度很大)
            bmpW = Math.Min(bmpW, 8000);
            bmpH = Math.Min(bmpH, 8000);

            Bitmap bmp = new Bitmap(bmpW, bmpH);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(System.Drawing.Color.Black);
                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    using (System.Drawing.Pen thinPen = new System.Drawing.Pen(System.Drawing.Color.White, 1.2f))
                    {
                        foreach (SelectedObject so in ss)
                        {
                            Curve curve = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                            if (curve == null) continue;

                            // 修复点：检查退化曲线（长度几乎为0的线）
                            if (Math.Abs(curve.EndParam - curve.StartParam) < 1e-6) continue;

                            int segments = 200;
                            PointF[] pts = new PointF[segments + 1];
                            double startP = curve.StartParam;
                            double endP = curve.EndParam;
                            double step = (endP - startP) / segments;

                            for (int i = 0; i <= segments; i++)
                            {
                                // 修复点：防止浮点数计算结果微弱超出 [StartParam, EndParam] 范围
                                double currentParam = startP + (step * i);
                                if (currentParam < startP) currentParam = startP;
                                if (currentParam > endP) currentParam = endP;

                                try
                                {
                                    AcadPoint3d p = curve.GetPointAtParameter(currentParam);
                                    float px = (float)((p.X - ext.MinPoint.X) * scale) + 50f;
                                    float py = (float)((ext.MaxPoint.Y - p.Y) * scale) + 50f;
                                    pts[i] = new PointF(px, py);
                                }
                                catch { continue; } // 如果某个点还是失败，跳过该点
                            }
                            g.DrawLines(thinPen, pts);
                        }
                    }
                    tr.Commit();
                }
            }
            return bmp;
        }

        // 4. OpenCV 提取轮廓 (显式指定点类型解决冲突)
        // 4. OpenCV 提取轮廓 - 修改版：识别内部闭合区域
        private List<List<AcadPoint2d>> ExtractCadContoursWithOpenCv(Bitmap bmp, Extents3d ext, double scale)
        {
            List<List<AcadPoint2d>> results = new List<List<AcadPoint2d>>();

            using (Mat src = bmp.ToMat())
            using (Mat gray = new Mat())
            {
                Cv2.CvtColor(src, gray, ColorConversionCodes.BGRA2GRAY);

                using (Mat binary = new Mat())
                {
                    // 阈值处理，确保线是纯白色(255)，背景是纯黑色(0)
                    Cv2.Threshold(gray, binary, _settings.Threshold, 255, ThresholdTypes.Binary);

                    CvPoint[][] contours;
                    HierarchyIndex[] hierarchy;

                    // 【核心修改1点】：使用 RetrievalModes.Tree 以获取层级关系
                    Cv2.FindContours(binary, out contours, out hierarchy, RetrievalModes.Tree, ContourApproximationModes.ApproxSimple);

                    for (int i = 0; i < contours.Length; i++)
                    {
                        // 【核心修改2点】：层级过滤
                        // 在 Tree 模式下，hierarchy[i].Parent != -1 表示该轮廓是一个“洞”或者嵌套在内的形状
                        // 我们绘制的线是白色，闭合区域（空腔）在 OpenCV 眼中就是白色线条包围的“黑色洞”
                        // 只有 Parent 不为 -1 的轮廓才是我们要的闭合区域内部边界
                        if (hierarchy[i].Parent == -1)
                            continue; // 跳过最外层的包络线（即你说的“把所有线生成一遍”的那个大圈）

                        var contour = contours[i];

                        // 过滤掉面积过小的杂质区域（噪点）
                        double area = Cv2.ContourArea(contour);
                        if (area < 100) continue;

                        double epsilon = 1.0 + (_settings.SmoothLevel * 0.2);
                        CvPoint[] approx = Cv2.ApproxPolyDP(contour, epsilon, true);

                        List<AcadPoint2d> cadPath = new List<AcadPoint2d>();
                        foreach (var p in approx)
                        {
                            // 映射像素坐标到 CAD 世界坐标
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
        private void DrawInCad(Document doc, List<List<AcadPoint2d>> contours)
        {
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
                        pl.AddVertexAt(i, pts[i], 0, 0, 0);
                    }
                    pl.Closed = true;
                    pl.Layer = layerName;
                    pl.ColorIndex = 256;
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