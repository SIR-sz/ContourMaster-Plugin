using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using OpenCvSharp;
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

        public void ProcessContour()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                PromptSelectionResult psr = ed.GetSelection(new PromptSelectionOptions { MessageForAdding = "\n请选择要生成轮廓的元素: " });
                if (psr.Status != PromptStatus.OK) return;

                Extents3d totalExt = GetSelectionExtents(psr.Value);

                // 逻辑转换：将用户输入的个位数（如 5）换算为千位（如 5000）
                int targetSize = _settings.PrecisionLevel * 1000;

                double worldW = totalExt.MaxPoint.X - totalExt.MinPoint.X;
                double worldH = totalExt.MaxPoint.Y - totalExt.MinPoint.Y;

                if (worldW <= 0 || worldH <= 0) return;
                double scale = targetSize / Math.Max(worldW, worldH);

                // 性能预检：在创建位图前检查内存安全性
                if (!IsMemorySafe(worldW, worldH, scale, ed)) return;

                using (Bitmap bmp = RasterizeSelection(doc, psr.Value, totalExt, scale))
                {
                    if (bmp == null) return;

                    List<List<AcadPoint2d>> cadContours = ExtractCadContoursWithOpenCv(bmp, totalExt, scale);

                    if (cadContours.Count == 0)
                    {
                        ed.WriteMessage("\n[未发现轮廓] 请确保选中的线能够构成闭合区域。建议调低“识别阈值”或增加“像素精度”。");
                        return;
                    }

                    DrawInCad(doc, cadContours, psr.Value, scale);
                    ed.WriteMessage($"\n[ContourMaster] 轮廓生成完毕，成功识别 {cadContours.Count} 个闭合区域。");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[异常] 运算中止: {ex.Message}");
            }
            finally
            {
                // 手动触发垃圾回收，释放大内存位图
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private bool IsMemorySafe(double worldW, double worldH, double scale, Editor ed)
        {
            int bmpW = (int)(worldW * scale) + 100;
            int bmpH = (int)(worldH * scale) + 100;
            long estimatedBytes = (long)bmpW * bmpH * 4;
            double estimatedMB = estimatedBytes / (1024.0 * 1024.0);

            if (bmpW > 12000 || bmpH > 12000 || estimatedMB > 600)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(
                    "❌ 精度过高：当前的设置可能导致内存溢出 (OOM)。\n请调低“像素精度”等级后再试。");
                return false;
            }
            return true;
        }

        private Bitmap RasterizeSelection(Document doc, SelectionSet ss, Extents3d ext, double scale)
        {
            int bmpW = (int)((ext.MaxPoint.X - ext.MinPoint.X) * scale) + 100;
            int bmpH = (int)((ext.MaxPoint.Y - ext.MinPoint.Y) * scale) + 100;

            if (bmpW <= 0 || bmpH <= 0 || bmpW > 15000 || bmpH > 15000) return null;

            Bitmap bmp = new Bitmap(bmpW, bmpH);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(System.Drawing.Color.Black);
                g.SmoothingMode = SmoothingMode.AntiAlias;

                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    // ✨ 核心修复：将画笔宽度设为 3.0f，线条必须有厚度才能产生“内孔”层级
                    using (System.Drawing.Pen thinPen = new System.Drawing.Pen(System.Drawing.Color.White, 3.0f))
                    {
                        foreach (SelectedObject so in ss)
                        {
                            Curve curve = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                            if (curve == null) continue;

                            double startParam = curve.StartParam;
                            double endParam = curve.EndParam;
                            if (Math.Abs(endParam - startParam) < 1e-7) continue;

                            int segments = 200;
                            List<PointF> pts = new List<PointF>();
                            double step = (endParam - startParam) / segments;

                            for (int i = 0; i <= segments; i++)
                            {
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
                            if (pts.Count > 1) g.DrawLines(thinPen, pts.ToArray());
                        }
                    }
                    tr.Commit();
                }
            }
            return bmp;
        }

        // sir-sz/contourmaster-plugin/ContourMaster-Plugin-fc9a6c7e482c9a7960aee9d2ceb674cbf4746bcd/Plugin_ContourMaster/Services/ContourEngine.cs

        private List<List<AcadPoint2d>> ExtractCadContoursWithOpenCv(Bitmap bmp, Extents3d ext, double scale)
        {
            List<List<AcadPoint2d>> results = new List<List<AcadPoint2d>>();

            using (Mat src = bmp.ToMat())
            using (Mat gray = new Mat())
            {
                Cv2.CvtColor(src, gray, ColorConversionCodes.BGRA2GRAY);

                using (Mat binary = new Mat())
                {
                    int thresholdValue = (int)Math.Max(10.0, Math.Min(240.0, _settings.Threshold));
                    Cv2.Threshold(gray, binary, thresholdValue, 255, ThresholdTypes.Binary);

                    // ✨ 核心逻辑：将 CAD 补缝距离换算为像素核大小
                    // 计算公式：核尺寸 = 补缝距离(CAD) * 缩放比例(Scale) * 1.2(冗余系数)
                    int kSize = (int)Math.Ceiling(_settings.SimplifyTolerance * scale * 1.2);

                    // 确保核大小至少为 3，且必须为奇数
                    if (kSize < 3) kSize = 3;
                    if (kSize % 2 == 0) kSize++;

                    // 限制核上限（如101），防止超大数值导致内存崩溃
                    kSize = Math.Min(kSize, 101);

                    using (Mat kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(kSize, kSize)))
                    {
                        Cv2.MorphologyEx(binary, binary, MorphTypes.Close, kernel);
                    }

                    CvPoint[][] contours;
                    HierarchyIndex[] hierarchy;
                    Cv2.FindContours(binary, out contours, out hierarchy, RetrievalModes.Tree, ContourApproximationModes.ApproxSimple);

                    if (contours == null || contours.Length == 0) return results;

                    for (int i = 0; i < contours.Length; i++)
                    {
                        if (hierarchy[i].Parent == -1) continue;

                        var contour = contours[i];
                        if (contour.Length < 4) continue;
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
                        if (_settings.SmoothLevel > 3) cadPath = SmoothCadPoints(cadPath);
                        results.Add(cadPath);
                    }
                }
            }
            return results;
        }

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
                    int vertexIndex = 0;
                    AcadPoint2d lastAddedPt = new AcadPoint2d(double.NaN, double.NaN);

                    for (int i = 0; i < pts.Count; i++)
                    {
                        AcadPoint2d snappedPt = SnapPointToOriginalCurves(pts[i], originalLines, tr, pixelSize);
                        if (vertexIndex > 0 && snappedPt.GetDistanceTo(lastAddedPt) < 0.001) continue;
                        pl.AddVertexAt(vertexIndex++, snappedPt, 0, 0, 0);
                        lastAddedPt = snappedPt;
                    }

                    if (pl.NumberOfVertices > 2)
                    {
                        if (pl.GetPoint2dAt(0).GetDistanceTo(pl.GetPoint2dAt(pl.NumberOfVertices - 1)) < 0.01)
                            pl.RemoveVertexAt(pl.NumberOfVertices - 1);
                        pl.Closed = true;
                    }

                    pl.Layer = layerName;
                    pl.ColorIndex = 256;
                    btr.AppendEntity(pl);
                    tr.AddNewlyCreatedDBObject(pl, true);
                }
                tr.Commit();
            }
        }

        // sir-sz/contourmaster-plugin/ContourMaster-Plugin-fc9a6c7e482c9a7960aee9d2ceb674cbf4746bcd/Plugin_ContourMaster/Services/ContourEngine.cs

        private AcadPoint2d SnapPointToOriginalCurves(AcadPoint2d pt, SelectionSet originalSs, Transaction tr, double pixelSize)
        {
            AcadPoint3d pt3d = new AcadPoint3d(pt.X, pt.Y, 0);

            // ✨ 核心逻辑：吸附阈值应大于等于补缝距离的一半，通常取 1.5 倍补缝距离以确保覆盖
            double baseSnap = pixelSize * 6.0;
            double snapThreshold = Math.Max(baseSnap, _settings.SimplifyTolerance * 1.5);
            double searchThreshold = snapThreshold + (pixelSize * 2.0);

            List<Curve> nearbyCurves = new List<Curve>();
            List<AcadPoint3d> criticalPoints = new List<AcadPoint3d>();

            foreach (SelectedObject so in originalSs)
            {
                Curve curve = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                if (curve == null) continue;

                try
                {
                    Extents3d ext = curve.GeometricExtents;
                    if (pt.X < ext.MinPoint.X - searchThreshold || pt.X > ext.MaxPoint.X + searchThreshold ||
                        pt.Y < ext.MinPoint.Y - searchThreshold || pt.Y > ext.MaxPoint.Y + searchThreshold)
                        continue;

                    nearbyCurves.Add(curve);
                    criticalPoints.Add(curve.StartPoint);
                    criticalPoints.Add(curve.EndPoint);
                }
                catch { }
            }

            AcadPoint3d bestSnapPt = pt3d;
            double minCriticalDist = snapThreshold;
            bool foundCritical = false;

            // 1. 交点/端点优先
            foreach (AcadPoint3d cp in criticalPoints)
            {
                double d = pt3d.DistanceTo(cp);
                if (d < minCriticalDist) { minCriticalDist = d; bestSnapPt = cp; foundCritical = true; }
            }

            if (nearbyCurves.Count >= 2)
            {
                for (int i = 0; i < nearbyCurves.Count; i++)
                {
                    for (int j = i + 1; j < nearbyCurves.Count; j++)
                    {
                        using (Point3dCollection interPts = new Point3dCollection())
                        {
                            nearbyCurves[i].IntersectWith(nearbyCurves[j], Intersect.OnBothOperands, interPts, IntPtr.Zero, IntPtr.Zero);
                            foreach (AcadPoint3d ip in interPts)
                            {
                                double d = pt3d.DistanceTo(ip);
                                if (d < minCriticalDist) { minCriticalDist = d; bestSnapPt = ip; foundCritical = true; }
                            }
                        }
                    }
                }
            }

            if (foundCritical) return new AcadPoint2d(bestSnapPt.X, bestSnapPt.Y);

            // 2. 曲线吸附
            AcadPoint3d bestLinePt = pt3d;
            double minLineDist = snapThreshold;
            foreach (Curve curve in nearbyCurves)
            {
                try
                {
                    AcadPoint3d closest = curve.GetClosestPointTo(pt3d, false);
                    double d = pt3d.DistanceTo(closest);
                    if (d < minLineDist) { minLineDist = d; bestLinePt = closest; }
                }
                catch { }
            }
            return new AcadPoint2d(bestLinePt.X, bestLinePt.Y);
        }

        private void CheckAndCreateLayer(Database db, Transaction tr, string layerName)
        {
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                LayerTableRecord ltr = new LayerTableRecord { Name = layerName };
                ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 2);
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }
    }
}