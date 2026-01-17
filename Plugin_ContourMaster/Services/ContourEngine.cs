using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using PdfiumViewer; // ✨ 需确保已安装并引用 PdfiumViewer
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

        #region 公开入口方法

        // 1. 处理 CAD 选区提取 (识别闭合区域)
        public void ProcessContour()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                PromptSelectionResult psr = ed.GetSelection(new PromptSelectionOptions { MessageForAdding = "\n请选择要生成轮廓的元素: " });
                if (psr.Status != PromptStatus.OK) return;

                Extents3d totalExt = GetSelectionExtents(psr.Value);

                // 换算精度：等级 * 1000
                int targetSize = _settings.PrecisionLevel * 1000;
                double worldW = totalExt.MaxPoint.X - totalExt.MinPoint.X;
                double worldH = totalExt.MaxPoint.Y - totalExt.MinPoint.Y;

                if (worldW <= 0 || worldH <= 0) return;
                double scale = targetSize / Math.Max(worldW, worldH);

                if (!IsMemorySafe(worldW, worldH, scale, ed)) return;

                using (Bitmap bmp = RasterizeSelection(doc, psr.Value, totalExt, scale))
                {
                    if (bmp == null) return;

                    // ✨ 核心调用：CAD 模式 onlyHoles 设为 true
                    List<List<AcadPoint2d>> cadContours = ExtractContoursFromBitmap(bmp, totalExt, scale, true);

                    if (cadContours.Count == 0)
                    {
                        ed.WriteMessage("\n[提示] 未发现闭合区域。");
                        return;
                    }

                    DrawInCad(doc, cadContours, psr.Value, scale);
                    ed.WriteMessage($"\n[ContourMaster] 成功识别 {cadContours.Count} 个闭合区域。");
                }
            }
            catch (System.Exception ex) { ed.WriteMessage($"\n[异常] {ex.Message}"); }
            finally { CleanMemory(); }
        }

        // 2. 识别外部图片
        public void ProcessImageContour()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            var ofd = new System.Windows.Forms.OpenFileDialog { Filter = "图片文件|*.jpg;*.png;*.bmp;*.jpeg" };
            if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

            try
            {
                using (Bitmap bmp = new Bitmap(ofd.FileName))
                {
                    ProcessExternalBitmapInternal(bmp, "图片提取结果");
                }
            }
            catch (System.Exception ex) { ed.WriteMessage($"\n[错误] 加载图片失败: {ex.Message}"); }
            finally { CleanMemory(); }
        }

        // 3. 识别 PDF 文件
        public void ProcessPdfContour()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            var ofd = new System.Windows.Forms.OpenFileDialog { Filter = "PDF文件|*.pdf" };
            if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

            try
            {
                ed.WriteMessage("\n正在解析 PDF，请稍候...");
                using (var pdfDoc = PdfDocument.Load(ofd.FileName))
                {
                    // 渲染第一页，300 DPI 保证识别精度
                    using (Bitmap bmp = (Bitmap)pdfDoc.Render(0, 300, 300, true))
                    {
                        ProcessExternalBitmapInternal(bmp, "PDF提取结果");
                    }
                }
            }
            catch (System.Exception ex) { ed.WriteMessage($"\n[错误] PDF 解析失败: {ex.Message}"); }
            finally { CleanMemory(); }
        }

        #endregion

        #region 内部算法逻辑

        private void ProcessExternalBitmapInternal(Bitmap bmp, string layerName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            // 外部文件放置在原点，比例为 1.0
            Extents3d dummyExt = new Extents3d(new AcadPoint3d(0, 0, 0), new AcadPoint3d(bmp.Width, bmp.Height, 0));

            // ✨ 外部识别模式：onlyHoles 设为 false 以提取所有线条
            List<List<AcadPoint2d>> cadContours = ExtractContoursFromBitmap(bmp, dummyExt, 1.0, false);

            if (cadContours.Count > 0)
            {
                DrawInCadSimple(doc, cadContours, layerName);
            }
        }

        private List<List<AcadPoint2d>> ExtractContoursFromBitmap(Bitmap bmp, Extents3d ext, double scale, bool onlyHoles)
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

                    // 自动纠偏：识别图片/PDF时，如果是白底黑线则反色
                    if (!onlyHoles && Cv2.CountNonZero(binary) > (binary.Rows * binary.Cols / 2))
                        Cv2.BitwiseNot(binary, binary);

                    // 动态补缝距离映射
                    int kSize = (int)Math.Ceiling(_settings.SimplifyTolerance * scale * 1.2);
                    if (kSize < 3) kSize = 3; if (kSize % 2 == 0) kSize++;
                    kSize = Math.Min(kSize, 101);

                    using (Mat kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(kSize, kSize)))
                    {
                        Cv2.MorphologyEx(binary, binary, MorphTypes.Close, kernel);
                    }

                    CvPoint[][] contours;
                    HierarchyIndex[] hierarchy;
                    var mode = onlyHoles ? RetrievalModes.Tree : RetrievalModes.List;
                    Cv2.FindContours(binary, out contours, out hierarchy, mode, ContourApproximationModes.ApproxSimple);

                    if (contours == null) return results;

                    for (int i = 0; i < contours.Length; i++)
                    {
                        if (onlyHoles && hierarchy[i].Parent == -1) continue;
                        if (contours[i].Length < 4 || Cv2.ContourArea(contours[i]) < 10) continue;

                        double epsilon = 1.0 + (_settings.SmoothLevel * 0.2);
                        CvPoint[] approx = Cv2.ApproxPolyDP(contours[i], epsilon, true);

                        List<AcadPoint2d> cadPath = new List<AcadPoint2d>();
                        foreach (var p in approx)
                        {
                            double offset = onlyHoles ? 50.0 : 0.0;
                            double wx = (p.X + 0.5 - offset) / scale + ext.MinPoint.X;
                            double wy = ext.MaxPoint.Y - (p.Y + 0.5 - offset) / scale;
                            cadPath.Add(new AcadPoint2d(wx, wy));
                        }
                        results.Add(cadPath);
                    }
                }
            }
            return results;
        }

        #endregion

        #region 辅助与绘制方法 (解决 CS0103)

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

        private bool IsMemorySafe(double w, double h, double s, Editor ed)
        {
            int bw = (int)(w * s) + 100; int bh = (int)(h * s) + 100;
            if (bw > 12000 || bh > 12000)
            {
                Application.ShowAlertDialog("❌ 精度过高，可能导致内存溢出，请调低精度等级。");
                return false;
            }
            return true;
        }

        private Bitmap RasterizeSelection(Document doc, SelectionSet ss, Extents3d ext, double scale)
        {
            int bw = (int)((ext.MaxPoint.X - ext.MinPoint.X) * scale) + 100;
            int bh = (int)((ext.MaxPoint.Y - ext.MinPoint.Y) * scale) + 100;
            if (bw <= 0 || bh <= 0) return null;

            Bitmap bmp = new Bitmap(bw, bh);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(System.Drawing.Color.Black);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    using (System.Drawing.Pen p = new System.Drawing.Pen(System.Drawing.Color.White, 3.0f))
                    {
                        foreach (SelectedObject so in ss)
                        {
                            Curve c = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                            if (c == null) continue;
                            int segs = 200; double step = (c.EndParam - c.StartParam) / segs;
                            List<PointF> pts = new List<PointF>();
                            for (int i = 0; i <= segs; i++)
                            {
                                AcadPoint3d pt = c.GetPointAtParameter(c.StartParam + step * i);
                                pts.Add(new PointF((float)((pt.X - ext.MinPoint.X) * scale) + 50f, (float)((ext.MaxPoint.Y - pt.Y) * scale) + 50f));
                            }
                            if (pts.Count > 1) g.DrawLines(p, pts.ToArray());
                        }
                    }
                    tr.Commit();
                }
            }
            return bmp;
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
                    Polyline pl = new Polyline();
                    int idx = 0; AcadPoint2d lastAdded = new AcadPoint2d(double.NaN, double.NaN);
                    for (int i = 0; i < pts.Count; i++)
                    {
                        AcadPoint2d snapped = SnapPointToOriginalCurves(pts[i], originalLines, tr, pixelSize);
                        if (idx > 0 && snapped.GetDistanceTo(lastAdded) < 0.001) continue;
                        pl.AddVertexAt(idx++, snapped, 0, 0, 0); lastAdded = snapped;
                    }
                    if (pl.NumberOfVertices > 2)
                    {
                        if (pl.GetPoint2dAt(0).GetDistanceTo(pl.GetPoint2dAt(pl.NumberOfVertices - 1)) < 0.01)
                            pl.RemoveVertexAt(pl.NumberOfVertices - 1);
                        pl.Closed = true;
                    }
                    pl.Layer = layerName; btr.AppendEntity(pl); tr.AddNewlyCreatedDBObject(pl, true);
                }
                tr.Commit();
            }
        }

        private AcadPoint2d SnapPointToOriginalCurves(AcadPoint2d pt, SelectionSet originalSs, Transaction tr, double pixelSize)
        {
            AcadPoint3d pt3d = new AcadPoint3d(pt.X, pt.Y, 0);
            double baseSnap = pixelSize * 6.0;
            double snapThreshold = Math.Max(baseSnap, _settings.SimplifyTolerance * 1.5);
            double searchThreshold = snapThreshold + (pixelSize * 2.0);

            List<Curve> curves = new List<Curve>();
            List<AcadPoint3d> criticalPts = new List<AcadPoint3d>();

            foreach (SelectedObject so in originalSs)
            {
                Curve c = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                if (c == null) continue;
                try
                {
                    Extents3d ex = c.GeometricExtents;
                    if (pt.X < ex.MinPoint.X - searchThreshold || pt.X > ex.MaxPoint.X + searchThreshold ||
                        pt.Y < ex.MinPoint.Y - searchThreshold || pt.Y > ex.MaxPoint.Y + searchThreshold) continue;
                    curves.Add(c); criticalPts.Add(c.StartPoint); criticalPts.Add(c.EndPoint);
                }
                catch { }
            }

            AcadPoint3d bestSnap = pt3d; double minDist = snapThreshold; bool found = false;
            foreach (AcadPoint3d cp in criticalPts)
            {
                double d = pt3d.DistanceTo(cp);
                if (d < minDist) { minDist = d; bestSnap = cp; found = true; }
            }
            if (curves.Count >= 2)
            {
                for (int i = 0; i < curves.Count; i++)
                {
                    for (int j = i + 1; j < curves.Count; j++)
                    {
                        using (Point3dCollection ips = new Point3dCollection())
                        {
                            curves[i].IntersectWith(curves[j], Intersect.OnBothOperands, ips, IntPtr.Zero, IntPtr.Zero);
                            foreach (AcadPoint3d ip in ips)
                            {
                                double d = pt3d.DistanceTo(ip);
                                if (d < minDist) { minDist = d; bestSnap = ip; found = true; }
                            }
                        }
                    }
                }
            }
            if (found) return new AcadPoint2d(bestSnap.X, bestSnap.Y);
            foreach (Curve c in curves)
            {
                try
                {
                    AcadPoint3d close = c.GetClosestPointTo(pt3d, false);
                    double d = pt3d.DistanceTo(close);
                    if (d < minDist) { minDist = d; bestSnap = close; }
                }
                catch { }
            }
            return new AcadPoint2d(bestSnap.X, bestSnap.Y);
        }

        private void DrawInCadSimple(Document doc, List<List<AcadPoint2d>> contours, string layerName)
        {
            using (var loc = doc.LockDocument())
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(doc.Database), OpenMode.ForWrite);
                CheckAndCreateLayer(doc.Database, tr, layerName);
                foreach (var pts in contours)
                {
                    Polyline pl = new Polyline();
                    for (int i = 0; i < pts.Count; i++) pl.AddVertexAt(i, pts[i], 0, 0, 0);
                    pl.Closed = true; pl.Layer = layerName;
                    btr.AppendEntity(pl); tr.AddNewlyCreatedDBObject(pl, true);
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
                ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 2);
                lt.Add(ltr); tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }

        private void CleanMemory() { GC.Collect(); GC.WaitForPendingFinalizers(); }

        #endregion
    }
}