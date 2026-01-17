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
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Tesseract; // 引入 Tesseract

// ✨ 定义别名以彻底解决命名冲突
using AcadPoint2d = Autodesk.AutoCAD.Geometry.Point2d;
using AcadPoint3d = Autodesk.AutoCAD.Geometry.Point3d;
using CvPoint = OpenCvSharp.Point;

namespace Plugin_ContourMaster.Services
{
    public class ContourEngine
    {
        private readonly ContourSettings _settings;

        public ContourEngine(ContourSettings settings) => _settings = settings;

        #region 1. CAD 选区识别逻辑

        public void ProcessContour()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                PromptSelectionResult psr = ed.GetSelection(new PromptSelectionOptions { MessageForAdding = "\n请选择要生成轮廓的矢量元素: " });
                if (psr.Status != PromptStatus.OK) return;

                Extents3d totalExt = GetSelectionExtents(psr.Value);

                int targetSize = _settings.PrecisionLevel * 1000;
                double worldW = totalExt.MaxPoint.X - totalExt.MinPoint.X;
                double worldH = totalExt.MaxPoint.Y - totalExt.MinPoint.Y;

                if (worldW <= 0 || worldH <= 0) return;
                double scale = targetSize / Math.Max(worldW, worldH);

                if (!IsMemorySafe(worldW, worldH, scale, ed)) return;

                using (Bitmap bmp = RasterizeSelection(doc, psr.Value, totalExt, scale))
                {
                    if (bmp == null) return;

                    List<List<AcadPoint2d>> cadContours = ExtractContoursFromBitmap(bmp, totalExt, scale, true);

                    if (cadContours.Count == 0)
                    {
                        ed.WriteMessage("\n[提示] 未发现闭合区域。");
                        return;
                    }

                    DrawInCad(doc, cadContours, psr.Value, scale);
                    ed.WriteMessage($"\n[ContourMaster] 成功识别 {cadContours.Count} 个区域。");
                }
            }
            catch (System.Exception ex) { ed.WriteMessage($"\n[异常] {ex.Message}"); }
            finally { CleanMemory(); }
        }

        #endregion

        #region 2. 图像参照原位识别逻辑

        public void ProcessReferencedImage()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                PromptEntityOptions peo = new PromptEntityOptions("\n请选择要识别的图像参照: ");
                peo.SetRejectMessage("\n选择的对象必须是有效的图像参照。");
                peo.AddAllowedClass(typeof(RasterImage), false);

                PromptEntityResult per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;

                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    RasterImage img = tr.GetObject(per.ObjectId, OpenMode.ForRead) as RasterImage;
                    if (img == null) return;

                    RasterImageDef imgDef = tr.GetObject(img.ImageDefId, OpenMode.ForRead) as RasterImageDef;
                    if (!File.Exists(imgDef.ActiveFileName)) throw new Exception("找不到图像源文件。");

                    using (Bitmap bmp = new Bitmap(imgDef.ActiveFileName))
                    {
                        Matrix3d pixelToWorld = GetEntityTransform(img, bmp.Width, bmp.Height);

                        if (_settings.IsOcrMode)
                        {
                            // 执行 OCR 模式
                            ProcessImageOcr(bmp, pixelToWorld);
                        }
                        else
                        {
                            // 执行原有的线条提取模式
                            List<List<AcadPoint2d>> contours = ExtractContoursReferenced(bmp, pixelToWorld);

                            if (contours.Count > 0)
                            {
                                DrawInCadSimple(doc, contours, "图片_原位识别");
                                ed.WriteMessage($"\n[成功] 已在原位生成 {contours.Count} 条轮廓。");
                            }
                            else
                            {
                                ed.WriteMessage("\n[提示] 未识别到有效轮廓，请调低“识别阈值”。");
                            }
                        }
                        tr.Commit();
                    }
                }
            catch (Exception ex) { ed.WriteMessage($"\n[错误] 操作失败: {ex.Message}"); }
            finally { CleanMemory(); }
        }

        #endregion

        #region 3. 核心算法逻辑

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

                    if (!onlyHoles && Cv2.CountNonZero(binary) > (binary.Rows * binary.Cols / 2))
                        Cv2.BitwiseNot(binary, binary);

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
                        if (contours[i].Length < 4 || Cv2.ContourArea(contours[i]) < 5) continue; // 降低面积过滤保留细节

                        double epsilon = _settings.SmoothLevel * 0.2;
                        CvPoint[] approx = epsilon > 0 ? Cv2.ApproxPolyDP(contours[i], epsilon, true) : contours[i];

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
        private void ProcessImageOcr(Bitmap bmp, Matrix3d transform)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // 1. 初始化 Tesseract 引擎
                string tessdataPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "tessdata");
                using (var engine = new TesseractEngine(tessdataPath, "chi_sim+eng", EngineMode.Default))
                {
                    // 2. 将 Bitmap 转换为 Tesseract 识别格式
                    using (var pix = PixConverter.ToPix(bmp))
                    using (var page = engine.Process(pix))
                    {
                        // 3. 遍历识别结果
                        using (var iter = page.GetIterator())
                        {
                            iter.Begin();

                            using (var loc = doc.LockDocument())
                            using (Transaction tr = doc.TransactionManager.StartTransaction())
                            {
                                var btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(doc.Database), OpenMode.ForWrite);
                                CheckAndCreateLayer(doc.Database, tr, "OCR_识别文字");

                                do
                                {
                                    if (iter.TryGetBoundingBox(PageIteratorLevel.Word, out Rect bounds))
                                    {
                                        string text = iter.GetText(PageIteratorLevel.Word);
                                        if (string.IsNullOrWhiteSpace(text)) continue;

                                        // 4. 坐标转换：像素点 -> CAD 世界坐标
                                        // Tesseract 返回的 Y 坐标是自顶向下的，需翻转
                                        AcadPoint3d pixelLoc = new AcadPoint3d(bounds.X1, bmp.Height - bounds.Y2, 0);
                                        AcadPoint3d worldLoc = pixelLoc.TransformBy(transform);

                                        // 5. 计算文字高度 (像素高度 * 缩放比)
                                        double textHeight = bounds.Height * transform.GetScaleExtent();

                                        // 6. 创建 CAD 文字对象
                                        using (DBText dbText = new DBText())
                                        {
                                            dbText.Contents = text;
                                            dbText.Position = worldLoc;
                                            dbText.Height = textHeight;
                                            dbText.Layer = "OCR_识别文字";

                                            btr.AppendEntity(dbText);
                                            tr.AddNewlyCreatedDBObject(dbText, true);
                                        }
                                    }
                                } while (iter.Next(PageIteratorLevel.Word));

                                tr.Commit();
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { ed.WriteMessage($"\n[OCR错误] {ex.Message}"); }
        }
        private List<List<AcadPoint2d>> ExtractContoursReferenced(Bitmap bmp, Matrix3d transform)
        {
            List<List<AcadPoint2d>> results = new List<List<AcadPoint2d>>();
            using (Mat src = bmp.ToMat())
            using (Mat gray = new Mat())
            {
                Cv2.CvtColor(src, gray, ColorConversionCodes.BGRA2GRAY);
                using (Mat binary = new Mat())
                {
                    Cv2.Threshold(gray, binary, _settings.Threshold, 255, ThresholdTypes.Binary);
                    if (Cv2.CountNonZero(binary) > (binary.Rows * binary.Cols / 2)) Cv2.BitwiseNot(binary, binary);

                    CvPoint[][] contours;
                    HierarchyIndex[] hierarchy;
                    Cv2.FindContours(binary, out contours, out hierarchy, RetrievalModes.List, ContourApproximationModes.ApproxSimple);

                    foreach (var c in contours)
                    {
                        if (c.Length < 4 || Cv2.ContourArea(c) < 5) continue; // 降低面积过滤

                        double epsilon = _settings.SmoothLevel * 0.2;
                        CvPoint[] approx = epsilon > 0 ? Cv2.ApproxPolyDP(c, epsilon, true) : c;

                        List<AcadPoint2d> pts = new List<AcadPoint2d>();
                        foreach (var p in approx)
                        {
                            AcadPoint3d pixelPt = new AcadPoint3d(p.X, bmp.Height - p.Y, 0);
                            AcadPoint3d worldPt = pixelPt.TransformBy(transform);
                            pts.Add(new AcadPoint2d(worldPt.X, worldPt.Y));
                        }
                        results.Add(pts);
                    }
                }
            }
            return results;
        }

        #endregion

        #region 4. 坐标转换与纠偏逻辑

        private Matrix3d GetEntityTransform(Entity ent, int imgW, int imgH)
        {
            Extents3d ext = ent.GeometricExtents;
            double cadW = ext.MaxPoint.X - ext.MinPoint.X;
            double cadH = ext.MaxPoint.Y - ext.MinPoint.Y;
            double scX = cadW / imgW;
            double scY = cadH / imgH;
            return Matrix3d.Displacement(ext.MinPoint.GetAsVector()) * Matrix3d.Scaling(scX, ext.MinPoint);
        }

        private AcadPoint2d SnapPointToOriginalCurves(AcadPoint2d pt, SelectionSet originalSs, Transaction tr, double pixelSize)
        {
            AcadPoint3d pt3d = new AcadPoint3d(pt.X, pt.Y, 0);
            double snapThreshold = Math.Max(pixelSize * 6.0, _settings.SimplifyTolerance * 1.5);
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

        #endregion

        #region 5. 辅助与绘制方法

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
                Application.ShowAlertDialog("❌ 精度过高，可能导致内存溢出。");
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
                    foreach (var p in pts)
                    {
                        AcadPoint2d snapped = SnapPointToOriginalCurves(p, originalLines, tr, pixelSize);
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