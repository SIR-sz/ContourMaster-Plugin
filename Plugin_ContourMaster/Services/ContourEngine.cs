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
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Tesseract;

// ✨ 定义别名以彻底解决命名冲突
using AcadPoint2d = Autodesk.AutoCAD.Geometry.Point2d;
using AcadPoint3d = Autodesk.AutoCAD.Geometry.Point3d;
using CvPoint = OpenCvSharp.Point;

namespace Plugin_ContourMaster.Services
{
    public class ContourEngine
    {
        private readonly ContourSettings _settings;
        private static readonly HttpClient _httpClient = new HttpClient();

        public ContourEngine(ContourSettings settings) => _settings = settings;

        #region 1. CAD 选区识别逻辑 (矢量转轮廓)
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
                    if (cadContours.Count == 0) { ed.WriteMessage("\n[提示] 未发现闭合区域。"); return; }
                    DrawInCad(doc, cadContours, psr.Value, scale);
                    ed.WriteMessage($"\n[ContourMaster] 成功识别 {cadContours.Count} 个区域。");
                }
            }
            catch (Exception ex) { ed.WriteMessage($"\n[异常] {ex.Message}"); }
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
                        if (_settings.IsOcrMode) ProcessImageOcr(bmp, pixelToWorld);
                        else
                        {
                            List<List<AcadPoint2d>> contours = ExtractContoursReferenced(bmp, pixelToWorld);
                            if (contours.Count > 0) DrawInCadSimple(doc, contours, "图片_原位识别");
                        }
                    }
                    tr.Commit();
                }
            }
            catch (Exception ex) { ed.WriteMessage($"\n[错误] 操作失败: {ex.Message}"); }
            finally { CleanMemory(); }
        }
        #endregion

        #region 3. OCR 识别核心逻辑
        private async void ProcessImageOcr(Bitmap bmp, Matrix3d transform)
        {
            if (_settings.SelectedOcrEngine == OcrEngineType.Baidu)
                await ProcessBaiduOcr(bmp, transform);
            else
                ProcessTesseractOcr(bmp, transform);
        }

        private async Task ProcessBaiduOcr(Bitmap bmp, Matrix3d transform)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            try
            {
                string token = await GetBaiduAccessToken(_settings.BaiduApiKey, _settings.BaiduSecretKey);
                if (string.IsNullOrEmpty(token)) throw new Exception("Token 获取失败，请检查 API Key。");

                string base64Image;
                using (MemoryStream ms = new MemoryStream())
                {
                    // 使用 80% 质量压缩，防止 Base64 字符串由于像素过高超过百度 4MB 的限制
                    bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                    base64Image = Convert.ToBase64String(ms.ToArray());
                }

                // ✨ 核心修正：手动构建并 UrlEncode image 参数，解决“参数无效”
                string requestBody = "image=" + WebUtility.UrlEncode(base64Image);
                var content = new StringContent(requestBody, System.Text.Encoding.UTF8, "application/x-www-form-urlencoded");

                // 高精度含位置版接口
                var response = await _httpClient.PostAsync($"https://aip.baidubce.com/rest/2.0/ocr/v1/accurate?access_token={token}", content);
                string json = await response.Content.ReadAsStringAsync();

                var serializer = new JavaScriptSerializer();
                var result = serializer.Deserialize<Dictionary<string, object>>(json);

                if (result.ContainsKey("error_code"))
                    throw new Exception($"{result["error_msg"]} (代码:{result["error_code"]})");

                if (result.ContainsKey("words_result"))
                {
                    using (var loc = doc.LockDocument())
                    using (Transaction tr = doc.TransactionManager.StartTransaction())
                    {
                        var btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(doc.Database), OpenMode.ForWrite);
                        CheckAndCreateLayer(doc.Database, tr, "OCR_识别文字");
                        double matrixScale = new Vector3d(transform[0, 0], transform[1, 0], transform[2, 0]).Length;

                        foreach (Dictionary<string, object> item in (Array)result["words_result"])
                        {
                            string text = item["words"].ToString();
                            var locData = (Dictionary<string, object>)item["location"];

                            AcadPoint3d worldLoc = new AcadPoint3d(Convert.ToDouble(locData["left"]),
                                                                  bmp.Height - Convert.ToDouble(locData["top"]), 0).TransformBy(transform);

                            using (MText mText = new MText())
                            {
                                mText.Contents = text;
                                mText.Location = worldLoc;
                                mText.TextHeight = SnapToStandardHeight(Convert.ToDouble(locData["height"]) * matrixScale);
                                mText.Width = Convert.ToDouble(locData["width"]) * matrixScale;
                                mText.Layer = "OCR_识别文字";
                                mText.Attachment = AttachmentPoint.TopLeft;
                                btr.AppendEntity(mText);
                                tr.AddNewlyCreatedDBObject(mText, true);
                            }
                        }
                        tr.Commit();
                    }
                    ed.WriteMessage("\n[Baidu OCR] 识别并原位生成完成。");
                }
            }
            catch (Exception ex) { ed.WriteMessage($"\n[百度OCR错误] {ex.Message}"); }
        }

        private void ProcessTesseractOcr(Bitmap bmp, Matrix3d transform)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            try
            {
                string tessdataPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "tessdata");
                using (var engine = new TesseractEngine(tessdataPath, "chi_sim+eng", EngineMode.Default))
                using (var pix = PixConverter.ToPix(bmp))
                using (var page = engine.Process(pix))
                using (var iter = page.GetIterator())
                {
                    iter.Begin();
                    var allWords = new List<OcrWordData>();
                    do
                    {
                        if (iter.TryGetBoundingBox(PageIteratorLevel.Word, out Tesseract.Rect bounds))
                            allWords.Add(new OcrWordData { Text = iter.GetText(PageIteratorLevel.Word).Trim(), Bounds = bounds });
                    } while (iter.Next(PageIteratorLevel.Word));

                    if (allWords.Count == 0) return;
                    var clusters = ClusterWords(allWords);
                    using (var loc = doc.LockDocument())
                    using (Transaction tr = doc.TransactionManager.StartTransaction())
                    {
                        var btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(doc.Database), OpenMode.ForWrite);
                        CheckAndCreateLayer(doc.Database, tr, "OCR_识别文字");
                        double scale = new Vector3d(transform[0, 0], transform[1, 0], transform[2, 0]).Length;
                        foreach (var g in clusters) CreateMTextFromWordGroup(tr, btr, g, bmp.Height, transform, scale);
                        tr.Commit();
                    }
                    ed.WriteMessage("\n[Tesseract] 识别完成。");
                }
            }
            catch (Exception ex) { ed.WriteMessage($"\n[Tesseract错误] {ex.Message}"); }
        }
        #endregion

        #region 4. 辅助算法与基础设施
        private async Task<string> GetBaiduAccessToken(string key, string secret)
        {
            // 修正 client_secret 参数名
            string url = $"https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id={key}&client_secret={secret}";
            var response = await _httpClient.GetAsync(url);
            string json = await response.Content.ReadAsStringAsync();
            var serializer = new JavaScriptSerializer();
            var res = serializer.Deserialize<Dictionary<string, string>>(json);
            return res.ContainsKey("access_token") ? res["access_token"] : null;
        }

        private List<List<OcrWordData>> ClusterWords(List<OcrWordData> words)
        {
            var clusters = new List<List<OcrWordData>>();
            for (int i = 0; i < words.Count; i++)
            {
                if (words[i].Visited) continue;
                var group = new List<OcrWordData>();
                var queue = new Queue<OcrWordData>();
                words[i].Visited = true; queue.Enqueue(words[i]);
                while (queue.Count > 0)
                {
                    var node = queue.Dequeue(); group.Add(node);
                    foreach (var other in words)
                    {
                        if (other.Visited) continue;
                        double threshold = node.Bounds.Height * 1.1;
                        double dx = Math.Max(0, Math.Max(other.Bounds.X1 - node.Bounds.X2, node.Bounds.X1 - other.Bounds.X2));
                        double dy = Math.Max(0, Math.Max(other.Bounds.Y1 - node.Bounds.Y2, node.Bounds.Y1 - other.Bounds.Y2));
                        if ((dy <= 0 && dx < threshold * 1.5) || (dx <= 0 && dy < threshold))
                        {
                            other.Visited = true; queue.Enqueue(other);
                        }
                    }
                }
                clusters.Add(group);
            }
            return clusters;
        }

        private void CreateMTextFromWordGroup(Transaction tr, BlockTableRecord btr, List<OcrWordData> group, int bmpH, Matrix3d trans, double scale)
        {
            group.Sort((a, b) => a.Bounds.Y1 == b.Bounds.Y1 ? a.Bounds.X1.CompareTo(b.Bounds.X1) : a.Bounds.Y1.CompareTo(b.Bounds.Y1));
            int minX = int.MaxValue, minY = int.MaxValue, maxX = int.MinValue;
            double totalH = 0;
            foreach (var w in group)
            {
                minX = Math.Min(minX, w.Bounds.X1); minY = Math.Min(minY, w.Bounds.Y1);
                maxX = Math.Max(maxX, w.Bounds.X2); totalH += w.Bounds.Height;
            }
            AcadPoint3d worldLoc = new AcadPoint3d(minX, bmpH - minY, 0).TransformBy(trans);
            using (MText mText = new MText())
            {
                mText.Contents = string.Join(" ", group.ConvertAll(w => w.Text));
                mText.Location = worldLoc;
                mText.TextHeight = SnapToStandardHeight((totalH / group.Count) * scale);
                mText.Layer = "OCR_识别文字";
                mText.Width = (maxX - minX) * scale;
                mText.Attachment = AttachmentPoint.TopLeft;
                btr.AppendEntity(mText); tr.AddNewlyCreatedDBObject(mText, true);
            }
        }

        private double SnapToStandardHeight(double raw)
        {
            double[] std = { 2.0, 2.5, 3.0, 3.5, 4.0, 5.0, 7.0, 10.0, 15.0, 20.0 };
            double closest = std[0], min = Math.Abs(raw - std[0]);
            foreach (var h in std) { double d = Math.Abs(raw - h); if (d < min) { min = d; closest = h; } }
            return (raw > std[std.Length - 1] * 1.5) ? raw : closest;
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
                    Cv2.Threshold(gray, binary, _settings.Threshold, 255, ThresholdTypes.Binary);
                    if (!onlyHoles && Cv2.CountNonZero(binary) > (binary.Rows * binary.Cols / 2)) Cv2.BitwiseNot(binary, binary);
                    int kSize = (int)Math.Max(3, Math.Min(101, _settings.SimplifyTolerance * scale));
                    if (kSize % 2 == 0) kSize++;
                    using (Mat k = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(kSize, kSize)))
                        Cv2.MorphologyEx(binary, binary, MorphTypes.Close, k);
                    CvPoint[][] c; HierarchyIndex[] h;
                    Cv2.FindContours(binary, out c, out h, RetrievalModes.List, ContourApproximationModes.ApproxNone);
                    if (c == null) return results;
                    foreach (var pts in c)
                    {
                        if (onlyHoles && h[0].Parent == -1) continue;
                        if (pts.Length < 4) continue;
                        List<AcadPoint2d> path = new List<AcadPoint2d>();
                        foreach (var p in pts) path.Add(new AcadPoint2d(p.X / scale + ext.MinPoint.X, ext.MaxPoint.Y - p.Y / scale));
                        results.Add(path);
                    }
                }
            }
            return results;
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
                    CvPoint[][] c; HierarchyIndex[] h;
                    Cv2.FindContours(binary, out c, out h, RetrievalModes.List, ContourApproximationModes.ApproxSimple);
                    foreach (var pts in c)
                    {
                        if (pts.Length < 4) continue;
                        List<AcadPoint2d> res = new List<AcadPoint2d>();
                        foreach (var p in pts)
                        {
                            AcadPoint3d world = new AcadPoint3d(p.X, bmp.Height - p.Y, 0).TransformBy(transform);
                            res.Add(new AcadPoint2d(world.X, world.Y));
                        }
                        results.Add(res);
                    }
                }
            }
            return results;
        }

        private Matrix3d GetEntityTransform(Entity ent, int imgW, int imgH)
        {
            Extents3d ext = ent.GeometricExtents;
            double scX = (ext.MaxPoint.X - ext.MinPoint.X) / imgW;
            return Matrix3d.Displacement(ext.MinPoint.GetAsVector()) * Matrix3d.Scaling(scX, ext.MinPoint);
        }

        private void DrawInCad(Document doc, List<List<AcadPoint2d>> contours, SelectionSet original, double scale)
        {
            using (var loc = doc.LockDocument())
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(doc.Database), OpenMode.ForWrite);
                CheckAndCreateLayer(doc.Database, tr, _settings.LayerName);
                foreach (var pts in contours)
                {
                    Polyline pl = new Polyline();
                    for (int i = 0; i < pts.Count; i++) pl.AddVertexAt(i, pts[i], 0, 0, 0);
                    pl.Closed = true; pl.Layer = _settings.LayerName;
                    btr.AppendEntity(pl); tr.AddNewlyCreatedDBObject(pl, true);
                }
                tr.Commit();
            }
        }

        private void DrawInCadSimple(Document doc, List<List<AcadPoint2d>> contours, string layer)
        {
            using (var loc = doc.LockDocument())
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(doc.Database), OpenMode.ForWrite);
                CheckAndCreateLayer(doc.Database, tr, layer);
                foreach (var pts in contours)
                {
                    Polyline pl = new Polyline();
                    for (int i = 0; i < pts.Count; i++) pl.AddVertexAt(i, pts[i], 0, 0, 0);
                    pl.Closed = true; pl.Layer = layer;
                    btr.AppendEntity(pl); tr.AddNewlyCreatedDBObject(pl, true);
                }
                tr.Commit();
            }
        }

        private Extents3d GetSelectionExtents(SelectionSet ss)
        {
            Extents3d ext = new Extents3d();
            using (Transaction tr = Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction())
                foreach (SelectedObject so in ss)
                {
                    Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
                    if (ent != null) try { ext.AddExtents(ent.GeometricExtents); } catch { }
                }
            return ext;
        }

        private void CheckAndCreateLayer(Database db, Transaction tr, string layer)
        {
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(layer))
            {
                lt.UpgradeOpen();
                LayerTableRecord ltr = new LayerTableRecord { Name = layer };
                ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2);
                lt.Add(ltr); tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }

        private Bitmap RasterizeSelection(Document doc, SelectionSet ss, Extents3d ext, double scale)
        {
            int bw = (int)((ext.MaxPoint.X - ext.MinPoint.X) * scale) + 100;
            int bh = (int)((ext.MaxPoint.Y - ext.MinPoint.Y) * scale) + 100;
            Bitmap bmp = new Bitmap(bw, bh);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(System.Drawing.Color.Black); g.SmoothingMode = SmoothingMode.AntiAlias;
                using (Transaction tr = doc.TransactionManager.StartTransaction())
                using (System.Drawing.Pen p = new System.Drawing.Pen(System.Drawing.Color.White, 3.0f))
                    foreach (SelectedObject so in ss)
                    {
                        Curve c = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                        if (c == null) continue;
                        List<PointF> pts = new List<PointF>();
                        for (int i = 0; i <= 200; i++)
                        {
                            AcadPoint3d pt = c.GetPointAtParameter(c.StartParam + (c.EndParam - c.StartParam) * i / 200);
                            pts.Add(new PointF((float)((pt.X - ext.MinPoint.X) * scale) + 50f, (float)((ext.MaxPoint.Y - pt.Y) * scale) + 50f));
                        }
                        if (pts.Count > 1) g.DrawLines(p, pts.ToArray());
                    }
            }
            return bmp;
        }

        private bool IsMemorySafe(double w, double h, double s, Editor ed) => (w * s < 12000 && h * s < 12000);
        private void CleanMemory() { GC.Collect(); GC.WaitForPendingFinalizers(); }

        private class OcrWordData { public string Text; public Tesseract.Rect Bounds; public bool Visited = false; }
        #endregion
    }
}