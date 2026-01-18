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
using System.Diagnostics;
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
// 明确指定使用 System.Drawing 的 ImageFormat
using DrawingImageFormat = System.Drawing.Imaging.ImageFormat;

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
            Stopwatch sw = new Stopwatch();

            try
            {
                PromptSelectionResult psr = ed.GetSelection(new PromptSelectionOptions { MessageForAdding = "\n请选择要生成轮廓的元素: " });
                if (psr.Status != PromptStatus.OK) return;

                sw.Start();
                ed.WriteMessage($"\n[ContourMaster] 开始处理，已选择 {psr.Value.Count} 个元素...");

                Extents3d totalExt = GetSelectionExtents(psr.Value);

                // 逻辑转换
                int targetSize = _settings.PrecisionLevel * 1000;
                double worldW = totalExt.MaxPoint.X - totalExt.MinPoint.X;
                double worldH = totalExt.MaxPoint.Y - totalExt.MinPoint.Y;

                if (worldW <= 0 || worldH <= 0) return;
                double scale = targetSize / Math.Max(worldW, worldH);

                // 内存安全检查
                if (!IsMemorySafe(worldW, worldH, scale, ed)) return;

                ed.WriteMessage("\n[1/4] 正在将矢量元素光栅化为位图...");
                using (Bitmap bmp = RasterizeSelection(doc, psr.Value, totalExt, scale))
                {
                    if (bmp == null) return;
                    ed.WriteMessage($"\n -> 位图创建成功: {bmp.Width}x{bmp.Height}");

                    ed.WriteMessage("\n[2/4] 正在使用 OpenCV 提取图像轮廓...");
                    List<List<AcadPoint2d>> cadContours = ExtractCadContoursWithOpenCv(bmp, totalExt, scale, ed);

                    if (cadContours.Count == 0)
                    {
                        ed.WriteMessage("\n[未发现轮廓] 请确保选中的线能够构成闭合区域。建议调低“识别阈值”或增加“像素精度”。");
                        return;
                    }

                    ed.WriteMessage($"\n[3/4] 提取完成，共识别到 {cadContours.Count} 个闭合区域。");
                    ed.WriteMessage("\n[4/4] 正在生成 AutoCAD 多段线并执行几何吸附...");

                    DrawInCad(doc, cadContours, psr.Value, scale, ed);
                }

                sw.Stop();
                ed.WriteMessage($"\n[ContourMaster] 处理完毕！总耗时: {sw.ElapsedMilliseconds} ms。");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[致命异常] 运算中止: {ex.Message}\n{ex.StackTrace}");
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        #endregion

        #region 2. 图像参照原位识别逻辑
        public async Task ProcessReferencedImage()
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

                    // 必须在 using 块内 await，保证 OCR 完成前 bmp 不被销毁
                    using (Bitmap bmp = new Bitmap(imgDef.ActiveFileName))
                    {
                        Matrix3d pixelToWorld = GetEntityTransform(img, bmp.Width, bmp.Height);
                        if (_settings.IsOcrMode)
                        {
                            await ProcessImageOcr(bmp, pixelToWorld); // 关键：增加 await
                        }
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
        private async Task ProcessImageOcr(Bitmap bmp, Matrix3d transform)
        {
            if (_settings.SelectedOcrEngine == OcrEngineType.Baidu)
                await ProcessBaiduOcr(bmp, transform);
            else
                ProcessTesseractOcr(bmp, transform);
        }

        private async Task ProcessBaiduOcr(Bitmap bmp, Matrix3d transform)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            try
            {
                ed.WriteMessage("\n[调试] 开始百度 OCR 流程...");

                string token = await GetBaiduAccessToken(_settings.BaiduApiKey, _settings.BaiduSecretKey);
                if (string.IsNullOrEmpty(token)) throw new Exception("Token 获取失败。");

                // 1. 计算缩放比例
                Bitmap processedBmp = bmp;
                double scaleRatio = 1.0; // 还原比例：原图 / 缩放图
                int maxSide = Math.Max(bmp.Width, bmp.Height);

                if (maxSide > 4000)
                {
                    scaleRatio = (double)maxSide / 4000.0; // 记录缩放倍数
                    int newW = (int)(bmp.Width / scaleRatio);
                    int newH = (int)(bmp.Height / scaleRatio);
                    processedBmp = new Bitmap(bmp, new System.Drawing.Size(newW, newH));
                    ed.WriteMessage($"\n[调试] 缩放比例: {scaleRatio:F2}, 识别尺寸: {newW}x{newH}");
                }

                string base64Image;
                using (MemoryStream ms = new MemoryStream())
                {
                    var jpgEncoder = GetEncoder(System.Drawing.Imaging.ImageFormat.Jpeg);
                    var encoderParams = new EncoderParameters(1);
                    encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 80L);
                    processedBmp.Save(ms, jpgEncoder, encoderParams);
                    base64Image = Convert.ToBase64String(ms.ToArray());
                }

                string postBody = "image=" + System.Net.WebUtility.UrlEncode(base64Image) + "&probability=true";
                var content = new StringContent(postBody, System.Text.Encoding.UTF8, "application/x-www-form-urlencoded");

                var response = await _httpClient.PostAsync($"https://aip.baidubce.com/rest/2.0/ocr/v1/accurate?access_token={token}", content);
                string json = await response.Content.ReadAsStringAsync();

                var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                var result = serializer.Deserialize<Dictionary<string, object>>(json);

                if (result.ContainsKey("error_code"))
                    throw new Exception($"{result["error_msg"]} (代码:{result["error_code"]})");

                if (result.ContainsKey("words_result") && result["words_result"] is System.Collections.IEnumerable wordsList)
                {
                    using (var loc = doc.LockDocument())
                    using (Transaction tr = doc.TransactionManager.StartTransaction())
                    {
                        var btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(doc.Database), OpenMode.ForWrite);
                        CheckAndCreateLayer(doc.Database, tr, "OCR_识别文字");

                        // 矩阵本身的单位比例
                        double matrixScale = new Vector3d(transform[0, 0], transform[1, 0], transform[2, 0]).Length;

                        foreach (object itemObj in wordsList)
                        {
                            var item = itemObj as Dictionary<string, object>;
                            if (item == null) continue;

                            string text = item["words"].ToString();
                            var locData = (Dictionary<string, object>)item["location"];

                            // 2. ✨ 核心修正：将 OCR 返回的坐标还原到原图尺寸
                            double rawLeft = Convert.ToDouble(locData["left"]) * scaleRatio;
                            double rawTop = Convert.ToDouble(locData["top"]) * scaleRatio;
                            double rawWidth = Convert.ToDouble(locData["width"]) * scaleRatio;
                            double rawHeight = Convert.ToDouble(locData["height"]) * scaleRatio;

                            // 3. 计算 CAD 世界坐标 (使用原图高度 bmp.Height 进行 Y 翻转)
                            AcadPoint3d worldLoc = new AcadPoint3d(rawLeft, bmp.Height - rawTop, 0).TransformBy(transform);

                            using (MText mText = new MText())
                            {
                                mText.Contents = text;
                                mText.Location = worldLoc;
                                // 字高和宽度也需要乘回 scaleRatio 才能匹配原始比例
                                mText.TextHeight = SnapToStandardHeight(rawHeight * matrixScale);
                                mText.Width = rawWidth * matrixScale;

                                mText.Layer = "OCR_识别文字";
                                mText.Attachment = AttachmentPoint.TopLeft; // 对应识别结果的左上角

                                btr.AppendEntity(mText);
                                tr.AddNewlyCreatedDBObject(mText, true);
                            }
                        }
                        tr.Commit();
                        ed.WriteMessage("\n[Baidu OCR] 坐标修正识别完成。");
                    }
                }

                if (processedBmp != bmp) processedBmp.Dispose();
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\n[百度OCR错误] {ex.Message}");
            }
        }

        // 辅助方法：获取图像编码器
        private ImageCodecInfo GetEncoder(DrawingImageFormat format) // 使用别名
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid) return codec;
            }
            return null;
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

        private List<List<AcadPoint2d>> ExtractCadContoursWithOpenCv(Bitmap bmp, Extents3d ext, double scale, Editor ed)
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

                    int kSize = (int)Math.Ceiling(_settings.SimplifyTolerance * scale * 1.2);
                    if (kSize < 3) kSize = 3;
                    if (kSize % 2 == 0) kSize++;
                    kSize = Math.Min(kSize, 101);

                    ed.WriteMessage($"\n -> 图像算法参数: 阈值={thresholdValue}, 闭运算核={kSize}");

                    using (Mat kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(kSize, kSize)))
                    {
                        Cv2.MorphologyEx(binary, binary, MorphTypes.Close, kernel);
                    }

                    CvPoint[][] contours;
                    HierarchyIndex[] hierarchy;
                    Cv2.FindContours(binary, out contours, out hierarchy, RetrievalModes.Tree, ContourApproximationModes.ApproxSimple);

                    if (contours == null) return results;

                    for (int i = 0; i < contours.Length; i++)
                    {
                        // 只处理有父级的轮廓（内孔逻辑）
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
                        results.Add(cadPath);
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
        private AcadPoint2d SnapPointToOriginalCurves(AcadPoint2d pt, SelectionSet originalSs, Transaction tr, double pixelSize)
        {
            AcadPoint3d pt3d = new AcadPoint3d(pt.X, pt.Y, 0);
            // 吸附阈值设定：通常取补缝距离的 1.5 倍
            double snapThreshold = Math.Max(pixelSize * 6.0, _settings.SimplifyTolerance * 1.5);
            double searchThreshold = snapThreshold + (pixelSize * 2.0);

            List<Curve> nearbyCurves = new List<Curve>();
            List<AcadPoint3d> criticalPoints = new List<AcadPoint3d>();

            foreach (SelectedObject so in originalSs)
            {
                try
                {
                    Curve curve = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                    if (curve == null) continue;

                    // 过滤距离太远的曲线，减少计算量
                    Extents3d ext = curve.GeometricExtents;
                    if (pt.X < ext.MinPoint.X - searchThreshold || pt.X > ext.MaxPoint.X + searchThreshold ||
                        pt.Y < ext.MinPoint.Y - searchThreshold || pt.Y > ext.MaxPoint.Y + searchThreshold)
                        continue;

                    nearbyCurves.Add(curve);
                    criticalPoints.Add(curve.StartPoint);
                    criticalPoints.Add(curve.EndPoint);
                }
                catch { continue; }
            }

            AcadPoint3d bestSnapPt = pt3d;
            double minCriticalDist = snapThreshold;
            bool foundCritical = false;

            // 1. 端点吸附优先
            foreach (AcadPoint3d cp in criticalPoints)
            {
                double d = pt3d.DistanceTo(cp);
                if (d < minCriticalDist) { minCriticalDist = d; bestSnapPt = cp; foundCritical = true; }
            }

            // 2. 交点吸附（增加异常保护，解决 eInvalidInput）
            if (!foundCritical && nearbyCurves.Count >= 2 && nearbyCurves.Count < 50)
            {
                for (int i = 0; i < nearbyCurves.Count; i++)
                {
                    for (int j = i + 1; j < nearbyCurves.Count; j++)
                    {
                        try
                        {
                            using (Point3dCollection interPts = new Point3dCollection())
                            {
                                // 核心风险点：必须 catch 异常
                                nearbyCurves[i].IntersectWith(nearbyCurves[j], Intersect.OnBothOperands, interPts, IntPtr.Zero, IntPtr.Zero);
                                foreach (AcadPoint3d ip in interPts)
                                {
                                    double d = pt3d.DistanceTo(ip);
                                    if (d < minCriticalDist) { minCriticalDist = d; bestSnapPt = ip; foundCritical = true; }
                                }
                            }
                        }
                        catch { continue; }
                    }
                }
            }

            if (foundCritical) return new AcadPoint2d(bestSnapPt.X, bestSnapPt.Y);

            // 3. 曲线最近点吸附
            double minLineDist = snapThreshold;
            foreach (Curve curve in nearbyCurves)
            {
                try
                {
                    AcadPoint3d closest = curve.GetClosestPointTo(pt3d, false);
                    double d = pt3d.DistanceTo(closest);
                    if (d < minLineDist) { minLineDist = d; bestSnapPt = closest; }
                }
                catch { continue; }
            }
            return new AcadPoint2d(bestSnapPt.X, bestSnapPt.Y);
        }
        private void DrawInCad(Document doc, List<List<AcadPoint2d>> contours, SelectionSet originalLines, double scale, Editor ed)
        {
            double pixelSize = 1.0 / scale;
            int count = 0;

            using (var loc = doc.LockDocument())
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(doc.Database), OpenMode.ForWrite);
                string layerName = _settings.LayerName ?? "LK_XS";
                CheckAndCreateLayer(doc.Database, tr, layerName);

                foreach (var pts in contours)
                {
                    count++;
                    if (count % 10 == 0) // 每处理10个区域更新一次进度
                    {
                        ed.WriteMessage($"\n -> 正在处理第 {count}/{contours.Count} 个区域...");
                    }

                    Polyline pl = new Polyline();
                    int vertexIndex = 0;
                    AcadPoint2d lastAddedPt = new AcadPoint2d(double.NaN, double.NaN);

                    for (int i = 0; i < pts.Count; i++)
                    {
                        // 吸附逻辑
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
                    btr.AppendEntity(pl);
                    tr.AddNewlyCreatedDBObject(pl, true);
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
            // 根据缩放比例计算位图尺寸，并在四周预留 50 像素（共 100 像素）的边距防止边缘裁剪
            int bmpW = (int)((ext.MaxPoint.X - ext.MinPoint.X) * scale) + 100;
            int bmpH = (int)((ext.MaxPoint.Y - ext.MinPoint.Y) * scale) + 100;

            // 内存安全检查：限制位图最大尺寸，防止超大选区导致内存溢出
            if (bmpW <= 0 || bmpH <= 0 || bmpW > 15000 || bmpH > 15000) return null;

            Bitmap bmp = new Bitmap(bmpW, bmpH);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(System.Drawing.Color.Black);
                g.SmoothingMode = SmoothingMode.AntiAlias;

                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    // 使用白色 3.0f 宽度的画笔。
                    // 提示：线条必须有厚度，OpenCV 才能准确识别出轮廓的内外边界。
                    using (System.Drawing.Pen thinPen = new System.Drawing.Pen(System.Drawing.Color.White, 3.0f))
                    {
                        foreach (SelectedObject so in ss)
                        {
                            try
                            {
                                Curve curve = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                                if (curve == null) continue;

                                double startParam = curve.StartParam;
                                double endParam = curve.EndParam;

                                // 忽略零长度或退化的曲线
                                if (Math.Abs(endParam - startParam) < 1e-7) continue;

                                int segments = 100; // 采样频率。对于 800+ 元素，100 次采样是平衡速度与精度的较好选择
                                List<PointF> pts = new List<PointF>();
                                double step = (endParam - startParam) / segments;

                                for (int i = 0; i <= segments; i++)
                                {
                                    try
                                    {
                                        double t = startParam + (step * i);

                                        // 严格约束参数范围，防止浮点数精度误差导致 t 略微超出边界而触发 eInvalidInput
                                        if (t < startParam) t = startParam;
                                        if (t > endParam) t = endParam;

                                        AcadPoint3d p = curve.GetPointAtParameter(t);

                                        // 坐标转换逻辑：
                                        // 1. (p.X - ext.MinPoint.X) * scale: 将 CAD 坐标平移并缩放
                                        // 2. (ext.MaxPoint.Y - p.Y) * scale: CAD Y轴向上，位图 Y轴向下，需进行翻转
                                        // 3. + 50f: 加上预留的边距偏移量
                                        float px = (float)((p.X - ext.MinPoint.X) * scale) + 50f;
                                        float py = (float)((ext.MaxPoint.Y - p.Y) * scale) + 50f;
                                        pts.Add(new PointF(px, py));
                                    }
                                    catch
                                    {
                                        // 捕获单个采样点的几何运算异常，允许跳过该点以保证多段线继续绘制
                                        continue;
                                    }
                                }

                                // 仅当成功获取到至少两个点时才进行绘制
                                if (pts.Count > 1)
                                {
                                    g.DrawLines(thinPen, pts.ToArray());
                                }
                            }
                            catch
                            {
                                // 捕获无法识别或已损坏实体的异常，确保不中断后续实体的处理
                                continue;
                            }
                        }
                    }
                    tr.Commit();
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