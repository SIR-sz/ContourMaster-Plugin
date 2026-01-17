using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Plugin_ContourMaster.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

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
        private List<Point2d> SimplifyCadPath(List<Point2d> pts, double tol)
        {
            if (pts.Count < 3) return pts;
            int idx = -1;
            double maxD = 0;
            for (int i = 1; i < pts.Count - 1; i++)
            {
                double d = GetDistancePointLine(pts[i], pts[0], pts[pts.Count - 1]);
                if (d > maxD) { maxD = d; idx = i; }
            }
            if (maxD > tol)
            {
                var l = SimplifyCadPath(pts.GetRange(0, idx + 1), tol);
                var r = SimplifyCadPath(pts.GetRange(idx, pts.Count - idx), tol);
                l.RemoveAt(l.Count - 1);
                l.AddRange(r);
                return l;
            }
            return new List<Point2d> { pts[0], pts[pts.Count - 1] };
        }
        private void CheckAndCreateLayer(Database db, Transaction tr, string layerName)
        {
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                LayerTableRecord ltr = new LayerTableRecord { Name = layerName };
                // 显式引用 AutoCAD 颜色，防止与 System.Drawing 冲突
                ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2); // 黄色
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }

        // 配合 SimplifyCadPath 使用的距离计算函数
        private double GetDistancePointLine(Point2d p, Point2d s, Point2d e)
        {
            double area = Math.Abs(0.5 * (s.X * (e.Y - p.Y) + e.X * (p.Y - s.Y) + p.X * (s.Y - e.Y)));
            double bottom = s.GetDistanceTo(e);
            if (bottom < 1e-9) return p.GetDistanceTo(s);
            return (area * 2.0) / bottom;
        }

        // 修改自
        public void ProcessContour()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // 1. 选择与范围获取 (保持原有逻辑)
                PromptSelectionResult psr = ed.GetSelection(new PromptSelectionOptions { MessageForAdding = "\n请选择元素: " });
                if (psr.Status != PromptStatus.OK) return;

                Extents3d totalExt = GetSelectionExtents(psr.Value);

                // 2. 预计算变换参数 (零偏差映射)
                int targetSize = 3000;
                double worldW = totalExt.MaxPoint.X - totalExt.MinPoint.X;
                double worldH = totalExt.MaxPoint.Y - totalExt.MinPoint.Y;
                double scale = targetSize / Math.Max(worldW, worldH);

                using (Bitmap bmp = RasterizeSelection(doc, psr.Value, totalExt, scale))
                {
                    if (bmp == null) return;

                    int gapRadius = (int)(_settings.SimplifyTolerance * scale);
                    if (gapRadius < 1) gapRadius = 1;

                    // 3. 图像处理提取孔洞
                    bool[,] holeMap = IdentifyEnclosedHolesWithMorphology(bmp, gapRadius);

                    // ✨ 核心改变：直接在提取阶段生成准确的 CAD 坐标点
                    // 传入 scale 和 totalExt，让生成过程“心中有坐标”
                    List<List<Point2d>> cadContours = ExtractCadContours(holeMap, totalExt, scale);

                    // 4. 直接绘制生成的 CAD 点位
                    DrawInCad(doc, cadContours);
                }

                ed.WriteMessage("\n[像素轮廓专家] 已按准确坐标直接生成轮廓。");
            }
            catch (Exception ex) { ed.WriteMessage($"\n[异常] {ex.Message}"); }
        }

        /// <summary>
        /// 将选中的 CAD 实体渲染为位图，用于后续的像素轮廓提取。
        /// 此版本已移除 out scale 关键字，改由外部传入，确保全流程坐标一致。
        /// </summary>
        private Bitmap RasterizeSelection(Document doc, SelectionSet ss, Extents3d ext, double scale)
        {
            // 计算实体的世界坐标宽度和高度
            double worldW = ext.MaxPoint.X - ext.MinPoint.X;
            double worldH = ext.MaxPoint.Y - ext.MinPoint.Y;

            // 根据 scale 计算位图尺寸，并预留双边共 100 像素的边距 (左右各50，上下各50)
            int bmpW = (int)(worldW * scale) + 100;
            int bmpH = (int)(worldH * scale) + 100;

            // 内存安全检查：防止因图形过大或 scale 过高导致创建超大位图
            if (bmpW > 10000) bmpW = 10000;
            if (bmpH > 10000) bmpH = 10000;
            if (bmpW < 100) bmpW = 100; // 确保最小尺寸
            if (bmpH < 100) bmpH = 100;

            // 创建位图
            Bitmap bmp = new Bitmap(bmpW, bmpH);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                // 关键设置：关闭抗锯齿，确保提取轮廓时像素边缘清晰，不产生模糊的中间色
                g.Clear(System.Drawing.Color.Black);
                g.SmoothingMode = SmoothingMode.None;

                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    // 使用 1 像素宽度的白色画笔绘制实体
                    using (System.Drawing.Pen thinPen = new System.Drawing.Pen(System.Drawing.Color.White, 1f))
                    {
                        foreach (SelectedObject so in ss)
                        {
                            // 只处理曲线实体 (Polyline, Line, Circle, Arc 等)
                            Curve curve = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                            if (curve == null) continue;

                            try
                            {
                                // 将曲线离散化为 150 个线段进行绘制
                                int segments = 150;
                                PointF[] pts = new PointF[segments + 1];
                                double curveStart = curve.StartParam;
                                double curveEnd = curve.EndParam;
                                double step = (curveEnd - curveStart) / segments;

                                for (int i = 0; i <= segments; i++)
                                {
                                    // 获取曲线上的点
                                    Point3d p = curve.GetPointAtParameter(curveStart + step * i);

                                    // 核心坐标变换逻辑：
                                    // 1. (p.X - ext.MinPoint.X) * scale: 将 CAD 坐标平移到 0,0 并缩放至像素空间
                                    // 2. + 50: 加上预留的 50 像素左/上边距
                                    // 3. Y 轴需要翻转，因为 CAD Y 轴向上，位图 Y 轴向下
                                    float px = (float)((p.X - ext.MinPoint.X) * scale) + 50f;
                                    float py = (float)((ext.MaxPoint.Y - p.Y) * scale) + 50f;

                                    pts[i] = new PointF(px, py);
                                }

                                // 绘制离散后的路径
                                if (pts.Length > 1)
                                {
                                    g.DrawLines(thinPen, pts);
                                }
                            }
                            catch
                            {
                                // 忽略无法获取参数的特定实体类型，防止算法中断
                            }
                        }
                    }
                    tr.Commit();
                }
            }
            return bmp;
        }

        // 优化 中的形态学逻辑
        private bool[,] IdentifyEnclosedHolesWithMorphology(Bitmap bmp, int radius)
        {
            int w = bmp.Width, h = bmp.Height;
            bool[,] isLine = ExtractLineMap(bmp); // 提取线条

            // 1. 膨胀：封死缺口
            bool[,] dilated = Dilate(isLine, radius);

            // 2. 种子填充识别外部区域 (代码保持原逻辑)
            bool[,] isOutside = FloodFillOutside(dilated);

            // 3. 提取内部孔洞区域
            bool[,] holeMap = new bool[w, h];
            for (int y = 0; y < h; y++)
                for (int x = 0; x < w; x++)
                    holeMap[x, y] = !dilated[x, y] && !isOutside[x, y];

            // ✨ 核心改进：使用 radius + 1 或 radius + 2
            // 增加 1-2 像素的额外腐蚀，彻底抵消平滑带来的内缩感
            int compensation = (_settings.SmoothLevel > 5) ? radius + 2 : radius + 1;
            return ErodeHole(holeMap, compensation);
        }
        // 补全：提取像素位图到布尔数组
        private bool[,] ExtractLineMap(Bitmap bmp)
        {
            int w = bmp.Width, h = bmp.Height;
            bool[,] isLine = new bool[w, h];
            BitmapData data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            byte[] pixels = new byte[data.Stride * h];
            Marshal.Copy(data.Scan0, pixels, 0, pixels.Length);
            bmp.UnlockBits(data);

            for (int y = 0; y < h; y++)
            {
                for (int x = 0; x < w; x++)
                {
                    // 使用设置中的阈值识别线条
                    isLine[x, y] = pixels[y * data.Stride + x * 4 + 2] > _settings.Threshold;
                }
            }
            return isLine;
        }

        // 补全：泛洪算法识别外部空间
        private bool[,] FloodFillOutside(bool[,] map)
        {
            int w = map.GetLength(0), h = map.GetLength(1);
            bool[,] isOutside = new bool[w, h];
            Queue<System.Drawing.Point> q = new Queue<System.Drawing.Point>();

            // 从图像四周边界注入种子
            for (int x = 0; x < w; x++) { q.Enqueue(new System.Drawing.Point(x, 0)); q.Enqueue(new System.Drawing.Point(x, h - 1)); }
            for (int y = 0; y < h; y++) { q.Enqueue(new System.Drawing.Point(0, y)); q.Enqueue(new System.Drawing.Point(w - 1, y)); }

            while (q.Count > 0)
            {
                var p = q.Dequeue();
                if (p.X < 0 || p.X >= w || p.Y < 0 || p.Y >= h || isOutside[p.X, p.Y] || map[p.X, p.Y]) continue;
                isOutside[p.X, p.Y] = true;
                q.Enqueue(new System.Drawing.Point(p.X + 1, p.Y)); q.Enqueue(new System.Drawing.Point(p.X - 1, p.Y));
                q.Enqueue(new System.Drawing.Point(p.X, p.Y + 1)); q.Enqueue(new System.Drawing.Point(p.X, p.Y - 1));
            }
            return isOutside;
        }
        private bool[,] Dilate(bool[,] source, int r)
        {
            int w = source.GetLength(0), h = source.GetLength(1);
            bool[,] dest = (bool[,])source.Clone();
            for (int y = r; y < h - r; y++)
                for (int x = r; x < w - r; x++)
                    if (source[x, y])
                        for (int ky = -r; ky <= r; ky++)
                            for (int kx = -r; kx <= r; kx++)
                                dest[x + kx, y + ky] = true;
            return dest;
        }

        private bool[,] ErodeHole(bool[,] hole, int r)
        {
            int w = hole.GetLength(0), h = hole.GetLength(1);
            bool[,] dest = (bool[,])hole.Clone();
            for (int y = r; y < h - r; y++)
                for (int x = r; x < w - r; x++)
                    if (!hole[x, y])
                        for (int ky = -r; ky <= r; ky++)
                            for (int kx = -r; kx <= r; kx++)
                                if (hole[x + kx, y + ky])
                                {
                                    dest[x, y] = true;
                                    goto NextPixel;
                                }
                            NextPixel:;
            return dest;
        }
        // ✨ 新增：拉普拉斯平滑，消除像素点的局部抖动
        private List<PointF> SmoothPoints(List<PointF> pts)
        {
            if (_settings.SmoothLevel <= 0 || pts.Count < 5) return pts;

            List<PointF> smoothed = new List<PointF>(pts.Count);
            smoothed.Add(pts[0]);

            // 减弱平滑强度，只对局部微小锯齿生效
            // 使用 0.2/0.6/0.2 的加权平均，比单纯的 1/3 平均更不容易收缩
            for (int i = 1; i < pts.Count - 1; i++)
            {
                float sx = pts[i - 1].X * 0.2f + pts[i].X * 0.6f + pts[i + 1].X * 0.2f;
                float sy = pts[i - 1].Y * 0.2f + pts[i].Y * 0.6f + pts[i + 1].Y * 0.2f;
                smoothed.Add(new PointF(sx, sy));
            }

            smoothed.Add(pts[pts.Count - 1]);
            return smoothed;
        }

        private List<List<Point2d>> ExtractCadContours(bool[,] map, Extents3d ext, double scale)
        {
            int w = map.GetLength(0), h = map.GetLength(1);
            List<List<Point2d>> results = new List<List<Point2d>>();
            bool[,] visited = new bool[w, h];

            // 基于 SmoothLevel 动态调整平滑容差
            double simplifyTol = 0.5 + (_settings.SmoothLevel * 0.5);

            for (int y = 1; y < h - 1; y++)
            {
                for (int x = 1; x < w - 1; x++)
                {
                    // 寻找边界像素
                    if (map[x, y] && !visited[x, y] && (!map[x - 1, y] || !map[x + 1, y] || !map[x, y - 1] || !map[x, y + 1]))
                    {
                        // 1. 像素空间追踪
                        var pixelPath = Trace(map, x, y, visited);
                        if (pixelPath.Count < 10) continue;

                        // 2. 路径预平滑 (防止锯齿)
                        var smoothedPixelPath = SmoothPoints(pixelPath);

                        // 3. ✨ 关键步骤：在简化/生成阶段直接计算 CAD 坐标
                        List<Point2d> cadPath = new List<Point2d>();
                        foreach (var p in smoothedPixelPath)
                        {
                            // 零偏差公式：(像素坐标 - 边距) / 缩放 + 原始基准点
                            double wx = (p.X - 50.0) / scale + ext.MinPoint.X;
                            double wy = ext.MaxPoint.Y - (p.Y - 50.0) / scale;
                            cadPath.Add(new Point2d(wx, wy));
                        }

                        // 4. 矢量简化 (Douglas-Peucker 算法，此时 tol 的单位已是 CAD 单位)
                        results.Add(SimplifyCadPath(cadPath, simplifyTol / scale));
                    }
                }
            }
            return results;
        }

        private List<PointF> Trace(bool[,] map, int sx, int sy, bool[,] v)
        {
            List<PointF> path = new List<PointF>();
            int[] dx = { -1, -1, 0, 1, 1, 1, 0, -1 }, dy = { 0, -1, -1, -1, 0, 1, 1, 1 };
            int cx = sx, cy = sy, bx = sx - 1, by = sy;
            do
            {
                path.Add(new PointF(cx, cy)); v[cx, cy] = true;
                int si = 0;
                for (int i = 0; i < 8; i++) if (cx + dx[i] == bx && cy + dy[i] == by) { si = i; break; }
                bool found = false;
                for (int i = 1; i <= 8; i++)
                {
                    int idx = (si + i) % 8, nx = cx + dx[idx], ny = cy + dy[idx];
                    if (nx >= 0 && nx < map.GetLength(0) && ny >= 0 && ny < map.GetLength(1) && map[nx, ny])
                    {
                        bx = cx; by = cy; cx = nx; cy = ny; found = true; break;
                    }
                }
                if (!found) break;
            } while ((cx != sx || cy != sy) && path.Count < 30000);
            return path;
        }

        private List<PointF> Simplify(List<PointF> pts, double tol)
        {
            if (pts.Count < 3) return pts;
            int idx = -1; double maxD = 0;
            for (int i = 1; i < pts.Count - 1; i++)
            {
                double d = GetDistance(pts[i], pts[0], pts[pts.Count - 1]);
                if (d > maxD) { maxD = d; idx = i; }
            }
            if (maxD > tol)
            {
                var l = Simplify(pts.GetRange(0, idx + 1), tol); var r = Simplify(pts.GetRange(idx, pts.Count - idx), tol);
                l.RemoveAt(l.Count - 1); l.AddRange(r); return l;
            }
            return new List<PointF> { pts[0], pts[pts.Count - 1] };
        }

        private double GetDistance(PointF p, PointF s, PointF e)
        {
            double a = Math.Abs(0.5 * (s.X * (e.Y - p.Y) + e.X * (p.Y - s.Y) + p.X * (s.Y - e.Y)));
            double b = Math.Sqrt(Math.Pow(s.X - e.X, 2) + Math.Pow(s.Y - e.Y, 2));
            return b == 0 ? 0 : a / (0.5 * b);
        }

        private void DrawInCad(Document doc, List<List<Point2d>> contours)
        {
            using (var loc = doc.LockDocument())
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                Database db = doc.Database;
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                string layerName = _settings.LayerName ?? "LK_XS";
                CheckAndCreateLayer(db, tr, layerName);

                foreach (var pts in contours)
                {
                    if (pts.Count < 2) continue;

                    Polyline pl = new Polyline();
                    for (int i = 0; i < pts.Count; i++)
                    {
                        // 无需再做坐标转换，直接使用传入的 Point2d
                        pl.AddVertexAt(i, pts[i], 0, 0, 0);
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
    }
}