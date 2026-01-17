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

        public void ProcessContour()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                PromptSelectionOptions pso = new PromptSelectionOptions();
                pso.MessageForAdding = "\n请选择要生成轮廓的元素: ";
                PromptSelectionResult psr = ed.GetSelection(pso);
                if (psr.Status != PromptStatus.OK) return;

                Extents3d totalExt = new Extents3d();
                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject so in psr.Value)
                    {
                        Entity ent = (Entity)tr.GetObject(so.ObjectId, OpenMode.ForRead);
                        try { totalExt.AddExtents(ent.GeometricExtents); } catch { }
                    }
                    tr.Commit();
                }

                // 1. 采用高分辨率 1 像素采样绘制
                using (Bitmap bmp = RasterizeSelection(doc, psr.Value, totalExt, out double scale))
                {
                    if (bmp == null) return;

                    // 计算膨胀/腐蚀半径（像素单位）
                    int gapRadius = (int)(_settings.SimplifyTolerance * scale);
                    if (gapRadius < 1) gapRadius = 1;

                    // 2. 核心：膨胀补缺 -> 识别孔洞 -> 腐蚀还原位置
                    bool[,] holeMap = IdentifyEnclosedHolesWithMorphology(bmp, gapRadius);

                    // 3. 追踪轮廓
                    List<List<PointF>> contours = FindContours(holeMap);

                    // 4. 还原绘制到 CAD
                    DrawInCad(doc, contours, totalExt, scale);
                }

                ed.WriteMessage("\n[像素轮廓专家] 处理完成：轮廓已精准贴合生成。");
            }
            catch (Exception ex) { ed.WriteMessage($"\n[算法异常] {ex.Message}"); }
        }

        private Bitmap RasterizeSelection(Document doc, SelectionSet ss, Extents3d ext, out double scale)
        {
            int targetSize = 3000; // 提高分辨率以确保精度
            double worldW = ext.MaxPoint.X - ext.MinPoint.X;
            double worldH = ext.MaxPoint.Y - ext.MinPoint.Y;
            scale = targetSize / Math.Max(worldW, worldH);

            int bmpW = (int)(worldW * scale) + 100;
            int bmpH = (int)(worldH * scale) + 100;

            Bitmap bmp = new Bitmap(bmpW, bmpH);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(System.Drawing.Color.Black);
                g.SmoothingMode = SmoothingMode.None; // 1像素采样不使用抗锯齿

                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    using (System.Drawing.Pen thinPen = new System.Drawing.Pen(System.Drawing.Color.White, 1f))
                    {
                        foreach (SelectedObject so in ss)
                        {
                            Curve curve = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                            if (curve == null) continue;

                            int segments = 150;
                            PointF[] pts = new PointF[segments + 1];
                            double step = (curve.EndParam - curve.StartParam) / segments;
                            for (int i = 0; i <= segments; i++)
                            {
                                Point3d p = curve.GetPointAtParameter(curve.StartParam + step * i);
                                pts[i] = new PointF((float)((p.X - ext.MinPoint.X) * scale) + 50, (float)((ext.MaxPoint.Y - p.Y) * scale) + 50);
                            }
                            g.DrawLines(thinPen, pts);
                        }
                    }
                    tr.Commit();
                }
            }
            return bmp;
        }

        private bool[,] IdentifyEnclosedHolesWithMorphology(Bitmap bmp, int radius)
        {
            int w = bmp.Width, h = bmp.Height;
            bool[,] isLine = new bool[w, h];

            BitmapData data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            byte[] pixels = new byte[data.Stride * h];
            Marshal.Copy(data.Scan0, pixels, 0, pixels.Length);
            bmp.UnlockBits(data);
            for (int y = 0; y < h; y++)
                for (int x = 0; x < w; x++)
                    isLine[x, y] = pixels[y * data.Stride + x * 4 + 2] > 128;

            // 1. 膨胀：封死缺口
            bool[,] dilated = Dilate(isLine, radius);

            // 2. 种子填充识别内部孔洞
            bool[,] isOutside = new bool[w, h];
            Queue<System.Drawing.Point> q = new Queue<System.Drawing.Point>();
            for (int x = 0; x < w; x++) { q.Enqueue(new System.Drawing.Point(x, 0)); q.Enqueue(new System.Drawing.Point(x, h - 1)); }
            for (int y = 0; y < h; y++) { q.Enqueue(new System.Drawing.Point(0, y)); q.Enqueue(new System.Drawing.Point(w - 1, y)); }

            while (q.Count > 0)
            {
                var p = q.Dequeue();
                if (p.X < 0 || p.X >= w || p.Y < 0 || p.Y >= h || isOutside[p.X, p.Y] || dilated[p.X, p.Y]) continue;
                isOutside[p.X, p.Y] = true;
                q.Enqueue(new System.Drawing.Point(p.X + 1, p.Y)); q.Enqueue(new System.Drawing.Point(p.X - 1, p.Y));
                q.Enqueue(new System.Drawing.Point(p.X, p.Y + 1)); q.Enqueue(new System.Drawing.Point(p.X, p.Y - 1));
            }

            bool[,] holeMap = new bool[w, h];
            for (int y = 0; y < h; y++)
                for (int x = 0; x < w; x++)
                    holeMap[x, y] = !dilated[x, y] && !isOutside[x, y];

            // 3. 腐蚀：将孔洞外扩，消除内缩，撞回 1 像素原始位置
            return ErodeHole(holeMap, radius);
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

        private List<List<PointF>> FindContours(bool[,] map)
        {
            int w = map.GetLength(0), h = map.GetLength(1);
            List<List<PointF>> result = new List<List<PointF>>();
            bool[,] visited = new bool[w, h];
            for (int y = 1; y < h - 1; y++)
                for (int x = 1; x < w - 1; x++)
                    if (map[x, y] && !visited[x, y] && (!map[x - 1, y] || !map[x + 1, y] || !map[x, y - 1] || !map[x, y + 1]))
                    {
                        var path = Trace(map, x, y, visited);
                        if (path.Count > 15) result.Add(Simplify(path, 0.4));
                    }
            return result;
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

        private void DrawInCad(Document doc, List<List<PointF>> contours, Extents3d ext, double scale)
        {
            using (var loc = doc.LockDocument())
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                Database db = doc.Database;
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                string layerName = "LK_XS";
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (!lt.Has(layerName))
                {
                    lt.UpgradeOpen();
                    LayerTableRecord ltr = new LayerTableRecord { Name = layerName };
                    ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2);
                    lt.Add(ltr); tr.AddNewlyCreatedDBObject(ltr, true);
                }

                foreach (var pts in contours)
                {
                    Polyline pl = new Polyline();
                    for (int i = 0; i < pts.Count; i++)
                    {
                        double wx = (pts[i].X - 50) / scale + ext.MinPoint.X;
                        double wy = ext.MaxPoint.Y - (pts[i].Y - 50) / scale;
                        pl.AddVertexAt(i, new Point2d(wx, wy), 0, 0, 0);
                    }
                    pl.Closed = true; pl.Layer = layerName; pl.ColorIndex = 256;
                    btr.AppendEntity(pl); tr.AddNewlyCreatedDBObject(pl, true);
                }
                tr.Commit();
            }
        }
    }
}