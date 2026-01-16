using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Plugin_ContourMaster.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
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
                using (Bitmap bmp = CaptureViewport())
                {
                    if (bmp == null) return;
                    bool[,] binaryMap = ApplyThreshold(bmp, (int)_settings.Threshold);
                    List<List<System.Drawing.Point>> contours = FindContours(binaryMap);
                    DrawInCad(doc, contours, bmp.Height);
                }
            }
            catch (Exception ex) { ed.WriteMessage($"\n[算法异常] {ex.Message}"); }
        }

        private Bitmap CaptureViewport()
        {
            Point2d screenSize = (Point2d)Application.GetSystemVariable("SCREENSIZE");
            Bitmap bmp = new Bitmap((int)screenSize.X, (int)screenSize.Y);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(new System.Drawing.Point(0, 0), new System.Drawing.Point(0, 0), bmp.Size);
            }
            return bmp;
        }

        private bool[,] ApplyThreshold(Bitmap bmp, int threshold)
        {
            int w = bmp.Width, h = bmp.Height;
            bool[,] map = new bool[w, h];
            BitmapData data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            byte[] pixels = new byte[Math.Abs(data.Stride) * h];
            Marshal.Copy(data.Scan0, pixels, 0, pixels.Length);
            bmp.UnlockBits(data);

            for (int y = 0; y < h; y++)
                for (int x = 0; x < w; x++)
                {
                    int i = y * data.Stride + x * 4;
                    map[x, y] = (pixels[i] * 0.11 + pixels[i + 1] * 0.59 + pixels[i + 2] * 0.3) > threshold;
                }
            return map;
        }

        private List<List<System.Drawing.Point>> FindContours(bool[,] map)
        {
            int w = map.GetLength(0), h = map.GetLength(1);
            List<List<System.Drawing.Point>> result = new List<List<System.Drawing.Point>>();
            bool[,] visited = new bool[w, h];
            for (int y = 1; y < h - 1; y++)
                for (int x = 1; x < w - 1; x++)
                    if (map[x, y] && !visited[x, y] && (!map[x - 1, y] || !map[x + 1, y]))
                    {
                        var path = Trace(map, x, y, visited);
                        if (path.Count > 5) result.Add(Simplify(path, _settings.SimplifyTolerance));
                    }
            return result;
        }

        private List<System.Drawing.Point> Trace(bool[,] map, int sx, int sy, bool[,] visited)
        {
            List<System.Drawing.Point> path = new List<System.Drawing.Point>();
            int[] dx = { -1, -1, 0, 1, 1, 1, 0, -1 }, dy = { 0, -1, -1, -1, 0, 1, 1, 1 };
            int cx = sx, cy = sy, bx = sx - 1, by = sy;
            do
            {
                path.Add(new System.Drawing.Point(cx, cy)); visited[cx, cy] = true;
                int si = 0;
                for (int i = 0; i < 8; i++) if (cx + dx[i] == bx && cy + dy[i] == by) { si = i; break; }
                for (int i = 1; i <= 8; i++)
                {
                    int idx = (si + i) % 8, nx = cx + dx[idx], ny = cy + dy[idx];
                    if (nx >= 0 && nx < map.GetLength(0) && ny >= 0 && ny < map.GetLength(1) && map[nx, ny])
                    {
                        bx = cx; by = cy; cx = nx; cy = ny; break;
                    }
                }
            } while ((cx != sx || cy != sy) && path.Count < 10000);
            return path;
        }

        private List<System.Drawing.Point> Simplify(List<System.Drawing.Point> pts, double tol)
        {
            if (pts.Count < 3) return pts;
            int idx = -1; double max = 0;
            for (int i = 1; i < pts.Count - 1; i++)
            {
                double d = Math.Abs((pts[i].X - pts[0].X) * (pts[pts.Count - 1].Y - pts[0].Y) - (pts[i].Y - pts[0].Y) * (pts[pts.Count - 1].X - pts[0].X));
                if (d > max) { max = d; idx = i; }
            }
            if (max > tol)
            {
                var l = Simplify(pts.GetRange(0, idx + 1), tol);
                var r = Simplify(pts.GetRange(idx, pts.Count - idx), tol);
                l.RemoveAt(l.Count - 1); l.AddRange(r); return l;
            }
            return new List<System.Drawing.Point> { pts[0], pts[pts.Count - 1] };
        }

        private void DrawInCad(Document doc, List<List<System.Drawing.Point>> contours, int imgH)
        {
            using (var loc = doc.LockDocument())
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                foreach (var pts in contours)
                {
                    Polyline pl = new Polyline();
                    for (int i = 0; i < pts.Count; i++) pl.AddVertexAt(i, new Point2d(pts[i].X, imgH - pts[i].Y), 0, 0, 0);
                    pl.Closed = true; pl.Layer = _settings.LayerName;
                    btr.AppendEntity(pl); tr.AddNewlyCreatedDBObject(pl, true);
                }
                tr.Commit();
            }
        }
    }
}