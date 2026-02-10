using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using SolidWorks.Interop.sldworks;

namespace SolidWorksImageTracerAddin;

internal sealed class ImageTraceService
{
    public IReadOnlyList<List<PointF>> TraceToPolylines(string pngPath)
    {
        using var bitmap = new Bitmap(pngPath);

        bool[,] mask = BuildBinaryMask(bitmap);
        var segments = BuildBoundarySegments(mask);
        var polylines = StitchSegmentsIntoPolylines(segments);

        if (polylines.Count == 0)
        {
            throw new InvalidOperationException("No traceable black pixels were found in the PNG.");
        }

        return polylines;
    }

    public void DrawPolylinesIntoSketch(IModelDoc2 model, IReadOnlyList<List<PointF>> polylines, double targetWidthMeters)
    {
        var sketchManager = model.SketchManager;
        bool wasSketching = sketchManager.ActiveSketch != null;

        if (!wasSketching)
        {
            sketchManager.InsertSketch(true);
        }

        var bounds = GetBounds(polylines);
        float width = bounds.maxX - bounds.minX;
        if (width <= 0)
        {
            throw new InvalidOperationException("Cannot scale image because traced width is zero.");
        }

        double scale = targetWidthMeters / width;

        foreach (var polyline in polylines)
        {
            if (polyline.Count < 2)
            {
                continue;
            }

            for (int i = 1; i < polyline.Count; i++)
            {
                PointF p1 = polyline[i - 1];
                PointF p2 = polyline[i];

                double x1 = (p1.X - bounds.minX) * scale;
                double y1 = -(p1.Y - bounds.minY) * scale;
                double x2 = (p2.X - bounds.minX) * scale;
                double y2 = -(p2.Y - bounds.minY) * scale;

                sketchManager.CreateLine(x1, y1, 0, x2, y2, 0);
            }
        }

        if (!wasSketching)
        {
            sketchManager.InsertSketch(true);
        }

        model.GraphicsRedraw2();
    }

    private static bool[,] BuildBinaryMask(Bitmap bitmap)
    {
        int width = bitmap.Width;
        int height = bitmap.Height;
        var mask = new bool[width, height];

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                Color c = bitmap.GetPixel(x, y);
                int luminance = (c.R + c.G + c.B) / 3;
                mask[x, y] = luminance < 128;
            }
        }

        return mask;
    }

    private static List<(PointF A, PointF B)> BuildBoundarySegments(bool[,] mask)
    {
        int width = mask.GetLength(0);
        int height = mask.GetLength(1);
        var segments = new List<(PointF A, PointF B)>();

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                if (!mask[x, y])
                {
                    continue;
                }

                if (x == 0 || !mask[x - 1, y])
                    segments.Add((new PointF(x, y), new PointF(x, y + 1)));

                if (x == width - 1 || !mask[x + 1, y])
                    segments.Add((new PointF(x + 1, y + 1), new PointF(x + 1, y)));

                if (y == 0 || !mask[x, y - 1])
                    segments.Add((new PointF(x + 1, y), new PointF(x, y)));

                if (y == height - 1 || !mask[x, y + 1])
                    segments.Add((new PointF(x, y + 1), new PointF(x + 1, y + 1)));
            }
        }

        return segments;
    }

    private static List<List<PointF>> StitchSegmentsIntoPolylines(List<(PointF A, PointF B)> segments)
    {
        var adjacency = new Dictionary<PointKey, List<PointF>>();
        foreach (var seg in segments)
        {
            var aKey = new PointKey(seg.A);
            var bKey = new PointKey(seg.B);

            if (!adjacency.TryGetValue(aKey, out var aList))
            {
                aList = new List<PointF>();
                adjacency[aKey] = aList;
            }

            if (!adjacency.TryGetValue(bKey, out var bList))
            {
                bList = new List<PointF>();
                adjacency[bKey] = bList;
            }

            aList.Add(seg.B);
            bList.Add(seg.A);
        }

        var unused = new HashSet<(PointKey, PointKey)>(segments.Select(s => NormalizeEdge(s.A, s.B)));
        var polylines = new List<List<PointF>>();

        while (unused.Count > 0)
        {
            var edge = unused.First();
            var polyline = new List<PointF> { edge.Item1.ToPointF(), edge.Item2.ToPointF() };
            unused.Remove(edge);

            ExtendPolyline(polyline, unused, adjacency, forward: true);
            ExtendPolyline(polyline, unused, adjacency, forward: false);

            if (polyline.Count > 2)
            {
                if (!polyline[0].Equals(polyline[^1]))
                {
                    polyline.Add(polyline[0]);
                }

                polylines.Add(polyline);
            }
        }

        return polylines;
    }

    private static void ExtendPolyline(
        List<PointF> polyline,
        HashSet<(PointKey, PointKey)> unused,
        Dictionary<PointKey, List<PointF>> adjacency,
        bool forward)
    {
        while (true)
        {
            PointF current = forward ? polyline[^1] : polyline[0];
            PointKey currentKey = new(current);

            if (!adjacency.TryGetValue(currentKey, out var neighbors))
            {
                return;
            }

            bool advanced = false;
            foreach (PointF candidate in neighbors)
            {
                var edge = NormalizeEdge(current, candidate);
                if (!unused.Contains(edge))
                {
                    continue;
                }

                unused.Remove(edge);

                if (forward)
                {
                    polyline.Add(candidate);
                }
                else
                {
                    polyline.Insert(0, candidate);
                }

                advanced = true;
                break;
            }

            if (!advanced)
            {
                return;
            }
        }
    }

    private static (PointKey, PointKey) NormalizeEdge(PointF a, PointF b)
    {
        var ka = new PointKey(a);
        var kb = new PointKey(b);
        return ka.CompareTo(kb) <= 0 ? (ka, kb) : (kb, ka);
    }

    private static (float minX, float maxX, float minY, float maxY) GetBounds(IReadOnlyList<List<PointF>> polylines)
    {
        float minX = float.MaxValue;
        float maxX = float.MinValue;
        float minY = float.MaxValue;
        float maxY = float.MinValue;

        foreach (var polyline in polylines)
        {
            foreach (PointF p in polyline)
            {
                minX = Math.Min(minX, p.X);
                maxX = Math.Max(maxX, p.X);
                minY = Math.Min(minY, p.Y);
                maxY = Math.Max(maxY, p.Y);
            }
        }

        return (minX, maxX, minY, maxY);
    }

    private readonly struct PointKey : IEquatable<PointKey>, IComparable<PointKey>
    {
        public PointKey(PointF p)
        {
            X = p.X;
            Y = p.Y;
        }

        public float X { get; }
        public float Y { get; }

        public PointF ToPointF() => new(X, Y);

        public bool Equals(PointKey other) => X.Equals(other.X) && Y.Equals(other.Y);

        public override bool Equals(object? obj) => obj is PointKey other && Equals(other);

        public override int GetHashCode() => HashCode.Combine(X, Y);

        public int CompareTo(PointKey other)
        {
            int xCompare = X.CompareTo(other.X);
            return xCompare != 0 ? xCompare : Y.CompareTo(other.Y);
        }
    }
}
