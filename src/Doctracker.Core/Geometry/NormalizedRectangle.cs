using System;

namespace Doctracker.Core.Geometry
{
    public sealed class NormalizedRectangle
    {
        public NormalizedRectangle(double x, double y, double width, double height)
        {
            if (x < 0 || y < 0 || width <= 0 || height <= 0 ||
                x + width > 1.000001 || y + height > 1.000001)
            {
                throw new ArgumentOutOfRangeException(nameof(width), "The snip rectangle must be inside the page.");
            }

            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        public double X { get; }
        public double Y { get; }
        public double Width { get; }
        public double Height { get; }
    }
}
