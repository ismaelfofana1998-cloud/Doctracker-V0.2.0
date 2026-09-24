using System;
using System.Drawing;
using System.Linq;
using Doctracker.Core.Models;

namespace Doctracker.AddIn.UI
{
    internal static class DocumentOverlay
    {
        public static RectangleF Bounds(DocumentComment comment, Size size) => new RectangleF((float)comment.X * size.Width,
            (float)comment.Y * size.Height, (float)comment.Width * size.Width, (float)comment.Height * size.Height);
        public static float FontPixels(int width, double fontSize = 16) => (float)Math.Max(1, width * fontSize / 595d);
        public static void Draw(Graphics graphics, Size size, DocumentRecord document, int page)
        {
            if (document == null || size.Width < 1 || size.Height < 1) return;
            using (var red = new SolidBrush(Color.Red))
            using (var pen = new Pen(Color.Red, Math.Max(1f, size.Width * .0015f)))
            {
                foreach (var comment in document.Comments.Where(c => c.PageNumber == page))
                {
                    var box = Bounds(comment, size); if (box.Width < 2 || box.Height < 2) continue;
                    graphics.FillRectangle(Brushes.White, box); graphics.DrawRectangle(pen, box.X, box.Y, box.Width, box.Height);
                    var padding = Math.Max(1, size.Width * .006f); box.Inflate(-padding, -padding);
                    if (box.Width <= 0 || box.Height <= 0) continue;
                    using (var font = new Font("Segoe UI", FontPixels(size.Width, comment.FontSize), FontStyle.Regular, GraphicsUnit.Pixel))
                    using (var format = new StringFormat { Trimming = StringTrimming.EllipsisWord })
                        graphics.DrawString(comment.Text, font, red, box, format);
                }
                if (page == 1 && !string.IsNullOrWhiteSpace(document.TestReference))
                {
                    var label = document.TestReference + " / " + document.ReferenceNumber.ToString("D2");
                    var box = new RectangleF(size.Width * .03f, size.Height * .012f, size.Width * .94f, Math.Max(size.Height * .08f, size.Width * .08f));
                    using (var font = new Font("Segoe UI", size.Width * .03f, FontStyle.Bold, GraphicsUnit.Pixel))
                    using (var format = new StringFormat { Alignment = StringAlignment.Far })
                        graphics.DrawString(label, font, red, box, format);
                }
            }
        }
    }
}
