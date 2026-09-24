using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using Doctracker.AddIn.Infrastructure;
using Doctracker.Core.Models;
using PdfiumViewer;

namespace Doctracker.AddIn.UI
{
    /// <summary>
    /// PDF/image surface with a real scrollable zoom surface.
    /// Left drag selects a snip; middle drag pans the zoomed document.
    /// </summary>
    internal sealed class DocumentCanvas : UserControl
    {
        private readonly Panel viewport;
        private readonly PictureBox picture;
        private readonly Label pageLabel;
        private readonly Label zoomLabel;
        private readonly Button previousButton;
        private readonly Button nextButton;
        private readonly Button zoomOutButton;
        private readonly Button fitButton;
        private readonly Button zoomInButton;

        private PdfDocument pdf;
        private Image currentImage;
        private string currentPath;
        private int pageIndex;
        private int imagePageCount = 1;
        private double zoom = 1d;
        private bool fitToViewport = true;
        private Point dragStart;
        private Point dragEnd;
        private bool dragging;
        private bool panning;
        private Point panStart;
        private Point panOrigin;
        private RectangleF? normalizedSelection;
        private readonly List<SnipRecord> proofs = new List<SnipRecord>();
        private SnipType selectionType = SnipType.Text;
        private string selectedProofId;
        public SnipType? ActiveType { get; set; }
        public void SetProofs(IEnumerable<SnipRecord> records)
        { proofs.Clear(); proofs.AddRange(records); picture.Invalidate(); }


        public DocumentCanvas()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer, true);

            Dock = DockStyle.Fill;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96,96);
            BackColor = SnipTheme.Surface;

            var toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(6, 5, 6, 4),
                BackColor = Color.White
            };

            previousButton = CreateToolbarButton("", "Page précédente", "Previous");
            nextButton = CreateToolbarButton("", "Page suivante", "Next");
            pageLabel = new Label
            {
                Text = "Aucun document",
                AutoSize = true,
                Padding = new Padding(8, 5, 8, 0),
                ForeColor = Color.FromArgb(45, 53, 64)
            };
            zoomOutButton = CreateToolbarButton("", "Réduire le zoom", "Minus");
            fitButton = CreateToolbarButton("Adapter", "Adapter la page à la fenêtre", "Fit");
            zoomInButton = CreateToolbarButton("", "Agrandir le zoom", "Plus");
            zoomLabel = new Label
            {
                Text = "100 %",
                AutoSize = true,
                Padding = new Padding(4, 5, 4, 0),
                ForeColor = Color.FromArgb(80, 88, 98)
            };

            previousButton.Click += (sender, args) => ShowPage(pageIndex - 1);
            nextButton.Click += (sender, args) => ShowPage(pageIndex + 1);
            zoomOutButton.Click += (sender, args) => SetZoom(zoom - 0.15d, false);
            fitButton.Click += (sender, args) => FitPage();
            zoomInButton.Click += (sender, args) => SetZoom(zoom + 0.15d, false);

            toolbar.Controls.Add(previousButton);
            toolbar.Controls.Add(nextButton);
            toolbar.Controls.Add(pageLabel);
            toolbar.Controls.Add(zoomOutButton);
            toolbar.Controls.Add(fitButton);
            toolbar.Controls.Add(zoomInButton);
            toolbar.Controls.Add(zoomLabel);

            viewport = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.FromArgb(235, 239, 245),
                Padding = new Padding(12)
            };
            viewport.Resize += (sender, args) =>
            {
                if (currentImage != null) UpdatePictureLayout();
            };

            picture = new PictureBox
            {
                SizeMode = PictureBoxSizeMode.StretchImage,
                BackColor = Color.White,
                Cursor = Cursors.Cross,
                TabStop = false
            };
            picture.MouseDown += Picture_MouseDown;
            picture.MouseMove += Picture_MouseMove;
            picture.MouseUp += Picture_MouseUp;
            picture.Paint += Picture_Paint;

            viewport.Controls.Add(picture);
            Controls.Add(viewport);
            Controls.Add(toolbar);

            UpdateNavigationState();
        }

        public event EventHandler SelectionCompleted;

        public string CurrentPath => currentPath;
        public void ClearDocument() => DisposeDocument();
        public int CurrentPageNumber => pageIndex + 1;
        public bool HasDocument => currentImage != null;
        public bool HasSelection => normalizedSelection.HasValue;

        public void LoadDocument(string path)
        {
            DisposeDocument();
            proofs.Clear(); selectedProofId = null;
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                throw new FileNotFoundException("The document cannot be found.", path);
            }

            try
            {
                currentPath = path;
                pageIndex = 0;
                fitToViewport = true;
                normalizedSelection = null;

                if (string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    NativePdfiumLoader.EnsureLoaded();
                    pdf = PdfDocument.Load(path);
                }
                if (pdf == null)
                    using (var image = Image.FromFile(path)) imagePageCount = DocumentIndexer.ImagePageCount(image);
                RenderCurrentPage();
            }
            catch
            {
                DisposeDocument();
                throw;
            }
        }

        public void NavigateTo(string path, SnipRecord snip)
        {
            if (!string.Equals(currentPath, path, StringComparison.OrdinalIgnoreCase))
            {
                LoadDocument(path);
            }

            ShowPage(Math.Max(0, snip.PageNumber - 1));
            selectionType = snip.SourceType ?? snip.Type;
            selectedProofId = snip.Id;
            normalizedSelection = new RectangleF(
                (float)snip.X, (float)snip.Y, (float)snip.Width, (float)snip.Height);
            RevealSelection();
            picture.Invalidate();
        }

        public void NavigateTo(string path, int targetPageNumber)
        {
            if (!string.Equals(currentPath, path, StringComparison.OrdinalIgnoreCase))
            {
                LoadDocument(path);
            }

            ShowPage(Math.Max(0, targetPageNumber - 1));
            normalizedSelection = null;
            FitPage();
            picture.Invalidate();
        }

        public RectangleF GetNormalizedSelection()
        {
            if (!normalizedSelection.HasValue)
                throw new InvalidOperationException("Draw a zone on the document first.");
            return normalizedSelection.Value;
        }

        public Bitmap CropSelection()
        {
            if (currentImage == null) throw new InvalidOperationException("No document is open.");
            var selection = GetNormalizedSelection();
            var crop = Rectangle.FromLTRB(
                Math.Max(0, (int)Math.Floor(selection.Left * currentImage.Width)),
                Math.Max(0, (int)Math.Floor(selection.Top * currentImage.Height)),
                Math.Min(currentImage.Width, (int)Math.Ceiling(selection.Right * currentImage.Width)),
                Math.Min(currentImage.Height, (int)Math.Ceiling(selection.Bottom * currentImage.Height)));
            if (crop.Width < 2 || crop.Height < 2)
                throw new InvalidOperationException("The selected zone is too small.");

            var output = new Bitmap(crop.Width, crop.Height);
            var dpiX = currentImage.HorizontalResolution > 0 ? currentImage.HorizontalResolution : 96f;
            var dpiY = currentImage.VerticalResolution > 0 ? currentImage.VerticalResolution : 96f;
            output.SetResolution(dpiX, dpiY);
            using (var graphics = Graphics.FromImage(output))
            {
                graphics.DrawImage(currentImage,
                    new Rectangle(0, 0, crop.Width, crop.Height),
                    crop,
                    GraphicsUnit.Pixel);
            }
            return output;
        }

        public void ClearSelection()
        {
            normalizedSelection = null;
            selectedProofId = null;
            picture.Invalidate();
        }

        private void ShowPage(int requestedIndex)
        {
            var pageCount = pdf == null ? (currentPath == null ? 0 : imagePageCount) : pdf.PageCount;
            if (requestedIndex < 0 || requestedIndex >= pageCount || (requestedIndex == pageIndex && currentImage != null)) return;
            pageIndex = requestedIndex;
            selectedProofId = null;
            normalizedSelection = null;
            RenderCurrentPage();
        }

        private void RenderCurrentPage()
        {
            if (currentImage != null)
            {
                picture.Image = null;
                currentImage.Dispose();
                currentImage = null;
            }

            if (pdf != null)
            {
                var size = DocumentIndexer.RenderSize(pdf.PageSizes[pageIndex]);
                currentImage = pdf.Render(
                    pageIndex, size.Width, size.Height, 144, 144,
                    PdfRenderFlags.Annotations | PdfRenderFlags.LcdText);
                pageLabel.Text = "Page " + (pageIndex + 1) + " / " + pdf.PageCount;
            }
            else if (!string.IsNullOrWhiteSpace(currentPath))
            {
                using (var source = Image.FromFile(currentPath))
                {
                    if (imagePageCount > 1) source.SelectActiveFrame(FrameDimension.Page, pageIndex);
                    currentImage = new Bitmap(source);
                }
                pageLabel.Text = "Page " + (pageIndex + 1) + " / " + imagePageCount;
            }
            else
            {
                pageLabel.Text = "Aucun document";
            }

            picture.Image = currentImage;
            UpdatePictureLayout();
            UpdateNavigationState();
            picture.Invalidate();
        }

        private void UpdatePictureLayout()
        {
            if (currentImage == null) return;

            if (fitToViewport)
            {
                var availableWidth = Math.Max(120, viewport.ClientSize.Width - viewport.Padding.Horizontal - 4);
                var availableHeight = Math.Max(120, viewport.ClientSize.Height - viewport.Padding.Vertical - 4);
                var widthRatio = availableWidth / (double)currentImage.Width;
                var heightRatio = availableHeight / (double)currentImage.Height;
                zoom = Math.Min(widthRatio, heightRatio);
                zoom = Math.Max(0.02d, Math.Min(1.0d, zoom));
            }

            var width = Math.Max(2, (int)Math.Round(currentImage.Width * zoom));
            var height = Math.Max(2, (int)Math.Round(currentImage.Height * zoom));
            picture.Size = new Size(width, height);
            // Centre fitted pages; preserve the real scroll origin when zoomed in.
            picture.Location = new Point(Math.Max(viewport.Padding.Left, (viewport.ClientSize.Width - width) / 2) + viewport.AutoScrollPosition.X,
                viewport.Padding.Top + viewport.AutoScrollPosition.Y);
            picture.Image = currentImage;
            picture.Invalidate();
            zoomLabel.Text = Math.Round(zoom * 100d) + " %";
        }

        private void SetZoom(double requestedZoom, bool fit)
        {
            if (currentImage == null) return;
            fitToViewport = fit;
            if (!fitToViewport)
            {
                zoom = Math.Max(0.02d, Math.Min(3.0d, requestedZoom));
            }
            UpdatePictureLayout();
        }

        private void FitPage()
        {
            SetZoom(1d, true);
        }

        private void Picture_MouseDown(object sender, MouseEventArgs e)
        {
            if (currentImage == null) return;

            if (e.Button == MouseButtons.Middle)
            {
                panning = true;
                picture.Capture = true;
                panStart = picture.PointToScreen(e.Location);
                panOrigin = viewport.AutoScrollPosition;
                picture.Cursor = Cursors.SizeAll;
                return;
            }

            if (e.Button != MouseButtons.Left) return;
            picture.Capture = true;
            dragging = true;
            selectionType = ActiveType ?? SnipType.Text;
            selectedProofId = null;
            dragStart = Clamp(e.Location, picture.ClientRectangle);
            dragEnd = dragStart;
            normalizedSelection = null;
        }

        private void Picture_MouseMove(object sender, MouseEventArgs e)
        {
            if (panning)
            {
                var screen = picture.PointToScreen(e.Location);
                var x = -panOrigin.X - (screen.X - panStart.X);
                var y = -panOrigin.Y - (screen.Y - panStart.Y);
                viewport.AutoScrollPosition = new Point(Math.Max(0, x), Math.Max(0, y));
                return;
            }

            if (!dragging) return;
            dragEnd = Clamp(e.Location, picture.ClientRectangle);
            picture.Invalidate();
        }

        private void Picture_MouseUp(object sender, MouseEventArgs e)
        {
            if (panning && e.Button == MouseButtons.Middle)
            {
                panning = false;
                picture.Capture = false;
                picture.Cursor = Cursors.Cross;
                return;
            }

            if (!dragging || e.Button != MouseButtons.Left) return;
            dragging = false;
            picture.Capture = false;
            dragEnd = Clamp(e.Location, picture.ClientRectangle);
            var rectangle = NormalizeScreenRectangle(dragStart, dragEnd);
            if (rectangle.Width >= 4 && rectangle.Height >= 4)
            {
                normalizedSelection = new RectangleF(
                    rectangle.Left / (float)picture.ClientSize.Width,
                    rectangle.Top / (float)picture.ClientSize.Height,
                    rectangle.Width / (float)picture.ClientSize.Width,
                    rectangle.Height / (float)picture.ClientSize.Height);
                SelectionCompleted?.Invoke(this, EventArgs.Empty);
            }
            picture.Invalidate();
        }

        private void Picture_Paint(object sender, PaintEventArgs e)
        {
            foreach (var proof in proofs.Where(item => item.PageNumber == CurrentPageNumber && item.Id != selectedProofId))
            {
                var bounds = new Rectangle((int)(proof.X * picture.Width), (int)(proof.Y * picture.Height),
                    (int)(proof.Width * picture.Width), (int)(proof.Height * picture.Height));
                DrawHighlight(e.Graphics, bounds, SnipTheme.ColorFor(proof.SourceType ?? proof.Type), false);
            }
            Rectangle rectangle;
            if (dragging)
            {
                rectangle = NormalizeScreenRectangle(dragStart, dragEnd);
            }
            else if (normalizedSelection.HasValue)
            {
                var selection = normalizedSelection.Value;
                rectangle = new Rectangle(
                    (int)(selection.X * picture.ClientSize.Width),
                    (int)(selection.Y * picture.ClientSize.Height),
                    (int)(selection.Width * picture.ClientSize.Width),
                    (int)(selection.Height * picture.ClientSize.Height));
            }
            else
            {
                return;
            }

            DrawHighlight(e.Graphics, rectangle, SnipTheme.ColorFor(selectionType), true);
        }

        private static void DrawHighlight(Graphics graphics, Rectangle bounds, Color color, bool selected)
        {
            if (bounds.Width <= 0 || bounds.Height <= 0) return;
            using (var fill = new SolidBrush(Color.FromArgb(selected ? 52 : 22, color)))
            using (var pen = new Pen(color, selected ? 2.5f : 1f))
            {
                graphics.FillRectangle(fill, bounds);
                graphics.DrawRectangle(pen, bounds);
                if (selected)
                    using (var handle = new SolidBrush(color))
                    {
                        graphics.FillRectangle(handle, bounds.Left - 3, bounds.Top - 3, 6, 6);
                        graphics.FillRectangle(handle, bounds.Right - 3, bounds.Bottom - 3, 6, 6);
                    }
            }
        }

        private void RevealSelection()
        {
            if (!normalizedSelection.HasValue || fitToViewport) return;
            var zone = normalizedSelection.Value;
            var x = (int)((zone.X + zone.Width / 2) * picture.Width) + viewport.Padding.Left;
            var y = (int)((zone.Y + zone.Height / 2) * picture.Height) + viewport.Padding.Top;
            viewport.AutoScrollPosition = new Point(Math.Max(0, x - viewport.ClientSize.Width / 2), Math.Max(0, y - viewport.ClientSize.Height / 2));
        }

        private void UpdateNavigationState()
        {
            var pageCount = pdf == null ? (currentPath == null ? 0 : imagePageCount) : pdf.PageCount;
            previousButton.Enabled = pageIndex > 0;
            nextButton.Enabled = pageIndex >= 0 && pageIndex < pageCount - 1;
            zoomOutButton.Enabled = currentImage != null;
            fitButton.Enabled = currentImage != null;
            zoomInButton.Enabled = currentImage != null;
        }

        private static Button CreateToolbarButton(string text, string tooltip, string icon)
        {
            var button = SnipTheme.Button(text, icon);
            button.AccessibleName = tooltip;
            var tip = new ToolTip(); tip.SetToolTip(button, tooltip);
            button.Disposed += (sender, args) => tip.Dispose();
            return button;
        }

        private static Point Clamp(Point point, Rectangle bounds)
        {
            return new Point(
                Math.Max(bounds.Left, Math.Min(bounds.Right, point.X)),
                Math.Max(bounds.Top, Math.Min(bounds.Bottom, point.Y)));
        }

        private static Rectangle NormalizeScreenRectangle(Point first, Point second)
        {
            return Rectangle.FromLTRB(
                Math.Min(first.X, second.X),
                Math.Min(first.Y, second.Y),
                Math.Max(first.X, second.X),
                Math.Max(first.Y, second.Y));
        }

        private void DisposeDocument()
        {
            dragging = false; panning = false; picture.Capture = false;
            picture.Image = null;
            if (currentImage != null) currentImage.Dispose();
            if (pdf != null) pdf.Dispose();
            currentImage = null;
            pdf = null;
            currentPath = null;
            normalizedSelection = null;
            selectedProofId = null; proofs.Clear();
            pageIndex = 0;
            zoomLabel.Text = "100 %";
            pageLabel.Text = "Aucun document";
            UpdateNavigationState();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) DisposeDocument();
            base.Dispose(disposing);
        }
    }
}
