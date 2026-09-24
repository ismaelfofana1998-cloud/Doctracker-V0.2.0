using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using System.Text.RegularExpressions;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using PdfiumViewer;

namespace Doctracker.AddIn.Infrastructure
{
    internal sealed class DocumentIndexer
    {
        private readonly ProjectStore store;
        private readonly IOcrEngine ocr;
        public DocumentIndexer(ProjectStore store, IOcrEngine ocr) { this.store = store; this.ocr = ocr; }

        public void Index(ProjectState state, DocumentRecord document, Action<int, int> progress, CancellationToken cancellation, bool forceOcr = false)
        {
            var path = store.ResolveDocumentPath(document);
            // Build separately: a failed/cancelled re-index must not erase the previous index.
            var pages = new List<PageTextRecord>();
            if (string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                NativePdfiumLoader.EnsureLoaded();
                using (var pdf = PdfDocument.Load(path))
                {
                    for (var index = 0; index < pdf.PageCount; index++)
                    {
                        cancellation.ThrowIfCancellationRequested();
                        var native = pdf.GetPdfText(index);
                        PageTextRecord page;
                        if (!forceOcr && !string.IsNullOrWhiteSpace(native)) page = ReadNativePage(pdf, index, native);
                        else
                        {
                            var size = RenderSize(pdf.PageSizes[index]);
                            using (var rendered = pdf.Render(index, size.Width, size.Height, 144, 144, PdfRenderFlags.Annotations))
                            using (var bitmap = new Bitmap(rendered)) page = ocr.Recognize(bitmap);
                        }
                        page.PageNumber = index + 1;
                        pages.Add(page);
                        progress?.Invoke(index + 1, pdf.PageCount);
                    }
                }
            }
            else
            {
                using (var original = Image.FromFile(path))
                {
                    var count = ImagePageCount(original);
                    for (var index = 0; index < count; index++)
                    {
                        cancellation.ThrowIfCancellationRequested();
                        if (count > 1) original.SelectActiveFrame(FrameDimension.Page, index);
                        using (var bitmap = new Bitmap(original))
                        {
                            var page = ocr.Recognize(bitmap);
                            page.PageNumber = index + 1;
                            pages.Add(page);
                        }
                        progress?.Invoke(index + 1, count);
                    }
                }
            }
            cancellation.ThrowIfCancellationRequested();
            List<PageTextRecord> oldPages;
            try{oldPages=document.IndexedPages;}
            catch(Exception failure) when(failure is InvalidOperationException || failure is System.Xml.XmlException || failure is IOException || failure is InvalidDataException)
            {oldPages=new List<PageTextRecord>();}
            var oldCount = document.PageCount;
            var oldComplete = document.IndexComplete;
            var oldError = document.IndexError;
            document.IndexedPages = pages;
            document.PageCount = pages.Count;
            document.IndexComplete = true;
            document.IndexError = "";
            var entry = new AuditEventRecord { Actor = Environment.UserName, Action = "DocumentIndexed",
                EntityType = "Document", EntityId = document.Id, Details = pages.Count + " page(s)" };
            state.AuditTrail.Add(entry);
            try { store.Save(state); document.ReleaseIndex(); }
            catch
            {
                document.IndexedPages = oldPages; document.PageCount = oldCount;
                document.IndexComplete = oldComplete; document.IndexError = oldError;
                state.AuditTrail.Remove(entry);
                throw;
            }
        }

        public List<string> IndexMissing(ProjectState state, Action<string, int, int> progress, CancellationToken cancellation, IEnumerable<DocumentRecord> scope = null, bool retryFailed = true, bool forceReindex = false)
        {
            var errors = new List<string>();
            var changed = false;
            foreach (var document in (scope ?? state.Documents).ToList())
            {
                cancellation.ThrowIfCancellationRequested();
                if (!retryFailed && !string.IsNullOrEmpty(document.IndexError))
                { errors.Add(document.OriginalName + " : " + document.IndexError); continue; }
                if(!forceReindex && store.ValidateIndex(document))continue;
                try { Index(state, document, (page, count) => progress?.Invoke(document.OriginalName, page, count), cancellation); }
                catch (OperationCanceledException) { throw; }
                catch (Exception exception)
                {
                    document.IndexError = exception.Message; changed = true;
                    errors.Add(document.OriginalName + " : " + exception.Message);
                }
            }
            if (changed) store.Save(state);
            return errors;
        }

        internal static PageTextRecord ReadNativePage(string path, int pageNumber)
        {
            if (!string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase)) return null;
            NativePdfiumLoader.EnsureLoaded();
            using (var pdf = PdfDocument.Load(path))
            {
                if (pageNumber < 1 || pageNumber > pdf.PageCount) return null;
                var text = pdf.GetPdfText(pageNumber - 1);
                if (string.IsNullOrWhiteSpace(text)) return null;
                var page = ReadNativePage(pdf, pageNumber - 1, text);
                page.PageNumber = pageNumber; return page;
            }
        }

        private static PageTextRecord ReadNativePage(PdfDocument pdf, int index, string text)
        {
            var page = new PageTextRecord { Text = text };
            var size = pdf.PageSizes[index];
            var line = 0;
            double previousY = -1;
            foreach (Match match in Regex.Matches(text, @"\S+"))
            {
                var boxes = pdf.GetTextBounds(new PdfTextSpan(index, match.Index, match.Length));
                if (boxes.Count == 0) continue;
                // PDFium uses a bottom-left origin and negative rectangle heights.
                var x = Math.Max(0, boxes.Min(box => box.Bounds.Left)) / size.Width;
                var y = Math.Max(0, size.Height - boxes.Max(box => box.Bounds.Top)) / size.Height;
                var right = Math.Min(size.Width, boxes.Max(box => box.Bounds.Right)) / size.Width;
                var bottom = Math.Min(size.Height, size.Height - boxes.Min(box => box.Bounds.Bottom)) / size.Height;
                if (right <= x || bottom <= y) continue;
                if (previousY < 0 || Math.Abs(y - previousY) > (bottom - y) * 0.6) line++;
                previousY = y;
                page.Words.Add(new WordRecord { Text = match.Value, Line = line, X = x, Y = y, Width = right - x, Height = bottom - y });
            }
            return page;
        }

        internal static Size RenderSize(SizeF page)
        {
            var scale = Math.Min(2.5, 3000d / Math.Max(page.Width, page.Height));
            return new Size(Math.Max(1, (int)Math.Ceiling(page.Width * scale)), Math.Max(1, (int)Math.Ceiling(page.Height * scale)));
        }
        internal static int ImagePageCount(Image image) => image.FrameDimensionsList.Contains(FrameDimension.Page.Guid)
            ? image.GetFrameCount(FrameDimension.Page) : 1;
    }
}
