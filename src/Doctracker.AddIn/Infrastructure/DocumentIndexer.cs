using System;
using System.Drawing;
using System.IO;
using System.Linq;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using PdfiumViewer;

namespace Doctracker.AddIn.Infrastructure
{
    internal sealed class DocumentIndexer
    {
        private readonly ProjectStore store;
        private readonly IOcrEngine ocr;

        public DocumentIndexer(ProjectStore store, IOcrEngine ocr)
        {
            this.store = store ?? throw new ArgumentNullException(nameof(store));
            this.ocr = ocr ?? throw new ArgumentNullException(nameof(ocr));
        }

        public void Index(ProjectState state, DocumentRecord document, Action<int, int> progress)
        {
            var path = store.ResolveDocumentPath(document);
            document.IndexedPages.Clear();

            if (string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                NativePdfiumLoader.EnsureLoaded();
                using (var pdf = PdfDocument.Load(path))
                {
                    document.PageCount = pdf.PageCount;
                    for (var index = 0; index < pdf.PageCount; index++)
                    {
                        using (var rendered = pdf.Render(
                            index, 1800, 2400, 144, 144,
                            PdfRenderFlags.Annotations | PdfRenderFlags.LcdText))
                        using (var bitmap = new Bitmap(rendered))
                        {
                            document.IndexedPages.Add(new PageTextRecord
                            {
                                PageNumber = index + 1,
                                Text = ocr.Recognize(bitmap)
                            });
                        }
                        progress?.Invoke(index + 1, pdf.PageCount);
                    }
                }
            }
            else
            {
                using (var original = Image.FromFile(path))
                using (var bitmap = new Bitmap(original))
                {
                    document.PageCount = 1;
                    document.IndexedPages.Add(new PageTextRecord
                    {
                        PageNumber = 1,
                        Text = ocr.Recognize(bitmap)
                    });
                    progress?.Invoke(1, 1);
                }
            }

            state.AuditTrail.Add(new AuditEventRecord
            {
                Actor = Environment.UserName,
                Action = "DocumentIndexed",
                EntityType = "Document",
                EntityId = document.Id,
                Details = document.IndexedPages.Count + " page(s)"
            });
            store.Save(state);
        }

        public void IndexMissing(ProjectState state, Action<string, int, int> progress)
        {
            foreach (var document in state.Documents.Where(item => item.IndexedPages.Count == 0))
            {
                Index(state, document,
                    (page, count) => progress?.Invoke(document.OriginalName, page, count));
            }
        }
    }
}
