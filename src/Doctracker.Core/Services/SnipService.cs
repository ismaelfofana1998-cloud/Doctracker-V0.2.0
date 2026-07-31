using System;
using System.Linq;
using Doctracker.Core.Geometry;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    public sealed class SnipService
    {
        private readonly ProjectStore store;
        private readonly TextValueParser parser;

        public SnipService(ProjectStore store, TextValueParser parser)
        {
            this.store = store ?? throw new ArgumentNullException(nameof(store));
            this.parser = parser ?? throw new ArgumentNullException(nameof(parser));
        }

        public SnipRecord Create(
            ProjectState state,
            string documentId,
            int pageNumber,
            NormalizedRectangle rectangle,
            SnipType type,
            string rawText,
            string worksheetName,
            string cellAddress,
            string actor)
        {
            if (state == null) throw new ArgumentNullException(nameof(state));
            if (rectangle == null) throw new ArgumentNullException(nameof(rectangle));
            var sourceDocument = state.Documents.FirstOrDefault(document => document.Id == documentId);
            if (sourceDocument == null)
                throw new InvalidOperationException("The source document is not part of this project.");
            if (pageNumber < 1) throw new ArgumentOutOfRangeException(nameof(pageNumber));
            if (sourceDocument.PageCount > 0 && pageNumber > sourceDocument.PageCount)
                throw new ArgumentOutOfRangeException(nameof(pageNumber), "The requested page does not exist.");

            var snip = new SnipRecord
            {
                DocumentId = documentId,
                PageNumber = pageNumber,
                X = rectangle.X,
                Y = rectangle.Y,
                Width = rectangle.Width,
                Height = rectangle.Height,
                Type = type,
                RawText = rawText ?? string.Empty,
                ExtractedValue = parser.Parse(type, rawText),
                WorksheetName = worksheetName ?? string.Empty,
                CellAddress = cellAddress ?? string.Empty,
                PreparedBy = actor ?? string.Empty,
                PreparedAtUtc = DateTime.UtcNow,
                Status = ReviewStatus.Prepared
            };

            state.Snips.Add(snip);
            state.AuditTrail.Add(new AuditEventRecord
            {
                Actor = actor ?? string.Empty,
                Action = "SnipCreated",
                EntityType = "Snip",
                EntityId = snip.Id,
                Details = worksheetName + "!" + cellAddress
            });
            store.Save(state);
            return snip;
        }

        public void SetReview(
            ProjectState state,
            string snipId,
            ReviewStatus status,
            string comment,
            string actor)
        {
            if (state == null) throw new ArgumentNullException(nameof(state));
            if (string.IsNullOrWhiteSpace(snipId))
                throw new ArgumentException("A snip identifier is required.", nameof(snipId));

            var snip = state.Snips.FirstOrDefault(item => item.Id == snipId);
            if (snip == null) throw new InvalidOperationException("Snip not found.");

            snip.Status = status;
            snip.Comment = comment ?? string.Empty;
            snip.ReviewedBy = actor ?? string.Empty;
            snip.ReviewedAtUtc = DateTime.UtcNow;
            state.AuditTrail.Add(new AuditEventRecord
            {
                Actor = actor ?? string.Empty,
                Action = "SnipReviewed",
                EntityType = "Snip",
                EntityId = snip.Id,
                Details = status + ": " + snip.Comment
            });
            store.Save(state);
        }
    }
}
