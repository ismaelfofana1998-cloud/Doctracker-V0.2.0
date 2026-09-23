using System;
using System.Linq;
using System.Collections.Generic;
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

        public SnipRecord Prepare(
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

            return snip;
        }

        public void Commit(ProjectState state, IReadOnlyList<SnipRecord> snips)
        {
            var events = snips.Select(snip => new AuditEventRecord
            {
                Actor = snip.PreparedBy, Action = "SnipCreated", EntityType = "Snip", EntityId = snip.Id,
                Details = snip.WorksheetName + "!" + snip.CellAddress
            }).ToList();
            state.Snips.AddRange(snips);
            state.AuditTrail.AddRange(events);
            try { store.Save(state); }
            catch
            {
                foreach (var snip in snips) state.Snips.Remove(snip);
                foreach (var entry in events) state.AuditTrail.Remove(entry);
                throw;
            }
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

            var oldStatus = snip.Status;
            var oldComment = snip.Comment;
            var oldReviewer = snip.ReviewedBy;
            var oldDate = snip.ReviewedAtUtc;
            snip.Status = status;
            snip.Comment = comment ?? string.Empty;
            snip.ReviewedBy = status == ReviewStatus.Prepared ? string.Empty : actor ?? string.Empty;
            snip.ReviewedAtUtc = status == ReviewStatus.Prepared ? (DateTime?)null : DateTime.UtcNow;
            state.AuditTrail.Add(new AuditEventRecord
            {
                Actor = actor ?? string.Empty,
                Action = "SnipReviewed",
                EntityType = "Snip",
                EntityId = snip.Id,
                Details = status + ": " + snip.Comment
            });
            try { store.Save(state); }
            catch
            {
                state.AuditTrail.RemoveAt(state.AuditTrail.Count - 1);
                snip.Status = oldStatus; snip.Comment = oldComment;
                snip.ReviewedBy = oldReviewer; snip.ReviewedAtUtc = oldDate;
                throw;
            }
        }
    }
}
