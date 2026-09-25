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

        public void Delete(ProjectState state, string snipId, string actor)
        {
            var snip = state.Snips.FirstOrDefault(item => item.Id == snipId);
            if (snip == null) throw new InvalidOperationException("Snip introuvable.");
            var position = state.Snips.IndexOf(snip);
            var previousLinks = state.CellLinks;
            var entry = new AuditEventRecord { Actor = actor ?? "", Action = "SnipDeleted", EntityType = "Snip", EntityId = snipId,
                Details = snip.WorksheetName + "!" + snip.CellAddress };
            state.Snips.RemoveAt(position);
            state.CellLinks = previousLinks.Select(link => new CellLinkRecord { WorksheetName = link.WorksheetName,
                CellAddress = link.CellAddress, SnipIds = link.SnipIds.Where(id => id != snipId).ToList() }).Where(link => link.SnipIds.Count > 0).ToList();
            state.AuditTrail.Add(entry);
            try { store.Save(state); }
            catch { state.Snips.Insert(position, snip); state.CellLinks = previousLinks; state.AuditTrail.Remove(entry); throw; }
        }

        public void UpdateGeometry(ProjectState state,string id,NormalizedRectangle rectangle,string rawText,string actor,Action<SnipRecord> applyCells=null)
        {
            var snip=state.Snips.FirstOrDefault(s=>s.Id==id) ?? throw new InvalidOperationException("Snip introuvable.");
            var value=parser.Parse(snip.Type,rawText);
            var oldRectangle=new NormalizedRectangle(snip.X,snip.Y,snip.Width,snip.Height);
            var oldText=snip.RawText;var oldValue=snip.ExtractedValue;var oldStatus=snip.Status;
            var oldReviewer=snip.ReviewedBy;var oldReviewDate=snip.ReviewedAtUtc;
            var entry=new AuditEventRecord {Actor=actor??"",Action="SnipResized",EntityType="Snip",EntityId=id};
            snip.X=rectangle.X;snip.Y=rectangle.Y;snip.Width=rectangle.Width;snip.Height=rectangle.Height;
            snip.RawText=rawText;snip.ExtractedValue=value;snip.Status=ReviewStatus.Prepared;snip.ReviewedBy="";snip.ReviewedAtUtc=null;
            state.AuditTrail.Add(entry);
            try{applyCells?.Invoke(snip);store.Save(state);}
            catch
            {
                snip.X=oldRectangle.X;snip.Y=oldRectangle.Y;snip.Width=oldRectangle.Width;snip.Height=oldRectangle.Height;
                snip.RawText=oldText;snip.ExtractedValue=oldValue;snip.Status=oldStatus;snip.ReviewedBy=oldReviewer;snip.ReviewedAtUtc=oldReviewDate;
                state.AuditTrail.Remove(entry);throw;
            }
        }

        public void SetReview(
            ProjectState state,
            string snipId,
            ReviewStatus status,
            string comment,
            string actor,
            Action<SnipRecord> applyMetadata = null)
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
            try { applyMetadata?.Invoke(snip); store.Save(state); }
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
