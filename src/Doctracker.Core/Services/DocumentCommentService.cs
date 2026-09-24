using System;
using System.Linq;
using Doctracker.Core.Geometry;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    public sealed class DocumentCommentService
    {
        private readonly ProjectStore store;
        public DocumentCommentService(ProjectStore store) { this.store = store; }
        public DocumentComment Save(ProjectState state, string documentId, int page, NormalizedRectangle zone, string text, string id = null)
        {
            var document = state.Documents.First(d => d.Id == documentId);
            if (page < 1 || (document.PageCount > 0 && page > document.PageCount)) throw new ArgumentOutOfRangeException(nameof(page));
            if (zone == null || string.IsNullOrWhiteSpace(text) || text.Length > 2000) throw new ArgumentException("Saisissez un commentaire de 1 à 2 000 caractères.");
            var previous = document.Comments.FirstOrDefault(c => c.Id == id);
            if (id != null && previous == null) throw new InvalidOperationException("Commentaire introuvable.");
            var comment = new DocumentComment { Id = id ?? Guid.NewGuid().ToString("N"), PageNumber = page,
                X = zone.X, Y = zone.Y, Width = zone.Width, Height = zone.Height, Text = text.Trim() };
            var index = previous == null ? document.Comments.Count : document.Comments.IndexOf(previous);
            if (previous != null) document.Comments.RemoveAt(index);
            document.Comments.Insert(index, comment);
            try { store.Save(state); }
            catch { document.Comments.RemoveAt(index); if (previous != null) document.Comments.Insert(index, previous); throw; }
            return comment;
        }
        public void Delete(ProjectState state, string documentId, string id)
        {
            var document = state.Documents.First(d => d.Id == documentId);
            var comment = document.Comments.First(c => c.Id == id);
            var index = document.Comments.IndexOf(comment); document.Comments.RemoveAt(index);
            try { store.Save(state); } catch { document.Comments.Insert(index, comment); throw; }
        }
    }
}
