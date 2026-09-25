using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    public sealed class ImportBatchResult
    {
        public int ImportedCount { get; internal set; }
        public string LastDocumentId { get; internal set; }
        public bool Cancelled { get; internal set; }
        public List<string> Errors { get; } = new List<string>();
    }

    // Sequential import only. Text extraction and OCR belong to an explicit search/index operation.
    public static class DocumentImportBatch
    {
        public static ImportBatchResult Run(ProjectStore store, ProjectState state, IEnumerable<string> paths,
            Func<string, DocumentRecord> importWithoutSaving, Action<int, string> progress, CancellationToken cancellation)
        {
            var result = new ImportBatchResult();
            var checkpoint = new Checkpoint(state);
            var pending = 0;
            Action flush = () => {
                if (pending == 0) return;
                try { store.Save(state); }
                catch { pending = 0; checkpoint.Restore(state); throw; }
                pending = 0; checkpoint = new Checkpoint(state);
            };
            try
            {
                foreach (var path in paths)
                {
                    cancellation.ThrowIfCancellationRequested();
                    progress?.Invoke(result.ImportedCount, Path.GetFileName(path));
                    try
                    {
                        var doc = importWithoutSaving(path);
                        result.LastDocumentId = doc.Id; result.ImportedCount++; pending++;
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception failure) { result.Errors.Add(Path.GetFileName(path) + " : " + failure.Message); }
                    // A failed file must never trigger an extra checkpoint.
                    if (pending >= 25 && !store.DeferMetadataWrites) flush();
                }
            }
            catch (OperationCanceledException) { result.Cancelled = true; }
            finally { flush(); }
            return result;
        }

        private sealed class Checkpoint
        {
            private readonly List<DocumentRecord> documents;
            private readonly List<AuditEventRecord> audit;
            private readonly List<Tuple<DocumentRecord, DateTime, List<string>>> metadata;
            public Checkpoint(ProjectState state)
            {
                documents = state.Documents.ToList(); audit = state.AuditTrail.ToList();
                metadata = documents.Select(d => Tuple.Create(d, d.LastImportedAtUtc, d.Categories.ToList())).ToList();
            }
            public void Restore(ProjectState state)
            {
                state.Documents = documents; state.AuditTrail = audit;
                foreach (var item in metadata) { item.Item1.LastImportedAtUtc = item.Item2; item.Item1.Categories = item.Item3; }
            }
        }
    }
}
