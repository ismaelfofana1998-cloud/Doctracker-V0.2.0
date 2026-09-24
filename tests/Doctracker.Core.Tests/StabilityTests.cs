using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class StabilityTests : IDisposable
    {
        private readonly string root = Path.Combine(Path.GetTempPath(), "doctracker-stability-" + Guid.NewGuid().ToString("N"));
        private ProjectStore Store(string name) => new ProjectStore(Path.Combine(root, name));
        private static ProjectState Fixture(ProjectStore store)
        {
            var state = new ProjectState();
            state.Documents.Add(new DocumentRecord { PageCount = 1, IndexComplete = true,
                IndexedPages = new List<PageTextRecord> { new PageTextRecord { PageNumber = 1, Text = "Facture X300" } } });
            // Index-only fixture keeps source embedding out of these failure-injection tests.
            state.SharedVaultPath = "shared";
            store.Save(state); return state;
        }
        [Fact]
        public void Failed_metadata_commit_keeps_old_index_and_unsaved_pages_for_retry()
        {
            var store = Store("rollback"); var state = Fixture(store); var doc = state.Documents[0];
            var key = doc.IndexKey; var revision = state.Revision; var updated = state.UpdatedAtUtc;
            doc.IndexedPages = new List<PageTextRecord> { new PageTextRecord { PageNumber = 1, Text = "Rebuilt X500" } };
            using (var locked = new FileStream(store.MetadataPath, FileMode.Open, FileAccess.Read, FileShare.None))
                Assert.ThrowsAny<IOException>(() => store.Save(state));
            Assert.Equal(key, doc.IndexKey); Assert.True(doc.IndexDirty);
            Assert.Equal(revision, state.Revision); Assert.Equal(updated, state.UpdatedAtUtc);
            Assert.Single(Directory.GetFiles(Path.Combine(store.ProjectDirectory, "indexes")));
            Assert.Equal("Facture X300", store.LoadOrCreate("").Documents[0].IndexedPages[0].Text);
            store.Save(state);
            Assert.False(doc.IndexDirty); Assert.NotEqual(key, doc.IndexKey);
            Assert.Equal("Rebuilt X500", store.LoadOrCreate("").Documents[0].IndexedPages[0].Text);
        }
        [Fact]
        public void Failing_save_observer_does_not_report_a_committed_write_as_failed()
        {
            var store = Store("observer"); var state = Fixture(store); var notified = false;
            store.Saved += () => throw new InvalidOperationException("Observer failed");
            store.Saved += () => notified = true;
            state.TestReference = "DAC"; store.Save(state);
            Assert.True(notified); Assert.Equal("DAC", store.LoadOrCreate("").TestReference);
        }
        [Fact]
        public void Opening_and_resaving_a_portable_project_does_not_read_any_index_blobs()
        {
            var source = Store("source"); var state = Fixture(source); var parts = new Parts();
            new PortableProject(parts).Save(source, state); parts.BlobReads = 0;
            var package = new PortableProject(parts); var target = Store("target"); var restored = package.Restore(target);
            Assert.Equal(0, parts.BlobReads);
            Assert.False(File.Exists(target.IndexPath(restored.Documents[0].IndexKey)));
            target.Save(restored); package.Save(target, restored);
            Assert.Equal(0, parts.BlobReads);
            Assert.Single(OccurrenceSearch.Find(restored, "X300"));
            Assert.Equal(1, parts.BlobReads);
            Assert.Single(OccurrenceSearch.Find(restored, "X300"));
            Assert.Equal(1, parts.BlobReads);
        }
        [Fact]
        public void Backup_materializes_unopened_indexes_and_preserves_search_after_restore()
        {
            var source = Store("archive-source"); var state = Fixture(source); var doc = state.Documents[0];
            Directory.CreateDirectory(source.DocumentsDirectory);
            doc.RelativePath = "documents/source.pdf"; doc.OriginalName = "source.pdf";
            File.WriteAllText(source.LocalDocumentPath(doc), "synthetic source");
            doc.Sha256 = DocumentImporter.ComputeSha256(source.LocalDocumentPath(doc));
            state.SharedVaultPath = ""; source.Save(state);
            var parts = new Parts(); new PortableProject(parts).Save(source, state);
            var target = Store("archive-target"); var restored = new PortableProject(parts).Restore(target);
            Assert.False(File.Exists(target.IndexPath(doc.IndexKey)));
            var archive = Path.Combine(root, "backup.dtpack");
            RecoveryArchive.Export(target, restored, archive, default);
            var recovered = RecoveryArchive.Restore(archive, Store("archive-recovery"), default);
            Assert.Single(OccurrenceSearch.Find(recovered, "X300"));
        }
        [Fact]
        public void Failed_attachment_cleanup_keeps_new_commit_and_retries_on_next_save()
        {
            var store = Store("cleanup"); var state = Fixture(store); var parts = new Parts();
            var package = new PortableProject(parts); package.Save(store, state);
            parts.FailDelete = true; state.TestReference = "Updated"; store.Save(state); package.Save(store, state);
            Assert.Equal(2, parts.Ids(PortableProject.ManifestNamespace).Count());
            Assert.Equal("Updated", new PortableProject(parts).Restore(Store("cleanup-read")).TestReference);
            parts.FailDelete = false; package.Save(store, state);
            Assert.Single(parts.Ids(PortableProject.ManifestNamespace));
        }
        public void Dispose() { if (Directory.Exists(root)) Directory.Delete(root, true); }
        private sealed class Parts : IWorkbookParts
        {
            private readonly Dictionary<string, string> values = new Dictionary<string, string>();
            public int BlobReads; public bool FailDelete;
            private static string Namespace(string xml) { var doc = new XmlDocument(); doc.LoadXml(xml); return doc.DocumentElement.NamespaceURI; }
            public IEnumerable<string> Ids(string ns) => values.Where(v => Namespace(v.Value) == ns).Select(v => v.Key).ToList();
            public string Read(string id) { if (Namespace(values[id]) == PortableProject.BlobNamespace) BlobReads++; return values[id]; }
            public string Add(string xml) { var id = Guid.NewGuid().ToString("N"); values.Add(id, xml); return id; }
            public void Delete(string id) { if (FailDelete) throw new IOException("Simulated attachment cleanup failure"); values.Remove(id); }
        }
    }
}
