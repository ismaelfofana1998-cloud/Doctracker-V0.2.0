using System;
using System.IO;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class ProjectStoreTests
    {
        [Fact]
        public void Project_round_trip_preserves_snips_and_audit_trail()
        {
            var directory = Path.Combine(Path.GetTempPath(), "doctracker-tests", Guid.NewGuid().ToString("N"));
            try
            {
                var store = new ProjectStore(directory);
                var state = store.LoadOrCreate(@"C:\Audit\Mission.xlsx");
                state.Documents.Add(new DocumentRecord
                {
                    Id = "doc-1",
                    OriginalName = "invoice.pdf",
                    RelativePath = Path.Combine("documents", "doc-1.pdf")
                });
                state.Snips.Add(new SnipRecord
                {
                    Id = "snip-1",
                    DocumentId = "doc-1",
                    WorksheetName = "Achats",
                    CellAddress = "F12",
                    ExtractedValue = "1250"
                });
                state.AuditTrail.Add(new AuditEventRecord
                {
                    Action = "SnipCreated",
                    EntityId = "snip-1"
                });
                store.Save(state);

                var loaded = store.LoadOrCreate(@"C:\Audit\Mission.xlsx");

                Assert.Single(loaded.Documents);
                Assert.Single(loaded.Snips);
                Assert.Single(loaded.AuditTrail);
                Assert.Equal("F12", loaded.Snips[0].CellAddress);
            }
            finally
            {
                if (Directory.Exists(directory)) Directory.Delete(directory, true);
            }
        }

        [Fact]
        public void Resolver_rejects_a_path_outside_the_project()
        {
            var store = new ProjectStore(Path.Combine(Path.GetTempPath(), "doctracker-safe"));
            var document = new DocumentRecord { RelativePath = Path.Combine("..", "secret.pdf") };
            Assert.Throws<InvalidOperationException>(() => store.ResolveDocumentPath(document));
        }
    }
}
