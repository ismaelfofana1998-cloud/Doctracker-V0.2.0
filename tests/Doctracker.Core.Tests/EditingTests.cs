using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Doctracker.Core.Geometry;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class EditingTests : IDisposable
    {
        private readonly string root = Path.Combine(Path.GetTempPath(), "doctracker-edit-" + Guid.NewGuid().ToString("N"));
        private ProjectStore Store => new ProjectStore(root);
        private static ProjectState LinkedState()
        {
            return new ProjectState {
                Documents = new List<DocumentRecord> { new DocumentRecord { Id="doc", PageCount=2 } },
                Snips = new List<SnipRecord> { new SnipRecord {Id="a",DocumentId="doc"},new SnipRecord {Id="b",DocumentId="doc"} },
                CellLinks = new List<CellLinkRecord> {
                    new CellLinkRecord {CellAddress="B2",SnipIds=new List<string>{"a","b"}},
                    new CellLinkRecord {CellAddress="B9",SnipIds=new List<string>{"a"}} }
            };
        }
        [Fact] public void Deleting_a_snip_cleans_all_copies_and_preserves_other_proofs()
        {
            var state=LinkedState();var store=Store;
            new SnipService(store,new TextValueParser()).Delete(state,"a","tester");
            var restored=store.LoadOrCreate("");Assert.Equal("b",Assert.Single(restored.Snips).Id);
            Assert.Equal("b",Assert.Single(Assert.Single(restored.CellLinks).SnipIds));
            Assert.Equal("SnipDeleted",Assert.Single(restored.AuditTrail).Action);
        }
        [Fact] public void Failed_deletion_restores_proofs_links_and_audit()
        {
            var state=LinkedState();var store=Store;Directory.CreateDirectory(store.MetadataPath);
            Assert.ThrowsAny<Exception>(()=>new SnipService(store,new TextValueParser()).Delete(state,"a","tester"));
            Assert.Equal(new[]{"a","b"},state.Snips.Select(s=>s.Id));Assert.Equal(2,state.CellLinks.Count);Assert.Empty(state.AuditTrail);
        }
        [Fact] public void Comments_and_xref_survive_full_backup_and_can_be_edited_and_deleted()
        {
            var store=Store;Directory.CreateDirectory(root);var source=Path.Combine(root,"source.pdf");File.WriteAllText(source,"fixture");
            var state=new ProjectState();var doc=new DocumentImporter(store).Import(state,source,"tester");doc.PageCount=2;
            CrossReferences.Assign(state,doc,"DAC B 30 040",1);
            var service=new DocumentCommentService(store);var zone=new NormalizedRectangle(.2,.3,.5,.2);
            var comment=service.Save(state,doc.Id,2,zone,"Pièce à rapprocher");
            service.Save(state,doc.Id,2,zone,"Rapprochement validé",comment.Id);
            var archive=Path.Combine(root,"mission.dtpack");RecoveryArchive.Export(store,state,archive,CancellationToken.None);
            var target=new ProjectStore(Path.Combine(root,"restore"));var restored=RecoveryArchive.Restore(archive,target,CancellationToken.None);
            Assert.Equal("Rapprochement validé",Assert.Single(restored.Documents[0].Comments).Text);Assert.Equal("DAC B 30 040",restored.Documents[0].TestReference);
            new DocumentCommentService(target).Delete(restored,doc.Id,comment.Id);Assert.Empty(target.LoadOrCreate("").Documents[0].Comments);
        }
        [Fact] public void Comment_edits_roll_back_on_storage_failure()
        {
            var store=Store;var state=LinkedState();var service=new DocumentCommentService(store);var zone=new NormalizedRectangle(.1,.1,.4,.3);
            var comment=service.Save(state,"doc",1,zone,"Original");File.Delete(store.MetadataPath);Directory.CreateDirectory(store.MetadataPath);
            Assert.ThrowsAny<Exception>(()=>service.Save(state,"doc",1,zone,"Modified",comment.Id));
            Assert.Equal("Original",Assert.Single(state.Documents[0].Comments).Text);
            Assert.ThrowsAny<Exception>(()=>service.Delete(state,"doc",comment.Id));Assert.Single(state.Documents[0].Comments);
        }
        [Fact] public void Reimport_is_deduplicated_and_refreshes_import_priority()
        {
            Directory.CreateDirectory(root);var source=Path.Combine(root,"fixture.pdf");File.WriteAllText(source,"fixture");
            var state=new ProjectState();var importer=new DocumentImporter(Store);var doc=importer.Import(state,source,"test");
            doc.LastImportedAtUtc=new DateTime(2000,1,1);var imported=importer.Import(state,source,"test");
            Assert.Same(doc,imported);Assert.Single(state.Documents);Assert.True(imported.LastImportedAtUtc.Year>2000);
        }
        public void Dispose() {if(Directory.Exists(root))Directory.Delete(root,true);}
    }
}
