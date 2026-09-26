using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Xml;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class PortableProjectTests : IDisposable
    {
        private readonly string root=Path.Combine(Path.GetTempPath(),"doctracker-portable-"+Guid.NewGuid().ToString("N"));
        private ProjectStore Store(string name)=>new ProjectStore(Path.Combine(root,name));
        private ProjectState Fixture(ProjectStore store)
        {
            Directory.CreateDirectory(root);var source=Path.Combine(root,"invoice.pdf");File.WriteAllBytes(source,Enumerable.Range(0,2300000).Select(i=>(byte)(i%251)).ToArray());
            var state=new ProjectState();var doc=new DocumentImporter(store).Import(state,source,"test","Client 1 / Factures");
            doc.IndexedPages=new List<PageTextRecord>{new PageTextRecord {PageNumber=1,Text="Facture 500X300 Z35 total 1500,00"}};
            doc.IndexComplete=true;doc.PageCount=1;
            state.Snips.Add(new SnipRecord {Id="proof",DocumentId=doc.Id,PageNumber=1,WorksheetName="Achats",CellAddress="B2"});
            state.CellLinks.Add(new CellLinkRecord {WorksheetName="Achats",CellAddress="B9",SnipIds=new List<string>{"proof"}});
            store.Save(state);return state;
        }
        [Fact] public void Workbook_transfer_restores_identical_sources_links_and_lazy_indexes()
        {
            var store=Store("pc1");var state=Fixture(store);var parts=new MemoryParts();new PortableProject(parts).Save(store,state);
            Directory.Delete(store.ProjectDirectory,true);
            var target=Store("pc2");var restored=new PortableProject(parts).Restore(target);
            Assert.Equal(state.ProjectId,restored.ProjectId);Assert.Equal("B9",restored.CellLinks[0].CellAddress);
            Assert.False(File.Exists(target.LocalDocumentPath(restored.Documents[0])));
            ProjectStore.VerifyHash(restored.Documents[0],target.ResolveDocumentPath(restored.Documents[0]));
            Assert.Contains("X300",restored.Documents[0].IndexedPages[0].Text);
            Assert.Equal("Client 1 / Factures",restored.Documents[0].Categories[0]);
        }
        [Fact] public void Unchanged_sources_are_not_reembedded_on_each_save()
        {
            var store=Store("first");var state=Fixture(store);var parts=new MemoryParts();var package=new PortableProject(parts);package.Save(store,state);
            var blobs=parts.Ids(PortableProject.BlobNamespace).OrderBy(x=>x).ToArray();state.Snips[0].Comment="Reviewed";store.Save(state);package.Save(store,state);
            Assert.Equal(blobs,parts.Ids(PortableProject.BlobNamespace).OrderBy(x=>x));
        }
        [Fact] public void Failed_embedding_preserves_previous_manifest_and_discards_new_chunks()
        {
            var store=Store("fail");var state=Fixture(store);var parts=new MemoryParts();var package=new PortableProject(parts);package.Save(store,state);
            var ids=parts.Values.Keys.OrderBy(x=>x).ToArray();parts.FailAdd=true;state.Snips[0].Comment="Not saved";
            Assert.Throws<IOException>(()=>package.Save(store,state));Assert.Equal(ids,parts.Values.Keys.OrderBy(x=>x));
            Assert.Equal("",new PortableProject(parts).Restore(Store("recovered")).Snips[0].Comment);
        }
        [Fact] public void Concurrent_manifest_updates_are_not_silently_overwritten()
        {
            var store=Store("conflict");var state=Fixture(store);var parts=new MemoryParts();var a=new PortableProject(parts);a.Save(store,state);var b=new PortableProject(parts);a.Save(store,state);
            Assert.Throws<IOException>(()=>b.Save(store,state));
        }
        [Fact] public void Corrupt_embedded_bytes_are_rejected()
        {
            var store=Store("corrupt");var state=Fixture(store);var parts=new MemoryParts();new PortableProject(parts).Save(store,state);
            var id=parts.Ids(PortableProject.BlobNamespace).First();parts.Values[id]="<blob xmlns=\""+PortableProject.BlobNamespace+"\">AAAA</blob>";
            var target=Store("restore");var restored=new PortableProject(parts).Restore(target);
            Assert.Throws<InvalidDataException>(()=>target.ResolveDocumentPath(restored.Documents[0]));
        }
        [Fact] public void Backup_recovers_from_loss_of_workbook_and_cache()
        {
            var store=Store("backup");var state=Fixture(store);var path=Path.Combine(root,"backup.dtpack");RecoveryArchive.Export(store,state,path,CancellationToken.None);Directory.Delete(store.ProjectDirectory,true);
            var target=Store("new");var restored=RecoveryArchive.Restore(path,target,CancellationToken.None);
            Assert.Single(restored.Snips);ProjectStore.VerifyHash(restored.Documents[0],target.ResolveDocumentPath(restored.Documents[0]));
        }
        [Fact] public void Archive_path_traversal_is_rejected()
        {
            Directory.CreateDirectory(root);var archive=Path.Combine(root,"bad.zip");using(var zip=ZipFile.Open(archive,ZipArchiveMode.Create))using(var file=zip.CreateEntry("../escape.txt").Open())file.WriteByte(1);
            Assert.Throws<InvalidDataException>(()=>RecoveryArchive.Restore(archive,Store("bad"),CancellationToken.None));Assert.False(File.Exists(Path.Combine(root,"escape.txt")));
        }
        [Fact] public void Xrefs_are_unique_persisted_and_do_not_rename_originals()
        {
            var store=Store("xref");var state=Fixture(store);var doc=state.Documents[0];CrossReferences.Assign(state,doc,"DAC B 30 040",1);store.Save(state);
            Assert.Equal("DAC B 30 040 - 01 - invoice.pdf",doc.DisplayName);Assert.Equal(2,CrossReferences.Next(state,"dac b 30 040"));
            Assert.Throws<InvalidOperationException>(()=>CrossReferences.Assign(state,new DocumentRecord(),"DAC B 30 040",1));Assert.True(File.Exists(store.ResolveDocumentPath(doc)));
            Assert.Single(store.LoadOrCreate("").XrefReservations);
        }
        [Fact] public void Shared_mode_transfers_references_without_embedding_source_bytes()
        {
            var store=Store("shared");var state=Fixture(store);var vault=Path.Combine(root,"vault");SharedVault.Publish(vault,store,state);state.SharedVaultPath=vault;store.Save(state);
            var parts=new MemoryParts();new PortableProject(parts).Save(store,state);
            Assert.Single(parts.Ids(PortableProject.BlobNamespace)); // index only, source is three chunks
            var target=Store("colleague");var restored=new PortableProject(parts).Restore(target);ProjectStore.VerifyHash(restored.Documents[0],target.ResolveDocumentPath(restored.Documents[0]));
        }
        [Fact] public void Shared_xref_reservation_rejects_competing_document()
        {
            var vault=Path.Combine(root,"reservations");SharedVault.Reserve(vault,"project","DAC",1,"doc-a");
            Assert.Throws<IOException>(()=>SharedVault.Reserve(vault,"project","DAC",1,"doc-b"));
        }
        [Fact] public void Metadata_corruption_recovers_previous_revision()
        {
            var store=Store("metadata");var state=Fixture(store);store.Save(state);File.WriteAllText(store.MetadataPath,"not xml");
            Assert.Single(store.LoadOrCreate("").Documents);Assert.NotNull(store.RecoveryNotice);
        }
        [Fact] public void Index_is_external_and_loaded_only_when_searched()
        {
            var store=Store("index");var state=Fixture(store);Assert.DoesNotContain("500X300",File.ReadAllText(store.MetadataPath));
            var restored=store.LoadOrCreate("");Assert.NotNull(restored.Documents[0].PageLoader);Assert.Single(new DocumentMatcher().FindAllFields(restored,new[]{"X300"},true));
        }
        [Theory][InlineData("X300")][InlineData("X300Z35")][InlineData("500X300 Z35")][InlineData("500X300Z35")]
        public void Partial_references_find_spaced_or_embedded_identifiers(string query)
        {
            var store=Store("partial");var state=Fixture(store);var match=Assert.Single(new DocumentMatcher().FindAllFields(state,new[]{query},true));Assert.Contains("X300",match.Fields[0].Evidence);
        }
        [Fact] public void Partial_matching_accepts_numeric_references_and_rejects_ambiguity()
        {
            var store=Store("strict");var state=Fixture(store);var matcher=new DocumentMatcher();Assert.Single(matcher.FindAllFields(state,new[]{"500"},true));
            state.Documents.Add(new DocumentRecord {IndexedPages=new List<PageTextRecord>{new PageTextRecord {PageNumber=1,Text="500X300Z35"}}});Assert.Equal(2,matcher.FindAllFields(state,new[]{"X300"},true).Count);
        }
        [Fact] public void Batch_scans_each_lazy_index_once_and_preserves_cancellation()
        {
            var loads=0;var document=new DocumentRecord {PageLoader=()=>{loads++;return new List<PageTextRecord>{new PageTextRecord {PageNumber=1,Text="Reference X300"}};}};
            var state=new ProjectState {Documents=new List<DocumentRecord>{document}};
            var rows=Enumerable.Range(0,1000).Select(i=>(IReadOnlyList<string>)new[]{"X300"}).ToList();var result=new DocumentMatcher().FindBatch(state,rows,true,CancellationToken.None);
            Assert.Equal(1,loads);Assert.All(result,x=>Assert.Single(x));
            Assert.Throws<OperationCanceledException>(()=>new DocumentMatcher().FindBatch(state,rows,true,new CancellationToken(true)));
        }
        public void Dispose(){if(Directory.Exists(root))Directory.Delete(root,true);}
        private sealed class MemoryParts : IWorkbookParts
        {
            public Dictionary<string,string> Values {get;}=new Dictionary<string,string>();public bool FailAdd;
            public IEnumerable<string> Ids(string ns)=>Values.Where(x=>{var d=new XmlDocument();d.LoadXml(x.Value);return d.DocumentElement.NamespaceURI==ns;}).Select(x=>x.Key).ToList();
            public string Read(string id)=>Values[id];
            public string Add(string xml){if(FailAdd)throw new IOException("Simulated COM/storage failure");var id=Guid.NewGuid().ToString();Values.Add(id,xml);return id;}
            public void Delete(string id)=>Values.Remove(id);
        }
    }
}
