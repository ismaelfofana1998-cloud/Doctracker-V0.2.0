using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Doctracker.Core.Geometry;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class DeferredEditingTests : IDisposable
    {
        private readonly string root=Path.Combine(Path.GetTempPath(),"doctracker-edit-"+Guid.NewGuid().ToString("N"));
        [Fact] public void Deferred_changes_write_no_metadata_or_history_until_workbook_flush()
        {
            var store=new ProjectStore(root){DeferMetadataWrites=true};var state=new ProjectState();var events=0;
            store.Saved+=()=>events++;
            for(var i=0;i<25;i++){state.TestReference="TEST "+i;store.Save(state);}
            Assert.Equal(25,events);Assert.False(File.Exists(store.MetadataPath));Assert.False(Directory.Exists(Path.Combine(root,"recovery")));
            store.Flush(state);Assert.Equal("TEST 24",new ProjectStore(root).LoadOrCreate("").TestReference);
            state.TestReference="Non enregistré";store.Save(state);
            Assert.Equal("TEST 24",new ProjectStore(root).LoadOrCreate("").TestReference);
        }
        [Fact] public void Deferred_import_publishes_one_change_at_end_of_batch()
        {
            var store=new ProjectStore(root){DeferMetadataWrites=true};var state=new ProjectState();var changes=0;store.Saved+=()=>changes++;
            var result=DocumentImportBatch.Run(store,state,Enumerable.Range(0,100).Select(i=>i.ToString()),path=>{
                var doc=new DocumentRecord {OriginalName=path};state.Documents.Add(doc);return doc;
            },null,System.Threading.CancellationToken.None);
            Assert.Equal(100,result.ImportedCount);Assert.Equal(1,changes);Assert.False(File.Exists(store.MetadataPath));
        }
        [Fact] public void Deferred_indexes_are_reusable_without_rewriting_the_project()
        {
            var store=new ProjectStore(root){DeferMetadataWrites=true};var state=new ProjectState();
            var doc=new DocumentRecord {PageCount=1,IndexComplete=true,IndexedPages=new List<PageTextRecord>{new PageTextRecord {PageNumber=1,Text="ABC123XYZ"}}};state.Documents.Add(doc);
            store.Save(state);var key=doc.IndexKey;doc.ReleaseIndex();
            Assert.Single(OccurrenceSearch.Find(state,"123"));Assert.False(File.Exists(store.MetadataPath));
            store.Save(state);Assert.Equal(key,doc.IndexKey);
            store.Flush(state);var loaded=new ProjectStore(root).LoadOrCreate("");Assert.Single(OccurrenceSearch.Find(loaded,"123"));
        }
        [Theory] [InlineData(SnipType.Text,"Nouveau texte","Nouveau texte")] [InlineData(SnipType.Sum,"10  25","35")] [InlineData(SnipType.Number,"125,50","125.5")]
        public void Resizing_updates_value_without_changing_link_identity_and_resets_review(SnipType type,string text,string expected)
        {
            var store=new ProjectStore(root){DeferMetadataWrites=true};var state=new ProjectState();var doc=new DocumentRecord();state.Documents.Add(doc);
            var snip=new SnipRecord {Type=type,DocumentId=doc.Id,Id="same-proof",X=.1,Y=.1,Width=.2,Height=.1,RawText="10",ExtractedValue="10",Status=ReviewStatus.Reviewed};state.Snips.Add(snip);
            var service=new SnipService(store,new TextValueParser());var callback=false;
            service.UpdateGeometry(state,snip.Id,new NormalizedRectangle(.2,.3,.4,.2),text,"test",updated=>{callback=true;Assert.Equal(expected,updated.ExtractedValue);});
            Assert.True(callback);Assert.Equal("same-proof",snip.Id);Assert.Equal(.2,snip.X);Assert.Equal(expected,snip.ExtractedValue);Assert.Equal(ReviewStatus.Prepared,snip.Status);
        }
        [Fact] public void Failed_cell_update_restores_old_geometry_text_and_review()
        {
            var store=new ProjectStore(root);var state=new ProjectState();var snip=new SnipRecord {Type=SnipType.Sum,X=.1,Y=.1,Width=.2,Height=.1,RawText="10  20",ExtractedValue="30",Status=ReviewStatus.Reviewed};state.Snips.Add(snip);
            Assert.Throws<IOException>(()=>new SnipService(store,new TextValueParser()).UpdateGeometry(state,snip.Id,new NormalizedRectangle(.3,.3,.2,.1),"40  50","test",s=>throw new IOException("Excel write failed")));
            Assert.Equal(.1,snip.X);Assert.Equal("30",snip.ExtractedValue);Assert.Equal("10  20",snip.RawText);Assert.Equal(ReviewStatus.Reviewed,snip.Status);Assert.Empty(state.AuditTrail);
        }
        [Fact] public void Failed_flush_leaves_pending_changes_available_for_retry()
        {
            var store=new ProjectStore(root){DeferMetadataWrites=true};var state=new ProjectState {TestReference="Keep me"};store.Save(state);
            Directory.CreateDirectory(store.MetadataPath);Assert.ThrowsAny<Exception>(()=>store.Flush(state));
            Assert.Equal("Keep me",state.TestReference);Directory.Delete(store.MetadataPath);store.Flush(state);
            Assert.Equal("Keep me",new ProjectStore(root).LoadOrCreate("").TestReference);
        }
        public void Dispose(){if(Directory.Exists(root))Directory.Delete(root,true);}
    }
}
