using System;
using System.IO;
using System.Linq;
using System.Threading;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class ImportBatchTests : IDisposable
    {
        private readonly string root = Path.Combine(Path.GetTempPath(), "doctracker-import-" + Guid.NewGuid().ToString("N"));
        private string[] Sources(int count)
        {
            Directory.CreateDirectory(root);
            return Enumerable.Range(0, count).Select(i => { var path = Path.Combine(root, i + ".pdf"); File.WriteAllText(path, "Synthetic document " + i); return path; }).ToArray();
        }
        [Fact]
        public void Import_is_sequential_without_loading_existing_indexes_and_commits_in_batches()
        {
            var files = Sources(61); var store = new ProjectStore(Path.Combine(root,"project")); var state = new ProjectState();
            state.Documents.Add(new DocumentRecord { IndexKey = Guid.NewGuid().ToString("N"), IndexComplete = true,
                PageLoader = () => throw new Exception("An import must not load an existing index") });
            var importer = new DocumentImporter(store); var saves = 0; store.Saved += () => saves++;
            var result = DocumentImportBatch.Run(store, state, files, p => importer.Import(state,p,"test",null,false),null,default);
            Assert.Equal(61,result.ImportedCount);Assert.Equal(3,saves);Assert.Empty(result.Errors);
            Assert.All(state.Documents.Skip(1),d => { Assert.False(d.IndexComplete);Assert.Equal("",d.IndexKey); });
            Assert.Equal(state.Documents.Last().Id,result.LastDocumentId);
            Assert.Equal(62,store.LoadOrCreate("").Documents.Count);
        }
        [Fact]
        public void Cancellation_saves_completed_files_once_and_does_not_start_the_next_file()
        {
            var files=Sources(9);var store=new ProjectStore(Path.Combine(root,"project"));var state=new ProjectState();var importer=new DocumentImporter(store);
            var cancellation=new CancellationTokenSource();var calls=0;var saves=0;store.Saved+=()=>saves++;
            var result=DocumentImportBatch.Run(store,state,files,p=>{
                var doc=importer.Import(state,p,"test",null,false);if(++calls==3)cancellation.Cancel();return doc;
            },null,cancellation.Token);
            Assert.True(result.Cancelled);Assert.Equal(3,calls);Assert.Equal(1,saves);Assert.Equal(3,store.LoadOrCreate("").Documents.Count);
        }
        [Fact]
        public void Failed_files_do_not_trigger_repeated_saves_or_stop_the_remaining_imports()
        {
            var files=Sources(27);var store=new ProjectStore(Path.Combine(root,"project"));var state=new ProjectState();var importer=new DocumentImporter(store);
            var saves=0;store.Saved+=()=>saves++;
            var result=DocumentImportBatch.Run(store,state,files,p=>{
                if(p==files[25])throw new IOException("Unavailable file");return importer.Import(state,p,"test",null,false);
            },null,default);
            Assert.Equal(26,result.ImportedCount);Assert.Single(result.Errors);Assert.Equal(2,saves);
        }
        [Fact]
        public void Failed_checkpoint_restores_previous_documents_and_duplicate_metadata_without_retrying_save()
        {
            var files=Sources(2);var store=new ProjectStore(Path.Combine(root,"project"));var state=new ProjectState();var importer=new DocumentImporter(store);
            var original=importer.Import(state,files[0],"test","BL");var date=original.LastImportedAtUtc;var revision=state.Revision;
            using(var locked=new FileStream(store.MetadataPath,FileMode.Open,FileAccess.Read,FileShare.None))
                Assert.ThrowsAny<IOException>(()=>DocumentImportBatch.Run(store,state,files,p=>importer.Import(state,p,"test","Factures",false),null,default));
            Assert.Single(state.Documents);Assert.Single(state.AuditTrail);Assert.Equal(date,original.LastImportedAtUtc);
            Assert.Equal(new[]{"BL"},original.Categories);Assert.Equal(revision,state.Revision);Assert.Single(store.LoadOrCreate("").Documents);
        }
        public void Dispose(){if(Directory.Exists(root))Directory.Delete(root,true);}
    }
}
