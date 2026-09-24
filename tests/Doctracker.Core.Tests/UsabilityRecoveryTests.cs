using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using Doctracker.Core.Geometry;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class UsabilityRecoveryTests : IDisposable
    {
        private readonly string root=Path.Combine(Path.GetTempPath(),"doctracker-v8-"+Guid.NewGuid().ToString("N"));
        private ProjectStore Store=>new ProjectStore(root);
        private ProjectState Fixture()
        {
            var page=new PageTextRecord {PageNumber=1,Text="Facture de BFA\u0001\nTotal : 6 489,83\u000b",Words=new List<WordRecord> {
                new WordRecord {Text="BFA\u0002",X=.1,Y=.2,Width=.2,Height=.1},new WordRecord {Text="6 489,83",X=.1,Y=.5,Width=.2,Height=.1} }};
            return new ProjectState {Documents=new List<DocumentRecord> {new DocumentRecord {Id="doc",PageCount=1,IndexComplete=true,IndexedPages=new List<PageTextRecord>{page}}},Snips=new List<SnipRecord> {new SnipRecord {Id="proof",DocumentId="doc",RawText="BFA\u0001",Comment="Revu\u0002"}}};
        }
        [Fact] public void Pdf_control_characters_do_not_break_saved_metadata_or_reloaded_index()
        {
            var store=Store;var state=Fixture();store.Save(state);
            using(var reader=XmlReader.Create(store.MetadataPath))while(reader.Read()){}
            var loaded=new ProjectStore(root).LoadOrCreate("");
            Assert.True(Store.ValidateIndex(loaded.Documents[0]));
            Assert.Single(OccurrenceSearch.Find(loaded,"BFA"));Assert.Single(OccurrenceSearch.Find(loaded,"6489,83"));
            Assert.Equal("BFA",loaded.Snips[0].RawText);
        }
        [Fact] public void Earlier_invalid_character_references_are_read_and_cleaned_without_losing_proofs()
        {
            var store=Store;var state=Fixture();store.Save(state);
            var xml="<ArrayOfPageTextRecord><PageTextRecord PageNumber=\"1\">Facture&#x1; BFA<Words><Word><Text>BFA&#xB;</Text><Line>1</Line><X>0.1</X><Y>0.2</Y><Width>0.2</Width><Height>0.1</Height></Word></Words></PageTextRecord></ArrayOfPageTextRecord>";
            Assert.Throws<XmlException>(()=>{using(var reader=XmlReader.Create(new StringReader(xml)))while(reader.Read()){};});
            WriteIndex(store.IndexPath(state.Documents[0].IndexKey),xml);
            var restored=new ProjectStore(root).LoadOrCreate("");
            Assert.True(new ProjectStore(root).ValidateIndex(restored.Documents[0]));
            var hit=Assert.Single(OccurrenceSearch.Find(restored,"BFA"));Assert.True(hit.HasLocation);Assert.Single(restored.Snips);
        }
        [Fact] public void Malformed_index_is_marked_for_rebuild_and_proofs_are_preserved()
        {
            var store=Store;var state=Fixture();store.Save(state);WriteIndex(store.IndexPath(state.Documents[0].IndexKey),"<ArrayOfPageTextRecord><broken>");
            var fresh=new ProjectStore(root);var loaded=fresh.LoadOrCreate("");Assert.False(fresh.ValidateIndex(loaded.Documents[0]));
            Assert.False(loaded.Documents[0].IndexComplete);Assert.Contains("Index",loaded.Documents[0].IndexError);Assert.Equal("proof",Assert.Single(loaded.Snips).Id);
        }
        [Fact] public void Dtd_and_external_entities_remain_forbidden()
        {
            var xml="<!DOCTYPE DoctrackerProject [<!ENTITY x SYSTEM 'file:///missing'>]><DoctrackerProject><WorkbookPath>&x;</WorkbookPath></DoctrackerProject>";
            Assert.ThrowsAny<Exception>(()=>ProjectStore.ReadMetadata(new MemoryStream(Encoding.UTF8.GetBytes(xml))));
        }
        [Theory] [InlineData("BFA")] [InlineData("FA")] [InlineData("B")] [InlineData("6489,83")] [InlineData("489.83")] [InlineData("X300")] [InlineData("*")]
        public void Interactive_search_finds_literal_occurrences_without_a_mode_switch(string query)
        {
            var state=Fixture();state.Documents[0].IndexedPages[0].Words.Clear();state.Documents[0].IndexedPages[0].Text="PréfixeBFA-suffixe Total=6 489,83EUR 500X300Z35 **";
            Assert.NotEmpty(OccurrenceSearch.Find(state,query));
        }
        [Theory] [InlineData("14/03/2026")] [InlineData("14-3-2026")] [InlineData("14.03.2026")]
        public void Excel_iso_date_query_finds_the_printed_document_date(string printed)
        {
            var state=Fixture();var page=state.Documents[0].IndexedPages[0];page.Words.Clear();page.Text="Date : "+printed;
            Assert.Single(OccurrenceSearch.Find(state,"2026-03-14"));
        }
        [Fact] public void Repeated_occurrences_on_one_page_have_distinct_locations()
        {
            var state=Fixture();state.Documents[0].IndexedPages[0].Words[1].Text="BFA";
            var hits=OccurrenceSearch.Find(state,"BFA");Assert.Equal(2,hits.Count);Assert.NotEqual(hits[0].Y,hits[1].Y);
            Assert.Single(OccurrenceSearch.Find(state,"BFA",1));
            Assert.Throws<OperationCanceledException>(()=>OccurrenceSearch.Find(state,"BFA",201,new CancellationToken(true)));
        }
        [Fact] public void Folders_create_move_rename_and_remove_without_changing_proofs()
        {
            var state=Fixture();var service=new ProjectFolders(Store);service.Create(state,"","Client 1");service.Create(state,"Client 1","BL");service.Create(state,"Client 1","Factures");
            service.Move(state,new[]{"doc"},"Client 1 / BL");Assert.Equal("Client 1 / BL",Assert.Single(state.Documents[0].Categories));
            service.Rename(state,"Client 1","Client A");Assert.Equal("Client A / BL",Assert.Single(state.Documents[0].Categories));
            var loaded=Store.LoadOrCreate("");Assert.Contains("Client A / Factures",ProjectFolders.Paths(loaded));
            service.Remove(state,"Client A");Assert.Empty(state.Documents[0].Categories);Assert.Empty(ProjectFolders.Paths(state));Assert.Equal("doc",state.Snips[0].DocumentId);
        }
        [Fact] public void Failed_folder_move_restores_membership()
        {
            var state=Fixture();var store=Store;var service=new ProjectFolders(store);service.Create(state,"","BL");service.Move(state,new[]{"doc"},"BL");
            File.Delete(store.MetadataPath);Directory.CreateDirectory(store.MetadataPath);
            Assert.ThrowsAny<Exception>(()=>service.Move(state,new[]{"doc"},""));Assert.Equal("BL",Assert.Single(state.Documents[0].Categories));
        }
        [Fact] public void Xref_edit_can_reuse_its_own_number_but_not_another_documents_number()
        {
            var state=Fixture();var doc=state.Documents[0];CrossReferences.Assign(state,doc,"DAC",1);
            var other=new DocumentRecord();CrossReferences.Assign(state,other,"DAC",2);
            CrossReferences.Reassign(state,doc,"DAC",3);Assert.Contains(1,CrossReferences.AvailableFor(state,"DAC",doc.Id));
            Assert.DoesNotContain(1,CrossReferences.Available(state,"DAC"));
            CrossReferences.Reassign(state,doc,"DAC",1);Assert.Equal(1,doc.ReferenceNumber);
            Assert.Throws<InvalidOperationException>(()=>CrossReferences.Reassign(state,doc,"DAC",2));Assert.Equal(1,doc.ReferenceNumber);
            CrossReferences.Reassign(state,doc,"DAC B",1);Assert.Equal("DAC B",doc.TestReference);Assert.Equal("proof",state.Snips[0].Id);
        }
        [Theory] [InlineData(0)] [InlineData(1)] [InlineData(2)] [InlineData(3)] [InlineData(4)] [InlineData(5)] [InlineData(6)] [InlineData(7)] [InlineData(8)]
        public void Graphical_comment_moves_and_resizes_stay_inside_the_page(int handle)
        {
            var original=new NormalizedRectangle(.2,.3,.4,.2);var updated=CommentGeometry.Transform(original,-2,3,handle);
            Assert.InRange(updated.X,0,1);Assert.InRange(updated.Y,0,1);Assert.True(updated.Width>0);Assert.True(updated.Height>0);
            Assert.True(updated.X+updated.Width<=1.000001);Assert.True(updated.Y+updated.Height<=1.000001);
            if(handle==0){Assert.Equal(original.Width,updated.Width);Assert.Equal(original.Height,updated.Height);}
        }
        private static void WriteIndex(string path,string xml){using(var stream=File.Create(path))using(var zip=new GZipStream(stream,CompressionMode.Compress))using(var writer=new StreamWriter(zip))writer.Write(xml);}
        public void Dispose(){if(Directory.Exists(root))Directory.Delete(root,true);}
    }
}
