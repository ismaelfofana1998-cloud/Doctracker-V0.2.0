using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class FirstTextMatchingTests
    {
        private static DocumentRecord Document(params string[] text) => new DocumentRecord {
            IndexedPages=text.Select((s,i)=>new PageTextRecord {PageNumber=i+1,Text=s}).ToList()};
        private static IReadOnlyList<MatchCandidate>[] Find(ProjectState state,params string[][] rows) =>
            new DocumentMatcher().FindFirstTextBatch(state,rows,CancellationToken.None);

        [Fact]
        public void Duplicates_retain_first_document_page_and_occurrence_without_loading_later_documents()
        {
            var first=Document("REF001A REF001B","REF001C");
            first.IndexedPages.Reverse(); // Stored page order must not change the first hit.
            var later=new DocumentRecord {PageLoader=()=>throw new Exception("Should stop after first hit")};
            var hit=Assert.Single(Find(new ProjectState {Documents=new List<DocumentRecord>{first,later}},new[]{"001"})[0]);
            Assert.Equal(first.Id,hit.DocumentId);Assert.Equal(1,hit.PageNumber);
            Assert.Equal("REF001A",hit.Fields[0].Evidence);
        }

        [Theory]
        [InlineData("Total -16489,83 EUR","6489,83","-16489,83")]
        [InlineData("FA-00123-04","00123","FA-00123-04")]
        [InlineData("REF 000123","000123","000123")]
        [InlineData("Total 6 489,83 EUR","6489.83","6 489,83")]
        [InlineData("REF PréfixeBFA-suffixe","bfa","PréfixeBFA-suffixe")]
        public void Values_are_text_fragments_without_amount_or_identifier_validation(string source,string query,string evidence)
        {
            var state=new ProjectState {Documents=new List<DocumentRecord>{Document(source)}};
            Assert.Equal(evidence,Assert.Single(Find(state,new[]{query})[0]).Fields[0].Evidence);
        }

        [Fact]
        public void Dates_and_leading_zeroes_are_not_reinterpreted()
        {
            var state=new ProjectState {Documents=new List<DocumentRecord>{Document("2026-09-26 REF123")}};
            var hits=Find(state,new[]{"26/09/2026"},new[]{"00123"},new[]{"2026-09-26"});
            Assert.Empty(hits[0]);Assert.Empty(hits[1]);Assert.Single(hits[2]);
        }

        [Fact]
        public void Ten_found_rows_are_retained_when_ten_other_rows_are_missing()
        {
            var loads=0;
            var document=new DocumentRecord {PageLoader=()=>{loads++;return new List<PageTextRecord>{new PageTextRecord {
                PageNumber=1,Text=string.Join(" ",Enumerable.Range(0,10).Select(i=>"REF"+i.ToString("D3")))}};}};
            var rows=Enumerable.Range(0,20).Select(i=>new[]{"REF"+i.ToString("D3")}).ToArray();
            var hits=Find(new ProjectState {Documents=new List<DocumentRecord>{document}},rows);
            Assert.All(hits.Take(10),h=>Assert.Single(h));Assert.All(hits.Skip(10),h=>Assert.Empty(h));Assert.Equal(1,loads);
        }

        [Fact]
        public void Criteria_of_a_row_share_a_page_and_blank_columns_keep_their_position()
        {
            var state=new ProjectState {Documents=new List<DocumentRecord>{Document("FA-01","CLIENT-X"),Document("FA-01 CLIENT-X")}};
            var hit=Assert.Single(Find(state,new[]{"FA-01"," ","CLIENT-X"})[0]);
            Assert.Equal(state.Documents[1].Id,hit.DocumentId);Assert.Equal(3,hit.Fields.Count);Assert.Null(hit.Fields[1]);
            Assert.Empty(Find(state,new[]{" "})[0]);
        }

        [Fact]
        public void First_hit_keeps_word_coordinates_and_cancellation_still_interrupts_matching()
        {
            var doc=Document("REF123 REF123");
            doc.IndexedPages[0].Words.Add(new WordRecord {Text="REF123",X=.1,Y=.2,Width=.3,Height=.04});
            var state=new ProjectState {Documents=new List<DocumentRecord>{doc}};
            var hit=Assert.Single(Find(state,new[]{"123"})[0]).Fields[0];
            Assert.True(hit.HasLocation);Assert.Equal(.1,hit.X);Assert.Equal(.2,hit.Y);
            Assert.Throws<OperationCanceledException>(()=>new DocumentMatcher().FindFirstTextBatch(state,new[]{new[]{"123"}},new CancellationToken(true)));
        }
    }
}
