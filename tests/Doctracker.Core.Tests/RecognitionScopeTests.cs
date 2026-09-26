using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class RecognitionScopeTests
    {
        private static DocumentRecord Document(string id,params string[] folders)=>new DocumentRecord {Id=id,Categories=folders.ToList()};
        [Fact] public void FolderIncludesChildrenButNotSimilarPrefixOrOtherCategories()
        {
            var state=new ProjectState {Documents=new List<DocumentRecord> {Document("bl","BL"),Document("child","BL / 2026"),Document("other","BLX"),Document("invoice","Factures"),Document("both","BL","Factures")}};
            Assert.Equal(new[]{"bl","child","both"},RecognitionScope.Select(state,"BL").Select(d=>d.Id));
        }
        [Fact] public void SelectionNeverAddsUnselectedDocumentsAndUnknownIdsAreIgnored()
        {
            var state=new ProjectState {Documents=new List<DocumentRecord> {Document("one","BL"),Document("two","BL"),Document("three","Factures")}};
            Assert.Equal(new[]{"one"},RecognitionScope.Select(state,"BL",new[]{"one","three","missing","one"}).Select(d=>d.Id));
            Assert.Empty(RecognitionScope.Select(state,null,new string[0]));
        }
        [Fact] public void NullMeansAllAndEmptyFolderMeansUnfiled()
        {
            var state=new ProjectState {Documents=new List<DocumentRecord> {Document("one"),Document("two","Factures")}};
            Assert.Equal(2,RecognitionScope.Select(state,null).Count);
            Assert.Equal("one",Assert.Single(RecognitionScope.Select(state,"")).Id);
        }
        [Fact] public void PreparedOnlyMatchingDoesNotLoadUnselectedDocument()
        {
            var ready=Document("ready","BL");ready.IndexComplete=true;
            ready.IndexedPages=new List<PageTextRecord> {new PageTextRecord {PageNumber=1,Text="FAC12345A"}};
            var other=Document("other","Factures");other.PageLoader=()=>throw new System.Exception("Out-of-scope index loaded");
            var scope=new ProjectState {Documents=RecognitionScope.Select(new ProjectState {Documents=new List<DocumentRecord>{ready,other}},"BL")};
            Assert.Single(new DocumentMatcher().FindBatch(scope,new[]{(IReadOnlyList<string>)new[]{"12345"}},true,CancellationToken.None)[0]);
        }
        [Fact] public void ExpandingPreparedScopeDetectsNewAmbiguity()
        {
            var one=Document("one");one.IndexedPages=new List<PageTextRecord>{new PageTextRecord {PageNumber=1,Text="FAC12345A"}};
            var two=Document("two");two.IndexedPages=new List<PageTextRecord>{new PageTextRecord {PageNumber=4,Text="FAC12345A"}};
            var state=new ProjectState {Documents=new List<DocumentRecord>{one}};
            var criteria=new[]{(IReadOnlyList<string>)new[]{"12345"}};var matcher=new DocumentMatcher();
            Assert.Single(matcher.FindBatch(state,criteria,true,CancellationToken.None)[0]);
            state.Documents.Add(two);
            Assert.Equal(2,matcher.FindBatch(state,criteria,true,CancellationToken.None)[0].Count);
        }
    }
}
