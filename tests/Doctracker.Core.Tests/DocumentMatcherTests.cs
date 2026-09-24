using System.Collections.Generic;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class DocumentMatcherTests
    {
        [Theory]
        [InlineData("Facture de BFA", "BFA", false)]
        [InlineData("Code : 500X300Z35", "X300", true)]
        [InlineData("Total HT : 6489,83 EUR", "6489,83", false)]
        [InlineData("Base TVA : 6 489,83", "6489,83", false)]
        [InlineData("Net à payer *******7787,80 EUR", "7787,80", false)]
        public void Finds_tokens_inside_a_positioned_block_even_when_page_text_is_incomplete(string text,string query,bool partial)
        {
            var page=new PageTextRecord {PageNumber=1,Text="Texte global incomplet",Words=new List<WordRecord> {
                new WordRecord {Text=text,Line=1,X=.2,Y=.3,Width=.4,Height=.1} }};
            var state=new ProjectState {Documents=new List<DocumentRecord> {new DocumentRecord {IndexedPages=new List<PageTextRecord>{page}}}};
            var hit=Assert.Single(new DocumentMatcher().Find(state,query,50,partial));
            Assert.True(hit.HasLocation);Assert.Equal(.2,hit.X);Assert.True(hit.IsExact||hit.IsPartial);
        }
        [Theory]
        [InlineData("6489,83", "6489,83")]
        [InlineData("6 489,83", "6489.83")]
        [InlineData("6\u202f489,83", "6489,83")]
        public void Finds_invoice_amounts_inside_plain_text(string amount,string query)
        {
            var state=TextState("Total HT : "+amount+" EUR   Total CA : 16245,77");
            Assert.True(Assert.Single(new DocumentMatcher().Find(state,query)).IsExact);
        }
        [Theory]
        [InlineData("16489,83")]
        [InlineData("6489,8301")]
        [InlineData("-6489,83")]
        [InlineData("(6489,83)")]
        [InlineData("6489,83-")]
        [InlineData("FA-6489,83-01")]
        public void Never_matches_a_different_or_negative_amount(string text)
        {
            Assert.Empty(new DocumentMatcher().Find(TextState("Total : "+text+" EUR"),"6489,83",50,true));
        }
        private static ProjectState TextState(string text) => new ProjectState {Documents=new List<DocumentRecord> {
            new DocumentRecord {IndexedPages=new List<PageTextRecord> {new PageTextRecord {PageNumber=1,Text=text}}} }};

        [Fact]
        public void Matching_prioritizes_invoice_reference_and_amount()
        {
            var expected = new DocumentRecord
            {
                Id = "invoice-a",
                OriginalName = "FA-2025-0198.pdf",
                IndexedPages = new List<PageTextRecord>
                {
                    new PageTextRecord
                    {
                        PageNumber = 1,
                        Text = "FACTURE FA-2025-0198 Client Alpha Total TTC 12 450,00 EUR"
                    }
                }
            };
            var other = new DocumentRecord
            {
                Id = "invoice-b",
                OriginalName = "FA-2025-0199.pdf",
                IndexedPages = new List<PageTextRecord>
                {
                    new PageTextRecord
                    {
                        PageNumber = 1,
                        Text = "FACTURE FA-2025-0199 Client Beta Total TTC 4 000,00 EUR"
                    }
                }
            };
            var state = new ProjectState
            {
                Documents = new List<DocumentRecord> { other, expected }
            };

            var result = new DocumentMatcher().Find(
                state, "FA-2025-0198 12 450,00", maximum: 2);

            Assert.NotEmpty(result);
            Assert.Equal("invoice-a", result[0].DocumentId);
            Assert.True(result[0].Score >= 0.55);
        }

        [Fact]
        public void Matching_returns_no_candidate_for_blank_input()
        {
            var result = new DocumentMatcher().Find(new ProjectState(), " ");
            Assert.Empty(result);
        }

        [Fact]
        public void Matching_returns_no_candidate_for_punctuation_only_input()
        {
            var state = new ProjectState
            {
                Documents = new List<DocumentRecord>
                {
                    new DocumentRecord
                    {
                        IndexedPages = new List<PageTextRecord>
                        {
                            new PageTextRecord { PageNumber = 1, Text = "Facture fournisseur" }
                        }
                    }
                }
            };

            var result = new DocumentMatcher().Find(state, "€ / -");

            Assert.Empty(result);
        }
    }
}
