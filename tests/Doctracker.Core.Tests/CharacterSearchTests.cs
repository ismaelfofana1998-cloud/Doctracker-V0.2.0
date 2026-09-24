using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace Doctracker.Core.Tests
{
    public sealed class CharacterSearchTests
    {
        private readonly ITestOutputHelper output;
        public CharacterSearchTests(ITestOutputHelper output) { this.output = output; }
        private static ProjectState State(string text, params WordRecord[] words) => new ProjectState {
            Documents = new List<DocumentRecord> { new DocumentRecord { IndexedPages = new List<PageTextRecord> {
                new PageTextRecord { PageNumber = 1, Text = text, Words = words.ToList() } } } } };
        private static WordRecord Word(string text, double x, int line = 1) => new WordRecord {
            Text = text, Line = line, X = x, Y = line * .1, Width = .03, Height = .02 };

        [Theory]
        [InlineData("FAC12345A", "12345")]
        [InlineData("500X300Z35", "500")]
        [InlineData("500X300Z35", "X300")]
        [InlineData("500X300Z35", "X300Z35")]
        [InlineData("500X300Z35", "500X300Z35")]
        [InlineData("500 X300 Z35", "500X300Z35")]
        [InlineData("500X300Z35", "X")]
        [InlineData("FAC00123A", "00123")]
        [InlineData("Référence PréfixeBFA-suffixe", "bfa")]
        [InlineData("BL\u2011001", "BL-001")]
        [InlineData("Réf : o\ufb03ce", "office")]
        [InlineData("FAC12\u200b345A", "12345")]
        public void Search_and_reference_matching_find_the_same_embedded_characters(string source, string query)
        {
            var state = State(source, Word(source, .2));
            var search = Assert.Single(OccurrenceSearch.Find(state, query));
            var matching = Assert.Single(new DocumentMatcher().FindAllFields(state, new[] { query }, true));
            Assert.True(search.HasLocation);
            Assert.True(matching.Fields[0].HasLocation);
            Assert.Equal(search.Evidence, matching.Fields[0].Evidence);
        }

        [Fact]
        public void References_can_span_more_than_eight_pdf_character_boxes()
        {
            const string reference = "500X300Z35";
            var words = reference.Select((c, i) => Word(c.ToString(), .1 + i * .04)).ToArray();
            var state = State(string.Join(" ", reference.ToCharArray()), words);
            var hit = Assert.Single(new DocumentMatcher().FindAllFields(state, new[] { reference }, true)).Fields[0];
            Assert.True(hit.HasLocation); Assert.Equal(.1, hit.X, 5); Assert.Equal(.39, hit.Width, 5);
        }

        [Fact]
        public void Missing_word_boxes_do_not_hide_other_occurrences_in_plain_text()
        {
            var state = State("REF123A REF123B", Word("REF123A", .1));
            var hits = OccurrenceSearch.Find(state, "123");
            Assert.Equal(2, hits.Count); Assert.True(hits[0].HasLocation); Assert.False(hits[1].HasLocation);
            Assert.Equal("REF123B", hits[1].Evidence);
        }

        [Fact]
        public void Words_in_different_order_from_native_text_still_have_locations()
        {
            var state = State("REF123A REF456B", Word("REF456B", .6), Word("REF123A", .1));
            var hit = Assert.Single(OccurrenceSearch.Find(state, "123"));
            Assert.True(hit.HasLocation); Assert.Equal(.1, hit.X);
        }

        [Fact]
        public void Box_only_text_is_searched_without_duplicating_plain_text_hits()
        {
            var state = State("REF123A", Word("REF123A", .1), Word("REF456B", .6));
            Assert.Single(OccurrenceSearch.Find(state, "123"));
            Assert.True(Assert.Single(OccurrenceSearch.Find(state, "456")).HasLocation);
        }

        [Theory]
        [InlineData("12\n34")]
        [InlineData("12\r\n34")]
        public void Separate_rows_are_not_concatenated_into_a_fictitious_reference(string text)
        {
            var state = State(text, Word("12", .1, 1), Word("34", .1, 2));
            Assert.Empty(OccurrenceSearch.Find(state, "1234"));
            Assert.Empty(new DocumentMatcher().FindAllFields(state, new[] { "1234" }, true));
        }

        [Fact]
        public void Overlapping_occurrences_are_all_visible_and_limit_is_respected()
        {
            var state = State("AAAA", Word("AAAA", .1));
            Assert.Equal(3, OccurrenceSearch.Find(state, "AA").Count);
            Assert.Equal(2, OccurrenceSearch.Find(state, "AA", 2).Count);
        }

        [Fact]
        public void Punctuation_is_literal_not_discarded_from_reference_matching()
        {
            Assert.Empty(new DocumentMatcher().FindAllFields(State("AB-123"), new[] { "AB/123" }, true));
            Assert.Empty(OccurrenceSearch.Find(State("AB-123"), "AB/123"));
        }

        [Fact]
        public void Search_finds_a_fragment_of_an_amount_but_automatic_financial_matching_does_not_validate_it()
        {
            var state = State("Total : 16489,83 EUR");
            Assert.Single(OccurrenceSearch.Find(state, "6489,83"));
            Assert.Empty(new DocumentMatcher().FindAllFields(state, new[] { "6489,83" }, true));
        }

        [Fact]
        public void Reference_matching_keeps_leading_zeros_and_does_not_guess_ocr_characters()
        {
            Assert.Empty(new DocumentMatcher().FindAllFields(State("FAC123A X3OO"), new[] { "00123" }, true));
            Assert.Empty(OccurrenceSearch.Find(State("X3OO"), "X300"));
        }

        [Fact]
        public void Matching_keeps_all_fields_on_one_page_and_reports_competing_sources()
        {
            var state = State("FAC12345A Total 100,00");
            state.Documents[0].IndexedPages.Add(new PageTextRecord { PageNumber = 2, Text = "Total 200,00" });
            var matcher = new DocumentMatcher();
            Assert.Single(matcher.FindAllFields(state, new[] { "12345", "100,00" }, true));
            Assert.Empty(matcher.FindAllFields(state, new[] { "12345", "200,00" }, true));
            state.Documents.Add(State("AUTRE12345B Total 100,00").Documents[0]);
            Assert.Equal(2, matcher.FindAllFields(state, new[] { "12345", "100,00" }, true).Count);
        }

        [Fact]
        public void Batch_reuses_page_preparation_and_does_not_reload_or_write_an_index_per_row()
        {
            var loads = 0;
            var words = Enumerable.Range(0, 10000).Select(i => Word("FAC" + i.ToString("D8") + "A", .1, i + 1)).ToList();
            var page = new PageTextRecord { PageNumber = 1, Text = string.Join("\n", words.Select(w => w.Text)), Words = words };
            var state = new ProjectState { Documents = new List<DocumentRecord> { new DocumentRecord {
                PageLoader = () => { loads++; return new List<PageTextRecord> { page }; } } } };
            var rows = Enumerable.Range(0, 1000).Select(i => (IReadOnlyList<string>)new[] { "00009999" }).ToList();
            var watch = Stopwatch.StartNew();
            var hits = new DocumentMatcher().FindBatch(state, rows, true, CancellationToken.None);
            output.WriteLine("1000 rows, 10000 positioned references: " + watch.ElapsedMilliseconds + " ms");
            Assert.Equal(1, loads); Assert.All(hits, h => Assert.Single(h));
            Assert.Throws<OperationCanceledException>(() => new DocumentMatcher().FindBatch(state, rows, true, new CancellationToken(true)));
        }

        [Fact]
        public void Index_save_can_skip_extra_recovery_copy_while_preserving_reload_and_previous_version()
        {
            var root = Path.Combine(Path.GetTempPath(), "doctracker-search-" + Guid.NewGuid().ToString("N"));
            try
            {
                var store = new ProjectStore(root); var state = State("FAC12345A");
                state.Snips.Add(new SnipRecord { DocumentId = state.Documents[0].Id, RawText = "Preuve" });
                store.Save(state);
                var recovery = Path.Combine(root, "recovery");
                var count = Directory.GetFiles(recovery).Length;
                state.Documents[0].IndexedPages = new List<PageTextRecord> { new PageTextRecord { PageNumber = 1, Text = "FAC12345B" } };
                store.Save(state, createRecoveryCheckpoint: false);
                Assert.Equal(count, Directory.GetFiles(recovery).Length);
                Assert.True(File.Exists(store.MetadataPath + ".bak"));
                var restored = new ProjectStore(root).LoadOrCreate("");
                Assert.Single(restored.Snips);
                Assert.Equal("FAC12345B", Assert.Single(OccurrenceSearch.Find(restored, "12345")).Evidence);
                store.Save(state); Assert.Equal(count + 1, Directory.GetFiles(recovery).Length);
            }
            finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        }
    }
}
