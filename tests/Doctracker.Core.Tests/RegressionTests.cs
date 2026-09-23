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
    public sealed class RegressionTests
    {
        [Theory]
        [InlineData("(1 250,50)", "-1250.5")]
        [InlineData("1\u202f250,50", "1250.5")]
        [InlineData("1\u00a0250,50", "1250.5")]
        [InlineData("1250,50-", "-1250.5")]
        [InlineData("−1250,50", "-1250.5")]
        [InlineData("1.234.567,89", "1234567.89")]
        [InlineData("1234.5678", "1234.5678")]
        public void Accounting_numbers_preserve_sign_and_precision(string text, string expected) =>
            Assert.Equal(expected, new TextValueParser().Parse(SnipType.Number, text));

        [Theory]
        [InlineData("10 20")]
        [InlineData("10\t20")]
        [InlineData("10\n20")]
        [InlineData("1,2,3")]
        public void Number_does_not_silently_choose_or_concatenate_multiple_amounts(string text) =>
            Assert.Throws<FormatException>(() => new TextValueParser().Parse(SnipType.Number, text));

        [Fact]
        public void Sum_does_not_join_two_columns() => Assert.Equal("30", new TextValueParser().Parse(SnipType.Sum, "10  20"));

        [Theory]
        [InlineData(SnipType.Text)]
        [InlineData(SnipType.Table)]
        public void Empty_ocr_is_not_a_success(SnipType type) => Assert.Throws<FormatException>(() => new TextValueParser().Parse(type, " "));

        [Theory]
        [InlineData("31/02/2026")]
        [InlineData("01/01/2026 02/01/2026")]
        public void Invalid_or_multiple_dates_are_rejected(string text) => Assert.Throws<FormatException>(() => TextValueParser.ParseDate(text));

        [Theory]
        [InlineData("123", "1234")]
        [InlineData("100,00", "-100,00")]
        [InlineData("FA-2025-0198", "FA-2025-0199")]
        [InlineData("123", "REF-123-001")]
        public void Similar_reference_or_amount_is_not_an_exact_match(string query, string source)
        {
            var state = Project(source);
            Assert.Empty(new DocumentMatcher().FindAllFields(state, new[] { query }));
        }

        [Fact]
        public void All_fields_must_exist_on_same_page()
        {
            var state = Project("FA-001 total 100,00 EUR");
            state.Documents[0].IndexedPages.Add(new PageTextRecord { PageNumber = 2, Text = "FA-002 total 200,00 EUR" });
            var matcher = new DocumentMatcher();
            Assert.Single(matcher.FindAllFields(state, new[] { "FA-002", "200.00" }));
            Assert.Empty(matcher.FindAllFields(state, new[] { "FA-001", "200.00" }));
            Assert.Empty(matcher.FindAllFields(state, new[] { "", "" }));
        }

        [Fact]
        public void Multiple_documents_remain_visible_as_ambiguous()
        {
            var state = Project("Référence FA-001");
            state.Documents.Add(Project("Référence FA-001").Documents[0]);
            Assert.Equal(2, new DocumentMatcher().FindAllFields(state, new[] { "FA-001" }).Count);
        }

        [Fact]
        public void Positional_match_returns_source_text_and_zone()
        {
            var state = Project("Facture 1 250,00");
            state.Documents[0].IndexedPages[0].Words.AddRange(new[]
            {
                Word("Facture", .1, .1, 1), Word("1", .5, .2, 2), Word("250,00", .6, .2, 2)
            });
            var match = Assert.Single(new DocumentMatcher().FindAllFields(state, new[] { "1250.00" }));
            Assert.True(match.Fields[0].HasLocation);
            Assert.Equal("1 250,00", match.Fields[0].Evidence);
            Assert.Equal(.5, match.Fields[0].X, 5);
        }

        [Fact]
        public void Date_matches_other_supported_format()
        {
            Assert.Single(new DocumentMatcher().FindAllFields(Project("Date : 31/12/2025"), new[] { "2025-12-31" }));
        }

        [Fact]
        public void Layout_preserves_empty_cells_and_right_aligned_numbers()
        {
            var words = new[] { Word("Article", .1, .1, 1), Word("Total", .7, .1, 1),
                Word("Service", .1, .2, 2), Word("100", .72, .2, 2), Word("Autre", .1, .3, 3) };
            var cells = LayoutExtractor.Extract(words);
            Assert.Equal(5, cells.Count);
            Assert.Equal(1, cells.Single(cell => cell.Text == "100").Column);
            Assert.Equal(0, cells.Single(cell => cell.Text == "Autre").Column);
        }

        [Fact]
        public void Legacy_xml_remains_readable_and_new_word_boxes_roundtrip()
        {
            InDirectory(directory =>
            {
                File.WriteAllText(Path.Combine(directory, "project.xml"),
                    "<DoctrackerProject SchemaVersion=\"1\"><Documents><Document Id=\"doc\"><PageCount>1</PageCount><IndexedPages><Page PageNumber=\"1\">Total 1 250,00</Page></IndexedPages></Document></Documents></DoctrackerProject>");
                var store = new ProjectStore(directory);
                var state = store.LoadOrCreate("mission.xlsx");
                Assert.Equal("Total 1 250,00", state.Documents[0].IndexedPages[0].Text);
                Assert.True(state.Documents[0].IndexComplete);
                state.Documents[0].IndexedPages[0].Words.Add(Word("Total", .1, .1, 1));
                store.Save(state);
                var loaded = store.LoadOrCreate("mission.xlsx");
                Assert.Equal("Total 1 250,00", loaded.Documents[0].IndexedPages[0].Text);
                Assert.Single(loaded.Documents[0].IndexedPages[0].Words);
            });
        }

        [Fact]
        public void Preparing_a_snip_does_not_publish_a_false_proof()
        {
            InDirectory(directory =>
            {
                var store = new ProjectStore(directory);
                var state = Project("100,00");
                var service = new SnipService(store, new TextValueParser());
                var snip = service.Prepare(state, state.Documents[0].Id, 1, new NormalizedRectangle(0, 0, 1, 1), SnipType.Number, "100,00", "Achats", "A1", "tester");
                Assert.Empty(state.Snips);
                Assert.False(File.Exists(store.MetadataPath));
                service.Commit(state, new[] { snip });
                Assert.Single(store.LoadOrCreate("mission.xlsx").Snips);
                service.SetReview(state, snip.Id, ReviewStatus.Reviewed, "OK", "reviewer");
                service.SetReview(state, snip.Id, ReviewStatus.Prepared, "À reprendre", "reviewer");
                Assert.Null(snip.ReviewedAtUtc);
                Assert.Empty(snip.ReviewedBy);
            });
        }

        [Fact]
        public void Failed_save_rolls_back_snips_and_audit_events()
        {
            InDirectory(directory =>
            {
                var store = new ProjectStore(directory);
                Directory.CreateDirectory(store.MetadataPath + ".tmp");
                var state = Project("100");
                var service = new SnipService(store, new TextValueParser());
                var snip = service.Prepare(state, state.Documents[0].Id, 1, new NormalizedRectangle(0, 0, 1, 1), SnipType.Text, "test", "A", "A1", "tester");
                Assert.ThrowsAny<Exception>(() => service.Commit(state, new[] { snip }));
                Assert.Empty(state.Snips);
                Assert.Empty(state.AuditTrail);
            });
        }

        [Fact]
        public void Non_finite_geometry_is_rejected() => Assert.Throws<ArgumentOutOfRangeException>(() => new NormalizedRectangle(double.NaN, 0, 1, 1));

        private static ProjectState Project(string text) => new ProjectState
        {
            Documents = new List<DocumentRecord> { new DocumentRecord { PageCount = 1,
                IndexedPages = new List<PageTextRecord> { new PageTextRecord { PageNumber = 1, Text = text } } } }
        };
        private static WordRecord Word(string text, double x, double y, int line) =>
            new WordRecord { Text = text, X = x, Y = y, Width = .08, Height = .03, Line = line };
        private static void InDirectory(Action<string> action)
        {
            var directory = Path.Combine(Path.GetTempPath(), "doctracker-tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try { action(directory); } finally { Directory.Delete(directory, true); }
        }
    }
}
