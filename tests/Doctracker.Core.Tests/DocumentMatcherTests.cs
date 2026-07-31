using System.Collections.Generic;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class DocumentMatcherTests
    {
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
