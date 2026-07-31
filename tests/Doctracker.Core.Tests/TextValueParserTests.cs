using System;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class TextValueParserTests
    {
        private readonly TextValueParser parser = new TextValueParser();

        [Theory]
        [InlineData("Montant TTC 12 345,67 €", "12345.67")]
        [InlineData("-1.250,50", "-1250.5")]
        [InlineData("Solde : 2,500.25 USD", "2500.25")]
        public void NumberSnip_understands_common_accounting_formats(string input, string expected)
        {
            Assert.Equal(expected, parser.Parse(SnipType.Number, input));
        }

        [Fact]
        public void SumSnip_adds_every_recognized_amount()
        {
            Assert.Equal("16.5", parser.Parse(SnipType.Sum, "10,00\n5,25\n1,25"));
        }

        [Theory]
        [InlineData("Facture du 31/12/2025", "2025-12-31")]
        [InlineData("2026-01-04", "2026-01-04")]
        public void DateSnip_normalizes_dates(string input, string expected)
        {
            Assert.Equal(expected, parser.Parse(SnipType.Date, input));
        }

        [Fact]
        public void NumberSnip_rejects_a_zone_without_number()
        {
            Assert.Throws<FormatException>(() => parser.Parse(SnipType.Number, "aucun montant"));
        }
    }
}
