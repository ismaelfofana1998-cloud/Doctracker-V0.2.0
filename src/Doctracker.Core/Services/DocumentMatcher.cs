using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    public sealed class DocumentMatcher
    {
        public IReadOnlyList<MatchCandidate> Find(ProjectState state, string query, int maximum = 5)
        {
            if (state == null) throw new ArgumentNullException(nameof(state));
            if (string.IsNullOrWhiteSpace(query)) return new List<MatchCandidate>();

            var normalizedQuery = Normalize(query);
            var queryTokens = Tokenize(normalizedQuery);
            decimal queryAmount;
            var hasAmount = TryExtractAmount(query, out queryAmount);
            if (normalizedQuery.Length == 0 && !hasAmount)
            {
                return new List<MatchCandidate>();
            }

            return state.Documents
                .SelectMany(document => document.IndexedPages.Select(page =>
                    Score(document.Id, page, normalizedQuery, queryTokens, hasAmount, queryAmount)))
                .Where(candidate => candidate.Score > 0)
                .OrderByDescending(candidate => candidate.Score)
                .Take(Math.Max(1, maximum))
                .ToList();
        }

        private static MatchCandidate Score(
            string documentId,
            PageTextRecord page,
            string normalizedQuery,
            HashSet<string> queryTokens,
            bool hasAmount,
            decimal queryAmount)
        {
            var normalizedPage = Normalize(page.Text);
            var score = 0d;

            if (normalizedQuery.Length > 0 && normalizedPage.Contains(normalizedQuery))
            {
                score += 0.65;
            }

            var pageTokens = Tokenize(normalizedPage);
            if (queryTokens.Count > 0)
            {
                var intersection = queryTokens.Count(token => pageTokens.Contains(token));
                score += 0.55 * intersection / queryTokens.Count;
            }

            if (hasAmount)
            {
                var amounts = ExtractAmounts(page.Text);
                if (amounts.Any(amount => Math.Abs(amount - queryAmount) <= 0.01m))
                {
                    score += 0.35;
                }
            }

            return new MatchCandidate
            {
                DocumentId = documentId,
                PageNumber = page.PageNumber,
                Score = Math.Min(1d, score),
                Evidence = CreateEvidence(page.Text, normalizedQuery)
            };
        }

        private static string Normalize(string value)
        {
            var decomposed = (value ?? string.Empty).ToLowerInvariant().Normalize(NormalizationForm.FormD);
            var builder = new StringBuilder();
            foreach (var character in decomposed)
            {
                var category = CharUnicodeInfo.GetUnicodeCategory(character);
                if (category != UnicodeCategory.NonSpacingMark)
                {
                    builder.Append(char.IsLetterOrDigit(character) ? character : ' ');
                }
            }
            return Regex.Replace(builder.ToString(), @"\s+", " ").Trim();
        }

        private static HashSet<string> Tokenize(string value)
        {
            return new HashSet<string>(
                Regex.Split(value ?? string.Empty, @"\s+").Where(token => token.Length >= 2),
                StringComparer.Ordinal);
        }

        private static bool TryExtractAmount(string value, out decimal amount)
        {
            amount = 0m;
            var parser = new TextValueParser();
            try
            {
                amount = parser.ParseNumber(value);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        private static IReadOnlyList<decimal> ExtractAmounts(string value)
        {
            var parser = new TextValueParser();
            try
            {
                return parser.ParseAllNumbers(value);
            }
            catch (FormatException)
            {
                return new List<decimal>();
            }
        }

        private static string CreateEvidence(string pageText, string normalizedQuery)
        {
            var singleLine = Regex.Replace(pageText ?? string.Empty, @"\s+", " ").Trim();
            if (singleLine.Length <= 180) return singleLine;
            var firstToken = normalizedQuery.Split(' ').FirstOrDefault() ?? string.Empty;
            var index = Normalize(singleLine).IndexOf(firstToken, StringComparison.Ordinal);
            var start = Math.Max(0, Math.Min(singleLine.Length - 180, index - 60));
            return singleLine.Substring(start, 180);
        }
    }
}
