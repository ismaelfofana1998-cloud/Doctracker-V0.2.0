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
        private readonly TextValueParser parser = new TextValueParser();

        public IReadOnlyList<MatchCandidate> Find(ProjectState state, string query, int maximum = 5)
        {
            if (state == null) throw new ArgumentNullException(nameof(state));
            if (Normalize(query).Length == 0 || maximum <= 0) return new List<MatchCandidate>();
            return state.Documents.SelectMany(document => document.IndexedPages.Select(page => Score(document, page, query)))
                .Where(candidate => candidate.Score > 0).OrderByDescending(candidate => candidate.Score)
                .ThenBy(candidate => candidate.DocumentId).ThenBy(candidate => candidate.PageNumber).Take(maximum).ToList();
        }

        // Automatic matching is deliberately stricter than interactive search:
        // every populated field must occur on the same page, and ambiguous documents stay unresolved.
        public IReadOnlyList<MatchCandidate> FindAllFields(ProjectState state, IReadOnlyList<string> queries)
        {
            if (state == null) throw new ArgumentNullException(nameof(state));
            if (queries == null || queries.Count == 0 || queries.All(string.IsNullOrWhiteSpace))
                return new List<MatchCandidate>();
            var result = new List<MatchCandidate>();
            foreach (var document in state.Documents)
            foreach (var page in document.IndexedPages)
            {
                var fields = queries.Select(query => string.IsNullOrWhiteSpace(query) ? null : Score(document, page, query)).ToList();
                if (fields.Where(field => field != null).Any(field => !field.IsExact)) continue;
                result.Add(new MatchCandidate { DocumentId = document.Id, PageNumber = page.PageNumber,
                    Score = 1, IsExact = true, Fields = fields });
            }
            return result;
        }

        private MatchCandidate Score(DocumentRecord document, PageTextRecord page, string query)
        {
            var normalizedQuery = Normalize(query);
            var candidate = new MatchCandidate { DocumentId = document.Id, PageNumber = page.PageNumber };
            decimal amount;
            var isAmount = parser.TryParseAmount(query, out amount) && !Regex.IsMatch(query.Trim(), @"^0\d");
            DateTime date;
            var isDate = TryDate(query, out date);
            Func<string, bool> exact = text =>
            {
                if (isDate) { DateTime other; return TryDate(text, out other) && other == date; }
                if (isAmount) { decimal other; return parser.TryParseAmount(text, out other) && other == amount; }
                return Normalize(text) == normalizedQuery;
            };

            // Prefer a real word location. Never join words from different lines for amounts.
            var words = page.Words ?? new List<WordRecord>();
            var maxWords = isAmount ? 6 : normalizedQuery.Split(' ').Length + 2;
            for (var start = 0; start < words.Count; start++)
            {
                var selected = new List<WordRecord>();
                for (var end = start; end < words.Count && end < start + maxWords; end++)
                {
                    if ((isAmount || isDate) && words[end].Line != words[start].Line) break;
                    selected.Add(words[end]);
                    var value = string.Join(" ", selected.Select(word => word.Text));
                    if (!exact(value)) continue;
                    candidate.IsExact = true;
                    candidate.Score = 1;
                    candidate.Evidence = value;
                    candidate.HasLocation = true;
                    candidate.X = selected.Min(word => word.X);
                    candidate.Y = selected.Min(word => word.Y);
                    candidate.Width = selected.Max(word => word.X + word.Width) - candidate.X;
                    candidate.Height = selected.Max(word => word.Y + word.Height) - candidate.Y;
                    return candidate;
                }
            }

            // Legacy projects/native PDF pages may have plain text without word boxes.
            var textValue = page.Text ?? "";
            if (isAmount)
            {
                foreach (Match match in Regex.Matches(textValue,
                    @"(?<![\p{L}\d/.-])\(?[-+−]?(?:\d{1,3}(?:[ \u00A0\u202F]\d{3}(?!\d))+|\d+)(?:[.,]\d+)*-?\)?(?![\p{L}\d/.-])"))
                    if (exact(match.Value)) return Exact(candidate, match.Value);
            }
            else if (isDate)
            {
                foreach (Match match in Regex.Matches(textValue, @"(?<!\d)\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}(?!\d)"))
                    if (exact(match.Value)) return Exact(candidate, match.Value);
            }
            else
            {
                if ((" " + Normalize(textValue) + " ").Contains(" " + normalizedQuery + " "))
                    return Exact(candidate, FindSourceText(textValue, normalizedQuery));
                if ((" " + Normalize(System.IO.Path.GetFileNameWithoutExtension(document.OriginalName)) + " ")
                    .Contains(" " + normalizedQuery + " ")) return Exact(candidate, document.OriginalName);
                var tokens = normalizedQuery.Split(' ').Distinct().ToArray();
                var available = new HashSet<string>(Normalize(textValue).Split(' '));
                candidate.Score = tokens.Count(token => available.Contains(token)) / (double)tokens.Length * 0.6;
            }
            candidate.Evidence = textValue.Length <= 180 ? textValue : textValue.Substring(0, 180);
            return candidate;
        }

        private static MatchCandidate Exact(MatchCandidate candidate, string evidence)
        {
            candidate.Score = 1; candidate.IsExact = true; candidate.Evidence = evidence;
            return candidate;
        }

        private static string FindSourceText(string text, string query)
        {
            var words = Regex.Split(text.Trim(), @"\s+");
            for (var i = 0; i < words.Length; i++)
            for (var count = 1; count <= 32 && i + count <= words.Length; count++)
            {
                var value = string.Join(" ", words.Skip(i).Take(count));
                if (Normalize(value) == query) return value;
            }
            return text.Trim();
        }

        private static bool TryDate(string text, out DateTime date)
        {
            date = default(DateTime);
            if (!Regex.IsMatch((text ?? "").Trim(), @"^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}$")) return false;
            try { date = TextValueParser.ParseDate(text); return true; }
            catch (FormatException) { return false; }
        }

        private static string Normalize(string text)
        {
            var builder = new StringBuilder();
            foreach (var c in (text ?? "").ToLowerInvariant().Normalize(NormalizationForm.FormD))
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                    builder.Append(char.IsLetterOrDigit(c) ? c : ' ');
            return Regex.Replace(builder.ToString(), @"\s+", " ").Trim();
        }
    }
}
