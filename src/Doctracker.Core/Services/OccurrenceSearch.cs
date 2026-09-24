using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    public static class OccurrenceSearch
    {
        public static IReadOnlyList<MatchCandidate> Find(ProjectState state, string query, int maximum = 201, CancellationToken cancellation = default)
        {
            if (state == null) throw new ArgumentNullException(nameof(state));
            var result = new List<MatchCandidate>();
            var terms = SearchTerms(query);
            if (terms[0].Length == 0 || maximum <= 0) return result;
            cancellation.ThrowIfCancellationRequested();
            foreach (var document in state.Documents)
            {
                try
                {
                    cancellation.ThrowIfCancellationRequested();
                    foreach (var page in document.IndexedPages)
                    {
                        var prepared = new PreparedPage(document, page, cancellation);
                        foreach (var hit in prepared.Find(terms, cancellation))
                        {
                            result.Add(hit);
                            if (result.Count >= maximum) return result;
                        }
                    }
                }
                finally { document.ReleaseIndex(); }
            }
            return result;
        }

        internal static string[] SearchTerms(string query)
        {
            var formats = new[] { "yyyy-MM-dd", "d/M/yyyy", "dd/MM/yyyy", "d-M-yyyy", "dd-MM-yyyy", "d.M.yyyy", "dd.MM.yyyy" };
            if (DateTime.TryParseExact((query ?? "").Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                return formats.Select(f => Normalize(date.ToString(f, CultureInfo.InvariantCulture))).Distinct().ToArray();
            return new[] { Normalize(query) };
        }

        internal static string Normalize(string value) => new TextMap(value).Text;

        // Keep offsets into the original text: search ignores layout spaces, while evidence
        // retains its spelling. A line break remains a boundary (two table rows aren't a reference).
        private sealed class TextMap
        {
            public readonly string Source, Text;
            public readonly int[] SourceOffsets;
            public TextMap(string value)
            {
                Source = SafeXmlWriter.Clean(value) ?? "";
                var text = new StringBuilder(); var offsets = new List<int>();
                for (var i = 0; i < Source.Length; i++)
                {
                    var c = Source[i];
                    if (c == '\r' || c == '\n')
                    {
                        if (text.Length == 0 || text[text.Length - 1] != '\n') { text.Append('\n'); offsets.Add(i); }
                        continue;
                    }
                    if (char.IsWhiteSpace(c) || c == '\u00ad' || c == '\u200b' || c == '\ufeff') continue;
                    // Normalize a surrogate pair together to avoid invalid UTF-16 in PDF text.
                    var length = char.IsHighSurrogate(c) && i + 1 < Source.Length && char.IsLowSurrogate(Source[i + 1]) ? 2 : 1;
                    foreach (var normalized in Source.Substring(i, length).Normalize(NormalizationForm.FormKD).ToUpperInvariant())
                    {
                        if (CharUnicodeInfo.GetUnicodeCategory(normalized) == UnicodeCategory.NonSpacingMark) continue;
                        text.Append(normalized == ',' ? '.' : "\u2010\u2011\u2012\u2013\u2212".IndexOf(normalized) >= 0 ? '-' : normalized);
                        offsets.Add(i);
                    }
                    i += length - 1;
                }
                Text = text.ToString(); SourceOffsets = offsets.ToArray();
            }
        }

        // One preparation per page, shared by all the rows of an Excel matching run.
        internal sealed class PreparedPage
        {
            private readonly DocumentRecord document;
            private readonly PageTextRecord page;
            private readonly List<WordRecord> words;
            private readonly TextMap raw, extra;
            private readonly int[] rawOwners, extraOwners;

            internal PreparedPage(DocumentRecord document, PageTextRecord page, CancellationToken cancellation)
            {
                this.document = document; this.page = page;
                words = page.Words ?? new List<WordRecord>();
                raw = new TextMap(page.Text);
                rawOwners = Enumerable.Repeat(-1, raw.Text.Length).ToArray();
                var additional = new StringBuilder(); var sourceOwners = new List<int>();
                var cursor = 0; var previousExtra = -1;
                for (var w = 0; w < words.Count; w++)
                {
                    if (w % 128 == 0) cancellation.ThrowIfCancellationRequested();
                    var token = Normalize(words[w].Text);
                    if (token.Length == 0) continue;
                    var at = raw.Text.IndexOf(token, cursor, StringComparison.Ordinal);
                    if (at < 0) at = raw.Text.IndexOf(token, StringComparison.Ordinal);
                    if (at >= 0)
                    {
                        // PDF engines sometimes return boxes in a different order from page text.
                        if (rawOwners[at] < 0)
                            for (var i = at; i < at + token.Length; i++) rawOwners[i] = w;
                        cursor = at + token.Length;
                        continue;
                    }
                    // Retain text from boxes even when the page's plain-text field is incomplete.
                    if (additional.Length > 0)
                    {
                        additional.Append(previousExtra == w - 1 && SameLine(words[previousExtra], words[w]) ? ' ' : '\n');
                        sourceOwners.Add(-1);
                    }
                    var source = SafeXmlWriter.Clean(words[w].Text) ?? "";
                    additional.Append(source);
                    for (var i = 0; i < source.Length; i++) sourceOwners.Add(w);
                    previousExtra = w;
                }
                extra = new TextMap(additional.ToString());
                extraOwners = extra.SourceOffsets.Select(i => sourceOwners[i]).ToArray();
                cancellation.ThrowIfCancellationRequested();
            }

            private static bool SameLine(WordRecord left, WordRecord right) => left.Line == right.Line &&
                (left.Height <= 0 || right.Height <= 0 || Math.Abs(left.Y - right.Y) <= Math.Max(left.Height, right.Height) * .6);

            internal IEnumerable<MatchCandidate> Find(string[] terms, CancellationToken cancellation)
            {
                foreach (var hit in Find(raw, rawOwners, terms, cancellation)) yield return hit;
                foreach (var hit in Find(extra, extraOwners, terms, cancellation)) yield return hit;
            }

            private IEnumerable<MatchCandidate> Find(TextMap map, int[] owners, string[] terms, CancellationToken cancellation)
            {
                // Merge date spellings by offset, without enumerating a whole large page before the first result.
                var positions = new int[terms.Length];
                for (var t = 0; t < terms.Length; t++) positions[t] = terms[t].Length == 0 ? -1 : map.Text.IndexOf(terms[t], StringComparison.Ordinal);
                while (positions.Any(p => p >= 0))
                {
                    cancellation.ThrowIfCancellationRequested();
                    var at = positions.Where(p => p >= 0).Min();
                    var length = Enumerable.Range(0, terms.Length).Where(t => positions[t] == at).Max(t => terms[t].Length);
                    yield return Hit(map, owners, at, length);
                    for (var t = 0; t < terms.Length; t++)
                        if (positions[t] == at) positions[t] = map.Text.IndexOf(terms[t], at + 1, StringComparison.Ordinal);
                }
            }

            private MatchCandidate Hit(TextMap map, int[] owners, int at, int length)
            {
                var first = map.SourceOffsets[at]; var last = map.SourceOffsets[at + length - 1] + 1;
                // Preserve the complete surrounding reference in matching output, not just the searched fragment.
                while (first > 0 && !char.IsWhiteSpace(map.Source[first - 1])) first--;
                while (last < map.Source.Length && !char.IsWhiteSpace(map.Source[last])) last++;
                var evidence = map.Source.Substring(first, last - first);
                var partial = Normalize(evidence) != map.Text.Substring(at, length);
                var hit = new MatchCandidate { DocumentId = document.Id, PageNumber = page.PageNumber,
                    Evidence = evidence, Score = partial ? .9 : 1, IsPartial = partial, IsExact = !partial };
                var selected = new HashSet<int>();
                for (var i = at; i < at + length; i++)
                {
                    if (owners[i] < 0) return hit; // Text remains searchable when some glyphs have no usable box.
                    selected.Add(owners[i]);
                }
                if (selected.Count == 0) return hit;
                var boxes = selected.Select(w => words[w]).ToList();
                if (boxes.Any(w => w.Width <= 0 || w.Height <= 0)) return hit;
                hit.X = boxes.Min(w => w.X); hit.Y = boxes.Min(w => w.Y);
                hit.Width = boxes.Max(w => w.X + w.Width) - hit.X;
                hit.Height = boxes.Max(w => w.Y + w.Height) - hit.Y;
                hit.HasLocation = true;
                return hit;
            }
        }
    }
}
