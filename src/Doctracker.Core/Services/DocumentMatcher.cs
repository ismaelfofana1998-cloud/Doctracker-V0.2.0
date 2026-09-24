using System;
using System.Threading;
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

        public IReadOnlyList<MatchCandidate> Find(ProjectState state, string query, int maximum = 5, bool partialReferences = false, CancellationToken cancellation = default(CancellationToken))
        {
            if (state == null) throw new ArgumentNullException(nameof(state));
            var result = new List<MatchCandidate>();
            if ((partialReferences ? OccurrenceSearch.Normalize(query) : Normalize(query)).Length == 0 || maximum <= 0) return result;
            foreach(var document in state.Documents)
            {
                try
                {
                    foreach(var page in document.IndexedPages)
                    {
                        cancellation.ThrowIfCancellationRequested();
                        var prepared = partialReferences ? new OccurrenceSearch.PreparedPage(document,page,cancellation) : null;
                        var candidate=Score(document,page,query,partialReferences,prepared,cancellation);
                        if(candidate.Score<=0)continue;
                        result.Add(candidate);
                        if(result.Count>maximum)result=result.OrderByDescending(c=>c.Score).ThenBy(c=>c.DocumentId).ThenBy(c=>c.PageNumber).Take(maximum).ToList();
                    }
                }
                finally {document.ReleaseIndex();}
            }
            return result.OrderByDescending(c=>c.Score).ThenBy(c=>c.DocumentId).ThenBy(c=>c.PageNumber).ToList();
        }
        public IReadOnlyList<MatchCandidate> FindAllFields(ProjectState state, IReadOnlyList<string> queries, bool partialReferences = false)
        {
            return FindBatch(state,new[]{queries},partialReferences,CancellationToken.None)[0];
        }
        // Pages are loaded once per batch, not once for every selected Excel row.
        // Two candidates are sufficient to reject an ambiguous row.
        public IReadOnlyList<MatchCandidate>[] FindBatch(ProjectState state, IReadOnlyList<IReadOnlyList<string>> rows, bool partialReferences, CancellationToken cancellation)
        {
            var result=rows.Select(row=>new List<MatchCandidate>()).ToArray();
            foreach(var document in state.Documents)
            {
                try
                {
                    foreach(var page in document.IndexedPages)
                    {
                        cancellation.ThrowIfCancellationRequested();
                        var prepared = partialReferences ? new OccurrenceSearch.PreparedPage(document,page,cancellation) : null;
                        var cache = new Dictionary<string,MatchCandidate>(StringComparer.Ordinal);
                        for(var row=0;row<rows.Count;row++)
                        {
                            if(row%25==0)cancellation.ThrowIfCancellationRequested();
                            var queries=rows[row];
                            if(result[row].Count>=2 || queries==null || queries.Count==0 || queries.All(string.IsNullOrWhiteSpace))continue;
                            var fields=new List<MatchCandidate>();var match=true;
                            foreach(var query in queries)
                            {
                                MatchCandidate field = null;
                                if (!string.IsNullOrWhiteSpace(query) && !cache.TryGetValue(query,out field))
                                {
                                    field = Score(document,page,query,partialReferences,prepared,cancellation);
                                    cache.Add(query,field);
                                }
                                fields.Add(field);
                                if(field!=null && !field.IsExact && !(partialReferences && field.IsPartial)){match=false;break;}
                            }
                            if(match)result[row].Add(new MatchCandidate {DocumentId=document.Id,PageNumber=page.PageNumber,Score=1,
                                IsExact=fields.Where(f=>f!=null).All(f=>f.IsExact),IsPartial=fields.Any(f=>f?.IsPartial==true),Fields=fields});
                        }
                    }
                }
                finally {document.ReleaseIndex();}
            }
            return result.Select(x=>(IReadOnlyList<MatchCandidate>)x).ToArray();
        }

        private MatchCandidate Score(DocumentRecord document, PageTextRecord page, string query, bool partialReferences, OccurrenceSearch.PreparedPage prepared, CancellationToken cancellation)
        {
            var normalizedQuery = Normalize(query);
            var candidate = new MatchCandidate { DocumentId = document.Id, PageNumber = page.PageNumber };
            decimal amount;
            var isAmount = parser.TryParseAmount(query, out amount) && !Regex.IsMatch(query.Trim(), @"^0\d");
            DateTime date;
            var isDate = TryDate(query, out date);
            // A bare integer may be an invoice/BL identifier. Decimal/currency/signed
            // values remain financial comparisons; free occurrence search is always literal.
            var explicitAmount = isAmount && Regex.IsMatch(query, @"[.,+−()€$£A-Za-z]|^-|-$");
            if (partialReferences && !isDate && !explicitAmount)
            {
                var hit = prepared.Find(OccurrenceSearch.SearchTerms(query),cancellation).FirstOrDefault();
                return hit ?? candidate;
            }

            Func<string, bool> exact = text =>
            {
                if (isDate) { DateTime other; return TryDate(text, out other) && other == date; }
                if (isAmount) { decimal other; return parser.TryParseAmount(text, out other) && other == amount; }
                return Normalize(text) == normalizedQuery;
            };

            // OCR/PDF engines can return a whole label and value as a single word box.
            // Match an exact token inside that box, retaining its reliable location.
            Func<string, string> contained = text => {
                if (exact(text)) return text;
                if (isAmount || isDate)
                {
                    var pattern = isDate ? @"(?<!\d)\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}(?!\d)" :
                        @"(?<![\p{L}\d/.,-])\(?[-+−]?(?:\d{1,3}(?:[ \u00A0\u202F]\d{3}(?!\d))+|\d+)(?:[.,]\d+)*-?\)?(?![\p{L}\d/.,-])";
                    foreach (Match match in Regex.Matches(text ?? "", pattern))
                        if (exact(match.Value)) return match.Value;
                }
                else if ((" " + Normalize(text) + " ").Contains(" " + normalizedQuery + " "))
                    return FindSourceText(text, normalizedQuery);
                return null;
            };
            var words = page.Words ?? new List<WordRecord>();
            var maxWords = isAmount ? 6 : normalizedQuery.Split(' ').Length + 2;
            for(var count=1;count<=maxWords;count++)
            for(var start=0;start+count<=words.Count;start++)
            {
                var selected=words.Skip(start).Take(count).ToList();
                if(selected.Any(word=>word.Line!=selected[0].Line))continue;
                var value=string.Join(" ",selected.Select(word=>word.Text));
                var evidence=contained(value);
                var exactMatch=evidence!=null;
                if(!exactMatch)continue;
                candidate.IsExact=true;candidate.Score=1;
                candidate.Evidence=evidence??value;candidate.HasLocation=true;
                candidate.X=selected.Min(word=>word.X);candidate.Y=selected.Min(word=>word.Y);
                candidate.Width=selected.Max(word=>word.X+word.Width)-candidate.X;
                candidate.Height=selected.Max(word=>word.Y+word.Height)-candidate.Y;
                return candidate;
            }

            // Legacy projects/native PDF pages may have plain text without word boxes.
            var textValue = page.Text ?? "";
            if (isAmount || isDate)
            {
                var evidence = contained(textValue);
                if (evidence != null) return Exact(candidate, evidence);
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
