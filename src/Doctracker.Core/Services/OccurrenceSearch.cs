using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    // Interactive text search is deliberately independent of automatic audit matching.
    // Every literal occurrence is useful here, including a number inside a reference.
    public static class OccurrenceSearch
    {
        public static IReadOnlyList<MatchCandidate> Find(ProjectState state,string query,int maximum=201,CancellationToken cancellation=default)
        {
            var result=new List<MatchCandidate>();var needle=Normalize(query);
            if(needle.Length==0 || maximum<=0)return result;
            foreach(var doc in state.Documents)
            {
                try
                {
                    foreach(var page in doc.IndexedPages)
                    {
                        cancellation.ThrowIfCancellationRequested();var before=result.Count;
                        var text=new StringBuilder();var owners=new List<int>();var words=page.Words??new List<WordRecord>();
                        for(var w=0;w<words.Count;w++)
                        {
                            var token=Normalize(words[w].Text);text.Append(token);
                            for(var i=0;i<token.Length;i++)owners.Add(w);
                        }
                        var haystack=text.ToString();
                        foreach(var offset in Offsets(haystack,needle))
                        {
                            cancellation.ThrowIfCancellationRequested();
                            var selected=words.Skip(owners[offset]).Take(owners[offset+needle.Length-1]-owners[offset]+1).ToList();
                            var hit=Hit(doc,page,string.Join(" ",selected.Select(w=>w.Text)));
                            hit.X=selected.Min(w=>w.X);hit.Y=selected.Min(w=>w.Y);
                            hit.Width=selected.Max(w=>w.X+w.Width)-hit.X;hit.Height=selected.Max(w=>w.Y+w.Height)-hit.Y;
                            hit.HasLocation=hit.Width>0 && hit.Height>0;
                            result.Add(hit);if(result.Count>=maximum)return result;
                        }
                        // Some older PDFs have text but no usable word boxes.
                        if(result.Count==before)
                            foreach(var offset in Offsets(Normalize(page.Text),needle))
                            {result.Add(Hit(doc,page,page.Text));if(result.Count>=maximum)return result;}
                    }
                }
                finally{doc.ReleaseIndex();}
            }
            return result;
        }
        private static MatchCandidate Hit(DocumentRecord doc,PageTextRecord page,string evidence)=>new MatchCandidate {
            DocumentId=doc.Id,PageNumber=page.PageNumber,Score=1,IsExact=true,
            Evidence=(evidence??"").Length>180 ? evidence.Substring(0,180)+"…" : evidence??"" };
        private static IEnumerable<int> Offsets(string text,string query)
        {
            var start=0;
            while(start<=text.Length-query.Length)
            {
                var at=text.IndexOf(query,start,StringComparison.Ordinal);if(at<0)yield break;
                yield return at;start=at+Math.Max(1,query.Length);
            }
        }
        private static string Normalize(string text)
        {
            var output=new StringBuilder();
            foreach(var c in (SafeXmlWriter.Clean(text)??"").Normalize(NormalizationForm.FormD).ToUpperInvariant())
            {
                if(char.IsWhiteSpace(c)||CharUnicodeInfo.GetUnicodeCategory(c)==UnicodeCategory.NonSpacingMark)continue;
                output.Append(c==',' ? '.' : c);
            }
            return output.ToString();
        }
    }
}
