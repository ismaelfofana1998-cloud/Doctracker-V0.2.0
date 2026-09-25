using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Doctracker.AddIn.Excel
{
    // Hidden workbook names track row/column insertion and survive Save As without cell notes.
    // Rebuild only after a worksheet edit; ordinary selection does not scan all names.
    internal static class ExcelProofLinks
    {
        private const string Prefix = "_DoctrackerProof_";
        private sealed class Link { public string Name, Id; public ExcelInterop.Range Cell; }
        private sealed class Index { public bool Valid; public readonly Dictionary<string,List<Link>> Cells = new Dictionary<string,List<Link>>(); }
        private static readonly ConditionalWeakTable<ExcelInterop.Workbook,Index> indexes = new ConditionalWeakTable<ExcelInterop.Workbook,Index>();
        private static ExcelInterop.Workbook Book(ExcelInterop.Range cell) => (ExcelInterop.Workbook)cell.Worksheet.Parent;
        private static string Key(ExcelInterop.Range cell) => cell.Worksheet.CodeName + "!" + cell.Address[true,true,ExcelInterop.XlReferenceStyle.xlA1];
        public static void Invalidate(ExcelInterop.Workbook book) { if (book != null && indexes.TryGetValue(book,out var index)) index.Valid=false; }
        private static Index Read(ExcelInterop.Workbook book)
        {
            var index=indexes.GetValue(book,_=>new Index()); if(index.Valid)return index;
            index.Cells.Clear();
            foreach(ExcelInterop.Name name in book.Names)
            {
                var local=name.Name; var bang=local.LastIndexOf('!'); if(bang>=0)local=local.Substring(bang+1);
                if(!local.StartsWith(Prefix,StringComparison.Ordinal))continue;
                var id=local.Substring(Prefix.Length).Split('_')[0];
                try
                {
                    var cell=name.RefersToRange;
                    if(cell==null || cell.Cells.CountLarge!=1 || !Equals(cell.Worksheet.Parent,book))continue;
                    var key=Key(cell); if(!index.Cells.TryGetValue(key,out var links))index.Cells[key]=links=new List<Link>();
                    links.Add(new Link {Name=name.Name,Id=id,Cell=cell});
                }
                catch(COMException) { /* Deleted cells leave #REF names, never reattach to another cell. */ }
            }
            index.Valid=true; return index;
        }
        public static IReadOnlyList<string> Get(ExcelInterop.Range cell)
        {
            if(cell==null || cell.Cells.CountLarge!=1)return new string[0];
            return Read(Book(cell)).Cells.TryGetValue(Key(cell),out var links)?links.Select(l=>l.Id).Distinct().ToList():(IReadOnlyList<string>)new string[0];
        }
        public static List<ExcelInterop.Range> Cells(ExcelInterop.Workbook book) => Read(book).Cells.Values.Select(v=>v[0].Cell).ToList();
        public static void Set(ExcelInterop.Range cell,IEnumerable<string> ids)
        {
            var book=Book(cell); var index=Read(book); var key=Key(cell); var wanted=ids.Distinct().ToList();
            if(!index.Cells.TryGetValue(key,out var links))index.Cells[key]=links=new List<Link>();
            try
            {
                foreach(var link in links.Where(l=>!wanted.Contains(l.Id)).ToList()) { book.Names.Item(link.Name).Delete(); links.Remove(link); }
                foreach(var id in wanted.Where(id=>!links.Any(l=>l.Id==id)))
                {
                    var name=Prefix+id+"_"+Guid.NewGuid().ToString("N");
                    var reference="='"+cell.Worksheet.Name.Replace("'","''")+"'!"+cell.Address[true,true,ExcelInterop.XlReferenceStyle.xlA1];
                    book.Names.Add(Name:name,RefersTo:reference,Visible:false);
                    links.Add(new Link {Name=name,Id=id,Cell=cell});
                }
                if(links.Count==0)index.Cells.Remove(key);
            }
            catch {index.Valid=false;throw;}
        }
    }
}
