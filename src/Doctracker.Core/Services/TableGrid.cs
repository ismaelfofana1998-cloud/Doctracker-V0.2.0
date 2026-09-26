using System;
using System.Collections.Generic;
using System.Linq;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    // Coordinates are relative to the selected table region, independent of PDF zoom.
    public sealed class TableGrid
    {
        private readonly List<double> columns=new List<double>{0,1};
        private readonly List<double> rows=new List<double>{0,1};
        private const double Gap=.001;
        public IReadOnlyList<double> Columns=>columns.AsReadOnly();
        public IReadOnlyList<double> Rows=>rows.AsReadOnly();
        public int ColumnCount=>columns.Count-1;
        public int RowCount=>rows.Count-1;
        public bool Add(bool column,double at)
        {
            var stops=column?columns:rows;
            if(!Finite(at) || at<=Gap || at>=1-Gap || stops.Any(p=>Math.Abs(p-at)<Gap))return false;
            if((column?ColumnCount+1:ColumnCount)*(column?RowCount:RowCount+1)>10000)
                throw new InvalidOperationException("Limitez le tableau à 10 000 cellules.");
            stops.Add(at);stops.Sort();return true;
        }
        public bool Move(bool column,int index,double at)
        {
            var stops=column?columns:rows;
            if(index<=0 || index>=stops.Count-1 || !Finite(at))return false;
            var next=Math.Max(stops[index-1]+Gap,Math.Min(stops[index+1]-Gap,at));
            if(stops[index]==next)return false;stops[index]=next;return true;
        }
        public bool Remove(bool column,int index)
        {
            var stops=column?columns:rows;
            if(index<=0 || index>=stops.Count-1)return false;stops.RemoveAt(index);return true;
        }
        public static TableGrid FromWords(IReadOnlyList<WordRecord> words)
        {
            var grid=new TableGrid();var cells=LayoutExtractor.Extract(words);
            var rowGroups=cells.GroupBy(c=>c.Row).OrderBy(g=>g.Key).ToList();
            for(var i=1;i<rowGroups.Count;i++)grid.Add(false,(rowGroups[i-1].Max(c=>c.Y+c.Height)+rowGroups[i].Min(c=>c.Y))/2);
            var template=cells.GroupBy(c=>c.Row).OrderByDescending(g=>g.Count()).FirstOrDefault();
            if(template!=null)
            {
                var anchors=template.OrderBy(c=>c.X).ToList();
                for(var i=1;i<anchors.Count;i++)grid.Add(true,(anchors[i-1].X+anchors[i-1].Width+anchors[i].X)/2);
            }
            return grid;
        }
        public IReadOnlyList<ExtractedCell> Extract(IReadOnlyList<WordRecord> words)
        {
            var buckets=new Dictionary<int,List<WordRecord>>();
            foreach(var word in words??new WordRecord[0])
            {
                var x=word.X+word.Width/2;var y=word.Y+word.Height/2;
                if(!Finite(x)||!Finite(y)||x<0||x>1||y<0||y>1)continue;
                var column=Band(columns,x);var row=Band(rows,y);var key=row*ColumnCount+column;
                if(!buckets.TryGetValue(key,out var group))buckets[key]=group=new List<WordRecord>();group.Add(word);
            }
            var result=new List<ExtractedCell>();
            for(var row=0;row<RowCount;row++)for(var column=0;column<ColumnCount;column++)
            {
                buckets.TryGetValue(row*ColumnCount+column,out var group);
                var text=group==null?"":string.Join("\n",group.GroupBy(w=>w.Line).OrderBy(g=>g.Min(w=>w.Y))
                    .Select(g=>string.Join(" ",g.OrderBy(w=>w.X).Select(w=>w.Text))));
                result.Add(new ExtractedCell {Row=row,Column=column,Text=text,X=columns[column],Y=rows[row],
                    Width=columns[column+1]-columns[column],Height=rows[row+1]-rows[row]});
            }
            return result;
        }
        private static int Band(List<double> stops,double at)
        {for(var i=1;i<stops.Count-1;i++)if(at<stops[i])return i-1;return stops.Count-2;}
        private static bool Finite(double value)=>!double.IsNaN(value)&&!double.IsInfinity(value);
    }
}
