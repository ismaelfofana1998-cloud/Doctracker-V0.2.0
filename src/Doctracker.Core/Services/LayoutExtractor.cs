using System;
using System.Collections.Generic;
using System.Linq;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    public sealed class ExtractedCell
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string Text { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
    }

    public static class LayoutExtractor
    {
        public static IReadOnlyList<ExtractedCell> Extract(IReadOnlyList<WordRecord> words)
        {
            var result = new List<ExtractedCell>();
            if (words == null || words.Count == 0) return result;
            var lines = words.GroupBy(word => word.Line).OrderBy(line => line.Min(word => word.Y)).ToList();
            var row = 0;
            foreach (var line in lines)
            {
                var ordered = line.OrderBy(word => word.X).ToList();
                var group = new List<WordRecord>();
                WordRecord previous = null;
                foreach (var word in ordered)
                {
                    var threshold = Math.Max(0.008, word.Height * 0.8);
                    if (previous != null && word.X - previous.X - previous.Width > threshold)
                    {
                        result.Add(Cell(group, row));
                        group.Clear();
                    }
                    group.Add(word);
                    previous = word;
                }
                if (group.Count > 0) result.Add(Cell(group, row));
                row++;
            }
            // Use the row with most cells as the column template. This preserves
            // multi-word headers and right-aligned amounts without inventing columns.
            var template = result.GroupBy(cell => cell.Row).OrderByDescending(line => line.Count()).First()
                .OrderBy(cell => cell.X).ToList();
            foreach (var cell in result)
            {
                cell.Column = template.Select((anchor, index) => new { index, distance = Distance(cell, anchor) })
                    .OrderBy(item => item.distance).First().index;
            }
            return result.GroupBy(cell => new { cell.Row, cell.Column }).Select(group => new ExtractedCell
            {
                Row = group.Key.Row, Column = group.Key.Column,
                Text = string.Join(" ", group.OrderBy(cell => cell.X).Select(cell => cell.Text)),
                X = group.Min(cell => cell.X), Y = group.Min(cell => cell.Y),
                Width = group.Max(cell => cell.X + cell.Width) - group.Min(cell => cell.X),
                Height = group.Max(cell => cell.Y + cell.Height) - group.Min(cell => cell.Y)
            }).OrderBy(cell => cell.Row).ThenBy(cell => cell.Column).ToList();
        }

        private static double Distance(ExtractedCell cell, ExtractedCell anchor) =>
            Math.Min(Math.Abs(cell.X - anchor.X), Math.Abs(cell.X + cell.Width - anchor.X - anchor.Width));

        private static ExtractedCell Cell(List<WordRecord> words, int row) => new ExtractedCell
        {
            Row = row, Text = string.Join(" ", words.Select(word => word.Text)),
            X = words.Min(word => word.X), Y = words.Min(word => word.Y),
            Width = words.Max(word => word.X + word.Width) - words.Min(word => word.X),
            Height = words.Max(word => word.Y + word.Height) - words.Min(word => word.Y)
        };
    }
}
