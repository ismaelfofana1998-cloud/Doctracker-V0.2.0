using System;
using System.Globalization;
using System.Linq;
using Doctracker.Core.Models;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Doctracker.AddIn.Excel
{
    internal sealed class ExcelCellGateway
    {
        internal const string MarkerPrefix = "DOCTRACKER-SNIP:";
        private readonly ExcelInterop.Application application;

        public ExcelCellGateway(ExcelInterop.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public ExcelInterop.Range GetSingleTarget()
        {
            var range = application.Selection as ExcelInterop.Range;
            if (range == null || range.Cells.CountLarge != 1)
            {
                throw new InvalidOperationException("Select exactly one target cell.");
            }
            return range;
        }

        public ExcelInterop.Range GetSelection()
        {
            var range = application.Selection as ExcelInterop.Range;
            if (range == null) throw new InvalidOperationException("Select one or more Excel cells.");
            return range;
        }

        public void WriteSnip(ExcelInterop.Range target, SnipRecord snip, DocumentRecord document)
        {
            if (snip.Type == SnipType.Table)
            {
                WriteTable(target, snip.ExtractedValue);
            }
            else if (snip.Type == SnipType.Number || snip.Type == SnipType.Sum)
            {
                decimal number;
                if (decimal.TryParse(snip.ExtractedValue, NumberStyles.Any,
                    CultureInfo.InvariantCulture, out number))
                {
                    target.Value2 = Convert.ToDouble(number);
                    target.NumberFormat = "#,##0.00";
                }
                else
                {
                    throw new FormatException("The extracted amount cannot be written to Excel.");
                }
            }
            else if (snip.Type == SnipType.Date)
            {
                DateTime date;
                if (DateTime.TryParseExact(snip.ExtractedValue, "yyyy-MM-dd",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                {
                    target.Value2 = date;
                    target.NumberFormat = "dd/mm/yyyy";
                }
                else
                {
                    throw new FormatException("The extracted date cannot be written to Excel.");
                }
            }
            else
            {
                target.Value2 = snip.ExtractedValue;
            }

            AttachProof(target, snip, document);
        }

        public void AttachProof(ExcelInterop.Range target, SnipRecord snip, DocumentRecord document)
        {
            var commentText =
                MarkerPrefix + snip.Id + Environment.NewLine +
                "Document : " + document.OriginalName + Environment.NewLine +
                "Page : " + snip.PageNumber + Environment.NewLine +
                "Statut : " + snip.Status + Environment.NewLine +
                "Double-cliquer pour ouvrir la preuve.";

            if (target.Comment == null)
            {
                target.AddComment(commentText);
            }
            else
            {
                target.Comment.Text(commentText);
            }

            target.Comment.Visible = false;
            target.Interior.Color = System.Drawing.ColorTranslator.ToOle(
                System.Drawing.Color.FromArgb(255, 244, 225));
            target.Font.Color = System.Drawing.ColorTranslator.ToOle(
                System.Drawing.Color.FromArgb(18, 28, 45));
        }

        public string GetSnipId(ExcelInterop.Range target)
        {
            try
            {
                if (target == null || target.Comment == null) return null;
                var text = target.Comment.Text() ?? string.Empty;
                var marker = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
                    .FirstOrDefault(line => line.StartsWith(MarkerPrefix, StringComparison.Ordinal));
                return marker == null ? null : marker.Substring(MarkerPrefix.Length).Trim();
            }
            catch
            {
                return null;
            }
        }

        private static void WriteTable(ExcelInterop.Range startCell, string tableText)
        {
            var rows = (tableText ?? string.Empty)
                .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                .Select(line => line.Split('\t'))
                .ToArray();
            if (rows.Length == 0) throw new InvalidOperationException("No table data was recognized.");
            var columns = rows.Max(row => row.Length);
            var values = new object[rows.Length, columns];

            for (var row = 0; row < rows.Length; row++)
            {
                for (var column = 0; column < rows[row].Length; column++)
                {
                    values[row, column] = rows[row][column];
                }
            }

            var destination = startCell.Resize[rows.Length, columns];
            destination.Value2 = values;
        }
    }
}
