using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Doctracker.Core.Models;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Doctracker.AddIn.Excel
{
    internal sealed class ExcelCellGateway
    {
        internal const string MarkerPrefix = "DOCTRACKER-SNIP:";
        private readonly ExcelInterop.Application application;
        public ExcelCellGateway(ExcelInterop.Application application) { this.application = application; }

        public ExcelInterop.Range GetSingleTarget()
        {
            var range = GetSelection();
            if (range.Cells.CountLarge != 1) throw new InvalidOperationException("Sélectionnez une seule cellule de destination.");
            return range;
        }
        public ExcelInterop.Range GetSelection()
        {
            var range = application.Selection as ExcelInterop.Range;
            if (range == null || range.Areas.Count != 1) throw new InvalidOperationException("Sélectionnez une plage Excel continue.");
            return range;
        }

        public static void ValidateWritable(ExcelInterop.Range range)
        {
            if (range == null || range.Areas.Count != 1) throw new InvalidOperationException("La destination doit être une plage continue.");
            if ((bool)range.Worksheet.ProtectContents) throw new InvalidOperationException("La feuille est protégée. Déprotégez-la avant l'insertion.");
            if (!Equals(range.MergeCells, false)) throw new InvalidOperationException("La destination contient des cellules fusionnées.");
            if (!Equals(range.HasArray, false)) throw new InvalidOperationException("La destination contient une formule matricielle.");
        }

        public static bool HasContent(ExcelInterop.Range target) =>
            target.Value2 != null || Equals(target.HasFormula, true) || target.Comment != null;

        public void WriteSnip(ExcelInterop.Range target, SnipRecord snip, DocumentRecord document, bool appendProof = false)
        {
            if ((snip.Type == SnipType.Validation || snip.Type == SnipType.Exception) && target.Value2 == null && !(bool)target.HasFormula)
                target.Value2 = snip.Type == SnipType.Validation ? "Validation" : "Exception";
            if (snip.Type != SnipType.Validation && snip.Type != SnipType.Exception)
            {
                if (snip.Type == SnipType.Number || snip.Type == SnipType.Sum)
                {
                    target.Value2 = double.Parse(snip.ExtractedValue, CultureInfo.InvariantCulture);
                    target.NumberFormat = "#,##0.00";
                }
                else if (snip.Type == SnipType.Date)
                {
                    var date = DateTime.ParseExact(snip.ExtractedValue, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    // Value2 expects a serial, not a COM Date. Account for 1904 workbooks.
                    var workbook = (ExcelInterop.Workbook)target.Worksheet.Parent;
                    target.Value2 = date.ToOADate() - (workbook.Date1904 ? 1462 : 0);
                    target.NumberFormat = "dd/mm/yyyy";
                }
                else
                {
                    // Prevent OCR content beginning with =,+,-,@ from becoming a formula.
                    target.NumberFormat = "@";
                    target.Value2 = snip.ExtractedValue;
                }
            }
            AttachProof(target, snip, document, appendProof);
        }

        public void AttachProof(ExcelInterop.Range target, SnipRecord snip, DocumentRecord document, bool append = true)
        {
            var old = target.Comment == null ? "" : target.Comment.Text() ?? "";
            // The entire Doctracker section is tagged; preserve pre-existing user notes.
            var userNote = old.Split(new[] { "\n[DOCTRACKER]\n" }, StringSplitOptions.None)[0];
            if (userNote.StartsWith(MarkerPrefix, StringComparison.Ordinal)) userNote = ""; // legacy metadata
            var ids = append ? GetSnipIds(target).Where(id => id != snip.Id).ToList() : new List<string>();
            ids.Add(snip.Id);
            var text = userNote + "\n[DOCTRACKER]\n" + string.Join("\n", ids.Select(id => MarkerPrefix + id)) +
                "\nDocument : " + document.DisplayName + "\nPage : " + snip.PageNumber + "\nType : " + snip.Type +
                "\nStatut : " + snip.Status + "\nCommentaire : " + snip.Comment + "\nSélectionner la cellule pour afficher la preuve.";
            if (target.Comment == null) target.AddComment(text); else target.Comment.Text(text);
            target.Comment.Visible = false;
            var color = UI.SnipTheme.Tint(snip.SourceType ?? snip.Type);
            target.Interior.Color = ColorTranslator.ToOle(color);
        }

        public void RemoveProof(ExcelInterop.Range target)
        {
            if (GetSnipIds(target).Count == 0) return;
            var text = target.Comment.Text() ?? "";
            var note = text.Split(new[] { "\n[DOCTRACKER]\n" }, StringSplitOptions.None)[0];
            if (note.StartsWith(MarkerPrefix, StringComparison.Ordinal)) note = "";
            target.Comment.Delete();
            if (!string.IsNullOrEmpty(note)) target.AddComment(note);
            target.Interior.Pattern = ExcelInterop.XlPattern.xlPatternNone;
        }

        public void DetachProof(ExcelInterop.Range target, string id, ProjectState state)
        {
            var remaining=GetSnipIds(target).Where(value=>value!=id).ToList();
            if(remaining.Count==0){RemoveProof(target);return;}
            var text=target.Comment.Text() ?? "";
            target.Comment.Text(string.Join("\n",Regex.Split(text,@"\r?\n").Where(line=>line.Trim()!=MarkerPrefix+id)));
            var last=state.Snips.LastOrDefault(s=>remaining.Contains(s.Id));
            var document=last==null?null:state.Documents.FirstOrDefault(d=>d.Id==last.DocumentId);
            if(document!=null)AttachProof(target,last,document,true);
        }

        public IReadOnlyList<string> GetSnipIds(ExcelInterop.Range target)
        {
            if (target == null || target.Cells.CountLarge != 1 || target.Comment == null) return new List<string>();
            var text = target.Comment.Text() ?? "";
            return Regex.Split(text, @"\r?\n").Where(line => line.StartsWith(MarkerPrefix, StringComparison.Ordinal))
                .Select(line => line.Substring(MarkerPrefix.Length).Trim()).Where(id => id.Length > 0).Distinct().ToList();
        }
        public string GetSnipId(ExcelInterop.Range target) => GetSnipIds(target).LastOrDefault();

        public static string QueryText(ExcelInterop.Range cell)
        {
            var value = cell.Value; // preserves VT_DATE when Excel formatted the value as a date
            if (value is DateTime) return ((DateTime)value).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            var raw = cell.Value2;
            if (raw is double)
            {
                var number = ((double)raw).ToString("0.################", CultureInfo.InvariantCulture);
                // Three decimal places must not look like a thousands group to the text parser.
                var dot = number.IndexOf('.');
                if (dot >= 0 && number.Length - dot - 1 == 3) number += "0";
                return number;
            }
            return Convert.ToString(raw, CultureInfo.InvariantCulture);
        }

        public CellSnapshot Snapshot(ExcelInterop.Range target) => new CellSnapshot(target);
        internal sealed class CellSnapshot
        {
            private readonly ExcelInterop.Range target;
            private readonly object formula, value, format, color, pattern;
            private readonly bool hasFormula;
            private readonly string note;
            public CellSnapshot(ExcelInterop.Range target)
            {
                this.target = target;
                hasFormula = Equals(target.HasFormula, true);
                formula = target.Formula; value = target.Value2; format = target.NumberFormat;
                color = target.Interior.Color; pattern = target.Interior.Pattern;
                note = target.Comment == null ? null : target.Comment.Text();
            }
            public void Restore()
            {
                if (hasFormula) { target.NumberFormat = format; target.Formula = formula; }
                else { target.NumberFormat = "@"; target.Value2 = value; target.NumberFormat = format; }
                target.Interior.Color = color; target.Interior.Pattern = pattern;
                if (target.Comment != null) target.Comment.Delete();
                if (note != null) target.AddComment(note);
            }
        }
    }
}
