using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    public sealed class TextValueParser
    {
        // A space joins thousands only when followed by exactly three digits.
        // Tabs, line breaks and two spaces delimit separate table cells.
        private static readonly Regex NumberPattern = new Regex(
            @"\(?[-+−]?(?:\d{1,3}(?:[ \u00A0\u202F]\d{3}(?!\d))+|\d+)(?:[.,]\d+)*-?\)?",
            RegexOptions.Compiled);
        private static readonly string[] DateFormats =
        {
            "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy",
            "yyyy-MM-dd", "dd.MM.yyyy", "d.M.yyyy"
        };

        public string Parse(SnipType type, string rawText)
        {
            var text = (rawText ?? string.Empty).Trim();
            if (type == SnipType.Validation || type == SnipType.Exception) return text;
            if (text.Length == 0) throw new FormatException("Aucun texte reconnu dans cette zone. Agrandissez la sélection.");
            switch (type)
            {
                case SnipType.Text: return Regex.Replace(text, @"\s+", " ");
                case SnipType.Number: return Format(ParseNumber(text));
                case SnipType.Date: return ParseDate(text).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                case SnipType.Sum: return Format(ParseAllNumbers(text).Sum());
                case SnipType.Table:
                    return string.Join("\n", Regex.Split(text, @"\r?\n")
                        .Select(line => Regex.Replace(line.Trim(), @"(?: {2,}|\t+)", "\t"))
                        .Where(line => line.Length > 0));
                default: throw new ArgumentOutOfRangeException(nameof(type));
            }
        }

        public decimal ParseNumber(string text)
        {
            var matches = NumberPattern.Matches(text ?? string.Empty);
            decimal value;
            if (matches.Count != 1 || !TryParseNumericToken(matches[0].Value, out value))
                throw new FormatException("Sélectionnez un seul nombre non ambigu (ou utilisez Somme).");
            return value;
        }

        public bool TryParseAmount(string text, out decimal value)
        {
            value = 0;
            var token = Regex.Replace(text ?? string.Empty, @"(?i)\b(?:EUR|USD|GBP|FCFA|XOF|CHF)\b|[€$£]", "").Trim();
            var match = NumberPattern.Match(token);
            return match.Success && match.Index == 0 && match.Length == token.Length &&
                TryParseNumericToken(token, out value);
        }

        public IReadOnlyList<decimal> ParseAllNumbers(string text)
        {
            var values = new List<decimal>();
            foreach (Match match in NumberPattern.Matches(text ?? string.Empty))
            {
                decimal value;
                if (!TryParseNumericToken(match.Value, out value))
                    throw new FormatException("Un montant est ambigu : " + match.Value);
                values.Add(value);
            }
            if (values.Count == 0) throw new FormatException("Aucun montant reconnu dans cette zone.");
            return values;
        }

        public static DateTime ParseDate(string text)
        {
            var dates = new List<DateTime>();
            foreach (Match match in Regex.Matches(text ?? string.Empty, @"(?<!\d)\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}(?!\d)"))
            {
                DateTime result;
                if (DateTime.TryParseExact(match.Value, DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
                    dates.Add(result);
            }
            if (dates.Count != 1) throw new FormatException("Sélectionnez une seule date valide (jour/mois/année).");
            return dates[0];
        }

        private static string Format(decimal value) => value.ToString("0.############################", CultureInfo.InvariantCulture);

        private static bool TryParseNumericToken(string token, out decimal value)
        {
            value = 0;
            var normalized = Regex.Replace(token ?? "", @"[ \u00A0\u202F]", "").Replace('−', '-');
            var parentheses = normalized.StartsWith("(") && normalized.EndsWith(")");
            if (parentheses) normalized = normalized.Substring(1, normalized.Length - 2);
            else if (normalized.Contains("(") || normalized.Contains(")")) return false;
            if (normalized.EndsWith("-"))
            {
                if (parentheses || normalized.StartsWith("-")) return false;
                normalized = "-" + normalized.Substring(0, normalized.Length - 1);
            }
            var comma = normalized.LastIndexOf(',');
            var dot = normalized.LastIndexOf('.');
            if (comma >= 0 && dot >= 0)
            {
                var decimalSeparator = comma > dot ? ',' : '.';
                var groupingSeparator = comma > dot ? '.' : ',';
                var pieces = normalized.Split(decimalSeparator);
                if (pieces.Length != 2 || !ValidGroupedInteger(pieces[0], groupingSeparator)) return false;
                normalized = pieces[0].Replace(groupingSeparator.ToString(), "") + "." + pieces[1];
            }
            else if (comma >= 0 || dot >= 0)
            {
                var separator = comma >= 0 ? ',' : '.';
                var parts = normalized.Split(separator);
                if (ValidGroupedInteger(normalized, separator)) normalized = normalized.Replace(separator.ToString(), "");
                else if (parts.Length == 2) normalized = parts[0] + "." + parts[1];
                else return false;
            }
            if (!decimal.TryParse(normalized, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture, out value)) return false;
            if (parentheses)
            {
                if (value < 0) return false;
                value = -value;
            }
            return true;
        }

        private static bool ValidGroupedInteger(string text, char separator)
        {
            return Regex.IsMatch(text, @"^[-+]?\d{1,3}(?:" + Regex.Escape(separator.ToString()) + @"\d{3})+$");
        }
    }
}
