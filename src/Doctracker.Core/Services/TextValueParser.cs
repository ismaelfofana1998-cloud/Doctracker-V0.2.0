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
        private static readonly Regex NumberPattern =
            new Regex(@"[-+]?\d(?:[\d \t\u00A0.,]*\d)?", RegexOptions.Compiled);

        private static readonly string[] DateFormats =
        {
            "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy",
            "yyyy-MM-dd", "dd.MM.yyyy", "d.M.yyyy"
        };

        public string Parse(SnipType type, string rawText)
        {
            var text = (rawText ?? string.Empty).Trim();
            switch (type)
            {
                case SnipType.Text:
                    return Regex.Replace(text, @"\s+", " ");
                case SnipType.Number:
                    return ParseNumber(text).ToString("0.################", CultureInfo.InvariantCulture);
                case SnipType.Date:
                    return ParseDate(text).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                case SnipType.Sum:
                    return ParseAllNumbers(text).Sum().ToString("0.################", CultureInfo.InvariantCulture);
                case SnipType.Table:
                    return NormalizeTable(text);
                default:
                    throw new ArgumentOutOfRangeException(nameof(type));
            }
        }

        public decimal ParseNumber(string text)
        {
            var match = NumberPattern.Match(text ?? string.Empty);
            if (!match.Success)
            {
                throw new FormatException("No number was recognized in the selected zone.");
            }

            return ParseNumericToken(match.Value);
        }

        public IReadOnlyList<decimal> ParseAllNumbers(string text)
        {
            var values = new List<decimal>();
            foreach (Match match in NumberPattern.Matches(text ?? string.Empty))
            {
                decimal value;
                if (TryParseNumericToken(match.Value, out value))
                {
                    values.Add(value);
                }
            }

            if (values.Count == 0)
            {
                throw new FormatException("No amount was recognized in the selected zone.");
            }

            return values;
        }

        private static DateTime ParseDate(string text)
        {
            DateTime result;
            foreach (Match match in Regex.Matches(text ?? string.Empty, @"\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}"))
            {
                if (DateTime.TryParseExact(match.Value, DateFormats, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out result))
                {
                    return result;
                }
            }

            throw new FormatException("No supported date was recognized in the selected zone.");
        }

        private static decimal ParseNumericToken(string token)
        {
            decimal value;
            if (!TryParseNumericToken(token, out value))
            {
                throw new FormatException("The recognized number is ambiguous.");
            }
            return value;
        }

        private static bool TryParseNumericToken(string token, out decimal value)
        {
            value = 0m;
            var normalized = Regex.Replace(token ?? string.Empty, @"[\s\u00A0]", string.Empty);
            if (normalized.Length == 0) return false;

            var lastComma = normalized.LastIndexOf(',');
            var lastDot = normalized.LastIndexOf('.');
            var decimalIndex = Math.Max(lastComma, lastDot);

            if (decimalIndex >= 0 && normalized.Length - decimalIndex - 1 <= 2)
            {
                var integerPart = Regex.Replace(normalized.Substring(0, decimalIndex), @"[.,]", string.Empty);
                var decimals = normalized.Substring(decimalIndex + 1);
                normalized = integerPart + "." + decimals;
            }
            else
            {
                normalized = Regex.Replace(normalized, @"[.,]", string.Empty);
            }

            return decimal.TryParse(normalized, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture, out value);
        }

        private static string NormalizeTable(string text)
        {
            var lines = Regex.Split(text ?? string.Empty, @"\r?\n")
                .Select(line => Regex.Replace(line.Trim(), @"\s{2,}", "\t"))
                .Where(line => line.Length > 0);
            return string.Join(Environment.NewLine, lines);
        }
    }
}
