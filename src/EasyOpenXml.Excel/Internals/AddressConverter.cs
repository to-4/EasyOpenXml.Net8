using System;
using System.Text.RegularExpressions;

namespace EasyOpenXml.Excel.Internals
{
    internal static class AddressConverter
    {
        private static readonly Regex A1Regex = new Regex(@"^([A-Za-z]+)(\d+)$", RegexOptions.Compiled);

        internal static string ToA1(int col, int row)
        {
            // 1. Convert 1-based column index to letters (1 -> A, 26 -> Z, 27 -> AA)
            var colLetters = ToColumnLetters(col);
            return colLetters + row.ToString();
        }

        internal static bool TryParseA1(string a1, out int col, out int row)
        {
            col = 0;
            row = 0;

            if (string.IsNullOrEmpty(a1)) return false;

            var m = A1Regex.Match(a1);
            if (!m.Success) return false;

            col = FromColumnLetters(m.Groups[1].Value);
            if (!int.TryParse(m.Groups[2].Value, out row)) return false;

            return col > 0 && row > 0;
        }

        private static string ToColumnLetters(int col)
        {
            if (col <= 0) throw new ArgumentOutOfRangeException(nameof(col));

            var result = string.Empty;
            var n = col;

            while (n > 0)
            {
                n--; // 1-based to 0-based
                var c = (char)('A' + (n % 26));
                result = c + result;
                n /= 26;
            }

            return result;
        }

        private static int FromColumnLetters(string letters)
        {
            if (string.IsNullOrEmpty(letters)) return 0;

            var n = 0;
            foreach (var ch in letters.ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') return 0;
                n = (n * 26) + (ch - 'A' + 1);
            }
            return n;
        }

        internal static bool TryParseA1Range(string a1Range, out int sx, out int sy, out int ex, out int ey)
        {
            sx = sy = ex = ey = 0;
            if (string.IsNullOrWhiteSpace(a1Range)) return false;

            var text = a1Range.Trim();

            // allow absolute refs like $A$1:$D$20
            text = text.Replace("$", "");

            // single cell like "B2"
            var parts = text.Split(':');
            if (parts.Length == 1)
            {
                if (!TryParseA1(parts[0], out var c, out var r)) return false;
                sx = ex = c;
                sy = ey = r;
                return true;
            }

            // range like "A1:D20"
            if (parts.Length == 2)
            {
                if (!TryParseA1(parts[0], out var c1, out var r1)) return false;
                if (!TryParseA1(parts[1], out var c2, out var r2)) return false;

                sx = Math.Min(c1, c2);
                ex = Math.Max(c1, c2);
                sy = Math.Min(r1, r2);
                ey = Math.Max(r1, r2);
                return true;
            }

            return false;
        }

        internal static string ToAbsoluteA1(int col, int row)
        {
            // ToA1 は "A1" を返す前提
            var a1 = ToA1(col, row);
            // "A1" -> "$A$1"
            var letters = new string(a1.TakeWhile(char.IsLetter).ToArray());
            var digits = new string(a1.SkipWhile(char.IsLetter).ToArray());
            return $"${letters}${digits}";
        }

        internal static string ToColumnName(int col)
        {
            int c = col;
            string s = "";
            while (c > 0)
            {
                c--;
                s = (char)('A' + (c % 26)) + s;
                c /= 26;
            }
            return s;
        }
    }
}
