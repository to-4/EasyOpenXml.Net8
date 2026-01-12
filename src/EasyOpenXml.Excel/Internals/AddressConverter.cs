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
    }
}
