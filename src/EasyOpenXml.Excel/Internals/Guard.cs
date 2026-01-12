// src/EasyOpenXml.Excel/Internals/OpenXmlGuards.cs
using System;
using DocumentFormat.OpenXml.Packaging;

namespace EasyOpenXml.Excel.Internals
{
    internal static class Guards
    {
        // Excel (XLSX) limits
        internal const int MaxColumns = 16384;   // XFD
        internal const int MaxRows = 1048576;

        internal static void NotNull(object obj, string paramName)
        {
            if (obj == null) throw new ArgumentNullException(paramName);
        }

        internal static void EnsureOpened(bool opened)
        {
            if (!opened)
                throw new InvalidOperationException("The workbook is not opened. Call OpenBook first.");
        }

        internal static void EnsureWorksheetSelected(WorksheetPart worksheetPart)
        {
            if (worksheetPart == null)
                throw new InvalidOperationException("Worksheet is not selected. Set SheetNo first.");
        }

        internal static void Ensure1Based(int col, int row)
        {
            if (col <= 0 || row <= 0)
                throw new ArgumentOutOfRangeException("Coordinates must be 1-based and positive.");
        }

        internal static void EnsureRange(int sx, int sy, int ex, int ey)
        {
            Ensure1Based(sx, sy);
            Ensure1Based(ex, ey);

            // Normalize is done elsewhere; here we validate bounds only.
            var minCol = Math.Min(sx, ex);
            var maxCol = Math.Max(sx, ex);
            var minRow = Math.Min(sy, ey);
            var maxRow = Math.Max(sy, ey);

            if (minCol < 1 || maxCol > MaxColumns)
                throw new ArgumentOutOfRangeException($"Column must be in 1..{MaxColumns}.");
            if (minRow < 1 || maxRow > MaxRows)
                throw new ArgumentOutOfRangeException($"Row must be in 1..{MaxRows}.");
        }

        internal static void EnsureWorkbookPart(SpreadsheetDocument document)
        {
            NotNull(document, nameof(document));
            if (document.WorkbookPart == null)
                throw new InvalidOperationException("WorkbookPart is missing.");
        }
    }
}
