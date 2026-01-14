using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace EasyOpenXml.Excel.Internals
{
    internal static class ExcelInternalAccessor
    {
        private static readonly Dictionary<SpreadsheetDocument, CellSnapshot> _clipboards
            = new Dictionary<SpreadsheetDocument, CellSnapshot>();

        internal static void SetClipboard(SpreadsheetDocument doc, CellSnapshot snapshot)
        {
            _clipboards[doc] = snapshot;
        }

        internal static CellSnapshot GetClipboard(SpreadsheetDocument doc)
        {
            _clipboards.TryGetValue(doc, out var snapshot);
            return snapshot;
        }
    }
}
