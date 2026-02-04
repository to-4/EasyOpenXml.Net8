using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Globalization;
using System.Linq;

namespace EasyOpenXml.Excel.Tests.Helpers
{
    internal static class OpenXmlAssert
    {
        internal static void CellEqualsString(string path, string sheetName, string a1, string expected)
        {
            using (var doc = SpreadsheetDocument.Open(path, false))
            {
                var wsPart = GetWorksheetPartByName(doc, sheetName);
                var cell = GetCell(wsPart, a1);
                Assert.IsNotNull(cell, $"Cell not found: {a1}");

                var actual = ReadCellText(doc, cell);
                Assert.AreEqual(expected, actual);
            }
        }

        internal static void CellEqualsNumber(string path, string sheetName, string a1, double expected, double delta = 0.0000001)
        {
            using (var doc = SpreadsheetDocument.Open(path, false))
            {
                var wsPart = GetWorksheetPartByName(doc, sheetName);
                var cell = GetCell(wsPart, a1);
                Assert.IsNotNull(cell, $"Cell not found: {a1}");

                var raw = cell.CellValue?.Text;
                Assert.IsNotNull(raw, "CellValue is null.");

                var actual = double.Parse(raw, CultureInfo.InvariantCulture);
                Assert.IsTrue(Math.Abs(expected - actual) <= delta, $"Expected {expected}, Actual {actual}");
            }
        }

        internal static void CellEqualsBoolean(string path, string sheetName, string a1, bool expected)
        {
            using (var doc = SpreadsheetDocument.Open(path, false))
            {
                var wsPart = GetWorksheetPartByName(doc, sheetName);
                var cell = GetCell(wsPart, a1);
                Assert.IsNotNull(cell, $"Cell not found: {a1}");

                var raw = cell.CellValue?.Text ?? "0";
                var actual = raw == "1";
                Assert.AreEqual(expected, actual);
            }
        }

        private static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument doc, string sheetName)
        {
            var sheets = doc.WorkbookPart.Workbook.Sheets.Elements<Sheet>();
            var sheet = sheets.First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal));
            return (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
        }

        private static Cell GetCell(WorksheetPart wsPart, string a1)
        {
            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return null;

            return sheetData.Descendants<Cell>()
                .FirstOrDefault(c => string.Equals(c.CellReference?.Value, a1, StringComparison.OrdinalIgnoreCase));
        }

        private static string ReadCellText(SpreadsheetDocument doc, Cell cell)
        {
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var sst = doc.WorkbookPart.SharedStringTablePart.SharedStringTable;
                var index = int.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture);
                return sst.Elements<SharedStringItem>().ElementAt(index).InnerText ?? string.Empty;
            }

            return cell.CellValue?.Text ?? string.Empty;
        }
    }
}
