using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace EasyOpenXml.Excel.Tests.Net8.TestHelpers
{
    internal static class WorkbookFactory
    {
        internal static void CreateWorkbook(string path, params string[] sheetNames)
        {
            if (sheetNames == null || sheetNames.Length == 0)
            {
                sheetNames = new[] { "Sheet1" };
            }

            using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();

            var sheets = wbPart.Workbook.AppendChild(new Sheets());

            uint sheetId = 1;
            foreach (var name in sheetNames)
            {
                var wsPart = wbPart.AddNewPart<WorksheetPart>();
                wsPart.Worksheet = new Worksheet(new SheetData());
                wsPart.Worksheet.Save();

                var relId = wbPart.GetIdOfPart(wsPart);
                var sheet = new Sheet
                {
                    Id = relId,
                    SheetId = sheetId,
                    Name = name
                };
                sheets.Append(sheet);
                sheetId++;
            }

            wbPart.Workbook.Save();
        }
    }
}
