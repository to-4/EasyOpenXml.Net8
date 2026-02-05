using DocumentFormat.OpenXml.Packaging;
using EasyOpenXml.Excel;
using EasyOpenXml.Excel.Tests.Net8.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EasyOpenXml.Excel.Tests.Net8.ExcelDocumentTests
{
    [TestClass]
    public sealed class ExcelDocumentRowDeleteTests
    {
        [TestMethod]
        public void RowDelete_RemovesAndShiftsRows()
        {
            var path = TestFiles.CreateTempPath("RowDelete");
            WorkbookFactory.CreateWorkbook(path);

            using (var doc = new ExcelDocument())
            {
                doc.InitializeFile(path);
                doc.SetValue(1, 1, 1);
                doc.SetValue(1, 2, 2);
                doc.SetValue(1, 3, 3);
                doc.SetValue(1, 4, 4);
                doc.SetValue(1, 5, 5);

                doc.RowDelete(1, 2); // delete rows 2-3 (0-based)
                doc.FinalizeFile(true);
            }

            OpenXmlAssert.CellEqualsNumber(path, "Sheet1", "A1", 1);
            OpenXmlAssert.CellEqualsNumber(path, "Sheet1", "A2", 4);
            OpenXmlAssert.CellEqualsNumber(path, "Sheet1", "A3", 5);

            using var xdoc = SpreadsheetDocument.Open(path, false);
            var wsPart = OpenXmlAssert.GetWorksheetPartByName(xdoc, "Sheet1");
            var a4 = OpenXmlAssert.GetCell(wsPart, "A4");
            Assert.IsNull(a4, "Expected A4 to be removed after row delete shift.");
        }

        [TestMethod]
        public void RowDelete_InvalidArguments_Throw()
        {
            var path = TestFiles.CreateTempPath("RowDeleteInvalid");
            WorkbookFactory.CreateWorkbook(path);

            using var doc = new ExcelDocument();
            doc.InitializeFile(path);

            Assert.ThrowsException<System.ArgumentOutOfRangeException>(() => doc.RowDelete(-1, 1));
            Assert.ThrowsException<System.ArgumentOutOfRangeException>(() => doc.RowDelete(0, 0));
        }
    }
}
