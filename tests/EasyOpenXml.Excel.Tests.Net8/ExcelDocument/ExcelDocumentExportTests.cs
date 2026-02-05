using EasyOpenXml.Excel;
using EasyOpenXml.Excel.Tests.Net8.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;

namespace EasyOpenXml.Excel.Tests.Net8.ExcelDocumentTests
{
    [TestClass]
    public sealed class ExcelDocumentExportTests
    {
        [TestMethod]
        public void ExportSharedFormulasCsv_WritesHeader()
        {
            var path = TestFiles.CreateTempPath("Export");
            WorkbookFactory.CreateWorkbook(path);

            var csv = TestFiles.CreateTempPath("shared_formulas", ".csv");

            using (var doc = new ExcelDocument())
            {
                doc.InitializeFile(path);
                doc.ExportSharedFormulasCsv(csv);
            }

            Assert.IsTrue(File.Exists(csv));
            var firstLine = File.ReadLines(csv).First();
            Assert.AreEqual("SheetName,Cell,Row,Col,SharedIndex,Formula,Reference", firstLine);
        }

        [TestMethod]
        public void ExportSharedFormulasCsv_InvalidPath_Throws()
        {
            var path = TestFiles.CreateTempPath("ExportInvalid");
            WorkbookFactory.CreateWorkbook(path);

            using var doc = new ExcelDocument();
            doc.InitializeFile(path);

            Assert.ThrowsException<System.ArgumentException>(() => doc.ExportSharedFormulasCsv(" "));
        }
    }
}
