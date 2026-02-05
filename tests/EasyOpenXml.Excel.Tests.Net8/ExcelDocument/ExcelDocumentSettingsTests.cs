using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EasyOpenXml.Excel;
using EasyOpenXml.Excel.Models;
using EasyOpenXml.Excel.Tests.Net8.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace EasyOpenXml.Excel.Tests.Net8.ExcelDocumentTests
{
    [TestClass]
    public sealed class ExcelDocumentSettingsTests
    {
        [TestMethod]
        public void SetCalculationMode_UpdatesWorkbook()
        {
            var path = TestFiles.CreateTempPath("CalcMode");
            WorkbookFactory.CreateWorkbook(path);

            using (var doc = new ExcelDocument())
            {
                doc.InitializeFile(path);
                doc.SetCalculationMode(CalculationMode.Manual);
                doc.FinalizeFile(true);
            }

            using var xdoc = SpreadsheetDocument.Open(path, false);
            var calc = xdoc.WorkbookPart!.Workbook.CalculationProperties;
            Assert.IsNotNull(calc);
            Assert.AreEqual(CalculateModeValues.Manual, calc.CalculationMode!.Value);
        }

        [TestMethod]
        public void PrintArea_IntRange_CreatesDefinedName()
        {
            var path = TestFiles.CreateTempPath("PrintAreaInt");
            WorkbookFactory.CreateWorkbook(path);

            using (var doc = new ExcelDocument())
            {
                doc.InitializeFile(path);
                doc.PrintArea(1, 1, 2, 2);
                doc.FinalizeFile(true);
            }

            using var xdoc = SpreadsheetDocument.Open(path, false);
            var definedNames = xdoc.WorkbookPart!.Workbook.DefinedNames;
            Assert.IsNotNull(definedNames);

            var printArea = definedNames.Elements<DefinedName>()
                .FirstOrDefault(d => d.Name?.Value == "_xlnm.Print_Area");

            Assert.IsNotNull(printArea);
            Assert.AreEqual("'Sheet1'!$A$1:$B$2", printArea!.Text);
        }

        [TestMethod]
        public void PrintArea_A1Range_CreatesDefinedName()
        {
            var path = TestFiles.CreateTempPath("PrintAreaA1");
            WorkbookFactory.CreateWorkbook(path);

            using (var doc = new ExcelDocument())
            {
                doc.InitializeFile(path);
                doc.PrintArea("A1:C3");
                doc.FinalizeFile(true);
            }

            using var xdoc = SpreadsheetDocument.Open(path, false);
            var printArea = xdoc.WorkbookPart!.Workbook.DefinedNames!
                .Elements<DefinedName>()
                .FirstOrDefault(d => d.Name?.Value == "_xlnm.Print_Area");

            Assert.IsNotNull(printArea);
            Assert.AreEqual("'Sheet1'!$A$1:$C$3", printArea!.Text);
        }

        [TestMethod]
        public void PrintArea_InvalidRange_Throws()
        {
            var path = TestFiles.CreateTempPath("PrintAreaInvalid");
            WorkbookFactory.CreateWorkbook(path);

            using var doc = new ExcelDocument();
            doc.InitializeFile(path);

            Assert.ThrowsException<ArgumentException>(() => doc.PrintArea("A1:"));
        }
    }
}
