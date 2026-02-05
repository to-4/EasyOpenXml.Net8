using EasyOpenXml.Excel;
using EasyOpenXml.Excel.Tests.Net8.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace EasyOpenXml.Excel.Tests.Net8.ExcelDocumentTests
{
    [TestClass]
    public sealed class ExcelDocumentSetGetTests
    {
        [TestMethod]
        public void SheetSelect_SetsActiveSheet()
        {
            var path = TestFiles.CreateTempPath("SheetSelect");
            WorkbookFactory.CreateWorkbook(path, "Sheet1", "Sheet2");

            using var doc = new ExcelDocument();
            doc.InitializeFile(path);
            doc.SheetSelect(1);
            doc.SetValue("A1", "OnSheet2");
            doc.FinalizeFile(true);

            OpenXmlAssert.CellEqualsString(path, "Sheet2", "A1", "OnSheet2");
        }

        [TestMethod]
        public void SetValue_IntRange_WritesAllCells()
        {
            var path = TestFiles.CreateTempPath("SetValueRange");
            WorkbookFactory.CreateWorkbook(path);

            using var doc = new ExcelDocument();
            doc.InitializeFile(path);
            doc.SetValue(1, 1, 2, 2, 5);
            doc.FinalizeFile(true);

            OpenXmlAssert.CellEqualsNumber(path, "Sheet1", "A1", 5);
            OpenXmlAssert.CellEqualsNumber(path, "Sheet1", "B1", 5);
            OpenXmlAssert.CellEqualsNumber(path, "Sheet1", "A2", 5);
            OpenXmlAssert.CellEqualsNumber(path, "Sheet1", "B2", 5);
        }

        [TestMethod]
        public void SetValue_StringCell_WritesSharedString()
        {
            var path = TestFiles.CreateTempPath("SetValueString");
            WorkbookFactory.CreateWorkbook(path);

            using var doc = new ExcelDocument();
            doc.InitializeFile(path);
            doc.SetValue("B2", "Hello");
            doc.FinalizeFile(true);

            OpenXmlAssert.CellEqualsString(path, "Sheet1", "B2", "Hello");
        }

        [TestMethod]
        public void SetValue_CellWithOffsets_WritesRange()
        {
            var path = TestFiles.CreateTempPath("SetValueOffsets");
            WorkbookFactory.CreateWorkbook(path);

            using var doc = new ExcelDocument();
            doc.InitializeFile(path);
            doc.SetValue("B2", 1, 1, "X");
            doc.FinalizeFile(true);

            OpenXmlAssert.CellEqualsString(path, "Sheet1", "B2", "X");
            OpenXmlAssert.CellEqualsString(path, "Sheet1", "C2", "X");
            OpenXmlAssert.CellEqualsString(path, "Sheet1", "B3", "X");
            OpenXmlAssert.CellEqualsString(path, "Sheet1", "C3", "X");
        }

        [TestMethod]
        public void GetValue_ReturnsTypedValues()
        {
            var path = TestFiles.CreateTempPath("GetValue");
            WorkbookFactory.CreateWorkbook(path);

            using var doc = new ExcelDocument();
            doc.InitializeFile(path);
            doc.SetValue(1, 1, "Hello");
            doc.SetValue(2, 1, true);
            doc.SetValue(3, 1, 123.5);
            doc.SetValue(4, 1, new DateTime(2026, 2, 4));

            Assert.AreEqual("Hello", doc.GetValue("A1"));
            Assert.AreEqual(true, doc.GetValue("B1"));
            Assert.AreEqual(123.5, (double)doc.GetValue("C1"), 0.0000001);

            var expectedOa = new DateTime(2026, 2, 4).ToOADate();
            Assert.AreEqual(expectedOa, (double)doc.GetValue("D1"), 0.0000001);
        }

        [TestMethod]
        public void Pos_CanWriteSingleCell()
        {
            var path = TestFiles.CreateTempPath("PosSingle");
            WorkbookFactory.CreateWorkbook(path);

            using var doc = new ExcelDocument();
            doc.InitializeFile(path);
            var pos = doc.Pos(2, 2);
            pos.Value = 10;
            doc.FinalizeFile(true);

            OpenXmlAssert.CellEqualsNumber(path, "Sheet1", "B2", 10);
        }

        [TestMethod]
        public void Pos_CanWriteRange()
        {
            var path = TestFiles.CreateTempPath("PosRange");
            WorkbookFactory.CreateWorkbook(path);

            using var doc = new ExcelDocument();
            doc.InitializeFile(path);
            doc.Pos(1, 1, 2, 2).Value = 7;
            doc.FinalizeFile(true);

            OpenXmlAssert.CellEqualsNumber(path, "Sheet1", "A1", 7);
            OpenXmlAssert.CellEqualsNumber(path, "Sheet1", "B1", 7);
            OpenXmlAssert.CellEqualsNumber(path, "Sheet1", "A2", 7);
            OpenXmlAssert.CellEqualsNumber(path, "Sheet1", "B2", 7);
        }
    }
}
