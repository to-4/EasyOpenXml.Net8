using EasyOpenXml.Excel;
using EasyOpenXml.Excel.Tests.Net8.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace EasyOpenXml.Excel.Tests.Net8.ExcelDocumentTests
{
    [TestClass]
    public sealed class ExcelDocumentInitializationTests
    {
        [TestMethod]
        public void InitializeFile_LoadsSheetNames()
        {
            var path = TestFiles.CreateTempPath("Init");
            WorkbookFactory.CreateWorkbook(path, "Sheet1", "Sheet2");

            using var doc = new ExcelDocument();
            doc.InitializeFile(path);

            var names = doc.SheetNames.ToArray();
            CollectionAssert.AreEqual(new[] { "Sheet1", "Sheet2" }, names);
        }

        [TestMethod]
        public void InitializeFile_InvalidPath_Throws()
        {
            using var doc = new ExcelDocument();

            Assert.ThrowsException<ExcelDocumentException>(
                () => doc.InitializeFile("/no/such/file.xlsx"));
        }

        [TestMethod]
        public void Dispose_PreventsFurtherUse()
        {
            var path = TestFiles.CreateTempPath("Dispose");
            WorkbookFactory.CreateWorkbook(path);

            var doc = new ExcelDocument();
            doc.InitializeFile(path);
            doc.Dispose();

            Assert.ThrowsException<ObjectDisposedException>(
                () => _ = doc.SheetNames);
        }
    }
}
