using EasyOpenXml.Excel;
using EasyOpenXml.Excel.Tests.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace EasyOpenXml.Excel.Tests
{
    [TestClass]
    public sealed class ExcelDocument_SetValue_Tests
    {
        [TestMethod]
        public void SetValue_String_WritesSharedString()
        {
            var path = TestFiles.CreateTempPath("SetValue_String");

            using (var doc = ExcelDocument.Create(path))
            {
                doc.SelectSheet("Sheet1")
                   .Pos(1, 1)
                   .SetValue("Hello", isString: true);

                doc.Save();
            }

            OpenXmlAssert.CellEqualsString(path, "Sheet1", "A1", "Hello");
        }

        [TestMethod]
        public void SetValue_Number_WritesNumeric()
        {
            var path = TestFiles.CreateTempPath("SetValue_Number");

            using (var doc = ExcelDocument.Create(path))
            {
                doc.SelectSheet("Sheet1")
                   .Pos(2, 3)
                   .SetValue(123.45, isString: false);

                doc.Save();
            }

            OpenXmlAssert.CellEqualsNumber(path, "Sheet1", "C2", 123.45);
        }

        [TestMethod]
        public void SetValue_Bool_WritesBoolean()
        {
            var path = TestFiles.CreateTempPath("SetValue_Bool");

            using (var doc = ExcelDocument.Create(path))
            {
                doc.SelectSheet("Sheet1")
                   .Pos(3, 2)
                   .SetValue(true, isString: false);

                doc.Save();
            }

            OpenXmlAssert.CellEqualsBoolean(path, "Sheet1", "B3", true);
        }

        [TestMethod]
        public void SetValue_DateTime_WritesOADate()
        {
            var path = TestFiles.CreateTempPath("SetValue_DateTime");
            var dt = new DateTime(2026, 2, 4, 0, 0, 0, DateTimeKind.Unspecified);
            var expectedOa = dt.ToOADate();

            using (var doc = ExcelDocument.Create(path))
            {
                doc.SelectSheet("Sheet1")
                   .Pos(4, 1)
                   .SetValue(dt, isString: false);

                doc.Save();
            }

            OpenXmlAssert.CellEqualsNumber(path, "Sheet1", "A4", expectedOa);
        }
    }
}
