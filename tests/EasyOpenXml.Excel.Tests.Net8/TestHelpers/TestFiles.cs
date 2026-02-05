using System;
using System.IO;

namespace EasyOpenXml.Excel.Tests.Net8.TestHelpers
{
    internal static class TestFiles
    {
        internal static string CreateTempPath(string prefix, string ext = ".xlsx")
        {
            var dir = Path.Combine(Path.GetTempPath(), "EasyOpenXml.Excel.Tests.Net8");
            Directory.CreateDirectory(dir);

            var name = $"{prefix}_{Guid.NewGuid():N}{ext}";
            return Path.Combine(dir, name);
        }
    }
}
