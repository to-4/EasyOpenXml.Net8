using System;
using System.IO;

namespace EasyOpenXml.Excel.Tests.Helpers
{
    internal static class TestFiles
    {
        internal static string CreateTempPath(string prefix, string ext = ".xlsx")
        {
            // Create a unique temp file path per test to avoid race conditions.
            var dir = Path.Combine(Path.GetTempPath(), "EasyOpenXml.Excel.Tests");
            Directory.CreateDirectory(dir);

            var name = $"{prefix}_{Guid.NewGuid():N}{ext}";
            return Path.Combine(dir, name);
        }
    }
}
