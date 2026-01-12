using EasyOpenXml.Excel;

namespace EasyOpenXml.Excel.DemoConsole.Demos;

internal static class Demo01_CreateNew
{
    public static void Run()
    {
        var path = Paths.OutFile("demo01_new.xlsx");

        //using var excel = ExcelDocument.Create(path);
        //excel.SetValue("A1", "Hello");
        //excel.SetValue("B1", "World");
        //excel.Save();

        Console.WriteLine("Created: " + path);
    }
}