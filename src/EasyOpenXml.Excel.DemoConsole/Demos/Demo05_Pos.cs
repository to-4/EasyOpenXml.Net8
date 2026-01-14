using EasyOpenXml.Excel;
using EasyOpenXml.Excel.Models;

namespace EasyOpenXml.Excel.DemoConsole.Demos;

internal static class Demo05_Pos
{
    public static void Run()
    {
        // テンプレ想定（DemoConsole 配下の "Assets/template.xlsx" ）
        var template = Path.Combine(AppContext.BaseDirectory, "Assets", "template.xlsx");
        if (!File.Exists(template))
        {
            Console.WriteLine("Template not found: " + template);
            Console.WriteLine("※ Assets/template.xlsx を配置してください（任意）");
            return;
        }

        var path = Paths.OutFile("Demo05_Pos.xlsx"); // ※ ファイル名を変更
        File.Copy(template, path, overwrite: true);

        var excel = new ExcelDocument();
        excel.InitializeFile(path, template);

        // == ↓ 確認用コード ↓ == //
        excel.Pos(1, 1).Str = "Hello";
        excel.Pos(2, 1).Value = 123;

        excel.Pos(3, 1, 5, 1).Str = "Range";
        // == ↑ 確認用コード ↑ == //

        excel.FinalizeFile();

        Console.WriteLine("Overwritten: " + path);
    }
}