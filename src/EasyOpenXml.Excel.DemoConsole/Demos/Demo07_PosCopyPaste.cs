using EasyOpenXml.Excel;
using EasyOpenXml.Excel.Models;
using System.Drawing;

namespace EasyOpenXml.Excel.DemoConsole.Demos;

internal static class Demo07_PosCopyPaste
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

        var path = Paths.OutFile("Demo07_PosCopyPaste.xlsx"); // ※ ファイル名を変更
        File.Copy(template, path, overwrite: true);

        var excel = new ExcelDocument();
        excel.InitializeFile(path, template);

        // == ↓ 確認用コード ↓ == //
        excel.Pos(1, 1).Str = "セルコピー元";
        excel.Pos(1, 1).Attr.BackColor = Color.LightYellow;
        excel.Pos(1, 1).Copy();

        // Paste destination
        excel.Pos(1, 3).Paste();
        // == ↑ 確認用コード ↑ == //

        excel.FinalizeFile();

        Console.WriteLine("Overwritten: " + path);
    }
}