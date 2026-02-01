using EasyOpenXml.Excel;
using EasyOpenXml.Excel.Models;
using System.Drawing;
using System.Diagnostics;

namespace EasyOpenXml.Excel.DemoConsole.Demos;

internal static class Demo14_BackColor
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

        var path = Paths.OutFile("Demo14_BackColor.xlsx"); // ※ ファイル名を変更
        File.Copy(template, path, overwrite: true);

        var excel = new ExcelDocument();
        excel.InitializeFile(path, template);

        // == ↓ 確認用コード ↓ == //

        // 背景色を変更
        var color = Color.Gray;
        //excel.Pos(4, 12, 8, 12).Attr.BackColor = color;
        excel.Pos(4, 12, 8, 12).Attr.BackColor = color;
        // == ↑ 確認用コード ↑ == //

        excel.FinalizeFile();

        Console.WriteLine("Overwritten: " + path);
    }
}