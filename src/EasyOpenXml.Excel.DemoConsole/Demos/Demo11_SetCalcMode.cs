using EasyOpenXml.Excel;
using EasyOpenXml.Excel.Models;
using System.Drawing;

namespace EasyOpenXml.Excel.DemoConsole.Demos;

internal static class Demo11_SetCalcMode
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

        var path = Paths.OutFile("Demo10_RowDelete.xlsx"); // ※ ファイル名を変更
        File.Copy(template, path, overwrite: true);

        var excel = new ExcelDocument();
        excel.InitializeFile(path, template);

        // == ↓ 確認用コード ↓ == //
        // 1. 一時的に手動モード（大量書き込み対策）
        excel.SetCalculationMode(CalculationMode.Manual);

        // データ大量投入
        excel.Pos(1, 1).Value = 100;
        excel.Pos(1, 2).Value = 200;

        // 2. 自動に戻す
        excel.SetCalculationMode(CalculationMode.Automatic);
        // == ↑ 確認用コード ↑ == //

        excel.FinalizeFile();

        Console.WriteLine("Overwritten: " + path);
    }
}