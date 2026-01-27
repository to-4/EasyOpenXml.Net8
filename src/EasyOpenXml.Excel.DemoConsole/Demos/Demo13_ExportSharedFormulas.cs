using EasyOpenXml.Excel;
using EasyOpenXml.Excel.Models;
using System.Drawing;
using System.Diagnostics;

namespace EasyOpenXml.Excel.DemoConsole.Demos;

internal static class Demo13_ExportSharedFormulas
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

        var path = Paths.OutFile("Demo13_ExportSharedFormulas.xlsx"); // ※ ファイル名を変更
        File.Copy(template, path, overwrite: true);

        var excel = new ExcelDocument();
        excel.InitializeFile(path, template);

        // == ↓ 確認用コード ↓ == //

        // ログファイル ※ logFile = @"C:\Users\[~(端末固有)]\デスクトップ\OxLog.log"; でもOK
        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        var csvFile = System.IO.Path.Combine(desktop, "sharedFormula.csv");

        excel.ExportSharedFormulasCsv(csvFile);

        // == ↑ 確認用コード ↑ == //

        excel.FinalizeFile();

        Console.WriteLine("Overwritten: " + path);
    }
}