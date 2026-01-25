using EasyOpenXml.Excel;
using EasyOpenXml.Excel.Models;
using System.Drawing;
using System.Diagnostics;

namespace EasyOpenXml.Excel.DemoConsole.Demos;

internal static class Demo12_Logger
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

        var path = Paths.OutFile("Demo12_Logger.xlsx"); // ※ ファイル名を変更
        File.Copy(template, path, overwrite: true);

        var excel = new ExcelDocument();
        excel.InitializeFile(path, template);

        // == ↓ 確認用コード ↓ == //

        // ログファイル ※ logFile = @"C:\Users\[~(端末固有)]\デスクトップ\OxLog.log"; でもOK
        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        var logFile = System.IO.Path.Combine(desktop, "OxLog.log");

        // ログオブジェクト生成
        var logger = new Logger(logFile);

        logger.Start(""); // 開始ログ

        Thread.Sleep(1000);

        logger.Log("process 1 done."); // 途中ログ出力、引数はメッセージ

        Thread.Sleep(1000);

        logger.Log("process 2 done."); // 途中ログ出力、引数はメッセージ

        logger.End(); // 終了ログ

        // == ↑ 確認用コード ↑ == //

        excel.FinalizeFile();

        Console.WriteLine("Overwritten: " + path);
    }
}