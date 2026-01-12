using EasyOpenXml.Excel;

namespace EasyOpenXml.Excel.DemoConsole.Demos;

internal static class Demo04_GetCellValue
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

        var path = Paths.OutFile("demo04_GetCellValue.xlsx"); // ※ ファイル名を変更
        File.Copy(template, path, overwrite: true);

        var excel = new ExcelDocument();
        excel.InitializeFile(path, template);

        // == ↓ 確認用コード ↓ == //
        var a1 = excel.GetValue("A1").ToString();
        var a2 = excel.GetValue("A2").ToString();
        excel.SetValue("A4", $"A1: {a1}, A2: {a2}");

        // == ↑ 確認用コード ↑ == //

        excel.FinalizeFile();

        Console.WriteLine("Overwritten: " + path);
    }
}