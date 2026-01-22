using EasyOpenXml.Excel;

namespace EasyOpenXml.Excel.DemoConsole.Demos;

internal static class Demo02_SheetSelect
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

        var path = Paths.OutFile("demo02_SheetSelect.xlsx"); // ※ ファイル名を変更
        File.Copy(template, path, overwrite: true);

        var excel = new ExcelDocument();
        excel.InitializeFile(path, template);

        // == ↓ 確認用コード ↓ == //

        excel.SheetSelect(1); // Select Sheet2
        excel.SetValue(1, 1, "Hello");
        excel.SetValue(1, 3, DateTime.Now); // 日付（OAデータをセット、書式を変えれば日付や時刻を表示できる）

        // == ↑ 確認用コード ↑ == //

        excel.FinalizeFile();

        Console.WriteLine("Overwritten: " + path);
    }
}