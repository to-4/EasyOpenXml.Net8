namespace EasyOpenXml.Excel.DemoConsole;

internal static class DemoRunner
{
    public static void Run(string[] args)
    {
        var selection = args.FirstOrDefault(); // "1", "all", null

        Console.WriteLine("EasyOpenXml.Excel DemoConsole");
        Console.WriteLine("Output folder: " + Paths.OutputDir);
        Console.WriteLine();

        var demos = new Dictionary<string, Action>
        {
            //{ "1", Demos.Demo01_InitializeFile.Run },
            //{ "2", Demos.Demo02_SheetSelect.Run },
            //{ "3", Demos.Demo03_GetValue.Run },
            //{ "4", Demos.Demo04_GetCellValue.Run },
            //{ "5", Demos.Demo05_Pos.Run },
            //{ "6", Demos.Demo06_PosAttr.Run },
            //{ "7", Demos.Demo07_PosCopyPaste.Run },
            //{ "8", Demos.Demo08_PosMerge.Run },
            //{ "9", Demos.Demo09_PrintArea.Run },
            //{ "10", Demos.Demo10_RowDelete.Run },
            //{ "11", Demos.Demo11_SetCalcMode.Run },
            //{ "12", Demos.Demo12_Logger.Run },
            { "13", Demos.Demo13_ExportSharedFormulas.Run },
        };

        // 引数なし or all → 全実行
        if (string.IsNullOrWhiteSpace(selection)
            || selection.Equals("all", StringComparison.OrdinalIgnoreCase))
        {
            RunAll(demos);
            return;
        }

        // 指定実行
        if (demos.TryGetValue(selection, out var demo))
        {
            RunSingle(selection, demo);
            return;
        }

        // 不正引数
        ShowUsage();
    }

    private static void RunAll(Dictionary<string, Action> demos)
    {
        foreach (var (key, demo) in demos)
        {
            Console.WriteLine($"--- Demo {key} start ---");
            demo();
            Console.WriteLine($"--- Demo {key} end ---");
            Console.WriteLine();
        }
    }

    private static void RunSingle(string key, Action demo)
    {
        Console.WriteLine($"--- Demo {key} start ---");
        demo();
        Console.WriteLine($"--- Demo {key} end ---");
        Console.WriteLine();
    }

    private static void ShowUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  dotnet run --project src/EasyOpenXml.Excel.DemoConsole");
        Console.WriteLine("  dotnet run --project src/EasyOpenXml.Excel.DemoConsole -- all");
        Console.WriteLine("  dotnet run --project src/EasyOpenXml.Excel.DemoConsole -- 1");
        Console.WriteLine("  dotnet run --project src/EasyOpenXml.Excel.DemoConsole -- 2");
    }
}

internal static class Paths
{
    public static string OutputDir { get; } =
        Path.Combine(AppContext.BaseDirectory, "Output");

    public static string OutFile(string fileName)
    {
        Directory.CreateDirectory(OutputDir);
        return Path.Combine(OutputDir, fileName);
    }
}