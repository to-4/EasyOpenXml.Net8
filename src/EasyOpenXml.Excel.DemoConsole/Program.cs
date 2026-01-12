using EasyOpenXml.Excel.DemoConsole;

try
{

    //Console.WriteLine("Hello, World!");
    DemoRunner.Run(args);
    //DemoRunner.Run(new string[] { "all" });
    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine("=== ERROR ===");
    Console.Error.WriteLine(ex);
    return 1;
}

