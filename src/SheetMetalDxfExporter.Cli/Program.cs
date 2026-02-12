using System;
using System.IO;
using SheetMetalDxfExporter;

internal static class Program
{
    private static int Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.Error.WriteLine("Użycie: SheetMetalDxfExporter.Cli <plik.SLDPRT> [output.dxf]");
            return 1;
        }

        var input = Path.GetFullPath(args[0]);
        var output = args.Length > 1
            ? Path.GetFullPath(args[1])
            : Path.ChangeExtension(input, ".flat.dxf")!;

        try
        {
            var service = new SolidWorksAutomationService();
            service.ExportFlatPatternDxf(new ExportOptions
            {
                InputPartPath = input,
                OutputDxfPath = output,
            });

            Console.WriteLine($"Zapisano DXF: {output}");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 2;
        }
    }
}
