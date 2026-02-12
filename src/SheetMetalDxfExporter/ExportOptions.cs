using System;
using System.IO;

namespace SheetMetalDxfExporter;

public sealed class ExportOptions
{
    public required string InputPartPath { get; init; }
    public required string OutputDxfPath { get; init; }

    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(InputPartPath) || !File.Exists(InputPartPath))
        {
            throw new FileNotFoundException("Nie znaleziono pliku części SolidWorks.", InputPartPath);
        }

        if (string.IsNullOrWhiteSpace(OutputDxfPath))
        {
            throw new ArgumentException("Ścieżka wyjściowa DXF nie może być pusta.", nameof(OutputDxfPath));
        }

        var folder = Path.GetDirectoryName(OutputDxfPath);
        if (!string.IsNullOrWhiteSpace(folder) && !Directory.Exists(folder))
        {
            Directory.CreateDirectory(folder);
        }
    }
}
