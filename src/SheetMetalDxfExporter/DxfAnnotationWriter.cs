using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace SheetMetalDxfExporter;

public static class DxfAnnotationWriter
{
    public static void AppendBottomRightLabels(string dxfPath, string partName, string thickness)
    {
        var text = File.ReadAllText(dxfPath, Encoding.ASCII);
        var insertionPointX = 10.0;
        var insertionPointY = 10.0;

        var label1 = BuildTextEntity(insertionPointX, insertionPointY, partName);
        var label2 = BuildTextEntity(insertionPointX, insertionPointY - 5.0, $"material thickness: {thickness}");

        var marker = "0\nENDSEC";
        if (!text.Contains(marker))
        {
            throw new InvalidOperationException("Nie znaleziono sekcji ENDSEC w pliku DXF.");
        }

        text = text.Replace(marker, label1 + label2 + marker);
        File.WriteAllText(dxfPath, text, Encoding.ASCII);
    }

    private static string BuildTextEntity(double x, double y, string value)
    {
        static string fmt(double number) => number.ToString("0.###", CultureInfo.InvariantCulture);
        return string.Join("\n", new[]
        {
            "0",
            "TEXT",
            "8",
            "ANNOTATION",
            "10",
            fmt(x),
            "20",
            fmt(y),
            "30",
            "0.0",
            "40",
            "2.5",
            "1",
            value,
            string.Empty,
        });
    }
}
