using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace SheetMetalDxfExporter;

public sealed class SolidWorksAutomationService
{
    private const int SwDocPart = 1;
    private const int SwOpenDocOptionsSilent = 1;
    private const int ExportSheetMetalAction = 2;
    private const int SwSaveAsCurrentVersion = 0;
    private const int SwSaveAsOptionsSilent = 1;

    // Geometry + bend lines + forming tools + hidden edges + merge coplanar faces.
    private const int FlatPatternOptions = 1 + 4 + 8 + 16 + 64;

    public void ExportFlatPatternDxf(ExportOptions options)
    {
        options.Validate();

        var swType = Type.GetTypeFromProgID("SldWorks.Application")
            ?? throw new InvalidOperationException("Nie znaleziono zainstalowanego SolidWorks (ProgID: SldWorks.Application).");

        dynamic? swApp = null;
        dynamic? model = null;

        try
        {
            swApp = Activator.CreateInstance(swType) ?? throw new InvalidOperationException("Nie można uruchomić instancji SolidWorks.");
            swApp.Visible = false;

            int errors = 0;
            int warnings = 0;
            model = swApp.OpenDoc6(options.InputPartPath, SwDocPart, SwOpenDocOptionsSilent, string.Empty, ref errors, ref warnings);
            if (model == null)
            {
                throw new InvalidOperationException($"SolidWorks nie otworzył części. Errors={errors}, Warnings={warnings}");
            }

            var exported = TryExportViaExportToDwg2(model, options.OutputDxfPath, options.InputPartPath)
                || TryExportViaSaveAs(model, options.OutputDxfPath);
            if (!exported)
            {
                throw new InvalidOperationException("Eksport DXF nie powiódł się. Sprawdź, czy plik jest częścią blachową z rozłożeniem i czy można wykonać 'Zapisz jako -> DXF' ręcznie.");
            }

            var fileName = Path.GetFileNameWithoutExtension(options.InputPartPath);
            var thickness = ReadThicknessFromEquation(model) ?? "n/a";
            DxfAnnotationWriter.AppendBottomRightLabels(options.OutputDxfPath, fileName, thickness);
        }
        finally
        {
            if (model != null)
            {
                var title = (string)model.GetTitle();
                swApp?.CloseDoc(title);
                Marshal.FinalReleaseComObject(model);
            }

            if (swApp != null)
            {
                swApp.ExitApp();
                Marshal.FinalReleaseComObject(swApp);
            }
        }
    }

    private static bool TryExportViaExportToDwg2(dynamic model, string outputDxfPath, string inputPartPath)
    {
        var alignment = new double[12];
        var views = Array.Empty<string>();

        if (File.Exists(outputDxfPath))
        {
            File.Delete(outputDxfPath);
        }

        try
        {
            var result = model.ExportToDWG2(
                outputDxfPath,
                inputPartPath,
                ExportSheetMetalAction,
                true,
                alignment,
                false,
                false,
                FlatPatternOptions,
                views);

            return result is bool ok && ok && File.Exists(outputDxfPath);
        }
        catch
        {
            return false;
        }
    }

    private static bool TryExportViaSaveAs(dynamic model, string outputDxfPath)
    {
        if (File.Exists(outputDxfPath))
        {
            File.Delete(outputDxfPath);
        }

        try
        {
            int errors = 0;
            int warnings = 0;

            var result = model.Extension.SaveAs(
                outputDxfPath,
                SwSaveAsCurrentVersion,
                SwSaveAsOptionsSilent,
                null,
                ref errors,
                ref warnings);

            return result is bool ok && ok && File.Exists(outputDxfPath);
        }
        catch
        {
            return false;
        }
    }

    private static string? ReadThicknessFromEquation(dynamic model)
    {
        try
        {
            dynamic equationMgr = model.GetEquationMgr();
            int count = (int)equationMgr.GetCount();

            for (var i = 0; i < count; i++)
            {
                string equationLine = equationMgr.Equation[i];
                if (!equationLine.Contains("Grubość", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var rightSide = equationLine.Split('=').Skip(1).FirstOrDefault()?.Trim();
                return string.IsNullOrWhiteSpace(rightSide) ? null : rightSide.Trim('"');
            }
        }
        catch
        {
            return null;
        }

        return null;
    }
}
