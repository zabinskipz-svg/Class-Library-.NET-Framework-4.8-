using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

namespace SolidWorksImageTracerAddin;

[ComVisible(true)]
[Guid("D14D2A24-B8C8-4E38-8B54-6E5DA811AC89")]
[ProgId("SolidWorksImageTracerAddin.Connect")]
[SwAddin(
    Description = "Trace black-and-white PNG images to sketch polylines.",
    Title = "Image Tracer",
    LoadAtStartup = true
)]
public sealed class ImageTracerAddin : ISwAddin
{
    private const string TracePngCallback = nameof(OnTracePng);
    private const string EnableAlways = nameof(IsAlwaysEnabled);

    private SldWorks? _app;
    private int _addinCookie;
    private ICommandManager? _commandManager;

    public bool ConnectToSW(object ThisSW, int cookie)
    {
        _app = (SldWorks)ThisSW;
        _addinCookie = cookie;

        _app.SetAddinCallbackInfo2(0, this, _addinCookie);
        _commandManager = _app.GetCommandManager(_addinCookie);

        AddCommandManager();
        return true;
    }

    public bool DisconnectFromSW()
    {
        RemoveCommandManager();
        Marshal.ReleaseComObject(_commandManager!);
        Marshal.ReleaseComObject(_app!);
        _commandManager = null;
        _app = null;
        return true;
    }

    public int IsAlwaysEnabled() => 1;

    public void OnTracePng()
    {
        if (_app is null)
        {
            return;
        }

        var model = _app.IActiveDoc2;
        if (model is null)
        {
            _app.SendMsgToUser2("Open a part document first.",
                (int)swMessageBoxIcon_e.swMbStop,
                (int)swMessageBoxBtn_e.swMbOk);
            return;
        }

        using var dialog = new OpenFileDialog
        {
            Filter = "PNG files (*.png)|*.png",
            Multiselect = false,
            Title = "Select black-and-white PNG to trace"
        };

        if (dialog.ShowDialog() != DialogResult.OK || !File.Exists(dialog.FileName))
        {
            return;
        }

        try
        {
            var tracer = new ImageTraceService();
            var contourSet = tracer.TraceToPolylines(dialog.FileName);
            tracer.DrawPolylinesIntoSketch(model, contourSet, 0.3);

            _app.SendMsgToUser2($"Traced {contourSet.Count} polylines from {Path.GetFileName(dialog.FileName)}.",
                (int)swMessageBoxIcon_e.swMbInformation,
                (int)swMessageBoxBtn_e.swMbOk);
        }
        catch (Exception ex)
        {
            _app.SendMsgToUser2($"Tracing failed: {ex.Message}",
                (int)swMessageBoxIcon_e.swMbStop,
                (int)swMessageBoxBtn_e.swMbOk);
        }
    }

    [ComRegisterFunction]
    public static void RegisterFunction(Type t)
    {
        AddInRegistration.Register(t);
    }

    [ComUnregisterFunction]
    public static void UnregisterFunction(Type t)
    {
        AddInRegistration.Unregister(t);
    }

    private void AddCommandManager()
    {
        if (_commandManager is null)
        {
            return;
        }

        int errors = 0;
        bool ignorePrevious = false;

        var commandGroup = _commandManager.CreateCommandGroup2(
            CommandIds.MainGroupId,
            "Image Tracer",
            "Trace black-and-white PNG files into sketch geometry",
            "Image Tracer",
            -1,
            ignorePrevious,
            ref errors);

        commandGroup.HasToolbar = true;
        commandGroup.HasMenu = true;

        int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);

        commandGroup.AddCommandItem2(
            "Trace PNG",
            -1,
            "Trace a black-and-white PNG to 2D sketch polylines",
            "Trace PNG",
            CommandIds.TracePngCommandId,
            TracePngCallback,
            EnableAlways,
            menuToolbarOption);

        commandGroup.Activate();

        int[] docTypes =
        {
            (int)swDocumentTypes_e.swDocPART,
            (int)swDocumentTypes_e.swDocASSEMBLY,
            (int)swDocumentTypes_e.swDocDRAWING
        };

        foreach (int docType in docTypes)
        {
            var cmdTab = _commandManager.GetCommandTab(docType, "Image Tracer")
                         ?? _commandManager.AddCommandTab(docType, "Image Tracer");

            if (cmdTab == null)
            {
                continue;
            }

            var box = cmdTab.AddCommandTabBox();
            int[] commands = { commandGroup.get_CommandID(CommandIds.TracePngCommandId) };
            int[] textStyles = { (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal };
            box.AddCommands(commands, textStyles);
        }
    }

    private void RemoveCommandManager()
    {
        _commandManager?.RemoveCommandGroup(CommandIds.MainGroupId);
    }
}
