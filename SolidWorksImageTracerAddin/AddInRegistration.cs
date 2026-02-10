using System;
using Microsoft.Win32;

namespace SolidWorksImageTracerAddin;

internal static class AddInRegistration
{
    private const string AddinsRoot = @"SOFTWARE\SolidWorks\Addins";
    private const string StartupRoot = @"SOFTWARE\SolidWorks\AddInsStartup";
    private const string AddInTitle = "Image Tracer";
    private const string AddInDescription = "Trace black-and-white PNG images into 2D sketch polylines.";

    public static void Register(Type type)
    {
        string clsid = type.GUID.ToString("B");

        using (RegistryKey? addinKey = Registry.LocalMachine.CreateSubKey($"{AddinsRoot}\\{clsid}"))
        {
            addinKey?.SetValue(null, 1, RegistryValueKind.DWord);
            addinKey?.SetValue("Title", AddInTitle, RegistryValueKind.String);
            addinKey?.SetValue("Description", AddInDescription, RegistryValueKind.String);
        }

        using (RegistryKey? startupKey = Registry.CurrentUser.CreateSubKey($"{StartupRoot}\\{clsid}"))
        {
            startupKey?.SetValue(null, 1, RegistryValueKind.DWord);
        }
    }

    public static void Unregister(Type type)
    {
        string clsid = type.GUID.ToString("B");
        Registry.LocalMachine.DeleteSubKeyTree($"{AddinsRoot}\\{clsid}", false);
        Registry.CurrentUser.DeleteSubKeyTree($"{StartupRoot}\\{clsid}", false);
    }
}
