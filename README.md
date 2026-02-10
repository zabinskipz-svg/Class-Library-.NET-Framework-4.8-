# SolidWorks 2025 Image Trace COM Add-In (.NET Framework 4.8)

This repository contains a SolidWorks COM add-in written in C# targeting **.NET Framework 4.8**.

## Features

- Adds an **Image Tracer** command group to the SolidWorks CommandManager.
- Exposes a **Trace PNG** button on toolbar/menu/command tabs.
- Opens a black-and-white PNG and extracts outer pixel boundaries into polyline loops.
- Draws the resulting geometry as 2D sketch lines in the active document sketch.
- Automatically scales traced geometry to **300 mm width** (`0.3` meters in SolidWorks units).

## Project layout

- `SolidWorksImageTracerAddin/SolidWorksImageTracerAddin.csproj` - .NET Framework class library with COM interop enabled.
- `SolidWorksImageTracerAddin/ImageTracerAddin.cs` - `ISwAddin` implementation and CommandManager integration.
- `SolidWorksImageTracerAddin/AddInRegistration.cs` - COM registration hooks and SolidWorks add-in registry keys.
- `SolidWorksImageTracerAddin/ImageTraceService.cs` - PNG thresholding, contour extraction, and sketch geometry generation.

## Prerequisites

1. **SolidWorks 2025** installed.
2. **Visual Studio 2022** (or MSBuild with .NET Framework 4.8 targeting pack).
3. SolidWorks interop assemblies available in GAC (normally installed with SolidWorks):
   - `SolidWorks.Interop.sldworks`
   - `SolidWorks.Interop.swconst`
   - `SolidWorks.Interop.swpublished`
4. Administrator rights for COM registration steps.

## Build steps

### Visual Studio

1. Open the project folder in Visual Studio.
2. Build in `Release | Any CPU`.

### Command line (Developer Command Prompt)

```bat
msbuild SolidWorksImageTracerAddin\SolidWorksImageTracerAddin.csproj /p:Configuration=Release
```

Output DLL path:

```text
SolidWorksImageTracerAddin\bin\Release\net48\SolidWorksImageTracerAddin.dll
```

## Register the COM add-in

The assembly uses `[ComRegisterFunction]` and `[ComUnregisterFunction]` to create/remove required SolidWorks keys:

- `HKLM\SOFTWARE\SolidWorks\Addins\{CLSID}`
- `HKCU\SOFTWARE\SolidWorks\AddInsStartup\{CLSID}`

### Register

Run **as Administrator**:

```bat
"%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" /codebase "<full-path>\SolidWorksImageTracerAddin.dll"
```

### Unregister

```bat
"%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" /unregister "<full-path>\SolidWorksImageTracerAddin.dll"
```

## Usage

1. Start SolidWorks.
2. Confirm **Image Tracer** is loaded (Tools -> Add-Ins).
3. Open/create a part and activate a planar sketch (or let the add-in create one).
4. Click **Image Tracer > Trace PNG**.
5. Select a black-and-white PNG.
6. The add-in inserts sketch line segments that approximate image boundaries, scaled to **300 mm total width**.

## Notes

- Best results are with clean black-on-white bitmaps and moderate resolution.
- The trace is pixel-edge based and intentionally favors robust boundary loops over smoothing.
