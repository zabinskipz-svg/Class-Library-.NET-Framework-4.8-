# SheetMetal DXF Exporter (SolidWorks + Windows)

Ten projekt dostarcza prosty dodatek narzędziowy, który automatyzuje eksport części blachowych do DXF:

- eksport rozłożenia blachy do **DXF 1:1**,
- zachowanie **linii gięcia**,
- dopisanie w prawym dolnym rogu:
  - nazwy części,
  - `material thickness: <wartość z równania "Grubość">`.

## Co zawiera repo

- `src/SheetMetalDxfExporter` – biblioteka z logiką automatyzacji COM SolidWorks.
- `src/SheetMetalDxfExporter.Cli` – aplikacja konsolowa uruchamiana z Windows (np. z menu kontekstowego).
- `tools/install-context-menu.reg` – przykładowy wpis do menu kontekstowego Windows dla plików `.sldprt`.

## Budowanie (Windows + .NET Framework 4.8 + SolidWorks)

```powershell
msbuild .\SheetMetalDxfExporter.sln /p:Configuration=Release
```

## Użycie ręczne

```powershell
.\src\SheetMetalDxfExporter.Cli\bin\Release\SheetMetalDxfExporter.Cli.exe "C:\modele\detal.SLDPRT"
```

albo z własną ścieżką wyjściową:

```powershell
.\src\SheetMetalDxfExporter.Cli\bin\Release\SheetMetalDxfExporter.Cli.exe "C:\modele\detal.SLDPRT" "C:\out\detal.dxf"
```

## Integracja z poziomu Windows (Explorer)

1. Otwórz `tools/install-context-menu.reg`.
2. Zmień ścieżkę `C:\Tools\SheetMetalDxfExporter.Cli.exe` na docelową lokalizację EXE.
3. Uruchom plik `.reg` jako administrator.
4. Kliknij PPM na `.SLDPRT` -> **Export Flat Pattern DXF (laser)**.

## Integracja z poziomu SolidWorks

Najprościej: przypisz skrót klawiaturowy lub przycisk, który uruchamia CLI z aktywną częścią.
Jeżeli chcesz pełny Add-In z ikoną w CommandManager, ten kod biblioteki można wykorzystać jako backend eksportu.

## Uwaga dot. równania grubości

Program szuka wpisu zawierającego nazwę `Grubość` w Equations i pobiera wartość po znaku `=`.
Przykład oczekiwanej definicji:

```text
"Grubość" = 2.0mm
```

