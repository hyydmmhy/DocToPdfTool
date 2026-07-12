# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Publish

```powershell
# Debug build
dotnet build

# Release build
dotnet build -c Release

# Single-file publish (merges all dependencies into one exe via Costura.Fody)
.\publish_single.bat
# Or: dotnet publish DocToPdfTool.csproj -c Release -o publish\SingleFile -p:SingleFile=true
# Output: publish\SingleFile\DocToPdfTool.exe
```

## Project Architecture

**A WPF multi-function document conversion tool** targeting .NET Framework 4.8, with left-side navigation switching between three tools: Doc→PDF, PDF→Word, and PDF→Image.

### UI Layout

```
MainWindow (800x500)
└─ DockPanel
    ├─ Left sidebar (180px) — NavListBox with 3 items
    │   ├─ "文档转PDF" → Pages.DocToPdfPage
    │   ├─ "PDF转Word" → Pages.PdfToWordPage
    │   └─ "PDF转图片" → Pages.PdfToImagePage
    ├─ Top-right — "置顶" toggle button
    └─ ContentControl — switches between page UserControls (cached)
```

Pages are created once and cached in a `Dictionary<string, UserControl>`. Switch by `ListBoxItem.Tag`.

### Key Components

- **App.xaml.cs** — Entry point. Single-instance via `Mutex` + `EventWaitHandle`. MainWindow.Closed forces `Process.Kill()` to handle WinRT runtime thread residuals. Defines global styles (AccentBrush #0078D4, NavListBox, NavListBoxItem).

- **Pages/DocToPdfPage.xaml/cs** — Doc→PDF tool. WPS/Microsoft Office dual engine, Edge HTML converter. Unchanged logic, only `BtnExit_Click` adapted to `Window.GetWindow(this).Close()`.

- **Pages/PdfToWordPage.xaml/cs** — PDF→Word tool. File list (add/remove/drag-drop), progress bar, convert button. Runs Word COM on a dedicated STA thread.

- **Pages/PdfToImagePage.xaml/cs** — PDF→Image tool. Batch file list with drag-drop, settings (scale/format/DPI), convert button. Uses `Task.Run` with `PdfToImageConverter`.

- **Utils/PdfConverter.cs** — WPS Office engine (KWps, KET, KWPP COM). Edge headless HTML→PDF.

- **Utils/OfficePdfConverter.cs** — Microsoft Office engine (Word, Excel, PowerPoint COM). Includes window guard for Office popups.

- **Utils/PdfToWordConverter.cs** — Microsoft Word PDF→Word engine. Key features:
  - 3-tier genuine Word instance creation (HKLM CLSID → launch EXE → versioned ProgID)
  - Multi-layer window guard (WMI + WinEventHook + 80ms polling)
  - Registers `DisableConvertPDFWarning` in HKCU
  - Removes Mark-of-the-Web Zone.Identifier before opening

- **Utils/PdfToImageConverter.cs** — PDF→Image engine based on Windows.Data.Pdf (Windows 10+ built-in PDF renderer). Key features:
  - Zero external dependencies (no pdfium.dll, Ghostscript, or Office required)
  - Supports DPI (150/300/600), scale (1×~5×), output format (JPEG/PNG)
  - Uses `PdfPageRenderOptions.DestinationWidth/Height` for high-DPI rendering
  - Proper WinRT RCW cleanup via `Marshal.FinalReleaseComObject`
  - 250M pixel GDI+ limit protection via `ClampSize`

### Conversion Flow: PDF → Image

```
PdfToImagePage.BtnConvert_Click
  └─ await Task.Run
       └─ PdfToImageConverter.Convert (sync wrapper → ConvertAsync)
            ├─ StorageFile.GetFileFromPathAsync(pdfPath)
            ├─ PdfDocument.LoadFromFileAsync(file)
            └─ For each page:
                 ├─ pdfDoc.GetPage(index)
                 ├─ page.RenderToStreamAsync(stream, options)
                 │    └─ PdfPageRenderOptions { DestinationWidth, DestinationHeight }
                 ├─ new Bitmap(stream.AsStream())
                 ├─ image.SetResolution(Dpi, Dpi)
                 └─ SaveWithQuality(image, outputFile)
                  ├─ JPEG → Encoder.Quality 95%
                  └─ PNG → direct Save
```

### Important Technical Details

- **Late-bound COM** — all Office/WPS interop uses `Type.GetTypeFromProgID` + `Activator.CreateInstance` with `dynamic`. No primary interop assemblies (PIAs) needed.
- **Window guard** (`OfficePdfConverter.RunWithGuard`) — critical for unattended Office conversions. Uses `WinEventHook(EVENT_OBJECT_CREATE)` + a 5ms polling thread to hide Office windows immediately on creation. Hides via `ShowWindow(SW_HIDE)`, removes `WS_VISIBLE`/`WS_EX_APPWINDOW` styles, sets `WDA_MONITOR` display affinity.
- **PowerPoint on STA thread** — `OfficePdfConverter.ConvertPpt` runs on a dedicated STA thread with 120s timeout; errors are marshaled back via `threadError`.
- **Edge HTML conversion** — injects `@page { size: 1920px 1080px }` CSS, replaces `position:fixed/sticky` with `static`, then launches `msedge.exe --headless=new` with 60s timeout. Waits up to 6s for PDF file to stabilize after process exit.
- **WinRT RCW cleanup** — `PdfToImageConverter` explicitly releases all WinRT objects (`PdfDocument`, `PdfPage`, `PdfPageRenderOptions`, `StorageFile`) via `Marshal.FinalReleaseComObject` in `finally` blocks to prevent memory leaks and process hang.
- **ReleaseMemory()** — `PdfToImagePage` calls a comprehensive cleanup after conversion: `GCSettings.LargeObjectHeapCompactionMode = CompactOnce`, double `GC.Collect()`, and `EmptyWorkingSet()` to return freed pages to the OS.
- **Process exit guarantee** — `App.xaml.cs` hooks `MainWindow.Closed` to call `Process.Kill()` as a safety net, ensuring the process terminates even if WinRT runtime threads linger.
- **Single-file publish** — conditionally includes `Costura.Fody` only when `SingleFile=true` MSBuild property is set.

## Constraints

- **Windows-only** — targets `net48` (Windows-only), uses Win32 P/Invoke extensively.
- **PDF to Image requires Windows 10+** — `Windows.Data.Pdf` API is only available on Windows 10/11.
- **Doc to PDF requires installed Office** — either WPS Office or Microsoft Office must be installed on the target machine.
- **PDF to Word requires Microsoft Office** — Word COM automation is required; WPS Office is not supported.
- **PDF to Image needs no extra software** — uses Windows built-in PDF renderer, zero external dependencies.
- **x64 msedge.exe** — Edge is found via registry (`SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe`) or well-known install paths.