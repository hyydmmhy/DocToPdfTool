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

**A WPF multi-function document conversion tool** targeting .NET Framework 4.8, with left-side navigation switching between two tools: Doc→PDF and PDF→Word.

### UI Layout

```
MainWindow (800x500)
└─ DockPanel
    ├─ Left sidebar (180px) — NavListBox with 2 items
    │   ├─ "文档转PDF" → Pages.DocToPdfPage
    │   └─ "PDF转Word" → Pages.PdfToWordPage
    ├─ Top-right — "置顶" toggle button
    └─ ContentControl — switches between page UserControls (cached)
```

Pages are created once and cached in a `Dictionary<string, UserControl>`. Switch by `ListBoxItem.Tag`.

### Key Components

- **App.xaml.cs** — Entry point. Single-instance via `Mutex` + `EventWaitHandle`. Defines global styles (AccentBrush #0078D4, NavListBox, NavListBoxItem).

- **Pages/DocToPdfPage.xaml/cs** — Existing Doc→PDF tool (extracted from old MainWindow). WPS/Microsoft Office dual engine, Edge HTML converter. Unchanged logic, only `BtnExit_Click` adapted to `Window.GetWindow(this).Close()`.

- **Pages/PdfToWordPage.xaml/cs** — New PDF→Word tool. File list (add/remove/drag-drop), progress bar, convert button. Runs Word COM on a dedicated STA thread.

- **Utils/PdfConverter.cs** — WPS Office engine (KWps, KET, KWPP COM). Edge headless HTML→PDF.

- **Utils/OfficePdfConverter.cs** — Microsoft Office engine (Word, Excel, PowerPoint COM). Includes window guard for Office popups.

- **Utils/PdfToWordConverter.cs** — Microsoft Word PDF→Word engine. Key features:
  - 3-tier genuine Word instance creation (HKLM CLSID → launch EXE → versioned ProgID)
  - Multi-layer window guard (WMI + WinEventHook + 80ms polling)
  - Registers `DisableConvertPDFWarning` in HKCU
  - Removes Mark-of-the-Web Zone.Identifier before opening

### Conversion Flow: PDF → Word

```
PdfToWordPage.BtnConvert_Click
  └─ Task.Run → STA Thread → DoConversion
       └─ PdfToWordConverter.Initialize()
            ├─ StartWmiProcessWatcher()
            ├─ CreateGenuineWordApplication()  (3-tier fallback)
            │   ├─ HKLM CLSID (pure COM, bypasses WPS hijack)
            │   ├─ Launch WINWORD.EXE /automation → attach via ROT
            │   └─ Versioned ProgID: .16 → .15 → .14 → .12
            ├─ DisablePdfConversionWarningDialog()
            └─ StartWordWindowGuard()  (WinEventHook + polling)
       └─ For each file:
            ├─ UnblockFile() (remove Zone.Identifier)
            └─ ConvertPdfToWordViaWordInterop()
                 ├─ wordApp.Documents.Open(pdfPath)
                 ├─ doc.SaveAs2(docx, 12)
                 └─ doc.Close()

### Important Technical Details

- **Late-bound COM** — all Office/WPS interop uses `Type.GetTypeFromProgID` + `Activator.CreateInstance` with `dynamic`. No primary interop assemblies (PIAs) needed.
- **Window guard** (`OfficePdfConverter.RunWithGuard`) — critical for unattended Office conversions. Uses `WinEventHook(EVENT_OBJECT_CREATE)` + a 5ms polling thread to hide Office windows immediately on creation. Hides via `ShowWindow(SW_HIDE)`, removes `WS_VISIBLE`/`WS_EX_APPWINDOW` styles, sets `WDA_MONITOR` display affinity.
- **PowerPoint on STA thread** — `OfficePdfConverter.ConvertPpt` runs on a dedicated STA thread with 120s timeout; errors are marshaled back via `threadError`.
- **Edge HTML conversion** — injects `@page { size: 1920px 1080px }` CSS, replaces `position:fixed/sticky` with `static`, then launches `msedge.exe --headless=new` with 60s timeout. Waits up to 6s for PDF file to stabilize after process exit.
- **Single-file publish** — conditionally includes `Costura.Fody` only when `SingleFile=true` MSBuild property is set.

## Constraints

- **Windows-only** — targets `net48` (Windows-only), uses Win32 P/Invoke extensively.
- **Requires installed Office** — either WPS Office or Microsoft Office must be installed on the target machine.
- **x64 msedge.exe** — Edge is found via registry (`SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe`) or well-known install paths.