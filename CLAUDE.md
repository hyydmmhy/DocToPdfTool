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

**A WPF multi-function document conversion tool** targeting .NET Framework 4.8, with left-side navigation switching between six tools: Doc→PDF, PDF→Word, PDF→Excel, PDF→PPT, PDF→Image, and PDF→ScannedPdf.

### UI Layout

```
MainWindow (800x500)
└─ DockPanel
    ├─ Left sidebar (180px) — NavListBox with 6 items
    │   ├─ "文档转PDF" → Pages.DocToPdfPage
    │   ├─ "PDF转Word" → Pages.PdfToWordPage
    │   ├─ "PDF转Excel" → Pages.PdfToExcelPage
    │   ├─ "PDF转PPT" → Pages.PdfToPptPage
    │   ├─ "PDF转图片" → Pages.PdfToImagePage
    │   └─ "PDF转扫描件" → Pages.PdfToScannedPdfPage
    ├─ Top-right — "置顶" toggle button
    └─ ContentControl — switches between page UserControls (cached)
```

Pages are created once and cached in a `Dictionary<string, UserControl>`. Switch by `ListBoxItem.Tag`.

### Key Components

- **App.xaml.cs** — Entry point. Single-instance via `Mutex` + `EventWaitHandle`. MainWindow.Closed forces `Process.Kill()` to handle WinRT runtime thread residuals. Defines global styles (AccentBrush #0078D4, NavListBox, NavListBoxItem).

- **Pages/DocToPdfPage.xaml/cs** — Doc→PDF tool. WPS/Microsoft Office dual engine, Edge HTML converter. Unchanged logic, only `BtnExit_Click` adapted to `Window.GetWindow(this).Close()`.

- **Pages/PdfToWordPage.xaml/cs** — PDF→Word tool. File list (add/remove/drag-drop), progress bar, convert button. Runs Word COM on a dedicated STA thread. Calls `ReleaseMemory()` after conversion (same pattern as other pages).

- **Pages/PdfToExcelPage.xaml/cs** — PDF→Excel tool. File list with add/remove/drag-drop, convert button. Uses `PdfToExcelConverter` on a background thread.

- **Pages/PdfToImagePage.xaml/cs** — PDF→Image tool. Batch file list with drag-drop, settings (scale/format/DPI), convert button. Uses `Task.Run` with `PdfToImageConverter`.

- **Pages/PdfToPptPage.xaml/cs** — PDF→PPT tool. Batch file list with drag-drop, convert button. Uses `Task.Run` with `PdfToPptConverter`.

- **Pages/PdfToScannedPdfPage.xaml/cs** — PDF→ScannedPdf tool. Batch file list with drag-drop, convert button. Uses `Task.Run` with `PdfToScannedPdfConverter`.

- **Utils/PdfConverter.cs** — WPS Office engine (KWps, KET, KWPP COM). Edge headless HTML→PDF.

- **Utils/OfficePdfConverter.cs** — Microsoft Office engine (Word, Excel, PowerPoint COM). Includes window guard for Office popups.

- **Utils/PdfToWordConverter.cs** — Microsoft Word PDF→Word engine. Key features:
  - 3-tier genuine Word instance creation (HKLM CLSID → launch EXE → versioned ProgID)
  - Multi-layer window guard (WMI + WinEventHook + 80ms polling)
  - Registers `DisableConvertPDFWarning` in HKCU
  - Removes Mark-of-the-Web Zone.Identifier before opening

- **Utils/PdfToWordConverterSpire.cs** — Spire.PDF PDF→Word engine (no Office needed). Key features:
  - Uses local `Libs\Spire.Pdf.dll` reference (cracked, no page limit or watermark)
  - Large PDF batch processing: splits into 20-page batches, converts each batch, merges DOCX via OpenXML AltChunk
  - Memory-controlled: `ReleaseMemory()` after each batch (`GC.Collect` + `EmptyWorkingSet`)
  - Page count via PdfPig (lightweight, no memory spike from initial load)

- **Utils/PdfToImageConverter.cs** — PDF→Image engine based on Windows.Data.Pdf (Windows 10+ built-in PDF renderer). Key features:
  - Zero external dependencies (no pdfium.dll, Ghostscript, or Office required)
  - Supports DPI (150/300/600), scale (1×~5×), output format (JPEG/PNG)
  - Uses `PdfPageRenderOptions.DestinationWidth/Height` for high-DPI rendering
  - Proper WinRT RCW cleanup via `Marshal.FinalReleaseComObject`
  - 250M pixel GDI+ limit protection via `ClampSize`

- **Utils/PdfToExcelConverter.cs** — PDF→Excel engine based on PdfPig. Key features:
  - Table auto-detection: column boundary detection via left-edge alignment clustering
  - Split-row merging: recombines text-wrapped rows using Y-gap + X-overlap heuristics
  - Header/footer margin filtering to exclude page noise
  - Non-table fallback: outputs text lines as single-column rows when no table structure is detected
  - Pure XML xlsx writer: generates `.xlsx` via `System.IO.Compression.ZipArchive` + `XElement`, zero Office dependency

- **Utils/PdfToPptConverter.cs** — PDF→PPT engine based on Windows.Data.Pdf (Windows 10+ built-in PDF renderer). Key features:
  - Renders each PDF page as a high-DPI image and embeds it into PPTX slides
  - Generates valid OOXML with `p:defaultTextStyle`, `p:txStyles`, `p:clrMapOvr` for PowerPoint compatibility
  - No Office or WPS dependency — pure XML pptx generation via `ZipArchive` + `XElement`
  - Shares `PdfToImageConverter.ConvertAsync` with per-page `onPageComplete` callback for memory cleanup
  - DPI (300), scale (1.5×) — fixed high-quality settings

- **Utils/PdfToScannedPdfConverter.cs** — PDF→ScannedPdf engine based on Windows.Data.Pdf (Windows 10+ built-in PDF renderer). Key features:
  - Renders each PDF page as a JPEG image and embeds it into a new PDF via raw PDF writer
  - Pure .NET PDF writer: writes valid PDF 1.4 with DCTDecode (JPEG) image streams
  - Zero external dependencies — no iTextSharp, PDFsharp, or any PDF library needed
  - Generates Catalog, Pages, Page, Image XObject, and Content Stream objects with xref table
  - DPI (200), JPEG quality (85) — fixed scanned-document settings
  - Proper WinRT RCW cleanup via `Marshal.FinalReleaseComObject`

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

### Conversion Flow: PDF → Excel

```
PdfToExcelPage.BtnConvert_Click
  └─ await Task.Run
       └─ PdfToExcelConverter.Convert(pdfPath, outputDir)
            ├─ Convert.ExtractAllPages(pdfPath)
            │    └─ PdfDocument.Open → foreach page:
            │         ├─ page.GetWords()
            │         └─ ProcessPage(words, pageHeight)
            │              ├─ FilterMargins → exclude headers/footers
            │              ├─ ClusterRows → group words into visual lines
            │              ├─ IsTablePage? → column detection or text fallback
            │              ├─ If table: MergeSplitRows → DetectColumnsByAlignment → BuildTable
            │              └─ If not table: ExtractTextAsLines (single column)
            └─ WriteToExcel(pagesData, outputPath)
                 └─ ZipArchive (xlsx) → XElement sheet XML → WriteXmlEntry
```

### Conversion Flow: PDF → PPT

```
PdfToPptPage.BtnConvert_Click
  └─ await Task.Run
       └─ PdfToPptConverter.Convert(pdfPath, outputDir, onPageComplete)
            ├─ PdfToImageConverter.ConvertAsync(pdfPath, tempDir, onPageComplete)
            │    └─ (per page) → onPageComplete?.Invoke() → ReleaseMemory()
            └─ RenderAllPages(pdfPath, outputDir, pageCount, onPageComplete)
                 └─ For each page image:
                      ├─ CreateSlide(pptPart, imagePath, rels)
                      │    └─ ZipArchive → ppt/slides/slideN.xml, _rels/slideN.xml.rels
                      └─ onPageComplete?.Invoke() → ReleaseMemory()
            └─ WritePresentation(pptPart) → p:defaultTextStyle (9 levels)
            └─ WriteSlideMaster(masterPart) → p:txStyles (title/body/other)
            └─ WriteSlideLayout(layoutPart) → p:clrMapOvr with masterClrMapping
            └─ WriteContentTypes / WriteRels → OOXML boilerplate
```

### Conversion Flow: PDF → ScannedPdf

```
PdfToScannedPdfPage.BtnConvert_Click
  └─ await Task.Run
       └─ PdfToScannedPdfConverter.Convert(pdfPath, outputDir)
            ├─ StorageFile.GetFileFromPathAsync(pdfPath)
            ├─ PdfDocument.LoadFromFileAsync(file)
            └─ For each page:
                 ├─ pdfDoc.GetPage(index)
                 ├─ page.RenderToStreamAsync(stream, options)
                 │    └─ DPI=200, DestinationWidth/Height computed from page size
                 ├─ new Bitmap(stream) → EncodeToJpeg (Quality 85)
                 └─ WritePdf (raw binary):
                      ├─ Catalog obj / Pages obj (with Kids refs)
                      ├─ Page obj → Image XObject (JPEG / DCTDecode) + Content Stream
                      └─ xref table + trailer
```

### Important Technical Details

- **Late-bound COM** — all Office/WPS interop uses `Type.GetTypeFromProgID` + `Activator.CreateInstance` with `dynamic`. No primary interop assemblies (PIAs) needed.
- **Window guard** (`OfficePdfConverter.RunWithGuard`) — critical for unattended Office conversions. Uses `WinEventHook(EVENT_OBJECT_CREATE)` + a 5ms polling thread to hide Office windows immediately on creation. Hides via `ShowWindow(SW_HIDE)`, removes `WS_VISIBLE`/`WS_EX_APPWINDOW` styles, sets `WDA_MONITOR` display affinity.
- **PowerPoint on STA thread** — `OfficePdfConverter.ConvertPpt` runs on a dedicated STA thread with 120s timeout; errors are marshaled back via `threadError`.
- **Edge HTML conversion** — injects `@page { size: 1920px 1080px }` CSS, replaces `position:fixed/sticky` with `static`, then launches `msedge.exe --headless=new` with 60s timeout. Waits up to 6s for PDF file to stabilize after process exit.
- **WinRT RCW cleanup** — `PdfToImageConverter` explicitly releases all WinRT objects (`PdfDocument`, `PdfPage`, `PdfPageRenderOptions`, `StorageFile`) via `Marshal.FinalReleaseComObject` in `finally` blocks to prevent memory leaks and process hang.
- **ReleaseMemory()** — `PdfToImagePage`, `PdfToPptPage`, `PdfToExcelPage`, `PdfToScannedPdfPage`, and `PdfToWordPage` all call a comprehensive cleanup after conversion: `GCSettings.LargeObjectHeapCompactionMode = CompactOnce`, double `GC.Collect()`, and `EmptyWorkingSet()` to return freed pages to the OS. `PdfToWordConverterSpire` also calls this internally after each batch.
- **PdfPig text extraction** — `PdfToExcelConverter` uses `UglyToad.PdfPig` for PDF text extraction. No COM or native PDF library needed. Words are clustered into rows by Y-coordinate with adaptive tolerance, then columns detected by left-edge alignment clustering.
- **Pptx writer** — `PdfToPptConverter` generates a valid `.pptx` file using only `System.IO.Compression.ZipArchive` and `System.Xml.Linq.XElement`. Creates `[Content_Types].xml`, `_rels/.rels`, `ppt/presentation.xml`, `ppt/slideMasters/slideMaster1.xml`, `ppt/slideLayouts/slideLayout1.xml`, `ppt/slides/slideN.xml`, `ppt/_rels/presentation.xml.rels`, and per-slide rels files. Uses `p:blipFill` with `a:blip` to embed PNG images. Adds `p:defaultTextStyle` (9 levels lvl1pPr-lvl9pPr), `p:txStyles` (titleStyle/bodyStyle/otherStyle), and `p:clrMapOvr` with `masterClrMapping` for full PowerPoint compatibility.
- **Process exit guarantee** — `App.xaml.cs` hooks `MainWindow.Closed` to call `Process.Kill()` as a safety net, ensuring the process terminates even if WinRT runtime threads linger.
- **Single-file publish** — conditionally includes `Costura.Fody` only when `SingleFile=true` MSBuild property is set.

## Constraints

- **Windows-only** — targets `net48` (Windows-only), uses Win32 P/Invoke extensively.
- **PDF to Image requires Windows 10+** — `Windows.Data.Pdf` API is only available on Windows 10/11.
- **Doc to PDF requires installed Office** — either WPS Office or Microsoft Office must be installed on the target machine.
- **PDF to Word** — Spire.PDF engine (local DLL, no Office needed) or Microsoft Office COM automation. Spire engine is the default.
- **PDF to Excel needs no Office** — uses PdfPig (open-source) for PDF parsing and pure XML for xlsx generation. No Office or COM dependency.
- **PDF to PPT needs no Office** — uses Windows built-in PDF renderer for images and pure XML for pptx generation. No Office or COM dependency.
- **PDF to Image needs no extra software** — uses Windows built-in PDF renderer, zero external dependencies.
- **PDF to ScannedPdf needs no extra software** — uses Windows built-in PDF renderer plus raw PDF writer, zero external dependencies.
- **x64 msedge.exe** — Edge is found via registry (`SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe`) or well-known install paths.