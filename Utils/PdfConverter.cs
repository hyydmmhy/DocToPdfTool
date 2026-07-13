using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Win32;

namespace DocToPdfTool.Utils
{
    public class PdfConverter : IDisposable
    {
        public event Action<string> LogMessage;

        public static readonly HashSet<string> SupportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
            ".txt", ".text", ".rtf",
            ".html", ".htm", ".mhtml"
        };

        private dynamic _wpsApp;
        private bool _disposed;

        public void Convert(string sourceFile, string outputFile)
        {
            if (sourceFile == null) throw new ArgumentNullException(nameof(sourceFile));
            if (!File.Exists(sourceFile)) throw new FileNotFoundException("源文件不存在", sourceFile);

            if (outputFile == null)
                outputFile = Path.ChangeExtension(sourceFile, "pdf");

            var extension = Path.GetExtension(sourceFile).ToLowerInvariant();

            if (extension == ".html" || extension == ".htm" || extension == ".mhtml")
            {
                ConvertHtml(sourceFile, outputFile);
            }
            else if (extension == ".doc" || extension == ".docx" || extension == ".txt" ||
                     extension == ".text" || extension == ".rtf")
            {
                ConvertWord(sourceFile, outputFile);
            }
            else if (extension == ".xls" || extension == ".xlsx")
            {
                ConvertExcel(sourceFile, outputFile);
            }
            else if (extension == ".ppt" || extension == ".pptx")
            {
                ConvertPpt(sourceFile, outputFile);
            }
            else
            {
                throw new NotSupportedException($"不支持的文件格式: {extension}");
            }
        }

        private void ConvertHtml(string sourceFile, string outputFile)
        {
            var edgePath = FindEdgePath();
            if (edgePath == null)
            {
                Log("未找到 Edge 浏览器，使用 WPS 转换 HTML...");
                ConvertWord(sourceFile, outputFile);
                return;
            }

            Log("使用 Edge 浏览器引擎转换 HTML...");

            // 注入 @page CSS 强制宽视口（1920px），防止打印时触发
            // 响应式 @media 断点导致布局变成单列/上下排列
            var tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".html");
            try
            {
                var html = File.ReadAllText(sourceFile);

                // 1) 直接替换 CSS 中的 position:fixed/sticky → static（防止重复出现粘性头）
                html = Regex.Replace(html,
                    @"position\s*:\s*(?:fixed|sticky)\s*;?",
                    "position:static;",
                    RegexOptions.IgnoreCase);

                // 2) 注入 @page + 分页控制 CSS
                var headEnd = html.IndexOf("</head>", StringComparison.OrdinalIgnoreCase);
                var inject =
                    "<style>" +
                    "@page { size: 1920px 1080px; margin: 0; }" +
                    "@media print {" +
                    "  p,li,pre,blockquote,figure,figcaption," +
                    "  h1,h2,h3,h4,h5,h6,img,tr{" +
                    "    page-break-inside:avoid !important;" +
                    "  }" +
                    "  li,p,td,th{orphans:3;widows:3}" +
                    "}" +
                    "</style>";

                if (headEnd >= 0)
                    html = html.Insert(headEnd, inject);
                else
                    html = inject + html;
                File.WriteAllText(tempFile, html);

                var url = "file:///" + Path.GetFullPath(tempFile).Replace('\\', '/');

                var args = $"--headless=new --disable-gpu --no-first-run " +
                           $"--no-default-browser-check --disable-sync " +
                           $"--disable-breakpad --log-level=3 " +
                           $"--window-position=-32000,-32000 " +
                           $"--window-size=1920,1080 --hide-scrollbars " +
                           $"--print-to-pdf-no-header --print-to-pdf=\"{outputFile}\" \"{url}\"";

                var psi = new ProcessStartInfo
                {
                    FileName = edgePath,
                    Arguments = args,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (var process = Process.Start(psi))
                {
                    if (process == null)
                        throw new Exception("无法启动 Edge 浏览器进程");

                    process.StandardOutput.ReadToEnd();
                    process.StandardError.ReadToEnd();

                    if (!process.WaitForExit(60000))
                    {
                        process.Kill();
                        throw new Exception("Edge 浏览器转换超时（60秒）");
                    }
                }

                // Edge 退出后文件可能尚未完全写入，最多等 6 秒
                for (int i = 0; i < 20; i++)
                {
                    if (File.Exists(outputFile))
                    {
                        var fi = new FileInfo(outputFile);
                        if (fi.Length > 0) return;
                    }
                    Thread.Sleep(300);
                }

                throw new Exception($"Edge 浏览器未能生成 PDF 文件: {outputFile}");
            }
            finally
            {
                if (File.Exists(tempFile))
                    File.Delete(tempFile);
            }
        }

        private static string FindEdgePath()
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"))
                {
                    var path = key?.GetValue(null) as string;
                    if (!string.IsNullOrEmpty(path) && File.Exists(path))
                        return path;
                }

                using (var key = Registry.LocalMachine.OpenSubKey(
                    @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"))
                {
                    var path = key?.GetValue(null) as string;
                    if (!string.IsNullOrEmpty(path) && File.Exists(path))
                        return path;
                }
            }
            catch { }

            // 兜底：检查默认安装路径
            var candidates = new[]
            {
                @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                @"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
            };
            foreach (var c in candidates)
            {
                if (File.Exists(c)) return c;
            }

            return null;
        }

        private void ConvertWord(string sourceFile, string outputFile)
        {
            EnsureWpsApp("KWps.Application");
            _wpsApp.DisplayAlerts = false;

            dynamic doc = _wpsApp.Documents.Open(sourceFile, Visible: false);
            try
            {
                doc.Repaginate();

                // WdExportFormat.wdExportFormatPDF = 17
                doc.ExportAsFixedFormat(outputFile, 17);
            }
            finally
            {
                doc.Close();
            }
        }

        private void ConvertExcel(string sourceFile, string outputFile)
        {
            EnsureWpsApp("KET.Application");
            _wpsApp.DisplayAlerts = false;

            dynamic xls = _wpsApp.Application.Workbooks.Open(sourceFile);
            try
            {
                int sheetCount = xls.Worksheets.Count;
                for (int i = 1; i <= sheetCount; i++)
                {
                    dynamic sheet = xls.Worksheets[i];
                    sheet.Activate();

                    dynamic usedRange = sheet.UsedRange;
                    usedRange.Columns.AutoFit();

                    dynamic pageSetup = sheet.PageSetup;
                    pageSetup.Zoom = false;
                    pageSetup.FitToPagesWide = 1;
                    pageSetup.FitToPagesTall = 0;
                }

                ((dynamic)xls.Worksheets[1]).Activate();

                // XlFixedFormatType.xlTypePDF = 0, XlFixedFormatQuality.xlQualityStandard = 0
                xls.ExportAsFixedFormat(0, outputFile, 0, true, false);
            }
            finally
            {
                xls.Close();
            }
        }

        private void ConvertPpt(string sourceFile, string outputFile)
        {
            EnsureWpsApp("KWPP.Application");
            _wpsApp.DisplayAlerts = false;

            // MsoTriState.msoCTrue = 1
            dynamic ppt = _wpsApp.Presentations.Open(sourceFile, 1, 1, 1);
            try
            {
                // PpSaveAsFileType.ppSaveAsPDF = 32, MsoTriState.msoTrue = -1
                ppt.SaveAs(outputFile, 32, -1);
            }
            finally
            {
                ppt.Close();
            }
        }

        private void EnsureWpsApp(string progId)
        {
            if (_wpsApp != null)
            {
                try { _wpsApp.Quit(); } catch { }
                _wpsApp = null;
            }

            var type = Type.GetTypeFromProgID(progId);
            if (type == null)
                throw new Exception($"未检测到WPS Office，请先安装WPS。ProgID: {progId}");

            _wpsApp = Activator.CreateInstance(type);
        }

        private void Log(string message)
        {
            LogMessage?.Invoke(message);
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                if (_wpsApp != null)
                {
                    try { _wpsApp.Quit(); } catch { }
                    _wpsApp = null;
                }
                _disposed = true;
            }
        }
    }
}