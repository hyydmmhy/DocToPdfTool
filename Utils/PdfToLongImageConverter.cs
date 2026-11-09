using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace DocToPdfTool.Utils
{
    /// <summary>PDF 转长图参数。</summary>
    public class PdfToLongImageOptions
    {
        /// <summary>渲染 DPI（72/96/150/300）。</summary>
        public int Dpi { get; set; } = 150;

        /// <summary>输出格式：true = PNG，false = JPG。</summary>
        public bool OutputPng { get; set; }

        /// <summary>JPEG 质量 1~100（仅 JPG 输出生效）。</summary>
        public int JpegQuality { get; set; } = 92;
    }

    /// <summary>PDF 转长图结果。</summary>
    public class PdfToLongImageResult
    {
        public string OutputFile { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public long Size { get; set; }
        public int PageCount { get; set; }
    }

    /// <summary>
    /// PDF 转长图转换器（复用 Windows 内置 PDF 渲染器，零外部依赖）。
    ///
    /// 实现要点：
    ///   - 逐页用 PdfToImageConverter 渲染为位图（高 DPI），立即拼接到大画布；
    ///   - 每页渲染完立即释放，控制峰值内存；
    ///   - 拼接后总像素超过上限时整体等比缩小，避免 GDI+ 位图超限崩溃。
    /// </summary>
    public class PdfToLongImageConverter : IDisposable
    {
        /// <summary>GDI+ 位图最大像素面积保护（约 200 MP）。</summary>
        private const long MaxPixelArea = 200_000_000;

        /// <summary>转换 PDF 为一张长图。</summary>
        public PdfToLongImageResult Convert(string pdfPath, string outputDir, PdfToLongImageOptions options)
        {
            if (string.IsNullOrEmpty(pdfPath)) throw new ArgumentNullException(nameof(pdfPath));
            if (!File.Exists(pdfPath)) throw new FileNotFoundException("PDF 文件不存在", pdfPath);
            options = options ?? new PdfToLongImageOptions();

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            // 1) 用 WinRT 渲染器逐页渲染为位图并拼接
            var renderer = new PdfToImageConverter
            {
                Dpi = options.Dpi,
                Scale = 1.0,
                OutputFormat = ImageFormat.Png   // 先用 PNG 拼，保留透明；最终按需转 JPG
            };

            var pages = new List<Bitmap>();
            int sourcePageCount = 0;
            try
            {
                // 渲染到临时目录（避免直接写结果目录），每页回调释放内存
                string tempDir = Path.Combine(Path.GetTempPath(), "DocToPdfTool_LongImage");
                if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);

                var renderedFiles = renderer.Convert(pdfPath, tempDir,
                    onPageComplete: () => ReleaseMemory());
                sourcePageCount = renderedFiles.Count;

                // 2) 加载所有页位图
                foreach (var f in renderedFiles)
                {
                    var bytes = File.ReadAllBytes(f);
                    using (var ms = new MemoryStream(bytes))
                    {
                        pages.Add(new Bitmap(ms));
                    }
                    try { File.Delete(f); } catch { }
                }

                if (pages.Count == 0)
                    throw new InvalidOperationException("未渲染出任何页面");

                // 3) 计算拼接尺寸（纵向），超限时整体缩小
                int width = pages.Max(p => p.Width);
                long totalHeight = pages.Sum(p => (long)p.Height);
                double downscale = 1.0;
                if ((long)width * totalHeight > MaxPixelArea)
                {
                    downscale = Math.Sqrt((double)MaxPixelArea / ((long)width * totalHeight));
                    totalHeight = (long)(totalHeight * downscale);
                    width = (int)(width * downscale);
                }
                int canvasW = Math.Max(1, width);
                int canvasH = Math.Max(1, (int)totalHeight);

                string outputPath = UniquePath(outputDir,
                    Path.GetFileNameWithoutExtension(pdfPath) + "_长图",
                    options.OutputPng ? ".png" : ".jpg");

                // 4) 拼接
                using (var canvas = new Bitmap(canvasW, canvasH, PixelFormat.Format24bppRgb))
                using (var g = Graphics.FromImage(canvas))
                {
                    g.Clear(Color.White);
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = SmoothingMode.HighQuality;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    g.CompositingQuality = CompositingQuality.HighQuality;

                    int offset = 0;
                    foreach (var page in pages)
                    {
                        int drawW = Math.Min(canvasW, Math.Max(1, (int)Math.Round(page.Width * downscale)));
                        int drawH = Math.Min(canvasH - offset, Math.Max(1, (int)Math.Round(page.Height * downscale)));
                        g.DrawImage(page, new Rectangle(0, offset, drawW, drawH));
                        offset += drawH;
                        page.Dispose();
                    }
                    pages.Clear();

                    SaveWithQuality(canvas, outputPath, options);
                }

                return new PdfToLongImageResult
                {
                    OutputFile = outputPath,
                    Width = canvasW,
                    Height = canvasH,
                    Size = new FileInfo(outputPath).Length,
                    PageCount = sourcePageCount
                };
            }
            finally
            {
                foreach (var p in pages) { try { p.Dispose(); } catch { } }
                pages.Clear();
            }
        }

        // ───────────────── 编码 ─────────────────

        private static void SaveWithQuality(Bitmap bitmap, string path, PdfToLongImageOptions options)
        {
            if (options.OutputPng)
            {
                bitmap.Save(path, ImageFormat.Png);
                return;
            }

            var codec = ImageCodecInfo.GetImageEncoders().FirstOrDefault(c => c.MimeType == "image/jpeg");
            if (codec == null)
            {
                bitmap.Save(path, ImageFormat.Jpeg);
                return;
            }
            using (var parameters = new EncoderParameters(1))
            using (var qualityParam = new EncoderParameter(
                System.Drawing.Imaging.Encoder.Quality,
                (long)Math.Min(100, Math.Max(1, options.JpegQuality))))
            {
                parameters.Param[0] = qualityParam;
                bitmap.Save(path, codec, parameters);
            }
        }

        // ───────────────── 辅助 ─────────────────

        private static string UniquePath(string dir, string baseName, string ext)
        {
            var path = Path.Combine(dir, baseName + ext);
            int seq = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(dir, $"{baseName}_{seq++}{ext}");
            }
            return path;
        }

        private static void ReleaseMemory()
        {
            try
            {
                System.Runtime.GCSettings.LargeObjectHeapCompactionMode =
                    System.Runtime.GCLargeObjectHeapCompactionMode.CompactOnce;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch { }
        }

        public static string FormatSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024.0).ToString("0.0") + " KB";
            return (bytes / 1024.0 / 1024.0).ToString("0.00") + " MB";
        }

        public void Dispose() { }
    }
}
