using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace DocToPdfTool.Utils
{
    /// <summary>OCR 参数。</summary>
    public class OcrOptions
    {
        /// <summary>PDF 渲染 DPI（200 / 300 / 400）。</summary>
        public int Dpi { get; set; } = 300;

        /// <summary>页码范围，如 "1-5,8"；空表示全部（仅 PDF 有效）。</summary>
        public string PageRange { get; set; } = string.Empty;

        /// <summary>是否在文本中插入「第 N 页」分隔行（默认不插入）。</summary>
        public bool PageSeparator { get; set; } = false;
    }

    /// <summary>单文件识别结果。</summary>
    public class OcrFileResult
    {
        public string OutputFile { get; set; }
        public int PageCount { get; set; }
        public int CharCount { get; set; }
        public double Seconds { get; set; }
        public bool Empty => CharCount == 0;
    }

    /// <summary>
    /// 文字识别（OCR）入口：图片或 PDF → 文本。
    ///
    /// 依赖：
    ///   - 首次使用需下载 PP-OCRv4 模型（约 15 MB），见 <see cref="OcrModelManager"/>；
    ///   - 推理走 OnnxRuntime（已内嵌），不依赖 Office / 系统 OCR 语言包。
    /// </summary>
    public class OcrConverter : IDisposable
    {
        private static readonly string[] ImageExtensions =
            { ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff" };

        private PpOcrEngine _engine;

        /// <summary>识别单个文件并写出 txt。</summary>
        public async Task<OcrFileResult> RecognizeFileAsync(string sourcePath, string outputDir,
            OcrOptions options, Action<int, int> onPageProgress = null,
            Action<string, int> onModelProgress = null, CancellationToken cancellationToken = default)
        {
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("文件不存在", sourcePath);
            options = options ?? new OcrOptions();

            var started = DateTime.Now;
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            if (!OcrModelManager.IsReady)
                await OcrModelManager.EnsureModelsAsync(onModelProgress, cancellationToken).ConfigureAwait(false);

            _engine = _engine ?? new PpOcrEngine();

            var extension = Path.GetExtension(sourcePath).ToLowerInvariant();
            var pages = new List<string>();

            if (extension == ".pdf")
            {
                pages = await RecognizePdfAsync(sourcePath, options, onPageProgress, cancellationToken)
                    .ConfigureAwait(false);
            }
            else if (ImageExtensions.Contains(extension))
            {
                onPageProgress?.Invoke(1, 1);
                using (var bitmap = LoadBitmap(sourcePath))
                {
                    pages.Add(RenderPage(1, _engine.Recognize(bitmap), options));
                }
            }
            else
            {
                throw new NotSupportedException("暂不支持该格式：" + extension);
            }

            var text = string.Join(Environment.NewLine, pages);
            var outputFile = UniquePath(outputDir, Path.GetFileNameWithoutExtension(sourcePath), ".txt");
            File.WriteAllText(outputFile, text.TrimEnd() + Environment.NewLine, new UTF8Encoding(true));

            return new OcrFileResult
            {
                OutputFile = outputFile,
                PageCount = pages.Count,
                CharCount = text.Length,
                Seconds = (DateTime.Now - started).TotalSeconds
            };
        }

        private async Task<List<string>> RecognizePdfAsync(string pdfPath, OcrOptions options,
            Action<int, int> onPageProgress, CancellationToken cancellationToken)
        {
            var pages = new List<string>();
            var tempDir = Path.Combine(Path.GetTempPath(), "ocr_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);

            try
            {
                int totalPages;
                using (var pig = UglyToad.PdfPig.PdfDocument.Open(pdfPath))
                    totalPages = pig.NumberOfPages;

                var wanted = ParsePageRange(options.PageRange, totalPages);

                // 一次性加载 PDF，逐页渲染 → 识别 → 释放位图，避免大文档占满内存
                var storageFile = await Windows.Storage.StorageFile.GetFileFromPathAsync(pdfPath)
                    .AsTask().ConfigureAwait(false);

                Windows.Data.Pdf.PdfDocument pdfDocument = null;
                try
                {
                    pdfDocument = await Windows.Data.Pdf.PdfDocument.LoadFromFileAsync(storageFile)
                        .AsTask().ConfigureAwait(false);

                    for (int i = 0; i < wanted.Count; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        int pageNumber = wanted[i];
                        onPageProgress?.Invoke(i + 1, wanted.Count);

                        using (var bitmap = await RenderPdfPageAsync(pdfDocument, pageNumber, options.Dpi)
                            .ConfigureAwait(false))
                        {
                            pages.Add(RenderPage(pageNumber, _engine.Recognize(bitmap), options));
                        }
                    }
                }
                finally
                {
                    ReleaseWinRt(pdfDocument);
                    ReleaseWinRt(storageFile);
                }
            }
            finally
            {
                try { Directory.Delete(tempDir, true); } catch { }
            }

            return pages;
        }

        /// <summary>用系统 PDF 渲染器把指定页渲染为位图。</summary>
        private static async Task<Bitmap> RenderPdfPageAsync(Windows.Data.Pdf.PdfDocument document,
            int pageNumber, int dpi)
        {
            dpi = Math.Min(600, Math.Max(150, dpi));

            Windows.Data.Pdf.PdfPage page = null;
            Windows.Storage.Streams.InMemoryRandomAccessStream stream = null;
            Windows.Data.Pdf.PdfPageRenderOptions renderOptions = null;

            try
            {
                page = document.GetPage((uint)(pageNumber - 1));
                stream = new Windows.Storage.Streams.InMemoryRandomAccessStream();

                uint width = (uint)Math.Round(page.Size.Width / 72.0 * dpi);
                uint height = (uint)Math.Round(page.Size.Height / 72.0 * dpi);

                renderOptions = new Windows.Data.Pdf.PdfPageRenderOptions
                {
                    DestinationWidth = width,
                    DestinationHeight = height
                };

                await page.RenderToStreamAsync(stream, renderOptions).AsTask().ConfigureAwait(false);

                stream.Seek(0);
                var bytes = new byte[stream.Size];
                using (var reader = new Windows.Storage.Streams.DataReader(stream.GetInputStreamAt(0)))
                {
                    await reader.LoadAsync((uint)stream.Size).AsTask().ConfigureAwait(false);
                    reader.ReadBytes(bytes);
                    reader.Dispose();
                }

                return new Bitmap(new MemoryStream(bytes));
            }
            finally
            {
                ReleaseWinRt(renderOptions);
                ReleaseWinRt(stream);
                ReleaseWinRt(page);
            }
        }

        private static void ReleaseWinRt(object instance)
        {
            if (instance == null) return;
            try { Marshal.FinalReleaseComObject(instance); } catch { }
        }

        private static string RenderPage(int pageNumber, List<OcrLine> lines, OcrOptions options)
        {
            var builder = new StringBuilder();
            if (options.PageSeparator && pageNumber > 0)
                builder.AppendLine($"===== 第 {pageNumber} 页 =====");

            foreach (var line in lines)
                builder.AppendLine(line.Text);

            return builder.ToString().TrimEnd();
        }

        // ───────────────── 图片解码 ─────────────────

        /// <summary>
        /// 读取图片为 24 位位图。
        ///
        /// 带 alpha 的图片（截图、剪贴板图片几乎都有）走 WPF 解码：
        /// GDI+ 在 alpha=0 时会丢失 RGB（克隆后整张变黑），
        /// 而 WPF 的 Bgra32 是直通格式，能保留原始 RGB。
        /// 之后再根据 alpha 是否可用决定「铺底色」还是「忽略 alpha 只取 RGB」。
        /// </summary>
        private static Bitmap LoadBitmap(string path)
        {
            var bytes = File.ReadAllBytes(path);

            using (var stream = new MemoryStream(bytes))
            using (var source = Image.FromStream(stream, false, true))
            {
                var hasAlpha = Image.IsAlphaPixelFormat(source.PixelFormat);
                var bitmap = hasAlpha ? FlattenAlphaImage(path) : ConvertOpaqueImage(source);

                ApplyExifOrientation(source, bitmap);
                return bitmap;
            }
        }

        /// <summary>无 alpha 的图：直接铺白底转 24 位。</summary>
        private static Bitmap ConvertOpaqueImage(Image source)
        {
            var bitmap = new Bitmap(source.Width, source.Height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(System.Drawing.Color.White);
                graphics.DrawImage(source, new Rectangle(0, 0, source.Width, source.Height));
            }
            return bitmap;
        }

        /// <summary>带 alpha 的图：用 WPF 解码保留 RGB，再按 alpha 情况合成到 24 位位图。</summary>
        private static Bitmap FlattenAlphaImage(string path)
        {
            using (var stream = File.OpenRead(path))
            {
                var decoder = BitmapDecoder.Create(stream,
                    BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                BitmapSource frame = decoder.Frames[0];

                var bgra = frame.Format == PixelFormats.Bgra32
                    ? frame
                    : new FormatConvertedBitmap(frame, PixelFormats.Bgra32, null, 0);

                int width = bgra.PixelWidth;
                int height = bgra.PixelHeight;
                int stride = width * 4;
                var pixels = new byte[stride * height];
                bgra.CopyPixels(pixels, stride, 0);

                // 统计 alpha 与亮度，决定合成方式
                int opaqueCount = 0;
                double luminanceSum = 0;
                for (int i = 0; i < pixels.Length; i += 4)
                {
                    if (pixels[i + 3] < 32) continue;
                    opaqueCount++;
                    luminanceSum += 0.299 * pixels[i + 2] + 0.587 * pixels[i + 1] + 0.114 * pixels[i];
                }

                // alpha 几乎全 0 说明是不可信的 alpha（很多截图工具如此），直接忽略 alpha
                bool useAlpha = opaqueCount >= Math.Max(1, width * height / 100);

                int background = 255;
                if (useAlpha && opaqueCount > 0 && luminanceSum / opaqueCount > 140)
                    background = 24;   // 浅色文字（暗色截图）铺深底

                var bitmap = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                var target = bitmap.LockBits(new Rectangle(0, 0, width, height),
                    System.Drawing.Imaging.ImageLockMode.WriteOnly,
                    System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                try
                {
                    var output = new byte[target.Stride * height];
                    for (int y = 0; y < height; y++)
                    {
                        int srcRow = y * stride;
                        int dstRow = y * target.Stride;
                        for (int x = 0; x < width; x++)
                        {
                            int s = srcRow + x * 4;
                            int d = dstRow + x * 3;

                            if (useAlpha)
                            {
                                int alpha = pixels[s + 3];
                                output[d] = (byte)((pixels[s] * alpha + background * (255 - alpha)) / 255);         // B
                                output[d + 1] = (byte)((pixels[s + 1] * alpha + background * (255 - alpha)) / 255); // G
                                output[d + 2] = (byte)((pixels[s + 2] * alpha + background * (255 - alpha)) / 255); // R
                            }
                            else
                            {
                                output[d] = pixels[s];
                                output[d + 1] = pixels[s + 1];
                                output[d + 2] = pixels[s + 2];
                            }
                        }
                    }

                    Marshal.Copy(output, 0, target.Scan0, output.Length);
                }
                finally
                {
                    bitmap.UnlockBits(target);
                }

                return bitmap;
            }
        }

        /// <summary>按 EXIF 方向旋转位图（手机竖拍照片）。</summary>
        private static void ApplyExifOrientation(Image source, Bitmap bitmap)
        {
            try
            {
                if (Array.IndexOf(source.PropertyIdList, 0x0112) < 0) return;

                var item = source.GetPropertyItem(0x0112);
                int orientation = item?.Value != null && item.Value.Length > 0 ? item.Value[0] : 1;
                switch (orientation)
                {
                    case 2: bitmap.RotateFlip(RotateFlipType.RotateNoneFlipX); break;
                    case 3: bitmap.RotateFlip(RotateFlipType.Rotate180FlipNone); break;
                    case 4: bitmap.RotateFlip(RotateFlipType.Rotate180FlipX); break;
                    case 5: bitmap.RotateFlip(RotateFlipType.Rotate90FlipX); break;
                    case 6: bitmap.RotateFlip(RotateFlipType.Rotate90FlipNone); break;
                    case 7: bitmap.RotateFlip(RotateFlipType.Rotate270FlipX); break;
                    case 8: bitmap.RotateFlip(RotateFlipType.Rotate270FlipNone); break;
                }
            }
            catch { }
        }

        /// <summary>解析 "1-5,8" 形式的页码范围。</summary>
        public static List<int> ParsePageRange(string range, int totalPages)
        {
            var result = new List<int>();
            if (string.IsNullOrWhiteSpace(range))
            {
                for (int i = 1; i <= totalPages; i++) result.Add(i);
                return result;
            }

            var seen = new HashSet<int>();
            foreach (var part in range.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var token = part.Trim();
                if (token.Length == 0) continue;

                int dash = token.IndexOfAny(new[] { '-', '~', '－' });
                if (dash > 0)
                {
                    if (int.TryParse(token.Substring(0, dash).Trim(), out int start) &&
                        int.TryParse(token.Substring(dash + 1).Trim(), out int end))
                    {
                        if (start > end) { var t = start; start = end; end = t; }
                        for (int i = Math.Max(1, start); i <= Math.Min(totalPages, end); i++)
                            if (seen.Add(i)) result.Add(i);
                    }
                }
                else if (int.TryParse(token, out int single))
                {
                    if (single >= 1 && single <= totalPages && seen.Add(single)) result.Add(single);
                }
            }

            result.Sort();
            if (result.Count == 0)
                for (int i = 1; i <= totalPages; i++) result.Add(i);

            return result;
        }

        private static string UniquePath(string dir, string baseName, string extension)
        {
            var path = Path.Combine(dir, baseName + extension);
            int seq = 1;
            while (File.Exists(path)) path = Path.Combine(dir, $"{baseName}_{seq++}{extension}");
            return path;
        }

        public void Dispose()
        {
            _engine?.Dispose();
            _engine = null;
        }
    }
}
