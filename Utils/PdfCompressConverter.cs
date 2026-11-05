using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using Spire.Pdf;
using Spire.Pdf.Conversion.Compression;
using Spire.Pdf.Graphics;
using Spire.Pdf.Utilities;

namespace DocToPdfTool.Utils
{
    /// <summary>PDF 压缩档位。</summary>
    public enum PdfCompressPreset
    {
        /// <summary>轻度：不重采样，仅按高质量重编码图片，画质接近原样。</summary>
        Light = 0,

        /// <summary>标准：重采样到 200 DPI + 质量 88，体积与画质平衡（推荐）。</summary>
        Standard = 1,

        /// <summary>强力：重采样到 150 DPI + 质量 82。</summary>
        Strong = 2,

        /// <summary>极限：重采样到 110 DPI + 质量 75，体积最小。</summary>
        Extreme = 3,

        /// <summary>自定义：自行指定目标 DPI 与 JPEG 质量。</summary>
        Custom = 4
    }

    /// <summary>PDF 压缩参数。</summary>
    public class PdfCompressOptions
    {
        public PdfCompressPreset Preset { get; set; } = PdfCompressPreset.Standard;

        /// <summary>自定义目标 DPI；0 表示不重采样（仅重编码）。</summary>
        public int CustomDpi { get; set; } = 200;

        /// <summary>自定义 JPEG 质量（50~98）。</summary>
        public int CustomQuality { get; set; } = 88;
    }

    /// <summary>单个文件的压缩结果。</summary>
    public class PdfCompressResult
    {
        public string OutputFile { get; set; }
        public long OriginalSize { get; set; }
        public long NewSize { get; set; }

        /// <summary>压缩后体积没有明显减少而未输出。</summary>
        public bool Skipped { get; set; }

        /// <summary>实际被重采样/重编码的图片数。</summary>
        public int ReplacedImages { get; set; }

        /// <summary>保持原样的图片数（有效 DPI 已低于目标）。</summary>
        public int KeptImages { get; set; }

        public double Ratio => OriginalSize <= 0 ? 1 : (double)NewSize / OriginalSize;
    }

    /// <summary>
    /// PDF 压缩器（自控 DPI + JPEG 质量版）。
    ///
    /// 设计思路：
    ///   引擎自带的压缩档位把「降质」和「降分辨率」绑死，且不允许指定 JPEG 质量，
    ///   因此这里改为自行处理图片：
    ///     1) 逐张图片按「像素 ÷ 显示物理尺寸」算有效 DPI；
    ///     2) 有效 DPI 高于目标值时才用双三次重采样到目标尺寸；
    ///     3) 按指定质量重新编码 JPEG 并替换回页面；
    ///     4) 最后由引擎做一次容器级压缩（内容流、字体等）。
    ///   文字与矢量内容完全不动，仍可搜索、可选中。
    /// </summary>
    public class PdfCompressConverter : IDisposable
    {
        /// <summary>压缩单个 PDF。</summary>
        public PdfCompressResult Compress(string sourcePath, string outputDir, PdfCompressOptions options)
        {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("PDF 文件不存在", sourcePath);
            options = options ?? new PdfCompressOptions();

            if (PdfEncryptConverter.IsEncrypted(sourcePath))
                throw new InvalidOperationException("文件已加密，请先用「PDF加密 → 移除密码」解除后再压缩");

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            MapPreset(options, out int targetDpi, out long quality);

            long originalSize = new FileInfo(sourcePath).Length;
            var result = new PdfCompressResult { OriginalSize = originalSize };

            var outputFile = UniquePath(outputDir,
                Path.GetFileNameWithoutExtension(sourcePath) + "_已压缩");
            var tempFile = outputFile + ".tmp.pdf";

            try
            {
                // 1) 图片级处理
                ResizeImages(sourcePath, tempFile, targetDpi, quality, result);

                // 2) 容器级收尾压缩
                var finalTemp = outputFile + ".final.pdf";
                var compressor = new PdfCompressor(tempFile);
                compressor.Options.CompressContents = true;
                compressor.Options.RemoveMetadata = false;
                compressor.Options.ImageCompressionOptions.CompressImage = true;
                compressor.Options.ImageCompressionOptions.ImageQuality = ImageQuality.High;
                compressor.Options.ImageCompressionOptions.ResizeImages = false;
                compressor.Options.TextCompressionOptions.CompressFonts = true;
                compressor.CompressToFile(finalTemp);

                TryDelete(tempFile);
                if (File.Exists(finalTemp)) File.Move(finalTemp, outputFile);

                long newSize = File.Exists(outputFile) ? new FileInfo(outputFile).Length : -1;

                // 3) 收益过小则不输出，避免产生无意义的副本
                long minSaving = Math.Max(1024, originalSize / 100);
                if (newSize < 0 || newSize >= originalSize || (originalSize - newSize) < minSaving)
                {
                    TryDelete(outputFile);
                    result.Skipped = true;
                    result.NewSize = originalSize;
                    return result;
                }

                result.OutputFile = outputFile;
                result.NewSize = newSize;
                return result;
            }
            catch
            {
                TryDelete(outputFile);
                TryDelete(tempFile);
                throw;
            }
        }

        /// <summary>逐页重采样并重编码图片。</summary>
        private static void ResizeImages(string sourcePath, string destPath,
            int targetDpi, long quality, PdfCompressResult result)
        {
            using (var doc = new PdfDocument())
            {
                doc.LoadFromFile(sourcePath);
                var helper = new PdfImageHelper();

                for (int pageIndex = 0; pageIndex < doc.Pages.Count; pageIndex++)
                {
                    PdfPageBase page = doc.Pages[pageIndex];
                    PdfImageInfo[] images;
                    try { images = helper.GetImagesInfo(page); }
                    catch { continue; }

                    foreach (var info in images)
                    {
                        try
                        {
                            var image = info.Image;
                            if (image == null) continue;

                            var bounds = info.Bounds;
                            if (bounds.Width <= 1 || bounds.Height <= 1) continue;

                            double dpiX = image.Width / (bounds.Width / 72.0);
                            double dpiY = image.Height / (bounds.Height / 72.0);
                            double effectiveDpi = Math.Max(dpiX, dpiY);

                            // 未超过目标 DPI：保持原样，避免无谓的重编码损失
                            if (targetDpi > 0 && effectiveDpi <= targetDpi * 1.02)
                            {
                                result.KeptImages++;
                                continue;
                            }

                            int targetWidth = image.Width;
                            int targetHeight = image.Height;
                            if (targetDpi > 0)
                            {
                                double scale = targetDpi / effectiveDpi;
                                targetWidth = Math.Max(1, (int)Math.Round(image.Width * scale));
                                targetHeight = Math.Max(1, (int)Math.Round(image.Height * scale));
                            }

                            using (var bitmap = new Bitmap(targetWidth, targetHeight, PixelFormat.Format24bppRgb))
                            {
                                using (var g = Graphics.FromImage(bitmap))
                                {
                                    g.Clear(Color.White);          // PDF 无透明通道，铺白底
                                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                                    g.SmoothingMode = SmoothingMode.HighQuality;
                                    g.DrawImage(image, 0, 0, targetWidth, targetHeight);
                                }

                                using (var ms = new MemoryStream())
                                {
                                    SaveJpeg(bitmap, ms, quality);
                                    ms.Position = 0;
                                    helper.ReplaceImage(info, PdfImage.FromStream(ms));
                                }
                            }

                            result.ReplacedImages++;
                        }
                        catch
                        {
                            // 单张图片处理失败不影响整体
                        }
                    }
                }

                doc.SaveToFile(destPath);
            }
        }

        /// <summary>档位 → 目标 DPI / JPEG 质量。0 表示不重采样。</summary>
        private static void MapPreset(PdfCompressOptions options, out int targetDpi, out long quality)
        {
            switch (options.Preset)
            {
                case PdfCompressPreset.Light:
                    targetDpi = 0; quality = 92; break;
                case PdfCompressPreset.Strong:
                    targetDpi = 150; quality = 82; break;
                case PdfCompressPreset.Extreme:
                    targetDpi = 110; quality = 75; break;
                case PdfCompressPreset.Custom:
                    targetDpi = Math.Max(0, Math.Min(600, options.CustomDpi));
                    quality = Math.Min(98, Math.Max(50, options.CustomQuality));
                    break;
                default: // Standard
                    targetDpi = 200; quality = 88; break;
            }
        }

        private static void SaveJpeg(Image image, Stream output, long quality)
        {
            var codec = ImageCodecInfo.GetImageEncoders()
                .FirstOrDefault(c => c.MimeType == "image/jpeg");
            if (codec == null)
            {
                image.Save(output, ImageFormat.Jpeg);
                return;
            }

            using (var parameters = new EncoderParameters(1))
            using (var qualityParam = new EncoderParameter(
                System.Drawing.Imaging.Encoder.Quality, Math.Min(100, Math.Max(5, quality))))
            {
                parameters.Param[0] = qualityParam;
                image.Save(output, codec, parameters);
            }
        }

        private static string UniquePath(string dir, string baseName)
        {
            var path = Path.Combine(dir, baseName + ".pdf");
            int seq = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(dir, $"{baseName}_{seq++}.pdf");
            }
            return path;
        }

        private static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }

        /// <summary>格式化体积显示。</summary>
        public static string FormatSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024.0).ToString("0.0") + " KB";
            return (bytes / 1024.0 / 1024.0).ToString("0.00") + " MB";
        }

        public void Dispose() { }
    }
}
