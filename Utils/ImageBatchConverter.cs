using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace DocToPdfTool.Utils
{
    /// <summary>输出格式。</summary>
    public enum ImageTargetFormat
    {
        KeepOriginal = 0,
        Jpeg = 1,
        Png = 2,
        Bmp = 3
    }

    /// <summary>缩放方式。</summary>
    public enum ImageResizeMode
    {
        None = 0,
        ByWidth = 1,
        ByHeight = 2,
        ByLongSide = 3,
        ByPercent = 4
    }

    /// <summary>图片批量处理参数。</summary>
    public class ImageBatchOptions
    {
        public ImageTargetFormat Format { get; set; } = ImageTargetFormat.Jpeg;

        public ImageResizeMode ResizeMode { get; set; } = ImageResizeMode.None;

        /// <summary>缩放数值：像素值或百分比。</summary>
        public int ResizeValue { get; set; } = 1920;

        /// <summary>JPEG 质量 1~100（仅输出 JPG 时生效）。</summary>
        public int JpegQuality { get; set; } = 92;

        /// <summary>目标体积（KB），0 表示不限制；设置后自动搜索满足体积的最高质量。</summary>
        public int TargetSizeKb { get; set; }

        /// <summary>按 EXIF 方向自动纠正拍摄朝向。</summary>
        public bool AutoOrient { get; set; } = true;

        /// <summary>输出文件前缀，空则用原文件名。</summary>
        public string OutputSuffix { get; set; } = "_已处理";
    }

    /// <summary>单张图片的处理结果。</summary>
    public class ImageBatchResult
    {
        public string OutputFile { get; set; }
        public long OriginalSize { get; set; }
        public long NewSize { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public string Note { get; set; }

        /// <summary>内容未改动、直接复制的情况。</summary>
        public bool Copied { get; set; }
    }

    /// <summary>
    /// 图片批量处理器（基于 GDI+）。
    ///
    /// 质量策略：
    ///   - 不改格式、不缩放、不限体积时直接复制原文件，避免无谓重编码；
    ///   - 需要缩小超过 2 倍时采用逐级减半的双三次缩放，细节明显优于一次性缩放；
    ///   - JPEG 质量可指定；设置目标体积时按质量二分搜索，取满足体积的最高质量；
    ///   - 透明图转 JPG/BMP 时铺白底（可配置），PNG 输出保留透明通道。
    /// </summary>
    public class ImageBatchConverter : IDisposable
    {
        /// <summary>处理单张图片。</summary>
        public ImageBatchResult Process(string sourcePath, string outputDir, ImageBatchOptions options)
        {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("图片不存在", sourcePath);
            options = options ?? new ImageBatchOptions();

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            var result = new ImageBatchResult { OriginalSize = new FileInfo(sourcePath).Length };

            // 先读入内存，避免 Image.FromFile 锁定源文件
            var sourceBytes = File.ReadAllBytes(sourcePath);
            using (var inputStream = new MemoryStream(sourceBytes))
            using (var source = Image.FromStream(inputStream, false, true))
            {
                var frameCount = GetFrameCount(source);
                bool multiFrame = frameCount > 1;
                if (multiFrame) result.Note = $"多帧图片，仅处理首帧（共 {frameCount} 帧）";

                bool keepOriginalRequested = options.Format == ImageTargetFormat.KeepOriginal;
                var targetFormat = ResolveFormat(options.Format, sourcePath, source);
                bool needResize = options.ResizeMode != ImageResizeMode.None;
                bool needReencode = !keepOriginalRequested;
                bool needQualityControl = options.TargetSizeKb > 0;

                // 三项都不需要时直接复制，保证零损失
                if (!needResize && !needReencode && !needQualityControl)
                {
                    var copyPath = UniquePath(outputDir, BuildBaseName(sourcePath, options), Path.GetExtension(sourcePath));
                    File.WriteAllBytes(copyPath, sourceBytes);
                    result.OutputFile = copyPath;
                    result.NewSize = result.OriginalSize;
                    result.Width = source.Width;
                    result.Height = source.Height;
                    result.Copied = true;
                    result.Note = string.IsNullOrEmpty(result.Note) ? "参数未改动，已直接复制" : result.Note + "；已直接复制";
                    return result;
                }

                // 解码为位图（同时按 EXIF 纠正方向）
                using (var oriented = PrepareBitmap(source, options.AutoOrient))
                {
                    int targetWidth = oriented.Width;
                    int targetHeight = oriented.Height;
                    ComputeTargetSize(oriented.Width, oriented.Height, options, out targetWidth, out targetHeight);

                    bool resize = targetWidth != oriented.Width || targetHeight != oriented.Height;

                    using (var scaled = resize
                        ? ResizeMultiStep(oriented, targetWidth, targetHeight)
                        : (Bitmap)oriented.Clone())
                    {
                        bool keepAlpha = targetFormat == ImageTargetFormat.Png;
                        var extension = GetExtension(targetFormat, sourcePath);

                        var outputPath = UniquePath(outputDir, BuildBaseName(sourcePath, options), extension);

                        long size = SaveImage(scaled, outputPath, targetFormat, options, keepAlpha, out int usedQuality);
                        if (usedQuality > 0 && options.TargetSizeKb > 0)
                        {
                            var tip = $"质量 {usedQuality}（按目标体积自动匹配）";
                            result.Note = string.IsNullOrEmpty(result.Note) ? tip : result.Note + "；" + tip;
                        }
                        result.OutputFile = outputPath;
                        result.NewSize = size;
                        result.Width = scaled.Width;
                        result.Height = scaled.Height;
                    }
                }
            }

            return result;
        }

        // ───────────────── 格式与命名 ─────────────────

        private static ImageTargetFormat ResolveFormat(ImageTargetFormat requested, string sourcePath, Image source)
        {
            if (requested != ImageTargetFormat.KeepOriginal) return requested;

            var ext = Path.GetExtension(sourcePath).ToLowerInvariant();
            switch (ext)
            {
                case ".jpg":
                case ".jpeg": return ImageTargetFormat.Jpeg;
                case ".png": return ImageTargetFormat.Png;
                case ".bmp": return ImageTargetFormat.Bmp;
                default:
                    // GIF/TIFF 等：保持原格式需要专门编码器，这里统一落 PNG 以免失真
                    return ImageTargetFormat.Png;
            }
        }

        private static string GetExtension(ImageTargetFormat format, string sourcePath)
        {
            switch (format)
            {
                case ImageTargetFormat.Jpeg: return ".jpg";
                case ImageTargetFormat.Png: return ".png";
                case ImageTargetFormat.Bmp: return ".bmp";
                default: return Path.GetExtension(sourcePath);
            }
        }

        private static string BuildBaseName(string sourcePath, ImageBatchOptions options)
        {
            var name = Path.GetFileNameWithoutExtension(sourcePath);
            var suffix = options.OutputSuffix ?? string.Empty;
            if (suffix.Length == 0) return name;
            return name + suffix;
        }

        // ───────────────── 位图准备与缩放 ─────────────────

        /// <summary>解码为 32 位位图，并按 EXIF 方向纠正。</summary>
        private static Bitmap PrepareBitmap(Image source, bool autoOrient)
        {
            var format = HasAlpha(source) ? PixelFormat.Format32bppArgb : PixelFormat.Format24bppRgb;
            var bitmap = new Bitmap(source.Width, source.Height, format);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.CompositingMode = CompositingMode.SourceCopy;
                g.InterpolationMode = InterpolationMode.NearestNeighbor;
                g.PixelOffsetMode = PixelOffsetMode.Half;
                g.DrawImage(source, new Rectangle(0, 0, source.Width, source.Height));
            }

            if (!autoOrient) return bitmap;

            int orientation = ReadOrientation(source);
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
            return bitmap;
        }

        /// <summary>判断图片是否含透明通道。</summary>
        private static bool HasAlpha(Image image)
        {
            switch (image.PixelFormat)
            {
                case PixelFormat.Format32bppArgb:
                case PixelFormat.Format32bppPArgb:
                case PixelFormat.Format64bppArgb:
                case PixelFormat.Format64bppPArgb:
                    return true;
                default:
                    return false;
            }
        }

        private static int ReadOrientation(Image image)
        {
            try
            {
                if (Array.IndexOf(image.PropertyIdList, 0x0112) < 0) return 1;
                var item = image.GetPropertyItem(0x0112);
                return item?.Value != null && item.Value.Length > 0 ? item.Value[0] : 1;
            }
            catch { return 1; }
        }

        /// <summary>按参数计算目标尺寸（保持宽高比）。</summary>
        private static void ComputeTargetSize(int width, int height, ImageBatchOptions options,
            out int targetWidth, out int targetHeight)
        {
            targetWidth = width;
            targetHeight = height;

            switch (options.ResizeMode)
            {
                case ImageResizeMode.ByWidth:
                {
                    int w = Math.Max(1, options.ResizeValue);
                    targetWidth = w;
                    targetHeight = Math.Max(1, (int)Math.Round((double)height * w / width));
                    break;
                }
                case ImageResizeMode.ByHeight:
                {
                    int h = Math.Max(1, options.ResizeValue);
                    targetHeight = h;
                    targetWidth = Math.Max(1, (int)Math.Round((double)width * h / height));
                    break;
                }
                case ImageResizeMode.ByLongSide:
                {
                    int side = Math.Max(1, options.ResizeValue);
                    if (width >= height)
                    {
                        targetWidth = side;
                        targetHeight = Math.Max(1, (int)Math.Round((double)height * side / width));
                    }
                    else
                    {
                        targetHeight = side;
                        targetWidth = Math.Max(1, (int)Math.Round((double)width * side / height));
                    }
                    break;
                }
                case ImageResizeMode.ByPercent:
                {
                    double pct = Math.Max(1, Math.Min(1000, options.ResizeValue)) / 100.0;
                    targetWidth = Math.Max(1, (int)Math.Round(width * pct));
                    targetHeight = Math.Max(1, (int)Math.Round(height * pct));
                    break;
                }
            }
        }

        /// <summary>逐级减半缩放：缩小超过 2 倍时先多次折半，避免一次性缩放的锯齿与细节丢失。</summary>
        private static Bitmap ResizeMultiStep(Bitmap source, int targetWidth, int targetHeight)
        {
            var current = source;
            bool owned = false;

            while (current.Width >= targetWidth * 2 && current.Height >= targetHeight * 2)
            {
                int w = Math.Max(targetWidth, current.Width / 2);
                int h = Math.Max(targetHeight, current.Height / 2);
                var next = DrawScaled(current, w, h);
                if (owned) current.Dispose();
                current = next;
                owned = true;
            }

            var final = DrawScaled(current, targetWidth, targetHeight);
            if (owned) current.Dispose();
            return final;
        }

        private static Bitmap DrawScaled(Bitmap source, int width, int height)
        {
            bool alpha = source.PixelFormat == PixelFormat.Format32bppArgb
                      || source.PixelFormat == PixelFormat.Format32bppPArgb;
            var bitmap = new Bitmap(width, height, alpha ? PixelFormat.Format32bppArgb : PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(alpha ? Color.Transparent : Color.White);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.DrawImage(source, new Rectangle(0, 0, width, height));
            }
            return bitmap;
        }

        // ───────────────── 编码 ─────────────────

        /// <summary>保存图片，返回文件字节数；JPEG 在设置目标体积时按质量二分搜索。</summary>
        private static long SaveImage(Bitmap bitmap, string path, ImageTargetFormat format,
            ImageBatchOptions options, bool keepAlpha, out int usedQuality)
        {
            usedQuality = 0;
            if (format == ImageTargetFormat.Jpeg || format == ImageTargetFormat.Bmp)
            {
                // JPG/BMP 无透明通道：铺白底
                using (var flattened = Flatten(bitmap))
                {
                    if (format == ImageTargetFormat.Bmp)
                    {
                        flattened.Save(path, ImageFormat.Bmp);
                        return new FileInfo(path).Length;
                    }
                    return SaveJpeg(flattened, path, options, out usedQuality);
                }
            }

            // PNG 保留透明
            bitmap.Save(path, ImageFormat.Png);
            return new FileInfo(path).Length;
        }

        private static long SaveJpeg(Bitmap bitmap, string path, ImageBatchOptions options, out int usedQuality)
        {
            long limit = (long)Math.Max(0, options.TargetSizeKb) * 1024;

            if (limit <= 0)
            {
                usedQuality = Math.Min(100, Math.Max(1, options.JpegQuality));
                SaveJpegWithQuality(bitmap, path, usedQuality);
                return new FileInfo(path).Length;
            }

            // 目标体积：在 25~98 之间二分搜索满足体积的最高质量
            int low = 25, high = 98, best = low;
            long bestSize = long.MaxValue;

            while (low <= high)
            {
                int mid = (low + high) / 2;
                long size = EncodeJpeg(bitmap, mid).Length;
                if (size <= limit)
                {
                    best = mid;
                    bestSize = size;
                    low = mid + 1;
                }
                else
                {
                    high = mid - 1;
                }
            }

            // 最低质量仍超标时用最低质量输出
            usedQuality = bestSize == long.MaxValue ? 25 : best;
            SaveJpegWithQuality(bitmap, path, usedQuality);
            return new FileInfo(path).Length;
        }

        private static void SaveJpegWithQuality(Bitmap bitmap, string path, int quality)
        {
            var codec = ImageCodecInfo.GetImageEncoders().FirstOrDefault(c => c.MimeType == "image/jpeg");
            if (codec == null)
            {
                bitmap.Save(path, ImageFormat.Jpeg);
                return;
            }

            using (var parameters = new EncoderParameters(1))
            using (var qualityParam = new EncoderParameter(
                System.Drawing.Imaging.Encoder.Quality, (long)Math.Min(100, Math.Max(1, quality))))
            {
                parameters.Param[0] = qualityParam;
                bitmap.Save(path, codec, parameters);
            }
        }

        private static MemoryStream EncodeJpeg(Bitmap bitmap, int quality)
        {
            var stream = new MemoryStream();
            var codec = ImageCodecInfo.GetImageEncoders().FirstOrDefault(c => c.MimeType == "image/jpeg");
            using (var parameters = new EncoderParameters(1))
            using (var qualityParam = new EncoderParameter(
                System.Drawing.Imaging.Encoder.Quality, (long)Math.Min(100, Math.Max(1, quality))))
            {
                parameters.Param[0] = qualityParam;
                bitmap.Save(stream, codec, parameters);
            }
            return stream;
        }

        /// <summary>把带透明通道的位图铺到白底上。</summary>
        private static Bitmap Flatten(Bitmap source)
        {
            if (source.PixelFormat == PixelFormat.Format24bppRgb) return (Bitmap)source.Clone();
            var bitmap = new Bitmap(source.Width, source.Height, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.DrawImageUnscaled(source, 0, 0);
            }
            return bitmap;
        }

        // ───────────────── 辅助 ─────────────────

        private static int GetFrameCount(Image image)
        {
            try
            {
                var dimension = new FrameDimension(image.FrameDimensionsList[0]);
                return image.GetFrameCount(dimension);
            }
            catch { return 1; }
        }

        private static string UniquePath(string dir, string baseName, string extension)
        {
            var path = Path.Combine(dir, baseName + extension);
            int seq = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(dir, $"{baseName}_{seq++}{extension}");
            }
            return path;
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
