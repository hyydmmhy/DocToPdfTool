using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace DocToPdfTool.Utils
{
    /// <summary>拼接方向。</summary>
    public enum ImageStitchDirection
    {
        /// <summary>纵向拼接（上下排列，长图）。</summary>
        Vertical = 0,

        /// <summary>横向拼接（左右排列）。</summary>
        Horizontal = 1
    }

    /// <summary>统一宽度策略。</summary>
    public enum ImageStitchWidthMode
    {
        /// <summary>保持各图片原始尺寸，不做任何缩放。</summary>
        Original = 0,

        /// <summary>以最宽的图片为基准，其余按比例缩放。</summary>
        Widest = 1,

        /// <summary>以最窄的图片为基准，其余按比例缩放。</summary>
        Narrowest = 2,

        /// <summary>统一缩放到自定义像素宽度。</summary>
        Custom = 3
    }

    /// <summary>图片拼接参数。</summary>
    public class ImageStitchOptions
    {
        public ImageStitchDirection Direction { get; set; } = ImageStitchDirection.Vertical;

        public ImageStitchWidthMode WidthMode { get; set; } = ImageStitchWidthMode.Widest;

        /// <summary>自定义统一宽度（WidthMode = Custom 时生效，像素）。</summary>
        public int CustomWidth { get; set; } = 1920;

        /// <summary>图片间距（像素）。</summary>
        public int Spacing { get; set; } = 0;

        /// <summary>背景色；true 表示透明（仅 PNG 输出生效）。</summary>
        public bool TransparentBackground { get; set; }

        /// <summary>输出格式：true = PNG，false = JPG。</summary>
        public bool OutputPng { get; set; }

        /// <summary>JPEG 质量 1~100（仅 JPG 输出生效）。</summary>
        public int JpegQuality { get; set; } = 92;

        /// <summary>按 EXIF 方向自动纠正。</summary>
        public bool AutoOrient { get; set; } = true;
    }

    /// <summary>拼接结果。</summary>
    public class ImageStitchResult
    {
        public string OutputFile { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public long Size { get; set; }
        public int ImageCount { get; set; }
    }

    /// <summary>
    /// 图片拼接长图转换器（基于 GDI+，零依赖）。
    ///
    /// 实现要点：
    ///   - 先统一所有图片到同一目标宽度（最宽/最窄/自定义），保持各自宽高比；
    ///   - 计算总画布尺寸后一次性创建位图，逐张绘制，避免反复 Resize；
    ///   - 超出 GDI+ 位图尺寸上限（约 2 亿像素）时自动整体等比缩小；
    ///   - EXIF 方向纠正、双三次缩放、JPEG 编码与 ImageBatchConverter 一致。
    /// </summary>
    public class ImageStitchConverter : IDisposable
    {
        /// <summary>GDI+ 位图最大像素面积保护（约 200 MP）。</summary>
        private const long MaxPixelArea = 200_000_000;

        /// <summary>拼接一组图片为一张长图。</summary>
        public ImageStitchResult Stitch(IList<string> sourcePaths, string outputDir, ImageStitchOptions options)
        {
            if (sourcePaths == null || sourcePaths.Count == 0)
                throw new ArgumentException("请至少添加一张图片");
            options = options ?? new ImageStitchOptions();

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            // 1) 读取并纠正每张图片，记录解码后的尺寸
            var decoded = new List<DecodedImage>();
            var scaled = new List<ScaledPiece>();
            try
            {
                foreach (var path in sourcePaths)
                {
                    if (string.IsNullOrEmpty(path) || !File.Exists(path))
                        throw new FileNotFoundException("图片不存在", path);

                    decoded.Add(LoadDecoded(path, options.AutoOrient));
                }

                // 2) 统一目标宽度；&lt;=0 表示保持各图片原始尺寸（不缩放）
                int targetWidth = ResolveTargetWidth(decoded, options);

                // 3) 计算每个元素（图片 + 间隔）在拼接轴上的尺寸
                int gap = Math.Max(0, options.Spacing);
                long mainTotal = 0;   // 拼接方向总长
                int crossMax = 0;     // 交叉方向最大宽度

                foreach (var d in decoded)
                {
                    var piece = ScaleToWidth(d.Bitmap, targetWidth);
                    scaled.Add(piece);
                    mainTotal += piece.Height;
                    crossMax = Math.Max(crossMax, piece.Width);
                }
                mainTotal += (long)gap * (scaled.Count - 1);

                bool vertical = options.Direction == ImageStitchDirection.Vertical;
                int canvasW = vertical ? crossMax : (int)mainTotal;
                int canvasH = vertical ? (int)mainTotal : crossMax;

                // 4) 超出上限时整体等比缩小
                if ((long)canvasW * canvasH > MaxPixelArea)
                {
                    double ratio = Math.Sqrt((double)MaxPixelArea / ((long)canvasW * canvasH));
                    canvasW = Math.Max(1, (int)(canvasW * ratio));
                    canvasH = Math.Max(1, (int)(canvasH * ratio));
                }

                bool alpha = options.TransparentBackground;
                string ext = options.OutputPng ? ".png" : ".jpg";
                var outputPath = UniquePath(outputDir,
                    Path.GetFileNameWithoutExtension(sourcePaths[0]) + "_拼接", ext);

                using (var canvas = new Bitmap(canvasW, canvasH,
                    alpha ? PixelFormat.Format32bppArgb : PixelFormat.Format24bppRgb))
                using (var g = Graphics.FromImage(canvas))
                {
                    g.Clear(alpha ? Color.Transparent : Color.White);
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = SmoothingMode.HighQuality;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    g.CompositingQuality = CompositingQuality.HighQuality;

                    int offset = 0;
                    for (int i = 0; i < scaled.Count; i++)
                    {
                        var p = scaled[i];
                        if (vertical)
                        {
                            g.DrawImage(p.Bitmap, new Rectangle(0, offset, p.Width, p.Height));
                            offset += p.Height + gap;
                        }
                        else
                        {
                            g.DrawImage(p.Bitmap, new Rectangle(offset, 0, p.Width, p.Height));
                            offset += p.Width + gap;
                        }
                        p.Bitmap.Dispose();
                    }
                    scaled.Clear();

                    SaveWithQuality(canvas, outputPath, options);
                }

                return new ImageStitchResult
                {
                    OutputFile = outputPath,
                    Width = canvasW,
                    Height = canvasH,
                    Size = new FileInfo(outputPath).Length,
                    ImageCount = sourcePaths.Count
                };
            }
            finally
            {
                foreach (var d in decoded) d.Bitmap.Dispose();
                foreach (var p in scaled) p.Bitmap.Dispose();
            }
        }

        // ───────────────── 解码与缩放 ─────────────────

        private static DecodedImage LoadDecoded(string path, bool autoOrient)
        {
            var bytes = File.ReadAllBytes(path);
            using (var stream = new MemoryStream(bytes))
            using (var img = Image.FromStream(stream, false, true))
            {
                var bmp = PrepareBitmap(img, autoOrient);
                return new DecodedImage { Bitmap = bmp };
            }
        }

        private static Bitmap PrepareBitmap(Image source, bool autoOrient)
        {
            bool alpha = HasAlpha(source);
            var format = alpha ? PixelFormat.Format32bppArgb : PixelFormat.Format24bppRgb;
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

        private static ScaledPiece ScaleToWidth(Bitmap source, int targetWidth)
        {
            // targetWidth &lt;= 0 表示保持原始尺寸，直接克隆
            if (targetWidth <= 0)
                return new ScaledPiece { Bitmap = (Bitmap)source.Clone(), Width = source.Width, Height = source.Height };

            int w = Math.Max(1, targetWidth);
            int h = Math.Max(1, (int)Math.Round((double)source.Height * w / source.Width));
            if (w == source.Width && h == source.Height)
                return new ScaledPiece { Bitmap = (Bitmap)source.Clone(), Width = w, Height = h };

            bool alpha = source.PixelFormat == PixelFormat.Format32bppArgb
                      || source.PixelFormat == PixelFormat.Format32bppPArgb;
            var bitmap = new Bitmap(w, h, alpha ? PixelFormat.Format32bppArgb : PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(alpha ? Color.Transparent : Color.White);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.DrawImage(source, new Rectangle(0, 0, w, h));
            }
            return new ScaledPiece { Bitmap = bitmap, Width = w, Height = h };
        }

        private static int ResolveTargetWidth(List<DecodedImage> decoded, ImageStitchOptions options)
        {
            switch (options.WidthMode)
            {
                case ImageStitchWidthMode.Original:
                    return 0;   // 0 = 保持各图片原始尺寸
                case ImageStitchWidthMode.Custom:
                    return Math.Max(1, options.CustomWidth);
                case ImageStitchWidthMode.Narrowest:
                    return decoded.Min(d => d.Bitmap.Width);
                default:
                    return decoded.Max(d => d.Bitmap.Width);
            }
        }

        // ───────────────── 编码 ─────────────────

        private static void SaveWithQuality(Bitmap bitmap, string path, ImageStitchOptions options)
        {
            if (options.OutputPng)
            {
                bitmap.Save(path, ImageFormat.Png);
                return;
            }

            using (var flattened = Flatten(bitmap))
            {
                var codec = ImageCodecInfo.GetImageEncoders().FirstOrDefault(c => c.MimeType == "image/jpeg");
                if (codec == null)
                {
                    flattened.Save(path, ImageFormat.Jpeg);
                    return;
                }
                using (var parameters = new EncoderParameters(1))
                using (var qualityParam = new EncoderParameter(
                    System.Drawing.Imaging.Encoder.Quality,
                    (long)Math.Min(100, Math.Max(1, options.JpegQuality))))
                {
                    parameters.Param[0] = qualityParam;
                    flattened.Save(path, codec, parameters);
                }
            }
        }

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

        public static string FormatSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024.0).ToString("0.0") + " KB";
            return (bytes / 1024.0 / 1024.0).ToString("0.00") + " MB";
        }

        public void Dispose() { }

        private class DecodedImage
        {
            public Bitmap Bitmap { get; set; }
        }

        private class ScaledPiece
        {
            public Bitmap Bitmap { get; set; }
            public int Width { get; set; }
            public int Height { get; set; }
        }
    }
}
