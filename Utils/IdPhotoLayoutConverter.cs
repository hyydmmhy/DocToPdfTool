using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace DocToPdfTool.Utils
{
    /// <summary>证件照规格（国内常用）。</summary>
    public enum IdPhotoSize
    {
        SmallOneInch = 0,   // 小一寸 22×32 mm
        OneInch = 1,        // 一寸 25×35 mm
        LargeOneInch = 2,   // 大一寸 33×48 mm
        SmallTwoInch = 3,   // 小二寸 35×45 mm
        TwoInch = 4,        // 二寸 35×49 mm
        LargeTwoInch = 5    // 大二寸 35×53 mm
    }

    /// <summary>打印相纸规格。</summary>
    public enum PhotoPaper
    {
        Size5R = 0,   // 5 寸 127×89 mm
        Size6R = 1,   // 6 寸 152×102 mm
        SizeA4 = 2    // A4 210×297 mm
    }

    /// <summary>证件照排版参数。</summary>
    public class IdPhotoLayoutOptions
    {
        public IdPhotoSize Size { get; set; } = IdPhotoSize.OneInch;

        public PhotoPaper Paper { get; set; } = PhotoPaper.Size6R;

        /// <summary>输出分辨率，打印建议 300。</summary>
        public int Dpi { get; set; } = 300;

        /// <summary>照片之间的间距（毫米）。</summary>
        public int GapMm { get; set; } = 2;



        /// <summary>是否按证件照比例居中裁剪（关闭则拉伸变形）。</summary>
        public bool CenterCrop { get; set; } = true;
    }

    /// <summary>排版结果。</summary>
    public class IdPhotoLayoutResult
    {
        public string OutputFile { get; set; }
        public int Columns { get; set; }
        public int Rows { get; set; }
        public int PhotoCount { get; set; }
        public int PaperWidth { get; set; }
        public int PaperHeight { get; set; }
        public int PhotoWidth { get; set; }
        public int PhotoHeight { get; set; }
    }

    /// <summary>
    /// 证件照排版器：把一张人像按标准证件照尺寸排满整张相纸，用于直接打印。
    ///
    /// 实现要点：
    ///   - 尺寸与相纸都按毫米定义，按目标 DPI 换算为像素；
    ///   - 先按证件照长宽比居中裁剪，再缩放，避免人像被拉伸变形；
    ///   - 大幅缩小时逐级减半双三次缩放，细节优于一次性缩放；
    ///   - 照片之间留白边作为裁切间隔，整张排版区域在相纸上居中，不额外画线。
    /// </summary>
    public class IdPhotoLayoutConverter : IDisposable
    {
        private const double MmPerInch = 25.4;

        /// <summary>把一张照片排版成整张相纸。</summary>
        public IdPhotoLayoutResult Layout(string sourcePath, string outputDir, IdPhotoLayoutOptions options)
        {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("照片不存在", sourcePath);
            options = options ?? new IdPhotoLayoutOptions();

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            int dpi = Math.Min(1200, Math.Max(72, options.Dpi));
            GetPhotoSizeMm(options.Size, out double photoMmW, out double photoMmH);
            GetPaperSizeMm(options.Paper, out double paperMmW, out double paperMmH);

            int photoW = Math.Max(1, (int)Math.Round(photoMmW / MmPerInch * dpi));
            int photoH = Math.Max(1, (int)Math.Round(photoMmH / MmPerInch * dpi));
            int paperW = Math.Max(1, (int)Math.Round(paperMmW / MmPerInch * dpi));
            int paperH = Math.Max(1, (int)Math.Round(paperMmH / MmPerInch * dpi));
            int gap = Math.Max(0, (int)Math.Round(options.GapMm / MmPerInch * dpi));

            // 计算可排布的行列数
            int columns = Math.Max(1, (paperW + gap) / (photoW + gap));
            int rows = Math.Max(1, (paperH + gap) / (photoH + gap));

            int usedW = columns * photoW + (columns - 1) * gap;
            int usedH = rows * photoH + (rows - 1) * gap;
            int offsetX = Math.Max(0, (paperW - usedW) / 2);
            int offsetY = Math.Max(0, (paperH - usedH) / 2);

            var outputFile = UniquePath(outputDir,
                Path.GetFileNameWithoutExtension(sourcePath) + "_证件照排版", ".jpg");

            using (var photo = LoadAndPrepare(sourcePath, photoW, photoH, options.CenterCrop))
            using (var paper = new Bitmap(paperW, paperH, PixelFormat.Format24bppRgb))
            using (var g = Graphics.FromImage(paper))
            {
                g.Clear(Color.White);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.SmoothingMode = SmoothingMode.HighQuality;

                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < columns; col++)
                    {
                        int x = offsetX + col * (photoW + gap);
                        int y = offsetY + row * (photoH + gap);
                        g.DrawImage(photo, new Rectangle(x, y, photoW, photoH));
                    }
                }

                SaveJpeg(paper, outputFile, 95);
            }

            return new IdPhotoLayoutResult
            {
                OutputFile = outputFile,
                Columns = columns,
                Rows = rows,
                PhotoCount = columns * rows,
                PaperWidth = paperW,
                PaperHeight = paperH,
                PhotoWidth = photoW,
                PhotoHeight = photoH
            };
        }

        // ───────────────── 尺寸定义 ─────────────────

        public static void GetPhotoSizeMm(IdPhotoSize size, out double width, out double height)
        {
            switch (size)
            {
                case IdPhotoSize.SmallOneInch: width = 22; height = 32; break;
                case IdPhotoSize.LargeOneInch: width = 33; height = 48; break;
                case IdPhotoSize.SmallTwoInch: width = 35; height = 45; break;
                case IdPhotoSize.TwoInch: width = 35; height = 49; break;
                case IdPhotoSize.LargeTwoInch: width = 35; height = 53; break;
                default: width = 25; height = 35; break;   // 一寸
            }
        }

        public static void GetPaperSizeMm(PhotoPaper paper, out double width, out double height)
        {
            switch (paper)
            {
                case PhotoPaper.Size5R: width = 127; height = 89; break;
                case PhotoPaper.SizeA4: width = 210; height = 297; break;
                default: width = 152; height = 102; break;  // 6 寸
            }
        }

        /// <summary>规格说明文字，供界面展示。</summary>
        public static string GetPhotoSizeText(IdPhotoSize size)
        {
            GetPhotoSizeMm(size, out double w, out double h);
            switch (size)
            {
                case IdPhotoSize.SmallOneInch: return $"小一寸 {w}×{h} mm";
                case IdPhotoSize.LargeOneInch: return $"大一寸 {w}×{h} mm";
                case IdPhotoSize.SmallTwoInch: return $"小二寸 {w}×{h} mm";
                case IdPhotoSize.TwoInch: return $"二寸 {w}×{h} mm";
                case IdPhotoSize.LargeTwoInch: return $"大二寸 {w}×{h} mm";
                default: return $"一寸 {w}×{h} mm";
            }
        }

        // ───────────────── 图像准备 ─────────────────

        /// <summary>按 EXIF 纠正方向，再按证件照比例居中裁剪并缩放到目标尺寸。</summary>
        private static Bitmap LoadAndPrepare(string path, int targetW, int targetH, bool centerCrop)
        {
            var bytes = File.ReadAllBytes(path);
            using (var stream = new MemoryStream(bytes))
            using (var source = Image.FromStream(stream, false, true))
            {
                using (var oriented = ApplyOrientation(source))
                {
                    if (!centerCrop)
                        return ResizeMultiStep(oriented, targetW, targetH);

                    // 居中裁剪到目标比例
                    double targetRatio = (double)targetW / targetH;
                    double sourceRatio = (double)oriented.Width / oriented.Height;

                    Rectangle cropRect;
                    if (sourceRatio > targetRatio)
                    {
                        int cropW = Math.Max(1, (int)Math.Round(oriented.Height * targetRatio));
                        cropRect = new Rectangle((oriented.Width - cropW) / 2, 0, cropW, oriented.Height);
                    }
                    else
                    {
                        int cropH = Math.Max(1, (int)Math.Round(oriented.Width / targetRatio));
                        cropRect = new Rectangle(0, (oriented.Height - cropH) / 2, oriented.Width, cropH);
                    }

                    using (var cropped = new Bitmap(cropRect.Width, cropRect.Height, PixelFormat.Format24bppRgb))
                    {
                        using (var g = Graphics.FromImage(cropped))
                        {
                            g.DrawImage(oriented,
                                new Rectangle(0, 0, cropped.Width, cropped.Height),
                                cropRect, GraphicsUnit.Pixel);
                        }
                        return ResizeMultiStep(cropped, targetW, targetH);
                    }
                }
            }
        }

        private static Bitmap ApplyOrientation(Image source)
        {
            var bitmap = new Bitmap(source.Width, source.Height, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.DrawImage(source, new Rectangle(0, 0, source.Width, source.Height));
            }

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

        private static Bitmap ResizeMultiStep(Bitmap source, int targetW, int targetH)
        {
            var current = source;
            bool owned = false;

            while (current.Width >= targetW * 2 && current.Height >= targetH * 2)
            {
                int w = Math.Max(targetW, current.Width / 2);
                int h = Math.Max(targetH, current.Height / 2);
                var next = DrawScaled(current, w, h);
                if (owned) current.Dispose();
                current = next;
                owned = true;
            }

            var final = DrawScaled(current, targetW, targetH);
            if (owned) current.Dispose();
            return final;
        }

        private static Bitmap DrawScaled(Bitmap source, int width, int height)
        {
            var bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.DrawImage(source, new Rectangle(0, 0, width, height));
            }
            return bitmap;
        }

        // ───────────────── 输出 ─────────────────

        private static void SaveJpeg(Bitmap bitmap, string path, int quality)
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

        public void Dispose() { }
    }
}
