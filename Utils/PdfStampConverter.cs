using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using Spire.Pdf;
using Spire.Pdf.Graphics;

namespace DocToPdfTool.Utils
{
    /// <summary>盖章位置（九宫格）。</summary>
    public enum PdfStampPosition
    {
        TopLeft = 0,
        TopCenter = 1,
        TopRight = 2,
        MiddleLeft = 3,
        Center = 4,
        MiddleRight = 5,
        BottomLeft = 6,
        BottomCenter = 7,
        BottomRight = 8
    }

    /// <summary>盖章参数。</summary>
    public class PdfStampOptions
    {
        /// <summary>章图片路径（PNG 透明底最佳）。</summary>
        public string ImagePath { get; set; }

        /// <summary>九宫格位置。</summary>
        public PdfStampPosition Position { get; set; } = PdfStampPosition.BottomRight;

        /// <summary>章宽度 = 页面宽度 × 该比例（0.05~0.5），默认 0.18。</summary>
        public double WidthRatio { get; set; } = 0.18;

        /// <summary>旋转角度（度，顺时针），默认 0。</summary>
        public double Rotation { get; set; }

        /// <summary>内边距（占页面短边的比例，0~0.1），默认 0.03。</summary>
        public double MarginRatio { get; set; } = 0.03;

        /// <summary>页码范围，如 "1-3,5"；空或 "全部" 表示全部。</summary>
        public string PageRange { get; set; }
    }

    /// <summary>盖章结果。</summary>
    public class PdfStampResult
    {
        public string OutputFile { get; set; }
        public int PageCount { get; set; }
        public int StampedPages { get; set; }
        public long OriginalSize { get; set; }
        public long NewSize { get; set; }
    }

    /// <summary>
    /// PDF 电子签章 / 图片盖章转换器（基于 Spire.Pdf，零新增依赖）。
    ///
    /// 在指定位置把章图片绘制到页面上，支持九宫格定位、大小比例、旋转、
    /// 透明度、页码范围与批量。章图片用 PdfImage.FromFile 加载，透明 PNG 效果最佳。
    /// </summary>
    public class PdfStampConverter : IDisposable
    {
        public PdfStampResult Stamp(string sourcePath, string outputDir, PdfStampOptions options)
        {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("PDF 文件不存在", sourcePath);
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (string.IsNullOrEmpty(options.ImagePath) || !File.Exists(options.ImagePath))
                throw new FileNotFoundException("章图片不存在", options.ImagePath);

            if (PdfEncryptConverter.IsEncrypted(sourcePath))
                throw new InvalidOperationException("文件已加密，请先用「PDF加密 → 移除密码」解除后再盖章");

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            long originalSize = new FileInfo(sourcePath).Length;
            var result = new PdfStampResult { OriginalSize = originalSize };

            var outputFile = UniquePath(outputDir,
                Path.GetFileNameWithoutExtension(sourcePath) + "_已盖章");

            var ranges = ParsePageRanges(options.PageRange);

            int stamped = 0;
            Bitmap srcBmp = null, rotatedBmp = null;
            try
            {
                // 加载章图；非 0 旋转时用 GDI+ 高质量预旋转，避免 Spire 旋转插值导致细线变暗
                // 注意：PdfImage.FromImage 是延迟编码，位图必须存活到 SaveToFile 之后
                srcBmp = new Bitmap(options.ImagePath);
                float angle = (float)(options.Rotation % 360);
                if (Math.Abs(angle) > 0.01f)
                {
                    rotatedBmp = RotateBitmap(srcBmp, angle);
                }

                PdfImage sealImage = rotatedBmp != null
                    ? PdfImage.FromImage(rotatedBmp)
                    : PdfImage.FromImage(srcBmp);

                using (var doc = new PdfDocument())
                {
                    doc.LoadFromFile(sourcePath);
                    int total = doc.Pages.Count;
                    result.PageCount = total;

                    for (int i = 0; i < total; i++)
                    {
                        if (!InRanges(i + 1, ranges)) continue;

                        var page = doc.Pages[i];
                        var canvas = page.Canvas;
                        var size = page.Size;

                        float pageW = size.Width;
                        float pageH = size.Height;
                        float shortSide = Math.Min(pageW, pageH);

                        // 章宽 = 页面宽度 × 比例
                        float stampW = (float)(pageW * Clamp(options.WidthRatio, 0.05, 0.5));
                        float stampH = stampW * sealImage.Height / Math.Max(1, sealImage.Width);

                        float margin = (float)(shortSide * Clamp(options.MarginRatio, 0, 0.1));
                        var (cx, cy) = GetAnchor(pageW, pageH, options.Position, margin, stampW, stampH);

                        // 只平移定位；旋转已在 GDI+ 里完成，此处不再 RotateTransform
                        canvas.DrawImage(sealImage,
                            new RectangleF(cx - stampW / 2f, cy - stampH / 2f, stampW, stampH));
                        stamped++;
                    }

                    if (stamped == 0)
                        throw new InvalidOperationException("页码范围没有匹配到任何页面");

                    doc.SaveToFile(outputFile);
                }
            }
            finally
            {
                rotatedBmp?.Dispose();
                srcBmp?.Dispose();
            }

            result.OutputFile = outputFile;
            result.StampedPages = stamped;
            result.NewSize = new FileInfo(outputFile).Length;
            return result;
        }

        /// <summary>根据九宫格位置计算章的中心坐标。</summary>
        private static (float cx, float cy) GetAnchor(float pageW, float pageH,
            PdfStampPosition pos, float margin, float stampW, float stampH)
        {
            float m = margin;
            switch (pos)
            {
                case PdfStampPosition.TopLeft: return (m + stampW / 2f, m + stampH / 2f);
                case PdfStampPosition.TopCenter: return (pageW / 2f, m + stampH / 2f);
                case PdfStampPosition.TopRight: return (pageW - m - stampW / 2f, m + stampH / 2f);
                case PdfStampPosition.MiddleLeft: return (m + stampW / 2f, pageH / 2f);
                case PdfStampPosition.Center: return (pageW / 2f, pageH / 2f);
                case PdfStampPosition.MiddleRight: return (pageW - m - stampW / 2f, pageH / 2f);
                case PdfStampPosition.BottomLeft: return (m + stampW / 2f, pageH - m - stampH / 2f);
                case PdfStampPosition.BottomCenter: return (pageW / 2f, pageH - m - stampH / 2f);
                case PdfStampPosition.BottomRight: return (pageW - m - stampW / 2f, pageH - m - stampH / 2f);
                default: return (pageW / 2f, pageH / 2f);
            }
        }

        /// <summary>
        /// 用 GDI+ 高质量把章图绕中心旋转指定角度，返回新的 Bitmap。
        /// 使用 HighQualityBicubic + HighQuality 平滑 + HighQuality 合成，避免细线条被插值劣化变暗。
        /// </summary>
        private static Bitmap RotateBitmap(Bitmap src, float angle)
        {
            int w = src.Width, h = src.Height;
            var dst = new Bitmap(w, h, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(dst))
            {
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.Clear(Color.Transparent);
                g.TranslateTransform(w / 2f, h / 2f);
                g.RotateTransform(angle);
                g.TranslateTransform(-w / 2f, -h / 2f);
                g.DrawImage(src, 0, 0, w, h);
            }
            return dst;
        }

        /// <summary>解析页码范围，支持 "1-3,5"；空表示全部。</summary>
        private static List<int[]> ParsePageRanges(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return null;
            var text = input.Trim();
            if (text == "全部" || text.Equals("all", StringComparison.OrdinalIgnoreCase)) return null;
            try
            {
                return PdfSplitter.ParseRanges(text);
            }
            catch (Exception ex)
            {
                throw new FormatException("页码范围格式错误：" + ex.Message);
            }
        }

        private static bool InRanges(int page, List<int[]> ranges)
        {
            if (ranges == null || ranges.Count == 0) return true;
            foreach (var r in ranges)
            {
                if (page >= r[0] && page <= r[1]) return true;
            }
            return false;
        }

        private static double Clamp(double v, double min, double max)
        {
            if (v < min) return min;
            if (v > max) return max;
            return v;
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

        public void Dispose() { }
    }
}
