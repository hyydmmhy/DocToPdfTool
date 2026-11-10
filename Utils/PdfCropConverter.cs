using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using Spire.Pdf;
using Spire.Pdf.Graphics;

namespace DocToPdfTool.Utils
{
    /// <summary>裁剪模式。</summary>
    public enum PdfCropMode
    {
        /// <summary>自动去白边：按内容包围盒裁剪。</summary>
        Auto = 0,

        /// <summary>手动裁剪：按指定边距裁剪。</summary>
        Manual = 1
    }

    /// <summary>PDF 裁剪参数。</summary>
    public class PdfCropOptions
    {
        public PdfCropMode Mode { get; set; } = PdfCropMode.Auto;

        /// <summary>自动模式：像素亮度 &gt; 该值视为白（0~255），默认 245。</summary>
        public int WhiteThreshold { get; set; } = 245;

        /// <summary>自动模式：保留安全边距（毫米），默认 2。</summary>
        public double AutoMarginMm { get; set; } = 2;

        /// <summary>手动模式：四边裁剪量（毫米）。</summary>
        public double TopMm { get; set; }
        public double BottomMm { get; set; }
        public double LeftMm { get; set; }
        public double RightMm { get; set; }
    }

    /// <summary>单文件裁剪结果。</summary>
    public class PdfCropResult
    {
        public string OutputFile { get; set; }
        public int PageCount { get; set; }
        public long OriginalSize { get; set; }
        public long NewSize { get; set; }

        /// <summary>实际被裁剪的页数（自动模式检测到内容边界变化才计入）。</summary>
        public int AffectedPages { get; set; }

        /// <summary>未检测到可裁白边而跳过。</summary>
        public bool Skipped { get; set; }
    }

    /// <summary>
    /// PDF 裁剪 / 去白边转换器（基于 Spire.Pdf，矢量无损）。
    ///
    /// 实现要点：
    ///   - 修改每页 CropBox 裁切框，不重新渲染、不栅格化，文字层保留可搜索复制；
    ///   - 自动模式：把页面渲染为低分辨率位图，检测非白像素包围盒，反算页面坐标；
    ///   - 手动模式：按毫米换算为磅，从四周裁掉指定量；
    ///   - 仅当裁剪范围合法（不越界、不反向）才写入，异常时跳过该页。
    /// </summary>
    public class PdfCropConverter : IDisposable
    {
        /// <summary>自动检测时用于内容边界分析的渲染 DPI（足够低，速度快）。</summary>
        private const int AnalyzeDpi = 72;

        /// <summary>毫米转磅。</summary>
        private const double MmToPoint = 72.0 / 25.4;

        /// <summary>裁剪单个 PDF，返回结果。</summary>
        public PdfCropResult Crop(string sourcePath, string outputDir, PdfCropOptions options)
        {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("PDF 文件不存在", sourcePath);
            options = options ?? new PdfCropOptions();

            if (PdfEncryptConverter.IsEncrypted(sourcePath))
                throw new InvalidOperationException("文件已加密，请先用「PDF加密 → 移除密码」解除后再裁剪");

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            long originalSize = new FileInfo(sourcePath).Length;
            var result = new PdfCropResult { OriginalSize = originalSize };

            var outputFile = UniquePath(outputDir,
                Path.GetFileNameWithoutExtension(sourcePath) + "_已裁剪");

            int affected = 0;
            using (var doc = new PdfDocument())
            {
                doc.LoadFromFile(sourcePath);
                int total = doc.Pages.Count;
                result.PageCount = total;

                for (int i = 0; i < total; i++)
                {
                    var page = doc.Pages[i];
                    var media = page.MediaBox;

                    RectangleF crop;
                    if (options.Mode == PdfCropMode.Auto)
                    {
                        crop = DetectAutoCrop(doc, i, media, options);
                    }
                    else
                    {
                        crop = ManualCrop(media, options);
                    }

                    if (crop.Width <= 0 || crop.Height <= 0) continue;   // 非法，跳过该页

                    // 与原裁切框比较，有变化才写入
                    if (IsDifferent(page.CropBox, crop))
                    {
                        page.CropBox = crop;
                        affected++;
                    }
                }

                if (affected == 0)
                {
                    result.Skipped = true;
                    result.NewSize = originalSize;
                    return result;
                }

                doc.SaveToFile(outputFile);
            }

            result.OutputFile = outputFile;
            result.AffectedPages = affected;
            result.NewSize = new FileInfo(outputFile).Length;
            return result;
        }

        // ───────────────── 自动检测 ─────────────────

        /// <summary>自动检测内容包围盒，返回新的 CropBox（磅）。</summary>
        private static RectangleF DetectAutoCrop(PdfDocument doc, int pageIndex, RectangleF media, PdfCropOptions options)
        {
            try
            {
                // 把当前页渲染为低分辨率位图（用 MediaBox 尺寸，忽略旋转影响）
                var image = doc.SaveAsImage(pageIndex, PdfImageType.Bitmap, AnalyzeDpi, AnalyzeDpi);
                using (image)
                {
                    var bmp = image as Bitmap;
                    if (bmp == null) return media;   // 渲染失败，保持原样

                    var bounds = FindContentBounds(bmp, options.WhiteThreshold);
                    if (bounds == Rectangle.Empty) return media;   // 整页白，不裁

                    // 位图像素 → 页面坐标（磅）
                    float pageW = media.Width;
                    float pageH = media.Height;
                    float scaleX = pageW / bmp.Width;
                    float scaleY = pageH / bmp.Height;

                    float x = bounds.X * scaleX;
                    float y = bounds.Y * scaleY;
                    float w = bounds.Width * scaleX;
                    float h = bounds.Height * scaleY;

                    // 减去保留边距
                    double margin = options.AutoMarginMm * MmToPoint;
                    float left = x - (float)margin;
                    float top = y - (float)margin;
                    float right = x + w + (float)margin;
                    float bottom = y + h + (float)margin;

                    // 裁剪到 MediaBox 内
                    left = Math.Max(0, left);
                    top = Math.Max(0, top);
                    right = Math.Min(media.Width, right);
                    bottom = Math.Min(media.Height, bottom);

                    if (right - left <= 0 || bottom - top <= 0) return media;
                    return new RectangleF(left, top, right - left, bottom - top);
                }
            }
            catch
            {
                return media;   // 检测失败，保持原样
            }
        }

        /// <summary>在低分辨率位图上找非白像素包围盒。</summary>
        private static Rectangle FindContentBounds(Bitmap bmp, int threshold)
        {
            int minX = int.MaxValue, minY = int.MaxValue;
            int maxX = -1, maxY = -1;

            for (int y = 0; y < bmp.Height; y++)
            {
                for (int x = 0; x < bmp.Width; x++)
                {
                    var c = bmp.GetPixel(x, y);
                    // 亮度 = 加权平均；大于阈值视为白
                    double lum = 0.299 * c.R + 0.587 * c.G + 0.114 * c.B;
                    if (lum < threshold)
                    {
                        if (x < minX) minX = x;
                        if (x > maxX) maxX = x;
                        if (y < minY) minY = y;
                        if (y > maxY) maxY = y;
                    }
                }
            }

            if (maxX < minX || maxY < minY) return Rectangle.Empty;
            return Rectangle.FromLTRB(minX, minY, maxX + 1, maxY + 1);
        }

        // ───────────────── 手动裁剪 ─────────────────

        /// <summary>按指定毫米边距裁剪。</summary>
        private static RectangleF ManualCrop(RectangleF media, PdfCropOptions options)
        {
            float left = (float)(options.LeftMm * MmToPoint);
            float top = (float)(options.TopMm * MmToPoint);
            float right = (float)(options.RightMm * MmToPoint);
            float bottom = (float)(options.BottomMm * MmToPoint);

            float x = Math.Max(0, left);
            float y = Math.Max(0, top);
            float w = media.Width - left - right;
            float h = media.Height - top - bottom;

            if (w <= 0 || h <= 0) return RectangleF.Empty;
            return new RectangleF(x, y, w, h);
        }

        // ───────────────── 辅助 ─────────────────

        private static bool IsDifferent(RectangleF a, RectangleF b)
        {
            const float eps = 0.5f;   // 0.5 磅以内视为无变化
            return Math.Abs(a.X - b.X) > eps
                || Math.Abs(a.Y - b.Y) > eps
                || Math.Abs(a.Width - b.Width) > eps
                || Math.Abs(a.Height - b.Height) > eps;
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

        public static string FormatSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024.0).ToString("0.0") + " KB";
            return (bytes / 1024.0 / 1024.0).ToString("0.00") + " MB";
        }

        public void Dispose() { }
    }
}
