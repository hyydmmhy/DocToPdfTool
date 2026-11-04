using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Text;

namespace DocToPdfTool.Utils
{
    /// <summary>图片转 PDF 的页面尺寸策略。</summary>
    public enum ImagePageMode
    {
        /// <summary>A4 适应：等比缩放进 A4，四周留白（默认）。</summary>
        A4Fit = 0,

        /// <summary>A4 铺满：等比放大填满 A4，超出部分裁切。</summary>
        A4Fill = 1,

        /// <summary>原始尺寸：按图片自身 DPI 换算物理尺寸，1 页 = 1 张图原始大小。</summary>
        OriginalSize = 2
    }

    /// <summary>
    /// 图片转 PDF 转换器（零依赖，纯 .NET 写入 PDF 1.4）。
    ///
    /// 实现要点：
    ///   - JPEG 原图直接以 DCTDecode 嵌入，不重编码，画质无损、速度快；
    ///   - PNG/BMP/GIF/TIFF 统一铺白底后转 JPEG 再嵌入（PDF 无透明通道）；
    ///   - 自动读取 EXIF 方向并纠正旋转（手机竖拍照片）；
    ///   - 逐页处理、逐页释放，大批量图片内存占用平稳；
    ///   - 超大图可限制长边，避免解码内存峰值过高。
    /// </summary>
    public class ImageToPdfConverter : IDisposable
    {
        private const float MmToPt = 72f / 25.4f;
        private const float A4WidthMm = 210f;
        private const float A4HeightMm = 297f;

        /// <summary>页面尺寸策略。</summary>
        public ImagePageMode PageMode { get; set; } = ImagePageMode.A4Fit;

        /// <summary>页面边距（毫米）。</summary>
        public int MarginMm { get; set; } = 10;

        /// <summary>JPEG 重编码质量（1-100）。</summary>
        public int JpegQuality { get; set; } = 92;

        /// <summary>图片长边像素上限；0 表示不限制。</summary>
        public int MaxImageSide { get; set; } = 5000;

        /// <summary>
        /// 把多张图片按列表顺序合成为一个 PDF。
        /// </summary>
        /// <param name="imagePaths">图片路径列表（顺序即页序）</param>
        /// <param name="outputFile">输出 PDF 完整路径</param>
        /// <param name="onProgress">进度回调（已处理数, 总数）</param>
        /// <param name="failures">无法处理的图片及原因</param>
        /// <returns>成功写入的页数</returns>
        public int Convert(IList<string> imagePaths, string outputFile,
            Action<int, int> onProgress = null, List<string> failures = null)
        {
            if (imagePaths == null) throw new ArgumentNullException(nameof(imagePaths));
            if (imagePaths.Count == 0) throw new ArgumentException("图片列表为空", nameof(imagePaths));
            if (string.IsNullOrEmpty(outputFile)) throw new ArgumentNullException(nameof(outputFile));

            var dir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            var pages = new List<PageEntry>(imagePaths.Count);
            int total = imagePaths.Count;

            for (int i = 0; i < total; i++)
            {
                var path = imagePaths[i];
                try
                {
                    var entry = BuildPage(path);
                    if (entry != null) pages.Add(entry);
                    else failures?.Add($"{Path.GetFileName(path)}: 无法读取图片");
                }
                catch (Exception ex)
                {
                    failures?.Add($"{Path.GetFileName(path)}: {ex.Message}");
                }

                onProgress?.Invoke(i + 1, total);
            }

            if (pages.Count == 0)
                throw new InvalidOperationException("没有可转换的图片");

            WritePdf(pages, outputFile);
            return pages.Count;
        }

        // ───────────────────────── 页面构建 ─────────────────────────

        private sealed class PageEntry
        {
            public byte[] Jpeg;
            public int ImageWidth;
            public int ImageHeight;
            public float MediaWidth;
            public float MediaHeight;
            public float DrawWidth;
            public float DrawHeight;
            public float DrawX;
            public float DrawY;
            public bool Clip;
            public float ClipX;
            public float ClipY;
            public float ClipWidth;
            public float ClipHeight;
        }

        private PageEntry BuildPage(string path)
        {
            using (var raw = Image.FromFile(path))
            {
                int orientation = GetOrientation(raw);
                bool unchanged = IsPlainJpeg(path, raw, orientation);

                // JPEG 且无需旋转 / 缩放：直接嵌入原始字节
                if (unchanged && !NeedResize(raw.Width, raw.Height))
                {
                    var entry = new PageEntry
                    {
                        Jpeg = File.ReadAllBytes(path),
                        ImageWidth = raw.Width,
                        ImageHeight = raw.Height
                    };
                    Layout(entry, raw.Width, raw.Height, raw.HorizontalResolution, raw.VerticalResolution);
                    return entry;
                }

                // 其他情况：旋转纠正 + 白底 + 可选缩放 + JPEG 重编码
                using (var normalized = Normalize(raw, orientation, MaxImageSide))
                {
                    var entry = new PageEntry
                    {
                        Jpeg = EncodeJpeg(normalized, JpegQuality),
                        ImageWidth = normalized.Width,
                        ImageHeight = normalized.Height
                    };
                    Layout(entry, normalized.Width, normalized.Height,
                        raw.HorizontalResolution, raw.VerticalResolution);
                    return entry;
                }
            }
        }

        /// <summary>计算页面尺寸与图片摆放矩形。</summary>
        private void Layout(PageEntry entry, int imgW, int imgH, float dpiX, float dpiY)
        {
            float margin = Math.Max(0, MarginMm) * MmToPt;

            float pageW, pageH;
            switch (PageMode)
            {
                case ImagePageMode.OriginalSize:
                    pageW = PixelsToPoints(imgW, dpiX) + margin * 2;
                    pageH = PixelsToPoints(imgH, dpiY) + margin * 2;
                    break;

                default:
                    // A4 按图片横竖自动选择方向
                    bool landscape = imgW > imgH;
                    pageW = (landscape ? A4HeightMm : A4WidthMm) * MmToPt;
                    pageH = (landscape ? A4WidthMm : A4HeightMm) * MmToPt;
                    break;
            }

            entry.MediaWidth = pageW;
            entry.MediaHeight = pageH;

            float contentW = Math.Max(1f, pageW - margin * 2);
            float contentH = Math.Max(1f, pageH - margin * 2);

            if (PageMode == ImagePageMode.OriginalSize)
            {
                entry.DrawWidth = Math.Max(1f, contentW);
                entry.DrawHeight = Math.Max(1f, contentH);
                entry.DrawX = margin;
                entry.DrawY = margin;
                return;
            }

            float scale = PageMode == ImagePageMode.A4Fill
                ? Math.Max(contentW / imgW, contentH / imgH)
                : Math.Min(contentW / imgW, contentH / imgH);

            float drawW = imgW * scale;
            float drawH = imgH * scale;

            entry.DrawWidth = drawW;
            entry.DrawHeight = drawH;
            entry.DrawX = margin + (contentW - drawW) / 2f;
            entry.DrawY = margin + (contentH - drawH) / 2f;

            if (PageMode == ImagePageMode.A4Fill)
            {
                entry.Clip = true;
                entry.ClipX = margin;
                entry.ClipY = margin;
                entry.ClipWidth = contentW;
                entry.ClipHeight = contentH;
            }
        }

        private static float PixelsToPoints(int pixels, float dpi)
        {
            if (dpi <= 1f || float.IsNaN(dpi) || float.IsInfinity(dpi)) dpi = 96f;
            return pixels / dpi * 72f;
        }

        private bool NeedResize(int w, int h)
        {
            return MaxImageSide > 0 && Math.Max(w, h) > MaxImageSide;
        }

        /// <summary>是否为可直接嵌入的普通 JPEG（RGB、无 EXIF 旋转）。</summary>
        private static bool IsPlainJpeg(string path, Image image, int orientation)
        {
            var ext = Path.GetExtension(path);
            bool jpegExt = string.Equals(ext, ".jpg", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(ext, ".jpeg", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(ext, ".jpe", StringComparison.OrdinalIgnoreCase);
            if (!jpegExt) return false;
            if (image.RawFormat.Guid != ImageFormat.Jpeg.Guid) return false;
            if (orientation > 1) return false;
            return image.PixelFormat == PixelFormat.Format24bppRgb;
        }

        /// <summary>读取 EXIF 方向值（0x0112），失败返回 1。</summary>
        private static int GetOrientation(Image image)
        {
            try
            {
                var ids = image.PropertyIdList;
                if (ids == null || Array.IndexOf(ids, 0x0112) < 0) return 1;
                var prop = image.GetPropertyItem(0x0112);
                if (prop?.Value == null || prop.Value.Length < 2) return 1;
                return BitConverter.ToInt16(prop.Value, 0);
            }
            catch { return 1; }
        }

        /// <summary>按 EXIF 方向纠正旋转，铺白底，必要时缩放，输出 24bpp 位图。</summary>
        private static Bitmap Normalize(Image source, int orientation, int maxSide)
        {
            Bitmap rotated = null;
            try
            {
                if (orientation > 1)
                {
                    rotated = new Bitmap(source.Width, source.Height, PixelFormat.Format24bppRgb);
                    using (var g = Graphics.FromImage(rotated))
                    {
                        g.Clear(Color.White);
                        g.DrawImage(source, 0, 0, source.Width, source.Height);
                    }

                    switch (orientation)
                    {
                        case 2: rotated.RotateFlip(RotateFlipType.RotateNoneFlipX); break;
                        case 3: rotated.RotateFlip(RotateFlipType.Rotate180FlipNone); break;
                        case 4: rotated.RotateFlip(RotateFlipType.Rotate180FlipX); break;
                        case 5: rotated.RotateFlip(RotateFlipType.Rotate90FlipX); break;
                        case 6: rotated.RotateFlip(RotateFlipType.Rotate90FlipNone); break;
                        case 7: rotated.RotateFlip(RotateFlipType.Rotate270FlipX); break;
                        case 8: rotated.RotateFlip(RotateFlipType.Rotate270FlipNone); break;
                    }
                }

                var work = rotated ?? source;

                double scale = 1.0;
                if (maxSide > 0)
                {
                    int longest = Math.Max(work.Width, work.Height);
                    if (longest > maxSide) scale = (double)maxSide / longest;
                }

                int tw = Math.Max(1, (int)Math.Round(work.Width * scale));
                int th = Math.Max(1, (int)Math.Round(work.Height * scale));

                var result = new Bitmap(tw, th, PixelFormat.Format24bppRgb);
                using (var g = Graphics.FromImage(result))
                {
                    g.Clear(Color.White);
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    g.SmoothingMode = SmoothingMode.HighQuality;
                    g.DrawImage(work, 0, 0, tw, th);
                }
                return result;
            }
            finally
            {
                rotated?.Dispose();
            }
        }

        private static byte[] EncodeJpeg(Image image, int quality)
        {
            if (quality < 1) quality = 1;
            if (quality > 100) quality = 100;

            var codec = GetEncoder(ImageFormat.Jpeg);
            using (var ms = new MemoryStream())
            {
                if (codec == null)
                {
                    image.Save(ms, ImageFormat.Jpeg);
                }
                else
                {
                    using (var parameters = new EncoderParameters(1))
                    using (var q = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, (long)quality))
                    {
                        parameters.Param[0] = q;
                        image.Save(ms, codec, parameters);
                    }
                }
                return ms.ToArray();
            }
        }

        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            foreach (var codec in ImageCodecInfo.GetImageEncoders())
            {
                if (codec.FormatID == format.Guid) return codec;
            }
            return null;
        }

        // ───────────────────────── PDF 写入 ─────────────────────────

        private static void WritePdf(List<PageEntry> pages, string outputFile)
        {
            using (var fs = new FileStream(outputFile, FileMode.Create, FileAccess.Write, FileShare.None))
            using (var bw = new BinaryWriter(fs, Encoding.ASCII))
            {
                int pageCount = pages.Count;
                // 对象编号：1=Catalog, 2=Pages, 之后每页 3 个对象（Page / Contents / Image）
                var offsets = new long[3 + pageCount * 3];

                WriteAscii(bw, "%PDF-1.4\n");
                bw.Write(new byte[] { 0x25, 0xE2, 0xE3, 0xCF, 0xD3, 0x0A }); // 二进制标识

                // 1. Catalog
                offsets[1] = fs.Position;
                WriteAscii(bw, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");

                // 2. Pages
                var kids = new StringBuilder();
                for (int i = 0; i < pageCount; i++)
                {
                    if (i > 0) kids.Append(' ');
                    kids.Append(PageObjectNumber(i)).Append(" 0 R");
                }
                offsets[2] = fs.Position;
                WriteAscii(bw, $"2 0 obj\n<< /Type /Pages /Kids [ {kids} ] /Count {pageCount} >>\nendobj\n");

                for (int i = 0; i < pageCount; i++)
                {
                    var page = pages[i];
                    int pageObj = PageObjectNumber(i);
                    int contentObj = pageObj + 1;
                    int imageObj = pageObj + 2;

                    // Page
                    offsets[pageObj] = fs.Position;
                    WriteAscii(bw,
                        $"{pageObj} 0 obj\n<< /Type /Page /Parent 2 0 R " +
                        $"/MediaBox [ 0 0 {N(page.MediaWidth)} {N(page.MediaHeight)} ] " +
                        $"/Resources << /XObject << /Im0 {imageObj} 0 R >> >> " +
                        $"/Contents {contentObj} 0 R >>\nendobj\n");

                    // Contents
                    byte[] content = BuildContent(page);
                    offsets[contentObj] = fs.Position;
                    WriteAscii(bw, $"{contentObj} 0 obj\n<< /Length {content.Length} >>\nstream\n");
                    bw.Write(content);
                    WriteAscii(bw, "\nendstream\nendobj\n");

                    // Image XObject
                    offsets[imageObj] = fs.Position;
                    WriteAscii(bw,
                        $"{imageObj} 0 obj\n<< /Type /XObject /Subtype /Image " +
                        $"/Width {page.ImageWidth} /Height {page.ImageHeight} " +
                        $"/ColorSpace /DeviceRGB /BitsPerComponent 8 " +
                        $"/Filter /DCTDecode /Length {page.Jpeg.Length} >>\nstream\n");
                    bw.Write(page.Jpeg);
                    WriteAscii(bw, "\nendstream\nendobj\n");
                }

                // xref
                int totalObjects = 3 + pageCount * 3;
                long xrefPos = fs.Position;
                WriteAscii(bw, $"xref\n0 {totalObjects}\n");
                WriteAscii(bw, "0000000000 65535 f \n");
                for (int i = 1; i < totalObjects; i++)
                {
                    WriteAscii(bw, offsets[i].ToString("D10", CultureInfo.InvariantCulture) + " 00000 n \n");
                }

                WriteAscii(bw,
                    $"trailer\n<< /Size {totalObjects} /Root 1 0 R >>\n" +
                    $"startxref\n{xrefPos}\n%%EOF\n");
            }
        }

        private static int PageObjectNumber(int pageIndex)
        {
            return 3 + pageIndex * 3;
        }

        private static byte[] BuildContent(PageEntry page)
        {
            var sb = new StringBuilder();
            if (page.Clip)
            {
                sb.Append("q\n");
                sb.Append($"{N(page.ClipX)} {N(page.ClipY)} {N(page.ClipWidth)} {N(page.ClipHeight)} re W n\n");
                sb.Append($"{N(page.DrawWidth)} 0 0 {N(page.DrawHeight)} {N(page.DrawX)} {N(page.DrawY)} cm\n");
                sb.Append("/Im0 Do\nQ\n");
            }
            else
            {
                sb.Append("q\n");
                sb.Append($"{N(page.DrawWidth)} 0 0 {N(page.DrawHeight)} {N(page.DrawX)} {N(page.DrawY)} cm\n");
                sb.Append("/Im0 Do\nQ\n");
            }
            return Encoding.ASCII.GetBytes(sb.ToString());
        }

        private static string N(float value)
        {
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private static void WriteAscii(BinaryWriter bw, string text)
        {
            bw.Write(Encoding.ASCII.GetBytes(text));
        }

        public void Dispose() { }
    }
}
