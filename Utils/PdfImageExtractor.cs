using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using Spire.Pdf;
using Spire.Pdf.Utilities;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using SpirePdfDocument = Spire.Pdf.PdfDocument;

namespace DocToPdfTool.Utils
{
    /// <summary>提取图片的输出格式。</summary>
    public enum PdfImageOutputFormat
    {
        /// <summary>保持原始编码（可直接取到 JPEG 时原样导出，其余转 PNG）。</summary>
        Original = 0,

        /// <summary>全部转为 PNG（无损）。</summary>
        Png = 1,

        /// <summary>全部转为 JPEG（有损，体积小）。</summary>
        Jpeg = 2
    }

    /// <summary>PDF 提取图片参数。</summary>
    public class PdfImageExtractOptions
    {
        public PdfImageOutputFormat Format { get; set; } = PdfImageOutputFormat.Original;

        /// <summary>最小边长过滤：宽或高小于该值的图片视为图标跳过；0 表示不过滤。</summary>
        public int MinSize { get; set; } = 64;

        /// <summary>跳过内容完全相同的重复图片。</summary>
        public bool SkipDuplicates { get; set; } = true;

        /// <summary>转为 JPEG 时的质量（1~100）。</summary>
        public int JpegQuality { get; set; } = 92;
    }

    /// <summary>单个 PDF 的提取结果。</summary>
    public class PdfImageExtractResult
    {
        public string OutputDir { get; set; }
        public int TotalImages { get; set; }
        public int SavedImages { get; set; }
        public int SkippedSmall { get; set; }
        public int SkippedDuplicate { get; set; }
        public int SkippedUnsupported { get; set; }
        public int PageCount { get; set; }

        /// <summary>实际使用的解码通道（PdfPig 直取 / 引擎解码）。</summary>
        public string Channel { get; set; } = "PdfPig";
    }

    /// <summary>
    /// PDF 内嵌图片提取器。
    ///
    /// 双通道设计：
    ///   1) PdfPig 直取：DCTDecode（JPEG）图片可拿到原始编码字节，JPEG 原样落盘，画质与体积零损失；
    ///   2) 引擎解码兜底：部分 PDF（如经其它工具重新压缩、图片为 Flate/ICC/带蒙版）
    ///      无法由 PdfPig 转码，此时改用本地 PDF 引擎解码成位图后按需保存。
    /// 两者都支持最小边长过滤与 SHA-256 去重。
    /// </summary>
    public class PdfImageExtractor : IDisposable
    {
        private static readonly byte[] JpegMagic = { 0xFF, 0xD8, 0xFF };
        private static readonly byte[] PngMagic = { 0x89, 0x50, 0x4E, 0x47 };

        /// <summary>提取单个 PDF 中的全部图片。</summary>
        public PdfImageExtractResult Extract(string pdfPath, string outputRoot, PdfImageExtractOptions options)
        {
            if (string.IsNullOrEmpty(pdfPath)) throw new ArgumentNullException(nameof(pdfPath));
            if (!File.Exists(pdfPath)) throw new FileNotFoundException("PDF 文件不存在", pdfPath);
            options = options ?? new PdfImageExtractOptions();

            if (PdfEncryptConverter.IsEncrypted(pdfPath))
                throw new InvalidOperationException("文件已加密，请先用「PDF加密 → 移除密码」解除后再提取");

            var baseName = Path.GetFileNameWithoutExtension(pdfPath);
            var outDir = MakeUniqueDir(outputRoot, baseName);

            // 第一通道：PdfPig 直取（保留 JPEG 原始字节）
            var result = ExtractWithPdfPig(pdfPath, outDir, options);

            // 出现无法转码的图片时，整份文件改走引擎解码，保证不漏图
            if (result.SkippedUnsupported > 0)
            {
                ClearDirectory(outDir);
                var fallback = ExtractWithSpire(pdfPath, outDir, options);
                if (fallback.SavedImages > 0 || result.SavedImages == 0)
                    return fallback;
            }

            return result;
        }

        // ───────────────── 通道一：PdfPig ─────────────────

        private static PdfImageExtractResult ExtractWithPdfPig(string pdfPath, string outDir, PdfImageExtractOptions options)
        {
            var result = new PdfImageExtractResult { OutputDir = outDir, Channel = "PdfPig" };
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using (var doc = UglyToad.PdfPig.PdfDocument.Open(pdfPath))
            {
                result.PageCount = doc.NumberOfPages;

                for (int pageNumber = 1; pageNumber <= doc.NumberOfPages; pageNumber++)
                {
                    var page = doc.GetPage(pageNumber);
                    int indexOnPage = 0;

                    foreach (var image in page.GetImages())
                    {
                        result.TotalImages++;

                        if (image.IsImageMask) { result.SkippedUnsupported++; continue; }

                        int w = image.WidthInSamples;
                        int h = image.HeightInSamples;
                        if (options.MinSize > 0 && (w < options.MinSize || h < options.MinSize))
                        {
                            result.SkippedSmall++;
                            continue;
                        }

                        byte[] data;
                        string ext;
                        if (!TryGetRawBytesViaPdfPig(image, options, out data, out ext))
                        {
                            // 含 Flate/Predictor/ICC 等无法直取的编码，交回上层改用引擎解码
                            result.SkippedUnsupported++;
                            continue;
                        }

                        if (options.SkipDuplicates && !seen.Add(Hash(data)))
                        {
                            result.SkippedDuplicate++;
                            continue;
                        }

                        indexOnPage++;
                        File.WriteAllBytes(NextPath(outDir, pageNumber, indexOnPage, ext), data);
                        result.SavedImages++;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 从 PdfPig 取图片的原始编码字节并落盘。
        /// 该 PdfPig 版本的 TryGetBytes / TryGetPng 不可用，但 RawBytes 可用：
        ///   - DCTDecode（JPEG）→ 原样写出 .jpg，字画质与体积零损失；
        ///   - 内嵌 PNG → 原样写出 .png；
        ///   - 其他编码（Flate 等，可能带 Predictor）→ 返回 false 交由引擎解码。
        /// 需要「统一 PNG / 统一 JPEG」时，把原始字节解码后重新编码。
        /// </summary>
        private static bool TryGetRawBytesViaPdfPig(IPdfImage image, PdfImageExtractOptions options,
            out byte[] data, out string ext)
        {
            data = null;
            ext = null;

            try
            {
                var raw = image.RawBytes != null && image.RawBytes.Count > 0
                    ? image.RawBytes.ToArray()
                    : null;
                if (raw == null) return false;

                bool isJpeg = StartsWith(raw, JpegMagic);
                bool isPng = StartsWith(raw, PngMagic);
                if (!isJpeg && !isPng) return false;

                if (options.Format == PdfImageOutputFormat.Original)
                {
                    data = raw;
                    ext = isJpeg ? "jpg" : "png";
                    return true;
                }

                // 需要转码：先解码为位图再按目标格式写出
                using (var input = new MemoryStream(raw))
                using (var bitmap = new Bitmap(input))
                using (var output = new MemoryStream())
                {
                    if (options.Format == PdfImageOutputFormat.Jpeg)
                    {
                        SaveJpeg(bitmap, output, options.JpegQuality);
                        ext = "jpg";
                    }
                    else
                    {
                        bitmap.Save(output, ImageFormat.Png);
                        ext = "png";
                    }

                    data = output.ToArray();
                    return data.Length > 0;
                }
            }
            catch
            {
                return false;
            }
        }

        // ───────────────── 通道二：引擎解码 ─────────────────

        private static PdfImageExtractResult ExtractWithSpire(string pdfPath, string outDir, PdfImageExtractOptions options)
        {
            var result = new PdfImageExtractResult { OutputDir = outDir, Channel = "Engine" };
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using (var doc = new SpirePdfDocument())
            {
                doc.LoadFromFile(pdfPath);
                result.PageCount = doc.Pages.Count;
                var helper = new PdfImageHelper();

                for (int pageIndex = 0; pageIndex < doc.Pages.Count; pageIndex++)
                {
                    int pageNumber = pageIndex + 1;
                    int indexOnPage = 0;

                    PdfImageInfo[] images;
                    try { images = helper.GetImagesInfo(doc.Pages[pageIndex]); }
                    catch { continue; }

                    foreach (var info in images)
                    {
                        result.TotalImages++;

                        try
                        {
                            using (var image = info.Image)
                            {
                                if (image == null) { result.SkippedUnsupported++; continue; }

                                if (options.MinSize > 0 &&
                                    (image.Width < options.MinSize || image.Height < options.MinSize))
                                {
                                    result.SkippedSmall++;
                                    continue;
                                }

                                byte[] data;
                                string ext;
                                if (!EncodeImage(image, options, out data, out ext))
                                {
                                    result.SkippedUnsupported++;
                                    continue;
                                }

                                if (options.SkipDuplicates && !seen.Add(Hash(data)))
                                {
                                    result.SkippedDuplicate++;
                                    continue;
                                }

                                indexOnPage++;
                                File.WriteAllBytes(NextPath(outDir, pageNumber, indexOnPage, ext), data);
                                result.SavedImages++;
                            }
                        }
                        catch
                        {
                            result.SkippedUnsupported++;
                        }
                    }
                }
            }

            return result;
        }

        private static bool EncodeImage(Image image, PdfImageExtractOptions options, out byte[] data, out string ext)
        {
            data = null;
            ext = null;

            try
            {
                using (var ms = new MemoryStream())
                {
                    if (options.Format == PdfImageOutputFormat.Jpeg)
                    {
                        SaveJpeg(image, ms, options.JpegQuality);
                        ext = "jpg";
                    }
                    else
                    {
                        // Original / Png：解码后统一存 PNG，像素无损
                        image.Save(ms, ImageFormat.Png);
                        ext = "png";
                    }

                    data = ms.ToArray();
                    return data.Length > 0;
                }
            }
            catch
            {
                return false;
            }
        }

        private static void SaveJpeg(Image image, Stream output, int quality)
        {
            var codec = ImageCodecInfo.GetImageEncoders().FirstOrDefault(c => c.MimeType == "image/jpeg");
            if (codec == null)
            {
                image.Save(output, ImageFormat.Jpeg);
                return;
            }

            using (var parameters = new EncoderParameters(1))
            using (var qualityParam = new EncoderParameter(
                System.Drawing.Imaging.Encoder.Quality, (long)Math.Min(100, Math.Max(5, quality))))
            {
                parameters.Param[0] = qualityParam;
                image.Save(output, codec, parameters);
            }
        }

        // ───────────────── 公共辅助 ─────────────────

        private static string NextPath(string dir, int pageNumber, int indexOnPage, string ext)
        {
            var path = Path.Combine(dir, $"p{pageNumber}_{indexOnPage}.{ext}");
            int seq = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(dir, $"p{pageNumber}_{indexOnPage}_{seq++}.{ext}");
            }
            return path;
        }

        private static bool StartsWith(byte[] data, byte[] magic)
        {
            if (data == null || data.Length < magic.Length) return false;
            for (int i = 0; i < magic.Length; i++)
            {
                if (data[i] != magic[i]) return false;
            }
            return true;
        }

        private static string Hash(byte[] data)
        {
            using (var sha = SHA256.Create())
            {
                return BitConverter.ToString(sha.ComputeHash(data));
            }
        }

        private static void ClearDirectory(string dir)
        {
            try
            {
                foreach (var f in Directory.GetFiles(dir)) File.Delete(f);
            }
            catch { }
        }

        private static string MakeUniqueDir(string root, string baseName)
        {
            if (!Directory.Exists(root)) Directory.CreateDirectory(root);
            var dir = Path.Combine(root, baseName);
            int seq = 1;
            while (Directory.Exists(dir) && Directory.GetFiles(dir).Length > 0)
            {
                dir = Path.Combine(root, $"{baseName}_{seq++}");
            }
            Directory.CreateDirectory(dir);
            return dir;
        }

        public void Dispose() { }
    }
}
