using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;

namespace DocToPdfTool.Utils
{
    public class PdfToScannedPdfConverter : IDisposable
    {
        /// <summary>
        /// 扫描件渲染 DPI（200 = 标准扫描件质量）
        /// </summary>
        public int Dpi { get; set; } = 200;

        /// <summary>
        /// 转换单个 PDF 为扫描件 PDF（同步）
        /// </summary>
        public string Convert(string pdfPath, string outputDir)
        {
            return ConvertAsync(pdfPath, outputDir)
                .ConfigureAwait(false).GetAwaiter().GetResult();
        }

        /// <summary>
        /// 异步转换单个 PDF 为扫描件 PDF
        /// </summary>
        public async Task<string> ConvertAsync(string pdfPath, string outputDir)
        {
            object file = null;
            PdfDocument pdfDoc = null;

            try
            {
                file = await StorageFile.GetFileFromPathAsync(pdfPath).AsTask().ConfigureAwait(false);
                pdfDoc = await PdfDocument.LoadFromFileAsync((StorageFile)file).AsTask().ConfigureAwait(false);

                int pageCount = (int)pdfDoc.PageCount;
                string baseName = Path.GetFileNameWithoutExtension(pdfPath);
                string outputPath = Path.Combine(outputDir, $"{baseName}_扫描件.pdf");

                using (var fs = new FileStream(outputPath, FileMode.Create))
                using (var bw = new BinaryWriter(fs, Encoding.ASCII))
                {
                    WritePdf(bw, pdfDoc, pageCount);
                }

                return outputPath;
            }
            finally
            {
                SafeRelease(pdfDoc);
                SafeRelease(file);
            }
        }

        /// <summary>
        /// 写入完整 PDF 文件
        /// </summary>
        private void WritePdf(BinaryWriter bw, PdfDocument pdfDoc, int pageCount)
        {
            // 偏移量记录表，用于 xref
            var offsets = new List<long>();

            // 辅助：写入 ASCII 字符串
            void WriteAscii(string s) => bw.Write(Encoding.ASCII.GetBytes(s));

            // 对象总数：Catalog(1) + Pages(1) + 每页3个对象(Page + ImageXObject + ContentStream)
            int totalObjects = 2 + pageCount * 3;

            // 偏移量预留：0 号对象（free entry）+ 每个对象一个偏移
            offsets.Capacity = totalObjects + 1;

            // ===== PDF 头部 =====
            WriteAscii("%PDF-1.4\n");

            // ===== 对象 1: Catalog =====
            offsets.Add(bw.BaseStream.Position);
            WriteAscii("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");

            // ===== 对象 2: Pages（占位，稍后填充 Kids 数组） =====
            long pagesObjStart = bw.BaseStream.Position;
            offsets.Add(bw.BaseStream.Position);
            // 先写入占位，最后再回填
            WriteAscii($"2 0 obj\n<< /Type /Pages /Kids [");
            long kidsArrayStart = bw.BaseStream.Position;

            // 计算每个子 Page 的对象编号（从 3 开始，每页占 3 个编号）
            var pageObjNumbers = new int[pageCount];
            for (int i = 0; i < pageCount; i++)
                pageObjNumbers[i] = 3 + i * 3;

            // 写入 Kids 引用
            for (int i = 0; i < pageCount; i++)
            {
                WriteAscii($"{pageObjNumbers[i]} 0 R");
                if (i < pageCount - 1) WriteAscii(" ");
            }

            WriteAscii($"] /Count {pageCount} >>\nendobj\n");

            // ===== 逐页处理：渲染 JPEG → 写入 PDF 对象 =====
            for (int i = 0; i < pageCount; i++)
            {
                object page = null;
                InMemoryRandomAccessStream stream = null;
                Stream managedStream = null;

                try
                {
                    page = pdfDoc.GetPage((uint)i);
                    var pageSize = ((PdfPage)page).Size;

                    int imgW = (int)(pageSize.Width * Dpi / 72.0);
                    int imgH = (int)(pageSize.Height * Dpi / 72.0);
                    double pageW = pageSize.Width;
                    double pageH = pageSize.Height;

                    var options = new PdfPageRenderOptions
                    {
                        DestinationWidth = (uint)imgW,
                        DestinationHeight = (uint)imgH,
                    };

                    // 渲染到内存流
                    stream = new InMemoryRandomAccessStream();
                    ((PdfPage)page).RenderToStreamAsync(stream, options).AsTask().ConfigureAwait(false).GetAwaiter().GetResult();
                    stream.Seek(0);
                    managedStream = stream.AsStream();

                    byte[] jpegData;
                    using (var image = new Bitmap(managedStream))
                    {
                        managedStream = null;
                        image.SetResolution(Dpi, Dpi);
                        jpegData = EncodeToJpeg(image);
                    }

                    int objNum = pageObjNumbers[i];
                    int imgObjNum = objNum + 1;
                    int contentObjNum = objNum + 2;

                    // ---- Page 对象 ----
                    offsets.Add(bw.BaseStream.Position);
                    WriteAscii($"{objNum} 0 obj\n");
                    WriteAscii($"<< /Type /Page /Parent 2 0 R");
                    WriteAscii($" /MediaBox [0 0 {pageW:F4} {pageH:F4}]");
                    WriteAscii($" /Resources << /XObject << /Im0 {imgObjNum} 0 R >> >>");
                    WriteAscii($" /Contents {contentObjNum} 0 R");
                    WriteAscii($" >>\nendobj\n");

                    // ---- Image XObject（JPEG 数据） ----
                    offsets.Add(bw.BaseStream.Position);
                    WriteAscii($"{imgObjNum} 0 obj\n");
                    WriteAscii($"<< /Type /XObject /Subtype /Image");
                    WriteAscii($" /Width {imgW} /Height {imgH}");
                    WriteAscii($" /ColorSpace /DeviceRGB /BitsPerComponent 8");
                    WriteAscii($" /Filter /DCTDecode /Length {jpegData.Length}");
                    WriteAscii($" >>\nstream\n");
                    bw.Write(jpegData);
                    WriteAscii("\nendstream\nendobj\n");

                    // ---- Content Stream（放置图片） ----
                    offsets.Add(bw.BaseStream.Position);
                    string content = $"q\n{pageW:F4} 0 0 {pageH:F4} 0 0 cm\n/Im0 Do\nQ\n";
                    byte[] contentBytes = Encoding.ASCII.GetBytes(content);
                    WriteAscii($"{contentObjNum} 0 obj\n");
                    WriteAscii($"<< /Length {contentBytes.Length} >>\nstream\n");
                    bw.Write(contentBytes);
                    WriteAscii("\nendstream\nendobj\n");
                }
                finally
                {
                    SafeRelease(page);
                    if (managedStream != null) try { managedStream.Dispose(); } catch { }
                    if (stream != null) try { stream.Dispose(); } catch { }
                }
            }

            // ===== xref 表 =====
            long xrefOffset = bw.BaseStream.Position;
            WriteAscii("xref\n");
            WriteAscii($"0 {totalObjects + 1}\n");
            // 0 号对象（free）
            WriteAscii($"0000000000 65535 f \n");
            // 每个对象的偏移量和 generation 号
            foreach (var off in offsets)
            {
                WriteAscii($"{off:D10} 00000 n \n");
            }

            // ===== trailer =====
            WriteAscii("trailer\n");
            WriteAscii($"<< /Size {totalObjects + 1} /Root 1 0 R >>\n");
            WriteAscii("startxref\n");
            WriteAscii($"{xrefOffset}\n");
            WriteAscii("%%EOF");
        }

        /// <summary>
        /// 将 Bitmap 编码为 JPEG 字节数组
        /// </summary>
        private static byte[] EncodeToJpeg(Image image)
        {
            var encoderParams = new EncoderParameters(1);
            encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 85L);
            var jpegCodec = GetEncoderInfo("image/jpeg");

            using (var ms = new MemoryStream())
            {
                image.Save(ms, jpegCodec, encoderParams);
                return ms.ToArray();
            }
        }

        private static ImageCodecInfo GetEncoderInfo(string mimeType)
        {
            foreach (var codec in ImageCodecInfo.GetImageEncoders())
                if (codec.MimeType == mimeType) return codec;
            return null;
        }

        private static void SafeRelease(object obj)
        {
            if (obj == null) return;
            try { Marshal.FinalReleaseComObject(obj); } catch { }
        }

        public void Dispose() { }
    }
}