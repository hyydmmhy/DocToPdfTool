using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;

namespace DocToPdfTool.Utils
{
    /// <summary>
    /// PDF 转图片转换器（基于 Windows 内置 PDF 引擎，Windows 10+）
    /// </summary>
    public class PdfToImageConverter : IDisposable
    {
        public int Dpi { get; set; } = 300;
        public double Scale { get; set; } = 1.5;
        public ImageFormat OutputFormat { get; set; } = ImageFormat.Jpeg;

        private const int MaxPixelArea = 250_000_000;

        /// <summary>
        /// 转换单个 PDF 所有页为图片
        /// </summary>
        public List<string> Convert(string pdfPath, string outputDir,
            int? startPage = null, int? endPage = null)
        {
            return ConvertAsync(pdfPath, outputDir, startPage, endPage)
                .GetAwaiter().GetResult();
        }

        /// <summary>
        /// 异步转换单个 PDF 所有页为图片
        /// </summary>
        public async Task<List<string>> ConvertAsync(string pdfPath, string outputDir,
            int? startPage = null, int? endPage = null)
        {
            var result = new List<string>();

            object file = null;   // StorageFile RCW
            PdfDocument pdfDoc = null;

            try
            {
                file = await StorageFile.GetFileFromPathAsync(pdfPath);
                pdfDoc = await PdfDocument.LoadFromFileAsync((StorageFile)file);

                int total = (int)pdfDoc.PageCount;
                int start = startPage ?? 1;
                int end = endPage ?? total;

                if (start < 1) start = 1;
                if (end > total) end = total;
                if (start > end) throw new ArgumentException("起始页码不能大于结束页码");

                string baseName = Path.GetFileNameWithoutExtension(pdfPath);
                string ext = GetExtension();

                for (int i = start; i <= end; i++)
                    ConvertPage(pdfDoc, i, baseName, ext, outputDir, result);
            }
            finally
            {
                // 释放 PdfDocument（WinRT RCW）
                SafeRelease(pdfDoc);
                // 释放 StorageFile（WinRT RCW）
                SafeRelease(file);
            }

            return result;
        }

        /// <summary>
        /// 转换单页
        /// </summary>
        private void ConvertPage(PdfDocument pdfDoc, int pageIndex,
            string baseName, string ext, string outputDir, List<string> result)
        {
            object page = null;
            InMemoryRandomAccessStream stream = null;
            Stream managedStream = null;
            object options = null;

            try
            {
                page = pdfDoc.GetPage((uint)(pageIndex - 1));
                var pageSize = ((PdfPage)page).Size;

                int width = (int)(pageSize.Width * Scale * Dpi / 72.0);
                int height = (int)(pageSize.Height * Scale * Dpi / 72.0);
                var size = ClampSize(width, height);

                options = new PdfPageRenderOptions
                {
                    DestinationWidth = (uint)size.Width,
                    DestinationHeight = (uint)size.Height,
                };

                stream = new InMemoryRandomAccessStream();
                var op = ((PdfPage)page).RenderToStreamAsync(stream, (PdfPageRenderOptions)options);
                var task = op.AsTask();
                task.GetAwaiter().GetResult();

                stream.Seek(0);
                managedStream = stream.AsStream();

                using (var image = new Bitmap(managedStream))
                {
                    managedStream = null;
                    image.SetResolution(Dpi, Dpi);

                    string outputFile = Path.Combine(outputDir, $"{baseName}_{pageIndex}{ext}");
                    SaveWithQuality(image, outputFile);
                    result.Add(outputFile);
                }
            }
            finally
            {
                if (managedStream != null) try { managedStream.Dispose(); } catch { }
                if (stream != null) try { stream.Dispose(); } catch { }
                // 释放 WinRT RCW: PdfPage 和 PdfPageRenderOptions
                SafeRelease(page);
                SafeRelease(options);
            }
        }

        /// <summary>
        /// 安全释放 WinRT COM RCW 对象
        /// </summary>
        private static void SafeRelease(object obj)
        {
            if (obj == null) return;
            try
            {
                Marshal.FinalReleaseComObject(obj);
            }
            catch
            {
                // 非 COM 对象或已释放，忽略
            }
        }

        private string GetExtension()
        {
            if (OutputFormat.Guid == ImageFormat.Jpeg.Guid) return ".jpg";
            if (OutputFormat.Guid == ImageFormat.Png.Guid) return ".png";
            if (OutputFormat.Guid == ImageFormat.Bmp.Guid) return ".bmp";
            if (OutputFormat.Guid == ImageFormat.Tiff.Guid) return ".tiff";
            if (OutputFormat.Guid == ImageFormat.Gif.Guid) return ".gif";
            return ".jpg";
        }

        private void SaveWithQuality(Image image, string filePath)
        {
            if (OutputFormat == ImageFormat.Jpeg)
            {
                var encoderParams = new EncoderParameters(1);
                encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 95L);
                var jpegCodec = GetEncoderInfo("image/jpeg");
                image.Save(filePath, jpegCodec, encoderParams);
            }
            else
            {
                image.Save(filePath, OutputFormat);
            }
        }

        private static ImageCodecInfo GetEncoderInfo(string mimeType)
        {
            foreach (var codec in ImageCodecInfo.GetImageEncoders())
                if (codec.MimeType == mimeType) return codec;
            return null;
        }

        private static Size ClampSize(int width, int height)
        {
            if (width * height <= MaxPixelArea)
                return new Size(width, height);
            double ratio = Math.Sqrt((double)MaxPixelArea / (width * height));
            return new Size((int)(width * ratio), (int)(height * ratio));
        }

        public void Dispose() { }
    }
}