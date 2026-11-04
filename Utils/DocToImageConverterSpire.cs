using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using Spire.Doc;
using Spire.Doc.Documents;

namespace DocToPdfTool.Utils
{
    /// <summary>
    /// 基于 Spire.Doc 的 Word 转图片转换器（内置引擎，无需 Office）。
    ///
    /// 实现要点：调用 SaveToImages(ImageType.Metafile) 取得矢量 EMF，
    /// 再按目标 DPI 绘制到高分辨率位图上。
    /// 相比 SaveToImages(ImageType.Bitmap) 固定 96 DPI 的位图输出，
    /// 矢量重渲染可按任意 DPI 输出且文字/表格完整。
    /// </summary>
    public class DocToImageConverterSpire : IDisposable
    {
        private bool _disposed;
        [DllImport("psapi.dll")] private static extern bool EmptyWorkingSet(IntPtr hProcess);

        /// <summary>Spire.Doc 输出 EMF 的原生分辨率（由 PhysicalDimension 推算，约 113.4 DPI）。</summary>
        private const double EmfBaseDpi = 113.41;

        /// <summary>转换单个 Word 文件的每一页为图片。</summary>
        /// <param name="docPath">源 Word 路径</param>
        /// <param name="outputDir">输出目录</param>
        /// <param name="dpi">目标渲染 DPI（如 96/150/300）</param>
        /// <param name="usePng">true 输出 PNG，false 输出 JPG</param>
        /// <param name="onPageComplete">每页完成后的回调（用于释放内存）</param>
        public List<string> Convert(string docPath, string outputDir, int dpi, bool usePng,
            Action onPageComplete = null)
        {
            if (string.IsNullOrEmpty(docPath)) throw new ArgumentNullException(nameof(docPath));
            if (!File.Exists(docPath)) throw new FileNotFoundException("Word 文件不存在", docPath);
            if (dpi <= 0) dpi = 150;

            var baseName = Path.GetFileNameWithoutExtension(docPath);
            var ext = usePng ? ".png" : ".jpg";
            var result = new List<string>();
            double scale = dpi / EmfBaseDpi;

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            using (var doc = new Document())
            {
                doc.LoadFromFile(docPath);

                Image[] pages = doc.SaveToImages(ImageType.Metafile);
                if (pages == null || pages.Length == 0)
                    throw new InvalidOperationException("未能从 Word 文档渲染出任何页面");

                try
                {
                    for (int i = 0; i < pages.Length; i++)
                    {
                        var meta = pages[i];
                        if (meta == null) continue;

                        // 留 1px 余量，避免边缘被裁切
                        int w = Math.Max(1, (int)Math.Ceiling(meta.Width * scale)) + 1;
                        int h = Math.Max(1, (int)Math.Ceiling(meta.Height * scale)) + 1;

                        var outputFile = Path.Combine(outputDir, $"{baseName}_{i + 1}{ext}");

                        using (var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb))
                        {
                            bmp.SetResolution(dpi, dpi);
                            using (var g = Graphics.FromImage(bmp))
                            {
                                g.Clear(Color.White);
                                g.SmoothingMode = SmoothingMode.HighQuality;
                                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                                g.CompositingQuality = CompositingQuality.HighQuality;
                                g.DrawImage(meta, new Rectangle(0, 0, w, h));
                            }
                            SaveImage(bmp, outputFile, usePng);
                        }

                        result.Add(outputFile);
                        onPageComplete?.Invoke();
                    }
                }
                finally
                {
                    foreach (var p in pages)
                    {
                        try { p?.Dispose(); } catch { }
                    }
                }
            }

            ReleaseMemory();
            return result;
        }

        /// <summary>保存图片；JPG 使用高质量编码。</summary>
        private static void SaveImage(Image image, string path, bool usePng)
        {
            if (usePng)
            {
                image.Save(path, ImageFormat.Png);
                return;
            }

            var encoder = GetEncoder(ImageFormat.Jpeg);
            if (encoder == null)
            {
                image.Save(path, ImageFormat.Jpeg);
                return;
            }

            using (var parameters = new EncoderParameters(1))
            using (var quality = new EncoderParameter(Encoder.Quality, 100L))
            {
                parameters.Param[0] = quality;
                image.Save(path, encoder, parameters);
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

        /// <summary>强制 GC 回收 + 压缩大对象堆 + 归还内存给操作系统。</summary>
        private static void ReleaseMemory()
        {
            try
            {
                System.Runtime.GCSettings.LargeObjectHeapCompactionMode =
                    System.Runtime.GCLargeObjectHeapCompactionMode.CompactOnce;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                EmptyWorkingSet(Process.GetCurrentProcess().Handle);
            }
            catch { }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
            }
        }
    }
}
