using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using Spire.Presentation;

namespace DocToPdfTool.Utils
{
    /// <summary>
    /// 基于本地 Presentation 引擎的 PPT 转图片转换器。
    /// 使用 Libs\Spire.Presentation.dll 本地引用，
    /// 无需 Microsoft Office，逐页导出高质量图片。
    /// </summary>
    public class PptToImageConverterSpire : IDisposable
    {
        private bool _disposed;
        [DllImport("psapi.dll")] private static extern bool EmptyWorkingSet(IntPtr hProcess);

        /// <summary>转换单个 PPT/PPTX 文件的每一页为图片。</summary>
        /// <param name="pptPath">源 PPT 路径</param>
        /// <param name="outputDir">输出目录</param>
        /// <param name="dpi">渲染 DPI（如 150/300/600）</param>
        /// <param name="usePng">true 输出 PNG，false 输出 JPG</param>
        /// <param name="onPageComplete">每页完成后的回调（用于释放内存）</param>
        public List<string> Convert(string pptPath, string outputDir, int dpi, bool usePng,
            Action onPageComplete = null)
        {
            if (string.IsNullOrEmpty(pptPath)) throw new ArgumentNullException(nameof(pptPath));
            if (!File.Exists(pptPath)) throw new FileNotFoundException("PPT 文件不存在", pptPath);
            if (dpi <= 0) dpi = 300;

            var baseName = Path.GetFileNameWithoutExtension(pptPath);
            var ext = usePng ? ".png" : ".jpg";
            var result = new List<string>();

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            using (var presentation = new Presentation())
            {
                presentation.LoadFromFile(pptPath);

                // 开启高质量渲染，显著改善文字/矢量边缘清晰度
                try { presentation.HighQualityImage = true; } catch { }

                int count = presentation.Slides.Count;
                for (int i = 0; i < count; i++)
                {
                    using (var image = presentation.SaveToImage(i, dpi, dpi))
                    {
                        var outputFile = Path.Combine(outputDir, $"{baseName}_{i + 1}{ext}");
                        SaveImage(image, outputFile, usePng);
                        result.Add(outputFile);
                    }

                    onPageComplete?.Invoke();
                }
            }

            ReleaseMemory();
            return result;
        }

        /// <summary>保存图片；JPG 使用 95% 质量编码。</summary>
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
