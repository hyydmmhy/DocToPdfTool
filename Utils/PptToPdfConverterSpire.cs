using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Spire.Presentation;

namespace DocToPdfTool.Utils
{
    /// <summary>
    /// 基于 Spire.Presentation 的 PPT 转 PDF 转换器。
    /// 使用本地破解版 Spire.Presentation.dll（Libs\Spire.Presentation.dll），
    /// 已解除评估水印（Evaluation Warning）和页数限制（试用版仅输出前10页），
    /// 无需 Microsoft Office，直接转换 PPT/PPTX → PDF。
    /// </summary>
    public class PptToPdfConverterSpire : IDisposable
    {
        private bool _disposed;
        [DllImport("psapi.dll")] private static extern bool EmptyWorkingSet(IntPtr hProcess);

        /// <summary>转换单个 PPT/PPTX 文件到 PDF。</summary>
        public void Convert(string pptPath, string outputDir)
        {
            if (string.IsNullOrEmpty(pptPath)) throw new ArgumentNullException(nameof(pptPath));
            if (!File.Exists(pptPath)) throw new FileNotFoundException("PPT 文件不存在", pptPath);

            var fileName = Path.GetFileNameWithoutExtension(pptPath);
            var savePath = Path.Combine(outputDir, fileName + ".pdf");

            // 确保输出目录存在
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // 使用 Spire.Presentation 加载并转换（破解版 DLL：无页数限制、无水印）
            using (var presentation = new Presentation())
            {
                presentation.LoadFromFile(pptPath);

                // 转换所有页（破解版已解除10页试用限制）
                presentation.SaveToFile(savePath, FileFormat.PDF);
            }

            // 释放内存
            ReleaseMemory();
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
