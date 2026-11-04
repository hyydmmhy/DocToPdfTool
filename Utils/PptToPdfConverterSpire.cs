using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Spire.Presentation;

namespace DocToPdfTool.Utils
{
    /// <summary>
    /// 基于本地 Presentation 引擎的 PPT 转 PDF 转换器。
    /// 使用 Libs\Spire.Presentation.dll 本地引用，
    /// 无需 Microsoft Office，直接转换 PPT/PPTX → PDF，大页数文档完整输出。
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

            // 使用本地 Presentation 引擎加载并转换
            using (var presentation = new Presentation())
            {
                presentation.LoadFromFile(pptPath);

                // 全量转换所有幻灯片
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
