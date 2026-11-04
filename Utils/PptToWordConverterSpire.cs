using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Spire.Presentation;
using Spire.Pdf;
using PptFileFormat = Spire.Presentation.FileFormat;
using PdfFileFormat = Spire.Pdf.FileFormat;

namespace DocToPdfTool.Utils
{
    /// <summary>
    /// 基于 Spire 的 PPT 转 Word 转换器。
    /// Spire.Presentation 不直接支持 DOCX 输出，采用两段式：
    /// PPT → 临时PDF（本地 Presentation 引擎）→ DOCX（本地 PDF 引擎）。
    /// 无需 Microsoft Office。
    /// </summary>
    public class PptToWordConverterSpire : IDisposable
    {
        private bool _disposed;
        [DllImport("psapi.dll")] private static extern bool EmptyWorkingSet(IntPtr hProcess);

        /// <summary>转换单个 PPT/PPTX 文件到 Word。</summary>
        public void Convert(string pptPath, string outputDir)
        {
            if (string.IsNullOrEmpty(pptPath)) throw new ArgumentNullException(nameof(pptPath));
            if (!File.Exists(pptPath)) throw new FileNotFoundException("PPT 文件不存在", pptPath);

            var fileName = Path.GetFileNameWithoutExtension(pptPath);
            var savePath = Path.Combine(outputDir, fileName + ".docx");
            var tempPdf = Path.Combine(Path.GetTempPath(), "PptToWord_" + Guid.NewGuid().ToString("N") + ".pdf");

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                // 第一步：PPT → 临时PDF（本地 Presentation 引擎）
                using (var presentation = new Presentation())
                {
                    presentation.LoadFromFile(pptPath);
                    presentation.SaveToFile(tempPdf, PptFileFormat.PDF);
                }

                // 第二步：临时PDF → DOCX（Spire.Pdf）
                using (var pdfDoc = new PdfDocument())
                {
                    pdfDoc.LoadFromFile(tempPdf);
                    pdfDoc.SaveToFile(savePath, PdfFileFormat.DOCX);
                }
            }
            finally
            {
                try { if (File.Exists(tempPdf)) File.Delete(tempPdf); } catch { }
            }

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