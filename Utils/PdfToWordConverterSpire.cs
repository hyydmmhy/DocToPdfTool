using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Spire.Pdf;

namespace DocToPdfTool.Utils
{
    /// <summary>
    /// 基于 Spire.PDF 的 PDF 转 Word 转换器（免费版，已破解页数限制与评估水印）。
    /// 大文件分批转换以控制内存峰值，每批转换后强制 GC + EmptyWorkingSet 归还内存。
    /// 无需 Microsoft Office。
    /// </summary>
    public class PdfToWordConverterSpire : IDisposable
    {
        private const int BATCH_SIZE = 20;
        private bool _disposed;

        [DllImport("psapi.dll")]
        private static extern bool EmptyWorkingSet(IntPtr hProcess);

        /// <summary>转换单个 PDF 文件到 DOCX。</summary>
        public void Convert(string pdfPath, string outputDir)
        {
            if (string.IsNullOrEmpty(pdfPath)) throw new ArgumentNullException(nameof(pdfPath));
            if (!File.Exists(pdfPath)) throw new FileNotFoundException("PDF 文件不存在", pdfPath);

            var fileName = Path.GetFileNameWithoutExtension(pdfPath);
            var savePath = Path.Combine(outputDir, fileName + ".docx");

            // ===== 用 PdfPig 轻量获取总页数（流式解析，不暴涨内存） =====
            int totalPages;
            using (var pdf = UglyToad.PdfPig.PdfDocument.Open(pdfPath))
            {
                totalPages = pdf.NumberOfPages;
            }

            if (totalPages <= BATCH_SIZE)
            {
                // 小文件：直接转换
                var tempPath = Path.Combine(outputDir, fileName + "_temp.docx");
                try
                {
                    using (var doc = new PdfDocument())
                    {
                        doc.LoadFromFile(pdfPath);
                        doc.SaveToFile(tempPath, FileFormat.DOCX);
                    }
                    RemoveWatermark(tempPath, savePath);
                }
                finally
                {
                    if (File.Exists(tempPath))
                        try { File.Delete(tempPath); } catch { }
                }
            }
            else
            {
                // 大文件：分批转换后合并，控制内存峰值
                // 每批创建一个新文档，用 ImportPage 从源 PDF 复制页面
                // 转换完成后释放批次文档 + 强制 GC，避免中间数据堆积
                var tempDocxFiles = new List<string>();
                try
                {
                    using (var sourceDoc = new PdfDocument())
                    {
                        sourceDoc.LoadFromFile(pdfPath);

                        for (int start = 0; start < totalPages; start += BATCH_SIZE)
                        {
                            int end = Math.Min(start + BATCH_SIZE, totalPages);
                            var batchPath = Path.Combine(outputDir, $"{fileName}_batch_{start}.docx");

                            // 创建批次文档，仅导入本批页面
                            using (var batchDoc = new PdfDocument())
                            {
                                // 逐页从源文档插入到批次文档
                                for (int i = start; i < end; i++)
                                    batchDoc.InsertPage(sourceDoc, i);

                                batchDoc.SaveToFile(batchPath, FileFormat.DOCX);
                            }

                            tempDocxFiles.Add(batchPath);

                            // 每批转换后强制释放内存，避免累积
                            ReleaseMemory();
                        }
                    }

                    // 合并所有临时 DOCX 文件
                    MergeDocxFiles(tempDocxFiles, savePath);

                    // 去除水印（如有）
                    RemoveWatermarkInPlace(savePath);
                }
                finally
                {
                    foreach (var f in tempDocxFiles)
                    {
                        try { if (File.Exists(f)) File.Delete(f); } catch { }
                    }
                }
            }

            // 全部转换完成 + 合并 + 清理后，强制归还内存
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

        /// <summary>合并多个 DOCX 文件为一个（AltChunk + 分页符）。</summary>
        private void MergeDocxFiles(List<string> sourcePaths, string targetPath)
        {
            if (sourcePaths == null || sourcePaths.Count == 0)
                throw new ArgumentException("没有可合并的文件");

            // 以第一个文件为基准
            File.Copy(sourcePaths[0], targetPath, true);
            if (sourcePaths.Count <= 1) return;

            using (var targetDoc = WordprocessingDocument.Open(targetPath, true))
            {
                var mainPart = targetDoc.MainDocumentPart;
                var body = mainPart.Document.Body;

                int chunkId = 1;
                foreach (var sourcePath in sourcePaths.Skip(1))
                {
                    // 添加分页符，确保每批内容从新页开始
                    body.Append(new Paragraph(
                        new Run(new Break { Type = BreakValues.Page })));

                    var altChunkPart = mainPart.AddAlternativeFormatImportPart(
                        AlternativeFormatImportPartType.WordprocessingML,
                        $"chunk{chunkId}");

                    using (var stream = File.Open(sourcePath, FileMode.Open, FileAccess.Read))
                    {
                        altChunkPart.FeedData(stream);
                    }

                    var altChunk = new AltChunk { Id = $"chunk{chunkId}" };
                    body.Append(altChunk);
                    chunkId++;
                }

                mainPart.Document.Save();
            }
        }

        /// <summary>原地去除 Spire.PDF 评估水印。</summary>
        private void RemoveWatermarkInPlace(string filePath)
        {
            using (var wd = WordprocessingDocument.Open(filePath, true))
            {
                var body = wd.MainDocumentPart.Document.Body;
                var elements = body.Elements().ToList();

                bool changed = false;
                foreach (var elem in elements)
                {
                    if (elem.InnerText.Contains("Evaluation Warning"))
                    {
                        elem.Remove();
                        changed = true;
                    }
                }

                if (changed)
                {
                    wd.MainDocumentPart.Document.Save();
                }
            }
        }

        /// <summary>去除 Spire.PDF 评估水印（复制模式）。</summary>
        private void RemoveWatermark(string tempPath, string savePath)
        {
            File.Copy(tempPath, savePath, true);

            using (var wd = WordprocessingDocument.Open(savePath, true))
            {
                var body = wd.MainDocumentPart.Document.Body;
                var elements = body.Elements().ToList();

                bool changed = false;
                foreach (var elem in elements)
                {
                    if (elem.InnerText.Contains("Evaluation Warning"))
                    {
                        elem.Remove();
                        changed = true;
                    }
                }

                if (changed)
                {
                    wd.MainDocumentPart.Document.Save();
                }
            }
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