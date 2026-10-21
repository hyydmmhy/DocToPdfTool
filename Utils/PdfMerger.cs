using System;
using System.Collections.Generic;
using System.IO;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Writer;

namespace DocToPdfTool.Utils
{
    /// <summary>
    /// PDF 合并器：基于 PdfPig 的 PdfDocumentBuilder 逐页复制，
    /// 保留原始文字、矢量和图片，不会栅格化。无需 Office 或任何外部依赖。
    /// </summary>
    public class PdfMerger : IDisposable
    {
        /// <summary>
        /// 按传入顺序将多个 PDF 合并为一个文件。
        /// </summary>
        /// <param name="pdfPaths">要合并的 PDF 文件路径，顺序即合并顺序</param>
        /// <param name="outputPath">输出文件路径</param>
        public string Merge(List<string> pdfPaths, string outputPath)
        {
            if (pdfPaths == null || pdfPaths.Count == 0)
                throw new InvalidOperationException("没有可合并的 PDF 文件");

            var builder = new PdfDocumentBuilder();

            int totalPages = 0;
            foreach (var path in pdfPaths)
            {
                using (var document = PdfDocument.Open(path))
                {
                    if (document.NumberOfPages == 0)
                        throw new InvalidOperationException($"文件没有可合并的页面: {path}");

                    for (int i = 1; i <= document.NumberOfPages; i++)
                    {
                        builder.AddPage(document, i);
                    }
                    totalPages += document.NumberOfPages;
                }
            }

            if (totalPages == 0)
                throw new InvalidOperationException("所有 PDF 文件均为空");

            File.WriteAllBytes(outputPath, builder.Build());
            return outputPath;
        }

        public void Dispose()
        {
        }
    }
}