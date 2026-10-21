using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Writer;

namespace DocToPdfTool.Utils
{
    /// <summary>
    /// PDF 拆分器：基于 PdfPig 的 PdfDocumentBuilder 逐页复制，
    /// 保留原始文字、矢量和图片，不会栅格化。无需 Office 或任何外部依赖。
    /// </summary>
    public class PdfSplitter : IDisposable
    {
        /// <summary>
        /// 将 PDF 每 pagesPerFile 页拆分为一个文件，返回生成的文件路径列表。
        /// </summary>
        public List<string> Split(string pdfPath, string outputDir, int pagesPerFile)
        {
            if (pagesPerFile < 1)
                pagesPerFile = 1;

            var results = new List<string>();
            string baseName = Path.GetFileNameWithoutExtension(pdfPath);

            using (var document = PdfDocument.Open(pdfPath))
            {
                int totalPages = document.NumberOfPages;
                if (totalPages == 0)
                    throw new InvalidOperationException("PDF 文件没有可拆分的页面");

                int partCount = (totalPages + pagesPerFile - 1) / pagesPerFile;

                for (int part = 0; part < partCount; part++)
                {
                    int startPage = part * pagesPerFile + 1; // PdfPig 页码从 1 开始
                    int endPage = Math.Min(startPage + pagesPerFile - 1, totalPages);

                    var builder = new PdfDocumentBuilder();
                    for (int i = startPage; i <= endPage; i++)
                    {
                        builder.AddPage(document, i);
                    }

                    string fileName = $"{baseName}_第{part + 1}部分({startPage}-{endPage}页).pdf";
                    string outputPath = Path.Combine(outputDir, fileName);
                    File.WriteAllBytes(outputPath, builder.Build());
                    results.Add(outputPath);
                }
            }

            return results;
        }

        /// <summary>
        /// 按页码范围拆分。范围格式如 "1-5,6,7-8"，每个逗号分隔段生成一个文件。
        /// </summary>
        public List<string> SplitByRanges(string pdfPath, string outputDir, List<int[]> ranges)
        {
            var results = new List<string>();
            string baseName = Path.GetFileNameWithoutExtension(pdfPath);

            using (var document = PdfDocument.Open(pdfPath))
            {
                int totalPages = document.NumberOfPages;
                if (totalPages == 0)
                    throw new InvalidOperationException("PDF 文件没有可拆分的页面");

                for (int part = 0; part < ranges.Count; part++)
                {
                    int startPage = ranges[part][0];
                    int endPage = ranges[part][1];

                    if (startPage < 1 || endPage > totalPages || startPage > endPage)
                        throw new ArgumentOutOfRangeException(
                            $"范围 {startPage}-{endPage} 超出文档页数（共 {totalPages} 页）");

                    var builder = new PdfDocumentBuilder();
                    for (int i = startPage; i <= endPage; i++)
                    {
                        builder.AddPage(document, i);
                    }

                    string rangeLabel = startPage == endPage ? $"{startPage}" : $"{startPage}-{endPage}";
                    string fileName = $"{baseName}_第{part + 1}部分({rangeLabel}页).pdf";
                    string outputPath = Path.Combine(outputDir, fileName);
                    File.WriteAllBytes(outputPath, builder.Build());
                    results.Add(outputPath);
                }
            }

            return results;
        }

        /// <summary>
        /// 解析页码范围字符串，如 "1-5,6,7-8"。
        /// 支持中文逗号和英文逗号。返回 int[] 列表，每项为 [start, end]。
        /// </summary>
        public static List<int[]> ParseRanges(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                throw new FormatException("页码范围不能为空");

            var ranges = new List<int[]>();
            // 统一中文逗号为英文逗号
            input = input.Replace('，', ',').Replace('；', ',').Replace(';', ',');

            var segments = input.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (segments.Length == 0)
                throw new FormatException("页码范围格式无效");

            foreach (var seg in segments)
            {
                var trimmed = seg.Trim();
                if (trimmed.Contains("-"))
                {
                    var parts = trimmed.Split('-');
                    if (parts.Length != 2)
                        throw new FormatException($"无效的范围: \"{trimmed}\"");

                    if (!int.TryParse(parts[0].Trim(), out int start) ||
                        !int.TryParse(parts[1].Trim(), out int end))
                        throw new FormatException($"无效的范围: \"{trimmed}\"");

                    if (start < 1 || end < 1 || start > end)
                        throw new FormatException($"无效的范围: \"{trimmed}\"（起始页必须小于等于结束页，且大于0）");

                    ranges.Add(new[] { start, end });
                }
                else
                {
                    if (!int.TryParse(trimmed, out int page))
                        throw new FormatException($"无效的页码: \"{trimmed}\"");
                    if (page < 1)
                        throw new FormatException($"页码必须大于0: \"{trimmed}\"");
                    ranges.Add(new[] { page, page });
                }
            }

            return ranges;
        }

        public void Dispose()
        {
        }
    }
}