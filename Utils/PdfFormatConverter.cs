using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Spire.Pdf;
using Spire.Pdf.Conversion;

namespace DocToPdfTool.Utils
{
    /// <summary>PDF 可转换的目标格式。</summary>
    public enum PdfTargetFormat
    {
        /// <summary>网页（HTML），保留版式与图片。</summary>
        Html = 0,

        /// <summary>矢量 SVG，每页一个文件。</summary>
        Svg = 1,

        /// <summary>Markdown 文本，便于二次编辑。</summary>
        Markdown = 2,

        /// <summary>灰度（黑白），体积通常明显变小。</summary>
        Gray = 3,

        /// <summary>PDF/A-2B 归档格式，长期保存用。</summary>
        PdfA = 4,

        /// <summary>XPS 固定版式文档。</summary>
        Xps = 5
    }

    /// <summary>单次转换结果。</summary>
    public class PdfFormatConvertResult
    {
        public string SourceFile { get; set; }

        /// <summary>输出文件路径；SVG 为输出目录里的首个文件，主要用于显示。</summary>
        public string OutputPath { get; set; }

        /// <summary>输出目录（SVG 会生成多个文件）。</summary>
        public string OutputDir { get; set; }

        public long OriginalSize { get; set; }
        public long NewSize { get; set; }

        /// <summary>输出文件数量（SVG 多页时为页数）。</summary>
        public int OutputCount { get; set; } = 1;

        /// <summary>输出文件扩展名，用于提示。</summary>
        public string Extension { get; set; }

        public TimeSpan Elapsed { get; set; }
    }

    /// <summary>
    /// PDF → 其他格式转换（HTML / SVG / Markdown / 灰度 / PDF/A / XPS）。
    /// 全部基于随程序分发的 Spire.Pdf，无外部依赖、无需 Office。
    /// </summary>
    public class PdfFormatConverter : IDisposable
    {
        /// <summary>各格式的中文名，供 UI 与日志使用。</summary>
        public static string GetFormatName(PdfTargetFormat format)
        {
            switch (format)
            {
                case PdfTargetFormat.Html: return "HTML 网页";
                case PdfTargetFormat.Svg: return "SVG 矢量图";
                case PdfTargetFormat.Markdown: return "Markdown 文本";
                case PdfTargetFormat.Gray: return "灰度 PDF";
                case PdfTargetFormat.PdfA: return "PDF/A 归档";
                case PdfTargetFormat.Xps: return "XPS 文档";
                default: return format.ToString();
            }
        }

        /// <summary>该格式会产出多个文件（按页拆分）。</summary>
        public static bool IsMultiFile(PdfTargetFormat format)
            => format == PdfTargetFormat.Svg;

        /// <summary>
        /// 转换单个 PDF。
        /// </summary>
        /// <param name="sourcePath">源 PDF 路径。</param>
        /// <param name="outputDir">输出根目录（本方法会在其中建子目录）。</param>
        /// <param name="format">目标格式。</param>
        public PdfFormatConvertResult Convert(string sourcePath, string outputDir, PdfTargetFormat format)
        {
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException("找不到源文件", sourcePath);

            var startedAt = DateTime.Now;

            string baseName = Path.GetFileNameWithoutExtension(sourcePath);
            string safeName = SanitizeFileName(baseName);

            // 直接输出到功能目录；同名时给文件名加序号，不单独建子目录
            Directory.CreateDirectory(outputDir);
            string targetDir = outputDir;
            safeName = MakeUniqueBaseName(targetDir, safeName, format);

            var result = new PdfFormatConvertResult
            {
                SourceFile = sourcePath,
                OutputDir = targetDir,
                OriginalSize = new FileInfo(sourcePath).Length
            };

            switch (format)
            {
                case PdfTargetFormat.Html:
                    result.OutputPath = ConvertToHtml(sourcePath, targetDir, safeName);
                    result.Extension = ".html";
                    break;

                case PdfTargetFormat.Svg:
                    result.OutputPath = ConvertToSvg(sourcePath, targetDir, safeName, result);
                    result.Extension = ".svg";
                    break;

                case PdfTargetFormat.Markdown:
                    result.OutputPath = ConvertToMarkdown(sourcePath, targetDir, safeName);
                    result.Extension = ".md";
                    break;

                case PdfTargetFormat.Gray:
                    result.OutputPath = ConvertToGray(sourcePath, targetDir, safeName);
                    result.Extension = ".pdf";
                    break;

                case PdfTargetFormat.PdfA:
                    result.OutputPath = ConvertToPdfA(sourcePath, targetDir, safeName);
                    result.Extension = ".pdf";
                    break;

                case PdfTargetFormat.Xps:
                    result.OutputPath = ConvertToXps(sourcePath, targetDir, safeName);
                    result.Extension = ".xps";
                    break;

                default:
                    throw new NotSupportedException("暂不支持的目标格式：" + format);
            }

            if (format == PdfTargetFormat.Svg)
            {
                var files = Directory.GetFiles(targetDir, "*.svg");
                result.NewSize = files.Sum(f => new FileInfo(f).Length);
            }
            else
            {
                result.NewSize = File.Exists(result.OutputPath) ? new FileInfo(result.OutputPath).Length : 0;
            }

            result.Elapsed = DateTime.Now - startedAt;
            return result;
        }

        // ───────────────── 各格式实现 ─────────────────

        private static string ConvertToHtml(string source, string dir, string safeName)
        {
            string output = Path.Combine(dir, safeName + ".html");

            using (var doc = new PdfDocument())
            {
                doc.LoadFromFile(source);
                doc.SaveToFile(output, FileFormat.HTML);
                doc.Close();
            }
            return output;
        }

        private static string ConvertToSvg(string source, string dir, string safeName, PdfFormatConvertResult result)
        {
            // Spire 的语义：传入「目标文件名」，实际按页输出到同目录下的
            // 「<文件名去扩展名>_1.svg」「_2.svg」…
            string probe = Path.Combine(dir, safeName + ".svg");

            var converter = new PdfToSvgConverter(source);
            converter.Convert(probe);

            var files = Directory.GetFiles(dir, safeName + "_*.svg")
                                 .OrderBy(f => ExtractPageNumber(f))
                                 .ToList();

            result.OutputCount = files.Count;
            return files.FirstOrDefault() ?? probe;
        }

        private static int ExtractPageNumber(string path)
        {
            var name = Path.GetFileNameWithoutExtension(path);
            int idx = name.LastIndexOf('_');
            if (idx < 0) return 0;
            return int.TryParse(name.Substring(idx + 1), out int n) ? n : 0;
        }

        private static string ConvertToMarkdown(string source, string dir, string safeName)
        {
            string output = Path.Combine(dir, safeName + ".md");

            var converter = new PdfToMarkdownConverter(source);
            converter.ConvertToMarkdown(output);
            converter.Dispose();
            return output;
        }

        private static string ConvertToGray(string source, string dir, string safeName)
        {
            string output = Path.Combine(dir, safeName + "_灰度.pdf");

            var converter = new PdfGrayConverter(source);
            converter.ToGrayPdf(output);
            converter.Dispose();
            return output;
        }

        private static string ConvertToPdfA(string source, string dir, string safeName)
        {
            string output = Path.Combine(dir, safeName + "_PDFA.pdf");

            var converter = new PdfStandardsConverter(source);
            converter.ToPdfA2B(output);
            return output;
        }

        private static string ConvertToXps(string source, string dir, string safeName)
        {
            string output = Path.Combine(dir, safeName + ".xps");

            using (var doc = new PdfDocument())
            {
                doc.LoadFromFile(source);
                doc.SaveToFile(output, FileFormat.XPS);
                doc.Close();
            }
            return output;
        }

        // ───────────────── 辅助 ─────────────────

        private static string GetFolderSuffix(PdfTargetFormat format)
        {
            switch (format)
            {
                case PdfTargetFormat.Html: return "HTML";
                case PdfTargetFormat.Svg: return "SVG";
                case PdfTargetFormat.Markdown: return "Markdown";
                case PdfTargetFormat.Gray: return "灰度";
                case PdfTargetFormat.PdfA: return "PDFA";
                case PdfTargetFormat.Xps: return "XPS";
                default: return "输出";
            }
        }

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "output";

            var invalid = Path.GetInvalidFileNameChars();
            var chars = name.Select(c => invalid.Contains(c) ? '_' : c).ToArray();
            var cleaned = new string(chars).Trim();

            foreach (var c in new[] { '.', ' ' })
                cleaned = cleaned.Trim(c);

            if (cleaned.Length == 0) cleaned = "output";
            if (cleaned.Length > 60) cleaned = cleaned.Substring(0, 60);
            return cleaned;
        }

        /// <summary>
        /// 在同一个输出目录里为本次转换取一个不冲突的基础名。
        /// 已存在「名字.html」或「名字_1.svg」这类同名产物时，依次尝试「名字(2)」「名字(3)」…
        /// </summary>
        private static string MakeUniqueBaseName(string dir, string baseName, PdfTargetFormat format)
        {
            if (string.IsNullOrEmpty(baseName)) baseName = "output";
            if (!Directory.Exists(dir)) return baseName;

            // 该格式会产生的文件名特征
            Func<string, bool> exists = candidate =>
            {
                switch (format)
                {
                    case PdfTargetFormat.Svg:
                        return Directory.GetFiles(dir, candidate + "_*.svg").Length > 0;
                    default:
                        {
                            var ext = GetResultExtension(format);
                            return File.Exists(Path.Combine(dir, candidate + ext));
                        }
                }
            };

            if (!exists(baseName)) return baseName;

            for (int i = 2; i < 1000; i++)
            {
                string candidate = $"{baseName}({i})";
                if (!exists(candidate)) return candidate;
            }
            return baseName + "(" + Guid.NewGuid().ToString("N").Substring(0, 6) + ")";
        }

        /// <summary>各格式主产物的扩展名（用于防重名判断）。</summary>
        private static string GetResultExtension(PdfTargetFormat format)
        {
            switch (format)
            {
                case PdfTargetFormat.Html: return ".html";
                case PdfTargetFormat.Svg: return "_1.svg";
                case PdfTargetFormat.Markdown: return ".md";
                case PdfTargetFormat.Gray: return "_灰度.pdf";
                case PdfTargetFormat.PdfA: return "_PDFA.pdf";
                case PdfTargetFormat.Xps: return ".xps";
                default: return ".bin";
            }
        }

        public void Dispose() { }
    }
}
