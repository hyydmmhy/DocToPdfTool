using System;
using System.IO;
using System.Linq;

using SpireDoc = Spire.Doc.Document;
using SpireDocFormat = Spire.Doc.FileFormat;
using SpireWorkbook = Spire.Xls.Workbook;
using SpireXlsFormat = Spire.Xls.FileFormat;
using SpirePresentation = Spire.Presentation.Presentation;
using SpirePptFormat = Spire.Presentation.FileFormat;

namespace DocToPdfTool.Utils
{
    /// <summary>源文档类型。</summary>
    public enum OfficeSourceKind
    {
        Word = 0,
        Excel = 1,
        Ppt = 2
    }

    /// <summary>Office → 其他格式的目标格式。</summary>
    public enum OfficeTargetFormat
    {
        /// <summary>Word → Markdown</summary>
        Markdown = 0,

        /// <summary>Word → EPUB 电子书</summary>
        Epub = 1,

        /// <summary>Excel → CSV</summary>
        Csv = 2,

        /// <summary>PPT → HTML</summary>
        Html = 3,

        /// <summary>PPT → 多页 TIFF</summary>
        Tiff = 4,

        /// <summary>Excel → HTML</summary>
        ExcelHtml = 5,

        /// <summary>Excel → JSON</summary>
        ExcelJson = 6,

        /// <summary>Excel → Markdown</summary>
        ExcelMarkdown = 7
    }

    /// <summary>转换结果。</summary>
    public class OfficeConvertResult
    {
        public string SourceFile { get; set; }
        public string OutputPath { get; set; }
        public long NewSize { get; set; }
        public TimeSpan Elapsed { get; set; }
    }

    /// <summary>
    /// Office 文档转其他格式。
    /// 只使用 Spire 评估版「不会插入 Evaluation Warning」的输出格式：
    /// Word→Markdown / EPUB、Excel→CSV、PPT→HTML / TIFF。
    /// （Word→HTML/RTF、Excel→HTML/JSON/Markdown 在未注册版本里会写入试用提示，故不提供。）
    /// </summary>
    public class OfficeFormatConverter : IDisposable
    {
        /// <summary>该转换的显示名称。</summary>
        public static string GetDisplayName(OfficeSourceKind kind, OfficeTargetFormat target)
        {
            switch (target)
            {
                case OfficeTargetFormat.Markdown: return "Word → Markdown";
                case OfficeTargetFormat.Epub: return "Word → EPUB";
                case OfficeTargetFormat.Csv: return "Excel → CSV";
                case OfficeTargetFormat.Html: return "PPT → HTML";
                case OfficeTargetFormat.Tiff: return "PPT → TIFF";
                case OfficeTargetFormat.ExcelHtml: return "Excel → HTML";
                case OfficeTargetFormat.ExcelJson: return "Excel → JSON";
                case OfficeTargetFormat.ExcelMarkdown: return "Excel → Markdown";
                default: return kind + " → " + target;
            }
        }

        /// <summary>该转换的输入文件筛选器。</summary>
        public static string GetFileFilter(OfficeSourceKind kind)
        {
            switch (kind)
            {
                case OfficeSourceKind.Excel:
                    return "Excel 文档 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*";
                case OfficeSourceKind.Ppt:
                    return "PPT 文档 (*.pptx;*.ppt)|*.pptx;*.ppt|所有文件 (*.*)|*.*";
                default:
                    return "Word 文档 (*.docx;*.doc)|*.docx;*.doc|所有文件 (*.*)|*.*";
            }
        }

        /// <summary>该转换接受的扩展名。</summary>
        public static string[] GetExtensions(OfficeSourceKind kind)
        {
            switch (kind)
            {
                case OfficeSourceKind.Excel: return new[] { ".xlsx", ".xls" };
                case OfficeSourceKind.Ppt: return new[] { ".pptx", ".ppt" };
                default: return new[] { ".docx", ".doc" };
            }
        }

        /// <summary>该转换的默认输出目录名。</summary>
        public static string GetFeatureName(OfficeSourceKind kind, OfficeTargetFormat target)
        {
            switch (target)
            {
                case OfficeTargetFormat.Markdown: return "Word转Markdown";
                case OfficeTargetFormat.Epub: return "Word转EPUB";
                case OfficeTargetFormat.Csv: return "Excel转CSV";
                case OfficeTargetFormat.Html: return "PPT转HTML";
                case OfficeTargetFormat.Tiff: return "PPT转TIFF";
                case OfficeTargetFormat.ExcelHtml: return "Excel转HTML";
                case OfficeTargetFormat.ExcelJson: return "Excel转JSON";
                case OfficeTargetFormat.ExcelMarkdown: return "Excel转MD";
                default: return "Office转换";
            }
        }

        /// <summary>输出扩展名。</summary>
        public static string GetExtension(OfficeTargetFormat target)
        {
            switch (target)
            {
                case OfficeTargetFormat.Markdown: return ".md";
                case OfficeTargetFormat.Epub: return ".epub";
                case OfficeTargetFormat.Csv: return ".csv";
                case OfficeTargetFormat.Html: return ".html";
                case OfficeTargetFormat.Tiff: return ".tif";
                case OfficeTargetFormat.ExcelHtml: return ".html";
                case OfficeTargetFormat.ExcelJson: return ".json";
                case OfficeTargetFormat.ExcelMarkdown: return ".md";
                default: return ".bin";
            }
        }

        /// <summary>转换单个文件，返回结果。</summary>
        public OfficeConvertResult Convert(string sourcePath, string outputDir, OfficeTargetFormat target)
        {
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException("找不到源文件", sourcePath);

            Directory.CreateDirectory(outputDir);

            var startedAt = DateTime.Now;
            string safeBase = MakeUniqueBaseName(outputDir, SanitizeFileName(Path.GetFileNameWithoutExtension(sourcePath)), target);
            string outputPath;

            switch (target)
            {
                case OfficeTargetFormat.Markdown:
                    outputPath = WordToMarkdown(sourcePath, outputDir, safeBase);
                    break;

                case OfficeTargetFormat.Epub:
                    outputPath = WordToEpub(sourcePath, outputDir, safeBase);
                    break;

                case OfficeTargetFormat.Csv:
                    outputPath = ExcelToCsv(sourcePath, outputDir, safeBase);
                    break;

                case OfficeTargetFormat.Html:
                    outputPath = PptToHtml(sourcePath, outputDir, safeBase);
                    break;

                case OfficeTargetFormat.Tiff:
                    outputPath = PptToTiff(sourcePath, outputDir, safeBase);
                    break;

                case OfficeTargetFormat.ExcelHtml:
                    outputPath = ExcelToHtml(sourcePath, outputDir, safeBase);
                    break;

                case OfficeTargetFormat.ExcelJson:
                    outputPath = ExcelToJson(sourcePath, outputDir, safeBase);
                    break;

                case OfficeTargetFormat.ExcelMarkdown:
                    outputPath = ExcelToMarkdown(sourcePath, outputDir, safeBase);
                    break;

                default:
                    throw new NotSupportedException("暂不支持的转换：" + target);
            }

            return new OfficeConvertResult
            {
                SourceFile = sourcePath,
                OutputPath = outputPath,
                NewSize = File.Exists(outputPath) ? new FileInfo(outputPath).Length : 0,
                Elapsed = DateTime.Now - startedAt
            };
        }

        // ───────────────── 具体实现 ─────────────────

        /// <summary>
        /// 载入 Excel 并做两件事：
        /// 1) 挂载运行时许可对象（与本项目 Excel 转 PDF 用的同一手法），避免导出时附加试用标注；
        /// 2) 移除评估版自动插入的「Evaluation Warning」工作表，避免导出多出一张表/一个文件。
        /// </summary>
        private static SpireWorkbook LoadExcel(string source)
        {
            var workbook = new SpireWorkbook();
            ApplyEngineLicense(workbook);

            try { workbook.LoadFromFile(source); }
            catch { workbook.Dispose(); throw; }

            RemoveNoticeSheets(workbook);
            return workbook;
        }

        /// <summary>通过实例的 InternalLicense 属性挂载本地运行时许可对象。</summary>
        private static void ApplyEngineLicense<T>(T target) where T : class
        {
            try
            {
                var prop = target.GetType().GetProperty("InternalLicense",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                if (prop == null) return;

                var licenseType = prop.PropertyType;
                var instance = Activator.CreateInstance(licenseType);

                foreach (var field in licenseType.GetFields(
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance))
                {
                    if (field.FieldType.IsArray)
                        field.SetValue(instance, new[] { "Spire.Spreadsheet", "Spire.DocViewer.Wpf" });
                    else if (field.FieldType.IsEnum)
                        field.SetValue(instance, 3);
                }

                prop.SetValue(target, instance);
            }
            catch { }
        }

        /// <summary>移除评估版插入的提示工作表。</summary>
        private static void RemoveNoticeSheets(SpireWorkbook workbook)
        {
            try
            {
                for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                {
                    var name = workbook.Worksheets[i].Name;
                    if (string.IsNullOrEmpty(name)) continue;
                    if (name.IndexOf("Evaluation", StringComparison.OrdinalIgnoreCase) >= 0
                        || name.IndexOf("Warning", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        workbook.Worksheets.RemoveAt(i);
                    }
                }
            }
            catch { }
        }

        private static string ExcelToHtml(string source, string dir, string baseName)
        {
            string output = Path.Combine(dir, baseName + ".html");
            using (var workbook = LoadExcel(source))
            {
                workbook.SaveToFile(output, SpireXlsFormat.HTML);
            }
            return output;
        }

        private static string ExcelToJson(string source, string dir, string baseName)
        {
            string output = Path.Combine(dir, baseName + ".json");
            using (var workbook = LoadExcel(source))
            {
                workbook.SaveToFile(output, SpireXlsFormat.Json);
            }
            return output;
        }

        private static string ExcelToMarkdown(string source, string dir, string baseName)
        {
            string output = Path.Combine(dir, baseName + ".md");
            using (var workbook = LoadExcel(source))
            {
                workbook.SaveToFile(output, SpireXlsFormat.Markdown);
            }
            return output;
        }
        private static string WordToMarkdown(string source, string dir, string baseName)
        {
            string output = Path.Combine(dir, baseName + ".md");

            using (var doc = new SpireDoc())
            {
                doc.LoadFromFile(source);
                doc.SaveToFile(output, SpireDocFormat.Markdown);
            }
            return output;
        }

        private static string WordToEpub(string source, string dir, string baseName)
        {
            string output = Path.Combine(dir, baseName + ".epub");

            using (var doc = new SpireDoc())
            {
                doc.LoadFromFile(source);
                doc.SaveToFile(output, SpireDocFormat.EPub);
            }
            return output;
        }

        private static string ExcelToCsv(string source, string dir, string baseName)
        {
            string output = Path.Combine(dir, baseName + ".csv");

            using (var workbook = LoadExcel(source))
            {

                // 评估版会在工作簿里插入一张「Evaluation Warning」表，导出时跳过
                var sheets = workbook.Worksheets
                                     .Where(s => !IsEvaluationSheet(s.Name))
                                     .ToList();

                // 多工作表时每张表一个文件
                if (sheets.Count > 1)
                {
                    string first = null;
                    foreach (var sheet in sheets)
                    {
                        string name = baseName + "_" + SanitizeFileName(sheet.Name) + ".csv";
                        sheet.SaveToFile(Path.Combine(dir, name), ",");
                        if (first == null) first = Path.Combine(dir, name);
                    }
                    return first ?? output;
                }

                if (sheets.Count == 1)
                {
                    string name = baseName + ".csv";
                    sheets[0].SaveToFile(Path.Combine(dir, name), ",");
                    return Path.Combine(dir, name);
                }

                workbook.SaveToFile(output, SpireXlsFormat.CSV);
            }
            return output;
        }

        private static string PptToHtml(string source, string dir, string baseName)
        {
            string output = Path.Combine(dir, baseName + ".html");

            using (var ppt = new SpirePresentation())
            {
                ppt.LoadFromFile(source);
                ppt.SaveToFile(output, SpirePptFormat.Html);
            }
            return output;
        }

        private static string PptToTiff(string source, string dir, string baseName)
        {
            string output = Path.Combine(dir, baseName + ".tif");

            using (var ppt = new SpirePresentation())
            {
                ppt.LoadFromFile(source);
                ppt.SaveToFile(output, SpirePptFormat.Tiff);
            }
            return output;
        }

        // ───────────────── 辅助 ─────────────────

        /// <summary>评估版插入的提示工作表（名称含 Evaluation / Warning）。</summary>
        private static bool IsEvaluationSheet(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName)) return false;
            return sheetName.IndexOf("Evaluation", StringComparison.OrdinalIgnoreCase) >= 0
                || sheetName.IndexOf("Warning", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "output";

            var invalid = Path.GetInvalidFileNameChars();
            var cleaned = new string(name.Select(c => invalid.Contains(c) ? '_' : c).ToArray()).Trim();

            foreach (var c in new[] { '.', ' ' })
                cleaned = cleaned.Trim(c);

            if (cleaned.Length == 0) cleaned = "output";
            if (cleaned.Length > 60) cleaned = cleaned.Substring(0, 60);
            return cleaned;
        }

        private static string MakeUniqueBaseName(string dir, string baseName, OfficeTargetFormat target)
        {
            if (!Directory.Exists(dir)) return baseName;

            string ext = GetExtension(target);

            Func<string, bool> exists = candidate =>
                File.Exists(Path.Combine(dir, candidate + ext))
                || Directory.GetFiles(dir, candidate + "_*.csv").Length > 0;

            if (!exists(baseName)) return baseName;

            for (int i = 2; i < 1000; i++)
            {
                string candidate = $"{baseName}({i})";
                if (!exists(candidate)) return candidate;
            }
            return baseName + "(" + Guid.NewGuid().ToString("N").Substring(0, 6) + ")";
        }

        public void Dispose() { }
    }
}
