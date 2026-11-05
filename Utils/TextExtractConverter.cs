using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Presentation;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using SpireDocDocument = Spire.Doc.Document;
using SpirePresentation = Spire.Presentation.Presentation;

namespace DocToPdfTool.Utils
{
    /// <summary>输出文本编码。</summary>
    public enum TextOutputEncoding
    {
        Utf8 = 0,
        Utf8Bom = 1,
        Gbk = 2
    }

    /// <summary>纯文本提取参数。</summary>
    public class TextExtractOptions
    {
        public TextOutputEncoding Encoding { get; set; } = TextOutputEncoding.Utf8Bom;

        /// <summary>剔除跨页重复的页眉页脚（含纯页码行）。</summary>
        public bool RemoveRepeatedHeaders { get; set; } = true;

        /// <summary>把连续空行压缩为一行。</summary>
        public bool CollapseBlankLines { get; set; } = true;

        /// <summary>页与页之间插入分隔标记。</summary>
        public bool PageSeparator { get; set; } = true;
    }

    /// <summary>单个文件的提取结果。</summary>
    public class TextExtractResult
    {
        public string OutputFile { get; set; }
        public int CharCount { get; set; }
        public int PageCount { get; set; }
        public string Note { get; set; }
        public bool Empty { get; set; }
    }

    /// <summary>
    /// 纯文本提取器：PDF / Word / PPT / TXT / MD → TXT。
    ///
    /// 要点：
    ///   - PDF 按「页 → Y 坐标 → X 坐标」重排，避免多栏排版顺序错乱；
    ///   - 中文与英文相邻时智能补空格：中文之间不加空格，西文之间保留空格；
    ///   - 跨页重复出现的短行（页眉页脚）与纯页码行会被自动剔除；
    ///   - 扫描件 PDF 无文字层时给出明确提示，而不是静默输出空文件。
    /// </summary>
    public class TextExtractConverter : IDisposable
    {
        /// <summary>行尾页码片段，如「第 4 页 / 共 27 页」「第2页」。</summary>
        private static readonly Regex PageMarkerTailRegex = new Regex(
            @"\s*第\s*\d+\s*页(\s*[/／]\s*共?\s*\d+\s*页?)?\s*$",
            RegexOptions.Compiled);

        private static readonly Regex PageNumberRegex = new Regex(
            @"^\s*(第\s*\d+\s*页(\s*[/／]\s*共?\s*\d+\s*页?)?|\d+\s*[/／]\s*\d+|-\s*\d+\s*-|page\s*\d+(\s*of\s*\d+)?)\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>提取单个文件的文本并写出 txt。</summary>
        public TextExtractResult Extract(string sourcePath, string outputDir, TextExtractOptions options)
        {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("文件不存在", sourcePath);
            options = options ?? new TextExtractOptions();

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            var extension = Path.GetExtension(sourcePath).ToLowerInvariant();
            var result = new TextExtractResult();

            string text;
            switch (extension)
            {
                case ".pdf":
                    text = ExtractPdf(sourcePath, options, result);
                    break;
                case ".doc":
                case ".docx":
                case ".rtf":
                    text = ExtractWord(sourcePath, result);
                    break;
                case ".ppt":
                case ".pptx":
                    text = ExtractPpt(sourcePath, result);
                    break;
                case ".txt":
                case ".md":
                    text = ReadPlainText(sourcePath);
                    break;
                default:
                    throw new NotSupportedException("暂不支持该格式：" + extension);
            }

            text = Normalize(text, options);

            var outputFile = UniquePath(outputDir, Path.GetFileNameWithoutExtension(sourcePath), ".txt");
            WriteText(outputFile, text, options.Encoding);

            result.OutputFile = outputFile;
            result.CharCount = text.Length;
            result.Empty = string.IsNullOrWhiteSpace(text);

            if (result.Empty && string.IsNullOrEmpty(result.Note))
                result.Note = extension == ".pdf"
                    ? "未提取到文字（可能是扫描件，需要 OCR）"
                    : "未提取到文字";

            return result;
        }

        // ───────────────── PDF ─────────────────

        private static string ExtractPdf(string path, TextExtractOptions options, TextExtractResult result)
        {
            if (PdfEncryptConverter.IsEncrypted(path))
                throw new InvalidOperationException("文件已加密，请先用「PDF加密 → 移除密码」解除后再提取");

            // 每页一组行：行 = (Y, 文本)
            var pages = new List<List<(double Y, string Text)>>();

            using (var doc = PdfDocument.Open(path))
            {
                result.PageCount = doc.NumberOfPages;

                for (int i = 1; i <= doc.NumberOfPages; i++)
                {
                    List<(double Y, string Text)> lines;
                    try
                    {
                        lines = BuildPageLines(doc.GetPage(i));
                    }
                    catch
                    {
                        lines = new List<(double Y, string Text)>();
                    }
                    pages.Add(lines);
                }
            }

            // 统计跨页重复行（页眉页脚候选）
            var repeated = new HashSet<string>();
            if (options.RemoveRepeatedHeaders && pages.Count >= 3)
            {
                var counter = new Dictionary<string, int>();
                foreach (var page in pages)
                {
                    // 只看每页首尾各 3 行
                    var candidates = page.Take(3).Concat(page.Skip(Math.Max(0, page.Count - 3)));
                    foreach (var line in candidates)
                    {
                        var key = NormalizeLine(line.Text);
                        if (key.Length == 0 || key.Length > 60) continue;
                        counter[key] = counter.TryGetValue(key, out int c) ? c + 1 : 1;
                    }
                }

                int threshold = Math.Max(3, (int)Math.Ceiling(pages.Count * 0.6));
                foreach (var pair in counter)
                {
                    if (pair.Value >= threshold) repeated.Add(pair.Key);
                }
            }

            var sb = new StringBuilder();
            for (int i = 0; i < pages.Count; i++)
            {
                var kept = new List<string>();

                foreach (var line in pages[i])
                {
                    var text = line.Text ?? string.Empty;

                    if (options.RemoveRepeatedHeaders)
                    {
                        // 先剥掉行尾的「第 N 页 / 共 M 页」片段
                        text = PageMarkerTailRegex.Replace(text, string.Empty).TrimEnd();
                        if (text.Length == 0) continue;
                        if (PageNumberRegex.IsMatch(text)) continue;
                        if (repeated.Contains(NormalizeLine(text))) continue;
                    }

                    kept.Add(text);
                }

                if (kept.Count == 0) continue;

                if (options.PageSeparator && pages.Count > 1)
                    sb.AppendLine($"===== 第 {i + 1} 页 =====");

                foreach (var line in kept) sb.AppendLine(line);
            }

            return sb.ToString();
        }

        /// <summary>把一页的文字按坐标重排成行。</summary>
        private static List<(double Y, string Text)> BuildPageLines(Page page)
        {
            var words = page.GetWords().ToList();
            if (words.Count == 0) return new List<(double Y, string Text)>();

            // 按 Y 从上到下（PDF 坐标 y 越大越靠上），同一行内按 X 从左到右
            var sorted = words
                .OrderByDescending(w => w.BoundingBox.Bottom)
                .ThenBy(w => w.BoundingBox.Left)
                .ToList();

            var lines = new List<(double Y, List<Word> Words)>();
            foreach (var word in sorted)
            {
                double y = word.BoundingBox.Bottom;
                var line = lines.FirstOrDefault(l => Math.Abs(l.Y - y) <= 2.5);
                if (line.Words == null)
                {
                    line = (y, new List<Word>());
                    lines.Add(line);
                }
                line.Words.Add(word);
            }

            var result = new List<(double Y, string Text)>();
            foreach (var line in lines.OrderByDescending(l => l.Y))
            {
                var ordered = line.Words.OrderBy(w => w.BoundingBox.Left).ToList();
                result.Add((line.Y, JoinWords(ordered)));
            }
            return result;
        }

        /// <summary>拼接一行内的单词：中文之间不留空格，西文之间补空格。</summary>
        private static string JoinWords(List<Word> words)
        {
            var sb = new StringBuilder();
            Word previous = null;

            foreach (var word in words)
            {
                if (previous != null)
                {
                    double gap = word.BoundingBox.Left - previous.BoundingBox.Right;
                    char prevLast = sb.Length > 0 ? sb[sb.Length - 1] : ' ';
                    char curFirst = word.Text.Length > 0 ? word.Text[0] : ' ';

                    bool needSpace = gap > 1.0;
                    if (IsCjk(prevLast) || IsCjk(curFirst)) needSpace = gap > 6.0;

                    if (needSpace && sb.Length > 0 && sb[sb.Length - 1] != ' ') sb.Append(' ');
                }
                sb.Append(word.Text);
                previous = word;
            }

            return sb.ToString();
        }

        private static bool IsCjk(char c)
        {
            return (c >= 0x4E00 && c <= 0x9FFF)     // 汉字
                || (c >= 0x3000 && c <= 0x303F)     // 中文标点
                || (c >= 0xFF00 && c <= 0xFFEF);    // 全角
        }

        // ───────────────── Word / PPT / TXT ─────────────────

        private static string ExtractWord(string path, TextExtractResult result)
        {
            using (var doc = new SpireDocDocument())
            {
                doc.LoadFromFile(path);
                try { result.PageCount = doc.GetPageCount(); } catch { }
                // GetText 已包含正文段落与表格单元格文字，按行返回
                return doc.GetText() ?? string.Empty;
            }
        }

        private static string ExtractPpt(string path, TextExtractResult result)
        {
            var sb = new StringBuilder();
            using (var ppt = new SpirePresentation())
            {
                ppt.LoadFromFile(path);
                result.PageCount = ppt.Slides.Count;

                for (int i = 0; i < ppt.Slides.Count; i++)
                {
                    var slide = ppt.Slides[i];
                    sb.AppendLine($"===== 第 {i + 1} 页 =====");

                    foreach (IShape shape in slide.Shapes)
                    {
                        ExtractShapeText(shape, sb);
                    }
                }
            }
            return sb.ToString();
        }

        private static void ExtractShapeText(IShape shape, StringBuilder sb)
        {
            try
            {
                if (shape is ITable table)
                {
                    foreach (var rowObj in table.TableRows)
                    {
                        if (!(rowObj is Spire.Presentation.TableRow row)) continue;
                        var cells = new List<string>();
                        for (int i = 0; i < row.Count; i++)
                        {
                            var cellText = row[i]?.TextFrame?.Text ?? string.Empty;
                            cells.Add(cellText.Replace("\r", " ").Replace("\n", " ").Trim());
                        }
                        sb.AppendLine(string.Join("\t", cells));
                    }
                    return;
                }

                if (shape is IAutoShape autoShape && autoShape.TextFrame != null)
                {
                    var text = autoShape.TextFrame.Text;
                    if (!string.IsNullOrWhiteSpace(text)) sb.AppendLine(text.TrimEnd());
                    return;
                }

            }
            catch
            {
                // 单个形状失败不影响整体
            }
        }

        private static string ReadPlainText(string path)
        {
            var bytes = File.ReadAllBytes(path);

            // BOM 识别
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
                return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
                return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);

            // 尝试 UTF-8 严格解码，失败则按 GBK
            try
            {
                var strict = new UTF8Encoding(false, true);
                return strict.GetString(bytes);
            }
            catch
            {
                return Encoding.GetEncoding(936).GetString(bytes);
            }
        }

        // ───────────────── 文本清理与写出 ─────────────────

        private static string Normalize(string text, TextExtractOptions options)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            text = text.Replace("\r\n", "\n").Replace('\r', '\n');

            if (options.CollapseBlankLines)
            {
                text = Regex.Replace(text, @"\n{3,}", "\n\n");
            }

            return text.Trim();
        }

        private static string NormalizeLine(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            return Regex.Replace(text.Trim(), @"\s+", " ");
        }

        private static void WriteText(string path, string text, TextOutputEncoding encoding)
        {
            switch (encoding)
            {
                case TextOutputEncoding.Utf8:
                    File.WriteAllText(path, text, new UTF8Encoding(false));
                    break;
                case TextOutputEncoding.Gbk:
                    File.WriteAllText(path, text, Encoding.GetEncoding(936));
                    break;
                default:
                    File.WriteAllText(path, text, new UTF8Encoding(true));
                    break;
            }
        }

        /// <summary>格式化体积显示。</summary>
        public static string FormatSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024.0).ToString("0.0") + " KB";
            return (bytes / 1024.0 / 1024.0).ToString("0.00") + " MB";
        }

        private static string UniquePath(string dir, string baseName, string extension)
        {
            var path = Path.Combine(dir, baseName + extension);
            int seq = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(dir, $"{baseName}_{seq++}{extension}");
            }
            return path;
        }

        public void Dispose() { }
    }
}
