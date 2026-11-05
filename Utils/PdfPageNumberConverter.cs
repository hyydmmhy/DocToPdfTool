using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using Spire.Pdf;
using Spire.Pdf.Graphics;

namespace DocToPdfTool.Utils
{
    /// <summary>页码位置。</summary>
    public enum PageNumberPosition
    {
        BottomCenter = 0,
        BottomLeft = 1,
        BottomRight = 2,
        TopCenter = 3,
        TopLeft = 4,
        TopRight = 5
    }

    /// <summary>页码格式。</summary>
    public enum PageNumberFormat
    {
        /// <summary>1</summary>
        Plain = 0,

        /// <summary>第 1 页</summary>
        WithPrefix = 1,

        /// <summary>1 / 20</summary>
        WithTotal = 2,

        /// <summary>- 1 -</summary>
        Dash = 3
    }

    /// <summary>PDF 加页码参数。</summary>
    public class PdfPageNumberOptions
    {
        public PageNumberPosition Position { get; set; } = PageNumberPosition.BottomCenter;

        public PageNumberFormat Format { get; set; } = PageNumberFormat.Plain;

        /// <summary>第一个被编号页面显示的页码。</summary>
        public int StartNumber { get; set; } = 1;

        /// <summary>跳过前 N 页（封面等），这些页不绘制也不占用编号。</summary>
        public int SkipFirstPages { get; set; }

        /// <summary>只对指定页编号，格式同「1-3,5」；为空表示全部页。</summary>
        public string PageRanges { get; set; }

        /// <summary>字号（磅）。</summary>
        public float FontSize { get; set; } = 10f;

        /// <summary>文字颜色。</summary>
        public Color Color { get; set; } = Color.Black;

        /// <summary>距页面边缘的距离（磅，1 磅约 0.35 mm）。</summary>
        public float MarginPt { get; set; } = 28f;
    }

    /// <summary>
    /// PDF 加页码转换器。
    ///
    /// 规则：
    ///   - 「跳过前 N 页」的页面既不绘制也不占用编号，适合封面不编号；
    ///   - 指定页码范围时，编号从范围内第一页开始按参与页顺序递增；
    ///   - 用页面自身尺寸计算位置，横版竖版都能贴边；
    ///   - 只叠加文字，不改动原有内容与文字层。
    /// </summary>
    public class PdfPageNumberConverter : IDisposable
    {
        /// <summary>为单个 PDF 添加页码，返回实际编号的页数。</summary>
        public int Apply(string sourcePath, string outputDir, PdfPageNumberOptions options)
        {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("PDF 文件不存在", sourcePath);
            options = options ?? new PdfPageNumberOptions();

            if (PdfEncryptConverter.IsEncrypted(sourcePath))
                throw new InvalidOperationException("文件已加密，请先用「PDF加密 → 移除密码」解除后再加页码");

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            var ranges = ParsePageRanges(options.PageRanges);
            var outputFile = UniquePath(outputDir,
                Path.GetFileNameWithoutExtension(sourcePath) + "_已加页码");

            int applied = 0;

            using (var doc = new PdfDocument())
            {
                doc.LoadFromFile(sourcePath);
                int pageCount = doc.Pages.Count;

                // 先选出参与编号的页（原页序），再按顺序分配页码
                var targets = new List<int>();
                for (int i = 0; i < pageCount; i++)
                {
                    int pageNumber = i + 1;
                    if (pageNumber <= options.SkipFirstPages) continue;
                    if (!InRanges(pageNumber, ranges)) continue;
                    targets.Add(i);
                }

                if (targets.Count == 0)
                    throw new InvalidOperationException("没有符合条件的页面可编号，请检查页码范围或跳过页数");

                int total = targets.Count;
                var font = new PdfTrueTypeFont(
                    new Font("Microsoft YaHei", Math.Max(5f, options.FontSize)), true);
                var brush = new PdfSolidBrush(new PdfRGBColor(options.Color));

                for (int k = 0; k < targets.Count; k++)
                {
                    int pageIndex = targets[k];
                    int displayNumber = options.StartNumber + k;
                    string text = FormatText(displayNumber, total, options.Format);

                    var page = doc.Pages[pageIndex];
                    var size = page.Size;

                    float margin = Math.Max(0f, options.MarginPt);
                    float lineHeight = Math.Max(options.FontSize * 1.8f, 12f);

                    bool isTop = options.Position == PageNumberPosition.TopCenter
                              || options.Position == PageNumberPosition.TopLeft
                              || options.Position == PageNumberPosition.TopRight;

                    float y = isTop ? margin : size.Height - margin - lineHeight;

                    var format = new PdfStringFormat
                    {
                        LineAlignment = PdfVerticalAlignment.Middle
                    };

                    switch (options.Position)
                    {
                        case PageNumberPosition.BottomLeft:
                        case PageNumberPosition.TopLeft:
                            format.Alignment = PdfTextAlignment.Left;
                            break;
                        case PageNumberPosition.BottomRight:
                        case PageNumberPosition.TopRight:
                            format.Alignment = PdfTextAlignment.Right;
                            break;
                        default:
                            format.Alignment = PdfTextAlignment.Center;
                            break;
                    }

                    var rect = new RectangleF(
                        margin,
                        y,
                        Math.Max(10f, size.Width - margin * 2f),
                        lineHeight);

                    page.Canvas.DrawString(text, font, brush, rect, format);
                    applied++;
                }

                doc.SaveToFile(outputFile);
            }

            return applied;
        }

        private static string FormatText(int number, int total, PageNumberFormat format)
        {
            switch (format)
            {
                case PageNumberFormat.WithPrefix:
                    return $"第 {number} 页";
                case PageNumberFormat.WithTotal:
                    return $"{number} / {total}";
                case PageNumberFormat.Dash:
                    return $"- {number} -";
                default:
                    return number.ToString();
            }
        }

        private static List<int[]> ParsePageRanges(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return null;
            var text = input.Trim();
            if (text == "全部" || text.Equals("all", StringComparison.OrdinalIgnoreCase)) return null;

            try
            {
                return PdfSplitter.ParseRanges(text);
            }
            catch (Exception ex)
            {
                throw new FormatException("页码范围格式错误：" + ex.Message);
            }
        }

        private static bool InRanges(int page, List<int[]> ranges)
        {
            if (ranges == null || ranges.Count == 0) return true;
            foreach (var r in ranges)
            {
                if (page >= r[0] && page <= r[1]) return true;
            }
            return false;
        }

        private static string UniquePath(string dir, string baseName)
        {
            var path = Path.Combine(dir, baseName + ".pdf");
            int seq = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(dir, $"{baseName}_{seq++}.pdf");
            }
            return path;
        }

        public void Dispose() { }
    }
}
