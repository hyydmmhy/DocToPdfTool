using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace DocToPdfTool.Utils
{
    /// <summary>输出纸张。</summary>
    public enum PdfNupPaper
    {
        A4 = 0,
        A3 = 1
    }

    /// <summary>页面方向。</summary>
    public enum PdfNupOrientation
    {
        Portrait = 0,
        Landscape = 1
    }

    /// <summary>页面顺序。</summary>
    public enum PdfNupOrder
    {
        /// <summary>从左到右，再换行。</summary>
        LeftToRight = 0,

        /// <summary>从上到下，再换列。</summary>
        TopToBottom = 1
    }

    /// <summary>拼版参数。</summary>
    public class PdfNupOptions
    {
        /// <summary>每张纸拼几个原页面（2/4/6/9/16）。</summary>
        public int N { get; set; } = 4;

        public PdfNupPaper Paper { get; set; } = PdfNupPaper.A4;

        public PdfNupOrientation Orientation { get; set; } = PdfNupOrientation.Portrait;

        public PdfNupOrder Order { get; set; } = PdfNupOrder.LeftToRight;

        /// <summary>页边距（毫米）。</summary>
        public double MarginMm { get; set; } = 10;

        /// <summary>格间隙（毫米）。</summary>
        public double GapMm { get; set; } = 4;

        /// <summary>是否绘制裁切线（虚线分隔）。</summary>
        public bool DrawCutLines { get; set; }

        /// <summary>页码范围，如 "1-3,5"；空或 "全部" 表示全部。</summary>
        public string PageRange { get; set; }

        /// <summary>渲染/拼图 DPI（清晰度）；默认 300。</summary>
        public int Dpi { get; set; } = 300;
    }

    /// <summary>拼版结果。</summary>
    public class PdfNupResult
    {
        public string OutputFile { get; set; }
        public int SourcePages { get; set; }
        public int OutputPages { get; set; }
        public long OriginalSize { get; set; }
        public long NewSize { get; set; }
    }

    /// <summary>
    /// PDF 拼版打印（N 合一）转换器（零新增依赖）。
    ///
    /// 方案：WinRT 渲染源页为高清图片（PdfToImageConverter）→ GDI+ 按宫格拼图
    /// （Graphics.DrawImage，像素级精确，不变形、不溢出）→ 图片转 PDF（ImageToPdfConverter）。
    ///
    /// 不使用 Spire 的 PdfTemplate / DrawImage 绘制，因为其变换矩阵存在系统性偏移与缩放偏差，
    /// 导致单元格内容错位、右侧被截断。
    /// </summary>
    public class PdfNupConverter : IDisposable
    {
        /// <summary>内容在单元格内留的安全边距比例（避免贴边）。</summary>
        private const double ContentInset = 0.02;

        public PdfNupResult Convert(string sourcePath, string outputDir, PdfNupOptions options)
        {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("PDF 文件不存在", sourcePath);
            options = options ?? new PdfNupOptions();
            if (options.N < 1) options.N = 1;

            if (PdfEncryptConverter.IsEncrypted(sourcePath))
                throw new InvalidOperationException("文件已加密，请先用「PDF加密 → 移除密码」解除后再拼版");

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            long originalSize = new FileInfo(sourcePath).Length;
            var result = new PdfNupResult { OriginalSize = originalSize };

            var outputFile = UniquePath(outputDir,
                Path.GetFileNameWithoutExtension(sourcePath) + "_拼版");

            // DPI（清晰度），范围限制 72-600
            int dpi = Math.Max(72, Math.Min(600, options.Dpi <= 0 ? 300 : options.Dpi));
            double pxPerMm = dpi / 25.4;

            var (paperWmm, paperHmm) = GetPaperSizeMm(options.Paper, options.Orientation);
            var (cols, rows) = GetGrid(options.N);
            var ranges = ParsePageRanges(options.PageRange);

            // ── 1. WinRT 渲染源页为图片 ──
            string tempDir = Path.Combine(Path.GetTempPath(), "DocToPdfTool_nup_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);

            List<string> allImages;
            try
            {
                using (var renderer = new PdfToImageConverter())
                {
                    renderer.Dpi = dpi;
                    renderer.Scale = 1.0;
                    renderer.OutputFormat = ImageFormat.Png;
                    // 同步渲染所有页；Convert 内部会调用 WinRT
                    allImages = renderer.Convert(sourcePath, tempDir);
                }
            }
            catch
            {
                // 某些环境 WinRT 不可用，尝试 Spire SaveAsImage 兜底
                allImages = RenderWithSpire(sourcePath, tempDir, dpi);
            }

            try
            {
                if (allImages == null || allImages.Count == 0)
                    throw new InvalidOperationException("无法渲染源 PDF 页面");

                // 按页码范围挑选页码
                var pageIndexes = new List<int>();
                for (int i = 0; i < allImages.Count; i++)
                {
                    if (InRanges(i + 1, ranges)) pageIndexes.Add(i);
                }
                if (pageIndexes.Count == 0)
                    throw new InvalidOperationException("页码范围没有匹配到任何页面");

                result.SourcePages = pageIndexes.Count;

                // ── 2. GDI+ 按宫格拼图 ──
                int sheetW = (int)Math.Round(paperWmm * pxPerMm);
                int sheetH = (int)Math.Round(paperHmm * pxPerMm);
                int marginPx = (int)Math.Round(options.MarginMm * pxPerMm);
                int gapPx = (int)Math.Round(options.GapMm * pxPerMm);

                int cellW = (sheetW - 2 * marginPx - (cols - 1) * gapPx) / cols;
                int cellH = (sheetH - 2 * marginPx - (rows - 1) * gapPx) / rows;

                int outputCount = (int)Math.Ceiling(pageIndexes.Count / (double)options.N);
                result.OutputPages = outputCount;

                var sheetFiles = new List<string>();
                for (int sheet = 0; sheet < outputCount; sheet++)
                {
                    using (var sheetBmp = new Bitmap(sheetW, sheetH, PixelFormat.Format32bppArgb))
                    {
                        sheetBmp.SetResolution(dpi, dpi);
                        using (var g = Graphics.FromImage(sheetBmp))
                        {
                            g.SmoothingMode = SmoothingMode.HighQuality;
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                            g.Clear(Color.White);

                            for (int k = 0; k < options.N; k++)
                            {
                                int idx = sheet * options.N + k;
                                if (idx >= pageIndexes.Count) break;

                                int col, row;
                                if (options.Order == PdfNupOrder.LeftToRight)
                                {
                                    col = k % cols;
                                    row = k / cols;
                                }
                                else
                                {
                                    col = k / rows;
                                    row = k % rows;
                                }

                                int x = marginPx + col * (cellW + gapPx);
                                int y = marginPx + row * (cellH + gapPx);

                                // 读取该页图片，裁剪到内容包围盒，再等比缩放进单元格（留 2% 安全边距），居中
                                var pagePath = allImages[pageIndexes[idx]];
                                if (!File.Exists(pagePath)) continue;
                                using (var pageImg = Image.FromFile(pagePath))
                                {
                                    using (var cropped = CropToContent(pageImg))
                                    {
                                        double cellW_d = cellW * (1 - 2 * ContentInset);
                                        double cellH_d = cellH * (1 - 2 * ContentInset);
                                        double scale = Math.Min(cellW_d / cropped.Width, cellH_d / cropped.Height);
                                        int w = (int)Math.Round(cropped.Width * scale);
                                        int h = (int)Math.Round(cropped.Height * scale);
                                        int dx = x + (cellW - w) / 2;
                                        int dy = y + (cellH - h) / 2;
                                        g.DrawImage(cropped, dx, dy, w, h);
                                    }
                                }
                            }

                            if (options.DrawCutLines && (cols > 1 || rows > 1))
                                DrawCutLines(g, cols, rows, marginPx, gapPx, sheetW, sheetH);
                        }

                        var sheetFile = Path.Combine(tempDir, $"sheet_{sheet + 1}.png");
                        sheetBmp.Save(sheetFile, ImageFormat.Png);
                        sheetFiles.Add(sheetFile);
                    }
                }

                // ── 3. 图片转 PDF ──
                // 用 OriginalSize + 0 边距：每张拼图纸精确按原始尺寸填满页面（无缩放/偏移）
                using (var imgConv = new ImageToPdfConverter())
                {
                    imgConv.PageMode = ImagePageMode.OriginalSize;
                    imgConv.MarginMm = 0;
                    imgConv.MaxImageSide = 0;   // 不限制图片尺寸，保证高清拼版
                    var failures = new List<string>();
                    imgConv.Convert(sheetFiles, outputFile, null, failures);
                    if (failures.Count > 0 && !File.Exists(outputFile))
                        throw new InvalidOperationException("拼版图片转 PDF 失败：" + string.Join("; ", failures));
                }

                result.OutputFile = outputFile;
                result.NewSize = new FileInfo(outputFile).Length;
                return result;
            }
            finally
            {
                try { Directory.Delete(tempDir, true); } catch { }
            }
        }

        /// <summary>Spire SaveAsImage 兜底渲染（WinRT 不可用时）。</summary>
        private static List<string> RenderWithSpire(string pdfPath, string tempDir, int dpi)
        {
            var files = new List<string>();
            using (var doc = new Spire.Pdf.PdfDocument())
            {
                doc.LoadFromFile(pdfPath);
                for (int i = 0; i < doc.Pages.Count; i++)
                {
                    using (var img = doc.SaveAsImage(i, Spire.Pdf.Graphics.PdfImageType.Bitmap, dpi, dpi))
                    {
                        if (img is Bitmap bmp)
                        {
                            var f = Path.Combine(tempDir, $"page_{i + 1}.png");
                            bmp.Save(f, ImageFormat.Png);
                            files.Add(f);
                        }
                    }
                }
            }
            return files;
        }

        /// <summary>
        /// 检测图片内容包围盒（非白/非透明像素），裁剪到内容区域。
        /// 消除源页面渲染时四周空白，使内容精确填充单元格，避免偏右/溢出。
        /// </summary>
        private static Bitmap CropToContent(Image img)
        {
            var bmp = img as Bitmap ?? new Bitmap(img);
            int w = bmp.Width, h = bmp.Height;
            int minX = w, minY = h, maxX = -1, maxY = -1;
            const int step = 2;

            // 遍历像素找非白/非透明包围盒
            for (int y = 0; y < h; y += step)
            {
                for (int x = 0; x < w; x += step)
                {
                    var c = bmp.GetPixel(x, y);
                    bool blank = (c.A < 16) || (c.R > 245 && c.G > 245 && c.B > 245);
                    if (!blank)
                    {
                        if (x < minX) minX = x;
                        if (x > maxX) maxX = x;
                        if (y < minY) minY = y;
                        if (y > maxY) maxY = y;
                    }
                }
            }

            // 完全空白 → 返回原图（无裁剪）
            if (maxX < 0) return bmp;

            // 扩大 1px 边界防切
            int x0 = Math.Max(0, minX - 1), y0 = Math.Max(0, minY - 1);
            int x1 = Math.Min(w - 1, maxX + 1), y1 = Math.Min(h - 1, maxY + 1);
            int cw = x1 - x0 + 1, ch = y1 - y0 + 1;

            if (cw >= w && ch >= h) return bmp;   // 无空白可裁

            var cropped = new Bitmap(cw, ch, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(cropped))
            {
                g.DrawImage(bmp, new Rectangle(0, 0, cw, ch), new Rectangle(x0, y0, cw, ch), GraphicsUnit.Pixel);
            }
            if (!ReferenceEquals(bmp, img)) bmp.Dispose();
            return cropped;
        }

        /// <summary>在 GDI+ 画布上绘制裁切线（虚线）。</summary>
        private static void DrawCutLines(Graphics g, int cols, int rows,
            int marginPx, int gapPx, int sheetW, int sheetH)
        {
            using (var pen = new Pen(Color.Black, 1f) { DashStyle = DashStyle.Dash })
            {
                int cellW = (sheetW - 2 * marginPx - (cols - 1) * gapPx) / cols;
                int cellH = (sheetH - 2 * marginPx - (rows - 1) * gapPx) / rows;

                for (int c = 1; c < cols; c++)
                {
                    int x = marginPx + c * (cellW + gapPx) - gapPx / 2;
                    g.DrawLine(pen, x, marginPx, x, sheetH - marginPx);
                }
                for (int r = 1; r < rows; r++)
                {
                    int y = marginPx + r * (cellH + gapPx) - gapPx / 2;
                    g.DrawLine(pen, marginPx, y, sheetW - marginPx, y);
                }
            }
        }

        private static (double w, double h) GetPaperSizeMm(PdfNupPaper paper, PdfNupOrientation orientation)
        {
            double w, h;
            switch (paper)
            {
                case PdfNupPaper.A3:
                    w = 297.0; h = 420.0;
                    break;
                default:
                    w = 210.0; h = 297.0;   // A4
                    break;
            }
            return orientation == PdfNupOrientation.Landscape
                ? (Math.Max(w, h), Math.Min(w, h))
                : (Math.Min(w, h), Math.Max(w, h));
        }

        private static (int cols, int rows) GetGrid(int n)
        {
            switch (n)
            {
                case 2: return (2, 1);
                case 6: return (3, 2);
                case 9: return (3, 3);
                case 16: return (4, 4);
                case 1: return (1, 1);
                default: return (2, 2);   // 4 合1
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
