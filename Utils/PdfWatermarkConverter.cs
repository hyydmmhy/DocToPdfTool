using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Spire.Pdf;
using Spire.Pdf.Graphics;

namespace DocToPdfTool.Utils
{
    /// <summary>水印类型。</summary>
    public enum PdfWatermarkType
    {
        /// <summary>文字水印。</summary>
        Text = 0,

        /// <summary>图片水印。</summary>
        Image = 1
    }

    /// <summary>PDF 水印参数。</summary>
    public class PdfWatermarkOptions
    {
        public PdfWatermarkType Type { get; set; } = PdfWatermarkType.Text;

        /// <summary>文字水印内容。</summary>
        public string Text { get; set; } = "机密";

        /// <summary>图片水印路径（Type = Image 时使用）。</summary>
        public string ImagePath { get; set; }

        /// <summary>图片水印宽度占页宽比例（0.05~0.9）。</summary>
        public float ImageScale { get; set; } = 0.3f;

        /// <summary>字号。</summary>
        public float FontSize { get; set; } = 48f;

        /// <summary>颜色。</summary>
        public Color Color { get; set; } = Color.Gray;

        /// <summary>透明度（0.05~1）。</summary>
        public float Opacity { get; set; } = 0.3f;

        /// <summary>旋转角度（度）。</summary>
        public float Rotation { get; set; } = 45f;

        /// <summary>是否平铺铺满整页。</summary>
        public bool Tile { get; set; }

        /// <summary>平铺时的水平/垂直步长（磅）。</summary>
        public float TileStepX { get; set; } = 220f;
        public float TileStepY { get; set; } = 160f;

        /// <summary>页码范围；为 null 或空表示全部页。格式同「1-3,5」。</summary>
        public string PageRanges { get; set; }
    }

    /// <summary>
    /// PDF 水印转换器（文字 / 图片，支持透明度、旋转、平铺、页码范围）。
    ///
    /// 实现要点：
    ///   - 逐页在原有内容之上绘制，不改变原页面内容与文字层；
    ///   - 中文使用系统 TrueType 字体（微软雅黑），避免缺字；
    ///   - 通过 Canvas 的 Save/Translate/Rotate/SetTransparency 实现倾斜与半透明。
    /// </summary>
    public class PdfWatermarkConverter : IDisposable
    {
        /// <summary>为单个 PDF 添加水印，返回处理的页数。</summary>
        public int Apply(string sourcePath, string outputDir, PdfWatermarkOptions options)
        {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("PDF 文件不存在", sourcePath);
            options = options ?? new PdfWatermarkOptions();

            if (PdfEncryptConverter.IsEncrypted(sourcePath))
                throw new InvalidOperationException("文件已加密，请先用「PDF加密 → 移除密码」解除后再加水印");

            if (options.Type == PdfWatermarkType.Text && string.IsNullOrWhiteSpace(options.Text))
                throw new ArgumentException("水印文字不能为空");
            if (options.Type == PdfWatermarkType.Image)
            {
                if (string.IsNullOrWhiteSpace(options.ImagePath) || !File.Exists(options.ImagePath))
                    throw new ArgumentException("请选择有效的水印图片");
            }

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            var ranges = ParsePageRanges(options.PageRanges);
            var outputFile = UniquePath(outputDir, Path.GetFileNameWithoutExtension(sourcePath) + "_已加水印");

            int affected = 0;
            using (var doc = new PdfDocument())
            {
                doc.LoadFromFile(sourcePath);

                var font = options.Type == PdfWatermarkType.Text
                    ? new PdfTrueTypeFont(new Font("Microsoft YaHei", options.FontSize), true)
                    : null;
                var brush = options.Type == PdfWatermarkType.Text
                    ? new PdfSolidBrush(new PdfRGBColor(options.Color))
                    : null;
                var image = options.Type == PdfWatermarkType.Image
                    ? PdfImage.FromFile(options.ImagePath)
                    : null;

                float alpha = Math.Min(1f, Math.Max(0.03f, options.Opacity));

                for (int i = 0; i < doc.Pages.Count; i++)
                {
                    int pageNumber = i + 1;
                    if (!InRanges(pageNumber, ranges)) continue;

                    var page = doc.Pages[i];
                    var canvas = page.Canvas;
                    var size = page.Size;

                    canvas.Save();
                    canvas.SetTransparency(alpha);

                    if (options.Tile)
                        DrawTiled(canvas, size, options, font, brush, image);
                    else
                        DrawSingle(canvas, size, options, font, brush, image);

                    canvas.Restore();
                    affected++;
                }

                if (affected == 0)
                    throw new InvalidOperationException("指定页码超出文件页数范围，未添加任何水印");

                doc.SaveToFile(outputFile);
            }

            return affected;
        }

        /// <summary>居中绘制单个水印。</summary>
        private static void DrawSingle(PdfCanvas canvas, SizeF size, PdfWatermarkOptions options,
            PdfTrueTypeFont font, PdfSolidBrush brush, PdfImage image)
        {
            float cx = size.Width / 2f;
            float cy = size.Height / 2f;

            canvas.Save();
            canvas.TranslateTransform(cx, cy);
            if (Math.Abs(options.Rotation) > 0.01f)
                canvas.RotateTransform(options.Rotation);

            if (options.Type == PdfWatermarkType.Text)
            {
                var format = new PdfStringFormat
                {
                    Alignment = PdfTextAlignment.Center,
                    LineAlignment = PdfVerticalAlignment.Middle
                };
                float boxH = Math.Max(options.FontSize * 2f, 40f);
                float boxW = Math.Max(size.Width, options.FontSize * options.Text.Length * 2f);
                canvas.DrawString(options.Text, font, brush,
                    new RectangleF(-boxW / 2f, -boxH / 2f, boxW, boxH), format);
            }
            else
            {
                float w = size.Width * Clamp(options.ImageScale, 0.05f, 0.9f);
                float h = w * image.Height / Math.Max(1, image.Width);
                canvas.DrawImage(image, new RectangleF(-w / 2f, -h / 2f, w, h));
            }

            canvas.Restore();
        }

        /// <summary>平铺绘制水印。</summary>
        private static void DrawTiled(PdfCanvas canvas, SizeF size, PdfWatermarkOptions options,
            PdfTrueTypeFont font, PdfSolidBrush brush, PdfImage image)
        {
            float stepX = Math.Max(60f, options.TileStepX);
            float stepY = Math.Max(50f, options.TileStepY);

            var format = new PdfStringFormat
            {
                Alignment = PdfTextAlignment.Center,
                LineAlignment = PdfVerticalAlignment.Middle
            };

            // 起点偏移半个步长，避免边缘出现半截水印
            for (float y = stepY / 2f; y < size.Height + stepY; y += stepY)
            {
                for (float x = stepX / 2f; x < size.Width + stepX; x += stepX)
                {
                    canvas.Save();
                    canvas.TranslateTransform(x, y);
                    if (Math.Abs(options.Rotation) > 0.01f)
                        canvas.RotateTransform(options.Rotation);

                    if (options.Type == PdfWatermarkType.Text)
                    {
                        float boxH = Math.Max(options.FontSize * 2f, 30f);
                        float boxW = Math.Max(stepX, options.FontSize * 6f);
                        canvas.DrawString(options.Text, font, brush,
                            new RectangleF(-boxW / 2f, -boxH / 2f, boxW, boxH), format);
                    }
                    else
                    {
                        float w = size.Width * Clamp(options.ImageScale, 0.05f, 0.9f) * 0.6f;
                        float h = w * image.Height / Math.Max(1, image.Width);
                        canvas.DrawImage(image, new RectangleF(-w / 2f, -h / 2f, w, h));
                    }

                    canvas.Restore();
                }
            }
        }

        /// <summary>解析页码范围，支持 "1-3,5"；为空表示全部。</summary>
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

        private static float Clamp(float v, float min, float max)
        {
            if (v < min) return min;
            if (v > max) return max;
            return v;
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
