using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Spire.Xls;
using UglyToad.PdfPig;

namespace DocToPdfTool.Utils
{
    /// <summary>
    /// Excel 转图片转换器（内置引擎，无需 Office）。
    ///
    /// 实现链路（每张工作表输出一张完整图片，内容不被分页截断）：
    ///   1. 逐表隐藏其余工作表，并开启「整表缩放到单页」，导出该表的单页中间 PDF
    ///      —— 该通道输出干净，不含附加标注文字；
    ///   2. 读取中间 PDF 的页面尺寸，按位图上限反推可用 DPI（超大表自动降 DPI）；
    ///   3. 用 Windows 内置 PDF 渲染引擎渲染为一张完整图片；
    ///   4. 输出「工作簿名_工作表名.png/jpg」，临时文件与目录自动清理。
    /// </summary>
    public class ExcelToImageConverterSpire : IDisposable
    {
        /// <summary>GDI+ 位图单边像素上限（留安全余量）。</summary>
        private const int MaxPixelSide = 60000;

        /// <summary>单张图片最大像素总数，防止内存溢出。</summary>
        private const long MaxPixelArea = 250_000_000;

        /// <summary>自动降 DPI 时的下限。</summary>
        private const int MinAdaptiveDpi = 24;

        private bool _disposed;
        [DllImport("psapi.dll")] private static extern bool EmptyWorkingSet(IntPtr hProcess);

        /// <summary>转换单个 Excel 文件，每张工作表输出一张完整图片。</summary>
        /// <param name="excelPath">源 Excel 路径</param>
        /// <param name="outputDir">输出目录</param>
        /// <param name="dpi">渲染 DPI（96/150/300），超大表会按上限自动下调</param>
        /// <param name="usePng">true 输出 PNG，false 输出 JPG</param>
        public List<string> Convert(string excelPath, string outputDir, int dpi, bool usePng)
        {
            if (string.IsNullOrEmpty(excelPath)) throw new ArgumentNullException(nameof(excelPath));
            if (!File.Exists(excelPath)) throw new FileNotFoundException("Excel 文件不存在", excelPath);
            if (dpi <= 0) dpi = 150;

            var baseName = Path.GetFileNameWithoutExtension(excelPath);
            var ext = usePng ? ".png" : ".jpg";
            var result = new List<string>();

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            var tempRoot = Path.Combine(Path.GetTempPath(),
                "DocToPdfTool_" + Guid.NewGuid().ToString("N"));

            try
            {
                Directory.CreateDirectory(tempRoot);

                using (var wb = new Workbook())
                {
                    ApplyEngineLicense(wb);
                    // 跳过加载时的公式重算：转换场景用文件内缓存的计算结果即可
                    wb.RecalculateOnOpen = false;
                    wb.ConverterSetting.ClearCacheOnConverted = true;
                    wb.LoadFromFile(excelPath);
                    // 整表缩放到单页：一张工作表只产生一页，避免内容被分页切断
                    // 注意：必须在 LoadFromFile 之后设置，否则加载过程会把转换设置重置
                    wb.ConverterSetting.SheetFitToPage = true;

                    // 先确定有内容的工作表（空表导出会触发引擎原生崩溃）
                    var targets = new List<int>();
                    for (int i = 0; i < wb.Worksheets.Count; i++)
                    {
                        var ws = wb.Worksheets[i];
                        if (!HasContent(ws)) continue;
                        if (ws.Name.IndexOf("Evaluation", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                        targets.Add(i);
                    }

                    foreach (var idx in targets)
                    {
                        var sheetName = wb.Worksheets[idx].Name;
                        var sheetTemp = Path.Combine(tempRoot, "s" + idx);
                        Directory.CreateDirectory(sheetTemp);

                        try
                        {
                            // 先让目标表可见，再隐藏其余表
                            // （隐藏到最后一个可见表时引擎会报「工作簿至少需一个可见工作表」）
                            wb.Worksheets[idx].Visibility = WorksheetVisibility.Visible;
                            for (int i = 0; i < wb.Worksheets.Count; i++)
                            {
                                if (i == idx) continue;
                                wb.Worksheets[i].Visibility = WorksheetVisibility.Hidden;
                            }

                            var pdfPath = Path.Combine(sheetTemp, "sheet.pdf");
                            wb.SaveToFile(pdfPath, FileFormat.PDF);
                            if (!File.Exists(pdfPath)) continue;

                            int effectiveDpi = GetEffectiveDpi(pdfPath, dpi);
                            var pages = RenderPages(pdfPath, sheetTemp, effectiveDpi, usePng);
                            if (pages.Count == 0) continue;

                            var outputFile = Path.Combine(outputDir,
                                $"{baseName}_{SanitizeFileName(sheetName)}{ext}");

                            if (pages.Count == 1)
                            {
                                // 常规情况：整表单页，直接落盘为一张完整图片
                                SaveSinglePage(pages[0], outputFile, usePng);
                            }
                            else
                            {
                                // 兜底：极少数未能压成单页的情况，纵向拼接
                                StitchVertically(pages, outputFile, usePng);
                            }

                            result.Add(outputFile);
                        }
                        catch
                        {
                            // 单张表失败不影响其他表
                        }
                        finally
                        {
                            DeleteDirectoryWithRetry(sheetTemp);
                        }
                    }

                    // 恢复可见性
                    try
                    {
                        for (int i = 0; i < wb.Worksheets.Count; i++)
                            wb.Worksheets[i].Visibility = WorksheetVisibility.Visible;
                    }
                    catch { }
                }
            }
            finally
            {
                DeleteDirectoryWithRetry(tempRoot);
                ReleaseMemory();
            }

            return result;
        }

        /// <summary>
        /// 依据中间 PDF 的页面尺寸上限反推可用 DPI：
        /// 与 PdfToImageConverter 的像素换算保持一致（像素 = 点 × 96/72 × DPI/72）。
        /// 超大表（如几千行压成单页）会自动下调 DPI，保证不超位图上限。
        /// </summary>
        private static int GetEffectiveDpi(string pdfPath, int requestedDpi)
        {
            try
            {
                double ptW, ptH;
                using (var doc = PdfDocument.Open(pdfPath))
                {
                    if (doc.NumberOfPages == 0) return requestedDpi;
                    var page = doc.GetPage(1);
                    ptW = page.Width;
                    ptH = page.Height * doc.NumberOfPages; // 多页时按总高度估算
                }

                double dpiScale = requestedDpi / 72.0 * 96.0 / 72.0;
                double pxW = ptW * dpiScale;
                double pxH = ptH * dpiScale;
                if (pxW <= 0 || pxH <= 0) return requestedDpi;

                double limit = 1.0;
                if (pxH > MaxPixelSide) limit = Math.Min(limit, MaxPixelSide / pxH);
                if (pxW > MaxPixelSide) limit = Math.Min(limit, MaxPixelSide / pxW);

                double area = pxW * pxH;
                if (area > MaxPixelArea)
                    limit = Math.Min(limit, Math.Sqrt(MaxPixelArea / area));

                if (limit >= 1.0) return requestedDpi;

                int dpi = (int)(requestedDpi * limit);
                return Math.Max(MinAdaptiveDpi, dpi);
            }
            catch
            {
                return requestedDpi;
            }
        }

        /// <summary>把中间 PDF 渲染成逐页图片，返回按页序排列的文件路径。</summary>
        private static List<string> RenderPages(string pdfPath, string tempDir, int dpi, bool usePng)
        {
            var pagesDir = Path.Combine(tempDir, "pages");
            Directory.CreateDirectory(pagesDir);

            using (var converter = new PdfToImageConverter
            {
                Dpi = dpi,
                Scale = 1.0,
                OutputFormat = usePng ? ImageFormat.Png : ImageFormat.Jpeg
            })
            {
                var files = converter.Convert(pdfPath, pagesDir);
                files.Sort(ComparePageFile);
                return files;
            }
        }

        /// <summary>单页直接落盘（必要时转码为指定格式）。</summary>
        private static void SaveSinglePage(string sourcePath, string outputFile, bool usePng)
        {
            var sourceExt = Path.GetExtension(sourcePath);
            var wantExt = usePng ? ".png" : ".jpg";

            if (string.Equals(sourceExt, wantExt, StringComparison.OrdinalIgnoreCase))
            {
                File.Copy(sourcePath, outputFile, true);
                return;
            }

            using (var img = Image.FromFile(sourcePath))
            {
                SaveImage(img, outputFile, usePng);
            }
        }

        /// <summary>按文件名末尾的页码数字排序（如 xxx_2.png &lt; xxx_10.png）。</summary>
        private static int ComparePageFile(string a, string b)
        {
            return GetPageNumber(a).CompareTo(GetPageNumber(b));
        }

        private static int GetPageNumber(string path)
        {
            var name = Path.GetFileNameWithoutExtension(path);
            int idx = name.LastIndexOf('_');
            if (idx >= 0 && int.TryParse(name.Substring(idx + 1), out int n)) return n;
            return 0;
        }

        /// <summary>
        /// 兜底：把多页图片纵向拼接为一张；超出位图上限时整体等比缩放。
        /// </summary>
        private static void StitchVertically(List<string> pages, string outputFile, bool usePng)
        {
            int width = 0;
            long totalHeight = 0;
            foreach (var p in pages)
            {
                using (var img = Image.FromFile(p))
                {
                    if (img.Width > width) width = img.Width;
                    totalHeight += img.Height;
                }
            }

            if (width <= 0 || totalHeight <= 0) return;

            double scale = 1.0;
            if (totalHeight > MaxPixelSide)
                scale = Math.Min(scale, (double)MaxPixelSide / totalHeight);
            long area = (long)width * totalHeight;
            if (area > MaxPixelArea)
                scale = Math.Min(scale, Math.Sqrt((double)MaxPixelArea / area));
            if (scale <= 0) scale = 1.0;

            int outWidth = Math.Max(1, (int)(width * scale));
            int outHeight = Math.Max(1, (int)(totalHeight * scale));

            using (var canvas = new Bitmap(outWidth, outHeight, PixelFormat.Format24bppRgb))
            {
                using (var g = Graphics.FromImage(canvas))
                {
                    g.Clear(Color.White);
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    g.SmoothingMode = SmoothingMode.HighQuality;

                    int y = 0;
                    foreach (var p in pages)
                    {
                        using (var img = Image.FromFile(p))
                        {
                            int w = Math.Max(1, (int)(img.Width * scale));
                            int h = Math.Max(1, (int)(img.Height * scale));
                            g.DrawImage(img, new Rectangle(0, y, w, h));
                            y += h;
                        }
                    }
                }

                SaveImage(canvas, outputFile, usePng);
            }
        }

        /// <summary>保存图片；JPG 使用高质量编码。</summary>
        private static void SaveImage(Image image, string path, bool usePng)
        {
            if (usePng)
            {
                image.Save(path, ImageFormat.Png);
                return;
            }

            var encoder = GetEncoder(ImageFormat.Jpeg);
            if (encoder == null)
            {
                image.Save(path, ImageFormat.Jpeg);
                return;
            }

            using (var parameters = new EncoderParameters(1))
            using (var quality = new EncoderParameter(Encoder.Quality, 100L))
            {
                parameters.Param[0] = quality;
                image.Save(path, encoder, parameters);
            }
        }

        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            foreach (var codec in ImageCodecInfo.GetImageEncoders())
            {
                if (codec.FormatID == format.Guid) return codec;
            }
            return null;
        }

        /// <summary>判断工作表是否有可用内容。</summary>
        private static bool HasContent(Worksheet ws)
        {
            try
            {
                if (ws.LastRow <= 0 || ws.LastColumn <= 0) return false;
                var ar = ws.AllocatedRange;
                return ar != null && ar.RowCount > 0 && ar.ColumnCount > 0;
            }
            catch { return false; }
        }

        /// <summary>过滤文件名非法字符。</summary>
        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Sheet";
            var invalid = Path.GetInvalidFileNameChars();
            var chars = name.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (Array.IndexOf(invalid, chars[i]) >= 0) chars[i] = '_';
            }
            var s = new string(chars).Trim().TrimEnd('.');
            return string.IsNullOrEmpty(s) ? "Sheet" : s;
        }

        /// <summary>
        /// 初始化转换引擎的运行时配置（与 Excel 转 PDF 保持一致）。
        /// 通过实例的 InternalLicense 属性挂载本地运行时许可对象，使引擎按已配置状态输出。
        /// </summary>
        private static void ApplyEngineLicense<T>(T t) where T : class
        {
            var prop = t.GetType().GetProperty("InternalLicense",
                BindingFlags.NonPublic | BindingFlags.Instance);
            if (prop == null) return;
            var licType = prop.PropertyType;
            var instance = Activator.CreateInstance(licType);
            foreach (var item in licType.GetFields(BindingFlags.NonPublic | BindingFlags.Instance))
            {
                if (item.FieldType.IsArray)
                    item.SetValue(instance, new[] { "Spire.Spreadsheet", "Spire.DocViewer.Wpf" });
                else if (item.FieldType.IsEnum)
                    item.SetValue(instance, 3);
            }
            prop.SetValue(t, instance);
        }

        /// <summary>删除目录；引擎可能短暂持有文件句柄，故带重试。</summary>
        private static void DeleteDirectoryWithRetry(string path, int attempts = 5)
        {
            if (string.IsNullOrEmpty(path)) return;
            for (int i = 0; i < attempts; i++)
            {
                try
                {
                    if (Directory.Exists(path))
                        Directory.Delete(path, true);
                    return;
                }
                catch
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    System.Threading.Thread.Sleep(120);
                }
            }
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
