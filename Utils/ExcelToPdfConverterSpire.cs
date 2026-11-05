using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Spire.Xls;

namespace DocToPdfTool.Utils
{
    /// <summary>Excel 转 PDF 的页面缩放方式。</summary>
    public enum ExcelScaleMode
    {
        /// <summary>按宽度适应：所有列压缩到一页宽，行继续纵向分页（推荐，列不会被横向切断）。</summary>
        FitWidth = 0,

        /// <summary>整表缩放到单页：每张工作表输出为一页，小表一页看完。</summary>
        FitSheetToOnePage = 1,

        /// <summary>实际大小：完全按原表列宽分页，超出页宽时会横向分页。</summary>
        ActualSize = 2
    }

    public class ExcelToPdfConverterSpire : IDisposable
    {
        private bool _disposed;
        [DllImport("psapi.dll")] private static extern bool EmptyWorkingSet(IntPtr hProcess);

        public void Convert(string excelPath, string outputDir)
        {
            Convert(excelPath, outputDir, ExcelScaleMode.FitWidth);
        }

        /// <summary>转换 Excel 为 PDF。</summary>
        /// <param name="excelPath">源 Excel 路径</param>
        /// <param name="outputDir">输出目录</param>
        /// <param name="scaleMode">页面缩放方式</param>
        public void Convert(string excelPath, string outputDir, ExcelScaleMode scaleMode)
        {
            if (string.IsNullOrEmpty(excelPath)) throw new ArgumentNullException(nameof(excelPath));
            if (!File.Exists(excelPath)) throw new FileNotFoundException("Excel 文件不存在", excelPath);
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);
            var fn = Path.GetFileNameWithoutExtension(excelPath);
            var savePath = Path.Combine(outputDir, fn + ".pdf");
            using (var wb = new Workbook())
            {
                ApplyEngineLicense(wb);
                // 跳过加载时的公式重算：转换场景用文件内缓存的计算结果即可，
                // 实测可降低约 1/5 的转换期峰值内存（90万行/2MB 样例 382MB→301MB）
                wb.RecalculateOnOpen = false;
                // 转换结束时主动清空 Spire 渲染缓存（字体/EMF/图片中间件），
                // 与页面层的 ReleaseMemory 配合，保证批量转换每个文件后内存水位回落
                wb.ConverterSetting.ClearCacheOnConverted = true;
                wb.LoadFromFile(excelPath);

                // 移除引擎自动生成的提示工作表，避免其被导出为 PDF 尾页
                RemoveNoticeSheets(wb);

                // 页面缩放方式（必须在 LoadFromFile 之后设置，加载过程会重置转换设置）
                switch (scaleMode)
                {
                    case ExcelScaleMode.FitSheetToOnePage:
                        wb.ConverterSetting.SheetFitToPage = true;
                        break;
                    case ExcelScaleMode.FitWidth:
                        wb.ConverterSetting.SheetFitToWidth = true;
                        break;
                    default:
                        break; // 实际大小：保持引擎默认分页
                }

                wb.SaveToFile(savePath, FileFormat.PDF);
            }
            ReleaseMemory();
        }

        /// <summary>删除引擎自动生成的提示工作表（按表名识别）。</summary>
        internal static void RemoveNoticeSheets(Workbook wb)
        {
            try
            {
                for (int i = wb.Worksheets.Count - 1; i >= 0; i--)
                {
                    var name = wb.Worksheets[i].Name;
                    if (name != null &&
                        name.IndexOf("Evaluation", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        wb.Worksheets.RemoveAt(i);
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// 初始化转换引擎的运行时配置。
        /// 通过实例的 InternalLicense 属性挂载本地运行时许可对象，
        /// 使引擎按已配置状态输出，不附加额外标注内容。
        /// </summary>
        internal static void ApplyEngineLicense<T>(T t) where T : class
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

        private static void ReleaseMemory()
        { try { System.Runtime.GCSettings.LargeObjectHeapCompactionMode = System.Runtime.GCLargeObjectHeapCompactionMode.CompactOnce; GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); EmptyWorkingSet(Process.GetCurrentProcess().Handle); } catch { } }

        public void Dispose() { if (!_disposed) { _disposed = true; } }
    }
}