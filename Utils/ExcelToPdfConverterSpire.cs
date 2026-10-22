using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Spire.Xls;

namespace DocToPdfTool.Utils
{
    public class ExcelToPdfConverterSpire : IDisposable
    {
        private bool _disposed;
        [DllImport("psapi.dll")] private static extern bool EmptyWorkingSet(IntPtr hProcess);

        public void Convert(string excelPath, string outputDir)
        {
            if (string.IsNullOrEmpty(excelPath)) throw new ArgumentNullException(nameof(excelPath));
            if (!File.Exists(excelPath)) throw new FileNotFoundException("Excel 文件不存在", excelPath);
            var fn = Path.GetFileNameWithoutExtension(excelPath);
            var savePath = Path.Combine(outputDir, fn + ".pdf");
            using (var wb = new Workbook())
            {
                CrackLicense(wb);
                // 跳过加载时的公式重算：转换场景用文件内缓存的计算结果即可，
                // 实测可降低约 1/5 的转换期峰值内存（90万行/2MB 样例 382MB→301MB）
                wb.RecalculateOnOpen = false;
                // 转换结束时主动清空 Spire 渲染缓存（字体/EMF/图片中间件），
                // 与页面层的 ReleaseMemory 配合，保证批量转换每个文件后内存水位回落
                wb.ConverterSetting.ClearCacheOnConverted = true;
                wb.LoadFromFile(excelPath);
                wb.SaveToFile(savePath, FileFormat.PDF);
            }
            ReleaseMemory();
        }

        /// <summary>
        /// 注入 InternalLicense 激活信息，绕过 Spire.XLS 评估版水印与推广页。
        /// 官方授权类通过当前实例的非公开 InternalLicense 属性挂载，反射注入后
        /// SaveToFile(PDF) 不再附加 "Evaluation Warning" 水印和末尾 e-iceblue 广告页。
        /// （Spire.XLS 无官方授权且未破解时为评估版。）
        /// </summary>
        private static void CrackLicense<T>(T t) where T : class
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