using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Spire.Xls;

namespace DocToPdfTool.Utils
{
    /// <summary>Excel 拆分 / 合并模式。</summary>
    public enum ExcelSplitMergeMode
    {
        /// <summary>按工作表拆分：一个工作簿 → 每张工作表一个文件。</summary>
        SplitBySheet = 0,

        /// <summary>合并为一个工作簿：多个工作簿 → 一个文件，每个源工作簿的表都保留。</summary>
        MergeToOne = 1
    }

    /// <summary>Excel 拆分 / 合并参数。</summary>
    public class ExcelSplitMergeOptions
    {
        public ExcelSplitMergeMode Mode { get; set; } = ExcelSplitMergeMode.SplitBySheet;

        /// <summary>拆分时跳过隐藏工作表。</summary>
        public bool SkipHiddenSheets { get; set; } = true;

        /// <summary>跳过没有任何内容的空工作表。</summary>
        public bool SkipEmptySheets { get; set; } = true;

        /// <summary>合并时保留原工作表名（重名自动加序号）；关闭则用「源文件名_表名」。</summary>
        public bool KeepSheetNames { get; set; } = true;
    }

    /// <summary>合并结果。</summary>
    public class ExcelMergeResult
    {
        public string OutputFile { get; set; }
        public List<string> SheetNames { get; set; } = new List<string>();
    }

    /// <summary>单个工作簿的拆分结果。</summary>
    public class ExcelSplitResult
    {
        public int SheetCount { get; set; }
        public List<string> OutputFiles { get; set; } = new List<string>();
        public int SkippedHidden { get; set; }

        public int SkippedEmpty { get; set; }
    }

    /// <summary>
    /// Excel 拆分 / 合并器（基于本地 Excel 引擎的工作表复制能力）。
    ///
    /// 实现要点：
    ///   - 拆分：逐个工作表复制到新工作簿后另存，保留公式、格式、行列宽；
    ///   - 合并：按文件顺序逐表复制到同一工作簿，可保留原表名并自动处理重名；
    ///   - 新建工作簿默认自带一张空表，复制完成后移除，避免留下空白页。
    /// </summary>
    public class ExcelSplitMergeConverter : IDisposable
    {
        /// <summary>按工作表拆分单个 Excel。</summary>
        public ExcelSplitResult Split(string excelPath, string outputDir, ExcelSplitMergeOptions options)
        {
            if (string.IsNullOrEmpty(excelPath)) throw new ArgumentNullException(nameof(excelPath));
            if (!File.Exists(excelPath)) throw new FileNotFoundException("Excel 文件不存在", excelPath);
            options = options ?? new ExcelSplitMergeOptions();

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            var baseName = Path.GetFileNameWithoutExtension(excelPath);
            var result = new ExcelSplitResult();

            using (var source = new Workbook())
            {
                ExcelToPdfConverterSpire.ApplyEngineLicense(source);
                source.RecalculateOnOpen = false;
                source.LoadFromFile(excelPath);
                ExcelToPdfConverterSpire.RemoveNoticeSheets(source);

                result.SheetCount = source.Worksheets.Count;

                foreach (Worksheet sheet in source.Worksheets)
                {
                    if (options.SkipHiddenSheets && sheet.Visibility != WorksheetVisibility.Visible)
                    {
                        result.SkippedHidden++;
                        continue;
                    }

                    if (options.SkipEmptySheets && IsEmptySheet(sheet))
                    {
                        result.SkippedEmpty++;
                        continue;
                    }

                    string sheetName = SafeName(string.IsNullOrWhiteSpace(sheet.Name) ? "Sheet" : sheet.Name);
                    var outputPath = UniquePath(outputDir, $"{baseName}_{sheetName}", ".xlsx");

                    using (var target = new Workbook())
                    {
                        ExcelToPdfConverterSpire.ApplyEngineLicense(target);
                        // 新工作簿自带若干默认空表：先记录，复制完成后再统一移除
                        var initial = target.Worksheets.Cast<Worksheet>().ToList();
                        var copied = target.Worksheets.AddCopy(sheet);
                        RemoveInitialSheets(target, initial);
                        if (copied != null)
                        {
                            copied.Name = sheetName;
                            // 单独拆出的隐藏表设为可见，否则打开文件看不到内容
                            if (sheet.Visibility != WorksheetVisibility.Visible)
                                copied.Visibility = WorksheetVisibility.Visible;
                        }
                        target.SaveToFile(outputPath, FileFormat.Version2016);
                    }

                    result.OutputFiles.Add(outputPath);
                }
            }

            if (result.OutputFiles.Count == 0)
            {
                throw new InvalidOperationException(
                    result.SkippedHidden > 0
                        ? "所有工作表都是隐藏表，已按设置跳过；可关闭「跳过隐藏工作表」后重试"
                        : "该文件没有可拆分的工作表");
            }

            return result;
        }

        /// <summary>把多个 Excel 合并为一个工作簿，返回输出文件路径。</summary>
        public ExcelMergeResult Merge(IList<string> excelPaths, string outputDir, ExcelSplitMergeOptions options)
        {
            if (excelPaths == null || excelPaths.Count == 0)
                throw new ArgumentException("请至少添加一个 Excel 文件");
            options = options ?? new ExcelSplitMergeOptions();

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            var outputPath = UniquePath(outputDir,
                "合并工作簿_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"), ".xlsx");

            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using (var target = new Workbook())
            {
                ExcelToPdfConverterSpire.ApplyEngineLicense(target);
                // 新建工作簿自带若干默认空表，合并完成后统一移除
                var initialSheets = target.Worksheets.Cast<Worksheet>().ToList();

                foreach (var file in excelPaths)
                {
                    if (!File.Exists(file)) throw new FileNotFoundException("Excel 文件不存在", file);

                    using (var source = new Workbook())
                    {
                        ExcelToPdfConverterSpire.ApplyEngineLicense(source);
                        source.RecalculateOnOpen = false;
                        source.LoadFromFile(file);
                        ExcelToPdfConverterSpire.RemoveNoticeSheets(source);

                        var fileBaseName = Path.GetFileNameWithoutExtension(file);

                        foreach (Worksheet sheet in source.Worksheets)
                        {
                            bool isHidden = sheet.Visibility != WorksheetVisibility.Visible;
                            if (options.SkipHiddenSheets && isHidden) continue;
                            if (options.SkipEmptySheets && IsEmptySheet(sheet)) continue;

                            var copied = target.Worksheets.AddCopy(sheet);
                            if (copied == null) continue;

                            string desired = options.KeepSheetNames
                                ? (string.IsNullOrWhiteSpace(sheet.Name) ? "Sheet" : sheet.Name)
                                : $"{fileBaseName}_{sheet.Name}";

                            // AddCopy 可能已自动改名，这里按目标工作簿的真实表名重新去重后再命名
                            ApplySheetName(target, copied, desired);

                            // 隐藏表被合入后设为可见，避免用户找不到内容
                            if (isHidden)
                            {
                                try { copied.Visibility = WorksheetVisibility.Visible; } catch { }
                            }
                        }
                    }
                }

                RemoveInitialSheets(target, initialSheets);

                target.SaveToFile(outputPath, FileFormat.Version2016);

                var result = new ExcelMergeResult { OutputFile = outputPath };
                result.SheetNames.AddRange(target.Worksheets.Cast<Worksheet>().Select(w => w.Name));
                return result;
            }
        }

        /// <summary>给复制后的工作表命名：去重 → 直接改名 → 失败则先改临时名再重试。</summary>
        private static void ApplySheetName(Workbook target, Worksheet copied, string desired)
        {
            var existing = new HashSet<string>(
                target.Worksheets.Cast<Worksheet>()
                    .Where(w => w.Index != copied.Index)
                    .Select(w => w.Name),
                StringComparer.OrdinalIgnoreCase);

            string finalName = MakeUniqueSheetName(desired, existing);

            if (TryRename(copied, finalName)) return;

            // 直接改名被拒绝时，先改成临时名再重试一次
            string tempName = "_tmp" + Guid.NewGuid().ToString("N").Substring(0, 6);
            if (TryRename(copied, tempName)) TryRename(copied, finalName);
        }

        private static bool TryRename(Worksheet sheet, string name)
        {
            try
            {
                sheet.Name = name;
                return string.Equals(sheet.Name, name, StringComparison.Ordinal);
            }
            catch { return false; }
        }

        /// <summary>判断工作表是否没有任何内容。</summary>
        private static bool IsEmptySheet(Worksheet sheet)
        {
            try
            {
                if (sheet.LastRow < 1 || sheet.LastColumn < 1) return true;
                var cell = sheet.Range[1, 1];
                var last = sheet.Range[sheet.LastRow, sheet.LastColumn];
                // 单格空白且行列范围为 1 时视为空表
                if (sheet.LastRow == 1 && sheet.LastColumn == 1 &&
                    string.IsNullOrWhiteSpace(cell.Text) && cell.NumberValue == 0)
                    return true;
                return false;
            }
            catch { return false; }
        }

        /// <summary>移除新建工作簿自带的默认空表（至少保留一张）。</summary>
        private static void RemoveInitialSheets(Workbook wb, IList<Worksheet> initial)
        {
            if (initial == null) return;
            foreach (var sheet in initial)
            {
                if (wb.Worksheets.Count <= 1) break;
                try { wb.Worksheets.Remove(sheet); } catch { }
            }
        }

        /// <summary>工作表名最长 31 字符，且不允许 : \ / ? * [ ]。</summary>
        private static string MakeUniqueSheetName(string desired, HashSet<string> used)
        {
            var name = desired ?? "Sheet";
            foreach (var ch in new[] { ':', '\\', '/', '?', '*', '[', ']' })
                name = name.Replace(ch, '_');
            if (name.Length > 31) name = name.Substring(0, 31);

            var candidate = name;
            int seq = 1;
            while (used.Contains(candidate))
            {
                var suffix = "_" + seq++;
                int keep = Math.Min(name.Length, 31 - suffix.Length);
                candidate = name.Substring(0, keep) + suffix;
            }
            used.Add(candidate);
            return candidate;
        }

        /// <summary>文件名非法字符替换。</summary>
        internal static string SafeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Sheet";
            foreach (var ch in Path.GetInvalidFileNameChars())
                name = name.Replace(ch, '_');
            name = name.Trim().TrimEnd('.');
            return string.IsNullOrEmpty(name) ? "Sheet" : name;
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
