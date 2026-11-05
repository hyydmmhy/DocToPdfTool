using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using DocToPdfTool.Utils;
using Microsoft.Win32;

namespace DocToPdfTool.Pages
{
    /// <summary>Excel 拆分合并页面的列表项。</summary>
    public class ExcelSplitMergeItem : INotifyPropertyChanged
    {
        private string _statusText = "待处理";

        public string FileName { get; set; }
        public string FilePath { get; set; }

        public string StatusText
        {
            get => _statusText;
            set { if (_statusText != value) { _statusText = value; OnPropertyChanged(); } }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public partial class ExcelSplitMergePage : UserControl
    {
        private static readonly string[] ExcelExtensions = { ".xlsx", ".xls", ".xlsm", ".xltx", ".xltm" };

        private readonly ObservableCollection<ExcelSplitMergeItem> _files =
            new ObservableCollection<ExcelSplitMergeItem>();

        public ExcelSplitMergePage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;
            UpdateEmptyHint();
            UpdateModeControls();
        }

        // ───────────────── 文件列表 ─────────────────

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog
            {
                Title = "选择 Excel 文件",
                Filter = "Excel 工作簿 (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm|所有文件 (*.*)|*.*",
                Multiselect = true
            };
            if (d.ShowDialog() == true) AddFiles(d.FileNames);
        }

        private void FileListView_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                AddFiles((string[])e.Data.GetData(DataFormats.FileDrop));
        }

        private void FileListView_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private void BtnRemoveFile_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement el && el.Tag is ExcelSplitMergeItem item)
            {
                _files.Remove(item);
                UpdateEmptyHint();
            }
        }

        private void AddFiles(IEnumerable<string> paths)
        {
            foreach (var p in paths)
            {
                var ext = Path.GetExtension(p);
                if (!ExcelExtensions.Any(e => string.Equals(e, ext, StringComparison.OrdinalIgnoreCase))) continue;
                if (_files.Any(f => f.FilePath.Equals(p, StringComparison.OrdinalIgnoreCase))) continue;

                _files.Add(new ExcelSplitMergeItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p,
                    StatusText = "待处理"
                });
            }
            UpdateEmptyHint();
        }

        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 3) return;
            double total = FileListView.ActualWidth - 14;
            double fixedWidth = 170 + 72;
            double name = Math.Max(120, total - fixedWidth);
            if (Math.Abs(gv.Columns[0].Width - name) > 1) gv.Columns[0].Width = name;
        }

        // ───────────────── 模式联动 ─────────────────

        private bool IsMergeMode => RdoMerge.IsChecked == true;

        private void Mode_Changed(object sender, RoutedEventArgs e)
        {
            UpdateModeControls();
        }

        private void UpdateModeControls()
        {
            if (ChkKeepNames == null || TxtHint == null || BtnExecute == null) return;

            bool merge = IsMergeMode;
            ChkKeepNames.IsEnabled = merge;
            ChkKeepNames.Opacity = merge ? 1.0 : 0.45;

            BtnExecute.Content = merge ? "开始合并" : "开始拆分";
            TxtHint.Text = merge
                ? "合并：按列表顺序把每个工作簿的全部工作表汇总到一个新工作簿；重名工作表自动追加序号"
                : "拆分：每个工作簿的每张工作表导出为独立文件；公式、格式、行列宽都会保留";
        }

        // ───────────────── 执行 ─────────────────

        private async void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            if (_files.Count == 0)
            {
                MessageBox.Show("请先添加 Excel 文件", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var invalid = _files.FirstOrDefault(f => !File.Exists(f.FilePath));
            if (invalid != null)
            {
                MessageBox.Show($"文件不存在: {invalid.FilePath}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (IsMergeMode && _files.Count < 2)
            {
                var ok = MessageBox.Show("当前只有 1 个文件，合并后与原文件内容一致，仍要继续吗？",
                    "确认", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                if (ok != MessageBoxResult.OK) return;
            }

            var options = new ExcelSplitMergeOptions
            {
                Mode = IsMergeMode ? ExcelSplitMergeMode.MergeToOne : ExcelSplitMergeMode.SplitBySheet,
                SkipHiddenSheets = ChkSkipHidden.IsChecked == true,
                SkipEmptySheets = ChkSkipEmpty.IsChecked == true,
                KeepSheetNames = ChkKeepNames.IsChecked == true
            };

            var snapshot = _files.ToList();
            SetBusy(true);
            TxtStatus.Text = "准备中...";
            try
            {
                if (options.Mode == ExcelSplitMergeMode.MergeToOne)
                    await Task.Run(() => DoMerge(snapshot, options));
                else
                    await Task.Run(() => DoSplit(snapshot, options));
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "处理失败: " + ex.Message;
                MessageBox.Show(ex.Message, "处理错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetBusy(false);
                ReleaseMemory();
            }
        }

        private void SetBusy(bool busy)
        {
            BtnExecute.IsEnabled = !busy;
            BtnAddFiles.IsEnabled = !busy;
            RdoSplit.IsEnabled = !busy;
            RdoMerge.IsEnabled = !busy;
            ChkSkipHidden.IsEnabled = !busy;
            ChkSkipEmpty.IsEnabled = !busy;
            if (busy) ChkKeepNames.IsEnabled = false;
            else UpdateModeControls();
        }

        /// <summary>拆分：逐文件处理，每个工作簿输出多个文件。</summary>
        private void DoSplit(List<ExcelSplitMergeItem> items, ExcelSplitMergeOptions options)
        {
            string dir = AppConfig.GetFeatureOutputDir("Excel拆分合并");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            int total = items.Count;
            int success = 0, outputFiles = 0, skippedEmpty = 0, skippedHidden = 0;
            var failures = new List<string>();

            for (int i = 0; i < total; i++)
            {
                int index = i;
                var item = items[i];
                Dispatcher.Invoke(() =>
                {
                    TxtStatus.Text = $"正在拆分 ({index + 1}/{total}) {item.FileName}";
                    item.StatusText = "拆分中...";
                });

                try
                {
                    using (var conv = new ExcelSplitMergeConverter())
                    {
                        var r = conv.Split(item.FilePath, dir, options);
                        Dispatcher.Invoke(() =>
                        {
                            item.StatusText = $"完成（{r.OutputFiles.Count} 个文件）";
                            success++;
                            outputFiles += r.OutputFiles.Count;
                            skippedEmpty += r.SkippedEmpty;
                            skippedHidden += r.SkippedHidden;
                        });
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        item.StatusText = "失败";
                        failures.Add($"{item.FileName}: " + Shorten(ex.Message));
                    });
                }
            }

            Dispatcher.Invoke(() =>
            {
                TxtStatus.Text = failures.Count > 0
                    ? $"完成：成功 {success} 个，失败 {failures.Count} 个"
                    : $"全部完成：{success} 个文件，共导出 {outputFiles} 个工作表文件";

                var msg = $"拆分完成\n\n处理文件：{success} 个\n导出工作表：{outputFiles} 个"
                        + (skippedEmpty > 0 ? $"\n跳过空表：{skippedEmpty} 张" : "")
                        + (skippedHidden > 0 ? $"\n跳过隐藏表：{skippedHidden} 张" : "")
                        + (failures.Count > 0 ? $"\n失败：{failures.Count} 个" : "")
                        + $"\n\n输出目录：\n{dir}";

                if (failures.Count > 0)
                {
                    msg += "\n\n失败原因：\n" + string.Join("\n", failures.Take(6))
                         + (failures.Count > 6 ? $"\n… 共 {failures.Count} 个" : "");
                }

                MessageBox.Show(msg, "拆分结果", MessageBoxButton.OK,
                    failures.Count > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);

                if (outputFiles > 0) { try { Process.Start(dir); } catch { } }
            });
        }

        /// <summary>合并：所有文件汇总到一个工作簿。</summary>
        private void DoMerge(List<ExcelSplitMergeItem> items, ExcelSplitMergeOptions options)
        {
            string dir = AppConfig.GetFeatureOutputDir("Excel拆分合并");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            Dispatcher.Invoke(() =>
            {
                TxtStatus.Text = $"正在合并 {items.Count} 个文件...";
                foreach (var it in items) it.StatusText = "合并中...";
            });

            try
            {
                ExcelMergeResult result;
                using (var conv = new ExcelSplitMergeConverter())
                {
                    result = conv.Merge(items.Select(i => i.FilePath).ToList(), dir, options);
                }

                Dispatcher.Invoke(() =>
                {
                    foreach (var it in items) it.StatusText = "已合并";
                    TxtStatus.Text = "合并完成：" + Path.GetFileName(result.OutputFile);

                    var sheetList = result.SheetNames.Count <= 8
                        ? string.Join("\n", result.SheetNames)
                        : string.Join("\n", result.SheetNames.Take(8)) + $"\n… 共 {result.SheetNames.Count} 张表";

                    MessageBox.Show(
                        $"合并完成\n\n合并后工作表（{result.SheetNames.Count} 张）：\n{sheetList}\n\n输出文件：\n{result.OutputFile}",
                        "合并结果", MessageBoxButton.OK, MessageBoxImage.Information);
                    try { Process.Start(dir); } catch { }
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    foreach (var it in items) it.StatusText = "失败";
                    TxtStatus.Text = "合并失败: " + ex.Message;
                    MessageBox.Show(ex.Message, "合并错误", MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }

        private static string Shorten(string s)
        {
            if (string.IsNullOrEmpty(s)) return "未知错误";
            return s.Length > 80 ? s.Substring(0, 80) + "…" : s;
        }

        private void UpdateEmptyHint()
        {
            EmptyOverlay.Visibility = _files.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

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

        [DllImport("psapi.dll")]
        private static extern bool EmptyWorkingSet(IntPtr hProcess);
    }
}
