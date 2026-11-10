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
    /// <summary>PDF 裁剪页面的列表项。</summary>
    public class PdfCropItem : INotifyPropertyChanged
    {
        private string _originalText = "—";
        private string _resultText = "—";
        private string _statusText = "待裁剪";

        public string FileName { get; set; }
        public string FilePath { get; set; }

        public string OriginalText
        {
            get => _originalText;
            set { if (_originalText != value) { _originalText = value; OnPropertyChanged(); } }
        }

        public string ResultText
        {
            get => _resultText;
            set { if (_resultText != value) { _resultText = value; OnPropertyChanged(); } }
        }

        public string StatusText
        {
            get => _statusText;
            set { if (_statusText != value) { _statusText = value; OnPropertyChanged(); } }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public partial class PdfCropPage : UserControl
    {
        private readonly ObservableCollection<PdfCropItem> _files =
            new ObservableCollection<PdfCropItem>();

        public PdfCropPage()
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
                Title = "选择要裁剪的PDF文件",
                Filter = "PDF 文档 (*.pdf)|*.pdf|所有文件 (*.*)|*.*",
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
            if (sender is FrameworkElement el && el.Tag is PdfCropItem item)
            {
                _files.Remove(item);
                UpdateEmptyHint();
            }
        }

        private void AddFiles(IEnumerable<string> paths)
        {
            foreach (var p in paths)
            {
                if (!string.Equals(Path.GetExtension(p), ".pdf", StringComparison.OrdinalIgnoreCase)) continue;
                if (_files.Any(f => f.FilePath.Equals(p, StringComparison.OrdinalIgnoreCase))) continue;

                long size = 0;
                try { size = new FileInfo(p).Length; } catch { }

                _files.Add(new PdfCropItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p,
                    OriginalText = size > 0 ? PdfCropConverter.FormatSize(size) : "—",
                    StatusText = PdfEncryptConverter.IsEncrypted(p) ? "已加密" : "待裁剪"
                });
            }
            UpdateEmptyHint();
        }

        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 5) return;
            double total = FileListView.ActualWidth - 14;
            double fixedWidth = 90 + 90 + 110 + 72;
            double name = Math.Max(110, total - fixedWidth);
            if (Math.Abs(gv.Columns[0].Width - name) > 1) gv.Columns[0].Width = name;
        }

        // ───────────────── 模式联动 ─────────────────

        private void Mode_Changed(object sender, RoutedEventArgs e)
        {
            UpdateModeControls();
        }

        private void UpdateModeControls()
        {
            if (AutoPanel == null || ManualPanel == null || TxtHint == null) return;

            bool auto = RdoAuto.IsChecked == true;
            AutoPanel.Visibility = auto ? Visibility.Visible : Visibility.Collapsed;
            ManualPanel.Visibility = auto ? Visibility.Collapsed : Visibility.Visible;
            TxtHint.Text = auto
                ? "自动检测每页内容边界，裁掉四周白边；矢量无损，文字层保留可搜索"
                : "手动输入四边裁剪量（毫米），适用于边距固定的文档";
        }

        private PdfCropMode GetMode()
        {
            return RdoAuto.IsChecked == true ? PdfCropMode.Auto : PdfCropMode.Manual;
        }

        private double GetDouble(TextBox box, double fallback, double min, double max)
        {
            if (double.TryParse((box.Text ?? "").Trim(), out double v) && v >= min && v <= max) return v;
            return fallback;
        }

        private int GetInt(TextBox box, int fallback, int min, int max)
        {
            if (int.TryParse((box.Text ?? "").Trim(), out int v) && v >= min && v <= max) return v;
            return fallback;
        }

        // ───────────────── 执行 ─────────────────

        private async void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            if (_files.Count == 0)
            {
                MessageBox.Show("请先添加PDF文件", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var invalid = _files.FirstOrDefault(f => !File.Exists(f.FilePath));
            if (invalid != null)
            {
                MessageBox.Show($"文件不存在: {invalid.FilePath}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var options = new PdfCropOptions
            {
                Mode = GetMode(),
                WhiteThreshold = GetInt(TxtWhiteThreshold, 245, 100, 254),
                AutoMarginMm = GetDouble(TxtAutoMargin, 2, 0, 50),
                TopMm = GetDouble(TxtTopMm, 5, 0, 500),
                BottomMm = GetDouble(TxtBottomMm, 5, 0, 500),
                LeftMm = GetDouble(TxtLeftMm, 5, 0, 500),
                RightMm = GetDouble(TxtRightMm, 5, 0, 500)
            };

            var snapshot = _files.ToList();
            SetBusy(true);
            TxtStatus.Text = "准备中...";
            try
            {
                await Task.Run(() => DoWork(snapshot, options));
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "裁剪失败: " + ex.Message;
                MessageBox.Show(ex.Message, "裁剪错误", MessageBoxButton.OK, MessageBoxImage.Error);
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
            RdoAuto.IsEnabled = !busy;
            RdoManual.IsEnabled = !busy;
            if (busy)
            {
                TxtWhiteThreshold.IsEnabled = false;
                TxtAutoMargin.IsEnabled = false;
                TxtTopMm.IsEnabled = false;
                TxtBottomMm.IsEnabled = false;
                TxtLeftMm.IsEnabled = false;
                TxtRightMm.IsEnabled = false;
            }
            else
            {
                UpdateModeControls();
                TxtWhiteThreshold.IsEnabled = true;
                TxtAutoMargin.IsEnabled = true;
                TxtTopMm.IsEnabled = true;
                TxtBottomMm.IsEnabled = true;
                TxtLeftMm.IsEnabled = true;
                TxtRightMm.IsEnabled = true;
            }
        }

        private void DoWork(List<PdfCropItem> items, PdfCropOptions options)
        {
            string dir = AppConfig.GetFeatureOutputDir("PDF裁剪");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            int total = items.Count;
            int success = 0, skipped = 0;
            var failures = new List<string>();

            for (int i = 0; i < total; i++)
            {
                int index = i;
                var item = items[i];
                Dispatcher.Invoke(() =>
                {
                    TxtStatus.Text = $"正在裁剪 ({index + 1}/{total}) {item.FileName}";
                    item.StatusText = "裁剪中...";
                });

                try
                {
                    using (var converter = new PdfCropConverter())
                    {
                        var r = converter.Crop(item.FilePath, dir, options);

                        Dispatcher.Invoke(() =>
                        {
                            if (r.Skipped)
                            {
                                item.ResultText = "无变化";
                                item.StatusText = "未检测到白边";
                                skipped++;
                            }
                            else
                            {
                                item.ResultText = PdfCropConverter.FormatSize(r.NewSize);
                                item.StatusText = $"已裁剪 {r.AffectedPages}/{r.PageCount} 页";
                                success++;
                            }
                        });
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        item.StatusText = "失败";
                        item.ResultText = "—";
                        failures.Add($"{item.FileName}: " + Shorten(ex.Message));
                    });
                }
            }

            Dispatcher.Invoke(() =>
            {
                TxtStatus.Text = $"完成：裁剪 {success} 个，跳过 {skipped} 个，失败 {failures.Count} 个";

                // 完成不弹窗，自动打开输出目录（仅失败时弹错误提示）
                if (success > 0)
                {
                    try { Process.Start(dir); } catch { }
                }

                if (failures.Count > 0)
                {
                    var msg = "部分文件裁剪失败：\n"
                            + string.Join("\n", failures.Take(6))
                            + (failures.Count > 6 ? $"\n… 共 {failures.Count} 个" : "");
                    MessageBox.Show(msg, "部分文件失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            });
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
