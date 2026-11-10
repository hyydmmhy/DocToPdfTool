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
    /// <summary>PDF 拼版页面的列表项。</summary>
    public class PdfNupItem : INotifyPropertyChanged
    {
        private string _originalText = "—";
        private string _resultText = "—";
        private string _statusText = "待拼版";

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

    public partial class PdfNupPage : UserControl
    {
        private readonly ObservableCollection<PdfNupItem> _files =
            new ObservableCollection<PdfNupItem>();

        public PdfNupPage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;
            UpdateEmptyHint();
        }

        // ───────────────── 文件列表 ─────────────────

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog
            {
                Title = "选择要拼版的PDF文件",
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
            if (sender is FrameworkElement el && el.Tag is PdfNupItem item)
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

                _files.Add(new PdfNupItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p,
                    OriginalText = size > 0 ? FormatSize(size) : "—",
                    StatusText = PdfEncryptConverter.IsEncrypted(p) ? "已加密" : "待拼版"
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

            var options = new PdfNupOptions
            {
                N = GetN(),
                Paper = CbPaper.SelectedIndex == 1 ? PdfNupPaper.A3 : PdfNupPaper.A4,
                Orientation = CbOrientation.SelectedIndex == 1 ? PdfNupOrientation.Landscape : PdfNupOrientation.Portrait,
                Order = CbOrder.SelectedIndex == 1 ? PdfNupOrder.TopToBottom : PdfNupOrder.LeftToRight,
                MarginMm = GetDouble(TxtMargin, 10, 0, 50),
                GapMm = GetDouble(TxtGap, 4, 0, 50),
                DrawCutLines = ChkCutLines.IsChecked == true,
                PageRange = TxtPageRange.Text.Trim(),
                Dpi = GetDpi()
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
                TxtStatus.Text = "拼版失败: " + ex.Message;
                MessageBox.Show(ex.Message, "拼版错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetBusy(false);
                ReleaseMemory();
            }
        }

        private int GetN()
        {
            switch (CbN.SelectedIndex)
            {
                case 0: return 2;
                case 1: return 4;
                case 2: return 6;
                case 3: return 9;
                case 4: return 16;
                default: return 4;
            }
        }

        private int GetDpi()
        {
            switch (CbDpi.SelectedIndex)
            {
                case 1: return 600;
                default: return 300;
            }
        }

        private double GetDouble(TextBox box, double fallback, double min, double max)
        {
            if (double.TryParse((box.Text ?? "").Trim(), out double v) && v >= min && v <= max) return v;
            return fallback;
        }

        private void SetBusy(bool busy)
        {
            BtnExecute.IsEnabled = !busy;
            BtnAddFiles.IsEnabled = !busy;
            CbN.IsEnabled = !busy;
            CbPaper.IsEnabled = !busy;
            CbOrientation.IsEnabled = !busy;
            CbOrder.IsEnabled = !busy;
            CbDpi.IsEnabled = !busy;
            ChkCutLines.IsEnabled = !busy;
            TxtMargin.IsEnabled = !busy;
            TxtGap.IsEnabled = !busy;
            TxtPageRange.IsEnabled = !busy;
        }

        private void DoWork(List<PdfNupItem> items, PdfNupOptions options)
        {
            string dir = AppConfig.GetFeatureOutputDir("PDF拼版");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            int total = items.Count;
            int success = 0;
            var failures = new List<string>();

            for (int i = 0; i < total; i++)
            {
                int index = i;
                var item = items[i];
                Dispatcher.Invoke(() =>
                {
                    TxtStatus.Text = $"正在拼版 ({index + 1}/{total}) {item.FileName}";
                    item.StatusText = "拼版中...";
                });

                try
                {
                    using (var converter = new PdfNupConverter())
                    {
                        var r = converter.Convert(item.FilePath, dir, options);

                        Dispatcher.Invoke(() =>
                        {
                            item.ResultText = FormatSize(r.NewSize);
                            item.StatusText = $"{r.SourcePages}页→{r.OutputPages}页";
                            success++;
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
                TxtStatus.Text = $"完成：拼版 {success} 个，失败 {failures.Count} 个";

                // 完成不弹窗，自动打开输出目录（仅失败时弹错误提示）
                if (success > 0)
                {
                    try { Process.Start(dir); } catch { }
                }

                if (failures.Count > 0)
                {
                    var msg = "部分文件拼版失败：\n"
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

        private static string FormatSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024.0).ToString("0.0") + " KB";
            return (bytes / 1024.0 / 1024.0).ToString("0.00") + " MB";
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
