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
    /// <summary>图片提取页面的列表项。</summary>
    public class PdfImageExtractItem : INotifyPropertyChanged
    {
        private string _pageText = "—";
        private string _resultText = "—";
        private string _statusText = "待处理";

        public string FileName { get; set; }
        public string FilePath { get; set; }

        public string PageText
        {
            get => _pageText;
            set { if (_pageText != value) { _pageText = value; OnPropertyChanged(); } }
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

    public partial class PdfImageExtractPage : UserControl
    {
        private readonly ObservableCollection<PdfImageExtractItem> _files =
            new ObservableCollection<PdfImageExtractItem>();

        public PdfImageExtractPage()
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
                Title = "选择要提取图片的PDF",
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
            if (sender is FrameworkElement el && el.Tag is PdfImageExtractItem item)
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

                bool encrypted = PdfEncryptConverter.IsEncrypted(p);
                int pages = encrypted ? 0 : PdfEncryptConverter.TryGetPageCount(p);
                _files.Add(new PdfImageExtractItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p,
                    PageText = pages > 0 ? pages + " 页" : "—",
                    StatusText = encrypted ? "已加密，需先解密" : "待处理"
                });
            }
            UpdateEmptyHint();
        }

        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 5) return;
            double total = FileListView.ActualWidth - 14;
            double fixedWidth = 70 + 160 + 110 + 60;
            double name = Math.Max(140, total - fixedWidth);
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

            var options = new PdfImageExtractOptions
            {
                Format = GetFormat(),
                MinSize = GetMinSize(),
                SkipDuplicates = ChkSkipDuplicates.IsChecked == true
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
                TxtStatus.Text = "提取失败: " + ex.Message;
                MessageBox.Show(ex.Message, "提取错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetBusy(false);
                ReleaseMemory();
            }
        }

        private PdfImageOutputFormat GetFormat()
        {
            switch (CmbFormat.SelectedIndex)
            {
                case 1: return PdfImageOutputFormat.Png;
                case 2: return PdfImageOutputFormat.Jpeg;
                default: return PdfImageOutputFormat.Original;
            }
        }

        private int GetMinSize()
        {
            if (int.TryParse((TxtMinSize.Text ?? "").Trim(), out int v) && v >= 0 && v <= 10000)
                return v;
            return 64;
        }

        private void SetBusy(bool busy)
        {
            BtnExecute.IsEnabled = !busy;
            BtnAddFiles.IsEnabled = !busy;
            CmbFormat.IsEnabled = !busy;
            TxtMinSize.IsEnabled = !busy;
            ChkSkipDuplicates.IsEnabled = !busy;
        }

        private void DoWork(List<PdfImageExtractItem> items, PdfImageExtractOptions options)
        {
            string root = AppConfig.GetFeatureOutputDir("PDF提取图片");
            if (!Directory.Exists(root)) Directory.CreateDirectory(root);

            int total = items.Count;
            int success = 0, totalImages = 0;
            var failures = new List<string>();
            string lastDir = root;

            for (int i = 0; i < total; i++)
            {
                int index = i;
                var item = items[i];
                Dispatcher.Invoke(() =>
                {
                    TxtStatus.Text = $"正在提取 ({index + 1}/{total}) {item.FileName}";
                    item.StatusText = "提取中...";
                });

                try
                {
                    using (var extractor = new PdfImageExtractor())
                    {
                        var r = extractor.Extract(item.FilePath, root, options);
                        lastDir = r.OutputDir;

                        Dispatcher.Invoke(() =>
                        {
                            item.ResultText = $"{r.SavedImages} 张";
                            item.StatusText = r.SavedImages > 0 ? "完成" : "未找到图片";
                            totalImages += r.SavedImages;
                            success++;
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
                    ? $"完成：{success} 个文件，共 {totalImages} 张图，失败 {failures.Count} 个"
                    : $"全部完成：{success} 个文件，共提取 {totalImages} 张图";

                if (failures.Count > 0)
                {
                    MessageBox.Show(
                        $"成功 {success} 个，共 {totalImages} 张图\n输出目录:\n{lastDir}\n\n失败 {failures.Count} 个:\n{FormatList(failures)}",
                        "部分文件失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    try { Process.Start(lastDir); } catch { }
                }
            });
        }

        private static string Shorten(string s)
        {
            if (string.IsNullOrEmpty(s)) return "未知错误";
            return s.Length > 80 ? s.Substring(0, 80) + "…" : s;
        }

        private static string FormatList(List<string> items)
        {
            return string.Join("\n", items.Take(8))
                 + (items.Count > 8 ? $"\n… 共 {items.Count} 个" : "");
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
