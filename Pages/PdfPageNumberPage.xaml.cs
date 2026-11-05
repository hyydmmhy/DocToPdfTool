using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
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
    /// <summary>加页码页面的列表项。</summary>
    public class PdfPageNumberItem : INotifyPropertyChanged
    {
        private string _pageText = "—";
        private string _statusText = "待处理";

        public string FileName { get; set; }
        public string FilePath { get; set; }

        public string PageText
        {
            get => _pageText;
            set { if (_pageText != value) { _pageText = value; OnPropertyChanged(); } }
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

    public partial class PdfPageNumberPage : UserControl
    {
        private readonly ObservableCollection<PdfPageNumberItem> _files =
            new ObservableCollection<PdfPageNumberItem>();

        public PdfPageNumberPage()
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
                Title = "选择要加页码的PDF",
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
            if (sender is FrameworkElement el && el.Tag is PdfPageNumberItem item)
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
                _files.Add(new PdfPageNumberItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p,
                    PageText = pages > 0 ? pages + " 页" : "—",
                    StatusText = encrypted ? "已加密" : "待处理"
                });
            }
            UpdateEmptyHint();
        }

        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 4) return;
            double total = FileListView.ActualWidth - 14;
            double fixedWidth = 70 + 120 + 72;
            double name = Math.Max(100, total - fixedWidth);
            if (Math.Abs(gv.Columns[0].Width - name) > 1) gv.Columns[0].Width = name;
        }

        // ───────────────── 参数读取 ─────────────────

        private PageNumberPosition GetPosition()
        {
            switch (CmbPosition.SelectedIndex)
            {
                case 1: return PageNumberPosition.BottomLeft;
                case 2: return PageNumberPosition.BottomRight;
                case 3: return PageNumberPosition.TopCenter;
                case 4: return PageNumberPosition.TopLeft;
                case 5: return PageNumberPosition.TopRight;
                default: return PageNumberPosition.BottomCenter;
            }
        }

        private PageNumberFormat GetFormat()
        {
            switch (CmbFormat.SelectedIndex)
            {
                case 1: return PageNumberFormat.WithPrefix;
                case 2: return PageNumberFormat.WithTotal;
                case 3: return PageNumberFormat.Dash;
                default: return PageNumberFormat.Plain;
            }
        }

        private float GetMarginPt()
        {
            switch (CmbMargin.SelectedIndex)
            {
                case 1: return 42.5f;   // 15 mm
                case 2: return 56.7f;   // 20 mm
                case 3: return 70.9f;   // 25 mm
                default: return 28.3f;  // 10 mm
            }
        }

        private int GetStartNumber()
        {
            if (int.TryParse((TxtStart.Text ?? "").Trim(), out int v) && v >= 0 && v <= 99999) return v;
            return 1;
        }

        private float GetFontSize()
        {
            if (float.TryParse((TxtFontSize.Text ?? "").Trim(), out float v) && v >= 5 && v <= 72) return v;
            return 10f;
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

            var options = new PdfPageNumberOptions
            {
                Position = GetPosition(),
                Format = GetFormat(),
                StartNumber = GetStartNumber(),
                SkipFirstPages = ChkSkipFirst.IsChecked == true ? 1 : 0,
                PageRanges = (TxtPageRanges.Text ?? "").Trim(),
                FontSize = GetFontSize(),
                Color = CmbColor.SelectedIndex == 1 ? Color.Gray : Color.Black,
                MarginPt = GetMarginPt()
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
            CmbPosition.IsEnabled = !busy;
            CmbFormat.IsEnabled = !busy;
            CmbMargin.IsEnabled = !busy;
            CmbColor.IsEnabled = !busy;
            TxtStart.IsEnabled = !busy;
            TxtFontSize.IsEnabled = !busy;
            ChkSkipFirst.IsEnabled = !busy;
            TxtPageRanges.IsEnabled = !busy;
        }

        private void DoWork(List<PdfPageNumberItem> items, PdfPageNumberOptions options)
        {
            string dir = AppConfig.GetFeatureOutputDir("PDF加页码");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            int total = items.Count;
            int success = 0, numberedPages = 0;
            var failures = new List<string>();

            for (int i = 0; i < total; i++)
            {
                int index = i;
                var item = items[i];
                Dispatcher.Invoke(() =>
                {
                    TxtStatus.Text = $"正在处理 ({index + 1}/{total}) {item.FileName}";
                    item.StatusText = "处理中...";
                });

                try
                {
                    using (var converter = new PdfPageNumberConverter())
                    {
                        int applied = converter.Apply(item.FilePath, dir, options);
                        Dispatcher.Invoke(() =>
                        {
                            item.StatusText = $"完成（{applied} 页）";
                            success++;
                            numberedPages += applied;
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
                    : $"全部完成：{success} 个文件，共编号 {numberedPages} 页";

                if (failures.Count > 0)
                {
                    MessageBox.Show(
                        $"成功 {success} 个，共编号 {numberedPages} 页\n输出目录:\n{dir}\n\n失败 {failures.Count} 个:\n{FormatList(failures)}",
                        "部分文件失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    try { Process.Start(dir); } catch { }
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
