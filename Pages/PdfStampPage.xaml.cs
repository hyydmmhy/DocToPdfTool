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
    /// <summary>PDF 盖章页面的列表项。</summary>
    public class PdfStampItem : INotifyPropertyChanged
    {
        private string _originalText = "—";
        private string _resultText = "—";
        private string _statusText = "待盖章";

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

    public partial class PdfStampPage : UserControl
    {
        private readonly ObservableCollection<PdfStampItem> _files =
            new ObservableCollection<PdfStampItem>();

        private string _sealPath;

        public PdfStampPage()
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
                Title = "选择要盖章的PDF文件",
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
            if (sender is FrameworkElement el && el.Tag is PdfStampItem item)
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

                _files.Add(new PdfStampItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p,
                    OriginalText = size > 0 ? FormatSize(size) : "—",
                    StatusText = PdfEncryptConverter.IsEncrypted(p) ? "已加密" : "待盖章"
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

        // ───────────────── 章图片 ─────────────────

        private void BtnSelectSeal_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog
            {
                Title = "选择章图片（建议透明底 PNG）",
                Filter = "图片文件|*.png;*.jpg;*.jpeg;*.bmp;*.gif|所有文件 (*.*)|*.*"
            };
            if (d.ShowDialog() == true)
            {
                _sealPath = d.FileName;
                TxtSealName.Text = Path.GetFileName(_sealPath);
            }
        }

        // ───────────────── 执行 ─────────────────

        private async void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            if (_files.Count == 0)
            {
                MessageBox.Show("请先添加PDF文件", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (string.IsNullOrEmpty(_sealPath) || !File.Exists(_sealPath))
            {
                MessageBox.Show("请先选择章图片", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var invalid = _files.FirstOrDefault(f => !File.Exists(f.FilePath));
            if (invalid != null)
            {
                MessageBox.Show($"文件不存在: {invalid.FilePath}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var options = new PdfStampOptions
            {
                ImagePath = _sealPath,
                Position = (PdfStampPosition)CbPosition.SelectedIndex,
                WidthRatio = GetDouble(TxtWidthRatio, 18, 5, 50) / 100.0,
                Rotation = GetRotation(),
                PageRange = TxtPageRange.Text.Trim()
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
                TxtStatus.Text = "盖章失败: " + ex.Message;
                MessageBox.Show(ex.Message, "盖章错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetBusy(false);
                ReleaseMemory();
            }
        }

        private double GetRotation()
        {
            if (CbRotation.SelectedIndex < 0) return 0;
            switch (CbRotation.SelectedIndex)
            {
                case 1: return 15;
                case 2: return 30;
                case 3: return 45;
                case 4: return 90;
                case 5: return 180;
                default: return 0;
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
            BtnSelectSeal.IsEnabled = !busy;
            CbPosition.IsEnabled = !busy;
            CbRotation.IsEnabled = !busy;
            TxtWidthRatio.IsEnabled = !busy;
            TxtPageRange.IsEnabled = !busy;
        }

        private void DoWork(List<PdfStampItem> items, PdfStampOptions options)
        {
            string dir = AppConfig.GetFeatureOutputDir("PDF盖章");
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
                    TxtStatus.Text = $"正在盖章 ({index + 1}/{total}) {item.FileName}";
                    item.StatusText = "盖章中...";
                });

                try
                {
                    using (var converter = new PdfStampConverter())
                    {
                        var r = converter.Stamp(item.FilePath, dir, options);

                        Dispatcher.Invoke(() =>
                        {
                            item.ResultText = FormatSize(r.NewSize);
                            item.StatusText = $"已盖章 {r.StampedPages}/{r.PageCount} 页";
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
                TxtStatus.Text = $"完成：盖章 {success} 个，失败 {failures.Count} 个";

                // 完成不弹窗，自动打开输出目录（仅失败时弹错误提示）
                if (success > 0)
                {
                    try { Process.Start(dir); } catch { }
                }

                if (failures.Count > 0)
                {
                    var msg = "部分文件盖章失败：\n"
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
