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
    /// <summary>图片批处理页面的列表项。</summary>
    public class ImageBatchItem : INotifyPropertyChanged
    {
        private string _sizeText = "—";
        private string _originalText = "—";
        private string _resultText = "—";
        private string _statusText = "待处理";

        public string FileName { get; set; }
        public string FilePath { get; set; }

        public string SizeText
        {
            get => _sizeText;
            set { if (_sizeText != value) { _sizeText = value; OnPropertyChanged(); } }
        }

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

    public partial class ImageBatchPage : UserControl
    {
        private static readonly string[] ImageExtensions =
            { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff" };

        private readonly ObservableCollection<ImageBatchItem> _files =
            new ObservableCollection<ImageBatchItem>();

        public ImageBatchPage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;
            UpdateEmptyHint();
            UpdateResizeControls();
        }

        // ───────────────── 文件列表 ─────────────────

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog
            {
                Title = "选择图片",
                Filter = "图片文件 (*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff)|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff|所有文件 (*.*)|*.*",
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
            if (sender is FrameworkElement el && el.Tag is ImageBatchItem item)
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
                if (!ImageExtensions.Any(e => string.Equals(e, ext, StringComparison.OrdinalIgnoreCase))) continue;
                if (_files.Any(f => f.FilePath.Equals(p, StringComparison.OrdinalIgnoreCase))) continue;

                var item = new ImageBatchItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p
                };

                try
                {
                    var size = new FileInfo(p).Length;
                    item.OriginalText = ImageBatchConverter.FormatSize(size);
                    using (var image = System.Drawing.Image.FromFile(p))
                    {
                        item.SizeText = image.Width + "×" + image.Height;
                    }
                }
                catch
                {
                    item.StatusText = "无法读取";
                    item.SizeText = "—";
                }

                _files.Add(item);
            }
            UpdateEmptyHint();
        }

        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 6) return;
            double total = FileListView.ActualWidth - 14;
            double fixedWidth = 90 + 80 + 120 + 80 + 72;
            double name = Math.Max(110, total - fixedWidth);
            if (Math.Abs(gv.Columns[0].Width - name) > 1) gv.Columns[0].Width = name;
        }

        // ───────────────── 参数联动 ─────────────────

        private bool IsResizeEnabled => CmbResize.SelectedIndex > 0;

        private void Resize_Changed(object sender, SelectionChangedEventArgs e)
        {
            UpdateResizeControls();
        }

        private void UpdateResizeControls()
        {
            if (TxtResizeValue == null || TxtResizeUnit == null) return;

            bool enabled = IsResizeEnabled;
            TxtResizeValue.IsEnabled = enabled;
            TxtResizeValue.Opacity = enabled ? 1.0 : 0.45;

            TxtResizeUnit.Text = CmbResize.SelectedIndex == 4 ? "%" : "px";

            if (enabled && string.IsNullOrWhiteSpace(TxtResizeValue.Text))
                TxtResizeValue.Text = CmbResize.SelectedIndex == 4 ? "50" : "1920";
        }

        private ImageTargetFormat GetFormat()
        {
            switch (CmbFormat.SelectedIndex)
            {
                case 1: return ImageTargetFormat.Jpeg;
                case 2: return ImageTargetFormat.Png;
                case 3: return ImageTargetFormat.Bmp;
                default: return ImageTargetFormat.KeepOriginal;
            }
        }

        private ImageResizeMode GetResizeMode()
        {
            switch (CmbResize.SelectedIndex)
            {
                case 1: return ImageResizeMode.ByWidth;
                case 2: return ImageResizeMode.ByHeight;
                case 3: return ImageResizeMode.ByLongSide;
                case 4: return ImageResizeMode.ByPercent;
                default: return ImageResizeMode.None;
            }
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
                MessageBox.Show("请先添加图片", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var invalid = _files.FirstOrDefault(f => !File.Exists(f.FilePath));
            if (invalid != null)
            {
                MessageBox.Show($"文件不存在: {invalid.FilePath}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var options = new ImageBatchOptions
            {
                Format = GetFormat(),
                ResizeMode = GetResizeMode(),
                ResizeValue = GetInt(TxtResizeValue, 1920, 1, 20000),
                JpegQuality = GetInt(TxtQuality, 92, 1, 100),
                TargetSizeKb = GetInt(TxtTargetKb, 0, 0, 1024 * 100),
                AutoOrient = ChkAutoOrient.IsChecked == true
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
            CmbFormat.IsEnabled = !busy;
            CmbResize.IsEnabled = !busy;
            TxtQuality.IsEnabled = !busy;
            TxtTargetKb.IsEnabled = !busy;
            ChkAutoOrient.IsEnabled = !busy;
            if (busy) TxtResizeValue.IsEnabled = false;
            else UpdateResizeControls();
        }

        private void DoWork(List<ImageBatchItem> items, ImageBatchOptions options)
        {
            string dir = AppConfig.GetFeatureOutputDir("图片批量处理");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            int total = items.Count;
            int success = 0, copied = 0;
            long totalBefore = 0, totalAfter = 0;
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
                    using (var converter = new ImageBatchConverter())
                    {
                        var r = converter.Process(item.FilePath, dir, options);

                        Dispatcher.Invoke(() =>
                        {
                            item.ResultText = ImageBatchConverter.FormatSize(r.NewSize)
                                            + "（" + r.Width + "×" + r.Height + "）";
                            item.StatusText = r.Copied ? "已复制" : "完成";
                            success++;
                            if (r.Copied) copied++;
                            totalBefore += r.OriginalSize;
                            totalAfter += r.NewSize;
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
                string ratio = totalBefore > 0
                    ? $"（{totalAfter * 100.0 / totalBefore:0.0}%）"
                    : "";
                TxtStatus.Text = failures.Count > 0
                    ? $"完成：成功 {success} 个，失败 {failures.Count} 个"
                    : $"全部完成：{success} 张，共 {ImageBatchConverter.FormatSize(totalAfter)}{ratio}";

                var msg = $"处理完成\n\n成功：{success} 张"
                        + (copied > 0 ? $"（其中 {copied} 张直接复制）" : "")
                        + $"\n体积：{ImageBatchConverter.FormatSize(totalBefore)} → {ImageBatchConverter.FormatSize(totalAfter)} {ratio}"
                        + (failures.Count > 0 ? $"\n失败：{failures.Count} 张" : "")
                        + $"\n\n输出目录：\n{dir}";

                if (failures.Count > 0)
                {
                    msg += "\n\n失败原因：\n" + string.Join("\n", failures.Take(6))
                         + (failures.Count > 6 ? $"\n… 共 {failures.Count} 张" : "");
                }

                MessageBox.Show(msg, "处理结果", MessageBoxButton.OK,
                    failures.Count > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);

                if (success > 0) { try { Process.Start(dir); } catch { } }
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
