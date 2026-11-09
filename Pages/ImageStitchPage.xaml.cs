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
    /// <summary>图片拼接页面的列表项。</summary>
    public class ImageStitchItem : INotifyPropertyChanged
    {
        private string _sizeText = "—";
        private string _originalText = "—";
        private string _statusText = "待拼接";

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

        public string StatusText
        {
            get => _statusText;
            set { if (_statusText != value) { _statusText = value; OnPropertyChanged(); } }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public partial class ImageStitchPage : UserControl
    {
        private static readonly string[] ImageExtensions =
            { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff" };

        private readonly ObservableCollection<ImageStitchItem> _files =
            new ObservableCollection<ImageStitchItem>();

        public ImageStitchPage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;
            UpdateEmptyHint();
            UpdateControls();
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
            if (sender is FrameworkElement el && el.Tag is ImageStitchItem item)
            {
                _files.Remove(item);
                UpdateEmptyHint();
            }
        }

        private void BtnMoveUp_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement el && el.Tag is ImageStitchItem item)
            {
                int index = _files.IndexOf(item);
                if (index > 0) _files.Move(index, index - 1);
            }
        }

        private void BtnMoveDown_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement el && el.Tag is ImageStitchItem item)
            {
                int index = _files.IndexOf(item);
                if (index >= 0 && index < _files.Count - 1) _files.Move(index, index + 1);
            }
        }

        private void AddFiles(IEnumerable<string> paths)
        {
            foreach (var p in paths)
            {
                var ext = Path.GetExtension(p);
                if (!ImageExtensions.Any(e => string.Equals(e, ext, StringComparison.OrdinalIgnoreCase))) continue;
                if (_files.Any(f => f.FilePath.Equals(p, StringComparison.OrdinalIgnoreCase))) continue;

                var item = new ImageStitchItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p
                };

                try
                {
                    var size = new FileInfo(p).Length;
                    item.OriginalText = ImageStitchConverter.FormatSize(size);
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
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 4) return;
            double total = FileListView.ActualWidth - 14;
            double fixedWidth = 90 + 80 + 150;
            double name = Math.Max(120, total - fixedWidth);
            if (Math.Abs(gv.Columns[0].Width - name) > 1) gv.Columns[0].Width = name;
        }

        // ───────────────── 参数联动 ─────────────────

        private void WidthMode_Changed(object sender, SelectionChangedEventArgs e)
        {
            UpdateControls();
        }

        private void Format_Changed(object sender, SelectionChangedEventArgs e)
        {
            UpdateControls();
        }

        private void UpdateControls()
        {
            if (CmbWidthMode == null || TxtCustomWidth == null || ChkTransparent == null
                || CmbFormat == null || TxtQuality == null || LblQuality == null) return;

            bool custom = CmbWidthMode.SelectedIndex == 3;
            TxtCustomWidth.IsEnabled = custom;
            TxtCustomWidth.Opacity = custom ? 1.0 : 0.45;

            bool png = CmbFormat.SelectedIndex == 1;
            ChkTransparent.IsEnabled = png;
            ChkTransparent.Opacity = png ? 1.0 : 0.45;
            if (!png) ChkTransparent.IsChecked = false;

            TxtQuality.IsEnabled = !png;
            TxtQuality.Opacity = png ? 0.45 : 1.0;
            LblQuality.Opacity = png ? 0.45 : 1.0;
        }

        private ImageStitchDirection GetDirection()
        {
            return CmbDirection.SelectedIndex == 1
                ? ImageStitchDirection.Horizontal
                : ImageStitchDirection.Vertical;
        }

        private ImageStitchWidthMode GetWidthMode()
        {
            switch (CmbWidthMode.SelectedIndex)
            {
                case 1: return ImageStitchWidthMode.Widest;
                case 2: return ImageStitchWidthMode.Narrowest;
                case 3: return ImageStitchWidthMode.Custom;
                default: return ImageStitchWidthMode.Original;
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
            if (_files.Count < 2)
            {
                MessageBox.Show("请至少添加两张图片", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var invalid = _files.FirstOrDefault(f => !File.Exists(f.FilePath));
            if (invalid != null)
            {
                MessageBox.Show($"文件不存在: {invalid.FilePath}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var options = new ImageStitchOptions
            {
                Direction = GetDirection(),
                WidthMode = GetWidthMode(),
                CustomWidth = GetInt(TxtCustomWidth, 1920, 1, 20000),
                Spacing = GetInt(TxtSpacing, 0, 0, 500),
                TransparentBackground = ChkTransparent.IsChecked == true,
                OutputPng = CmbFormat.SelectedIndex == 1,
                JpegQuality = GetInt(TxtQuality, 92, 1, 100)
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
                TxtStatus.Text = "拼接失败: " + ex.Message;
                MessageBox.Show(ex.Message, "拼接错误", MessageBoxButton.OK, MessageBoxImage.Error);
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
            CmbDirection.IsEnabled = !busy;
            CmbWidthMode.IsEnabled = !busy;
            CmbFormat.IsEnabled = !busy;
            ChkTransparent.IsEnabled = !busy;
            if (busy)
            {
                TxtCustomWidth.IsEnabled = false;
                TxtQuality.IsEnabled = false;
            }
            else
            {
                UpdateControls();
            }
        }

        private void DoWork(List<ImageStitchItem> items, ImageStitchOptions options)
        {
            string dir = AppConfig.GetFeatureOutputDir("图片拼接");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            // 所有 UI 访问必须在 UI 线程通过 Dispatcher.Invoke 完成（DoWork 运行在后台线程）
            Dispatcher.Invoke(() =>
            {
                TxtStatus.Text = "正在拼接 " + items.Count + " 张图片...";
            });

            try
            {
                using (var converter = new ImageStitchConverter())
                {
                    var paths = items.Select(i => i.FilePath).ToList();
                    var r = converter.Stitch(paths, dir, options);

                    Dispatcher.Invoke(() =>
                    {
                        TxtStatus.Text = $"拼接完成：{r.Width}×{r.Height}，{ImageStitchConverter.FormatSize(r.Size)}";

                        // 完成不弹窗，直接自动打开输出文件夹
                        try { Process.Start(dir); } catch { }
                    });
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    TxtStatus.Text = "拼接失败: " + Shorten(ex.Message);
                });
                throw;
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
