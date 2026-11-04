using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using DocToPdfTool.Utils;
using Microsoft.Win32;

namespace DocToPdfTool.Pages
{
    /// <summary>水印页面的列表项。</summary>
    public class PdfWatermarkItem : INotifyPropertyChanged
    {
        private string _statusText = "待处理";
        private string _pageText = "—";

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

    public partial class PdfWatermarkPage : UserControl
    {
        private readonly ObservableCollection<PdfWatermarkItem> _files =
            new ObservableCollection<PdfWatermarkItem>();

        private string _imagePath;

        public PdfWatermarkPage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;
            UpdateEmptyHint();
            UpdateTypeControls();
        }

        // ───────────────── 文件列表 ─────────────────

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog
            {
                Title = "选择要添加水印的PDF",
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
            if (sender is FrameworkElement el && el.Tag is PdfWatermarkItem item)
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

                int pages = PdfEncryptConverter.IsEncrypted(p) ? 0 : PdfEncryptConverter.TryGetPageCount(p);
                _files.Add(new PdfWatermarkItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p,
                    PageText = pages > 0 ? pages + " 页" : "—",
                    StatusText = PdfEncryptConverter.IsEncrypted(p) ? "已加密，需先解密" : "待处理"
                });
            }
            UpdateEmptyHint();
        }

        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 5) return;
            double total = FileListView.ActualWidth - 14;
            double fixedWidth = 70 + 120 + 60;
            double name = Math.Max(140, (total - fixedWidth) * 0.55);
            double path = Math.Max(80, total - fixedWidth - name);
            if (Math.Abs(gv.Columns[0].Width - name) > 1) gv.Columns[0].Width = name;
            if (Math.Abs(gv.Columns[3].Width - path) > 1) gv.Columns[3].Width = path;
        }

        // ───────────────── 类型联动 ─────────────────

        private bool IsTextMode => RdoText.IsChecked == true;

        private void Type_Changed(object sender, RoutedEventArgs e)
        {
            UpdateTypeControls();
        }

        /// <summary>切换文字/图片类型时只改可用性，不隐藏控件，避免布局跳动。</summary>
        private void UpdateTypeControls()
        {
            if (TxtText == null || BtnPickImage == null || TxtFontSize == null) return;

            bool text = IsTextMode;
            TxtText.IsEnabled = text;
            TxtFontSize.IsEnabled = text;
            TxtText.Opacity = text ? 1.0 : 0.45;
            TxtFontSize.Opacity = text ? 1.0 : 0.45;

            BtnPickImage.IsEnabled = !text;
            BtnPickImage.Opacity = text ? 0.45 : 1.0;

            if (TxtHint != null)
                TxtHint.Text = text
                    ? "水印绘制在原有内容之上，不改变原文字层；加密文件需先移除密码"
                    : "图片水印按页宽比例缩放，建议使用 PNG 透明背景图；加密文件需先移除密码";
        }

        private void BtnPickImage_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog
            {
                Title = "选择水印图片",
                Filter = "图片文件 (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp|所有文件 (*.*)|*.*"
            };
            if (d.ShowDialog() == true)
            {
                _imagePath = d.FileName;
                TxtImageName.Text = Path.GetFileName(d.FileName);
                TxtImageName.Foreground = System.Windows.Media.Brushes.DimGray;
            }
        }

        // ───────────────── 参数读取 ─────────────────

        private Color GetSelectedColor()
        {
            switch (CmbColor.SelectedIndex)
            {
                case 1: return Color.Red;
                case 2: return Color.SteelBlue;
                case 3: return Color.Black;
                default: return Color.Gray;
            }
        }

        private float GetSelectedOpacity()
        {
            switch (CmbOpacity.SelectedIndex)
            {
                case 0: return 0.1f;
                case 1: return 0.2f;
                case 3: return 0.5f;
                case 4: return 0.7f;
                default: return 0.3f;
            }
        }

        private float GetSelectedRotation()
        {
            switch (CmbRotation.SelectedIndex)
            {
                case 0: return 0f;
                case 2: return 30f;
                case 3: return 90f;
                default: return 45f;
            }
        }

        private float GetFontSize()
        {
            if (float.TryParse((TxtFontSize.Text ?? "").Trim(), out float size) && size >= 8 && size <= 300)
                return size;
            return 48f;
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

            bool textMode = IsTextMode;
            if (textMode && string.IsNullOrWhiteSpace(TxtText.Text))
            {
                MessageBox.Show("请输入水印文字", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                TxtText.Focus();
                return;
            }
            if (!textMode && string.IsNullOrEmpty(_imagePath))
            {
                MessageBox.Show("请选择水印图片", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var options = new PdfWatermarkOptions
            {
                Type = textMode ? PdfWatermarkType.Text : PdfWatermarkType.Image,
                Text = (TxtText.Text ?? "").Trim(),
                ImagePath = _imagePath,
                FontSize = GetFontSize(),
                Color = GetSelectedColor(),
                Opacity = GetSelectedOpacity(),
                Rotation = GetSelectedRotation(),
                Tile = ChkTile.IsChecked == true,
                PageRanges = (TxtPageRanges.Text ?? "").Trim()
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
            RdoText.IsEnabled = !busy;
            RdoImage.IsEnabled = !busy;
            CmbColor.IsEnabled = !busy;
            CmbOpacity.IsEnabled = !busy;
            CmbRotation.IsEnabled = !busy;
            ChkTile.IsEnabled = !busy;
            TxtPageRanges.IsEnabled = !busy;
            if (busy)
            {
                TxtText.IsEnabled = false;
                TxtFontSize.IsEnabled = false;
                BtnPickImage.IsEnabled = false;
            }
            else
            {
                UpdateTypeControls();
            }
        }

        private void DoWork(List<PdfWatermarkItem> items, PdfWatermarkOptions options)
        {
            string dir = AppConfig.GetFeatureOutputDir("PDF水印");
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
                    TxtStatus.Text = $"正在处理 ({index + 1}/{total}) {item.FileName}";
                    item.StatusText = "处理中...";
                });

                try
                {
                    using (var converter = new PdfWatermarkConverter())
                    {
                        int affected = converter.Apply(item.FilePath, dir, options);
                        Dispatcher.Invoke(() =>
                        {
                            item.StatusText = $"完成（{affected} 页）";
                            success++;
                        });
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        item.StatusText = "失败";
                        failures.Add($"{item.FileName}: {Shorten(ex.Message)}");
                    });
                }
            }

            Dispatcher.Invoke(() =>
            {
                TxtStatus.Text = failures.Count > 0
                    ? $"完成：成功 {success} 个，失败 {failures.Count} 个"
                    : $"全部完成，共 {success} 个文件";

                if (failures.Count > 0)
                {
                    MessageBox.Show(
                        $"成功 {success} 个，输出目录:\n{dir}\n\n失败 {failures.Count} 个:\n{FormatList(failures)}",
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
