using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using DocToPdfTool.Utils;
using Microsoft.Win32;

namespace DocToPdfTool.Pages
{
    public partial class ImageToPdfPage : UserControl
    {
        private readonly ObservableCollection<ImageFileItem> _files = new ObservableCollection<ImageFileItem>();

        /// <summary>支持的图片扩展名（系统解码器可识别的格式）。</summary>
        private static readonly HashSet<string> SupportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".jpg", ".jpeg", ".jpe", ".png", ".bmp", ".gif", ".tif", ".tiff"
        };

        public ImageToPdfPage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;
            TxtFileName.Text = "图片转PDF_" + DateTime.Now.ToString("yyyyMMdd");
            UpdateEmptyHint();
        }

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog
            {
                Title = "选择要转换的图片",
                Filter = "图片文件 (*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff)|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff|所有文件 (*.*)|*.*",
                Multiselect = true
            };
            if (d.ShowDialog() == true)
                AddFiles(d.FileNames);
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
            if (sender is FrameworkElement el && el.Tag is ImageFileItem item)
            {
                _files.Remove(item);
                UpdateEmptyHint();
            }
        }

        private void BtnMoveUp_Click(object sender, RoutedEventArgs e) => MoveSelected(-1);
        private void BtnMoveDown_Click(object sender, RoutedEventArgs e) => MoveSelected(1);

        private void MoveSelected(int delta)
        {
            int idx = FileListView.SelectedIndex;
            if (idx < 0) return;
            int target = idx + delta;
            if (target < 0 || target >= _files.Count) return;

            _files.Move(idx, target);
            FileListView.SelectedIndex = target;
            FileListView.ScrollIntoView(FileListView.SelectedItem);
        }

        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 5) return;
            double total = FileListView.ActualWidth - 14;
            double preview = 64;
            double size = 110;
            double action = 60;
            double name = Math.Max(120, (total - preview - size - action) * 0.45);
            double path = Math.Max(80, total - preview - size - action - name);
            if (Math.Abs(gv.Columns[1].Width - name) > 1) gv.Columns[1].Width = name;
            if (Math.Abs(gv.Columns[3].Width - path) > 1) gv.Columns[3].Width = path;
        }

        private async void BtnConvert_Click(object sender, RoutedEventArgs e)
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

            string fileName = (TxtFileName.Text ?? "").Trim();
            if (string.IsNullOrEmpty(fileName)) fileName = "图片转PDF_" + DateTime.Now.ToString("yyyyMMdd");
            foreach (var c in Path.GetInvalidFileNameChars())
                fileName = fileName.Replace(c, '_');
            if (fileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                fileName = fileName.Substring(0, fileName.Length - 4);

            var pageMode = GetSelectedPageMode();
            int marginMm = GetSelectedMargin();
            bool mergeToOne = CmbOutputMode.SelectedIndex != 1;
            var paths = _files.Select(f => f.FilePath).ToList();

            SetBusy(true);
            TxtConversionStatus.Text = "准备中...";
            try
            {
                await Task.Run(() => DoConversion(paths, fileName, pageMode, marginMm, mergeToOne));
            }
            catch (Exception ex)
            {
                TxtConversionStatus.Text = $"转换失败: {ex.Message}";
                MessageBox.Show(ex.Message, "转换错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetBusy(false);
                ReleaseMemory();
            }
        }

        private void SetBusy(bool busy)
        {
            BtnConvert.IsEnabled = !busy;
            BtnAddFiles.IsEnabled = !busy;
            BtnMoveUp.IsEnabled = !busy;
            BtnMoveDown.IsEnabled = !busy;
            CmbPageMode.IsEnabled = !busy;
            CmbMargin.IsEnabled = !busy;
            CmbOutputMode.IsEnabled = !busy;
            UpdateFileNameEnabled(!busy);
        }

        /// <summary>是否为「合成一个 PDF」模式。</summary>
        private bool IsMergeMode => CmbOutputMode.SelectedIndex != 1;

        /// <summary>切换输出方式时同步文件名输入框的可用状态。</summary>
        private void CmbOutputMode_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateFileNameEnabled(BtnConvert?.IsEnabled ?? true);
        }

        /// <summary>
        /// 文件名输入框仅在「合成一个 PDF」模式下可用；
        /// 「每张图各出一个 PDF」时按各图片原始文件名输出，故置灰。
        /// </summary>
        private void UpdateFileNameEnabled(bool idle)
        {
            // XAML 加载期间可能早于 TxtFileName 创建
            if (TxtFileName == null) return;
            bool enabled = idle && IsMergeMode;
            TxtFileName.IsEnabled = enabled;
            TxtFileName.Opacity = enabled ? 1.0 : 0.5;
        }

        private ImagePageMode GetSelectedPageMode()
        {
            switch (CmbPageMode.SelectedIndex)
            {
                case 1: return ImagePageMode.A4Fill;
                case 2: return ImagePageMode.OriginalSize;
                default: return ImagePageMode.A4Fit;
            }
        }

        private int GetSelectedMargin()
        {
            switch (CmbMargin.SelectedIndex)
            {
                case 0: return 0;
                case 1: return 5;
                case 3: return 20;
                default: return 10;
            }
        }

        private void DoConversion(List<string> paths, string fileName, ImagePageMode pageMode,
            int marginMm, bool mergeToOne)
        {
            string dir = AppConfig.GetFeatureOutputDir("图片转PDF");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            var failures = new List<string>();

            if (mergeToOne)
            {
                // 所有图片合成为一个 PDF
                string outputFile = UniquePdfPath(dir, fileName);
                int pageCount;
                using (var converter = new ImageToPdfConverter
                {
                    PageMode = pageMode,
                    MarginMm = marginMm,
                    JpegQuality = 92
                })
                {
                    pageCount = converter.Convert(paths, outputFile,
                        (done, total) => Dispatcher.Invoke(() =>
                            TxtConversionStatus.Text = $"正在处理 {done}/{total} ..."),
                        failures);
                }

                Dispatcher.Invoke(() =>
                {
                    if (failures.Count > 0)
                    {
                        TxtConversionStatus.Text = $"完成，{pageCount} 页，{failures.Count} 张图片被跳过";
                        MessageBox.Show(
                            $"已生成 {pageCount} 页 PDF:\n{outputFile}\n\n以下图片被跳过:\n{FormatFailures(failures)}",
                            "部分图片跳过", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else
                    {
                        TxtConversionStatus.Text = $"完成，共 {pageCount} 页";
                    }
                    try { Process.Start(dir); } catch { }
                });
                return;
            }

            // 每张图各出一个 PDF
            int totalCount = paths.Count;
            int success = 0;
            for (int i = 0; i < totalCount; i++)
            {
                int index = i;
                string path = paths[i];
                Dispatcher.Invoke(() => TxtConversionStatus.Text = $"正在处理 {index + 1}/{totalCount} ...");

                string outputFile = UniquePdfPath(dir, Path.GetFileNameWithoutExtension(path));
                try
                {
                    using (var converter = new ImageToPdfConverter
                    {
                        PageMode = pageMode,
                        MarginMm = marginMm,
                        JpegQuality = 92
                    })
                    {
                        converter.Convert(new[] { path }, outputFile, null, null);
                        success++;
                    }
                }
                catch (Exception ex)
                {
                    failures.Add($"{Path.GetFileName(path)}: {ex.Message}");
                }
            }

            Dispatcher.Invoke(() =>
            {
                if (failures.Count > 0)
                {
                    TxtConversionStatus.Text = $"完成，{success} 个 PDF，{failures.Count} 张失败";
                    MessageBox.Show(
                        $"已生成 {success} 个 PDF 到:\n{dir}\n\n以下图片失败:\n{FormatFailures(failures)}",
                        "部分图片失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    TxtConversionStatus.Text = $"完成，共 {success} 个 PDF";
                }
                try { Process.Start(dir); } catch { }
            });
        }

        /// <summary>生成不重名的 PDF 路径（重名自动追加序号）。</summary>
        private static string UniquePdfPath(string dir, string baseName)
        {
            string path = Path.Combine(dir, baseName + ".pdf");
            int seq = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(dir, $"{baseName}_{seq++}.pdf");
            }
            return path;
        }

        /// <summary>失败信息摘要（最多列出 10 条）。</summary>
        private static string FormatFailures(List<string> failures)
        {
            return string.Join("\n", failures.Take(10))
                 + (failures.Count > 10 ? $"\n… 共 {failures.Count} 张" : "");
        }

        private void AddFiles(IEnumerable<string> paths)
        {
            foreach (var p in paths)
            {
                if (!SupportedExtensions.Contains(Path.GetExtension(p))) continue;
                if (_files.Any(f => f.FilePath.Equals(p, StringComparison.OrdinalIgnoreCase))) continue;

                _files.Add(new ImageFileItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p,
                    SizeText = ReadSizeText(p),
                    Thumbnail = LoadThumbnail(p)
                });
            }
            UpdateEmptyHint();
        }

        private static string ReadSizeText(string path)
        {
            try
            {
                using (var img = System.Drawing.Image.FromFile(path))
                    return $"{img.Width} × {img.Height}";
            }
            catch { return "无法读取"; }
        }

        private static BitmapSource LoadThumbnail(string path)
        {
            try
            {
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.UriSource = new Uri(path, UriKind.Absolute);
                bmp.DecodePixelWidth = 96;
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.CreateOptions = BitmapCreateOptions.IgnoreColorProfile;
                bmp.EndInit();
                bmp.Freeze();
                return bmp;
            }
            catch { return null; }
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
