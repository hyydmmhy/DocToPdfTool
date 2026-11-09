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
using DocToPdfTool.Utils;
using Microsoft.Win32;

namespace DocToPdfTool.Pages
{
    public partial class PdfToLongImagePage : UserControl
    {
        private readonly ObservableCollection<PdfFileItem> _files = new ObservableCollection<PdfFileItem>();

        public PdfToLongImagePage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;
            UpdateEmptyHint();
            UpdateFormatControls();
        }

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "选择要生成长图的PDF文件",
                Filter = "PDF 文件 (*.pdf)|*.pdf|所有文件 (*.*)|*.*",
                Multiselect = true
            };

            if (dialog.ShowDialog() == true)
            {
                AddFiles(dialog.FileNames);
            }
        }

        private void FileListView_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                AddFiles(files);
            }
        }

        private void FileListView_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
                ? DragDropEffects.Copy
                : DragDropEffects.None;
            e.Handled = true;
        }

        private void BtnRemoveFile_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement element && element.Tag is PdfFileItem item)
            {
                _files.Remove(item);
                UpdateEmptyHint();
            }
        }

        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 3)
                return;

            double total = FileListView.ActualWidth - 14;
            double nameW = Math.Max(150, total * 0.30);
            double pathW = Math.Max(100, total - nameW - 60);

            if (Math.Abs(gv.Columns[0].Width - nameW) > 1)
                gv.Columns[0].Width = nameW;
            if (Math.Abs(gv.Columns[1].Width - pathW) > 1)
                gv.Columns[1].Width = pathW;
        }

        private void Format_Changed(object sender, SelectionChangedEventArgs e)
        {
            UpdateFormatControls();
        }

        private void UpdateFormatControls()
        {
            if (TxtQuality == null || LblQuality == null) return;
            bool png = CmbFormat.SelectedIndex == 1;
            TxtQuality.IsEnabled = !png;
            TxtQuality.Opacity = png ? 0.45 : 1.0;
            LblQuality.Opacity = png ? 0.45 : 1.0;
        }

        private async void BtnConvert_Click(object sender, RoutedEventArgs e)
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

            string defaultOutputDir = AppConfig.GetFeatureOutputDir("PDF转长图");

            int dpi = 150;
            if (CmbDpi.SelectedItem is ComboBoxItem dpiItem &&
                int.TryParse(dpiItem.Content.ToString(), out int parsedDpi))
            {
                dpi = parsedDpi;
            }

            bool usePng = CmbFormat.SelectedItem is ComboBoxItem fmtItem &&
                          fmtItem.Content.ToString().Equals("PNG", StringComparison.OrdinalIgnoreCase);

            int quality = 92;
            if (int.TryParse((TxtQuality.Text ?? "").Trim(), out int q) && q >= 1 && q <= 100) quality = q;

            BtnConvert.IsEnabled = false;
            BtnAddFiles.IsEnabled = false;
            TxtConversionStatus.Text = "准备中...";

            var snapshot = _files.ToList();

            try
            {
                bool hasError = false;
                string errorMsg = null;
                int success = 0;
                string lastDir = defaultOutputDir;

                await Task.Run(() =>
                {
                    int total = snapshot.Count;
                    int completed = 0;

                    var options = new PdfToLongImageOptions
                    {
                        Dpi = dpi,
                        OutputPng = usePng,
                        JpegQuality = quality
                    };

                    foreach (var file in snapshot)
                    {
                        completed++;

                        Dispatcher.Invoke(() =>
                        {
                            TxtConversionStatus.Text = $"正在生成 ({completed}/{total}) {file.FileName}";
                        });

                        try
                        {
                            using (var converter = new PdfToLongImageConverter())
                            {
                                converter.Convert(file.FilePath, defaultOutputDir, options);
                                success++;
                            }
                        }
                        catch (Exception ex)
                        {
                            hasError = true;
                            errorMsg = $"{file.FileName}: {ex.Message}";
                        }
                    }
                });

                if (hasError)
                {
                    TxtConversionStatus.Text = "部分文件生成失败";
                    MessageBox.Show($"以下文件生成失败:\n{errorMsg}",
                        "部分文件失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    TxtConversionStatus.Text = "全部生成完成";
                    try { Process.Start(lastDir); } catch { }
                }
            }
            catch (Exception ex)
            {
                TxtConversionStatus.Text = $"生成失败: {ex.Message}";
                MessageBox.Show(ex.Message, "生成错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ReleaseMemory();

                BtnConvert.IsEnabled = true;
                BtnAddFiles.IsEnabled = true;
            }
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

        private void AddFiles(IEnumerable<string> filePaths)
        {
            foreach (var path in filePaths)
            {
                if (string.Equals(Path.GetExtension(path), ".pdf",
                    StringComparison.OrdinalIgnoreCase))
                {
                    var name = Path.GetFileNameWithoutExtension(path);
                    if (!_files.Any(f => f.FilePath.Equals(path, StringComparison.OrdinalIgnoreCase)))
                    {
                        _files.Add(new PdfFileItem { FileName = name, FilePath = path });
                    }
                }
            }
            UpdateEmptyHint();
        }

        private void UpdateEmptyHint()
        {
            EmptyOverlay.Visibility = _files.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }
    }
}
