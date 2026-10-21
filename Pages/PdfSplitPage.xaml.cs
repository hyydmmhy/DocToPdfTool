using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using DocToPdfTool.Utils;
using Microsoft.Win32;

namespace DocToPdfTool.Pages
{
    public partial class PdfSplitPage : UserControl
    {
        private readonly ObservableCollection<PdfFileItem> _files = new ObservableCollection<PdfFileItem>();

        public PdfSplitPage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;
            UpdateEmptyHint();
        }

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "选择要拆分的PDF文件",
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
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 4)
                return;

            double total = FileListView.ActualWidth - 14;
            double nameW = Math.Max(120, total * 0.28);
            double pageW = 50;
            double pathW = Math.Max(80, total - nameW - pageW - 70);

            if (Math.Abs(gv.Columns[0].Width - nameW) > 1)
                gv.Columns[0].Width = nameW;
            if (Math.Abs(gv.Columns[1].Width - pathW) > 1)
                gv.Columns[1].Width = pathW;
        }

        private void SplitMode_Changed(object sender, RoutedEventArgs e)
        {
            if (PanelByPages == null) return; // 初始化期间
            PanelByPages.Visibility = RadioByPages.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
            PanelByRange.Visibility = RadioByRange.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        }

        private async void BtnSplit_Click(object sender, RoutedEventArgs e)
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

            bool byRange = RadioByRange.IsChecked == true;
            int pagesPerFile = 1;
            List<int[]> ranges = null;

            if (byRange)
            {
                try
                {
                    ranges = PdfSplitter.ParseRanges(TxtPageRanges.Text);
                }
                catch (FormatException ex)
                {
                    MessageBox.Show(ex.Message, "页码范围格式错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }
            else
            {
                if (!int.TryParse(TxtPagesPerFile.Text?.Trim(), out pagesPerFile) || pagesPerFile < 1)
                {
                    MessageBox.Show("请输入有效的拆分页数（大于等于1的整数）", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
            }

            BtnSplit.IsEnabled = false;
            BtnAddFiles.IsEnabled = false;
            TxtConversionStatus.Text = "准备中...";

            var snapshot = _files.ToList();

            try
            {
                await Task.Run(() => DoSplit(snapshot, pagesPerFile, ranges, byRange));
            }
            catch (Exception ex)
            {
                TxtConversionStatus.Text = $"拆分失败: {ex.Message}";
                MessageBox.Show(ex.Message, "拆分错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnSplit.IsEnabled = true;
                BtnAddFiles.IsEnabled = true;
                ReleaseMemory();
            }
        }

        private static void ReleaseMemory()
        {
            try
            {
                GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                EmptyWorkingSet(Process.GetCurrentProcess().Handle);
            }
            catch { }
        }

        [DllImport("psapi.dll")]
        private static extern bool EmptyWorkingSet(IntPtr hProcess);

        private void DoSplit(List<PdfFileItem> files, int pagesPerFile, List<int[]> ranges, bool byRange)
        {
            string outputDir = AppConfig.GetFeatureOutputDir("拆分PDF");

            var failedFiles = new List<string>();
            int totalOutputFiles = 0;

            using (var splitter = new PdfSplitter())
            {
                int total = files.Count;
                int completed = 0;

                foreach (var file in files)
                {
                    completed++;

                    Dispatcher.Invoke(() =>
                    {
                        TxtConversionStatus.Text = $"正在拆分 ({completed}/{total}) {file.FileName}";
                    });

                    try
                    {
                        List<string> outputs;
                        if (byRange)
                            outputs = splitter.SplitByRanges(file.FilePath, outputDir, ranges);
                        else
                            outputs = splitter.Split(file.FilePath, outputDir, pagesPerFile);
                        totalOutputFiles += outputs.Count;
                    }
                    catch (Exception ex)
                    {
                        failedFiles.Add($"{file.FileName}: {ex.Message}");
                    }

                    ReleaseMemory();
                }
            }

            Dispatcher.Invoke(() =>
            {
                if (failedFiles.Count > 0)
                {
                    TxtConversionStatus.Text = $"完成，{failedFiles.Count} 个文件拆分失败";
                    MessageBox.Show($"以下文件拆分失败:\n{string.Join("\n", failedFiles)}",
                        "部分文件失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    TxtConversionStatus.Text = $"全部拆分完成，共生成 {totalOutputFiles} 个文件";
                    try { Process.Start(outputDir); } catch { }
                }
            });
        }

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
                        int pageCount = 0;
                        try
                        {
                            using (var doc = UglyToad.PdfPig.PdfDocument.Open(path))
                            {
                                pageCount = doc.NumberOfPages;
                            }
                        }
                        catch { }

                        _files.Add(new PdfFileItem { FileName = name, FilePath = path, PageCount = pageCount });
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