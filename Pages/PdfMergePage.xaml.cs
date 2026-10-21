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
using UglyToad.PdfPig;

namespace DocToPdfTool.Pages
{
    public partial class PdfMergePage : UserControl
    {
        private readonly ObservableCollection<PdfFileItem> _files = new ObservableCollection<PdfFileItem>();

        public PdfMergePage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;
            UpdateEmptyHint();
        }

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "选择要合并的PDF文件",
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
                RefreshSeq();
                UpdateEmptyHint();
            }
        }

        private void BtnMoveUp_Click(object sender, RoutedEventArgs e)
        {
            var item = FileListView.SelectedItem as PdfFileItem;
            if (item == null)
            {
                MessageBox.Show("请先在列表中选择要上移的文件", "提示",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int index = _files.IndexOf(item);
            if (index > 0)
            {
                _files.Move(index, index - 1);
                RefreshSeq();
            }
        }

        private void BtnMoveDown_Click(object sender, RoutedEventArgs e)
        {
            var item = FileListView.SelectedItem as PdfFileItem;
            if (item == null)
            {
                MessageBox.Show("请先在列表中选择要下移的文件", "提示",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int index = _files.IndexOf(item);
            if (index >= 0 && index < _files.Count - 1)
            {
                _files.Move(index, index + 1);
                RefreshSeq();
            }
        }

        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 5)
                return;

            double total = FileListView.ActualWidth - 14;
            double nameW = Math.Max(120, total * 0.30);
            double pathW = Math.Max(80, total - 45 - nameW - 50 - 70);

            if (Math.Abs(gv.Columns[1].Width - nameW) > 1)
                gv.Columns[1].Width = nameW;
            if (Math.Abs(gv.Columns[2].Width - pathW) > 1)
                gv.Columns[2].Width = pathW;
        }

        private async void BtnMerge_Click(object sender, RoutedEventArgs e)
        {
            if (_files.Count < 2)
            {
                MessageBox.Show("请至少添加两个PDF文件", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var invalid = _files.FirstOrDefault(f => !File.Exists(f.FilePath));
            if (invalid != null)
            {
                MessageBox.Show($"文件不存在: {invalid.FilePath}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            BtnMerge.IsEnabled = false;
            BtnAddFiles.IsEnabled = false;
            TxtConversionStatus.Text = "准备中...";

            var snapshot = _files.ToList();

            try
            {
                await Task.Run(() => DoMerge(snapshot));
            }
            catch (Exception ex)
            {
                TxtConversionStatus.Text = $"合并失败: {ex.Message}";
                MessageBox.Show(ex.Message, "合并错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnMerge.IsEnabled = true;
                BtnAddFiles.IsEnabled = true;
                ReleaseMemory();
            }
        }

        private void DoMerge(List<PdfFileItem> files)
        {
            string outputDir = AppConfig.GetFeatureOutputDir("合并PDF");
            string baseName = Path.GetFileNameWithoutExtension(files[0].FilePath);

            // 已存在同名文件时自动追加序号，避免覆盖
            string outputPath = Path.Combine(outputDir, baseName + "_合并.pdf");
            int counter = 2;
            while (File.Exists(outputPath))
            {
                outputPath = Path.Combine(outputDir, $"{baseName}_合并({counter}).pdf");
                counter++;
            }

            var paths = files.Select(f => f.FilePath).ToList();

            Dispatcher.Invoke(() =>
            {
                TxtConversionStatus.Text = $"正在合并 {paths.Count} 个文件...";
            });

            using (var merger = new PdfMerger())
            {
                merger.Merge(paths, outputPath);
            }

            Dispatcher.Invoke(() =>
            {
                TxtConversionStatus.Text = "合并完成";
                try { Process.Start(outputDir); } catch { }
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
                            using (var doc = PdfDocument.Open(path))
                            {
                                pageCount = doc.NumberOfPages;
                            }
                        }
                        catch { }

                        _files.Add(new PdfFileItem { FileName = name, FilePath = path, PageCount = pageCount });
                    }
                }
            }
            RefreshSeq();
            UpdateEmptyHint();
        }

        private void RefreshSeq()
        {
            for (int i = 0; i < _files.Count; i++)
            {
                _files[i].Seq = i + 1;
            }
            FileListView.Items.Refresh();
        }

        private void UpdateEmptyHint()
        {
            EmptyOverlay.Visibility = _files.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
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
    }
}