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
    public partial class PdfToPptPage : UserControl
    {
        private readonly ObservableCollection<PdfFileItem> _files = new ObservableCollection<PdfFileItem>();

        public PdfToPptPage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;
            UpdateEmptyHint();
        }

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "选择要转换的PDF文件",
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
            double pathW = Math.Max(100, total - nameW - 70);

            if (Math.Abs(gv.Columns[0].Width - nameW) > 1)
                gv.Columns[0].Width = nameW;
            if (Math.Abs(gv.Columns[1].Width - pathW) > 1)
                gv.Columns[1].Width = pathW;
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

            BtnConvert.IsEnabled = false;
            BtnAddFiles.IsEnabled = false;
            TxtConversionStatus.Text = "准备中...";

            var snapshot = _files.ToList();

            try
            {
                await Task.Run(() => DoConversion(snapshot));
            }
            catch (Exception ex)
            {
                TxtConversionStatus.Text = $"转换失败: {ex.Message}";
                MessageBox.Show(ex.Message, "转换错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnConvert.IsEnabled = true;
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

        private void DoConversion(List<PdfFileItem> files)
        {
            string defaultOutputDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "PDF转PPT输出");

            var failedFiles = new List<string>();

            using (var converter = new PdfToPptConverter())
            {
                int total = files.Count;
                int completed = 0;

                foreach (var file in files)
                {
                    completed++;

                    Dispatcher.Invoke(() =>
                    {
                        TxtConversionStatus.Text = $"正在转换 ({completed}/{total}) {file.FileName}";
                    });

                    string outputDir = defaultOutputDir;
                    if (!Directory.Exists(outputDir))
                        Directory.CreateDirectory(outputDir);

                    try
                    {
                        converter.Convert(file.FilePath, outputDir, () => ReleaseMemory());
                    }
                    catch (Exception ex)
                    {
                        failedFiles.Add($"{file.FileName}: {ex.Message}");
                    }

                    // 再次回收（确保转换完成后所有对象已释放）
                    ReleaseMemory();
                }
            }

            // 最终回收
            ReleaseMemory();

            Dispatcher.Invoke(() =>
            {
                if (failedFiles.Count > 0)
                {
                    TxtConversionStatus.Text = $"完成，{failedFiles.Count} 个文件转换失败";
                    MessageBox.Show($"以下文件转换失败:\n{string.Join("\n", failedFiles)}",
                        "部分文件失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    TxtConversionStatus.Text = "全部转换完成";
                    try { System.Diagnostics.Process.Start(defaultOutputDir); } catch { }
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