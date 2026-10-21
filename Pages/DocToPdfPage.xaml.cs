using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using DocToPdfTool.Utils;
using Microsoft.Win32;

namespace DocToPdfTool.Pages
{
    public partial class DocToPdfPage : UserControl
    {
        private readonly ObservableCollection<PdfFileItem> _files = new ObservableCollection<PdfFileItem>();

        public DocToPdfPage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;
            UpdateEmptyHint();
        }

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "选择要转换的文档",
                Filter = "支持的文档格式|*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.txt;*.rtf;*.html;*.htm;*.mhtml|" +
                         "所有文件 (*.*)|*.*",
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
                MessageBox.Show("请先添加文档", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var invalid = _files.FirstOrDefault(f => !File.Exists(f.FilePath));
            if (invalid != null)
            {
                MessageBox.Show($"文件不存在: {invalid.FilePath}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var useOffice = RadioOffice.IsChecked == true;

            // 校验当前引擎支持的文件格式
            var unsupported = _files.FirstOrDefault(f =>
            {
                var ext = Path.GetExtension(f.FilePath).ToLowerInvariant();
                if (useOffice)
                    return !IsHtml(ext) && !OfficePdfConverter.SupportedExtensions.Contains(ext);
                return !PdfConverter.SupportedExtensions.Contains(ext);
            });
            if (unsupported != null)
            {
                MessageBox.Show($"当前引擎不支持的文件格式: {unsupported.FileName}\n\n" +
                    (useOffice
                        ? $"Office 引擎支持的格式: {string.Join(", ", OfficePdfConverter.SupportedExtensions)}"
                        : $"WPS 引擎支持的格式: {string.Join(", ", PdfConverter.SupportedExtensions)}"),
                    "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            BtnConvert.IsEnabled = false;
            BtnAddFiles.IsEnabled = false;
            TxtStatus.Text = "准备中...";

            var snapshot = _files.ToList();

            try
            {
                await Task.Run(() => DoConversion(snapshot, useOffice));
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"转换失败: {ex.Message}";
                MessageBox.Show(ex.Message, "转换错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnConvert.IsEnabled = true;
                BtnAddFiles.IsEnabled = true;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void DoConversion(List<PdfFileItem> files, bool useOffice)
        {
            string outputDir = AppConfig.GetFeatureOutputDir("文档转PDF");

            int total = files.Count;
            int completed = 0;
            var failedFiles = new List<string>();

            foreach (var file in files)
            {
                completed++;

                Dispatcher.Invoke(() =>
                {
                    TxtStatus.Text = $"正在转换 ({completed}/{total}) {file.FileName}";
                });

                var outputFile = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(file.FilePath) + ".pdf");

                try
                {
                    var ext = Path.GetExtension(file.FilePath).ToLowerInvariant();

                    if (useOffice && !IsHtml(ext))
                    {
                        using (var converter = new OfficePdfConverter())
                        {
                            converter.Convert(file.FilePath, outputFile);
                        }
                    }
                    else
                    {
                        using (var converter = new PdfConverter())
                        {
                            converter.Convert(file.FilePath, outputFile);
                        }
                    }
                }
                catch (Exception ex)
                {
                    failedFiles.Add($"{file.FileName}: {ex.Message}");
                }
            }

            Dispatcher.Invoke(() =>
            {
                if (failedFiles.Count > 0)
                {
                    TxtStatus.Text = $"完成，{failedFiles.Count} 个文件转换失败";
                    MessageBox.Show($"以下文件转换失败:\n{string.Join("\n", failedFiles)}",
                        "部分文件失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    TxtStatus.Text = "全部转换完成";
                    try { System.Diagnostics.Process.Start(outputDir); } catch { }
                }
            });
        }

        private static bool IsHtml(string ext)
        {
            return ext == ".html" || ext == ".htm" || ext == ".mhtml";
        }

        private void AddFiles(IEnumerable<string> filePaths)
        {
            foreach (var path in filePaths)
            {
                var ext = Path.GetExtension(path).ToLowerInvariant();
                if (IsHtml(ext)
                    || PdfConverter.SupportedExtensions.Contains(ext)
                    || OfficePdfConverter.SupportedExtensions.Contains(ext))
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