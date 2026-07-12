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
    public partial class PdfToImagePage : UserControl
    {
        private readonly ObservableCollection<PdfFileItem> _files = new ObservableCollection<PdfFileItem>();

        public PdfToImagePage()
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
            double pathW = Math.Max(100, total - nameW - 60);

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

            // 确定输出目录
            string defaultOutputDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "PDF转图片输出");
            if (!Directory.Exists(defaultOutputDir))
                Directory.CreateDirectory(defaultOutputDir);

            // 读取设置
            double scale = 1.5;
            if (CmbScale.SelectedItem is ComboBoxItem scaleItem &&
                double.TryParse(scaleItem.Content.ToString(), out double parsedScale))
            {
                scale = parsedScale;
            }

            int dpi = 300;
            if (CmbDpi.SelectedItem is ComboBoxItem dpiItem &&
                int.TryParse(dpiItem.Content.ToString(), out int parsedDpi))
            {
                dpi = parsedDpi;
            }

            bool usePng = CmbFormat.SelectedItem is ComboBoxItem fmtItem &&
                          fmtItem.Content.ToString().Equals("PNG", StringComparison.OrdinalIgnoreCase);

            BtnConvert.IsEnabled = false;
            BtnAddFiles.IsEnabled = false;
            TxtConversionStatus.Text = "准备中...";

            var snapshot = _files.ToList();

            try
            {
                bool hasError = false;
                string errorMsg = null;

                await Task.Run(() =>
                {
                    int total = snapshot.Count;
                    int completed = 0;

                    using (var converter = new PdfToImageConverter())
                    {
                        converter.Dpi = dpi;
                        converter.Scale = scale;
                        converter.OutputFormat = usePng
                            ? System.Drawing.Imaging.ImageFormat.Png
                            : System.Drawing.Imaging.ImageFormat.Jpeg;

                        foreach (var file in snapshot)
                        {
                            completed++;

                            Dispatcher.Invoke(() =>
                            {
                                TxtConversionStatus.Text = $"正在转换 ({completed}/{total}) {file.FileName}";
                            });

                            try
                            {
                                converter.Convert(file.FilePath, defaultOutputDir);
                            }
                            catch (Exception ex)
                            {
                                hasError = true;
                                errorMsg = $"{file.FileName}: {ex.Message}";
                            }
                            finally
                            {
                                // 每个文件转换完立即回收，避免高分辨率位图在大对象堆上持续堆积
                                ReleaseMemory();
                            }
                        }
                    }
                });

                if (hasError)
                {
                    TxtConversionStatus.Text = "部分文件转换失败";
                    MessageBox.Show($"以下文件转换失败:\n{errorMsg}",
                        "部分文件失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    TxtConversionStatus.Text = "全部转换完成";
                    try { System.Diagnostics.Process.Start(defaultOutputDir); } catch { }
                }
            }
            catch (Exception ex)
            {
                TxtConversionStatus.Text = $"转换失败: {ex.Message}";
                MessageBox.Show(ex.Message, "转换错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                // 强制回收内存，释放 WinRT 和 GDI+ 资源
                ReleaseMemory();

                BtnConvert.IsEnabled = true;
                BtnAddFiles.IsEnabled = true;
            }
        }

        /// <summary>
        /// 强制GC回收 + 压缩大对象堆 + 将释放出来的内存真正归还给操作系统。
        /// 高DPI渲染的位图会进入大对象堆(LOH)，普通GC.Collect()不会压缩LOH、
        /// 也不会把已释放的页面还给OS，所以任务管理器里数字不会掉——这里把这两步都补上。
        /// </summary>
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