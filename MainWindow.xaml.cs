using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using DocToPdfTool.Utils;

namespace DocToPdfTool
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnBrowseSource_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "选择要转换的文档",
                Filter = "支持的文档格式|*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.txt;*.rtf;*.html;*.htm;*.mhtml|" +
                         "Word 文档 (*.doc;*.docx)|*.doc;*.docx|" +
                         "Excel 表格 (*.xls;*.xlsx)|*.xls;*.xlsx|" +
                         "PPT 演示 (*.ppt;*.pptx)|*.ppt;*.pptx|" +
                         "文本文件 (*.txt;*.rtf)|*.txt;*.rtf|" +
                         "HTML 文件 (*.html;*.htm;*.mhtml)|*.html;*.htm;*.mhtml|" +
                         "所有文件 (*.*)|*.*"
            };

            if (dialog.ShowDialog() == true)
            {
                TxtSourceFile.Text = dialog.FileName;
                TxtOutputDir.Text = Path.GetDirectoryName(dialog.FileName);
            }
        }

        private void BtnBrowseOutput_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "选择PDF输出目录",
                SelectedPath = TxtOutputDir.Text
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                TxtOutputDir.Text = dialog.SelectedPath;
            }
        }

        private async void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            var sourceFile = TxtSourceFile.Text?.Trim();
            var outputDir = TxtOutputDir.Text?.Trim();

            if (string.IsNullOrEmpty(sourceFile))
            {
                MessageBox.Show("请先选择要转换的源文件", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!File.Exists(sourceFile))
            {
                MessageBox.Show("源文件不存在，请重新选择", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Path.GetDirectoryName(sourceFile);
                TxtOutputDir.Text = outputDir;
            }

            if (!Directory.Exists(outputDir))
            {
                try
                {
                    Directory.CreateDirectory(outputDir);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"创建输出目录失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }

            var extension = Path.GetExtension(sourceFile).ToLowerInvariant();
            var useOffice = RadioOffice.IsChecked == true;
            var isHtml = extension == ".html" || extension == ".htm" || extension == ".mhtml";

            if (!useOffice)
            {
                if (!PdfConverter.SupportedExtensions.Contains(extension))
                {
                    MessageBox.Show($"不支持的文件格式: {extension}\n\n支持的格式: {string.Join(", ", PdfConverter.SupportedExtensions)}",
                        "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }
            else if (!isHtml)
            {
                if (!OfficePdfConverter.SupportedExtensions.Contains(extension))
                {
                    MessageBox.Show($"不支持的文件格式: {extension}\n\nOffice 引擎支持的格式: {string.Join(", ", OfficePdfConverter.SupportedExtensions)}",
                        "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }

            var outputFile = Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(sourceFile) + ".pdf");

            BtnConvert.IsEnabled = false;
            TxtStatus.Text = "正在转换...";
            TxtStatus.Foreground = System.Windows.Media.Brushes.Gray;

            try
            {
                await Task.Run(() =>
                {
                    if (useOffice && !isHtml)
                    {
                        using (var converter = new OfficePdfConverter())
                        {
                            converter.Convert(sourceFile, outputFile);
                        }
                    }
                    else
                    {
                        using (var converter = new PdfConverter())
                        {
                            converter.Convert(sourceFile, outputFile);
                        }
                    }
                });
                TxtStatus.Text = "转换成功";
                TxtStatus.Foreground = System.Windows.Media.Brushes.Gray;
                System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{outputFile}\"");
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"转换失败: {ex.Message}";
                TxtStatus.Foreground = System.Windows.Media.Brushes.Gray;
            }
            finally
            {
                BtnConvert.IsEnabled = true;
            }
        }

        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void BtnTopmost_Click(object sender, RoutedEventArgs e)
        {
            Topmost = BtnTopmost.IsChecked == true;
        }
    }
}