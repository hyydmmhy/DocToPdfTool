using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;

namespace DocToPdfTool
{
    public partial class MainWindow : Window
    {
        private readonly Dictionary<string, UserControl> _pages = new Dictionary<string, UserControl>();
        private bool _isNavigatingFromHome;

        public MainWindow()
        {
            InitializeComponent();
            // 启动默认只显示首页（无左侧导航）
            ShowHomePage();
        }

        /// <summary>显示首页（隐藏左侧导航，全窗口内容区；右上角按钮常驻）。</summary>
        public void ShowHomePage()
        {
            SideNav.Visibility = Visibility.Collapsed;
            TopBar.Visibility = Visibility.Visible;
            NavList.SelectedItem = null;

            if (!_pages.TryGetValue("Home", out var page))
            {
                page = new Pages.HomePage();
                _pages["Home"] = page;
            }
            PageContent.Content = page;
        }

        /// <summary>从首页跳转到指定功能页：显示导航栏+顶栏，自动定位左侧菜单。</summary>
        public void NavigateToFunction(string functionTag)
        {
            _isNavigatingFromHome = true;
            try
            {
                SideNav.Visibility = Visibility.Visible;
                TopBar.Visibility = Visibility.Visible;

                // 自动选中左侧导航对应项，并滚动到可见位置
                foreach (var obj in NavList.Items)
                {
                    if (obj is ListBoxItem item && item.Tag?.ToString() == functionTag)
                    {
                        NavList.SelectedItem = item;
                        // 等布局完成后滚动，确保选中项在可视区内
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            try
                            {
                                NavList.UpdateLayout();
                                NavList.ScrollIntoView(item);
                            }
                            catch { }
                        }), System.Windows.Threading.DispatcherPriority.Loaded);
                        break;
                    }
                }

                // 若左侧无对应项（理论上都有），仍直接切页
                SwitchToPage(functionTag);
            }
            finally
            {
                _isNavigatingFromHome = false;
            }
        }

        private void BtnBackHome_Click(object sender, RoutedEventArgs e)
        {
            ShowHomePage();
        }

        private void NavList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 从首页跳转时已手动处理，避免重复触发
            if (_isNavigatingFromHome) return;

            if (NavList.SelectedItem is ListBoxItem item && item.Tag is string pageKey)
            {
                SwitchToPage(pageKey);
            }
        }

        private void SwitchToPage(string pageKey)
        {
            if (!_pages.TryGetValue(pageKey, out var page))
            {
                try
                {
                    page = pageKey switch
                    {
                        "DocToPdf" => new Pages.DocToPdfPage(),
                        "WordToPdf" => new Pages.WordToPdfPage(),
                        "ExcelToPdf" => new Pages.ExcelToPdfPage(),
                        "ExcelToImage" => new Pages.ExcelToImagePage(),
                        "PptToPdf" => new Pages.PptToPdfPage(),
                        "PptToWord" => new Pages.PptToWordPage(),
                        "PptToImage" => new Pages.PptToImagePage(),
                        "WordToImage" => new Pages.WordToImagePage(),
                        "PdfToWord" => new Pages.PdfToWordPage(),
                        "PdfToImage" => new Pages.PdfToImagePage(),
                        "ImageToPdf" => new Pages.ImageToPdfPage(),
                        "PdfEncrypt" => new Pages.PdfEncryptPage(),
                        "PdfWatermark" => new Pages.PdfWatermarkPage(),
                        "PdfPageManager" => new Pages.PdfPageManagerPage(),
                        "PdfImageExtract" => new Pages.PdfImageExtractPage(),
                        "PdfCompress" => new Pages.PdfCompressPage(),
                        "PdfToHtml" => new Pages.PdfToFormatPage(Utils.PdfTargetFormat.Html),
                        "PdfToSvg" => new Pages.PdfToFormatPage(Utils.PdfTargetFormat.Svg),
                        "PdfToMarkdown" => new Pages.PdfToFormatPage(Utils.PdfTargetFormat.Markdown),
                        "PdfToGray" => new Pages.PdfToFormatPage(Utils.PdfTargetFormat.Gray),
                        "PdfToPdfA" => new Pages.PdfToFormatPage(Utils.PdfTargetFormat.PdfA),
                        "PdfToXps" => new Pages.PdfToFormatPage(Utils.PdfTargetFormat.Xps),
                        "WordToMarkdown" => new Pages.OfficeToFormatPage(Utils.OfficeSourceKind.Word, Utils.OfficeTargetFormat.Markdown),
                        "WordToEpub" => new Pages.OfficeToFormatPage(Utils.OfficeSourceKind.Word, Utils.OfficeTargetFormat.Epub),
                        "ExcelToCsv" => new Pages.OfficeToFormatPage(Utils.OfficeSourceKind.Excel, Utils.OfficeTargetFormat.Csv),
                        "PptToHtml" => new Pages.OfficeToFormatPage(Utils.OfficeSourceKind.Ppt, Utils.OfficeTargetFormat.Html),
                        "PptToTiff" => new Pages.OfficeToFormatPage(Utils.OfficeSourceKind.Ppt, Utils.OfficeTargetFormat.Tiff),
                        "ExcelToHtml" => new Pages.OfficeToFormatPage(Utils.OfficeSourceKind.Excel, Utils.OfficeTargetFormat.ExcelHtml),
                        "ExcelToJson" => new Pages.OfficeToFormatPage(Utils.OfficeSourceKind.Excel, Utils.OfficeTargetFormat.ExcelJson),
                        "ExcelToMd" => new Pages.OfficeToFormatPage(Utils.OfficeSourceKind.Excel, Utils.OfficeTargetFormat.ExcelMarkdown),
                        "PdfPageNumber" => new Pages.PdfPageNumberPage(),
                        "ExcelSplitMerge" => new Pages.ExcelSplitMergePage(),
                        "ImageBatch" => new Pages.ImageBatchPage(),
                        "IdPhotoLayout" => new Pages.IdPhotoLayoutPage(),
                        "TextExtract" => new Pages.TextExtractPage(),
                        "Translate" => new Pages.TranslatePage(),
                        "Ocr" => new Pages.OcrPage(),
                        "QrCode" => new Pages.QrCodePage(),
                        "PdfToExcel" => new Pages.PdfToExcelPage(),
                        "PdfToPpt" => new Pages.PdfToPptPage(),
                        "PdfToScannedPdf" => new Pages.PdfToScannedPdfPage(),
                        "PdfToLongImage" => new Pages.PdfToLongImagePage(),
                        "PdfSplit" => new Pages.PdfSplitPage(),
                        "PdfMerge" => new Pages.PdfMergePage(),
                        "ImageStitch" => new Pages.ImageStitchPage(),
                        _ => throw new ArgumentException($"Unknown page: {pageKey}")
                    };
                    _pages[pageKey] = page;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"页面加载失败: {ex.Message}\n\n{ex.StackTrace}",
                        "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }
            PageContent.Content = page;
        }

        private void BtnTopmost_Click(object sender, RoutedEventArgs e)
        {
            Topmost = BtnTopmost.IsChecked == true;
        }

        private void BtnSettings_Click(object sender, RoutedEventArgs e)
        {
            var settingsWindow = new Pages.SettingsWindow { Owner = this };
            settingsWindow.ShowDialog();
        }

        /// <summary>打开 GitHub 仓库（默认浏览器）。</summary>
        private void BtnGitHub_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo("https://github.com/hyydmmhy/DocToPdfTool")
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("无法打开 GitHub 仓库: " + ex.Message,
                    "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}