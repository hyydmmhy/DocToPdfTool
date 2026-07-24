using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace DocToPdfTool
{
    public partial class MainWindow : Window
    {
        private readonly Dictionary<string, UserControl> _pages = new Dictionary<string, UserControl>();

        public MainWindow()
        {
            InitializeComponent();
            NavList.SelectedIndex = 0;
        }

        private void NavList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
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
                        "PdfToWord" => new Pages.PdfToWordPage(),
                        "PdfToImage" => new Pages.PdfToImagePage(),
                        "PdfToExcel" => new Pages.PdfToExcelPage(),
                        "PdfToPpt" => new Pages.PdfToPptPage(),
                        "PdfToScannedPdf" => new Pages.PdfToScannedPdfPage(),
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
    }
}