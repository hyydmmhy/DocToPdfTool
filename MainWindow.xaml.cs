using System;
using System.Collections.Generic;
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

        /// <summary>显示首页（隐藏左侧导航和顶栏，全窗口内容区）。</summary>
        public void ShowHomePage()
        {
            SideNav.Visibility = Visibility.Collapsed;
            TopBar.Visibility = Visibility.Collapsed;
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
                        "PdfToExcel" => new Pages.PdfToExcelPage(),
                        "PdfToPpt" => new Pages.PdfToPptPage(),
                        "PdfToScannedPdf" => new Pages.PdfToScannedPdfPage(),
                        "PdfSplit" => new Pages.PdfSplitPage(),
                        "PdfMerge" => new Pages.PdfMergePage(),
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
    }
}