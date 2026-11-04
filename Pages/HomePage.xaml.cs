using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace DocToPdfTool.Pages
{
    public partial class HomePage : UserControl
    {
        private readonly List<(string Tag, string Name, string Icon)> _functions = new List<(string, string, string)>
        {
            ("DocToPdf", "文档转PDF", "📄"),
            ("WordToPdf", "Word转PDF", "📝"),
            ("ExcelToPdf", "Excel转PDF", "📊"),
            ("ExcelToImage", "Excel转图片", "🖼️"),
            ("PptToPdf", "PPT转PDF", "📑"),
            ("PptToWord", "PPT转Word", "📄"),
            ("PptToImage", "PPT转图片", "🖼️"),
            ("WordToImage", "Word转图片", "🖼️"),
            ("PdfToWord", "PDF转Word", "📝"),
            ("PdfToExcel", "PDF转Excel", "📊"),
            ("PdfToPpt", "PDF转PPT", "📑"),
            ("PdfToImage", "PDF转图片", "🖼️"),
            ("ImageToPdf", "图片转PDF", "📄"),
            ("PdfEncrypt", "PDF加密", "🔒"),
            ("PdfWatermark", "PDF水印", "💧"),
            ("PdfPageManager", "PDF页面管理", "📑"),
            ("PdfImageExtract", "PDF提取图片", "🖼️"),
            ("PdfToScannedPdf", "PDF转扫描件", "📄"),
            ("PdfSplit", "拆分PDF", "✂️"),
            ("PdfMerge", "合并PDF", "🔗")
        };

        public HomePage()
        {
            InitializeComponent();
            LoadFunctionButtons(_functions);
        }

        private void LoadFunctionButtons(IEnumerable<(string Tag, string Name, string Icon)> functions)
        {
            FunctionPanel.Children.Clear();
            foreach (var (tag, name, icon) in functions)
            {
                var btn = new Button
                {
                    Tag = tag,
                    Content = $"{icon} {name}",
                    FontSize = 14,
                    Padding = new Thickness(16, 12, 16, 12),
                    Margin = new Thickness(8, 8, 8, 8),
                    Width = 160,
                    Height = 60
                };
                btn.Click += FunctionButton_Click;
                FunctionPanel.Children.Add(btn);
            }
        }

        private void FunctionButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string tag)
            {
                NavigateToFunction(tag);
            }
        }

        private void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            SearchFunctions();
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtSearch.Text))
            {
                LoadFunctionButtons(_functions);
            }
            else
            {
                var filtered = _functions.Where(f => f.Name.IndexOf(TxtSearch.Text, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
                LoadFunctionButtons(filtered);
            }
        }

        private void SearchFunctions()
        {
            if (string.IsNullOrWhiteSpace(TxtSearch.Text))
            {
                LoadFunctionButtons(_functions);
            }
            else
            {
                var filtered = _functions.Where(f => f.Name.IndexOf(TxtSearch.Text, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
                LoadFunctionButtons(filtered);
            }
        }

        private void NavigateToFunction(string functionTag)
        {
            if (Window.GetWindow(this) is MainWindow mainWindow)
            {
                mainWindow.NavigateToFunction(functionTag);
            }
        }
    }
}
