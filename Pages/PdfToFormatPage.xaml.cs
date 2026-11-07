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
    /// <summary>
    /// PDF 转单一格式的通用页面。
    /// 一个实现，六个入口（HTML / SVG / Markdown / 灰度 / PDF-A / XPS），
    /// 每个格式在导航与首页各自独立，便于搜索直达。
    /// </summary>
    public partial class PdfToFormatPage : UserControl
    {
        private readonly ObservableCollection<PdfToFormatItem> _files =
            new ObservableCollection<PdfToFormatItem>();

        private readonly PdfTargetFormat _format;

        /// <summary>供 XAML 设计器使用的默认构造（等同于 HTML 页）。</summary>
        public PdfToFormatPage() : this(PdfTargetFormat.Html) { }

        public PdfToFormatPage(PdfTargetFormat format)
        {
            InitializeComponent();
            _format = format;

            FileListView.ItemsSource = _files;

            TxtHint.Text = GetFormatDescription(format);
            BtnExecute.Content = "开始转换";

            UpdateEmptyHint();
        }

        /// <summary>各格式的用途说明（显示在设置区）。</summary>
        private static string GetFormatDescription(PdfTargetFormat format)
        {
            switch (format)
            {
                case PdfTargetFormat.Svg:
                    return "SVG：矢量格式，每页生成一个文件，可无限放大不失真";
                case PdfTargetFormat.Markdown:
                    return "Markdown：提取文字与表格结构，便于二次编辑或发布";
                case PdfTargetFormat.Gray:
                    return "灰度：转成黑白，体积通常明显变小（会丢失颜色）";
                case PdfTargetFormat.PdfA:
                    return "PDF/A-2B：国际归档标准，长期保存与合规提交用";
                case PdfTargetFormat.Xps:
                    return "XPS：Windows 固定版式文档，适合打印与归档";
                default:
                    return "HTML：保留版式与图片，适合网页预览与分享";
            }
        }

        // ───────────────── 文件列表 ─────────────────

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog
            {
                Title = "选择要转换的PDF文件",
                Filter = "PDF 文档 (*.pdf)|*.pdf|所有文件 (*.*)|*.*",
                Multiselect = true
            };
            if (d.ShowDialog() == true) AddFiles(d.FileNames);
        }

        private void FileListView_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                AddFiles((string[])e.Data.GetData(DataFormats.FileDrop));
        }

        private void FileListView_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private void BtnRemoveFile_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement el && el.Tag is PdfToFormatItem item)
            {
                _files.Remove(item);
                UpdateEmptyHint();
            }
        }

        private void AddFiles(IEnumerable<string> paths)
        {
            foreach (var p in paths)
            {
                if (!string.Equals(Path.GetExtension(p), ".pdf", StringComparison.OrdinalIgnoreCase)) continue;
                if (_files.Any(f => f.FilePath.Equals(p, StringComparison.OrdinalIgnoreCase))) continue;

                long size = 0;
                try { size = new FileInfo(p).Length; } catch { }

                _files.Add(new PdfToFormatItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p,
                    SizeText = FormatSize(size),
                    StatusText = "待转换"
                });
            }
            UpdateEmptyHint();
        }

        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 5) return;
            double total = FileListView.ActualWidth - 14;
            double fixedWidth = 70 + 132 + 76 + 72;
            double name = Math.Max(90, total - fixedWidth);
            if (Math.Abs(gv.Columns[0].Width - name) > 1) gv.Columns[0].Width = name;
        }

        // ───────────────── 执行 ─────────────────

        private async void BtnExecute_Click(object sender, RoutedEventArgs e)
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

            var snapshot = _files.ToList();

            SetBusy(true);
            TxtStatus.Text = "准备中...";

            try
            {
                await Task.Run(() => DoWork(snapshot));
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "转换失败: " + ex.Message;
                MessageBox.Show(ex.Message, "转换错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetBusy(false);
                ReleaseMemory();
            }
        }

        private void DoWork(List<PdfToFormatItem> items)
        {
            string dir = AppConfig.GetFeatureOutputDir(GetFeatureName());
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            int total = items.Count;
            int ok = 0;
            var failures = new List<string>();

            using (var converter = new PdfFormatConverter())
            {
                for (int i = 0; i < total; i++)
                {
                    int index = i;
                    var item = items[i];

                    Dispatcher.Invoke(() =>
                    {
                        TxtStatus.Text = $"正在转换 ({index + 1}/{total}) {item.FileName}";
                        item.StatusText = "转换中...";
                    });

                    try
                    {
                        var r = converter.Convert(item.FilePath, dir, _format);

                        Dispatcher.Invoke(() =>
                        {
                            item.StatusText = "已完成";
                            item.OutputText = _format == PdfTargetFormat.Svg
                                ? $"{r.OutputCount} 个 SVG"
                                : FormatSize(r.NewSize);
                            ok++;
                        });
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            item.StatusText = "失败";
                            item.OutputText = "—";
                            failures.Add($"{item.FileName}: " + Shorten(ex.Message));
                        });
                    }
                }
            }

            Dispatcher.Invoke(() =>
            {
                TxtStatus.Text = $"完成：成功 {ok} 个，失败 {failures.Count} 个（已打开输出目录）";
                TxtStatus.ToolTip = dir;

                if (ok > 0)
                {
                    try { Process.Start(dir); } catch { }
                }

                if (failures.Count > 0)
                {
                    var msg = $"有 {failures.Count} 个文件转换失败：\n\n"
                            + string.Join("\n", failures.Take(6))
                            + (failures.Count > 6 ? $"\n… 共 {failures.Count} 个" : "");
                    MessageBox.Show(msg, "转换失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            });
        }

        /// <summary>输出目录名：各格式独立（与导航名称一致）。</summary>
        private string GetFeatureName()
        {
            switch (_format)
            {
                case PdfTargetFormat.Svg: return "PDF转SVG";
                case PdfTargetFormat.Markdown: return "PDF转Markdown";
                case PdfTargetFormat.Gray: return "PDF转灰度";
                case PdfTargetFormat.PdfA: return "PDF转PDFA";
                case PdfTargetFormat.Xps: return "PDF转XPS";
                default: return "PDF转HTML";
            }
        }

        private void SetBusy(bool busy)
        {
            BtnExecute.IsEnabled = !busy;
            BtnAddFiles.IsEnabled = !busy;
        }

        private void UpdateEmptyHint()
        {
            EmptyOverlay.Visibility = _files.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private static string FormatSize(long bytes)
        {
            if (bytes <= 0) return "—";
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024.0).ToString("0.0") + " KB";
            return (bytes / 1024.0 / 1024.0).ToString("0.00") + " MB";
        }

        private static string Shorten(string s)
        {
            if (string.IsNullOrEmpty(s)) return "未知错误";
            return s.Length > 80 ? s.Substring(0, 80) + "…" : s;
        }

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
    }

    /// <summary>列表项模型（实现通知，界面才会实时刷新「输出」「状态」列）。</summary>
    public class PdfToFormatItem : System.ComponentModel.INotifyPropertyChanged
    {
        private string _outputText = "—";
        private string _statusText = "待转换";

        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string SizeText { get; set; }

        public string OutputText
        {
            get => _outputText;
            set { if (_outputText != value) { _outputText = value; OnPropertyChanged(nameof(OutputText)); } }
        }

        public string StatusText
        {
            get => _statusText;
            set { if (_statusText != value) { _statusText = value; OnPropertyChanged(nameof(StatusText)); } }
        }

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string name)
            => PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(name));
    }
}
