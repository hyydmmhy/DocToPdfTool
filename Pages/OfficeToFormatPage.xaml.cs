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
    /// Office 文档转其他格式的通用页面。
    /// 一个实现，多个入口（Word→Markdown / Word→EPUB / Excel→CSV / PPT→HTML / PPT→TIFF）。
    /// </summary>
    public partial class OfficeToFormatPage : UserControl
    {
        private readonly ObservableCollection<OfficeToFormatItem> _files =
            new ObservableCollection<OfficeToFormatItem>();

        private readonly OfficeSourceKind _kind;
        private readonly OfficeTargetFormat _target;
        private readonly string[] _extensions;

        /// <summary>设计器用默认构造。</summary>
        public OfficeToFormatPage() : this(OfficeSourceKind.Word, OfficeTargetFormat.Markdown) { }

        public OfficeToFormatPage(OfficeSourceKind kind, OfficeTargetFormat target)
        {
            InitializeComponent();

            _kind = kind;
            _target = target;
            _extensions = OfficeFormatConverter.GetExtensions(kind);

            FileListView.ItemsSource = _files;

            TxtHint.Text = GetHint(kind, target);
            TxtDropHint.Text = "或拖拽 " + GetKindLabel(kind) + " 文件到此处";
            TxtEmpty.Text = $"拖拽 {GetKindLabel(kind)} 文件到此处，或点击上方「添加文件」按钮";

            UpdateEmptyHint();
        }

        private static string GetKindLabel(OfficeSourceKind kind)
        {
            switch (kind)
            {
                case OfficeSourceKind.Excel: return "Excel";
                case OfficeSourceKind.Ppt: return "PPT";
                default: return "Word";
            }
        }

        private static string GetHint(OfficeSourceKind kind, OfficeTargetFormat target)
        {
            switch (target)
            {
                case OfficeTargetFormat.Markdown:
                    return "Markdown：保留标题层级、列表与表格，适合笔记与文档站发布";
                case OfficeTargetFormat.Epub:
                    return "EPUB：标准电子书格式，可在手机/平板阅读器打开";
                case OfficeTargetFormat.Csv:
                    return "CSV：纯文本表格，Excel/数据库/脚本都能直接读（多工作表时每张表一个文件）";
                case OfficeTargetFormat.Html:
                    return "HTML：幻灯片转成网页，可直接在浏览器翻页查看";
                case OfficeTargetFormat.Tiff:
                    return "TIFF：每页幻灯片一张图，合成多页 TIFF，适合打印与归档";
                case OfficeTargetFormat.ExcelHtml:
                    return "HTML：每张工作表一个网页，保留单元格样式与颜色";
                case OfficeTargetFormat.ExcelJson:
                    return "JSON：按工作表输出结构化数据，便于程序读取或导入系统";
                case OfficeTargetFormat.ExcelMarkdown:
                    return "Markdown：每张工作表转成表格，适合文档与笔记引用";
                default:
                    return string.Empty;
            }
        }

        // ───────────────── 文件列表 ─────────────────

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "选择要转换的文件",
                Filter = OfficeFormatConverter.GetFileFilter(_kind),
                Multiselect = true
            };
            if (dialog.ShowDialog() == true) AddFiles(dialog.FileNames);
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
            if (sender is FrameworkElement el && el.Tag is OfficeToFormatItem item)
            {
                _files.Remove(item);
                UpdateEmptyHint();
            }
        }

        private void AddFiles(IEnumerable<string> paths)
        {
            int skipped = 0;

            foreach (var p in paths)
            {
                var ext = Path.GetExtension(p) ?? string.Empty;
                if (!_extensions.Any(x => string.Equals(x, ext, StringComparison.OrdinalIgnoreCase)))
                {
                    skipped++;
                    continue;
                }
                if (_files.Any(f => f.FilePath.Equals(p, StringComparison.OrdinalIgnoreCase))) continue;

                long size = 0;
                try { size = new FileInfo(p).Length; } catch { }

                _files.Add(new OfficeToFormatItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p,
                    SizeText = FormatSize(size),
                    StatusText = "待转换"
                });
            }

            UpdateEmptyHint();

            if (skipped > 0)
                TxtStatus.Text = $"已忽略 {skipped} 个类型不匹配的文件（本页只处理 {GetKindLabel(_kind)}）";
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
                MessageBox.Show("请先添加文件", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
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

        private void DoWork(List<OfficeToFormatItem> items)
        {
            string dir = AppConfig.GetFeatureOutputDir(OfficeFormatConverter.GetFeatureName(_kind, _target));
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            int total = items.Count;
            int ok = 0;
            var failures = new List<string>();

            using (var converter = new OfficeFormatConverter())
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
                        var r = converter.Convert(item.FilePath, dir, _target);

                        Dispatcher.Invoke(() =>
                        {
                            item.StatusText = "已完成";
                            item.OutputText = FormatSize(r.NewSize);
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

    /// <summary>列表项模型（INotifyPropertyChanged 让「输出」「状态」列实时刷新）。</summary>
    public class OfficeToFormatItem : System.ComponentModel.INotifyPropertyChanged
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
