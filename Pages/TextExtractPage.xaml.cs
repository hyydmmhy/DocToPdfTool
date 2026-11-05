using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using DocToPdfTool.Utils;
using Microsoft.Win32;

namespace DocToPdfTool.Pages
{
    /// <summary>纯文本提取页面的列表项。</summary>
    public class TextExtractItem : INotifyPropertyChanged
    {
        private string _typeText = "—";
        private string _resultText = "—";
        private string _statusText = "待提取";

        public string FileName { get; set; }
        public string FilePath { get; set; }

        public string TypeText
        {
            get => _typeText;
            set { if (_typeText != value) { _typeText = value; OnPropertyChanged(); } }
        }

        public string ResultText
        {
            get => _resultText;
            set { if (_resultText != value) { _resultText = value; OnPropertyChanged(); } }
        }

        public string StatusText
        {
            get => _statusText;
            set { if (_statusText != value) { _statusText = value; OnPropertyChanged(); } }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public partial class TextExtractPage : UserControl
    {
        private static readonly string[] SupportedExtensions =
            { ".pdf", ".doc", ".docx", ".rtf", ".ppt", ".pptx", ".txt", ".md" };

        private readonly ObservableCollection<TextExtractItem> _files =
            new ObservableCollection<TextExtractItem>();

        public TextExtractPage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;
            UpdateEmptyHint();
        }

        // ───────────────── 文件列表 ─────────────────

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog
            {
                Title = "选择要提取文字的文件",
                Filter = "支持的文件|*.pdf;*.doc;*.docx;*.rtf;*.ppt;*.pptx;*.txt;*.md|所有文件 (*.*)|*.*",
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
            if (sender is FrameworkElement el && el.Tag is TextExtractItem item)
            {
                _files.Remove(item);
                UpdateEmptyHint();
            }
        }

        private void AddFiles(IEnumerable<string> paths)
        {
            foreach (var p in paths)
            {
                var ext = Path.GetExtension(p);
                if (!SupportedExtensions.Any(e => string.Equals(e, ext, StringComparison.OrdinalIgnoreCase))) continue;
                if (_files.Any(f => f.FilePath.Equals(p, StringComparison.OrdinalIgnoreCase))) continue;

                _files.Add(new TextExtractItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p,
                    TypeText = ext.TrimStart('.').ToUpperInvariant()
                });
            }
            UpdateEmptyHint();
        }

        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 5) return;
            double total = FileListView.ActualWidth - 14;
            double fixedWidth = 70 + 150 + 110 + 72;
            double name = Math.Max(120, total - fixedWidth);
            if (Math.Abs(gv.Columns[0].Width - name) > 1) gv.Columns[0].Width = name;
        }

        // ───────────────── 执行 ─────────────────

        private TextExtractOptions BuildOptions()
        {
            TextOutputEncoding enc = TextOutputEncoding.Utf8Bom;
            if (CmbEncoding.SelectedIndex == 0) enc = TextOutputEncoding.Utf8;
            else if (CmbEncoding.SelectedIndex == 2) enc = TextOutputEncoding.Gbk;

            return new TextExtractOptions
            {
                Encoding = enc,
                RemoveRepeatedHeaders = ChkRemoveHeaders.IsChecked == true,
                CollapseBlankLines = ChkCollapseBlank.IsChecked == true,
                PageSeparator = ChkPageSeparator.IsChecked == true
            };
        }

        private bool IsMergeMode => CmbOutputMode.SelectedIndex == 1;

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

            var options = BuildOptions();
            bool merge = IsMergeMode;
            var snapshot = _files.ToList();

            SetBusy(true);
            TxtStatus.Text = "准备中...";
            try
            {
                if (merge) await Task.Run(() => DoMerge(snapshot, options));
                else await Task.Run(() => DoExtractEach(snapshot, options));
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "提取失败: " + ex.Message;
                MessageBox.Show(ex.Message, "提取错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetBusy(false);
                ReleaseMemory();
            }
        }

        private void SetBusy(bool busy)
        {
            BtnExecute.IsEnabled = !busy;
            BtnAddFiles.IsEnabled = !busy;
            CmbEncoding.IsEnabled = !busy;
            CmbOutputMode.IsEnabled = !busy;
            ChkRemoveHeaders.IsEnabled = !busy;
            ChkCollapseBlank.IsEnabled = !busy;
            ChkPageSeparator.IsEnabled = !busy;
        }

        /// <summary>每个文件生成一个 txt。</summary>
        private void DoExtractEach(List<TextExtractItem> items, TextExtractOptions options)
        {
            string dir = AppConfig.GetFeatureOutputDir("提取纯文本");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            int total = items.Count, success = 0, empty = 0, chars = 0;
            var failures = new List<string>();

            for (int i = 0; i < total; i++)
            {
                int index = i;
                var item = items[i];
                Dispatcher.Invoke(() =>
                {
                    TxtStatus.Text = $"正在提取 ({index + 1}/{total}) {item.FileName}";
                    item.StatusText = "提取中...";
                });

                try
                {
                    using (var conv = new TextExtractConverter())
                    {
                        var r = conv.Extract(item.FilePath, dir, options);
                        Dispatcher.Invoke(() =>
                        {
                            item.ResultText = $"{r.CharCount} 字" + (r.PageCount > 0 ? $" / {r.PageCount} 页" : "");
                            item.StatusText = r.Empty ? "无文字" : "完成";
                            if (r.Empty) empty++; else success++;
                            chars += r.CharCount;
                        });
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        item.StatusText = "失败";
                        failures.Add($"{item.FileName}: " + Shorten(ex.Message));
                    });
                }
            }

            Dispatcher.Invoke(() =>
            {
                TxtStatus.Text = failures.Count > 0
                    ? $"完成：成功 {success} 个，失败 {failures.Count} 个"
                    : $"全部完成：{success} 个文件，共 {chars} 字";

                var msg = $"提取完成\n\n成功：{success} 个\n合计：{chars} 字"
                        + (empty > 0 ? $"\n无文字：{empty} 个（可能是扫描件，需 OCR）" : "")
                        + (failures.Count > 0 ? $"\n失败：{failures.Count} 个" : "")
                        + $"\n\n输出目录：\n{dir}";

                if (failures.Count > 0)
                {
                    msg += "\n\n失败原因：\n" + string.Join("\n", failures.Take(6))
                         + (failures.Count > 6 ? $"\n… 共 {failures.Count} 个" : "");
                }

                MessageBox.Show(msg, "提取结果", MessageBoxButton.OK,
                    failures.Count > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);

                if (success + empty > 0) { try { Process.Start(dir); } catch { } }
            });
        }

        /// <summary>全部合并成一个 txt。</summary>
        private void DoMerge(List<TextExtractItem> items, TextExtractOptions options)
        {
            string dir = AppConfig.GetFeatureOutputDir("提取纯文本");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            int total = items.Count;
            var sb = new StringBuilder();
            var failures = new List<string>();
            int success = 0, empty = 0;

            for (int i = 0; i < total; i++)
            {
                int index = i;
                var item = items[i];
                Dispatcher.Invoke(() =>
                {
                    TxtStatus.Text = $"正在提取 ({index + 1}/{total}) {item.FileName}";
                    item.StatusText = "提取中...";
                });

                string tempDir = Path.Combine(Path.GetTempPath(), "DocToPdfTool_txt_" + Guid.NewGuid().ToString("N"));
                try
                {
                    using (var conv = new TextExtractConverter())
                    {
                        var r = conv.Extract(item.FilePath, tempDir, options);
                        var text = string.IsNullOrEmpty(r.OutputFile) || !File.Exists(r.OutputFile)
                            ? string.Empty
                            : File.ReadAllText(r.OutputFile, Encoding.UTF8);

                        sb.AppendLine("==================================================");
                        sb.AppendLine($"文件：{item.FileName}");
                        if (r.PageCount > 0) sb.AppendLine($"页数：{r.PageCount}");
                        sb.AppendLine("==================================================");
                        sb.AppendLine(text);
                        sb.AppendLine();

                        Dispatcher.Invoke(() =>
                        {
                            item.ResultText = $"{r.CharCount} 字";
                            item.StatusText = r.Empty ? "无文字" : "完成";
                            if (r.Empty) empty++; else success++;
                        });
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        item.StatusText = "失败";
                        failures.Add($"{item.FileName}: " + Shorten(ex.Message));
                    });
                }
                finally
                {
                    try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
                }
            }

            string outputFile = UniquePath(dir, "合并文本_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            WriteText(outputFile, sb.ToString().TrimEnd() + Environment.NewLine, options.Encoding);
            long size = new FileInfo(outputFile).Length;

            Dispatcher.Invoke(() =>
            {
                TxtStatus.Text = $"合并完成：{Path.GetFileName(outputFile)}（{TextExtractConverter.FormatSize(size)}）";
                MessageBox.Show(
                    $"合并完成\n\n成功：{success} 个"
                    + (empty > 0 ? $"\n无文字：{empty} 个" : "")
                    + (failures.Count > 0 ? $"\n失败：{failures.Count} 个" : "")
                    + $"\n\n输出文件：\n{outputFile}",
                    "合并结果", MessageBoxButton.OK,
                    failures.Count > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);

                try { Process.Start(dir); } catch { }
            });
        }

        private static void WriteText(string path, string text, TextOutputEncoding encoding)
        {
            switch (encoding)
            {
                case TextOutputEncoding.Utf8:
                    File.WriteAllText(path, text, new UTF8Encoding(false));
                    break;
                case TextOutputEncoding.Gbk:
                    File.WriteAllText(path, text, Encoding.GetEncoding(936));
                    break;
                default:
                    File.WriteAllText(path, text, new UTF8Encoding(true));
                    break;
            }
        }

        private static string UniquePath(string dir, string baseName)
        {
            var path = Path.Combine(dir, baseName + ".txt");
            int seq = 1;
            while (File.Exists(path)) path = Path.Combine(dir, $"{baseName}_{seq++}.txt");
            return path;
        }

        private static string Shorten(string s)
        {
            if (string.IsNullOrEmpty(s)) return "未知错误";
            return s.Length > 80 ? s.Substring(0, 80) + "…" : s;
        }

        private void UpdateEmptyHint()
        {
            EmptyOverlay.Visibility = _files.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
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
}
