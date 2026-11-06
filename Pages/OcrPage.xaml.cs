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
    /// <summary>OCR 列表项。</summary>
    public class OcrItem : INotifyPropertyChanged
    {
        private string _typeText = "—";
        private string _resultText = "—";
        private string _statusText = "待识别";

        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string RecognizedText { get; set; }

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

    public partial class OcrPage : UserControl
    {
        private static readonly string[] SupportedExtensions =
            { ".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff" };

        private readonly ObservableCollection<OcrItem> _files = new ObservableCollection<OcrItem>();

        public OcrPage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;
            PreviewKeyDown += Page_PreviewKeyDown;
            UpdateEmptyHint();
            UpdateModelHint();
        }

        // ───────────────── 模型状态 ─────────────────

        private void UpdateModelHint()
        {
            if (OcrModelManager.IsReady)
            {
                TxtHint.Text = $"OCR 引擎已就绪（缓存 {OcrModelManager.CacheSizeMb:0.0} MB，离线可用）；手写体不支持";
                TxtHint.Foreground = System.Windows.Media.Brushes.SlateGray;
            }
            else
            {
                TxtHint.Text = "首次使用需下载约 15 MB 模型（之后离线可用），点「开始识别」会自动下载";
                TxtHint.Foreground = System.Windows.Media.Brushes.DarkOrange;
            }
        }

        // ───────────────── 剪贴板粘贴识别 ─────────────────

        /// <summary>识别剪贴板中最新的图片：先把图片落到临时文件，再加入列表并立即识别。</summary>
        private async void BtnPaste_Click(object sender, RoutedEventArgs e)
        {
            if (BtnExecute.IsEnabled == false) return;   // 正在识别中

            var item = TryGetClipboardImageItem();
            if (item == null)
            {
                MessageBox.Show("剪贴板里没有图片。\n先截图或复制一张图片，再点「粘贴识别」。",
                    "没有可识别的图片", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            _files.Add(item);
            UpdateEmptyHint();
            await RunRecognitionAsync(new List<OcrItem> { item });
        }

        /// <summary>Ctrl+V：只有在剪贴板含图片时拦截，文本粘贴仍交给输入框处理。</summary>
        private async void Page_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key != System.Windows.Input.Key.V) return;
            if ((System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control) != System.Windows.Input.ModifierKeys.Control) return;
            if (!System.Windows.Clipboard.ContainsImage()) return;
            if (BtnExecute.IsEnabled == false) return;

            e.Handled = true;

            var item = TryGetClipboardImageItem();
            if (item == null) return;

            _files.Add(item);
            UpdateEmptyHint();
            await RunRecognitionAsync(new List<OcrItem> { item });
        }

        /// <summary>把剪贴板图片保存为临时 PNG 并构造列表项；没有图片返回 null。</summary>
        /// <summary>清掉 3 天前的剪贴板临时图片，避免临时目录越积越多。</summary>
        private static void CleanupOldClipboardFiles(string dir)
        {
            try
            {
                var limit = DateTime.Now.AddDays(-3);
                foreach (var file in Directory.GetFiles(dir, "*.png"))
                {
                    if (File.GetLastWriteTime(file) < limit) File.Delete(file);
                }
            }
            catch { }
        }

        private OcrItem TryGetClipboardImageItem()
        {
            try
            {
                if (!System.Windows.Clipboard.ContainsImage()) return null;

                var source = System.Windows.Clipboard.GetImage();
                if (source == null) return null;

                var dir = Path.Combine(Path.GetTempPath(), "DocToPdfTool", "clipboard");
                Directory.CreateDirectory(dir);
                CleanupOldClipboardFiles(dir);

                var fileName = "剪贴板_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".png";
                var path = Path.Combine(dir, fileName);

                var encoder = new System.Windows.Media.Imaging.PngBitmapEncoder();
                encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(source));
                using (var stream = File.Create(path))
                    encoder.Save(stream);

                return new OcrItem
                {
                    FileName = fileName,
                    FilePath = path,
                    TypeText = "PNG",
                    StatusText = "来自剪贴板"
                };
            }
            catch (Exception ex)
            {
                MessageBox.Show("读取剪贴板图片失败：" + ex.Message, "粘贴失败",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return null;
            }
        }
        private void BtnOpenModels_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dir = OcrModelManager.ModelDirectory;
                Directory.CreateDirectory(dir);
                Process.Start("explorer.exe", "\"" + dir + "\"");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "打开失败", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ───────────────── 文件列表 ─────────────────

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "选择要识别的图片或 PDF",
                Filter = "图片与 PDF|*.pdf;*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff|所有文件 (*.*)|*.*",
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
            if (sender is FrameworkElement element && element.Tag is OcrItem item)
            {
                _files.Remove(item);
                UpdateEmptyHint();
            }
        }

        /// <summary>切到某个文件时，右侧结果区显示它已识别的文本。</summary>
        private void FileListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FileListView.SelectedItem is OcrItem item && !string.IsNullOrEmpty(item.RecognizedText))
            {
                ShowResult(item.FileName, item.RecognizedText);
            }
        }

        private void AddFiles(IEnumerable<string> paths)
        {
            foreach (var path in paths)
            {
                var extension = Path.GetExtension(path);
                if (!SupportedExtensions.Any(x => string.Equals(x, extension, StringComparison.OrdinalIgnoreCase)))
                    continue;
                if (_files.Any(f => f.FilePath.Equals(path, StringComparison.OrdinalIgnoreCase)))
                    continue;

                _files.Add(new OcrItem
                {
                    FileName = Path.GetFileName(path),
                    FilePath = path,
                    TypeText = extension.TrimStart('.').ToUpperInvariant()
                });
            }
            UpdateEmptyHint();
        }

        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!(FileListView.View is GridView gridView) || gridView.Columns.Count < 5) return;
            double total = FileListView.ActualWidth - 14;
            double fixedWidth = 62 + 140 + 100 + 70;
            double name = Math.Max(110, total - fixedWidth);
            if (Math.Abs(gridView.Columns[0].Width - name) > 1) gridView.Columns[0].Width = name;
        }

        // ───────────────── 结果展示 ─────────────────

        private void ShowResult(string fileName, string text)
        {
            TxtResultFile.Text = fileName;
            TxtResult.Text = text;
        }

        private void BtnCopy_Click(object sender, RoutedEventArgs e)
        {
            var text = TxtResult.Text;
            if (string.IsNullOrWhiteSpace(text))
            {
                TxtStatus.Text = "没有可复制的内容";
                return;
            }

            try
            {
                Clipboard.SetText(text);
                TxtStatus.Text = $"已复制 {text.Length} 个字符";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "复制失败：" + ex.Message;
            }
        }

        private void BtnSaveAs_Click(object sender, RoutedEventArgs e)
        {
            var text = TxtResult.Text;
            if (string.IsNullOrWhiteSpace(text))
            {
                TxtStatus.Text = "没有可保存的内容";
                return;
            }

            var dialog = new SaveFileDialog
            {
                Title = "保存识别结果",
                Filter = "文本文件 (*.txt)|*.txt",
                FileName = (TxtResultFile.Text ?? "识别结果").Replace(".pdf", "").Replace(".PDF", "")
                           + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt"
            };
            if (dialog.ShowDialog() != true) return;

            try
            {
                File.WriteAllText(dialog.FileName, text, new UTF8Encoding(true));
                TxtStatus.Text = "已保存：" + Path.GetFileName(dialog.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "保存失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ───────────────── 执行 ─────────────────

        private OcrOptions BuildOptions()
        {
            int dpi = 300;
            if (CmbDpi.SelectedIndex == 0) dpi = 200;
            else if (CmbDpi.SelectedIndex == 2) dpi = 400;

            return new OcrOptions
            {
                Dpi = dpi,
                PageRange = (TxtPageRange.Text ?? string.Empty).Trim(),
                PageSeparator = ChkPageSeparator.IsChecked == true
            };
        }

        private async void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            if (_files.Count == 0)
            {
                MessageBox.Show("请先添加图片或 PDF", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var missing = _files.FirstOrDefault(f => !File.Exists(f.FilePath));
            if (missing != null)
            {
                MessageBox.Show($"文件不存在: {missing.FilePath}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            await RunRecognitionAsync(_files.ToList());
        }

        /// <summary>统一的识别入口：按钮与「粘贴识别」都走这里。</summary>
        private async Task RunRecognitionAsync(List<OcrItem> items)
        {
            if (items == null || items.Count == 0) return;

            var options = BuildOptions();

            SetBusy(true);
            TxtStatus.Text = "准备中…";
            try
            {
                await Task.Run(() => DoWork(items, options));
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "识别失败：" + ex.Message;
                MessageBox.Show(ex.Message, "识别失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetBusy(false);
                ReleaseMemory();
                UpdateModelHint();
            }
        }

        private void SetBusy(bool busy)
        {
            BtnExecute.IsEnabled = !busy;
            BtnAddFiles.IsEnabled = !busy;
            CmbDpi.IsEnabled = !busy;
            TxtPageRange.IsEnabled = !busy;
            ChkPageSeparator.IsEnabled = !busy;
        }

        private void DoWork(List<OcrItem> items, OcrOptions options)
        {
            string outputDir = AppConfig.GetFeatureOutputDir("OCR文字识别");
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            if (!OcrModelManager.IsReady)
            {
                Dispatcher.Invoke(() => TxtStatus.Text = "正在下载 OCR 模型…");
                OcrModelManager.EnsureModelsAsync(
                    (message, percent) => Dispatcher.Invoke(() =>
                        TxtStatus.Text = percent >= 100 ? "模型准备完成" : $"{message} {percent}%"))
                    .GetAwaiter().GetResult();
                Dispatcher.Invoke(UpdateModelHint);
            }

            int total = items.Count, success = 0, empty = 0, chars = 0;
            var failures = new List<string>();

            using (var converter = new OcrConverter())
            {
                for (int i = 0; i < total; i++)
                {
                    int index = i;
                    var item = items[i];

                    Dispatcher.Invoke(() =>
                    {
                        TxtStatus.Text = $"正在识别 ({index + 1}/{total}) {item.FileName}";
                        item.StatusText = "识别中…";
                    });

                    try
                    {
                        var result = converter.RecognizeFileAsync(item.FilePath, outputDir, options,
                            (done, pages) => Dispatcher.Invoke(() =>
                                TxtStatus.Text = $"正在识别 ({index + 1}/{total}) {item.FileName} — 第 {done}/{pages} 页"))
                            .GetAwaiter().GetResult();

                        var text = result.OutputFile != null && File.Exists(result.OutputFile)
                            ? File.ReadAllText(result.OutputFile, Encoding.UTF8)
                            : string.Empty;

                        Dispatcher.Invoke(() =>
                        {
                            item.RecognizedText = text;
                            item.ResultText = $"{result.CharCount} 字 / {result.PageCount} 页";
                            item.StatusText = result.Empty ? "未识别到文字" : "完成";
                            if (result.Empty) empty++; else success++;
                            chars += result.CharCount;

                            // 识别完直接切到结果页展示，方便复制
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                ShowResult(item.FileName, text);
                                ResultTabs.SelectedItem = ResultTab;
                            }
                        });
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
            }

            Dispatcher.Invoke(() =>
            {
                TxtStatus.Text = failures.Count > 0
                    ? $"完成：成功 {success} 个，失败 {failures.Count} 个"
                    : $"全部完成：{success} 个文件，共 {chars} 字（结果可直接复制）";

                if (failures.Count > 0)
                {
                    MessageBox.Show(
                        $"识别完成\n\n成功：{success} 个\n失败：{failures.Count} 个\n\n失败原因：\n"
                        + string.Join("\n", failures.Take(6))
                        + (failures.Count > 6 ? $"\n… 共 {failures.Count} 个" : ""),
                        "识别结果", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            });
        }

        private static string Shorten(string text)
        {
            if (string.IsNullOrEmpty(text)) return "未知错误";
            return text.Length > 80 ? text.Substring(0, 80) + "…" : text;
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
