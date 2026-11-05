using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using DocToPdfTool.Utils;
using Microsoft.Win32;

namespace DocToPdfTool.Pages
{
    /// <summary>证件照排版页面的列表项。</summary>
    public class IdPhotoItem : INotifyPropertyChanged
    {
        private string _sizeText = "—";
        private string _resultText = "—";
        private string _statusText = "待排版";

        public string FileName { get; set; }
        public string FilePath { get; set; }

        public string SizeText
        {
            get => _sizeText;
            set { if (_sizeText != value) { _sizeText = value; OnPropertyChanged(); } }
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

    public partial class IdPhotoLayoutPage : UserControl
    {
        private static readonly string[] ImageExtensions =
            { ".jpg", ".jpeg", ".png", ".bmp" };

        private readonly ObservableCollection<IdPhotoItem> _files =
            new ObservableCollection<IdPhotoItem>();

        public IdPhotoLayoutPage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;
            UpdateEmptyHint();
            UpdatePreview();
        }

        // ───────────────── 文件列表 ─────────────────

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog
            {
                Title = "选择证件照原图",
                Filter = "图片文件 (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|所有文件 (*.*)|*.*",
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
            if (sender is FrameworkElement el && el.Tag is IdPhotoItem item)
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
                if (!ImageExtensions.Any(e => string.Equals(e, ext, StringComparison.OrdinalIgnoreCase))) continue;
                if (_files.Any(f => f.FilePath.Equals(p, StringComparison.OrdinalIgnoreCase))) continue;

                var item = new IdPhotoItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p
                };

                try
                {
                    using (var image = System.Drawing.Image.FromFile(p))
                        item.SizeText = image.Width + "×" + image.Height;
                }
                catch
                {
                    item.StatusText = "无法读取";
                }

                _files.Add(item);
            }
            UpdateEmptyHint();
        }

        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 5) return;
            double total = FileListView.ActualWidth - 14;
            double fixedWidth = 100 + 150 + 80 + 72;
            double name = Math.Max(110, total - fixedWidth);
            if (Math.Abs(gv.Columns[0].Width - name) > 1) gv.Columns[0].Width = name;
        }

        // ───────────────── 参数与预览 ─────────────────

        private IdPhotoSize GetPhotoSize()
        {
            switch (CmbSize.SelectedIndex)
            {
                case 0: return IdPhotoSize.SmallOneInch;
                case 2: return IdPhotoSize.LargeOneInch;
                case 3: return IdPhotoSize.SmallTwoInch;
                case 4: return IdPhotoSize.TwoInch;
                case 5: return IdPhotoSize.LargeTwoInch;
                default: return IdPhotoSize.OneInch;
            }
        }

        private PhotoPaper GetPaper()
        {
            switch (CmbPaper.SelectedIndex)
            {
                case 0: return PhotoPaper.Size5R;
                case 2: return PhotoPaper.SizeA4;
                default: return PhotoPaper.Size6R;
            }
        }

        private int GetGapMm()
        {
            if (int.TryParse((TxtGap.Text ?? "").Trim(), out int v) && v >= 0 && v <= 20) return v;
            return 2;
        }

        private void Option_Changed(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void Option_Changed(object sender, SelectionChangedEventArgs e)
        {
            UpdatePreview();
        }

        /// <summary>实时计算每张相纸能排多少张，给用户直观预期。</summary>
        private void UpdatePreview()
        {
            if (TxtPreview == null || CmbSize == null || CmbPaper == null) return;

            const int dpi = 300;
            IdPhotoLayoutConverter.GetPhotoSizeMm(GetPhotoSize(), out double pw, out double ph);
            IdPhotoLayoutConverter.GetPaperSizeMm(GetPaper(), out double aw, out double ah);

            int photoW = (int)Math.Round(pw / 25.4 * dpi);
            int photoH = (int)Math.Round(ph / 25.4 * dpi);
            int paperW = (int)Math.Round(aw / 25.4 * dpi);
            int paperH = (int)Math.Round(ah / 25.4 * dpi);
            int gap = (int)Math.Round(GetGapMm() / 25.4 * dpi);

            int columns = Math.Max(1, (paperW + gap) / (photoW + gap));
            int rows = Math.Max(1, (paperH + gap) / (photoH + gap));

            TxtPreview.Text = $"每张相纸可排 {columns} 列 × {rows} 行 = {columns * rows} 张"
                            + $"（单张 {photoW}×{photoH} px，相纸 {paperW}×{paperH} px，300 DPI）";
        }

        // ───────────────── 执行 ─────────────────

        private async void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            if (_files.Count == 0)
            {
                MessageBox.Show("请先添加照片", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var invalid = _files.FirstOrDefault(f => !File.Exists(f.FilePath));
            if (invalid != null)
            {
                MessageBox.Show($"文件不存在: {invalid.FilePath}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var options = new IdPhotoLayoutOptions
            {
                Size = GetPhotoSize(),
                Paper = GetPaper(),
                Dpi = 300,
                GapMm = GetGapMm(),
                CenterCrop = ChkCenterCrop.IsChecked == true
            };

            var snapshot = _files.ToList();
            SetBusy(true);
            TxtStatus.Text = "准备中...";
            try
            {
                await Task.Run(() => DoWork(snapshot, options));
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "排版失败: " + ex.Message;
                MessageBox.Show(ex.Message, "排版错误", MessageBoxButton.OK, MessageBoxImage.Error);
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
            CmbSize.IsEnabled = !busy;
            CmbPaper.IsEnabled = !busy;
            TxtGap.IsEnabled = !busy;
            ChkCenterCrop.IsEnabled = !busy;
        }

        private void DoWork(List<IdPhotoItem> items, IdPhotoLayoutOptions options)
        {
            string dir = AppConfig.GetFeatureOutputDir("证件照排版");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            int total = items.Count;
            int success = 0;
            var failures = new List<string>();
            int lastColumns = 0, lastRows = 0, lastCount = 0;

            for (int i = 0; i < total; i++)
            {
                int index = i;
                var item = items[i];
                Dispatcher.Invoke(() =>
                {
                    TxtStatus.Text = $"正在排版 ({index + 1}/{total}) {item.FileName}";
                    item.StatusText = "排版中...";
                });

                try
                {
                    using (var converter = new IdPhotoLayoutConverter())
                    {
                        var r = converter.Layout(item.FilePath, dir, options);
                        Dispatcher.Invoke(() =>
                        {
                            item.ResultText = $"{r.Columns}×{r.Rows} = {r.PhotoCount} 张";
                            item.StatusText = "完成";
                            success++;
                            lastColumns = r.Columns;
                            lastRows = r.Rows;
                            lastCount = r.PhotoCount;
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
                    : $"全部完成：{success} 张照片，每张相纸排 {lastColumns}×{lastRows} = {lastCount} 张";

                var msg = $"排版完成\n\n成功：{success} 张"
                        + $"\n每张相纸：{lastColumns} 列 × {lastRows} 行 = {lastCount} 张证件照"
                        + (failures.Count > 0 ? $"\n失败：{failures.Count} 张" : "")
                        + $"\n\n输出目录：\n{dir}\n\n用 300 DPI 或「实际大小、不缩放」打印即可。";

                if (failures.Count > 0)
                {
                    msg += "\n\n失败原因：\n" + string.Join("\n", failures.Take(6))
                         + (failures.Count > 6 ? $"\n… 共 {failures.Count} 张" : "");
                }

                MessageBox.Show(msg, "排版结果", MessageBoxButton.OK,
                    failures.Count > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);

                if (success > 0) { try { Process.Start(dir); } catch { } }
            });
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
