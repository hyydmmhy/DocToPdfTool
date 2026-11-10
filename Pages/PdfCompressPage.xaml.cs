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
    public partial class PdfCompressPage : UserControl
    {
        private readonly ObservableCollection<PdfCompressItem> _files =
            new ObservableCollection<PdfCompressItem>();

        public PdfCompressPage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;
            UpdateEmptyHint();
            UpdateCustomControls();
        }

        // ───────────────── 文件列表 ─────────────────

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog
            {
                Title = "选择要压缩的PDF文件",
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
            if (sender is FrameworkElement el && el.Tag is PdfCompressItem item)
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

                _files.Add(new PdfCompressItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p,
                    OriginalText = size > 0 ? PdfCompressConverter.FormatSize(size) : "—",
                    StatusText = PdfEncryptConverter.IsEncrypted(p) ? "已加密" : "待压缩"
                });
            }
            UpdateEmptyHint();
        }

        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 5) return;
            double total = FileListView.ActualWidth - 14;
            double fixedWidth = 68 + 132 + 76 + 72;   // 原始大小 + 压缩后 + 状态 + 操作
            double name = Math.Max(90, total - fixedWidth);
            if (Math.Abs(gv.Columns[0].Width - name) > 1) gv.Columns[0].Width = name;
        }

        // ───────────────── 档位联动 ─────────────────

        private PdfCompressPreset GetPreset()
        {
            if (RdoStandard.IsChecked == true) return PdfCompressPreset.Standard;
            if (RdoStrong.IsChecked == true) return PdfCompressPreset.Strong;
            if (RdoExtreme.IsChecked == true) return PdfCompressPreset.Extreme;
            if (RdoCustom.IsChecked == true) return PdfCompressPreset.Custom;
            return PdfCompressPreset.Light;
        }

        private void Level_Changed(object sender, RoutedEventArgs e)
        {
            UpdateCustomControls();
        }

        /// <summary>自定义档位时才启用下拉与质量框，其余置灰但不隐藏，避免布局跳动。</summary>
        private void UpdateCustomControls()
        {
            if (CmbDpi == null || TxtQuality == null || TxtHint == null) return;

            bool custom = GetPreset() == PdfCompressPreset.Custom;
            CmbDpi.IsEnabled = custom;
            TxtQuality.IsEnabled = custom;
            CmbDpi.Opacity = custom ? 1.0 : 0.45;
            TxtQuality.Opacity = custom ? 1.0 : 0.45;

            switch (GetPreset())
            {
                case PdfCompressPreset.Standard:
                    TxtHint.Text = "标准：图片重采样到 200 DPI + 质量 88，体积与画质平衡（推荐）";
                    break;
                case PdfCompressPreset.Strong:
                    TxtHint.Text = "强力：图片重采样到 150 DPI + 质量 82，体积更小，放大看会略糊";
                    break;
                case PdfCompressPreset.Extreme:
                    TxtHint.Text = "极限：图片重采样到 110 DPI + 质量 75，体积最小，适合只看内容";
                    break;
                case PdfCompressPreset.Custom:
                    TxtHint.Text = "自定义：目标分辨率与 JPEG 质量自由组合；有效 DPI 已低于目标的图片保持原样";
                    break;
                default:
                    TxtHint.Text = "轻度：不重采样，仅高质量重编码，画质接近原样；文字层保留，仍可搜索复制";
                    break;
            }
        }

        private int GetCustomDpi()
        {
            switch (CmbDpi.SelectedIndex)
            {
                case 0: return 0;      // 不重采样
                case 1: return 300;
                case 2: return 240;
                case 3: return 200;
                case 4: return 150;
                case 5: return 120;
                case 6: return 96;
                case 7: return 72;
                default: return 200;
            }
        }

        private int GetCustomQuality()
        {
            if (int.TryParse((TxtQuality.Text ?? "").Trim(), out int q) && q >= 50 && q <= 98) return q;
            return 88;
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

            var options = new PdfCompressOptions
            {
                Preset = GetPreset(),
                CustomDpi = GetCustomDpi(),
                CustomQuality = GetCustomQuality()
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
                TxtStatus.Text = "压缩失败: " + ex.Message;
                MessageBox.Show(ex.Message, "压缩错误", MessageBoxButton.OK, MessageBoxImage.Error);
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
            RdoLight.IsEnabled = !busy;
            RdoStandard.IsEnabled = !busy;
            RdoStrong.IsEnabled = !busy;
            RdoExtreme.IsEnabled = !busy;
            RdoCustom.IsEnabled = !busy;
            if (busy)
            {
                CmbDpi.IsEnabled = false;
                TxtQuality.IsEnabled = false;
            }
            else
            {
                UpdateCustomControls();
            }
        }

        private void DoWork(List<PdfCompressItem> items, PdfCompressOptions options)
        {
            string dir = AppConfig.GetFeatureOutputDir("PDF压缩");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            int total = items.Count;
            int compressed = 0, replacedTotal = 0;
            var skipped = new List<string>();
            var failures = new List<string>();

            for (int i = 0; i < total; i++)
            {
                int index = i;
                var item = items[i];
                Dispatcher.Invoke(() =>
                {
                    TxtStatus.Text = $"正在压缩 ({index + 1}/{total}) {item.FileName}";
                    item.StatusText = "压缩中...";
                });

                try
                {
                    using (var converter = new PdfCompressConverter())
                    {
                        var r = converter.Compress(item.FilePath, dir, options);

                        Dispatcher.Invoke(() =>
                        {
                            if (r.Skipped)
                            {
                                item.RatioText = "—";
                                item.NewText = PdfCompressConverter.FormatSize(r.OriginalSize) + "（无收益）";
                                item.StatusText = "已是最优";
                                skipped.Add(item.FileName);
                            }
                            else
                            {
                                item.NewText = PdfCompressConverter.FormatSize(r.NewSize)
                                             + "（" + (r.Ratio * 100).ToString("0.0") + "%）";
                                item.RatioText = (r.Ratio * 100).ToString("0.0") + "%";
                                item.StatusText = "已压缩";
                                compressed++;
                                replacedTotal += r.ReplacedImages;
                            }
                        });
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        item.StatusText = "失败";
                        item.RatioText = "—";
                        failures.Add($"{item.FileName}: " + Shorten(ex.Message));
                    });
                }
            }

            Dispatcher.Invoke(() =>
            {
                TxtStatus.Text = $"完成：压缩 {compressed} 个，跳过 {skipped.Count} 个，失败 {failures.Count} 个";

                // 完成不弹窗，自动打开输出目录（仅失败时弹错误提示）
                if (compressed > 0)
                {
                    try { Process.Start(dir); } catch { }
                }

                if (failures.Count > 0)
                {
                    var msg = "部分文件压缩失败：\n"
                            + string.Join("\n", failures.Take(6))
                            + (failures.Count > 6 ? $"\n… 共 {failures.Count} 个" : "");
                    MessageBox.Show(msg, "部分文件失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
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
