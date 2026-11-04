using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using DocToPdfTool.Utils;
using Microsoft.Win32;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;

namespace DocToPdfTool.Pages
{
    /// <summary>页面管理中的单个页面项。</summary>
    public class PdfPageItem : INotifyPropertyChanged
    {
        private int _rotation;
        private string _thumbnailPath;
        private string _displayText;

        /// <summary>源文件中的页码（1 基）。</summary>
        public int SourceIndex { get; set; }

        /// <summary>缩略图文件路径。</summary>
        public string ThumbnailPath
        {
            get => _thumbnailPath;
            set { if (_thumbnailPath != value) { _thumbnailPath = value; OnPropertyChanged(); } }
        }

        /// <summary>列表显示文本，如「第 3 页」。</summary>
        public string DisplayText
        {
            get => _displayText;
            set { if (_displayText != value) { _displayText = value; OnPropertyChanged(); } }
        }

        /// <summary>旋转角度 0/90/180/270。</summary>
        public int Rotation
        {
            get => _rotation;
            set
            {
                if (_rotation == value) return;
                _rotation = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(RotationText));
            }
        }

        public string RotationText => _rotation == 0 ? "" : "↻ " + _rotation + "°";

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public partial class PdfPageManagerPage : UserControl
    {
        private readonly ObservableCollection<PdfPageItem> _pages = new ObservableCollection<PdfPageItem>();
        private string _sourcePath;
        private string _thumbDir;

        public PdfPageManagerPage()
        {
            InitializeComponent();
            PageList.ItemsSource = _pages;
            UpdateUiState();
        }

        // ───────────────── 打开 / 重置 ─────────────────

        private async void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog
            {
                Title = "选择要管理页面的PDF",
                Filter = "PDF 文档 (*.pdf)|*.pdf|所有文件 (*.*)|*.*"
            };
            if (d.ShowDialog() != true) return;

            await LoadFileAsync(d.FileName);
        }

        private async void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_sourcePath)) return;
            await LoadFileAsync(_sourcePath);
        }

        private async Task LoadFileAsync(string path)
        {
            try
            {
                if (PdfEncryptConverter.IsEncrypted(path))
                {
                    MessageBox.Show("该 PDF 已加密，请先用「PDF加密 → 移除密码」解除后再管理页面",
                        "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                SetBusy(true);
                TxtStatus.Text = "正在读取页面...";

                _sourcePath = path;
                CleanupThumbnails();
                _pages.Clear();

                int pageCount = PdfEncryptConverter.TryGetPageCount(path);
                if (pageCount <= 0)
                {
                    TxtStatus.Text = "无法读取该 PDF 的页数";
                    MessageBox.Show("无法读取该 PDF，请确认文件是否损坏", "错误",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                _thumbDir = Path.Combine(Path.GetTempPath(),
                    "DocToPdfTool_pages_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(_thumbDir);

                // 先创建占位项，界面立即响应
                for (int i = 1; i <= pageCount; i++)
                    _pages.Add(new PdfPageItem { SourceIndex = i, DisplayText = "第 " + i + " 页" });

                TxtFileInfo.Text = $"{Path.GetFileName(path)}　共 {pageCount} 页";
                UpdateUiState();

                await RenderThumbnailsAsync(path);

                TxtStatus.Text = $"已加载 {pageCount} 页，可旋转 / 删除 / 调整顺序";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "读取失败: " + ex.Message;
                MessageBox.Show(ex.Message, "读取错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetBusy(false);
                UpdateUiState();
            }
        }

        /// <summary>逐页渲染缩略图（宽度 120px），渲染完成即刷新界面。</summary>
        private async Task RenderThumbnailsAsync(string pdfPath)
        {
            object file = null;
            PdfDocument pdfDoc = null;
            try
            {
                file = await StorageFile.GetFileFromPathAsync(pdfPath).AsTask().ConfigureAwait(false);
                pdfDoc = await PdfDocument.LoadFromFileAsync((StorageFile)file).AsTask().ConfigureAwait(false);

                for (int i = 0; i < _pages.Count; i++)
                {
                    var index = i;
                    object pageObj = null;
                    InMemoryRandomAccessStream stream = null;
                    try
                    {
                        pageObj = pdfDoc.GetPage((uint)index);
                        stream = new InMemoryRandomAccessStream();

                        var options = new PdfPageRenderOptions { DestinationWidth = 120 };
                        await ((PdfPage)pageObj).RenderToStreamAsync(stream, options)
                            .AsTask().ConfigureAwait(false);

                        stream.Seek(0);
                        var thumbPath = Path.Combine(_thumbDir, "p" + (index + 1) + ".png");
                        using (var managed = stream.AsStreamForRead())
                        using (var fs = File.Create(thumbPath))
                        {
                            await managed.CopyToAsync(fs).ConfigureAwait(false);
                        }

                        Dispatcher.Invoke(() =>
                        {
                            if (index < _pages.Count)
                                _pages[index].ThumbnailPath = thumbPath;
                        });
                    }
                    catch { }
                    finally
                    {
                        SafeRelease(pageObj);
                        try { stream?.Dispose(); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => TxtStatus.Text = "缩略图渲染失败: " + ex.Message);
            }
            finally
            {
                SafeRelease(pdfDoc);
                SafeRelease(file);
            }
        }

        private static void SafeRelease(object obj)
        {
            if (obj == null) return;
            try { Marshal.FinalReleaseComObject(obj); } catch { }
        }

        // ───────────────── 页面操作 ─────────────────

        private void BtnRotateLeft_Click(object sender, RoutedEventArgs e) => RotateSelected(-90);
        private void BtnRotateRight_Click(object sender, RoutedEventArgs e) => RotateSelected(90);

        private void RotateSelected(int delta)
        {
            var selected = PageList.SelectedItems.Cast<PdfPageItem>().ToList();
            if (selected.Count == 0) return;

            foreach (var item in selected)
                item.Rotation = ((item.Rotation + delta) % 360 + 360) % 360;

            TxtStatus.Text = $"已旋转 {selected.Count} 页";
        }

        private void BtnMoveUp_Click(object sender, RoutedEventArgs e) => MoveSelected(-1);
        private void BtnMoveDown_Click(object sender, RoutedEventArgs e) => MoveSelected(1);

        private void MoveSelected(int delta)
        {
            var selected = PageList.SelectedItems.Cast<PdfPageItem>().ToList();
            if (selected.Count == 0) return;

            // 上移从前往后处理，下移从后往前处理，避免相互挤占
            var ordered = delta < 0 ? selected.OrderBy(p => _pages.IndexOf(p)).ToList()
                                    : selected.OrderByDescending(p => _pages.IndexOf(p)).ToList();

            foreach (var item in ordered)
            {
                int idx = _pages.IndexOf(item);
                int target = idx + delta;
                if (target < 0 || target >= _pages.Count) continue;
                _pages.Move(idx, target);
            }

            foreach (var item in selected) PageList.SelectedItems.Add(item);
            TxtStatus.Text = "已调整 " + selected.Count + " 页顺序";
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            var selected = PageList.SelectedItems.Cast<PdfPageItem>().ToList();
            if (selected.Count == 0) return;

            if (selected.Count >= _pages.Count)
            {
                MessageBox.Show("不能删除全部页面，至少保留一页", "提示",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            foreach (var item in selected) _pages.Remove(item);
            UpdateUiState();
            TxtStatus.Text = $"已删除 {selected.Count} 页，剩余 {_pages.Count} 页";
        }

        // ───────────────── 导出 ─────────────────

        private async void BtnExportAll_Click(object sender, RoutedEventArgs e)
        {
            await ExportAsync(_pages.ToList(), "页面管理");
        }

        private async void BtnExportSelected_Click(object sender, RoutedEventArgs e)
        {
            var selected = PageList.SelectedItems.Cast<PdfPageItem>().ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show("请先在列表中选择要导出的页面", "提示",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // 按当前顺序导出选中页
            selected = selected.OrderBy(p => _pages.IndexOf(p)).ToList();
            await ExportAsync(selected, "选中页");
        }

        private async Task ExportAsync(List<PdfPageItem> pages, string suffix)
        {
            if (string.IsNullOrEmpty(_sourcePath) || pages.Count == 0) return;

            SetBusy(true);
            TxtStatus.Text = "正在导出...";
            try
            {
                var keep = pages.Select(p => p.SourceIndex).ToList();
                var rotations = pages.Select(p => p.Rotation).ToList();

                string output = await Task.Run(() =>
                {
                    string dir = AppConfig.GetFeatureOutputDir("PDF页面管理");
                    using (var manager = new PdfPageManager())
                        return manager.Save(_sourcePath, dir, keep, rotations, suffix);
                });

                TxtStatus.Text = $"已导出 {pages.Count} 页：" + Path.GetFileName(output);
                try { Process.Start(Path.GetDirectoryName(output)); } catch { }
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "导出失败: " + ex.Message;
                MessageBox.Show(ex.Message, "导出错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetBusy(false);
                UpdateUiState();
            }
        }

        // ───────────────── 状态维护 ─────────────────

        private void SetBusy(bool busy)
        {
            _busy = busy;
            BtnOpen.IsEnabled = !busy;
            BtnReset.IsEnabled = !busy && !string.IsNullOrEmpty(_sourcePath);
            UpdateUiState();
        }

        private bool _busy;

        private void UpdateUiState()
        {
            bool hasPages = _pages.Count > 0;
            EmptyOverlay.Visibility = hasPages ? Visibility.Collapsed : Visibility.Visible;

            bool idle = !_busy;
            BtnRotateLeft.IsEnabled = idle && hasPages;
            BtnRotateRight.IsEnabled = idle && hasPages;
            BtnMoveUp.IsEnabled = idle && hasPages;
            BtnMoveDown.IsEnabled = idle && hasPages;
            BtnDelete.IsEnabled = idle && hasPages;
            BtnExportAll.IsEnabled = idle && hasPages;
            BtnExportSelected.IsEnabled = idle && hasPages;
        }

        private void CleanupThumbnails()
        {
            try
            {
                if (!string.IsNullOrEmpty(_thumbDir) && Directory.Exists(_thumbDir))
                    Directory.Delete(_thumbDir, true);
            }
            catch { }
            _thumbDir = null;
        }

        [DllImport("psapi.dll")]
        private static extern bool EmptyWorkingSet(IntPtr hProcess);
    }
}
