using System;
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
    public partial class WordToPdfPage : UserControl
    {
        private readonly ObservableCollection<PdfFileItem> _files = new ObservableCollection<PdfFileItem>();
        public WordToPdfPage() { InitializeComponent(); FileListView.ItemsSource = _files; UpdateEmptyHint(); }
        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        { var d = new OpenFileDialog { Title = "选择要转换的Word文件", Filter = "Word 文档 (*.doc;*.docx)|*.doc;*.docx|所有文件 (*.*)|*.*", Multiselect = true }; if (d.ShowDialog() == true) AddFiles(d.FileNames); }
        private void FileListView_Drop(object sender, DragEventArgs e)
        { if (e.Data.GetDataPresent(DataFormats.FileDrop)) AddFiles((string[])e.Data.GetData(DataFormats.FileDrop)); }
        private void FileListView_PreviewDragOver(object sender, DragEventArgs e)
        { e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None; e.Handled = true; }
        private void BtnRemoveFile_Click(object sender, RoutedEventArgs e)
        { if (sender is FrameworkElement el && el.Tag is PdfFileItem item) { _files.Remove(item); UpdateEmptyHint(); } }
        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        { if (!(FileListView.View is GridView gv) || gv.Columns.Count < 3) return; double t = FileListView.ActualWidth - 14; double n = Math.Max(150, t * 0.30); double p = Math.Max(100, t - n - 70); if (Math.Abs(gv.Columns[0].Width - n) > 1) gv.Columns[0].Width = n; if (Math.Abs(gv.Columns[1].Width - p) > 1) gv.Columns[1].Width = p; }
        private async void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            if (_files.Count == 0) { MessageBox.Show("请先添加Word文件", "提示", MessageBoxButton.OK, MessageBoxImage.Information); return; }
            var invalid = _files.FirstOrDefault(f => !File.Exists(f.FilePath));
            if (invalid != null) { MessageBox.Show($"文件不存在: {invalid.FilePath}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); return; }
            BtnConvert.IsEnabled = false; BtnAddFiles.IsEnabled = false; TxtConversionStatus.Text = "准备中...";
            var snap = _files.ToList();
            try { await Task.Run(() => DoConversion(snap)); }
            catch (Exception ex) { TxtConversionStatus.Text = $"转换失败: {ex.Message}"; MessageBox.Show(ex.Message, "转换错误", MessageBoxButton.OK, MessageBoxImage.Error); }
            finally { BtnConvert.IsEnabled = true; BtnAddFiles.IsEnabled = true; ReleaseMemory(); }
        }
        private static void ReleaseMemory()
        { try { System.Runtime.GCSettings.LargeObjectHeapCompactionMode = System.Runtime.GCLargeObjectHeapCompactionMode.CompactOnce; GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); EmptyWorkingSet(Process.GetCurrentProcess().Handle); } catch { } }
        [DllImport("psapi.dll")] private static extern bool EmptyWorkingSet(IntPtr hProcess);
        private void DoConversion(System.Collections.Generic.List<PdfFileItem> files)
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Word转PDF输出");
            int t = files.Count; int c = 0; var fails = new System.Collections.Generic.List<string>();
            foreach (var f in files)
            { c++; int i = c; Dispatcher.Invoke(() => TxtConversionStatus.Text = $"正在转换 ({i}/{t}) {f.FileName}");
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                try { using (var conv = new WordToPdfConverterSpire()) conv.Convert(f.FilePath, dir); }
                catch (Exception ex) { fails.Add($"{f.FileName}: {ex.Message}"); }
            }
            Dispatcher.Invoke(() => {
                if (fails.Count > 0) { TxtConversionStatus.Text = $"完成，{fails.Count} 个文件转换失败"; MessageBox.Show($"以下文件转换失败:\n{string.Join("\n", fails)}", "部分文件失败", MessageBoxButton.OK, MessageBoxImage.Warning); }
                else { TxtConversionStatus.Text = "全部转换完成"; try { Process.Start(dir); } catch { } }
            });
        }
        private void AddFiles(System.Collections.Generic.IEnumerable<string> paths)
        { foreach (var p in paths) { var ext = Path.GetExtension(p); if (string.Equals(ext, ".doc", StringComparison.OrdinalIgnoreCase) || string.Equals(ext, ".docx", StringComparison.OrdinalIgnoreCase)) { if (!_files.Any(f => f.FilePath.Equals(p, StringComparison.OrdinalIgnoreCase))) _files.Add(new PdfFileItem { FileName = Path.GetFileNameWithoutExtension(p), FilePath = p }); } } UpdateEmptyHint(); }
        private void UpdateEmptyHint() { EmptyOverlay.Visibility = _files.Count == 0 ? Visibility.Visible : Visibility.Collapsed; }
    }
}