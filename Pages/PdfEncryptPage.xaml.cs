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
using System.Windows.Threading;
using DocToPdfTool.Utils;
using Microsoft.Win32;
using Spire.Pdf;

namespace DocToPdfTool.Pages
{
    public partial class PdfEncryptPage : UserControl
    {
        private readonly ObservableCollection<PdfEncryptItem> _files = new ObservableCollection<PdfEncryptItem>();

        public PdfEncryptPage()
        {
            InitializeComponent();
            FileListView.ItemsSource = _files;

            // 眼睛按钮：密文 / 明文切换
            SetupPasswordToggle(PwdOpen, TxtOpen, BtnEyeOpen, WmOpen);
            SetupPasswordToggle(PwdOpenConfirm, TxtOpenConfirm, BtnEyeOpenConfirm, WmOpenConfirm);
            SetupPasswordToggle(PwdOwner, TxtOwner, BtnEyeOwner, WmOwner, PwdOwner_Changed);
            SetupPasswordToggle(PwdOwnerConfirm, TxtOwnerConfirm, BtnEyeOwnerConfirm, WmOwnerConfirm);
            SetupPasswordToggle(PwdDecrypt, TxtDecrypt, BtnEyeDecrypt, WmDecrypt);

            UpdateEmptyHint();
            UpdateModePanels();
            UpdateOwnerDependentControls();
        }

        // ───────────────── 文件列表 ─────────────────

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog
            {
                Title = "选择要处理的PDF文件",
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
            if (sender is FrameworkElement el && el.Tag is PdfEncryptItem item)
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

                bool encrypted = PdfEncryptConverter.IsEncrypted(p);
                int pages = encrypted ? 0 : PdfEncryptConverter.TryGetPageCount(p);

                _files.Add(new PdfEncryptItem
                {
                    FileName = Path.GetFileName(p),
                    FilePath = p,
                    PageText = pages > 0 ? pages + " 页" : "—",
                    IsEncrypted = encrypted,
                    StatusText = encrypted ? "已加密" : "未加密"
                });
            }
            UpdateEmptyHint();
        }

        private void FileListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!(FileListView.View is GridView gv) || gv.Columns.Count < 5) return;
            double total = FileListView.ActualWidth - 14;
            double fixedWidth = 70 + 90 + 60;
            double name = Math.Max(140, (total - fixedWidth) * 0.45);
            double path = Math.Max(100, total - fixedWidth - name);
            if (Math.Abs(gv.Columns[0].Width - name) > 1) gv.Columns[0].Width = name;
            if (Math.Abs(gv.Columns[3].Width - path) > 1) gv.Columns[3].Width = path;
        }

        // ───────────────── 模式与联动 ─────────────────

        private bool IsEncryptMode => RdoEncrypt.IsChecked == true;

        private void Mode_Changed(object sender, RoutedEventArgs e)
        {
            if (PanelEncrypt == null || PanelDecrypt == null) return;
            UpdateModePanels();
        }

        private void UpdateModePanels()
        {
            bool encrypt = IsEncryptMode;
            PanelEncrypt.Visibility = encrypt ? Visibility.Visible : Visibility.Collapsed;
            PanelDecrypt.Visibility = encrypt ? Visibility.Collapsed : Visibility.Visible;
        }

        /// <summary>密码长度约束。</summary>
        private const int MinPasswordLength = 6;
        private const int MaxPasswordLength = 128;

        /// <summary>
        /// 为「密码框 + 明文本框 + 眼睛按钮」建立联动：
        /// 点击眼睛在密文 PasswordBox 与明文 TextBox 之间切换，并同步内容。
        /// </summary>
        private void SetupPasswordToggle(PasswordBox pwd, TextBox plain, Button eye,
            TextBlock watermark, Action onChanged = null)
        {
            if (pwd == null || plain == null || eye == null) return;

            eye.Click += (s, e) =>
            {
                bool showPlain = plain.Visibility != Visibility.Visible;
                if (showPlain)
                {
                    plain.Text = pwd.Password ?? string.Empty;
                    plain.Visibility = Visibility.Visible;
                    pwd.Visibility = Visibility.Collapsed;
                    eye.Opacity = 0.45;

                    // 光标置于末尾；等布局完成后再设置一次，避免焦点/位置被重置到开头
                    plain.Focus();
                    plain.CaretIndex = plain.Text.Length;
                    plain.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        plain.Focus();
                        plain.CaretIndex = plain.Text.Length;
                        plain.Select(plain.Text.Length, 0);
                    }), DispatcherPriority.Input);
                }
                else
                {
                    var text = plain.Text ?? string.Empty;
                    bool changed = !string.Equals(pwd.Password, text, StringComparison.Ordinal);

                    pwd.Visibility = Visibility.Visible;
                    plain.Visibility = Visibility.Collapsed;
                    eye.Opacity = 1.0;

                    // 内容未变时不重新赋值，保留密码框原有光标位置
                    if (changed) pwd.Password = text;

                    pwd.Focus();
                    if (changed)
                    {
                        // 内容有变化（用户在明文态改过）时全选，便于继续输入或直接覆盖
                        pwd.SelectAll();
                    }
                }

                UpdateWatermark(pwd, plain, watermark);
                onChanged?.Invoke();
            };

            pwd.PasswordChanged += (s, e) => { UpdateWatermark(pwd, plain, watermark); onChanged?.Invoke(); };
            plain.TextChanged += (s, e) => { UpdateWatermark(pwd, plain, watermark); onChanged?.Invoke(); };
        }

        /// <summary>输入为空时显示占位提示（明文/密文任一模式都判断）。</summary>
        private static void UpdateWatermark(PasswordBox pwd, TextBox plain, TextBlock watermark)
        {
            if (watermark == null) return;
            bool empty = plain != null && plain.Visibility == Visibility.Visible
                ? string.IsNullOrEmpty(plain.Text)
                : string.IsNullOrEmpty(pwd?.Password);
            watermark.Visibility = empty ? Visibility.Visible : Visibility.Collapsed;
        }

        /// <summary>读取密码：明文模式下取 TextBox，密文模式下取 PasswordBox。</summary>
        private static string GetPassword(PasswordBox pwd, TextBox plain)
        {
            if (plain != null && plain.Visibility == Visibility.Visible)
                return plain.Text ?? string.Empty;
            return pwd?.Password ?? string.Empty;
        }

        /// <summary>未填写编辑密码时，算法与权限均不可选（该模式固定 RC4-128 且不限制权限）。</summary>
        private void PwdOwner_Changed()
        {
            UpdateOwnerDependentControls();
        }

        private void UpdateOwnerDependentControls()
        {
            if (CmbAlgorithm == null || CmbPreset == null || TxtHint == null) return;
            bool hasOwner = GetPassword(PwdOwner, TxtOwner).Length > 0;

            CmbAlgorithm.IsEnabled = hasOwner;
            CmbPreset.IsEnabled = hasOwner;

            CmbAlgorithm.Opacity = hasOwner ? 1.0 : 0.5;
            CmbPreset.Opacity = hasOwner ? 1.0 : 0.5;

            UpdateCustomPanel();
        }

        private void CmbPreset_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateCustomPanel();
        }

        /// <summary>
        /// 统一维护「自定义权限」可用状态与提示文案。
        /// 所有元素始终占位，仅切换可用性与文字，避免布局跳动。
        /// </summary>
        private void UpdateCustomPanel()
        {
            if (PanelCustom == null || CmbPreset == null) return;

            bool hasOwner = GetPassword(PwdOwner, TxtOwner).Length > 0;
            bool custom = hasOwner && CmbPreset.SelectedIndex == 3;
            PanelCustom.IsEnabled = custom;
            PanelCustom.Opacity = custom ? 1.0 : 0.45;

            if (TxtHint == null) return;

            if (!hasOwner)
            {
                TxtHint.Foreground = System.Windows.Media.Brushes.SandyBrown;
                TxtHint.Text = "未填写编辑密码：仅设置打开密码，不限制编辑与权限（使用 RC4-128 加密）";
            }
            else if (CmbPreset.SelectedIndex == 0)
            {
                TxtHint.Foreground = System.Windows.Media.Brushes.SlateGray;
                TxtHint.Text = "权限为「全部允许」：打开后可自由编辑/打印，阅读器不会索要编辑密码";
            }
            else
            {
                TxtHint.Foreground = System.Windows.Media.Brushes.SeaGreen;
                TxtHint.Text = "受限权限已启用：阅读器执行被禁止的操作时会要求输入编辑密码";
            }
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

            bool encryptMode = IsEncryptMode;
            PdfEncryptOptions options = null;
            string decryptPassword = null;

            if (encryptMode)
            {
                var pwd = GetPassword(PwdOpen, TxtOpen);
                var confirm = GetPassword(PwdOpenConfirm, TxtOpenConfirm);

                if (pwd.Length == 0)
                {
                    MessageBox.Show("请输入打开密码（6-128位，区分大小写）", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    PwdOpen.Focus();
                    return;
                }
                if (pwd.Length < MinPasswordLength || pwd.Length > MaxPasswordLength)
                {
                    MessageBox.Show($"打开密码长度需为 {MinPasswordLength}-{MaxPasswordLength} 位，当前 {pwd.Length} 位",
                        "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    PwdOpen.Focus();
                    return;
                }
                if (!string.Equals(pwd, confirm, StringComparison.Ordinal))
                {
                    MessageBox.Show("两次输入的打开密码不一致", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    PwdOpenConfirm.Focus();
                    return;
                }

                var owner = GetPassword(PwdOwner, TxtOwner);
                var ownerConfirm = GetPassword(PwdOwnerConfirm, TxtOwnerConfirm);

                if (owner.Length > 0)
                {
                    if (owner.Length < MinPasswordLength || owner.Length > MaxPasswordLength)
                    {
                        MessageBox.Show($"编辑密码长度需为 {MinPasswordLength}-{MaxPasswordLength} 位，当前 {owner.Length} 位",
                            "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                        PwdOwner.Focus();
                        return;
                    }
                    if (!string.Equals(owner, ownerConfirm, StringComparison.Ordinal))
                    {
                        MessageBox.Show("两次输入的编辑密码不一致", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                        PwdOwnerConfirm.Focus();
                        return;
                    }
                    if (string.Equals(owner, pwd, StringComparison.Ordinal))
                    {
                        MessageBox.Show("编辑密码不能与打开密码相同", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                        PwdOwner.Focus();
                        return;
                    }
                }

                options = new PdfEncryptOptions
                {
                    UserPassword = pwd,
                    OwnerPassword = owner,
                    Algorithm = GetSelectedAlgorithm(),
                    Preset = owner.Length > 0 ? GetSelectedPreset() : PdfEncryptPreset.AllowAll,
                    AllowPrint = ChkPrint.IsChecked == true,
                    AllowCopy = ChkCopy.IsChecked == true,
                    AllowModify = ChkModify.IsChecked == true,
                    AllowAnnotate = ChkAnnotate.IsChecked == true,
                    AllowFillForm = ChkForm.IsChecked == true
                };
            }
            else
            {
                decryptPassword = GetPassword(PwdDecrypt, TxtDecrypt);
            }

            var paths = _files.Select(f => f.FilePath).ToList();

            SetBusy(true);
            TxtStatus.Text = "准备中...";
            try
            {
                await Task.Run(() => DoWork(paths, encryptMode, options, decryptPassword));
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "处理失败: " + ex.Message;
                MessageBox.Show(ex.Message, "处理错误", MessageBoxButton.OK, MessageBoxImage.Error);
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
            RdoEncrypt.IsEnabled = !busy;
            RdoDecrypt.IsEnabled = !busy;
            PwdOpen.IsEnabled = !busy;
            PwdOpenConfirm.IsEnabled = !busy;
            PwdOwner.IsEnabled = !busy;
            PwdOwnerConfirm.IsEnabled = !busy;
            PwdDecrypt.IsEnabled = !busy;

            TxtOpen.IsEnabled = !busy;
            TxtOpenConfirm.IsEnabled = !busy;
            TxtOwner.IsEnabled = !busy;
            TxtOwnerConfirm.IsEnabled = !busy;
            TxtDecrypt.IsEnabled = !busy;

            BtnEyeOpen.IsEnabled = !busy;
            BtnEyeOpenConfirm.IsEnabled = !busy;
            BtnEyeOwner.IsEnabled = !busy;
            BtnEyeOwnerConfirm.IsEnabled = !busy;
            BtnEyeDecrypt.IsEnabled = !busy;
            if (busy)
            {
                CmbAlgorithm.IsEnabled = false;
                CmbPreset.IsEnabled = false;
            }
            else
            {
                UpdateOwnerDependentControls();
            }
        }

        private PdfEncryptionAlgorithm GetSelectedAlgorithm()
        {
            switch (CmbAlgorithm.SelectedIndex)
            {
                case 1: return PdfEncryptionAlgorithm.AES_128;
                case 2: return PdfEncryptionAlgorithm.RC4_128;
                default: return PdfEncryptionAlgorithm.AES_256;
            }
        }

        private PdfEncryptPreset GetSelectedPreset()
        {
            switch (CmbPreset.SelectedIndex)
            {
                case 1: return PdfEncryptPreset.ViewOnly;
                case 2: return PdfEncryptPreset.AllowPrint;
                case 3: return PdfEncryptPreset.Custom;
                default: return PdfEncryptPreset.AllowAll;
            }
        }

        private void DoWork(List<string> paths, bool encryptMode, PdfEncryptOptions options, string decryptPassword)
        {
            string dir = AppConfig.GetFeatureOutputDir(encryptMode ? "PDF加密" : "PDF解密");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            int total = paths.Count;
            int success = 0;
            var failures = new List<string>();

            for (int i = 0; i < total; i++)
            {
                int index = i;
                var path = paths[i];
                Dispatcher.Invoke(() => TxtStatus.Text = $"正在处理 ({index + 1}/{total}) {Path.GetFileName(path)}");

                try
                {
                    using (var converter = new PdfEncryptConverter())
                    {
                        if (encryptMode)
                            converter.Encrypt(path, dir, options);
                        else
                            converter.Decrypt(path, dir, decryptPassword);
                        success++;
                    }
                }
                catch (Exception ex)
                {
                    failures.Add($"{Path.GetFileName(path)}: {Shorten(ex.Message)}");
                }
            }

            Dispatcher.Invoke(() =>
            {
                if (failures.Count > 0)
                {
                    TxtStatus.Text = $"完成，成功 {success} 个，失败 {failures.Count} 个";
                    MessageBox.Show(
                        $"成功 {success} 个，输出目录:\n{dir}\n\n失败 {failures.Count} 个:\n{FormatFailures(failures)}",
                        "部分文件失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    TxtStatus.Text = $"全部完成，共 {success} 个文件";
                }
                try { Process.Start(dir); } catch { }
            });
        }

        private static string Shorten(string s)
        {
            if (string.IsNullOrEmpty(s)) return "未知错误";
            return s.Length > 80 ? s.Substring(0, 80) + "…" : s;
        }

        private static string FormatFailures(List<string> failures)
        {
            return string.Join("\n", failures.Take(8))
                 + (failures.Count > 8 ? $"\n… 共 {failures.Count} 个" : "");
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
