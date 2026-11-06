using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using DocToPdfTool.Utils;
using Microsoft.Win32;

namespace DocToPdfTool.Pages
{
    public partial class TranslatePage : UserControl
    {
        private CancellationTokenSource _cts;

        public TranslatePage()
        {
            InitializeComponent();
        }

        // ───────────────── 导入 ─────────────────

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "选择要翻译的文本文件",
                Filter = "文本文件 (*.txt;*.md)|*.txt;*.md|所有文件 (*.*)|*.*"
            };
            if (dialog.ShowDialog() != true) return;

            try
            {
                TxtSource.Text = ReadText(dialog.FileName);
                TxtStatus.Text = $"已导入：{Path.GetFileName(dialog.FileName)}（{TxtSource.Text.Length} 字符）";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "读取失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>读取文本：识别 BOM，无 BOM 时先按 UTF-8 严格解码，失败退回 GBK。</summary>
        private static string ReadText(string path)
        {
            var bytes = File.ReadAllBytes(path);

            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
                return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
                return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);

            try
            {
                return new UTF8Encoding(false, true).GetString(bytes);
            }
            catch
            {
                return Encoding.GetEncoding(936).GetString(bytes);
            }
        }

        // ───────────────── 翻译 ─────────────────

        private TranslateEngine GetEngine()
        {
            return CmbEngine.SelectedIndex == 1 ? TranslateEngine.Tencent : TranslateEngine.Microsoft;
        }

        private string GetTargetLanguage()
        {
            if (CmbTarget.SelectedItem is ComboBoxItem item && item.Tag is string tag) return tag;
            return "en";
        }

        private async void BtnTranslate_Click(object sender, RoutedEventArgs e)
        {
            var text = TxtSource.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text))
            {
                MessageBox.Show("请先输入或导入要翻译的内容", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var options = new TranslateOptions
            {
                Engine = GetEngine(),
                TargetLanguage = GetTargetLanguage(),
                SourceLanguage = "auto",
                AutoFallback = true
            };

            SetBusy(true);
            _cts = new CancellationTokenSource();

            var stopwatch = Stopwatch.StartNew();
            try
            {
                using (var service = new TranslateService())
                {
                    var result = await service.TranslateAsync(text, options,
                        (done, total) => Dispatcher.Invoke(() =>
                            TxtStatus.Text = $"正在翻译… {done}/{total} 段"),
                        _cts.Token);

                    stopwatch.Stop();
                    TxtResult.Text = result.Text;

                    var engineName = TranslateService.GetEngineDisplayName(result.UsedEngine);
                    var summary = $"完成：{result.SegmentCount} 段，{engineName}，{stopwatch.Elapsed.TotalSeconds:0.0} 秒";
                    if (!string.IsNullOrEmpty(result.Note)) summary += " ｜ " + result.Note;
                    TxtStatus.Text = summary;

                    AutoSave(result.Text);
                }
            }
            catch (OperationCanceledException)
            {
                TxtStatus.Text = "已取消";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "翻译失败：" + ex.Message;
                MessageBox.Show(ex.Message, "翻译失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetBusy(false);
                _cts?.Dispose();
                _cts = null;
            }
        }

        private void SetBusy(bool busy)
        {
            BtnTranslate.IsEnabled = !busy;
            BtnImport.IsEnabled = !busy;
            CmbTarget.IsEnabled = !busy;
            CmbEngine.IsEnabled = !busy;
            TxtStatus.Text = busy ? "正在翻译…" : TxtStatus.Text;
        }

        /// <summary>自动保存一份到功能输出目录，避免翻完丢失。</summary>
        private void AutoSave(string text)
        {
            try
            {
                var dir = AppConfig.GetFeatureOutputDir("文本翻译");
                var path = Path.Combine(dir, $"译文_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                File.WriteAllText(path, text, new UTF8Encoding(true));
            }
            catch
            {
                // 自动保存失败不影响主流程
            }
        }

        // ───────────────── 复制与保存 ─────────────────

        private void BtnCopy_Click(object sender, RoutedEventArgs e)
        {
            var text = TxtResult.Text;
            if (string.IsNullOrEmpty(text))
            {
                TxtStatus.Text = "没有可复制的译文";
                return;
            }

            try
            {
                Clipboard.SetText(text);
                TxtStatus.Text = "译文已复制到剪贴板";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "复制失败：" + ex.Message;
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            var text = TxtResult.Text;
            if (string.IsNullOrEmpty(text))
            {
                TxtStatus.Text = "没有可保存的译文";
                return;
            }

            var dialog = new SaveFileDialog
            {
                Title = "保存译文",
                Filter = "文本文件 (*.txt)|*.txt",
                FileName = $"译文_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
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
    }
}
