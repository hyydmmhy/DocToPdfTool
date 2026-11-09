using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using ZXing;
using ZXing.Common;
using ZXing.QrCode;
using ZXing.QrCode.Internal;

namespace DocToPdfTool.Pages
{
    public partial class QrCodePage : UserControl
    {
        private string _logoPath;
        private BitmapSource _generatedQr;

        public QrCodePage()
        {
            InitializeComponent();
        }

        private void BtnSelectLogo_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "图片文件|*.png;*.jpg;*.jpeg;*.bmp;*.gif",
                Title = "选择Logo图片"
            };
            if (dlg.ShowDialog() == true)
            {
                _logoPath = dlg.FileName;
                TxtLogoPath.Text = Path.GetFileName(_logoPath);
            }
        }

        private void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            var text = TxtQrContent.Text?.Trim();
            if (string.IsNullOrEmpty(text))
            {
                MessageBox.Show("请输入二维码内容", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(TxtQrSize.Text, out int size) || size < 50 || size > 2000)
            {
                MessageBox.Show("大小请输入 50~2000 之间的数字", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var levelText = ((ComboBoxItem)CbErrorLevel.SelectedItem).Content.ToString();
            var errorLevel = ErrorCorrectionLevel.M;
            switch (levelText)
            {
                case "L": errorLevel = ErrorCorrectionLevel.L; break;
                case "M": errorLevel = ErrorCorrectionLevel.M; break;
                case "Q": errorLevel = ErrorCorrectionLevel.Q; break;
                case "H": errorLevel = ErrorCorrectionLevel.H; break;
            }

            try
            {
                var writer = new BarcodeWriterPixelData
                {
                    Format = BarcodeFormat.QR_CODE,
                    Options = new QrCodeEncodingOptions
                    {
                        Width = size,
                        Height = size,
                        ErrorCorrection = errorLevel,
                        CharacterSet = "UTF-8",
                        Margin = 1
                    }
                };

                var pixelData = writer.Write(text);
                var bitmap = BitmapSource.Create(
                    pixelData.Width, pixelData.Height,
                    96, 96,
                    PixelFormats.Bgra32,
                    null,
                    pixelData.Pixels,
                    pixelData.Width * 4);

                if (ChkAddLogo.IsChecked == true && !string.IsNullOrEmpty(_logoPath) && File.Exists(_logoPath))
                {
                    bitmap = OverlayLogo(bitmap, _logoPath, size);
                }

                bitmap.Freeze();
                _generatedQr = bitmap;
                ImgQrPreview.Source = _generatedQr;
            }
            catch (Exception ex)
            {
                MessageBox.Show("生成失败: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private BitmapSource OverlayLogo(BitmapSource qrBitmap, string logoPath, int qrSize)
        {
            var logoImage = new BitmapImage();
            logoImage.BeginInit();
            logoImage.UriSource = new Uri(logoPath);
            logoImage.CacheOption = BitmapCacheOption.OnLoad;
            logoImage.EndInit();

            int logoSize = qrSize / 5;
            int x = (qrSize - logoSize) / 2;
            int y = (qrSize - logoSize) / 2;

            var visual = new DrawingVisual();
            using (var ctx = visual.RenderOpen())
            {
                ctx.DrawImage(qrBitmap, new Rect(0, 0, qrSize, qrSize));
                ctx.DrawRectangle(Brushes.White, null, new Rect(x - 2, y - 2, logoSize + 4, logoSize + 4));
                ctx.DrawImage(logoImage, new Rect(x, y, logoSize, logoSize));
            }

            var rtb = new RenderTargetBitmap(qrSize, qrSize, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(visual);
            return rtb;
        }

        private void BtnSaveQr_Click(object sender, RoutedEventArgs e)
        {
            if (_generatedQr == null)
            {
                MessageBox.Show("请先生成二维码", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "PNG 图片|*.png",
                FileName = "qrcode.png"
            };
            if (dlg.ShowDialog() == true)
            {
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(_generatedQr));
                using (var fs = new FileStream(dlg.FileName, FileMode.Create))
                {
                    encoder.Save(fs);
                }
                MessageBox.Show("保存成功", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnSelectImage_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "图片文件|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tiff;*.webp",
                Title = "选择要识别的图片"
            };
            if (dlg.ShowDialog() != true) return;

            TxtImagePath.Text = Path.GetFileName(dlg.FileName);

            try
            {
                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.UriSource = new Uri(dlg.FileName);
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                ImgDecodePreview.Source = bitmapImage;

                using (var bitmap = new System.Drawing.Bitmap(dlg.FileName))
                {
                    var reader = new BarcodeReader();
                    var result = reader.Decode(bitmap);
                    if (result != null)
                    {
                        TxtDecodeResult.Text = result.Text;
                    }
                    else
                    {
                        TxtDecodeResult.Text = "未识别到二维码";
                    }
                }
            }
            catch (Exception ex)
            {
                TxtDecodeResult.Text = "识别失败: " + ex.Message;
            }
        }

        private void BtnCopyDecodeResult_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(TxtDecodeResult.Text))
            {
                Clipboard.SetText(TxtDecodeResult.Text);
                MessageBox.Show("已复制到剪切板", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("没有可复制内容", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
