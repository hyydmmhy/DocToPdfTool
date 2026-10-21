using System.Windows;
using DocToPdfTool.Utils;

namespace DocToPdfTool.Pages
{
    public partial class SettingsWindow : Window
    {
        public SettingsWindow()
        {
            InitializeComponent();
            TxtSaveDir.Text = AppConfig.GetDefaultSaveDir();
        }

        private void BtnBrowseDir_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "选择默认保存目录";
                dialog.SelectedPath = TxtSaveDir.Text;

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    AppConfig.SetDefaultSaveDir(dialog.SelectedPath);
                    TxtSaveDir.Text = dialog.SelectedPath;
                    TxtStatus.Text = "已保存，之后的转换将使用新目录";
                }
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}