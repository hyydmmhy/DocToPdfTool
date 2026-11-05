using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace DocToPdfTool.Pages
{
    /// <summary>PDF 压缩页面的列表项。</summary>
    public class PdfCompressItem : INotifyPropertyChanged
    {
        private string _originalText = "—";
        private string _newText = "—";
        private string _ratioText = "—";
        private string _statusText = "待处理";

        public string FileName { get; set; }
        public string FilePath { get; set; }

        public string OriginalText
        {
            get => _originalText;
            set { if (_originalText != value) { _originalText = value; OnPropertyChanged(); } }
        }

        public string NewText
        {
            get => _newText;
            set { if (_newText != value) { _newText = value; OnPropertyChanged(); } }
        }

        public string RatioText
        {
            get => _ratioText;
            set { if (_ratioText != value) { _ratioText = value; OnPropertyChanged(); } }
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
}
