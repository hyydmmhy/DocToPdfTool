using System.Windows.Media;

namespace DocToPdfTool.Pages
{
    /// <summary>PDF 加密页面的列表项。</summary>
    public class PdfEncryptItem
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string PageText { get; set; }
        public string StatusText { get; set; }
        public bool IsEncrypted { get; set; }
        public ImageSource Icon { get; set; }
    }
}
