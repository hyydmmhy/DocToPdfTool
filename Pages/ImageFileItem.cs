using System.Windows.Media;

namespace DocToPdfTool.Pages
{
    /// <summary>图片转 PDF 页面的列表项。</summary>
    public class ImageFileItem
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string SizeText { get; set; }
        public ImageSource Thumbnail { get; set; }
    }
}
