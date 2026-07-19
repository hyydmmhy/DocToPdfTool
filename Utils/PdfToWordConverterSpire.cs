using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Spire.Pdf;

namespace DocToPdfTool.Utils
{
    /// <summary>
    /// 基于 Spire.PDF 的 PDF 转 Word 转换器（免费版）。
    /// 转换后自动去除 Spire.PDF 免费版添加的评估水印。
    /// 无需 Microsoft Office。
    /// </summary>
    public class PdfToWordConverterSpire : IDisposable
    {
        private bool _disposed;

        /// <summary>转换单个 PDF 文件到 DOCX。</summary>
        public void Convert(string pdfPath, string outputDir)
        {
            if (string.IsNullOrEmpty(pdfPath)) throw new ArgumentNullException(nameof(pdfPath));
            if (!File.Exists(pdfPath)) throw new FileNotFoundException("PDF 文件不存在", pdfPath);

            var fileName = Path.GetFileNameWithoutExtension(pdfPath);
            var tempPath = Path.Combine(outputDir, fileName + "_temp.docx");
            var savePath = Path.Combine(outputDir, fileName + ".docx");

            try
            {
                // Step 1: 用 Spire.PDF 转换（免费版会加水印）
                using (var doc = new PdfDocument())
                {
                    doc.LoadFromFile(pdfPath);
                    doc.SaveToFile(tempPath, FileFormat.DOCX);
                }

                // Step 2: 用 OpenXML 去除水印
                RemoveWatermark(tempPath, savePath);
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); } catch { }
                }
            }
        }

        /// <summary>去除 Spire.PDF 评估水印。</summary>
        private void RemoveWatermark(string tempPath, string savePath)
        {
            // 复制一份，在原文件上修改
            File.Copy(tempPath, savePath, true);

            using (var wd = WordprocessingDocument.Open(savePath, true))
            {
                var body = wd.MainDocumentPart.Document.Body;
                var elements = body.Elements().ToList();

                bool changed = false;
                foreach (var elem in elements)
                {
                    if (elem.InnerText.Contains("Evaluation Warning"))
                    {
                        elem.Remove();
                        changed = true;
                    }
                }

                if (changed)
                {
                    wd.MainDocumentPart.Document.Save();
                }
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
            }
        }
    }
}