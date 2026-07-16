using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using UglyToad.PdfPig;

namespace DocToPdfTool.Utils
{
    public class PdfToPptConverter : IDisposable
    {
        private bool _disposed;

        // 渲染质量
        private const int RenderDpi = 200;
        private const double RenderScale = 1.0;

        // 输出图片格式（JPEG 无透明通道）
        private static readonly ImageFormat OutputImageFormat = ImageFormat.Jpeg;
        private const string ImageExt = "jpg";
        private const string ImageContentType = "image/jpeg";

        // XML 命名空间
        private static readonly XNamespace P =
            "http://schemas.openxmlformats.org/presentationml/2006/main";
        private static readonly XNamespace A =
            "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace R =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace Ct =
            "http://schemas.openxmlformats.org/package/2006/content-types";
        private static readonly XNamespace RelPkg =
            "http://schemas.openxmlformats.org/package/2006/relationships";

        public void Convert(string pdfPath, string outputDir, Action onPageComplete = null)
        {
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException("PDF文件不存在", pdfPath);

            string fileName = Path.GetFileNameWithoutExtension(pdfPath);
            string outputPath = Path.Combine(outputDir, $"{fileName}.pptx");

            if (File.Exists(outputPath))
            {
                int counter = 1;
                do { outputPath = Path.Combine(outputDir, $"{fileName}({counter}).pptx"); counter++; }
                while (File.Exists(outputPath));
            }

            // 获取页数
            int pageCount;
            using (var pdf = PdfDocument.Open(pdfPath))
            {
                pageCount = pdf.NumberOfPages;
            }

            // 临时目录存放渲染的 JPEG
            string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);

            try
            {
                // 一次性渲染所有页面（只开一次 WinRT，大幅降低内存）
                var imageFiles = RenderAllPages(pdfPath, tempDir, pageCount, onPageComplete);

                // 从第一页图片获取原始尺寸，计算幻灯片尺寸（保持原 PDF 比例）
                long slideW, slideH;
                using (var img = Image.FromFile(imageFiles[0]))
                {
                    slideW = (long)img.Width * 914400L / RenderDpi;
                    slideH = (long)img.Height * 914400L / RenderDpi;
                }

                // 创建 PPTX
                using (var archive = ZipFile.Open(outputPath, ZipArchiveMode.Create))
                {
                    WriteContentTypes(archive, pageCount);
                    WriteRels(archive);
                    WritePresProps(archive);
                    WriteViewProps(archive);
                    WriteTableStyles(archive);
                    WriteDocPropsCore(archive, fileName);
                    WriteDocPropsApp(archive, pageCount);
                    WritePresentation(archive, pageCount, slideW, slideH);
                    WritePresentationRels(archive, pageCount);
                    WriteSlideMaster(archive);
                    WriteSlideMasterRels(archive);
                    WriteSlideLayout(archive);
                    WriteSlideLayoutRels(archive);
                    WriteTheme(archive);

                    for (int i = 0; i < pageCount; i++)
                    {
                        int slideNum = i + 1;
                        WriteSlide(archive, slideNum, slideW, slideH);
                        WriteSlideRels(archive, slideNum);
                        WriteMediaImage(archive, slideNum, imageFiles[i]);
                        try { File.Delete(imageFiles[i]); } catch { }
                    }
                }
            }
            finally
            {
                try { Directory.Delete(tempDir, true); } catch { }
            }
        }

        // 一次性渲染所有页面（单次 WinRT 调用）
        private List<string> RenderAllPages(string pdfPath, string outputDir, int pageCount, Action onPageComplete = null)
        {
            List<string> images;
            using (var converter = new PdfToImageConverter
            {
                Dpi = RenderDpi,
                Scale = RenderScale,
                OutputFormat = OutputImageFormat
            })
            {
                images = converter.Convert(pdfPath, outputDir, 1, pageCount, onPageComplete);
                if (images == null || images.Count == 0)
                    throw new InvalidOperationException("PDF渲染失败，未能生成任何图片");
            }

            return images;
        }

        // ========== 辅助：写入 XML 条目（无 BOM） ==========

        private static void WriteXmlEntry(ZipArchive archive, string entryName, XElement element)
        {
            var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using (var sw = new StreamWriter(entry.Open(), new UTF8Encoding(false)))
            {
                element.Save(sw);
            }
        }

        private static void WriteStringEntry(ZipArchive archive, string entryName, string content)
        {
            var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using (var sw = new StreamWriter(entry.Open(), new UTF8Encoding(false)))
            {
                sw.Write(content);
            }
        }

        // ========== [Content_Types].xml ==========

        private void WriteContentTypes(ZipArchive archive, int slideCount)
        {
            var types = new XElement(Ct + "Types",
                new XElement(Ct + "Default", new XAttribute("Extension", "rels"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                new XElement(Ct + "Default", new XAttribute("Extension", ImageExt),
                    new XAttribute("ContentType", ImageContentType)),
                new XElement(Ct + "Override", new XAttribute("PartName", "/ppt/presentation.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")),
                new XElement(Ct + "Override", new XAttribute("PartName", "/ppt/slideMasters/slideMaster1.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml")),
                new XElement(Ct + "Override", new XAttribute("PartName", "/ppt/slideLayouts/slideLayout1.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml")),
                new XElement(Ct + "Override", new XAttribute("PartName", "/ppt/theme/theme1.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.theme+xml")),
                new XElement(Ct + "Override", new XAttribute("PartName", "/ppt/presProps.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml")),
                new XElement(Ct + "Override", new XAttribute("PartName", "/ppt/viewProps.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml")),
                new XElement(Ct + "Override", new XAttribute("PartName", "/ppt/tableStyles.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml")),
                new XElement(Ct + "Override", new XAttribute("PartName", "/docProps/core.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-package.core-properties+xml")),
                new XElement(Ct + "Override", new XAttribute("PartName", "/docProps/app.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml")));

            for (int i = 1; i <= slideCount; i++)
            {
                types.Add(new XElement(Ct + "Override",
                    new XAttribute("PartName", $"/ppt/slides/slide{i}.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")));
            }

            WriteXmlEntry(archive, "[Content_Types].xml", types);
        }

        // ========== _rels/.rels ==========

        private void WriteRels(ZipArchive archive)
        {
            WriteXmlEntry(archive, "_rels/.rels",
                new XElement(RelPkg + "Relationships",
                    new XElement(RelPkg + "Relationship",
                        new XAttribute("Id", "rId1"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
                        new XAttribute("Target", "ppt/presentation.xml")),
                    new XElement(RelPkg + "Relationship",
                        new XAttribute("Id", "rId2"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"),
                        new XAttribute("Target", "docProps/core.xml")),
                    new XElement(RelPkg + "Relationship",
                        new XAttribute("Id", "rId3"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"),
                        new XAttribute("Target", "docProps/app.xml"))));
        }

        // ========== ppt/presentation.xml ==========

        private void WritePresentation(ZipArchive archive, int slideCount, long slideW, long slideH)
        {
            var sldIdLst = new XElement(P + "sldIdLst");
            for (int i = 1; i <= slideCount; i++)
            {
                sldIdLst.Add(new XElement(P + "sldId",
                    new XAttribute("id", (uint)(256 + i)),
                    new XAttribute(R + "id", $"rId{i}")));
            }

            WriteXmlEntry(archive, "ppt/presentation.xml",
                new XElement(P + "presentation",
                    new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "p", P.NamespaceName),
                    new XElement(P + "sldMasterIdLst",
                        new XElement(P + "sldMasterId",
                            new XAttribute("id", 2147483648),
                            new XAttribute(R + "id", "rIdMaster"))),
                    sldIdLst,
                    new XElement(P + "sldSz",
                        new XAttribute("cx", slideW),
                        new XAttribute("cy", slideH)),
                    new XElement(P + "notesSz",
                        new XAttribute("cx", 6858000),
                        new XAttribute("cy", 9144000)),
                    DefaultTextStyle()));
        }

        private static XElement DefaultTextStyle()
        {
            XElement LevelDefRPr()
            {
                return new XElement(A + "defRPr",
                    new XAttribute("sz", 1800),
                    new XAttribute("kern", 1200),
                    new XElement(A + "solidFill",
                        new XElement(A + "schemeClr", new XAttribute("val", "tx1"))),
                    new XElement(A + "latin", new XAttribute("typeface", "+mn-lt")),
                    new XElement(A + "ea", new XAttribute("typeface", "+mn-ea")),
                    new XElement(A + "cs", new XAttribute("typeface", "+mn-cs")));
            }

            var lvls = new XElement[9];
            for (int i = 0; i < 9; i++)
            {
                int marL = i * 457200;
                lvls[i] = new XElement(A + $"lvl{i + 1}pPr",
                    new XAttribute("marL", marL),
                    new XAttribute("algn", "l"),
                    new XAttribute("defTabSz", 914400),
                    new XAttribute("rtl", 0),
                    new XAttribute("eaLnBrk", 1),
                    new XAttribute("latinLnBrk", 0),
                    new XAttribute("hangingPunct", 1),
                    LevelDefRPr());
            }

            return new XElement(P + "defaultTextStyle",
                new XElement(A + "defPPr",
                    new XElement(A + "defRPr", new XAttribute("lang", "zh-CN"))),
                lvls);
        }

        // ========== ppt/_rels/presentation.xml.rels ==========

        private void WritePresentationRels(ZipArchive archive, int slideCount)
        {
            var rels = new XElement(RelPkg + "Relationships",
                new XElement(RelPkg + "Relationship",
                    new XAttribute("Id", "rIdMaster"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"),
                    new XAttribute("Target", "slideMasters/slideMaster1.xml")));

            for (int i = 1; i <= slideCount; i++)
            {
                rels.Add(new XElement(RelPkg + "Relationship",
                    new XAttribute("Id", $"rId{i}"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"),
                    new XAttribute("Target", $"slides/slide{i}.xml")));
            }

            rels.Add(new XElement(RelPkg + "Relationship",
                    new XAttribute("Id", "rIdPresProps"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"),
                    new XAttribute("Target", "presProps.xml")));
            rels.Add(new XElement(RelPkg + "Relationship",
                    new XAttribute("Id", "rIdTableStyles"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"),
                    new XAttribute("Target", "tableStyles.xml")));
            rels.Add(new XElement(RelPkg + "Relationship",
                    new XAttribute("Id", "rIdViewProps"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"),
                    new XAttribute("Target", "viewProps.xml")));

            WriteXmlEntry(archive, "ppt/_rels/presentation.xml.rels", rels);
        }

        // ========== ppt/presProps.xml ==========

        private void WritePresProps(ZipArchive archive)
        {
            WriteXmlEntry(archive, "ppt/presProps.xml",
                new XElement(P + "presentationPr",
                    new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "p", P.NamespaceName)));
        }

        // ========== ppt/viewProps.xml ==========

        private void WriteViewProps(ZipArchive archive)
        {
            WriteXmlEntry(archive, "ppt/viewProps.xml",
                new XElement(P + "viewPr",
                    new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "p", P.NamespaceName),
                    new XElement(P + "normalViewPr",
                        new XElement(P + "restoredLeft", new XAttribute("sz", "15620")),
                        new XElement(P + "restoredTop", new XAttribute("sz", "94660"))),
                    new XElement(P + "slideViewPr",
                        new XElement(P + "cSldViewPr",
                            new XElement(P + "cViewPr", new XAttribute("varScale", "1"),
                                new XElement(P + "scale",
                                    new XElement(A + "sx", new XAttribute("n", "64"), new XAttribute("d", "100")),
                                    new XElement(A + "sy", new XAttribute("n", "64"), new XAttribute("d", "100"))),
                                new XElement(P + "origin", new XAttribute("x", "-1392"), new XAttribute("y", "-96"))))),
                    new XElement(P + "notesTextViewPr",
                        new XElement(P + "cViewPr",
                            new XElement(P + "scale",
                                new XElement(A + "sx", new XAttribute("n", "1"), new XAttribute("d", "1")),
                                new XElement(A + "sy", new XAttribute("n", "1"), new XAttribute("d", "1"))),
                            new XElement(P + "origin", new XAttribute("x", "0"), new XAttribute("y", "0")))),
                    new XElement(P + "gridSpacing", new XAttribute("cx", "76200"), new XAttribute("cy", "76200"))));
        }

        // ========== ppt/tableStyles.xml ==========

        private void WriteTableStyles(ZipArchive archive)
        {
            WriteXmlEntry(archive, "ppt/tableStyles.xml",
                new XElement(A + "tblStyleLst",
                    new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
                    new XAttribute("def", "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")));
        }

        // ========== docProps/core.xml ==========

        private void WriteDocPropsCore(ZipArchive archive, string fileName)
        {
            WriteStringEntry(archive, "docProps/core.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"" +
                " xmlns:dc=\"http://purl.org/dc/elements/1.1/\"" +
                " xmlns:dcterms=\"http://purl.org/dc/terms/\"" +
                " xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
                $"<dc:title>{EscapeXml(fileName)}</dc:title>" +
                "<dc:creator>DocToPdfTool</dc:creator>" +
                $"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{DateTime.UtcNow:yyyy-MM-ddTHH:mm:ssZ}</dcterms:created>" +
                "</cp:coreProperties>");
        }

        // ========== docProps/app.xml ==========

        private void WriteDocPropsApp(ZipArchive archive, int slideCount)
        {
            WriteStringEntry(archive, "docProps/app.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"" +
                " xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
                "<Application>DocToPdfTool</Application>" +
                "<TotalTime>0</TotalTime>" +
                $"<Slides>{slideCount}</Slides>" +
                "</Properties>");
        }

        private static string EscapeXml(string s)
        {
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        }

        // ========== ppt/slideMasters/slideMaster1.xml ==========

        private void WriteSlideMaster(ZipArchive archive)
        {
            WriteXmlEntry(archive, "ppt/slideMasters/slideMaster1.xml",
                new XElement(P + "sldMaster",
                    new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "p", P.NamespaceName),
                    new XElement(P + "cSld",
                        new XElement(P + "spTree",
                            new XElement(P + "nvGrpSpPr",
                                new XElement(P + "cNvPr", new XAttribute("id", "1"), new XAttribute("name", "")),
                                new XElement(P + "cNvGrpSpPr"),
                                new XElement(P + "nvPr")),
                            new XElement(P + "grpSpPr",
                                new XElement(A + "xfrm",
                                    new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                                    new XElement(A + "ext", new XAttribute("cx", "0"), new XAttribute("cy", "0")),
                                    new XElement(A + "chOff", new XAttribute("x", "0"), new XAttribute("y", "0")),
                                    new XElement(A + "chExt", new XAttribute("cx", "0"), new XAttribute("cy", "0")))))),
                    new XElement(P + "clrMap",
                        new XAttribute("bg1", "lt1"), new XAttribute("tx1", "dk1"),
                        new XAttribute("bg2", "lt2"), new XAttribute("tx2", "dk2"),
                        new XAttribute("accent1", "accent1"), new XAttribute("accent2", "accent2"),
                        new XAttribute("accent3", "accent3"), new XAttribute("accent4", "accent4"),
                        new XAttribute("accent5", "accent5"), new XAttribute("accent6", "accent6"),
                        new XAttribute("hlink", "hlink"), new XAttribute("folHlink", "folHlink")),
                    new XElement(P + "sldLayoutIdLst",
                        new XElement(P + "sldLayoutId",
                            new XAttribute("id", "2147483649"),
                            new XAttribute(R + "id", "rId1"))),
                    new XElement(P + "txStyles",
                        new XElement(P + "titleStyle",
                            new XElement(A + "lvl1pPr",
                                new XAttribute("algn", "l"),
                                new XAttribute("defTabSz", 914400),
                                new XAttribute("rtl", 0),
                                new XAttribute("eaLnBrk", 1),
                                new XAttribute("latinLnBrk", 0),
                                new XAttribute("hangingPunct", 1),
                                new XElement(A + "lnSpc",
                                    new XElement(A + "spcPct", new XAttribute("val", 90000))),
                                new XElement(A + "spcBef",
                                    new XElement(A + "spcPct", new XAttribute("val", 0))),
                                new XElement(A + "buNone"),
                                new XElement(A + "defRPr",
                                    new XAttribute("sz", 4400),
                                    new XAttribute("kern", 1200),
                                    new XElement(A + "solidFill",
                                        new XElement(A + "schemeClr", new XAttribute("val", "tx1"))),
                                    new XElement(A + "latin", new XAttribute("typeface", "+mj-lt")),
                                    new XElement(A + "ea", new XAttribute("typeface", "+mj-ea")),
                                    new XElement(A + "cs", new XAttribute("typeface", "+mj-cs"))))),
                        new XElement(P + "bodyStyle",
                            new XElement(A + "lvl1pPr",
                                new XAttribute("marL", 228600),
                                new XAttribute("indent", -228600),
                                new XAttribute("algn", "l"),
                                new XAttribute("defTabSz", 914400),
                                new XAttribute("rtl", 0),
                                new XAttribute("eaLnBrk", 1),
                                new XAttribute("latinLnBrk", 0),
                                new XAttribute("hangingPunct", 1),
                                new XElement(A + "lnSpc",
                                    new XElement(A + "spcPct", new XAttribute("val", 90000))),
                                new XElement(A + "spcBef",
                                    new XElement(A + "spcPts", new XAttribute("val", 1000))),
                                new XElement(A + "buFont",
                                    new XAttribute("typeface", "Arial")),
                                new XElement(A + "buChar",
                                    new XAttribute("char", "•")),
                                new XElement(A + "defRPr",
                                    new XAttribute("sz", 2800),
                                    new XAttribute("kern", 1200),
                                    new XElement(A + "solidFill",
                                        new XElement(A + "schemeClr", new XAttribute("val", "tx1"))),
                                    new XElement(A + "latin", new XAttribute("typeface", "+mn-lt")),
                                    new XElement(A + "ea", new XAttribute("typeface", "+mn-ea")),
                                    new XElement(A + "cs", new XAttribute("typeface", "+mn-cs"))))),
                        new XElement(P + "otherStyle",
                            new XElement(A + "defPPr",
                                new XElement(A + "defRPr", new XAttribute("lang", "zh-CN"))),
                            new XElement(A + "lvl1pPr",
                                new XAttribute("marL", 0),
                                new XAttribute("algn", "l"),
                                new XAttribute("defTabSz", 914400),
                                new XAttribute("rtl", 0),
                                new XAttribute("eaLnBrk", 1),
                                new XAttribute("latinLnBrk", 0),
                                new XAttribute("hangingPunct", 1),
                                new XElement(A + "defRPr",
                                    new XAttribute("sz", 1800),
                                    new XAttribute("kern", 1200),
                                    new XElement(A + "solidFill",
                                        new XElement(A + "schemeClr", new XAttribute("val", "tx1"))),
                                    new XElement(A + "latin", new XAttribute("typeface", "+mn-lt")),
                                    new XElement(A + "ea", new XAttribute("typeface", "+mn-ea")),
                                    new XElement(A + "cs", new XAttribute("typeface", "+mn-cs"))))))));
        }

        // ========== ppt/slideMasters/_rels/slideMaster1.xml.rels ==========

        private void WriteSlideMasterRels(ZipArchive archive)
        {
            WriteXmlEntry(archive, "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                new XElement(RelPkg + "Relationships",
                    new XElement(RelPkg + "Relationship",
                        new XAttribute("Id", "rId1"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"),
                        new XAttribute("Target", "../slideLayouts/slideLayout1.xml")),
                    new XElement(RelPkg + "Relationship",
                        new XAttribute("Id", "rId2"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"),
                        new XAttribute("Target", "../theme/theme1.xml"))));
        }

        // ========== ppt/slideLayouts/slideLayout1.xml ==========

        private void WriteSlideLayout(ZipArchive archive)
        {
            WriteXmlEntry(archive, "ppt/slideLayouts/slideLayout1.xml",
                new XElement(P + "sldLayout",
                    new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "p", P.NamespaceName),
                    new XAttribute("type", "blank"),
                    new XAttribute("preserve", "1"),
                    new XElement(P + "cSld",
                        new XAttribute("name", "Blank"),
                        new XElement(P + "spTree",
                            new XElement(P + "nvGrpSpPr",
                                new XElement(P + "cNvPr", new XAttribute("id", "1"), new XAttribute("name", "")),
                                new XElement(P + "cNvGrpSpPr"),
                                new XElement(P + "nvPr")),
                            new XElement(P + "grpSpPr",
                                new XElement(A + "xfrm",
                                    new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                                    new XElement(A + "ext", new XAttribute("cx", "0"), new XAttribute("cy", "0")),
                                    new XElement(A + "chOff", new XAttribute("x", "0"), new XAttribute("y", "0")),
                                    new XElement(A + "chExt", new XAttribute("cx", "0"), new XAttribute("cy", "0")))))),
                    new XElement(P + "clrMapOvr",
                        new XElement(A + "masterClrMapping"))));
        }

        // ========== ppt/slideLayouts/_rels/slideLayout1.xml.rels ==========

        private void WriteSlideLayoutRels(ZipArchive archive)
        {
            WriteXmlEntry(archive, "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                new XElement(RelPkg + "Relationships",
                    new XElement(RelPkg + "Relationship",
                        new XAttribute("Id", "rId1"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"),
                        new XAttribute("Target", "../theme/theme1.xml")),
                    new XElement(RelPkg + "Relationship",
                        new XAttribute("Id", "rId2"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"),
                        new XAttribute("Target", "../slideMasters/slideMaster1.xml"))));
        }

        // ========== ppt/theme/theme1.xml ==========

        private void WriteTheme(ZipArchive archive)
        {
            WriteXmlEntry(archive, "ppt/theme/theme1.xml",
                new XElement(A + "theme",
                    new XAttribute("name", "Default"),
                    new XElement(A + "themeElements",
                        new XElement(A + "clrScheme", new XAttribute("name", "Default"),
                            new XElement(A + "dk1", new XElement(A + "srgbClr", new XAttribute("val", "000000"))),
                            new XElement(A + "lt1", new XElement(A + "srgbClr", new XAttribute("val", "FFFFFF"))),
                            new XElement(A + "dk2", new XElement(A + "srgbClr", new XAttribute("val", "44546A"))),
                            new XElement(A + "lt2", new XElement(A + "srgbClr", new XAttribute("val", "E7E6E6"))),
                            new XElement(A + "accent1", new XElement(A + "srgbClr", new XAttribute("val", "4472C4"))),
                            new XElement(A + "accent2", new XElement(A + "srgbClr", new XAttribute("val", "ED7D31"))),
                            new XElement(A + "accent3", new XElement(A + "srgbClr", new XAttribute("val", "A5A5A5"))),
                            new XElement(A + "accent4", new XElement(A + "srgbClr", new XAttribute("val", "FFC000"))),
                            new XElement(A + "accent5", new XElement(A + "srgbClr", new XAttribute("val", "5B9BD5"))),
                            new XElement(A + "accent6", new XElement(A + "srgbClr", new XAttribute("val", "70AD47"))),
                            new XElement(A + "hlink", new XElement(A + "srgbClr", new XAttribute("val", "0563C1"))),
                            new XElement(A + "folHlink", new XElement(A + "srgbClr", new XAttribute("val", "954F72")))),
                        new XElement(A + "fontScheme", new XAttribute("name", "Default"),
                            new XElement(A + "majorFont",
                                new XElement(A + "latin", new XAttribute("typeface", "Calibri Light")),
                                new XElement(A + "ea", new XAttribute("typeface", "")),
                                new XElement(A + "cs", new XAttribute("typeface", ""))),
                            new XElement(A + "minorFont",
                                new XElement(A + "latin", new XAttribute("typeface", "Calibri")),
                                new XElement(A + "ea", new XAttribute("typeface", "")),
                                new XElement(A + "cs", new XAttribute("typeface", "")))),
                        new XElement(A + "fmtScheme", new XAttribute("name", "Default"),
                            new XElement(A + "fillStyleLst",
                                new XElement(A + "solidFill", new XElement(A + "schemeClr", new XAttribute("val", "phClr"))),
                                new XElement(A + "solidFill", new XElement(A + "schemeClr", new XAttribute("val", "phClr"))),
                                new XElement(A + "solidFill", new XElement(A + "schemeClr", new XAttribute("val", "phClr")))),
                            new XElement(A + "lnStyleLst",
                                new XElement(A + "ln", new XAttribute("w", "6350"),
                                    new XElement(A + "solidFill", new XElement(A + "schemeClr", new XAttribute("val", "phClr")))),
                                new XElement(A + "ln", new XAttribute("w", "6350"),
                                    new XElement(A + "solidFill", new XElement(A + "schemeClr", new XAttribute("val", "phClr")))),
                                new XElement(A + "ln", new XAttribute("w", "6350"),
                                    new XElement(A + "solidFill", new XElement(A + "schemeClr", new XAttribute("val", "phClr"))))),
                            new XElement(A + "effectStyleLst",
                                new XElement(A + "effectStyle", new XElement(A + "effectLst")),
                                new XElement(A + "effectStyle", new XElement(A + "effectLst")),
                                new XElement(A + "effectStyle", new XElement(A + "effectLst"))),
                            new XElement(A + "bgFillStyleLst",
                                new XElement(A + "solidFill", new XElement(A + "schemeClr", new XAttribute("val", "phClr"))),
                                new XElement(A + "solidFill", new XElement(A + "schemeClr", new XAttribute("val", "phClr"))),
                                new XElement(A + "solidFill", new XElement(A + "schemeClr", new XAttribute("val", "phClr"))))))));
        }

        // ========== ppt/slides/slideN.xml ==========

        private void WriteSlide(ZipArchive archive, int slideNum, long slideW, long slideH)
        {
            WriteXmlEntry(archive, $"ppt/slides/slide{slideNum}.xml",
                new XElement(P + "sld",
                    new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "p", P.NamespaceName),
                    new XElement(P + "cSld",
                        new XElement(P + "bg",
                            new XElement(P + "bgPr",
                                new XElement(A + "solidFill",
                                    new XElement(A + "srgbClr", new XAttribute("val", "FFFFFF"))))),
                        new XElement(P + "spTree",
                            new XElement(P + "nvGrpSpPr",
                                new XElement(P + "cNvPr", new XAttribute("id", "1"), new XAttribute("name", "")),
                                new XElement(P + "cNvGrpSpPr"),
                                new XElement(P + "nvPr")),
                            new XElement(P + "grpSpPr",
                                new XElement(A + "xfrm",
                                    new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                                    new XElement(A + "ext", new XAttribute("cx", "0"), new XAttribute("cy", "0")),
                                    new XElement(A + "chOff", new XAttribute("x", "0"), new XAttribute("y", "0")),
                                    new XElement(A + "chExt", new XAttribute("cx", "0"), new XAttribute("cy", "0")))),
                            new XElement(P + "pic",
                                new XElement(P + "nvPicPr",
                                    new XElement(P + "cNvPr",
                                        new XAttribute("id", "2"),
                                        new XAttribute("name", $"page_{slideNum:D3}")),
                                    new XElement(P + "cNvPicPr",
                                        new XElement(A + "picLocks", new XAttribute("noChangeAspect", "1"))),
                                    new XElement(P + "nvPr")),
                                new XElement(P + "blipFill",
                                    new XElement(A + "blip", new XAttribute(R + "embed", "rId1")),
                                    new XElement(A + "stretch", new XElement(A + "fillRect"))),
                                new XElement(P + "spPr",
                                    new XElement(A + "xfrm",
                                        new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                                        new XElement(A + "ext",
                                            new XAttribute("cx", slideW),
                                            new XAttribute("cy", slideH))),
                                    new XElement(A + "prstGeom",
                                        new XAttribute("prst", "rect"),
                                        new XElement(A + "avLst")))))),
                    new XElement(P + "clrMapOvr",
                        new XElement(A + "masterClrMapping"))));
        }

        // ========== ppt/slides/_rels/slideN.xml.rels ==========

        private void WriteSlideRels(ZipArchive archive, int slideNum)
        {
            WriteXmlEntry(archive, $"ppt/slides/_rels/slide{slideNum}.xml.rels",
                new XElement(RelPkg + "Relationships",
                    new XElement(RelPkg + "Relationship",
                        new XAttribute("Id", "rId1"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
                        new XAttribute("Target", $"../media/image{slideNum}.{ImageExt}")),
                    new XElement(RelPkg + "Relationship",
                        new XAttribute("Id", "rId2"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"),
                        new XAttribute("Target", "../slideLayouts/slideLayout1.xml"))));
        }

        // ========== ppt/media/imageN.jpg ==========

        private void WriteMediaImage(ZipArchive archive, int slideNum, string imagePath)
        {
            var entry = archive.CreateEntry($"ppt/media/image{slideNum}.{ImageExt}", CompressionLevel.Optimal);
            using (var entryStream = entry.Open())
            using (var fileStream = File.OpenRead(imagePath))
            {
                fileStream.CopyTo(entryStream);
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
            }
            GC.SuppressFinalize(this);
        }
    }
}