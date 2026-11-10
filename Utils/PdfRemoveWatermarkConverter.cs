using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Graphics.Operations.TextState;
using UglyToad.PdfPig.Writer;

namespace DocToPdfTool.Utils
{
    /// <summary>去水印结果。</summary>
    public class PdfRemoveWatermarkResult
    {
        public string OutputFile { get; set; }
        public int SourcePages { get; set; }
        public int OutputPages { get; set; }
        public int WatermarkBlocksRemoved { get; set; }
        public long OriginalSize { get; set; }
        public long NewSize { get; set; }
        public bool IsComplexPdf { get; set; }
    }

    /// <summary>
    /// PDF 去文字水印（基于 PdfPig，零新增依赖）。
    ///
    /// 识别：跨页重复出现的大字号文字块（水印特征：大字号、多页同位置重复）。
    /// 删除：用 PdfDocumentBuilder 复制每页，重建内容流时跳过水印块（BT...ET）。
    /// 正文矢量保留（可选中/可复制），水印消失。
    /// </summary>
    public class PdfRemoveWatermarkConverter : IDisposable
    {
        /// <summary>判定为水印的最小字号（低于此视为正文）。</summary>
        public double MinWatermarkFontSize { get; set; } = 24;

        /// <summary>水印需在至少多少页重复出现。</summary>
        public int MinRepeatedPages { get; set; } = 2;

        /// <summary>关键词需重复出现至少多少次才判为水印（避免正文长句误判）。</summary>
        public int MinKeywordRepeat { get; set; } = 6;

        /// <summary>用户手动指定的水印文字（自动关键词未覆盖时，可精确指定）。</summary>
        public List<string> WatermarkTexts { get; set; } = null;

        public PdfRemoveWatermarkResult Convert(string sourcePath, string outputDir,
            List<string> watermarkTexts = null)
        {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("PDF 文件不存在", sourcePath);
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            long originalSize = new FileInfo(sourcePath).Length;
            var result = new PdfRemoveWatermarkResult { OriginalSize = originalSize };
            var outputFile = UniquePath(outputDir,
                Path.GetFileNameWithoutExtension(sourcePath) + "_去水印");
            int removedBlocks = 0;

            // 统计每页的文字块（BT...ET），记录字号 + 文字字符数
            var blocksByPage = new List<List<TextBlockInfo>>();

            using (var doc = PdfDocument.Open(sourcePath))
            {
                int total = doc.NumberOfPages;
                result.SourcePages = total;
                result.OutputPages = total;

                // 第一遍：收集每页文字块，用于跨页重复统计
                for (int i = 1; i <= total; i++)
                {
                    var page = doc.GetPage(i);
                    blocksByPage.Add(ExtractTextBlocks(page));
                }

                // 识别水印字号：多页重复出现且字号 > 阈值
                var watermarkFontSizes = DetectWatermarkSizes(blocksByPage, result);

                // 安全策略：识别不到水印 → 原样复制，不碰内容流（避免破坏正文）
                if (watermarkFontSizes.Count == 0)
                {
                    File.Copy(sourcePath, outputFile, true);
                    result.OutputFile = outputFile;
                    result.NewSize = new FileInfo(outputFile).Length;
                    return result;
                }

                // 检测 PDF 是否复杂（Type3/Form 字体重建会丢内容）→ 原样复制并标记
                bool complex = IsComplexPdf(sourcePath);
                result.IsComplexPdf = complex;
                if (complex)
                {
                    File.Copy(sourcePath, outputFile, true);
                    result.OutputFile = outputFile;
                    result.NewSize = new FileInfo(outputFile).Length;
                    result.WatermarkBlocksRemoved = 0;   // 未实际删除（不安全）
                    return result;
                }

                // 第二遍：复制每页，重建内容流（跳过水印块）
                var builder = new PdfDocumentBuilder();
                for (int i = 1; i <= total; i++)
                {
                    var page = doc.GetPage(i);
                    var pb = builder.AddPage(doc, i);
                    // 清空复制的内容流（运行时为 List<IPageContentStream>，静态类型 IReadOnlyList 无 Clear）
                    if (pb.ContentStreams is IList list)
                        list.Clear();
                    pb.NewContentStreamBefore();
                    var stream = pb.CurrentStream;

                    // 遍历操作，跳过水印块（BT...ET 整体收集，非水印则整个重放）
                    // SetFontAndSize（Tf）可能在 BT 前，需全局跟踪字号
                    // /Artifact /Watermark BDC...EMC 整块删除（PDF 标准水印标记，含水印 Form Do）
                    var ops = page.Operations;
                    var blockOps = new List<UglyToad.PdfPig.Graphics.Operations.IGraphicsStateOperation>();
                    bool inBlock = false;
                    bool inWatermarkArtifact = false;
                    int artifactDepth = 0;
                    double lastFontSize = 0;
                    double blockMaxSize = 0;
                    for (int k = 0; k < ops.Count; k++)
                    {
                        var op = ops[k];
                        var name = op.GetType().Name;

                        // /Artifact /Watermark BDC → 进入跳过模式（删到 EMC）
                        if (name == "BeginMarkedContentWithProperties")
                        {
                            if (IsWatermarkArtifact(op))
                            {
                                inWatermarkArtifact = true;
                                artifactDepth = 1;
                                removedBlocks++;
                                continue;
                            }
                        }
                        if (inWatermarkArtifact)
                        {
                            if (name == "BeginMarkedContent" || name == "BeginMarkedContentWithProperties") artifactDepth++;
                            if (name == "EndMarkedContent")
                            {
                                artifactDepth--;
                                if (artifactDepth <= 0) { inWatermarkArtifact = false; artifactDepth = 0; }
                            }
                            continue;   // 跳过水印块内所有操作
                        }

                        if (name == "SetFontAndSize")
                        {
                            double sz = ReadFontSize(op);
                            if (sz > 0) lastFontSize = sz;
                            if (inBlock && sz > blockMaxSize) blockMaxSize = sz;
                            stream.Operations.Add(op);   // 字体设置必须保留
                            continue;
                        }

                        if (name == "BeginText")
                        {
                            inBlock = true;
                            blockMaxSize = lastFontSize;
                            blockOps.Clear();
                            blockOps.Add(op);
                            continue;
                        }

                        if (name == "EndText")
                        {
                            inBlock = false;
                            blockOps.Add(op);
                            // 水印判定：块字号在识别出的水印字号集合内即删除（不限阈值，兼容小字水印）
                            bool isWatermark = IsWatermarkSize(blockMaxSize, watermarkFontSizes);
                            if (!isWatermark)
                            {
                                foreach (var bo in blockOps)
                                    stream.Operations.Add(bo);
                            }
                            else
                            {
                                removedBlocks++;
                            }
                            continue;
                        }

                        if (inBlock)
                        {
                            blockOps.Add(op);
                        }
                        else
                        {
                            stream.Operations.Add(op);
                        }
                    }
                }

                using (doc)
                {
                    var bytes = builder.Build();
                    File.WriteAllBytes(outputFile, bytes);
                }

                // 安全网：验证重建后正文是否保留（读输出文字），若丢失/读取失败 → 回退原样
                // 重建对某些复杂 PDF（Type0/CID 字体）会丢正文，必须原样保护
                if (!VerifyBodyPreserved(sourcePath, outputFile))
                {
                    File.Copy(sourcePath, outputFile, true);
                    result.OutputFile = outputFile;
                    result.NewSize = new FileInfo(outputFile).Length;
                    result.WatermarkBlocksRemoved = 0;   // 回退，未实际删除
                    return result;
                }
            }

            result.WatermarkBlocksRemoved = removedBlocks;
            result.OutputFile = outputFile;
            result.NewSize = new FileInfo(outputFile).Length;
            return result;
        }

        /// <summary>验证重建后正文是否保留：读原页1 与输出页1 文字，若输出丢失正文/读取失败 → 回退原样。</summary>
        private static bool VerifyBodyPreserved(string sourcePath, string outputFile)
        {
            try
            {
                string origText = "";
                using (var d = UglyToad.PdfPig.PdfDocument.Open(sourcePath))
                    origText = d.GetPage(1).Text ?? "";

                string outText = "";
                using (var d = UglyToad.PdfPig.PdfDocument.Open(outputFile))
                    outText = d.GetPage(1).Text ?? "";

                if (string.IsNullOrWhiteSpace(outText)) return false;   // 输出无文字 → 丢正文

                // 原文字前 15 个非空白字符作正文锚点（正文通常在水印前）
                var anchor = new string((origText ?? "").Where(c => !char.IsWhiteSpace(c)).Take(15).ToArray());
                if (anchor.Length >= 4 && !outText.Contains(anchor))
                    return false;   // 输出不含原正文开头 → 丢正文

                // 输出文字过少（< 原60%）→ 可能丢正文
                if (origText.Length > 50 && outText.Length < origText.Length * 0.6)
                    return false;

                return true;
            }
            catch
            {
                return false;   // 读取输出失败（复杂字体）→ 回退原样，安全
            }
        }

        /// <summary>检测 PDF 是否含 Type3 字体（重建会丢内容，需原样保护）。
        /// 不保护 Form XObject：水印常在 Form 里（/Artifact /Watermark），删 BDC 块即可处理。</summary>
        private static bool IsComplexPdf(string sourcePath)
        {
            try
            {
                var bytes = File.ReadAllBytes(sourcePath);
                var txt = System.Text.Encoding.ASCII.GetString(bytes);
                if (txt.Contains("/Type3")) return true;
            }
            catch { }
            return false;
        }

        /// <summary>提取一页的所有文字块：Letters 按字号聚类（CJK 可读），并结合 Operations 分析旋转/灰色（水印特征）。</summary>
        private static List<TextBlockInfo> ExtractTextBlocks(Page page)
        {
            var blocks = new List<TextBlockInfo>();
            try
            {
                // 1. 遍历 Operations：按字号标记旋转45°(斜排 Tm/cm)和灰色(半透明 g/rg) —— 水印典型特征
                // 旋转只认45°(|b|=|c|≈0.7071)，正文也可能旋转但通常非标准45°；配合短文本+单页防误判
                var rotatedSizes = new HashSet<double>();
                var graySizes = new HashSet<double>();
                double lastSize = 0;
                foreach (var op in page.Operations)
                {
                    var n = op.GetType().Name;
                    if (n == "SetFontAndSize")
                    {
                        double sz = Math.Round(ReadFontSize(op), 1);
                        if (sz > 0) lastSize = sz;
                    }
                    else if (n == "SetTextMatrix" || n == "ModifyCurrentTransformationMatrix")
                    {
                        // Tm a b c d e f —— 旋转45°: |b|=|c|≈0.7071
                        try
                        {
                            double b = 0, c = 0;
                            var valProp = op.GetType().GetProperty("Value");
                            if (valProp != null && valProp.GetValue(op) is Array arr && arr.Length >= 3)
                            {
                                b = System.Convert.ToDouble(arr.GetValue(1));
                                c = System.Convert.ToDouble(arr.GetValue(2));
                            }
                            else
                            {
                                var m = op.GetType().GetProperty("Matrix").GetValue(op);
                                b = System.Convert.ToDouble(m.GetType().GetProperty("B").GetValue(m));
                                c = System.Convert.ToDouble(m.GetType().GetProperty("C").GetValue(m));
                            }
                            if (Math.Abs(Math.Abs(b) - 0.7071) < 0.03 && Math.Abs(Math.Abs(c) - 0.7071) < 0.03)
                                if (lastSize > 0) rotatedSizes.Add(lastSize);
                        }
                        catch { }
                    }
                    else if (n == "SetNonStrokeColorDeviceRgb" || n == "SetStrokeColorDeviceRgb"
                             || n == "SetNonStrokeColorDeviceGray" || n == "SetStrokeColorDeviceGray"
                             || n == "SetNonStrokeColorAdvanced" || n == "SetStrokeColorAdvanced")
                    {
                        // rg r g b / g gray / scn —— 灰色(半透明): 三通道近似且 0.05<v<0.9
                        try
                        {
                            var ct = op.GetType();
                            double r = 0, g = 0, bl = 0;
                            var rProp = ct.GetProperty("R");
                            var grayProp = ct.GetProperty("Gray");
                            var operandsProp = ct.GetProperty("Operands");
                            if (rProp != null)
                            {
                                r = System.Convert.ToDouble(rProp.GetValue(op));
                                g = System.Convert.ToDouble(ct.GetProperty("G").GetValue(op));
                                bl = System.Convert.ToDouble(ct.GetProperty("B").GetValue(op));
                            }
                            else if (grayProp != null)
                            {
                                r = g = bl = System.Convert.ToDouble(grayProp.GetValue(op));
                            }
                            else if (operandsProp != null && operandsProp.GetValue(op) is IEnumerable olist)
                            {
                                var oa = olist.Cast<object>().ToArray();
                                if (oa.Length >= 3)
                                {
                                    r = System.Convert.ToDouble(oa[0]);
                                    g = System.Convert.ToDouble(oa[1]);
                                    bl = System.Convert.ToDouble(oa[2]);
                                }
                            }
                            else
                            {
                                var col = ct.GetProperty("Colour").GetValue(op);
                                r = System.Convert.ToDouble(col.GetType().GetProperty("R").GetValue(col));
                                g = System.Convert.ToDouble(col.GetType().GetProperty("G").GetValue(col));
                                bl = System.Convert.ToDouble(col.GetType().GetProperty("B").GetValue(col));
                            }
                            if (Math.Abs(r - g) < 0.02 && Math.Abs(g - bl) < 0.02 && r > 0.05 && r < 0.9)
                                if (lastSize > 0) graySizes.Add(lastSize);
                        }
                        catch { }
                    }
                }

                // 2. Letters 按字号聚类文本（CJK 可读）
                var bySize = new Dictionary<double, System.Text.StringBuilder>();
                foreach (var l in page.Letters)
                {
                    double sz = Math.Round((double)l.FontSize, 1);
                    if (!bySize.TryGetValue(sz, out var sb)) { sb = new System.Text.StringBuilder(); bySize[sz] = sb; }
                    sb.Append(l.Value);
                }
                foreach (var kv in bySize)
                {
                    blocks.Add(new TextBlockInfo
                    {
                        FontSize = kv.Key,
                        Text = kv.Value.ToString(),
                        Count = kv.Value.Length,
                        IsRotated = rotatedSizes.Contains(kv.Key),
                        IsGray = graySizes.Contains(kv.Key)
                    });
                }
            }
            catch { }
            return blocks;
        }

        /// <summary>判断 BDC 是否为 /Artifact 且 /Subtype /Watermark（PDF 标准水印标记）。</summary>
        private static bool IsWatermarkArtifact(object op)
        {
            try
            {
                var t = op.GetType();
                var nameProp = t.GetProperty("Name");
                var nm = nameProp?.GetValue(op);
                if (nm != null && nm.ToString().Contains("Artifact"))
                {
                    var propsProp = t.GetProperty("Properties");
                    var props = propsProp?.GetValue(op);
                    if (props != null && props.ToString().Contains("Watermark"))
                        return true;
                }
            }
            catch { }
            return false;
        }

        /// <summary>读取 SetFontAndSize 的字号（反射，兼容不同命名空间的类型）。</summary>
        private static double ReadFontSize(object op)
        {
            try
            {
                var sizeProp = op.GetType().GetProperty("Size");
                if (sizeProp != null)
                {
                    var v = sizeProp.GetValue(op);
                    if (v != null) return System.Convert.ToDouble(v);
                }
            }
            catch { }
            return 0;
        }

        /// <summary>字号是否在水印集合内（容差比较）。</summary>
        private static bool IsWatermarkSize(double size, List<double> watermarkSizes)
        {
            foreach (var w in watermarkSizes)
                if (Math.Abs(w - size) < 0.5) return true;
            return false;
        }

        /// <summary>水印关键词（常见水印专用词，含即可判为水印；避免用"测试"等正文常见词）。</summary>
        private static readonly string[] WatermarkKeywords = new[]
        {
            "机密", "内部", "草稿", "绝密", "秘密", "内部资料", "受控文件",
            "COPY", "DRAFT", "CONFIDENTIAL", "INTERNAL", "PREVIEW", "VOID"
        };

        /// <summary>识别水印字号：关键词匹配，或大字号 + 跨页重复 / 页内大量重复。</summary>
        private List<double> DetectWatermarkSizes(List<List<TextBlockInfo>> blocksByPage, PdfRemoveWatermarkResult result)
        {
            // 统计每个字号：出现页数 + 页内总次数 + 字符序列（拼接所有页）+ 旋转/灰色标记
            var sizePages = new Dictionary<double, HashSet<int>>();
            var sizeCount = new Dictionary<double, int>();
            var sizeText = new Dictionary<double, System.Text.StringBuilder>();
            var sizeRotated = new Dictionary<double, bool>();
            var sizeGray = new Dictionary<double, bool>();

            for (int i = 0; i < blocksByPage.Count; i++)
            {
                foreach (var b in blocksByPage[i])
                {
                    double sz = b.FontSize;
                    if (!sizePages.TryGetValue(sz, out var pages)) { pages = new HashSet<int>(); sizePages[sz] = pages; }
                    pages.Add(i);
                    sizeCount[sz] = sizeCount.TryGetValue(sz, out int c) ? c + b.Count : b.Count;
                    if (!sizeText.TryGetValue(sz, out var sb)) { sb = new System.Text.StringBuilder(); sizeText[sz] = sb; }
                    sb.Append(b.Text);
                    if (b.IsRotated) sizeRotated[sz] = true;
                    if (b.IsGray) sizeGray[sz] = true;
                }
            }

            var watermarks = new List<double>();
            foreach (var kv in sizePages)
            {
                double sz = kv.Key;
                int pageCount = kv.Value.Count;
                string text = sizeText.TryGetValue(sz, out var sb) ? sb.ToString() : "";

                bool matched = false;
                bool isGray = sizeGray.TryGetValue(sz, out bool gry) && gry;
                bool isRotated = sizeRotated.TryGetValue(sz, out bool rot) && rot;

                // 规则0：灰色(半透明) 或 旋转45° + 大字号 + 短文本 + 单页 → 水印
                // 支持单页1个/多个水印；正文长句/跨页大标题不误判
                if (!matched && sz >= MinWatermarkFontSize && (isGray || isRotated)
                    && text.Length > 0 && text.Length < 25 && pageCount == 1)
                    matched = true;

                // 规则1：大字号 + 跨页重复 + 短文本（水印是短词重复，正文长句/页眉标题不误判）
                if (!matched && sz >= MinWatermarkFontSize && pageCount >= MinRepeatedPages
                    && text.Length > 0 && text.Length < 25)
                    matched = true;

                // 规则2：关键词重复 ≥ 阈值（"机密/内部/草稿/COPY"等）
                if (!matched && !string.IsNullOrEmpty(text))
                {
                    var upper = text.ToUpperInvariant();
                    foreach (var kw in WatermarkKeywords)
                    {
                        if (CountOccurrences(upper, kw) >= MinKeywordRepeat) { matched = true; break; }
                    }
                }

                // 规则3：用户手动指定的水印文字（重复 ≥ 阈值）
                if (!matched && WatermarkTexts != null && WatermarkTexts.Count > 0 && !string.IsNullOrEmpty(text))
                {
                    foreach (var wt in WatermarkTexts)
                    {
                        if (!string.IsNullOrEmpty(wt)
                            && CountOccurrences(text, wt) >= MinKeywordRepeat) { matched = true; break; }
                    }
                }

                if (matched)
                    watermarks.Add(sz);
            }
            result.WatermarkBlocksRemoved = watermarks.Count;
            return watermarks;
        }

        /// <summary>统计关键词在文本中出现的次数（非重叠）。</summary>
        private static int CountOccurrences(string text, string keyword)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(keyword)) return 0;
            int count = 0, idx = 0;
            while ((idx = text.IndexOf(keyword, idx, StringComparison.OrdinalIgnoreCase)) >= 0)
            {
                count++;
                idx += keyword.Length;
            }
            return count;
        }

        private class TextBlockInfo
        {
            public double FontSize { get; set; }
            public string Text { get; set; }
            public int Count { get; set; }
            public bool IsRotated { get; set; }
            public bool IsGray { get; set; }
        }

        private static string UniquePath(string dir, string baseName)
        {
            var path = Path.Combine(dir, baseName + ".pdf");
            int seq = 1;
            while (File.Exists(path))
                path = Path.Combine(dir, $"{baseName}_{seq++}.pdf");
            return path;
        }

        public void Dispose() { }
    }
}
