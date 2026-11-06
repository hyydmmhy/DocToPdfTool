using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;

namespace DocToPdfTool.Utils
{
    /// <summary>识别出的一行文字及其包围盒。</summary>
    public class OcrLine
    {
        public string Text { get; set; }
        public Rectangle Box { get; set; }
        public float Score { get; set; }
    }

    /// <summary>单页识别结果。</summary>
    public class OcrPageResult
    {
        public int PageNumber { get; set; }
        public List<OcrLine> Lines { get; set; } = new List<OcrLine>();
        public string Text => string.Join(Environment.NewLine, Lines.Select(l => l.Text));
    }

    /// <summary>
    /// PP-OCRv4 推理引擎（检测 + 识别），基于 OnnxRuntime。
    ///
    /// 流程：
    ///   1. 检测模型输出文本区域概率图 → 二值化 → 连通域 → 外接框 → 按比例外扩；
    ///   2. 按检测框裁剪原图，统一缩放到高 48 像素后送识别模型；
    ///   3. CTC 贪心解码，得到每行文字。
    ///
    /// 说明：
    ///   - 仅支持横排文字（扫描件、打印件、截图的主流场景），不做任意角度矫正；
    ///   - 输入按 BGR、mean=0.5、std=0.5 归一化，与 PaddleOCR 训练时一致。
    /// </summary>
    public class PpOcrEngine : IDisposable
    {
        private const int DetLimitSideLenDefault = 960;
        private const int DetLimitSideLenRetry = 1600;
        private int _detLimitSideLen = DetLimitSideLenDefault;
        private const float DetThreshold = 0.3f;
        private const float BoxThreshold = 0.5f;
        private const float UnclipRatio = 1.6f;
        private const int RecHeight = 48;
        private const int RecMaxWidth = 320;

        private readonly InferenceSession _detSession;
        private readonly InferenceSession _recSession;
        private readonly string _detInputName;
        private readonly string _recInputName;
        private readonly List<string> _charset;

        /// <summary>调试用：true 表示按 BGR 送入模型，false 表示 RGB。</summary>
        public static bool UseBgr = true;

        private static bool _nativeReady;
        private static readonly object NativeLock = new object();

        /// <summary>
        /// 确保 OnnxRuntime 原生库可被加载。
        /// 目录发布时 DLL 就在程序目录；单文件发布时从嵌入资源释放到本地缓存目录，
        /// 再通过 SetDllDirectory 让 P/Invoke 能解析到它。
        /// </summary>
        private static void EnsureNativeLibrary()
        {
            if (_nativeReady) return;
            lock (NativeLock)
            {
                if (_nativeReady) return;

                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                if (File.Exists(Path.Combine(baseDir, "onnxruntime.dll")))
                {
                    _nativeReady = true;
                    return;
                }

                var targetDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "DocToPdfTool", "native");
                Directory.CreateDirectory(targetDir);

                var targetPath = Path.Combine(targetDir, "onnxruntime.dll");
                if (!File.Exists(targetPath) || new FileInfo(targetPath).Length < 1_000_000)
                {
                    var assembly = Assembly.GetExecutingAssembly();
                    // 单文件发布内嵌的是 gzip 版本（约 4.2 MB），解压后才是可加载的 DLL
                    using (var compressed = assembly.GetManifestResourceStream("DocToPdfTool.onnxruntime.dll.gz"))
                    {
                        if (compressed == null)
                            throw new InvalidOperationException("缺少 OCR 推理引擎的原生库 onnxruntime.dll");

                        using (var gzip = new System.IO.Compression.GZipStream(
                            compressed, System.IO.Compression.CompressionMode.Decompress))
                        using (var output = File.Create(targetPath))
                            gzip.CopyTo(output);
                    }
                }

                SetDllDirectory(targetDir);
                var handle = LoadLibrary(targetPath);
                if (handle == IntPtr.Zero)
                    throw new InvalidOperationException("加载 onnxruntime.dll 失败，错误码 " +
                        Marshal.GetLastWin32Error());

                _nativeReady = true;
            }
        }

        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr LoadLibrary(string path);

        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool SetDllDirectory(string path);

        public PpOcrEngine()
        {
            if (!OcrModelManager.IsReady)
                throw new InvalidOperationException("OCR 模型尚未下载");

            EnsureNativeLibrary();

            var options = new SessionOptions
            {
                IntraOpNumThreads = Math.Max(1, Environment.ProcessorCount / 2),
                InterOpNumThreads = 1,
                GraphOptimizationLevel = GraphOptimizationLevel.ORT_ENABLE_ALL
            };

            _detSession = new InferenceSession(OcrModelManager.DetModelPath, options);
            _recSession = new InferenceSession(OcrModelManager.RecModelPath, options);

            _detInputName = _detSession.InputMetadata.Keys.First();
            _recInputName = _recSession.InputMetadata.Keys.First();
            _charset = LoadCharset();
        }

        // ───────────────── 字符集 ─────────────────

        private static List<string> LoadCharset()
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream("DocToPdfTool.ocr_keys.txt"))
            {
                if (stream == null) throw new InvalidOperationException("缺少内置字符集资源");
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                {
                    var text = reader.ReadToEnd().Replace("\r\n", "\n").Replace('\r', '\n');
                    var lines = text.Split('\n').ToList();
                    while (lines.Count > 0 && lines[lines.Count - 1].Length == 0) lines.RemoveAt(lines.Count - 1);

                    // 模型输出索引 0 为 CTC blank，索引 1 对应字典第 1 行，
                    // 因此这里直接存字典本身（解码时统一按 best - 1 取），末尾补一个空格字符。
                    var charset = new List<string>(lines);
                    charset.Add(" ");
                    return charset;
                }
            }
        }

        // ───────────────── 对外入口 ─────────────────

        /// <summary>识别一张图片。</summary>
        public List<OcrLine> Recognize(Bitmap bitmap)
        {
            var boxes = Detect(bitmap);

            // 大图缩到 960 后小字号可能消失，没检测到任何框时用更大的边长再试一次
            if (boxes.Count == 0 && Math.Max(bitmap.Width, bitmap.Height) > 1200)
            {
                _detLimitSideLen = DetLimitSideLenRetry;
                try { boxes = Detect(bitmap); }
                finally { _detLimitSideLen = DetLimitSideLenDefault; }
            }

            var lines = new List<OcrLine>();

            foreach (var box in boxes)
            {
                var padded = Rectangle.Inflate(box, 2, 2);
                padded = Rectangle.Intersect(padded, new Rectangle(0, 0, bitmap.Width, bitmap.Height));
                var text = RecognizeBox(bitmap, padded.Width > 0 ? padded : box);
                if (string.IsNullOrWhiteSpace(text)) continue;
                lines.Add(new OcrLine { Text = text.Trim(), Box = box });
            }

            return lines;
        }

        // ───────────────── 检测 ─────────────────

        private List<Rectangle> Detect(Bitmap bitmap)
        {
            ComputeDetSize(bitmap.Width, bitmap.Height, out int inputW, out int inputH);

            var input = new DenseTensor<float>(new[] { 1, 3, inputH, inputW });
            FillDetInput(bitmap, input, inputW, inputH);

            var inputs = new List<NamedOnnxValue>
            {
                NamedOnnxValue.CreateFromTensor(_detInputName, input)
            };

            var boxes = new List<Rectangle>();
            using (var results = _detSession.Run(inputs))
            {
                var output = results.First().AsTensor<float>();
                var dims = output.Dimensions;
                int outH = dims[dims.Length - 2];
                int outW = dims[dims.Length - 1];

                var mask = new bool[outH * outW];
                var values = new float[outH * outW];
                for (int y = 0; y < outH; y++)
                    for (int x = 0; x < outW; x++)
                    {
                        int idx = y * outW + x;
                        float v = output[0, 0, y, x];
                        values[idx] = v;
                        mask[idx] = v > DetThreshold;
                    }

                // 掩码坐标 → 输入图坐标 → 原图坐标（两次缩放都要乘上）
                float toOriginX = (inputW / (float)outW) * (bitmap.Width / (float)inputW);
                float toOriginY = (inputH / (float)outH) * (bitmap.Height / (float)inputH);

                foreach (var component in FindComponents(mask, outW, outH))
                {
                    if (component.Area < 12) continue;

                    float sum = 0;
                    int count = 0;
                    foreach (var idx in component.Pixels)
                    {
                        sum += values[idx];
                        count++;
                    }
                    float score = count > 0 ? sum / count : 0;
                    if (score < BoxThreshold) continue;

                    var rect = component.Bounds;

                    float originWidth = rect.Width * toOriginX;
                    float originHeight = rect.Height * toOriginY;
                    float expand = ComputeUnclip((int)originWidth, (int)originHeight);

                    int left = (int)Math.Round(rect.Left * toOriginX - expand);
                    int top = (int)Math.Round(rect.Top * toOriginY - expand);
                    int width = (int)Math.Round(originWidth + expand * 2);
                    int height = (int)Math.Round(originHeight + expand * 2);

                    left = Math.Max(0, left);
                    top = Math.Max(0, top);
                    width = Math.Min(bitmap.Width - left, Math.Max(1, width));
                    height = Math.Min(bitmap.Height - top, Math.Max(1, height));

                    if (width < 6 || height < 6) continue;
                    boxes.Add(new Rectangle(left, top, width, height));
                }
            }

            return MergeLineBoxes(boxes);
        }

        private void ComputeDetSize(int width, int height, out int inputW, out int inputH)
        {
            float scale;
            int maxSide = Math.Max(width, height);
            int minSide = Math.Min(width, height);

            if (maxSide > _detLimitSideLen) scale = (float)_detLimitSideLen / maxSide;
            else if (minSide < _detLimitSideLen / 2) scale = (float)(_detLimitSideLen / 2) / minSide;
            else scale = 1f;

            inputW = Math.Max(32, RoundTo32((int)(width * scale)));
            inputH = Math.Max(32, RoundTo32((int)(height * scale)));
        }

        private static int RoundTo32(int value)
        {
            int rem = value % 32;
            return rem == 0 ? value : value + (32 - rem);
        }

        /// <summary>按 BGR、mean=0.5、std=0.5 归一化填充检测输入张量。</summary>
        private static void FillDetInput(Bitmap bitmap, DenseTensor<float> input, int inputW, int inputH)
        {
            using (var resized = new Bitmap(inputW, inputH, PixelFormat.Format24bppRgb))
            {
                using (var graphics = Graphics.FromImage(resized))
                {
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    graphics.DrawImage(bitmap, new Rectangle(0, 0, inputW, inputH));
                }

                var data = resized.LockBits(new Rectangle(0, 0, inputW, inputH),
                    ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                try
                {
                    var buffer = new byte[data.Stride * inputH];
                    Marshal.Copy(data.Scan0, buffer, 0, buffer.Length);

                    int plane = inputW * inputH;
                    for (int y = 0; y < inputH; y++)
                    {
                        int rowOffset = y * data.Stride;
                        for (int x = 0; x < inputW; x++)
                        {
                            int src = rowOffset + x * 3;
                            int idx = y * inputW + x;
                            int c0 = UseBgr ? buffer[src] : buffer[src + 2];
                            int c2 = UseBgr ? buffer[src + 2] : buffer[src];
                            input[0, 0, y, x] = (c0 / 255f - 0.5f) / 0.5f;
                            input[0, 1, y, x] = (buffer[src + 1] / 255f - 0.5f) / 0.5f;
                            input[0, 2, y, x] = (c2 / 255f - 0.5f) / 0.5f;
                        }
                    }
                }
                finally
                {
                    resized.UnlockBits(data);
                }
            }
        }

        private static float ComputeUnclip(int width, int height)
        {
            float area = width * (float)height;
            float perimeter = 2f * (width + height);
            if (perimeter <= 0) return 0;
            return area * UnclipRatio / perimeter;
        }

        private class Component
        {
            public Rectangle Bounds;
            public int Area;
            public List<int> Pixels = new List<int>();
        }

        /// <summary>四邻域连通域标记（BFS）。</summary>
        private static List<Component> FindComponents(bool[] mask, int width, int height)
        {
            var visited = new bool[mask.Length];
            var components = new List<Component>();
            var queue = new Queue<int>();

            for (int start = 0; start < mask.Length; start++)
            {
                if (!mask[start] || visited[start]) continue;

                var component = new Component();
                int minX = int.MaxValue, maxX = -1, minY = int.MaxValue, maxY = -1;

                queue.Clear();
                queue.Enqueue(start);
                visited[start] = true;

                while (queue.Count > 0)
                {
                    int idx = queue.Dequeue();
                    component.Pixels.Add(idx);
                    component.Area++;

                    int x = idx % width;
                    int y = idx / width;
                    if (x < minX) minX = x;
                    if (x > maxX) maxX = x;
                    if (y < minY) minY = y;
                    if (y > maxY) maxY = y;

                    if (x > 0) TryEnqueue(idx - 1);
                    if (x < width - 1) TryEnqueue(idx + 1);
                    if (y > 0) TryEnqueue(idx - width);
                    if (y < height - 1) TryEnqueue(idx + width);
                }

                component.Bounds = Rectangle.FromLTRB(minX, minY, maxX + 1, maxY + 1);
                components.Add(component);

                void TryEnqueue(int next)
                {
                    if (visited[next] || !mask[next]) return;
                    visited[next] = true;
                    queue.Enqueue(next);
                }
            }

            return components;
        }

        /// <summary>
        /// 合并同一行内被拆碎的检测框，再按「行 → 列」排序。
        ///
        /// DB 检测输出在字符间距较大时会切成多个连通域，直接输出会把一句话拆成好几段，
        /// 这里按「垂直重叠 &gt; 50% 且水平间距 &lt; 0.8 倍行高」判定为同一行并合并。
        /// </summary>
        private static List<Rectangle> MergeLineBoxes(List<Rectangle> boxes)
        {
            var list = boxes.OrderBy(b => b.Left).ToList();
            bool changed = true;

            while (changed)
            {
                changed = false;
                for (int i = 0; i < list.Count; i++)
                {
                    for (int j = i + 1; j < list.Count; j++)
                    {
                        var a = list[i];
                        var b = list[j];

                        int overlapY = Math.Min(a.Bottom, b.Bottom) - Math.Max(a.Top, b.Top);
                        float minHeight = Math.Min(a.Height, b.Height);
                        if (minHeight <= 0 || overlapY < minHeight * 0.5f) continue;

                        int gap = Math.Max(a.Left, b.Left) - Math.Min(a.Right, b.Right);
                        float maxGap = Math.Max(a.Height, b.Height) * 0.8f;
                        if (gap > maxGap) continue;

                        list[i] = Rectangle.Union(a, b);
                        list.RemoveAt(j);
                        changed = true;
                        j = i;
                    }
                }
            }

            // 行聚类：垂直中心相近的归为一行
            var rows = new List<List<Rectangle>>();
            foreach (var box in list.OrderBy(b => b.Top + b.Height / 2.0))
            {
                double center = box.Top + box.Height / 2.0;
                var row = rows.FirstOrDefault(r =>
                {
                    double rowCenter = r.Average(x => x.Top + x.Height / 2.0);
                    double tolerance = Math.Max(6, Math.Min(box.Height, r.Min(x => x.Height)) * 0.6);
                    return Math.Abs(center - rowCenter) <= tolerance;
                });

                if (row == null) rows.Add(new List<Rectangle> { box });
                else row.Add(box);
            }

            return rows
                .OrderBy(r => r.Average(b => b.Top + b.Height / 2.0))
                .SelectMany(r => r.OrderBy(b => b.Left))
                .ToList();
        }
        // ───────────────── 识别 ─────────────────

        private string RecognizeBox(Bitmap bitmap, Rectangle box)
        {
            box = Rectangle.Intersect(box, new Rectangle(0, 0, bitmap.Width, bitmap.Height));
            if (box.Width < 2 || box.Height < 2) return string.Empty;

            int targetW = (int)Math.Ceiling(RecHeight * (box.Width / (double)box.Height));
            targetW = Math.Max(8, Math.Min(RecMaxWidth, targetW));

            var input = new DenseTensor<float>(new[] { 1, 3, RecHeight, targetW });

            using (var cropped = new Bitmap(targetW, RecHeight, PixelFormat.Format24bppRgb))
            {
                using (var graphics = Graphics.FromImage(cropped))
                {
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    graphics.DrawImage(bitmap, new Rectangle(0, 0, targetW, RecHeight), box, GraphicsUnit.Pixel);
                }

                var data = cropped.LockBits(new Rectangle(0, 0, targetW, RecHeight),
                    ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                try
                {
                    var buffer = new byte[data.Stride * RecHeight];
                    Marshal.Copy(data.Scan0, buffer, 0, buffer.Length);

                    for (int y = 0; y < RecHeight; y++)
                    {
                        int rowOffset = y * data.Stride;
                        for (int x = 0; x < targetW; x++)
                        {
                            int src = rowOffset + x * 3;
                            int c0 = UseBgr ? buffer[src] : buffer[src + 2];
                            int c2 = UseBgr ? buffer[src + 2] : buffer[src];
                            input[0, 0, y, x] = (c0 / 255f - 0.5f) / 0.5f;
                            input[0, 1, y, x] = (buffer[src + 1] / 255f - 0.5f) / 0.5f;
                            input[0, 2, y, x] = (c2 / 255f - 0.5f) / 0.5f;
                        }
                    }
                }
                finally
                {
                    cropped.UnlockBits(data);
                }
            }

            var inputs = new List<NamedOnnxValue>
            {
                NamedOnnxValue.CreateFromTensor(_recInputName, input)
            };

            using (var results = _recSession.Run(inputs))
            {
                var output = results.First().AsTensor<float>();
                var dims = output.Dimensions;
                int steps = dims[dims.Length - 2];
                int classes = dims[dims.Length - 1];

                var builder = new StringBuilder();
                int previous = -1;

                for (int t = 0; t < steps; t++)
                {
                    int best = 0;
                    float bestScore = float.MinValue;
                    for (int c = 0; c < classes; c++)
                    {
                        float v = output[0, t, c];
                        if (v <= bestScore) continue;
                        bestScore = v;
                        best = c;
                    }

                    if (best != 0 && best != previous)
                    {
                        int index = best - 1;
                        if (index >= 0 && index < _charset.Count) builder.Append(_charset[index]);
                    }
                    previous = best;
                }

                return builder.ToString();
            }
        }

        /// <summary>调试：识别指定区域，返回文本与平均置信度。</summary>
        public string RecognizeRegion(Bitmap bitmap, Rectangle box, out float score, out int steps, out int classes)
        {
            score = 0; steps = 0; classes = 0;
            box = Rectangle.Intersect(box, new Rectangle(0, 0, bitmap.Width, bitmap.Height));
            if (box.Width < 2 || box.Height < 2) return string.Empty;

            int targetW = (int)Math.Ceiling(RecHeight * (box.Width / (double)box.Height));
            targetW = Math.Max(8, Math.Min(RecMaxWidth, targetW));

            var input = new DenseTensor<float>(new[] { 1, 3, RecHeight, targetW });
            using (var cropped = new Bitmap(targetW, RecHeight, PixelFormat.Format24bppRgb))
            {
                using (var graphics = Graphics.FromImage(cropped))
                {
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.DrawImage(bitmap, new Rectangle(0, 0, targetW, RecHeight), box, GraphicsUnit.Pixel);
                }
                var data = cropped.LockBits(new Rectangle(0, 0, targetW, RecHeight), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                try
                {
                    var buffer = new byte[data.Stride * RecHeight];
                    Marshal.Copy(data.Scan0, buffer, 0, buffer.Length);
                    for (int y = 0; y < RecHeight; y++)
                    {
                        int rowOffset = y * data.Stride;
                        for (int x = 0; x < targetW; x++)
                        {
                            int src = rowOffset + x * 3;
                            int c0 = UseBgr ? buffer[src] : buffer[src + 2];
                            int c2 = UseBgr ? buffer[src + 2] : buffer[src];
                            input[0, 0, y, x] = (c0 / 255f - 0.5f) / 0.5f;
                            input[0, 1, y, x] = (buffer[src + 1] / 255f - 0.5f) / 0.5f;
                            input[0, 2, y, x] = (c2 / 255f - 0.5f) / 0.5f;
                        }
                    }
                }
                finally { cropped.UnlockBits(data); }
            }

            var inputs = new List<NamedOnnxValue> { NamedOnnxValue.CreateFromTensor(_recInputName, input) };
            using (var results = _recSession.Run(inputs))
            {
                var output = results.First().AsTensor<float>();
                var dims = output.Dimensions;
                steps = dims[dims.Length - 2];
                classes = dims[dims.Length - 1];

                var builder = new StringBuilder();
                int previous = -1;
                float total = 0; int counted = 0;

                for (int t = 0; t < steps; t++)
                {
                    int best = 0; float bestScore = float.MinValue;
                    for (int c = 0; c < classes; c++)
                    {
                        float v = output[0, t, c];
                        if (v <= bestScore) continue;
                        bestScore = v; best = c;
                    }
                    if (best != 0)
                    {
                        total += bestScore; counted++;
                        if (best != previous)
                        {
                            int index = best - 1;
                            if (index >= 0 && index < _charset.Count) builder.Append(_charset[index]);
                        }
                    }
                    previous = best;
                }
                score = counted > 0 ? total / counted : 0;
                return builder.ToString();
            }
        }

        public void Dispose()
        {
            _detSession?.Dispose();
            _recSession?.Dispose();
        }
    }
}
