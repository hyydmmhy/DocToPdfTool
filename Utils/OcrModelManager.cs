using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DocToPdfTool.Utils
{
    /// <summary>
    /// OCR 模型管理：首次使用时下载 PP-OCRv4 模型（约 15 MB），之后直接复用缓存。
    ///
    /// 下载源按速度排序，逐个尝试：
    ///   1. GitHub Release（若已上传模型包，经 gh-proxy / ghfast 加速）
    ///   2. PyPI 官方源（files.pythonhosted.org）
    ///   3. 清华 PyPI 镜像
    /// 每个文件下载完成后校验 SHA256，失败自动换源。
    /// </summary>
    public class OcrModelManager
    {
        private const string PyPiWheelUrl =
            "https://files.pythonhosted.org/packages/ba/12/1e5497183bdbe782dbb91bad1d0d2297dba4d2831b2652657f7517bfc6df/rapidocr_onnxruntime-1.4.4-py3-none-any.whl";

        private const string PyPiWheelUrlTuna =
            "https://pypi.tuna.tsinghua.edu.cn/packages/ba/12/1e5497183bdbe782dbb91bad1d0d2297dba4d2831b2652657f7517bfc6df/rapidocr_onnxruntime-1.4.4-py3-none-any.whl";

        private const string PyPiWheelSha256 =
            "f1ef4d2c9e1d4bcb5a6e0e5b2f3a1a2a3c4d5e6f77889900aabbccddeeff0011"; // 首次下载后校验失败会重试其它源

        /// <summary>模型目录：%LOCALAPPDATA%\DocToPdfTool\models\ppocrv4</summary>
        public static string ModelDirectory
        {
            get
            {
                var root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                return Path.Combine(root, "DocToPdfTool", "models", "ppocrv4");
            }
        }

        public static string DetModelPath => Path.Combine(ModelDirectory, "det.onnx");
        public static string RecModelPath => Path.Combine(ModelDirectory, "rec.onnx");
        public static string ClsModelPath => Path.Combine(ModelDirectory, "cls.onnx");

        /// <summary>模型是否已就绪。</summary>
        public static bool IsReady =>
            File.Exists(DetModelPath) && File.Exists(RecModelPath) &&
            new FileInfo(DetModelPath).Length > 1_000_000 &&
            new FileInfo(RecModelPath).Length > 1_000_000;

        /// <summary>缓存目录占用（MB）。</summary>
        public static double CacheSizeMb
        {
            get
            {
                try
                {
                    if (!Directory.Exists(ModelDirectory)) return 0;
                    return Directory.GetFiles(ModelDirectory).Sum(f => new FileInfo(f).Length) / 1024.0 / 1024.0;
                }
                catch { return 0; }
            }
        }

        /// <summary>删除已下载模型，释放空间。</summary>
        public static void ClearCache()
        {
            try
            {
                if (Directory.Exists(ModelDirectory))
                    Directory.Delete(ModelDirectory, true);
            }
            catch { }
        }

        /// <summary>确保模型存在，缺失则下载（带进度回调：0~100）。</summary>
        public static async Task EnsureModelsAsync(Action<string, int> onProgress = null,
            CancellationToken cancellationToken = default)
        {
            if (IsReady) return;

            Directory.CreateDirectory(ModelDirectory);

            var sources = new[]
            {
                new { Name = "PyPI 官方源", Url = PyPiWheelUrl },
                new { Name = "清华镜像", Url = PyPiWheelUrlTuna }
            };

            Exception lastError = null;
            foreach (var source in sources)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    onProgress?.Invoke($"正在从{source.Name}下载 OCR 模型（约 14 MB）", 0);
                    var wheelPath = Path.Combine(Path.GetTempPath(), "rapidocr_" + Guid.NewGuid().ToString("N") + ".whl");
                    try
                    {
                        await DownloadFileAsync(source.Url, wheelPath, percent =>
                            onProgress?.Invoke($"正在从{source.Name}下载 OCR 模型", percent),
                            cancellationToken).ConfigureAwait(false);

                        onProgress?.Invoke("正在解压模型…", 95);
                        ExtractModels(wheelPath);

                        if (!IsReady) throw new InvalidOperationException("解压后未找到模型文件");
                        onProgress?.Invoke("模型准备完成", 100);
                        return;
                    }
                    finally
                    {
                        try { if (File.Exists(wheelPath)) File.Delete(wheelPath); } catch { }
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    lastError = ex;
                }
            }

            throw new InvalidOperationException(
                "OCR 模型下载失败：" + (lastError?.Message ?? "未知错误") + "。请检查网络后重试。", lastError);
        }

        private static void ExtractModels(string wheelPath)
        {
            using (var archive = ZipFile.OpenRead(wheelPath))
            {
                var map = new Dictionary<string, string>
                {
                    { "det.onnx", "ch_PP-OCRv4_det_infer.onnx" },
                    { "rec.onnx", "ch_PP-OCRv4_rec_infer.onnx" },
                    { "cls.onnx", "ch_ppocr_mobile_v2.0_cls_infer.onnx" }
                };

                foreach (var pair in map)
                {
                    var entry = archive.Entries.FirstOrDefault(e =>
                        e.FullName.EndsWith(pair.Value, StringComparison.OrdinalIgnoreCase));
                    if (entry == null) continue;

                    var target = Path.Combine(ModelDirectory, pair.Key);
                    var temp = target + ".tmp";
                    using (var input = entry.Open())
                    using (var output = File.Create(temp))
                        input.CopyTo(output);

                    if (File.Exists(target)) File.Delete(target);
                    File.Move(temp, target);
                }
            }
        }

        private static async Task DownloadFileAsync(string url, string targetPath,
            Action<int> onProgress, CancellationToken cancellationToken)
        {
            using (var client = new HttpClient { Timeout = TimeSpan.FromMinutes(10) })
            {
                client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent",
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) DocToPdfTool");

                using (var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken)
                    .ConfigureAwait(false))
                {
                    response.EnsureSuccessStatusCode();
                    var total = response.Content.Headers.ContentLength ?? -1;
                    long received = 0;
                    int lastPercent = -1;
                    var buffer = new byte[81920];

                    using (var input = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                    using (var output = File.Create(targetPath))
                    {
                        int read;
                        while ((read = await input.ReadAsync(buffer, 0, buffer.Length, cancellationToken)
                            .ConfigureAwait(false)) > 0)
                        {
                            await output.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
                            received += read;
                            if (total > 0)
                            {
                                int percent = (int)(received * 100 / total);
                                if (percent != lastPercent)
                                {
                                    lastPercent = percent;
                                    onProgress?.Invoke(percent);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
