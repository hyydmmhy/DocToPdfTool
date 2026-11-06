using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DocToPdfTool.Utils
{
    /// <summary>翻译引擎。</summary>
    public enum TranslateEngine
    {
        /// <summary>微软翻译（免密钥签名通道）。</summary>
        Microsoft = 0,
        /// <summary>腾讯翻译君（transmart）。</summary>
        Tencent = 1
    }

    /// <summary>翻译参数。</summary>
    public class TranslateOptions
    {
        public TranslateEngine Engine { get; set; } = TranslateEngine.Microsoft;

        /// <summary>源语言代码，auto 表示自动识别。</summary>
        public string SourceLanguage { get; set; } = "auto";

        /// <summary>目标语言代码，如 zh-Hans / en / ja。</summary>
        public string TargetLanguage { get; set; } = "en";

        /// <summary>首选引擎失败时自动切换另一个引擎。</summary>
        public bool AutoFallback { get; set; } = true;
    }

    /// <summary>翻译结果。</summary>
    public class TranslateResult
    {
        public string Text { get; set; }
        public TranslateEngine UsedEngine { get; set; }
        public int SegmentCount { get; set; }
        public bool FallbackUsed { get; set; }
        public string Note { get; set; }
    }

    /// <summary>
    /// 文本翻译器，支持「微软翻译」与「腾讯翻译君」两个免密钥通道。
    ///
    /// 要点：
    ///   - 按行 + 句子边界切分，每段不超过 1000 字符（微软单次请求上限）；
    ///   - 同一批中重复段落只请求一次，译文按原顺序回填；
    ///   - 首选引擎失败（超时、限流、接口变更）时自动切换到另一个引擎；
    ///   - 空行与段落结构原样保留。
    /// </summary>
    public class TranslateService : IDisposable
    {
        private static readonly HttpClient Http = CreateClient();

        /// <summary>微软翻译签名通道使用的 64 字节密钥。</summary>
                private static readonly byte[] MicrosoftKey =
        {
            0xa2, 0x29, 0x3a, 0x3d, 0xd0, 0xdd, 0x32, 0x73,
            0x97, 0x7a, 0x64, 0xdb, 0xc2, 0xf3, 0x27, 0xf5,
            0xd7, 0xbf, 0x87, 0xd9, 0x45, 0x9d, 0xf0, 0x5a,
            0x09, 0x66, 0xc6, 0x30, 0xc6, 0x6a, 0xaa, 0x84,
            0x9a, 0x41, 0xaa, 0x94, 0x3a, 0xa8, 0xd5, 0x1a,
            0x6e, 0x4d, 0xaa, 0xc9, 0xa3, 0x70, 0x12, 0x35,
            0xc7, 0xeb, 0x12, 0xf6, 0xe8, 0x23, 0x07, 0x9e,
            0x47, 0x10, 0x95, 0x91, 0x88, 0x55, 0xd8, 0x17,
        };

        private const string TencentClientKey =
            "browser-chrome-110.0.0-Mac OS-df4bd4c5-a65d-44b2-a40f-42f34f3535f2-1677486696487";

        private const int MaxSegmentLength = 1000;

        private static HttpClient CreateClient()
        {
            var client = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
            client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36");
            return client;
        }

        public static string GetLanguageDisplayName(string code)
        {
            switch ((code ?? string.Empty).ToLowerInvariant())
            {
                case "zh-hans": case "zh-cn": case "zh": return "简体中文";
                case "zh-hant": case "zh-tw": return "繁体中文";
                case "en": return "英语";
                case "ja": return "日语";
                case "ko": return "韩语";
                case "fr": return "法语";
                case "de": return "德语";
                case "es": return "西班牙语";
                case "ru": return "俄语";
                default: return code;
            }
        }

        public static string GetEngineDisplayName(TranslateEngine engine)
        {
            return engine == TranslateEngine.Microsoft ? "微软翻译" : "腾讯翻译";
        }

        // ───────────────── 对外入口 ─────────────────

        public async Task<TranslateResult> TranslateAsync(string text, TranslateOptions options,
            Action<int, int> onProgress = null, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(text))
                throw new ArgumentException("没有可翻译的内容", nameof(text));

            options = options ?? new TranslateOptions();
            var segments = SplitIntoSegments(text, MaxSegmentLength);

            // 去重
            var unique = new List<string>();
            var map = new Dictionary<string, int>(StringComparer.Ordinal);
            var indexes = new int[segments.Count];
            for (int i = 0; i < segments.Count; i++)
            {
                if (!map.TryGetValue(segments[i], out int idx))
                {
                    idx = unique.Count;
                    unique.Add(segments[i]);
                    map[segments[i]] = idx;
                }
                indexes[i] = idx;
            }

            var translations = new string[unique.Count];
            var usedEngine = options.Engine;
            bool fallbackUsed = false;
            string note = null;

            try
            {
                await TranslateBatchAsync(unique, translations, options.Engine, options, onProgress, cancellationToken)
                    .ConfigureAwait(false);
            }
            catch (Exception ex) when (options.AutoFallback)
            {
                var alternative = options.Engine == TranslateEngine.Microsoft
                    ? TranslateEngine.Tencent
                    : TranslateEngine.Microsoft;

                note = $"{GetEngineDisplayName(options.Engine)}不可用（{Reason(ex)}），已自动切换到{GetEngineDisplayName(alternative)}";
                fallbackUsed = true;
                usedEngine = alternative;

                Array.Clear(translations, 0, translations.Length);
                await TranslateBatchAsync(unique, translations, alternative, options, onProgress, cancellationToken)
                    .ConfigureAwait(false);
            }

            var sb = new StringBuilder();
            for (int i = 0; i < segments.Count; i++)
                sb.AppendLine(translations[indexes[i]] ?? string.Empty);

            return new TranslateResult
            {
                Text = sb.ToString().TrimEnd() + Environment.NewLine,
                UsedEngine = usedEngine,
                SegmentCount = unique.Count,
                FallbackUsed = fallbackUsed,
                Note = note
            };
        }

        private static string Reason(Exception ex)
        {
            var message = ex is TaskCanceledException ? "请求超时" : ex.Message;
            message = (message ?? string.Empty).Replace("\r", " ").Replace("\n", " ");
            return message.Length > 80 ? message.Substring(0, 80) + "…" : message;
        }

        // ───────────────── 分段 ─────────────────

        public static List<string> SplitIntoSegments(string text, int maxLength)
        {
            maxLength = Math.Max(100, maxLength);
            var result = new List<string>();

            var normalized = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');

            foreach (var rawLine in normalized.Split('\n'))
            {
                var line = rawLine.TrimEnd();
                if (line.Trim().Length == 0)
                {
                    result.Add(string.Empty);
                    continue;
                }

                if (line.Length <= maxLength)
                {
                    result.Add(line);
                    continue;
                }

                result.AddRange(SplitLongLine(line, maxLength));
            }

            while (result.Count > 0 && result[result.Count - 1].Length == 0)
                result.RemoveAt(result.Count - 1);

            return result;
        }

        private static IEnumerable<string> SplitLongLine(string line, int maxLength)
        {
            var sentences = new List<string>();
            var builder = new StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                builder.Append(c);

                bool boundary = c == '。' || c == '！' || c == '？' || c == '；'
                             || c == '.' || c == '!' || c == '?' || c == ';';

                if (!boundary) continue;
                if (c == '.' && i + 1 < line.Length && char.IsLetter(line[i + 1])) continue;

                if (builder.Length >= maxLength)
                {
                    sentences.Add(builder.ToString().Trim());
                    builder.Clear();
                }
            }

            if (builder.Length > 0) sentences.Add(builder.ToString().Trim());

            var result = new List<string>();
            foreach (var sentence in sentences)
            {
                if (sentence.Length == 0) continue;
                if (sentence.Length <= maxLength) { result.Add(sentence); continue; }

                for (int i = 0; i < sentence.Length; i += maxLength)
                {
                    int len = Math.Min(maxLength, sentence.Length - i);
                    result.Add(sentence.Substring(i, len));
                }
            }
            return result;
        }

        // ───────────────── 批量调度 ─────────────────

        private async Task TranslateBatchAsync(List<string> segments, string[] output,
            TranslateEngine engine, TranslateOptions options, Action<int, int> onProgress,
            CancellationToken cancellationToken)
        {
            var pending = new List<int>();
            for (int i = 0; i < segments.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(segments[i])) output[i] = segments[i];
                else pending.Add(i);
            }
            if (pending.Count == 0) return;

            int batchSize = engine == TranslateEngine.Microsoft ? 20 : 10;
            int done = 0;

            for (int start = 0; start < pending.Count; start += batchSize)
            {
                var batch = pending.Skip(start).Take(batchSize).ToList();
                var texts = batch.Select(i => segments[i]).ToList();

                var translated = engine == TranslateEngine.Microsoft
                    ? await TranslateWithMicrosoftAsync(texts, options, cancellationToken).ConfigureAwait(false)
                    : await TranslateWithTencentAsync(texts, options, cancellationToken).ConfigureAwait(false);

                if (translated == null || translated.Count != texts.Count)
                    throw new InvalidOperationException("翻译接口返回的数据不完整");

                for (int i = 0; i < batch.Count; i++) output[batch[i]] = translated[i];

                done += batch.Count;
                onProgress?.Invoke(done, pending.Count);
            }
        }

        // ───────────────── 微软 ─────────────────

        private async Task<List<string>> TranslateWithMicrosoftAsync(List<string> texts,
            TranslateOptions options, CancellationToken cancellationToken)
        {
            var target = NormalizeMicrosoftLanguage(options.TargetLanguage, true);
            var path = "api.cognitive.microsofttranslator.com/translate?api-version=3.0&to="
                     + Uri.EscapeDataString(target);

            var source = NormalizeMicrosoftLanguage(options.SourceLanguage, false);
            if (!string.Equals(source, "auto", StringComparison.OrdinalIgnoreCase))
                path += "&from=" + Uri.EscapeDataString(source);

            var requestItems = texts.Select(t => new MsRequestItem { Text = t }).ToArray();
            var body = Serialize(requestItems);

            using (var request = new HttpRequestMessage(HttpMethod.Post, "https://" + path))
            {
                request.Content = new StringContent(body, Encoding.UTF8, "application/json");
                request.Headers.TryAddWithoutValidation("X-MT-Signature", BuildMicrosoftSignature(path));

                var response = await Http.SendAsync(request, cancellationToken).ConfigureAwait(false);
                var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                if (!response.IsSuccessStatusCode)
                    throw new InvalidOperationException(
                        $"微软翻译返回 {(int)response.StatusCode}：{Truncate(payload, 100)}");

                var entries = Deserialize<MsResponseItem[]>(payload);
                if (entries == null) throw new InvalidOperationException("微软翻译返回内容异常");

                return entries.Select(e =>
                {
                    var first = e.Translations != null && e.Translations.Length > 0 ? e.Translations[0] : null;
                    return first?.Text ?? string.Empty;
                }).ToList();
            }
        }

        private static string BuildMicrosoftSignature(string requestPath)
        {
            var guid = Guid.NewGuid().ToString("N");
            var escaped = Uri.EscapeDataString(requestPath);
            var dateTime = DateTimeOffset.UtcNow.ToString("ddd, dd MMM yyyy HH:mm:ss",
                CultureInfo.InvariantCulture) + "GMT";

            var payload = ("MSTranslatorAndroidApp" + escaped + dateTime + guid).ToLowerInvariant();

            using (var hmac = new HMACSHA256(MicrosoftKey))
            {
                var hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(payload));
                return "MSTranslatorAndroidApp::" + Convert.ToBase64String(hash) + "::" + dateTime + "::" + guid;
            }
        }

        private static string NormalizeMicrosoftLanguage(string code, bool isTarget)
        {
            var value = (code ?? "auto").Trim();
            switch (value.ToLowerInvariant())
            {
                case "auto": return isTarget ? "en" : "auto";
                case "zh": case "zh-cn": case "zh-hans": return "zh-Hans";
                case "zh-tw": case "zh-hant": return "zh-Hant";
                default: return value;
            }
        }

        // ───────────────── 腾讯 ─────────────────

        private async Task<List<string>> TranslateWithTencentAsync(List<string> texts,
            TranslateOptions options, CancellationToken cancellationToken)
        {
            var requestBody = new TencentRequest
            {
                Header = new TencentHeader { Fn = "auto_translation", ClientKey = TencentClientKey },
                Type = "plain",
                ModelCategory = "normal",
                Source = new TencentSource
                {
                    Lang = NormalizeTencentLanguage(options.SourceLanguage),
                    TextList = texts.ToArray()
                },
                Target = new TencentTarget { Lang = NormalizeTencentLanguage(options.TargetLanguage) }
            };

            using (var request = new HttpRequestMessage(HttpMethod.Post, "https://transmart.qq.com/api/imt"))
            {
                request.Content = new StringContent(Serialize(requestBody), Encoding.UTF8, "application/json");
                request.Headers.TryAddWithoutValidation("Referer", "https://yi.qq.com/zh-CN/index");

                var response = await Http.SendAsync(request, cancellationToken).ConfigureAwait(false);
                var payload = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                if (!response.IsSuccessStatusCode)
                    throw new InvalidOperationException(
                        $"腾讯翻译返回 {(int)response.StatusCode}：{Truncate(payload, 100)}");

                var parsed = Deserialize<TencentResponse>(payload);
                if (parsed == null) throw new InvalidOperationException("腾讯翻译返回内容异常");

                if (parsed.Header != null && !string.IsNullOrEmpty(parsed.Header.RetCode)
                    && parsed.Header.RetCode != "succ")
                    throw new InvalidOperationException("腾讯翻译返回错误：" + parsed.Header.RetCode);

                if (parsed.AutoTranslation == null) throw new InvalidOperationException("腾讯翻译返回内容异常");

                return parsed.AutoTranslation.Select(x => x ?? string.Empty).ToList();
            }
        }

        private static string NormalizeTencentLanguage(string code)
        {
            var value = (code ?? "auto").Trim();
            switch (value.ToLowerInvariant())
            {
                case "auto": return "auto";
                case "zh": case "zh-cn": case "zh-hans": return "zh";
                case "zh-tw": case "zh-hant": return "zh-TW";
                default: return value;
            }
        }

        // ───────────────── JSON ─────────────────

        private static string Serialize<T>(T value)
        {
            var serializer = new DataContractJsonSerializer(typeof(T));
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, value);
                return Encoding.UTF8.GetString(stream.ToArray());
            }
        }

        private static T Deserialize<T>(string json) where T : class
        {
            try
            {
                var serializer = new DataContractJsonSerializer(typeof(T));
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                    return serializer.ReadObject(stream) as T;
            }
            catch
            {
                return null;
            }
        }

        // ───────────────── 契约类型 ─────────────────

        [DataContract]
        private class MsRequestItem
        {
            [DataMember(Name = "Text", Order = 0)]
            public string Text { get; set; }
        }

        [DataContract]
        private class MsTranslation
        {
            [DataMember(Name = "text", Order = 0)]
            public string Text { get; set; }

            [DataMember(Name = "to", Order = 1)]
            public string To { get; set; }
        }

        [DataContract]
        private class MsResponseItem
        {
            [DataMember(Name = "translations", Order = 0)]
            public MsTranslation[] Translations { get; set; }
        }

        [DataContract]
        private class TencentHeader
        {
            [DataMember(Name = "fn", Order = 0)]
            public string Fn { get; set; }

            [DataMember(Name = "client_key", Order = 1)]
            public string ClientKey { get; set; }
        }

        [DataContract]
        private class TencentSource
        {
            [DataMember(Name = "lang", Order = 0)]
            public string Lang { get; set; }

            [DataMember(Name = "text_list", Order = 1)]
            public string[] TextList { get; set; }
        }

        [DataContract]
        private class TencentTarget
        {
            [DataMember(Name = "lang", Order = 0)]
            public string Lang { get; set; }
        }

        [DataContract]
        private class TencentRequest
        {
            [DataMember(Name = "header", Order = 0)]
            public TencentHeader Header { get; set; }

            [DataMember(Name = "type", Order = 1)]
            public string Type { get; set; }

            [DataMember(Name = "model_category", Order = 2)]
            public string ModelCategory { get; set; }

            [DataMember(Name = "source", Order = 3)]
            public TencentSource Source { get; set; }

            [DataMember(Name = "target", Order = 4)]
            public TencentTarget Target { get; set; }
        }

        [DataContract]
        private class TencentResponseHeader
        {
            [DataMember(Name = "ret_code", Order = 0)]
            public string RetCode { get; set; }
        }

        [DataContract]
        private class TencentResponse
        {
            [DataMember(Name = "header", Order = 0)]
            public TencentResponseHeader Header { get; set; }

            [DataMember(Name = "auto_translation", Order = 1)]
            public string[] AutoTranslation { get; set; }
        }

        private static string Truncate(string text, int max)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            return text.Length <= max ? text : text.Substring(0, max) + "…";
        }

        public void Dispose() { }
    }
}
