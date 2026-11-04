using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Spire.Pdf;
using SpirePdfDocument = Spire.Pdf.PdfDocument;

namespace DocToPdfTool.Utils
{
    /// <summary>PDF 权限预设。</summary>
    public enum PdfEncryptPreset
    {
        /// <summary>全部允许。</summary>
        AllowAll = 0,

        /// <summary>仅查看：禁止打印、复制、修改、批注、填表。</summary>
        ViewOnly = 1,

        /// <summary>允许打印：仅允许打印，其余禁止。</summary>
        AllowPrint = 2,

        /// <summary>自定义：按复选框逐项设置。</summary>
        Custom = 3
    }

    /// <summary>PDF 加密参数。</summary>
    public class PdfEncryptOptions
    {
        /// <summary>打开密码（用户密码），必填。</summary>
        public string UserPassword { get; set; }

        /// <summary>编辑密码（所有者密码），可空；为空时不设置编辑密码。</summary>
        public string OwnerPassword { get; set; }

        /// <summary>加密算法（仅在设置了编辑密码时可选）。</summary>
        public PdfEncryptionAlgorithm Algorithm { get; set; } = PdfEncryptionAlgorithm.AES_256;

        /// <summary>权限预设。</summary>
        public PdfEncryptPreset Preset { get; set; } = PdfEncryptPreset.AllowAll;

        public bool AllowPrint { get; set; } = true;
        public bool AllowCopy { get; set; } = true;
        public bool AllowModify { get; set; } = true;
        public bool AllowAnnotate { get; set; } = true;
        public bool AllowFillForm { get; set; } = true;
    }

    /// <summary>
    /// PDF 加密 / 移除密码转换器（基于本地 PDF 引擎）。
    ///
    /// 实测约束：
    ///   - 引擎不允许打开密码与编辑密码相同，故两种密码必须不同；
    ///   - 编辑密码传空字符串会导致打开密码失效（空密码即可打开），因此「不设编辑密码」时
    ///     改用只设置打开密码的接口，此模式下加密算法固定为 RC4-128；
    ///   - 移除密码需要编辑密码；若原文件未设置编辑密码，传打开密码或空即可解除。
    /// </summary>
    public class PdfEncryptConverter : IDisposable
    {
        private const string EncryptMarker = "/Encrypt";

        /// <summary>判断 PDF 是否已加密（扫描文件中的加密字典标记）。</summary>
        public static bool IsEncrypted(string pdfPath)
        {
            try
            {
                if (!File.Exists(pdfPath)) return false;
                using (var fs = new FileStream(pdfPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(fs, Encoding.GetEncoding(28591), false))
                {
                    var buffer = new char[64 * 1024];
                    int read;
                    var tail = new StringBuilder();
                    while ((read = reader.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        var chunk = new string(buffer, 0, read);
                        // 保留上一块尾部，避免标记被切断
                        var probe = tail.Length > 0 ? tail + chunk : chunk;
                        if (probe.IndexOf(EncryptMarker, StringComparison.Ordinal) >= 0) return true;
                        tail.Clear();
                        if (probe.Length >= 32) tail.Append(probe.Substring(probe.Length - 32));
                    }
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 加密单个 PDF，返回输出文件路径。
        /// </summary>
        public string Encrypt(string sourcePath, string outputDir, PdfEncryptOptions options)
        {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("PDF 文件不存在", sourcePath);
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (string.IsNullOrEmpty(options.UserPassword))
                throw new ArgumentException("打开密码不能为空", nameof(options));

            var owner = options.OwnerPassword ?? string.Empty;
            if (owner.Length > 0 && string.Equals(owner, options.UserPassword, StringComparison.Ordinal))
                throw new ArgumentException("编辑密码不能与打开密码相同");

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            var outputFile = UniquePath(outputDir,
                Path.GetFileNameWithoutExtension(sourcePath) + "_已加密", ".pdf");

            using (var doc = new SpirePdfDocument())
            {
                doc.LoadFromFile(sourcePath);

                if (owner.Length == 0)
                {
                    // 不设置编辑密码：只设置打开密码。
                    // 引擎不提供「仅打开密码」的非过时接口，此处沿用 Security.Encrypt(密码)。
#pragma warning disable 618
                    doc.Security.Encrypt(options.UserPassword);
#pragma warning restore 618
                }
                else
                {
                    var policy = new PdfPasswordSecurityPolicy(options.UserPassword, owner)
                    {
                        EncryptionAlgorithm = options.Algorithm,
                        DocumentPrivilege = BuildPrivilege(options),
                        EncryptMetadata = true
                    };
                    doc.Encrypt(policy);
                }

                doc.SaveToFile(outputFile);
            }

            return outputFile;
        }

        /// <summary>
        /// 移除 PDF 密码（解密），返回输出文件路径。
        /// 需要编辑密码；若原文件未设置编辑密码，传打开密码或留空均可。
        /// </summary>
        public string Decrypt(string sourcePath, string outputDir, string password)
        {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("PDF 文件不存在", sourcePath);
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            var outputFile = UniquePath(outputDir,
                Path.GetFileNameWithoutExtension(sourcePath) + "_已解密", ".pdf");

            // 先尝试以「编辑密码」解密；失败再尝试无参解密（针对未设置编辑密码的文件）
            var errors = new List<Exception>();

            if (!string.IsNullOrEmpty(password) && TryDecrypt(sourcePath, outputFile, password, errors))
                return outputFile;

            if (TryDecrypt(sourcePath, outputFile, null, errors))
                return outputFile;

            var inner = errors.Count > 0 ? errors[errors.Count - 1] : null;
            throw new InvalidOperationException("密码不正确，无法移除密码（需要编辑密码）", inner);
        }

        /// <summary>以指定密码加载并解除加密；成功返回 true。</summary>
        private static bool TryDecrypt(string sourcePath, string outputFile, string password, List<Exception> errors)
        {
            try
            {
                using (var doc = new SpirePdfDocument())
                {
                    if (!string.IsNullOrEmpty(password))
                        doc.LoadFromFile(sourcePath, password);
                    else
                        doc.LoadFromFile(sourcePath);

                    if (!string.IsNullOrEmpty(password))
                        doc.Decrypt(password);
                    else
                        doc.Decrypt();

                    doc.SaveToFile(outputFile);
                }
                return true;
            }
            catch (Exception ex)
            {
                errors.Add(ex);
                SafeDelete(outputFile);
                return false;
            }
        }

        /// <summary>删除半成品文件。</summary>
        private static void SafeDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }

        /// <summary>按预设/自定义选项构建权限对象。</summary>
        private static PdfDocumentPrivilege BuildPrivilege(PdfEncryptOptions options)
        {
            switch (options.Preset)
            {
                case PdfEncryptPreset.ViewOnly:
                {
                    var p = PdfDocumentPrivilege.ForbidAll;
                    return p;
                }
                case PdfEncryptPreset.AllowPrint:
                {
                    var p = PdfDocumentPrivilege.ForbidAll;
                    p.AllowPrint = true;
                    p.AllowDegradedPrinting = true;
                    return p;
                }
                case PdfEncryptPreset.Custom:
                {
                    var p = PdfDocumentPrivilege.ForbidAll;
                    p.AllowPrint = options.AllowPrint;
                    p.AllowDegradedPrinting = options.AllowPrint;
                    p.AllowContentCopying = options.AllowCopy;
                    p.AllowCopyContentAccessibility = options.AllowCopy;
                    p.AllowModifyContents = options.AllowModify;
                    p.AllowModifyAnnotations = options.AllowAnnotate;
                    p.AllowFillFormFields = options.AllowFillForm;
                    return p;
                }
                default:
                    return PdfDocumentPrivilege.AllowAll;
            }
        }

        /// <summary>生成不重名的输出路径。</summary>
        private static string UniquePath(string dir, string baseName, string ext)
        {
            var path = Path.Combine(dir, baseName + ext);
            int seq = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(dir, $"{baseName}_{seq++}{ext}");
            }
            return path;
        }

        /// <summary>读取 PDF 页数（加密文件用 PdfPig 无法解析时返回 0）。</summary>
        public static int TryGetPageCount(string pdfPath)
        {
            try
            {
                using (var doc = UglyToad.PdfPig.PdfDocument.Open(pdfPath))
                    return doc.NumberOfPages;
            }
            catch { return 0; }
        }

        public void Dispose() { }
    }
}
