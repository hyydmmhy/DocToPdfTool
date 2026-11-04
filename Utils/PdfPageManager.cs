using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Spire.Pdf;

namespace DocToPdfTool.Utils
{
    /// <summary>
    /// PDF 页面管理器（按指定顺序保留页面、设置旋转）。
    ///
    /// 引擎细节：
    ///   - PdfPageCollection.ReArrange 的数组是「0 基源页索引」，且长度必须等于当前页数，
    ///     因此先按目标顺序重排、把要删除的页追加到尾部，再从尾部逐页删除；
    ///   - 旋转取值 0/90/180/270，写入页面 Rotation 属性。
    /// </summary>
    public class PdfPageManager : IDisposable
    {
        /// <summary>
        /// 按指定顺序导出页面。
        /// </summary>
        /// <param name="sourcePath">源 PDF</param>
        /// <param name="outputDir">输出目录</param>
        /// <param name="keepPages">保留的源页码（1 基），数组顺序即输出顺序</param>
        /// <param name="rotations">与 keepPages 对应的旋转角度（0/90/180/270），可为 null</param>
        /// <param name="suffix">输出文件名后缀</param>
        public string Save(string sourcePath, string outputDir,
            IList<int> keepPages, IList<int> rotations = null, string suffix = "页面管理")
        {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("PDF 文件不存在", sourcePath);
            if (keepPages == null || keepPages.Count == 0)
                throw new ArgumentException("请至少保留一页");
            if (PdfEncryptConverter.IsEncrypted(sourcePath))
                throw new InvalidOperationException("文件已加密，请先用「PDF加密 → 移除密码」解除后再处理");

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            var outputFile = UniquePath(outputDir,
                Path.GetFileNameWithoutExtension(sourcePath) + "_" + suffix);

            using (var doc = new PdfDocument())
            {
                doc.LoadFromFile(sourcePath);
                int total = doc.Pages.Count;

                if (keepPages.Any(p => p < 1 || p > total))
                    throw new ArgumentOutOfRangeException(nameof(keepPages), "页码超出文件页数范围");

                // 目标顺序（0 基），被删除页追加到末尾以保证数组长度等于当前页数
                var order = keepPages.Select(p => p - 1).ToList();
                for (int i = 0; i < total; i++)
                {
                    if (!order.Contains(i)) order.Add(i);
                }

                if (order.Count != total)
                    throw new InvalidOperationException("页面顺序数组长度与原页数不一致");

                doc.Pages.ReArrange(order.ToArray());

                // 从尾部删除被丢弃的页
                for (int i = total - 1; i >= keepPages.Count; i--)
                    doc.Pages.RemoveAt(i);

                // 逐页应用旋转
                for (int i = 0; i < keepPages.Count; i++)
                {
                    int angle = rotations != null && i < rotations.Count ? rotations[i] : 0;
                    doc.Pages[i].Rotation = ToRotateAngle(angle);
                }

                doc.SaveToFile(outputFile);
            }

            return outputFile;
        }

        private static PdfPageRotateAngle ToRotateAngle(int degrees)
        {
            switch (((degrees % 360) + 360) % 360)
            {
                case 90: return PdfPageRotateAngle.RotateAngle90;
                case 180: return PdfPageRotateAngle.RotateAngle180;
                case 270: return PdfPageRotateAngle.RotateAngle270;
                default: return PdfPageRotateAngle.RotateAngle0;
            }
        }

        private static string UniquePath(string dir, string baseName)
        {
            var path = Path.Combine(dir, baseName + ".pdf");
            int seq = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(dir, $"{baseName}_{seq++}.pdf");
            }
            return path;
        }

        public void Dispose() { }
    }
}
