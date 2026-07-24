using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Spire.Doc;

namespace DocToPdfTool.Utils
{
    public class WordToPdfConverterSpire : IDisposable
    {
        private bool _disposed;
        [DllImport("psapi.dll")] private static extern bool EmptyWorkingSet(IntPtr hProcess);
        public void Convert(string wordPath, string outputDir)
        {
            if (string.IsNullOrEmpty(wordPath)) throw new ArgumentNullException(nameof(wordPath));
            if (!File.Exists(wordPath)) throw new FileNotFoundException("Word 文件不存在", wordPath);
            var fn = Path.GetFileNameWithoutExtension(wordPath);
            var savePath = Path.Combine(outputDir, fn + ".pdf");
            using (var doc = new Document()) { doc.LoadFromFile(wordPath); doc.SaveToFile(savePath, FileFormat.PDF); }
            ReleaseMemory();
        }
        private static void ReleaseMemory()
        { try { System.Runtime.GCSettings.LargeObjectHeapCompactionMode = System.Runtime.GCLargeObjectHeapCompactionMode.CompactOnce; GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); EmptyWorkingSet(Process.GetCurrentProcess().Handle); } catch { } }
        public void Dispose() { if (!_disposed) { _disposed = true; } }
    }
}