using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace DocToPdfTool.Utils
{
    public class OfficePdfConverter : IDisposable
    {
        public static readonly HashSet<string> SupportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"
        };

        private dynamic _officeApp;
        private bool _disposed;

        public void Convert(string sourceFile, string outputFile)
        {
            if (sourceFile == null) throw new ArgumentNullException(nameof(sourceFile));
            if (!File.Exists(sourceFile)) throw new FileNotFoundException("源文件不存在", sourceFile);

            if (outputFile == null)
                outputFile = Path.ChangeExtension(sourceFile, "pdf");

            var extension = Path.GetExtension(sourceFile).ToLowerInvariant();

            if (extension == ".doc" || extension == ".docx")
            {
                ConvertWord(sourceFile, outputFile);
            }
            else if (extension == ".xls" || extension == ".xlsx")
            {
                ConvertExcel(sourceFile, outputFile);
            }
            else if (extension == ".ppt" || extension == ".pptx")
            {
                ConvertPpt(sourceFile, outputFile);
            }
            else
            {
                throw new NotSupportedException($"不支持的文件格式: {extension}");
            }
        }

        private void ConvertWord(string sourceFile, string outputFile)
        {
            EnsureOfficeApp("Word.Application");
            _officeApp.DisplayAlerts = false;

            dynamic doc = _officeApp.Documents.Open(sourceFile);
            try
            {
                doc.Repaginate();
                doc.ExportAsFixedFormat(outputFile, 17);
            }
            finally
            {
                try { doc.Close(); } catch { }
                try { Marshal.ReleaseComObject((object)doc); } catch { }
                doc = null;
            }
        }

        private void ConvertExcel(string sourceFile, string outputFile)
        {
            EnsureOfficeApp("Excel.Application");
            var appType = _officeApp.GetType();
            appType.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, _officeApp, new object[] { 0 });
            appType.InvokeMember("ScreenUpdating", BindingFlags.SetProperty, null, _officeApp, new object[] { 0 });
            appType.InvokeMember("Visible", BindingFlags.SetProperty, null, _officeApp, new object[] { 0 });

            RunWithGuard("EXCEL", () =>
            {
                dynamic xls = _officeApp.Application.Workbooks.Open(sourceFile);
                try
                {
                    foreach (dynamic sheet in xls.Worksheets)
                    {
                        dynamic usedRange = null;
                        dynamic pageSetup = null;
                        try
                        {
                            sheet.Activate();
                            usedRange = sheet.UsedRange;
                            usedRange.Columns.AutoFit();

                            pageSetup = sheet.PageSetup;
                            var psType = pageSetup.GetType();
                            psType.InvokeMember("Zoom", BindingFlags.SetProperty, null, pageSetup, new object[] { false });
                            psType.InvokeMember("FitToPagesWide", BindingFlags.SetProperty, null, pageSetup, new object[] { 1 });
                        }
                        finally
                        {
                            if (pageSetup != null) try { Marshal.ReleaseComObject((object)pageSetup); } catch { }
                            if (usedRange != null) try { Marshal.ReleaseComObject((object)usedRange); } catch { }
                            try { Marshal.ReleaseComObject((object)sheet); } catch { }
                        }
                    }

                    dynamic firstSheet = xls.Worksheets[1];
                    try
                    {
                        firstSheet.Activate();
                        xls.ExportAsFixedFormat(0, outputFile, 0, true, false);
                        xls.Saved = true;
                    }
                    finally
                    {
                        try { Marshal.ReleaseComObject((object)firstSheet); } catch { }
                    }
                }
                finally
                {
                    try { xls.Close(); } catch { }
                    try { Marshal.ReleaseComObject((object)xls); } catch { }
                    xls = null;
                }
            });
        }

        private void ConvertPpt(string sourceFile, string outputFile)
        {
            Exception threadError = null;
            var staThread = new Thread(() =>
            {
                try
                {
                    var type = Type.GetTypeFromProgID("PowerPoint.Application");
                    if (type == null)
                        throw new Exception("未检测到 Microsoft Office，请先安装 Office PowerPoint。");

                    dynamic app = Activator.CreateInstance(type);
                    try
                    {
                        var t = app.GetType();
                        t.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, app, new object[] { 1 });

                        RunWithGuard("POWERPNT", () =>
                        {
                            dynamic ppt = app.Presentations.Open(sourceFile, -1, 0, 0);
                            try
                            {
                                ppt.SaveAs(outputFile, 32, -1);
                                ppt.Saved = true;
                            }
                            finally
                            {
                                try { ppt.Close(); } catch { }
                                try { Marshal.ReleaseComObject((object)ppt); } catch { }
                                ppt = null;
                            }
                        });
                    }
                    finally
                    {
                        try { app.Quit(); } catch { }
                        try { Marshal.ReleaseComObject((object)app); } catch { }
                    }
                }
                catch (Exception ex) { threadError = ex; }
            });
            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join(TimeSpan.FromSeconds(120));
            if (threadError != null) throw threadError;
        }

        // ==================== 窗口守卫（WinEventHook + 轮询） ====================

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc,
            WinEventProc lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool SetWindowDisplayAffinity(IntPtr hWnd, uint dwAffinity);

        [DllImport("user32.dll")]
        private static extern bool PostThreadMessage(uint threadId, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        private delegate void WinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
            int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const int SW_HIDE = 0;
        private const int GWL_STYLE = -16;
        private const int GWL_EXSTYLE = -20;
        private const int WS_VISIBLE = 0x10000000;
        private const int WS_EX_APPWINDOW = 0x00040000;
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const uint EVENT_OBJECT_CREATE = 0x8000;
        private const uint WINEVENT_OUTOFCONTEXT = 0;
        private const uint WM_QUIT = 0x0012;
        private const uint WDA_MONITOR = 1;

        private static void ForceHideWindow(IntPtr hWnd)
        {
            try
            {
                ShowWindow(hWnd, SW_HIDE);
                int style = GetWindowLong(hWnd, GWL_STYLE);
                if ((style & WS_VISIBLE) != 0)
                    SetWindowLong(hWnd, GWL_STYLE, style & ~WS_VISIBLE);
                int exStyle = GetWindowLong(hWnd, GWL_EXSTYLE);
                int newExStyle = exStyle;
                if ((newExStyle & WS_EX_APPWINDOW) != 0)
                    newExStyle &= ~WS_EX_APPWINDOW;
                if ((newExStyle & WS_EX_TOOLWINDOW) == 0)
                    newExStyle |= WS_EX_TOOLWINDOW;
                if (newExStyle != exStyle)
                    SetWindowLong(hWnd, GWL_EXSTYLE, newExStyle);

                // 阻止 DWM 渲染窗口内容到屏幕（Windows 8+），窗口即使显示也是空白
                SetWindowDisplayAffinity(hWnd, WDA_MONITOR);
            }
            catch { }
        }

        private static HashSet<uint> GetProcessPids(string processName)
        {
            return new HashSet<uint>(Process.GetProcessesByName(processName).Select(p => (uint)p.Id));
        }

        private static void HideAllWindows(HashSet<uint> pids)
        {
            if (pids == null || pids.Count == 0) return;
            EnumWindows((hWnd, lParam) =>
            {
                uint pid;
                GetWindowThreadProcessId(hWnd, out pid);
                if (pids.Contains(pid))
                    ForceHideWindow(hWnd);
                return true;
            }, IntPtr.Zero);
        }

        /// <summary>运行 action 期间启动窗口守卫（WinEventHook + 轮询），结束自动关闭。</summary>
        private static void RunWithGuard(string processName, Action action)
        {
            // 共享 PID 列表，由轮询线程持续更新，WinEventHook 回调读取
            var sharedPids = GetProcessPids(processName);
            if (sharedPids.Count == 0) { action(); return; }

            // 初始隐藏所有已有窗口
            HideAllWindows(sharedPids);

            var guardRunning = true;
            IntPtr hook = IntPtr.Zero;
            Thread pollThread = null;
            uint hookThreadId = 0;

            // WinEventHook 需要 STA 线程泵消息
            Thread hookThread = new Thread(() =>
            {
                hookThreadId = GetCurrentThreadId();
                hook = SetWinEventHook(EVENT_OBJECT_CREATE, EVENT_OBJECT_CREATE, IntPtr.Zero,
                    (hHook, evt, hwnd, idObj, idChild, dwThread, dwTime) =>
                    {
                        if (!guardRunning) return;
                        if (idObj != 0) return;
                        // 从共享 PID 列表快速判断，O(1) 不阻塞
                        uint pid;
                        GetWindowThreadProcessId(hwnd, out pid);
                        if (sharedPids.Contains(pid))
                            ForceHideWindow(hwnd);
                    }, 0, 0, WINEVENT_OUTOFCONTEXT);

                while (guardRunning && GetMessage(out MSG msg, IntPtr.Zero, 0, 0))
                {
                    if (msg.message == WM_QUIT) break;
                    DispatchMessage(ref msg);
                }
                if (hook != IntPtr.Zero)
                    UnhookWinEvent(hook);
            });
            hookThread.SetApartmentState(ApartmentState.STA);
            hookThread.IsBackground = true; // 关键修复：前台线程会阻止进程退出
            hookThread.Start();

            while (hook == IntPtr.Zero && hookThread.IsAlive)
                Thread.Sleep(50);

            // 轮询线程：5ms 高频轮询，持续更新共享 PID 列表并隐藏窗口
            pollThread = new Thread(() =>
            {
                while (guardRunning)
                {
                    Thread.Sleep(5);
                    sharedPids = GetProcessPids(processName);
                    HideAllWindows(sharedPids);
                    // 按已知窗口类名精确查找发布弹框
                    IntPtr dlg = FindWindow("#32770", null); // 标准对话框类
                    if (dlg != IntPtr.Zero)
                    {
                        uint pid;
                        GetWindowThreadProcessId(dlg, out pid);
                        if (sharedPids.Contains(pid))
                            ForceHideWindow(dlg);
                    }
                }
            });
            pollThread.IsBackground = true;
            pollThread.Start();

            try { action(); }
            finally
            {
                guardRunning = false;
                if (hookThreadId != 0)
                    PostThreadMessage(hookThreadId, WM_QUIT, IntPtr.Zero, IntPtr.Zero);
                try { hookThread.Join(1000); } catch { }
                try { pollThread.Join(500); } catch { }
            }
        }

        // ==================== 消息结构 ====================

        [StructLayout(LayoutKind.Sequential)]
        private struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public int ptX;
            public int ptY;
        }

        [DllImport("user32.dll")]
        private static extern bool GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

        [DllImport("user32.dll")]
        private static extern IntPtr DispatchMessage([In] ref MSG lpMsg);

        [DllImport("user32.dll")]
        private static extern bool PeekMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax, uint wRemoveMsg);

        // ==================== COM 辅助方法 ====================

        private void EnsureOfficeApp(string progId)
        {
            if (_officeApp != null)
            {
                try { _officeApp.Quit(); } catch { }
                try { Marshal.ReleaseComObject((object)_officeApp); } catch { }
                _officeApp = null;
            }

            var type = Type.GetTypeFromProgID(progId);
            if (type == null)
                throw new Exception($"未检测到 Microsoft Office，请先安装 Office。ProgID: {progId}");

            _officeApp = Activator.CreateInstance(type);
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                if (_officeApp != null)
                {
                    try { _officeApp.Quit(); } catch { }
                    try { Marshal.ReleaseComObject((object)_officeApp); } catch { }
                    _officeApp = null;
                }
                _disposed = true;
            }
        }
    }
}