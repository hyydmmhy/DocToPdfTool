using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32;

namespace DocToPdfTool.Utils
{
    public class PdfToWordConverter : IDisposable
    {
        private dynamic _wordApp;
        private bool _disposed;

        // ==================== 窗口守卫 ====================

        private volatile bool _windowGuardRunning;
        private volatile uint _guardedProcessId;
        private Thread _windowGuardThread;
        private Thread _hookPumpThread;
        private uint _hookPumpThreadId;
        private IntPtr _winEventHook = IntPtr.Zero;
        private WinEventDelegate _winEventDelegate;
        private ManagementEventWatcher _wmiWatcher;

        // ==================== P/Invoke ====================

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern bool SetWindowDisplayAffinity(IntPtr hWnd, uint dwAffinity);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostThreadMessage(uint threadId, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax,
            IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc,
            uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild,
            uint dwEventThread, uint dwmsEventTime);

        private const int SW_HIDE = 0;
        private const int GWL_STYLE = -16;
        private const int GWL_EXSTYLE = -20;
        private const int WS_VISIBLE = 0x10000000;
        private const int WS_EX_APPWINDOW = 0x00040000;
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const uint WDA_MONITOR = 1;
        private const uint EVENT_OBJECT_SHOW = 0x8002;
        private const uint WINEVENT_OUTOFCONTEXT = 0;
        private const uint WM_QUIT = 0x0012;

        // ==================== 公开方法 ====================

        /// <summary>必须在 STA 线程上调用。创建 Word 实例并启动窗口守卫。</summary>
        public void Initialize()
        {
            StartWmiProcessWatcher();

            _wordApp = CreateGenuineWordApplication();
            if (_wordApp == null)
            {
                _windowGuardRunning = false;
                throw new Exception(
                    "未能定位到真正的 Microsoft Word。\n\n" +
                    "常见解决办法：\n" +
                    "1) 打开微软 Word，在 文件->选项->信任中心 里确认没有被 WPS 覆盖注册；\n" +
                    "2) 到 控制面板->程序 里对 Microsoft Office 做一次\"快速修复/联机修复\"；\n" +
                    "3) 或者先卸载/更新 WPS 到最新版（新版一般不会抢注 Word.Application）。");
            }

            _wordApp.Visible = false;
            _wordApp.DisplayAlerts = 0;

            DisablePdfConversionWarningDialog();
            StartWordWindowGuard();
        }

        /// <summary>转换单个 PDF 文件到 DOCX。</summary>
        public void Convert(string pdfPath, string outputDir)
        {
            if (string.IsNullOrEmpty(pdfPath)) throw new ArgumentNullException(nameof(pdfPath));
            if (!File.Exists(pdfPath)) throw new FileNotFoundException("PDF 文件不存在", pdfPath);

            if (_wordApp == null)
                throw new InvalidOperationException("请先调用 Initialize()");

            var fileName = Path.GetFileNameWithoutExtension(pdfPath);
            var savePath = Path.Combine(outputDir, fileName + ".docx");

            ConvertPdfToWordViaWordInterop(pdfPath, savePath);
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                StopWordWindowGuard();
                if (_wordApp != null)
                {
                    try { _wordApp.Quit(); } catch { }
                    try { Marshal.ReleaseComObject((object)_wordApp); } catch { }
                    _wordApp = null;
                }
                _disposed = true;
            }
        }

        // ==================== 核心转换 ====================

        private void ConvertPdfToWordViaWordInterop(string pdfPath, string savePath)
        {
            UnblockFile(pdfPath);

            dynamic doc = _wordApp.Documents.Open(pdfPath);
            if (doc == null)
            {
                throw new InvalidOperationException(
                    "Word 未能成功打开并转换该 PDF（Documents.Open 返回了空值）。" +
                    "常见原因：1) PDF 被加密/加了密码；2) PDF 已损坏；3) Word 转换确认弹窗未被成功关闭。");
            }

            try
            {
                // 12 = wdFormatXMLDocument (docx)
                doc.SaveAs2(savePath, 12);
            }
            finally
            {
                doc.Close(false);
                Marshal.ReleaseComObject((object)doc);
            }
        }

        // ==================== 创建真正的 Word 实例（三级回退） ====================

        private dynamic CreateGenuineWordApplication()
        {
            dynamic hklmApp = CreateWordApplicationFromHklmClsid();
            if (hklmApp != null) return hklmApp;

            dynamic exeApp = CreateWordApplicationByLaunchingExe();
            if (exeApp != null) return exeApp;

            string[] candidateProgIds = new[]
            {
                "Word.Application.16",
                "Word.Application.15",
                "Word.Application.14",
                "Word.Application.12",
                "Word.Application"
            };

            foreach (string progId in candidateProgIds)
            {
                Type t = Type.GetTypeFromProgID(progId);
                if (t == null) continue;

                dynamic app = null;
                try
                {
                    app = Activator.CreateInstance(t);
                    app.Visible = false;
                    string path = (string)app.Path;
                    if (!string.IsNullOrEmpty(path) &&
                        path.IndexOf("Microsoft", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return app;
                    }
                    try { app.Quit(); } catch { }
                    try { Marshal.ReleaseComObject((object)app); } catch { }
                }
                catch
                {
                    if (app != null)
                    {
                        try { Marshal.ReleaseComObject((object)app); } catch { }
                    }
                }
            }

            return null;
        }

        private dynamic CreateWordApplicationFromHklmClsid()
        {
            string[] clsidRegPaths = new[]
            {
                @"SOFTWARE\Classes\Word.Application\CLSID",
                @"SOFTWARE\Wow6432Node\Classes\Word.Application\CLSID"
            };

            foreach (string regPath in clsidRegPaths)
            {
                Guid clsid;
                if (!TryGetClsidFromHklm(regPath, out clsid)) continue;

                dynamic app = null;
                try
                {
                    Type t = Type.GetTypeFromCLSID(clsid);
                    if (t == null) continue;

                    app = Activator.CreateInstance(t);
                    app.Visible = false;
                    string path = (string)app.Path;
                    if (!string.IsNullOrEmpty(path) &&
                        path.IndexOf("Microsoft", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return app;
                    }

                    try { app.Quit(); } catch { }
                    try { Marshal.ReleaseComObject((object)app); } catch { }
                }
                catch
                {
                    if (app != null)
                    {
                        try { Marshal.ReleaseComObject((object)app); } catch { }
                    }
                }
            }

            return null;
        }

        private string FindGenuineWinwordExePath()
        {
            string[] appPathsKeys = new[]
            {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE",
                @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE"
            };

            foreach (string keyPath in appPathsKeys)
            {
                try
                {
                    using (RegistryKey key = Registry.LocalMachine.OpenSubKey(keyPath))
                    {
                        if (key == null) continue;
                        string exePath = key.GetValue(null) as string;
                        if (!string.IsNullOrEmpty(exePath) && File.Exists(exePath))
                            return exePath;
                    }
                }
                catch { }
            }

            return null;
        }

        private dynamic CreateWordApplicationByLaunchingExe()
        {
            string exePath = FindGenuineWinwordExePath();
            if (string.IsNullOrEmpty(exePath)) return null;

            Process proc = null;
            Thread startupGuard = null;
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = "/automation",
                    WindowStyle = ProcessWindowStyle.Hidden,
                    CreateNoWindow = true,
                    UseShellExecute = false
                };
                proc = Process.Start(psi);
                int pid = proc.Id;

                startupGuard = StartStartupWindowGuard(pid);

                dynamic app = null;
                for (int i = 0; i < 20; i++)
                {
                    Thread.Sleep(300);
                    try
                    {
                        object obj = Marshal.GetActiveObject("Word.Application");
                        app = obj;
                        if (app != null) break;
                    }
                    catch (COMException) { }
                }

                if (app == null)
                {
                    try { if (proc != null && !proc.HasExited) proc.Kill(); } catch { }
                    return null;
                }

                try
                {
                    string path = (string)app.Path;
                    if (string.IsNullOrEmpty(path) || path.IndexOf("Microsoft", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        try { app.Quit(); } catch { }
                        try { Marshal.ReleaseComObject((object)app); } catch { }
                        return null;
                    }
                }
                catch { return null; }

                return app;
            }
            catch
            {
                try { if (proc != null && !proc.HasExited) proc.Kill(); } catch { }
                return null;
            }
            finally
            {
                if (startupGuard != null)
                {
                    try { startupGuard.Abort(); } catch { }
                }
            }
        }

        private Thread StartStartupWindowGuard(int pid)
        {
            Thread guard = new Thread(() =>
            {
                try
                {
                    for (int j = 0; j < 30; j++)
                    {
                        EnumWindows((hWnd, lParam) =>
                        {
                            uint windowPid;
                            GetWindowThreadProcessId(hWnd, out windowPid);
                            if (windowPid == pid && IsWindowVisible(hWnd))
                                ForceHideWindow(hWnd);
                            return true;
                        }, IntPtr.Zero);
                        Thread.Sleep(100);
                    }
                }
                catch { }
            });
            guard.IsBackground = true;
            guard.Start();
            return guard;
        }

        private bool TryGetClsidFromHklm(string regPath, out Guid clsid)
        {
            clsid = Guid.Empty;
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(regPath))
                {
                    if (key == null) return false;
                    string clsidStr = key.GetValue(null) as string;
                    if (string.IsNullOrEmpty(clsidStr)) return false;
                    clsid = new Guid(clsidStr);
                    return true;
                }
            }
            catch
            {
                clsid = Guid.Empty;
                return false;
            }
        }

        // ==================== 注册表：禁用 PDF 转换警告 ====================

        private void DisablePdfConversionWarningDialog()
        {
            try
            {
                string version = (string)_wordApp.Version;
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(
                    string.Format(@"SOFTWARE\Microsoft\Office\{0}\Word\Options", version)))
                {
                    if (key != null)
                        key.SetValue("DisableConvertPDFWarning", 1, RegistryValueKind.DWord);
                }
            }
            catch { }
        }

        // ==================== 解除 Mark of the Web 锁定 ====================

        private void UnblockFile(string filePath)
        {
            try
            {
                string zoneIdentifierPath = filePath + ":Zone.Identifier";
                if (File.Exists(zoneIdentifierPath))
                    File.Delete(zoneIdentifierPath);
            }
            catch { }
        }

        // ==================== 窗口守卫 ====================

        private void StartWmiProcessWatcher()
        {
            try
            {
                WqlEventQuery query = new WqlEventQuery(
                    "Win32_ProcessStartTrace",
                    "ProcessName = 'WINWORD.EXE'");
                _wmiWatcher = new ManagementEventWatcher(query);
                _wmiWatcher.EventArrived += OnWinwordProcessStarted;
                _wmiWatcher.Start();
            }
            catch
            {
                _wmiWatcher = null;
            }
        }

        private void OnWinwordProcessStarted(object sender, EventArrivedEventArgs e)
        {
            try
            {
                if (_windowGuardRunning) return;

                uint pid = System.Convert.ToUInt32(e.NewEvent.Properties["ProcessID"].Value);
                _guardedProcessId = pid;
                _windowGuardRunning = true;

                _windowGuardThread = new Thread(() =>
                {
                    while (_windowGuardRunning)
                    {
                        HideAllWindowsOfProcess(_guardedProcessId);
                        Thread.Sleep(80);
                    }
                });
                _windowGuardThread.IsBackground = true;
                _windowGuardThread.Start();
            }
            catch { }
        }

        private void StartWordWindowGuard()
        {
            try
            {
                IntPtr hwnd = new IntPtr(System.Convert.ToInt64(_wordApp.Hwnd));
                uint pid;
                GetWindowThreadProcessId(hwnd, out pid);
                if (pid == 0) return;

                _guardedProcessId = pid;
                _windowGuardRunning = true;

                HideAllWindowsOfProcess(pid);

                _winEventDelegate = OnWordWindowShown;
                _hookPumpThread = new Thread(() =>
                {
                    _hookPumpThreadId = GetCurrentThreadId();
                    _winEventHook = SetWinEventHook(
                        EVENT_OBJECT_SHOW, EVENT_OBJECT_SHOW,
                        IntPtr.Zero, _winEventDelegate,
                        pid, 0, WINEVENT_OUTOFCONTEXT);
                    System.Windows.Forms.Application.Run();
                });
                _hookPumpThread.SetApartmentState(ApartmentState.STA);
                _hookPumpThread.IsBackground = true;
                _hookPumpThread.Start();

                if (_windowGuardThread == null || !_windowGuardThread.IsAlive)
                {
                    _windowGuardThread = new Thread(() =>
                    {
                        while (_windowGuardRunning)
                        {
                            HideAllWindowsOfProcess(_guardedProcessId);
                            Thread.Sleep(80);
                        }
                    });
                    _windowGuardThread.IsBackground = true;
                    _windowGuardThread.Start();
                }
            }
            catch { }
        }

        private void OnWordWindowShown(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild,
            uint dwEventThread, uint dwmsEventTime)
        {
            if (!_windowGuardRunning) return;
            if (idObject != 0) return;
            ForceHideWindow(hwnd);
        }

        private void ForceHideWindow(IntPtr hWnd)
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

                SetWindowDisplayAffinity(hWnd, WDA_MONITOR);
            }
            catch { }
        }

        private void StopWordWindowGuard()
        {
            _windowGuardRunning = false;

            if (_hookPumpThread != null && _hookPumpThread.IsAlive)
            {
                try { PostThreadMessage(_hookPumpThreadId, WM_QUIT, IntPtr.Zero, IntPtr.Zero); } catch { }
                try { _hookPumpThread.Join(1000); } catch { }
                _hookPumpThread = null;
            }

            if (_winEventHook != IntPtr.Zero)
            {
                try { UnhookWinEvent(_winEventHook); } catch { }
                _winEventHook = IntPtr.Zero;
            }

            if (_windowGuardThread != null)
            {
                try { _windowGuardThread.Join(500); } catch { }
                _windowGuardThread = null;
            }

            if (_wmiWatcher != null)
            {
                try { _wmiWatcher.Stop(); } catch { }
                try { _wmiWatcher.Dispose(); } catch { }
                _wmiWatcher = null;
            }
        }

        private void HideAllWindowsOfProcess(uint processId)
        {
            try
            {
                EnumWindows((hWnd, lParam) =>
                {
                    uint pid;
                    GetWindowThreadProcessId(hWnd, out pid);
                    if (pid == processId && IsWindowVisible(hWnd))
                        ForceHideWindow(hWnd);
                    return true;
                }, IntPtr.Zero);
            }
            catch { }
        }
    }
}