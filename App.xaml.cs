using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Interop;

namespace DocToPdfTool
{
    public partial class App : Application
    {
        private static Mutex _mutex;
        private static EventWaitHandle _eventWaitHandle;

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool AllowSetForegroundWindow(int dwProcessId);

        private const int SW_RESTORE = 9;

        private void OnStartup(object sender, StartupEventArgs e)
        {
            bool createdNew;
            _mutex = new Mutex(true, "DocToPdfTool_SingleInstance", out createdNew);
            _eventWaitHandle = new EventWaitHandle(false, EventResetMode.AutoReset, "DocToPdfTool_ActivateEvent");

            if (!createdNew)
            {
                try
                {
                    var current = Process.GetCurrentProcess();
                    var other = Process.GetProcessesByName(current.ProcessName)
                        .FirstOrDefault(p => p.Id != current.Id);
                    if (other != null)
                        AllowSetForegroundWindow(other.Id);
                }
                catch { }

                _eventWaitHandle.Set();
                Environment.Exit(0);
                return;
            }

            var waitThread = new Thread(() =>
            {
                while (true)
                {
                    _eventWaitHandle.WaitOne();
                    Dispatcher.BeginInvoke((Action)(() =>
                    {
                        var w = Current.MainWindow;
                        if (w == null) return;

                        if (w.WindowState == WindowState.Minimized)
                            w.WindowState = WindowState.Normal;

                        var hwnd = new WindowInteropHelper(w).Handle;
                        if (hwnd != IntPtr.Zero)
                        {
                            if (IsIconic(hwnd))
                                ShowWindow(hwnd, SW_RESTORE);
                            SetForegroundWindow(hwnd);
                        }

                        w.Show();
                        w.Activate();
                        w.Topmost = true;
                        w.Topmost = false;
                    }));
                }
            })
            { IsBackground = true };
            waitThread.Start();

            new MainWindow().Show();
        }
    }
}