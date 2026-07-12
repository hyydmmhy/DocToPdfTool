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
        private static bool _shutdownRequested;
        private Thread _waitThread;

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
            // 全局异常处理
            this.DispatcherUnhandledException += (s, ex) =>
            {
                MessageBox.Show($"程序遇到未处理的异常:\n{ex.Exception.Message}\n\n堆栈:\n{ex.Exception.StackTrace}",
                    "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                ex.Handled = true;
            };

            AppDomain.CurrentDomain.UnhandledException += (s, ex) =>
            {
                var exObj = ex.ExceptionObject as Exception;
                MessageBox.Show($"后台线程异常:\n{exObj?.Message}\n\n堆栈:\n{exObj?.StackTrace}",
                    "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            };

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

            // 注册退出事件，确保线程退出
            this.Exit += (s, e) =>
            {
                _shutdownRequested = true;
                _eventWaitHandle?.Set();

                // 保险丝：正常的 Dispatcher 关闭流程走完后，如果还有游离的
                // 前台线程（比如某个转换器忘了设 IsBackground）导致进程赖着不退，
                // 这里强制终止进程，保证点右上角关闭一定能真正退出。
                Environment.Exit(0);
            };

            _waitThread = new Thread(() =>
            {
                while (!_shutdownRequested)
                {
                    try
                    {
                        // WaitOne 返回 true 表示收到了激活信号，false 是超时
                        bool signaled = _eventWaitHandle.WaitOne(1000);
                        if (_shutdownRequested || !signaled) continue;

                        Current.Dispatcher.BeginInvoke((Action)(() =>
                        {
                            if (_shutdownRequested) return;
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
                    catch { break; }
                }
            })
            { IsBackground = true };
            _waitThread.Start();

            var mainWindow = new MainWindow();
            mainWindow.Closed += (s, e) =>
            {
                _shutdownRequested = true;
                // 强制终止进程——WPF 正常关闭流程可能被 WinRT 残留线程阻塞
                try
                {
                    var proc = Process.GetCurrentProcess();
                    proc.Kill();
                }
                catch { }
            };
            mainWindow.Show();
        }
    }
}