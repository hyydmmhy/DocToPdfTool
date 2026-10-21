using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace DocToPdfTool.Utils
{
    /// <summary>
    /// 应用配置：读写 %APPDATA%\DocToPdfTool\config.ini，提供默认保存目录。
    /// 每个转换功能的输出目录 = 默认保存目录 + 功能名（与左侧菜单一致）。
    /// </summary>
    public static class AppConfig
    {
        private static string ConfigDir => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "DocToPdfTool");

        private static string ConfigPath => Path.Combine(ConfigDir, "config.ini");

        /// <summary>获取默认保存目录（优先取配置，未配置时为下载目录）</summary>
        public static string GetDefaultSaveDir()
        {
            var dir = ReadValue("Settings", "OutputDir");
            if (!string.IsNullOrWhiteSpace(dir))
                return dir;
            return GetDownloadsFolder();
        }

        /// <summary>保存默认保存目录到配置文件，永久生效</summary>
        public static void SetDefaultSaveDir(string dir)
        {
            WriteValue("Settings", "OutputDir", dir);
        }

        /// <summary>获取某功能的输出目录（默认保存目录下以功能名命名的子目录），不存在则创建</summary>
        public static string GetFeatureOutputDir(string featureName)
        {
            var dir = Path.Combine(GetDefaultSaveDir(), featureName);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            return dir;
        }

        private static readonly Guid FOLDERID_Downloads =
            new Guid("{374DE290-123F-4565-9164-39C4925E467B}");

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHGetKnownFolderPath(
            [MarshalAs(UnmanagedType.LPStruct)] Guid rfid,
            uint dwFlags, IntPtr hToken, out IntPtr pszPath);

        /// <summary>获取当前用户的下载目录（失败时回退到 用户目录\Downloads）</summary>
        public static string GetDownloadsFolder()
        {
            try
            {
                IntPtr pathPtr;
                int hr = SHGetKnownFolderPath(FOLDERID_Downloads, 0, IntPtr.Zero, out pathPtr);
                if (hr == 0 && pathPtr != IntPtr.Zero)
                {
                    var path = Marshal.PtrToStringUni(pathPtr);
                    Marshal.FreeCoTaskMem(pathPtr);
                    if (!string.IsNullOrEmpty(path))
                        return path;
                }
            }
            catch { }
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
        }

        private static string ReadValue(string section, string key)
        {
            try
            {
                if (!File.Exists(ConfigPath))
                    return null;

                bool inSection = false;
                foreach (var line in File.ReadAllLines(ConfigPath))
                {
                    var trimmed = line.Trim();
                    if (trimmed.StartsWith("[") && trimmed.EndsWith("]"))
                    {
                        inSection = string.Equals(trimmed, "[" + section + "]",
                            StringComparison.OrdinalIgnoreCase);
                        continue;
                    }

                    if (inSection)
                    {
                        int eq = trimmed.IndexOf('=');
                        if (eq > 0 && string.Equals(trimmed.Substring(0, eq).Trim(), key,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            return trimmed.Substring(eq + 1).Trim();
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        private static void WriteValue(string section, string key, string value)
        {
            try
            {
                Directory.CreateDirectory(ConfigDir);

                var lines = new List<string>();
                if (File.Exists(ConfigPath))
                    lines.AddRange(File.ReadAllLines(ConfigPath));

                // 定位节
                int sectionStart = -1;
                for (int i = 0; i < lines.Count; i++)
                {
                    if (string.Equals(lines[i].Trim(), "[" + section + "]",
                        StringComparison.OrdinalIgnoreCase))
                    {
                        sectionStart = i;
                        break;
                    }
                }

                if (sectionStart < 0)
                {
                    // 节不存在：追加到文件末尾
                    if (lines.Count > 0 && !string.IsNullOrWhiteSpace(lines[lines.Count - 1]))
                        lines.Add("");
                    lines.Add("[" + section + "]");
                    lines.Add($"{key}={value}");
                }
                else
                {
                    // 节内查找键：记录节结束位置用于插入
                    int keyIndex = -1;
                    int insertIndex = lines.Count;
                    for (int i = sectionStart + 1; i < lines.Count; i++)
                    {
                        var t = lines[i].Trim();
                        if (t.StartsWith("["))
                        {
                            insertIndex = i;
                            break;
                        }
                        int eq = t.IndexOf('=');
                        if (eq > 0 && string.Equals(t.Substring(0, eq).Trim(), key,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            keyIndex = i;
                            break;
                        }
                    }

                    if (keyIndex >= 0)
                        lines[keyIndex] = $"{key}={value}";
                    else
                        lines.Insert(insertIndex, $"{key}={value}");
                }

                File.WriteAllLines(ConfigPath, lines);
            }
            catch { }
        }
    }
}