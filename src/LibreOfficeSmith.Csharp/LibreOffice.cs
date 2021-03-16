using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace LibreOfficeSmith.Csharp
{
    public static class LibreOffice
    {
        private static readonly string uninstallRegkey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
        private static readonly string uninstallRegkey86 = @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall";
        private static readonly string[] validExtensions = new string[]
        {
            ".txt", ".html",
            ".doc", ".docs", ".xls", ".xlsx", ".ppt", ".pptx",
            ".odt", ".ods", ".odp", ".odg", ".odc", ".odf", ".odi",
        };

        /// <summary>
        /// サポートしているファイルの種類一覧
        /// </summary>
        public static string[] SupportedFileKinds
            => validExtensions;

        /// <summary>
        /// sofficeのファイルパス
        /// </summary>
        public static string ProgramPath => programPath.Value;
        private static Lazy<string> programPath = new Lazy<string>(() =>
        {
            if (!(OperatingSystem.IsWindows() || OperatingSystem.IsMacOS() || OperatingSystem.IsLinux()))
                throw new NotSupportedException("Only Linux, Windows, and MacOS are supported.");

            if (InstallLocation == null)
                throw new FileNotFoundException("LibreOffice is not installed");

            if (OperatingSystem.IsWindows())
                return Path.Combine(InstallLocation, @"program\soffice");

            if (OperatingSystem.IsMacOS())
                return Path.Combine(InstallLocation, "Contents/MacOS/soffice");

            if (OperatingSystem.IsLinux())
                throw new NotSupportedException("Linux OS is not yet supported.");

            throw new InvalidOperationException("");
        });

        /// <summary>
        /// LibreOfficeのインストール先ディレクトリパス
        /// </summary>
        public static string InstallLocation => installLocation.Value;
        private static Lazy<string> installLocation = new Lazy<string>(() =>
        {
            if (OperatingSystem.IsWindows())
            {
                var dir = FindLibreOfficeInstallDirectoryInner(uninstallRegkey);
                return dir == null
                    ? FindLibreOfficeInstallDirectoryInner(uninstallRegkey86)
                    : dir;
            }

            if (OperatingSystem.IsMacOS())
            {
                return FindLibreOfficeApp();
            }

            if (OperatingSystem.IsLinux())
            {
                throw new NotSupportedException("Linux OS is not yet supported.");
            }

            throw new NotSupportedException("Only Linux, Windows, and MacOS are supported.");

            [SupportedOSPlatform("windows")]
            string FindLibreOfficeInstallDirectoryInner(string rootKey)
            {
                if (!OperatingSystem.IsWindows())
                    return null;

                var reg = Registry.LocalMachine.OpenSubKey(rootKey, false);
                if (reg != null)
                {
                    foreach (string subKey in reg.GetSubKeyNames())
                    {
                        var appkey = Registry.LocalMachine.OpenSubKey(rootKey + "\\" + subKey, false);

                        var appName = appkey.GetValue("DisplayName") == null
                            ? subKey
                            : appkey.GetValue("DisplayName").ToString();

                        if (!appName.StartsWith("LibreOffice"))
                            continue;

                        return appkey.GetValue("InstallLocation")?.ToString();
                    }
                }
                return null;
            }

            [SupportedOSPlatform("macos")]
            string FindLibreOfficeApp(string dirpath = "/Applications")
            {
                foreach (var dir in Directory.EnumerateDirectories(dirpath))
                {
                    if (dir.EndsWith("LibreOffice.app"))
                        return dir;

                    if (!dir.EndsWith(".app"))
                    {
                        var found = FindLibreOfficeApp(dir);
                        if (found != null)
                            return found;
                    }
                }
                return null;
            }
        });

        /// <summary>
        /// true: LibreOfficeがインストールされている
        /// </summary>
        public static bool Exists => InstallLocation != null;

        /// <summary>
        /// sourceファイルをPDFに変換する
        /// </summary>
        /// <param name="source">変換したいファイル</param>
        /// <param name="outputDirectory">出力先ディレクトリ</param>
        /// <returns></returns>
        public static async ValueTask ConvertToPdfAsync(string source, string outputDirectory = "./")
        {
            if (!validExtensions.Any(extension => Path.GetExtension(source) == extension))
                throw new NotSupportedException($"\'{Path.GetExtension(source)}\' is not supported.");

            var libreoffice = ProgramPath;
            var src = Path.GetFullPath(source);
            var working = Path.GetFullPath(outputDirectory);

            using var process = new Process();
            process.StartInfo.FileName = libreoffice;
            process.StartInfo.Arguments = $"-norestore -nologo -nofirststartwizard -headless -convert-to pdf \"{src}\"";
            process.StartInfo.WorkingDirectory = working;
            process.Start();
            await process.WaitForExitAsync();
        }
    }
}
