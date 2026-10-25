// ---------------------------------------------------------------------------------------------------------------
// Name:     Logger.cs
// Purpose:  Writes one line per event to %LOCALAPPDATA%\OpenTools\OpenTools.log. Never throws, because
//           no exception may escape into Office. When the log passes 1 MB it moves to OpenTools.old.log,
//           replacing the previous .old file. This is the add-in's own diagnostic log, not user data.
// Revision: 2
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace OpenTools
{
    internal static class Logger
    {
        private const long MaxLogBytes = 1048576;

        private static readonly object _lock = new object();
        private static string _context = "Unknown";

        // ----- Setup -----

        public static void SetContext(string appName)
        {
            _context = appName;
        }

        // ----- Writing -----

        public static void Info(string message)
        {
            Write("INFO ", message);
        }

        public static void Error(string where, Exception ex)
        {
            string detail = ex == null ? "(no exception)" : ex.GetType().Name + ": " + ex.Message;
            Write("ERROR", where + " - " + detail);
        }

        private static void Write(string level, string message)
        {
            try
            {
                lock (_lock)
                {
                    string folder = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "OpenTools");
                    Directory.CreateDirectory(folder);

                    string path = Path.Combine(folder, "OpenTools.log");
                    RollIfLarge(folder, path);

                    string line = string.Format("{0:yyyy-MM-dd HH:mm:ss} {1} {2} pid={3} {4}{5}",
                        DateTime.Now, level, _context, Process.GetCurrentProcess().Id, message,
                        Environment.NewLine);
                    File.AppendAllText(path, line, new UTF8Encoding(false));
                }
            }
            catch (Exception)
            {
                // Logging must never disturb Office. Nothing else can be done here.
            }
        }

        private static void RollIfLarge(string folder, string path)
        {
            FileInfo info = new FileInfo(path);
            if (!info.Exists || info.Length < MaxLogBytes)
            {
                return;
            }

            string oldPath = Path.Combine(folder, "OpenTools.old.log");
            if (File.Exists(oldPath))
            {
                File.Delete(oldPath);
            }

            File.Move(path, oldPath);
        }
    }
}
