// ---------------------------------------------------------------------------------------------------------------
// Name:     InstalledApps.cs
// Purpose:  Finds which supported Office apps are installed, so the settings window lists only those. An app
//           counts as installed when its App Paths entry exists (WINWORD.EXE, EXCEL.EXE, POWERPNT.EXE,
//           WINPROJ.EXE, VISIO.EXE) in the 64-bit or 32-bit registry view, under the machine or the user, and the
//           file it points to exists. The app the window was opened from always counts. If nothing is found, every
//           app is listed, so the window is never empty.
// Revision: 1
// DTG:      261024210000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Win32;

namespace OpenTools
{
    internal static class InstalledApps
    {
        private const string AppPathsKey = @"Software\Microsoft\Windows\CurrentVersion\App Paths\";

        private static readonly string[] AppNames = { "Word", "Excel", "PowerPoint", "Project", "Visio" };
        private static readonly string[] ExeNames =
        {
            "WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE", "WINPROJ.EXE", "VISIO.EXE"
        };

        // ----- Listing -----

        public static string[] List(string currentApp)
        {
            List<string> found = new List<string>();
            for (int i = 0; i < AppNames.Length; i++)
            {
                if (AppNames[i] == currentApp || IsInstalled(ExeNames[i]))
                {
                    found.Add(AppNames[i]);
                }
            }

            return found.Count == 0 ? (string[])AppNames.Clone() : found.ToArray();
        }

        private static bool IsInstalled(string exeName)
        {
            RegistryHive[] hives = { RegistryHive.LocalMachine, RegistryHive.CurrentUser };
            RegistryView[] views = { RegistryView.Registry64, RegistryView.Registry32 };
            foreach (RegistryHive hive in hives)
            {
                foreach (RegistryView view in views)
                {
                    if (PathExists(hive, view, exeName))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool PathExists(RegistryHive hive, RegistryView view, string exeName)
        {
            try
            {
                using (RegistryKey root = RegistryKey.OpenBaseKey(hive, view))
                using (RegistryKey key = root.OpenSubKey(AppPathsKey + exeName, false))
                {
                    if (key == null)
                    {
                        return false;
                    }

                    string path = key.GetValue(string.Empty) as string;
                    if (string.IsNullOrEmpty(path))
                    {
                        return false;
                    }

                    path = Environment.ExpandEnvironmentVariables(path.Trim().Trim('"'));
                    return File.Exists(path);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("InstalledApps.PathExists " + exeName, ex);
                return false;
            }
        }
    }
}
