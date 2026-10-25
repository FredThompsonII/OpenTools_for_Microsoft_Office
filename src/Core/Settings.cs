// ---------------------------------------------------------------------------------------------------------------
// Name:     Settings.cs
// Purpose:  Reads and writes HKEY_CURRENT_USER\Software\OpenTools. A value is looked up in the app's
//           override key first (Apps\<App>), then in the global key, then in the built-in default. Defaults
//           match section 8 of the specification. The registry is read in Reload and the results are cached,
//           so painting never touches the registry. The static methods at the end read and write single stored
//           values for the settings window, which shows global values and per-app overrides separately.
// Revision: 11
// DTG:      261025100000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using Microsoft.Win32;

namespace OpenTools
{
    internal sealed class Settings
    {
        private const string RootPath = @"Software\OpenTools";
        private const string AppsFolder = "Apps";

        private static readonly Dictionary<string, object> _defaults = BuildDefaults();

        private readonly string _appName;
        private readonly Dictionary<string, object> _values = new Dictionary<string, object>();

        public Settings(string appName)
        {
            _appName = appName;
            Reload();
        }

        // ----- Defaults -----

        private static Dictionary<string, object> BuildDefaults()
        {
            Dictionary<string, object> d = new Dictionary<string, object>();
            d[SettingNames.TabsEnabled] = true;
            d[SettingNames.ToolbarsEnabled] = true;
            d[SettingNames.StartOnOpenToolsTab] = true;
            d[SettingNames.LoremEnabled] = true;
            d[SettingNames.LoremSections] = 1;
            d[SettingNames.LoremParagraphsLow] = 3;
            d[SettingNames.LoremParagraphsHigh] = 5;
            d[SettingNames.LoremSentences] = 5;
            d[SettingNames.LoremTitles] = false;
            d[SettingNames.DoubleClickClose] = true;
            d[SettingNames.MiddleClickClose] = true;
            d[SettingNames.DragToReorder] = true;
            d[SettingNames.BoldCurrentTab] = true;
            d[SettingNames.FileIcons] = true;
            d[SettingNames.NewDocumentButton] = true;
            d[SettingNames.CloseAllButton] = true;
            d[SettingNames.SingleWindowMode] = true;
            d[SettingNames.SyncRibbonTab] = true;
            d[SettingNames.EmbedTabBar] = true;
            d[SettingNames.TabBarPosition] = (int)TabBarPosition.Above;
            d[SettingNames.AutoWidth] = true;
            d[SettingNames.MaxTabChars] = 85;
            d[SettingNames.FixedTabWidth] = 150;
            d[SettingNames.MultiLine] = true;
            d[SettingNames.UseCustomFont] = false;
            d[SettingNames.CustomFontName] = string.Empty;
            d[SettingNames.UseCustomFontSize] = true;
            d[SettingNames.CustomFontSize] = 9;
            d[SettingNames.OverlayLeftOffset] = 0;
            d[SettingNames.OverlayRightReserve] = 60;
            return d;
        }

        // ----- Reading -----

        public void Reload()
        {
            try
            {
                using (RegistryKey hive = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
                using (RegistryKey global = hive.OpenSubKey(RootPath, false))
                using (RegistryKey appKey = hive.OpenSubKey(RootPath + "\\" + AppsFolder + "\\" + _appName, false))
                {
                    Dictionary<string, object> fresh = new Dictionary<string, object>();
                    foreach (KeyValuePair<string, object> pair in _defaults)
                    {
                        object found = ReadValue(appKey, pair.Key, pair.Value);
                        if (found == null)
                        {
                            found = ReadValue(global, pair.Key, pair.Value);
                        }

                        fresh[pair.Key] = found ?? pair.Value;
                    }

                    _values.Clear();
                    foreach (KeyValuePair<string, object> pair in fresh)
                    {
                        _values[pair.Key] = pair.Value;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Settings.Reload", ex);
                _values.Clear();
                foreach (KeyValuePair<string, object> pair in _defaults)
                {
                    _values[pair.Key] = pair.Value;
                }
            }
        }

        private static object ReadValue(RegistryKey key, string name, object defaultValue)
        {
            if (key == null)
            {
                return null;
            }

            object raw = key.GetValue(name, null);
            if (raw == null)
            {
                return null;
            }

            if (defaultValue is bool)
            {
                return raw is int ? (object)((int)raw != 0) : null;
            }

            if (defaultValue is int)
            {
                return raw is int ? raw : null;
            }

            return raw as string;
        }

        public bool GetBool(string name)
        {
            return (bool)_values[name];
        }

        public int GetInt(string name)
        {
            return (int)_values[name];
        }

        public string GetString(string name)
        {
            return (string)_values[name];
        }

        public TabBarPosition GetPosition()
        {
            int raw = GetInt(SettingNames.TabBarPosition);
            return raw == (int)TabBarPosition.Below ? TabBarPosition.Below : TabBarPosition.Above;
        }

        // ----- Writing -----

        public void SetBool(string name, bool value, bool appOverride)
        {
            WriteValue(name, value ? 1 : 0, RegistryValueKind.DWord, appOverride);
        }

        public void SetInt(string name, int value, bool appOverride)
        {
            WriteValue(name, value, RegistryValueKind.DWord, appOverride);
        }

        public void SetString(string name, string value, bool appOverride)
        {
            WriteValue(name, value, RegistryValueKind.String, appOverride);
        }

        private void WriteValue(string name, object value, RegistryValueKind kind, bool appOverride)
        {
            try
            {
                string path = appOverride ? RootPath + "\\" + AppsFolder + "\\" + _appName : RootPath;
                using (RegistryKey hive = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
                using (RegistryKey key = hive.CreateSubKey(path))
                {
                    key.SetValue(name, value, kind);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Settings.WriteValue " + name, ex);
            }

            Reload();
        }

        // ----- Single stored values, for the settings window -----

        public const string KeyPath = RootPath;

        public static object GetDefault(string name)
        {
            return _defaults[name];
        }

        // The value stored for one app (appName) or globally (appName null), typed like its default, or null when
        // nothing is stored there.
        public static object ReadStored(string appName, string name)
        {
            try
            {
                using (RegistryKey hive = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
                using (RegistryKey key = hive.OpenSubKey(StoredPath(appName), false))
                {
                    return ReadValue(key, name, _defaults[name]);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Settings.ReadStored " + name, ex);
                return null;
            }
        }

        public static void WriteStored(string appName, string name, object value)
        {
            try
            {
                using (RegistryKey hive = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
                using (RegistryKey key = hive.CreateSubKey(StoredPath(appName)))
                {
                    if (value is bool)
                    {
                        key.SetValue(name, (bool)value ? 1 : 0, RegistryValueKind.DWord);
                    }
                    else if (value is int)
                    {
                        key.SetValue(name, (int)value, RegistryValueKind.DWord);
                    }
                    else
                    {
                        key.SetValue(name, Convert.ToString(value), RegistryValueKind.String);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Settings.WriteStored " + name, ex);
            }
        }

        public static void DeleteStored(string appName, string name)
        {
            try
            {
                using (RegistryKey hive = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
                using (RegistryKey key = hive.OpenSubKey(StoredPath(appName), true))
                {
                    if (key != null)
                    {
                        key.DeleteValue(name, false);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Settings.DeleteStored " + name, ex);
            }
        }

        private static string StoredPath(string appName)
        {
            return appName == null ? RootPath : RootPath + "\\" + AppsFolder + "\\" + appName;
        }
    }
}
