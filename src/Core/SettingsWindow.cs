// ---------------------------------------------------------------------------------------------------------------
// Name:     SettingsWindow.cs
// Purpose:  The OpenTools settings window, opened from the ribbon button in each app. Pages: Apps (one row per
//           installed app with Tabs, Toolbars, Lorem Ipsum and Start on OpenTools tab; all but Tabs are for
//           Word, Excel and PowerPoint), Tabs (tab behavior), Windows (window options), Look (width, rows,
//           font), Lorem Ipsum (sections, paragraphs, sentences, titles) and About (name and version, what it does,
//           copyright, license and the GitHub address). The window is just wide enough for every page name on one row.
//           Only settings the add-in actually uses are shown. "Settings for" chooses between the global
//           settings and one app's own overrides; in an app's view, a value set for that app only is shown in
//           bold, and "Clear this app's overrides" returns the app to the global settings. Window options are
//           grayed out for Project and Visio, which keep every document in one main window. Backup writes
//           HKCU\Software\OpenTools to a .reg file with reg.exe. Restore checks that the file only touches that
//           key and deletes nothing, then replaces the current settings with it.
// Revision: 11
// DTG:      261025120000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OpenTools
{
    internal sealed class SettingsWindow : Form
    {
        private const string AllApps = "All apps";
        private const string RegistryHeader = "Windows Registry Editor Version 5.00";
        private const string KeyPrefix = "[HKEY_CURRENT_USER\\" + Settings.KeyPath;
        private const int RegTimeoutMilliseconds = 15000;

        private const int MinimumPageWidth = 400;
        private const int PageHeight = 270;
        private const int TextMargin = 40;
        private const int TabPadding = 16;
        private const int TabRowSlack = 8;
        private const string AboutTitle = "OpenTools for Microsoft Office 1.00";
        private const string AboutText = "Classic toolbars, lorem ipsum, and tab support\n"
            + "for Microsoft Office 2016-2024";
        private const string CopyrightText = "(c) Fred Thompson 2026";
        private const string LicenseText = "MIT License";
        private const string GitHubAddress = "https://github.com/FredThompsonII/OpenTools_for_Microsoft_Office";
        private static readonly string[] PageNames = { "Apps", "Tabs", "Windows", "Look", "Lorem Ipsum", "About" };

        private int _pageWidth;
        private int _textWidth;
        private static readonly string[] ToolbarApps = { "Word", "Excel", "PowerPoint" };
        private static readonly string[] AppColumns =
        {
            SettingNames.TabsEnabled, SettingNames.ToolbarsEnabled, SettingNames.LoremEnabled,
            SettingNames.StartOnOpenToolsTab
        };
        private static readonly string[] AppColumnTitles =
        {
            "Tabs", "Toolbars", "Lorem Ipsum", "Start on OpenTools tab"
        };

        private readonly string[] _appNames;

        private readonly Action _saved;
        private readonly ComboBox _scope = new ComboBox();
        private readonly Button _clearOverrides = new Button();
        private readonly Dictionary<string, CheckBox> _appBoxes = new Dictionary<string, CheckBox>();
        private readonly Dictionary<string, Control> _controls = new Dictionary<string, Control>();
        private readonly HashSet<string> _changed = new HashSet<string>();
        private readonly HashSet<string> _appsChanged = new HashSet<string>();
        private readonly List<Control> _windowOptions = new List<Control>();
        private RadioButton _autoWidth;
        private RadioButton _fixedWidth;
        private string _scopeApp;
        private bool _loading;

        // currentApp selects that app's view first; null opens the global settings. saved runs after each save.
        public SettingsWindow(string currentApp, Action saved)
        {
            _saved = saved;
            _appNames = InstalledApps.List(currentApp);
            Text = "OpenTools for Microsoft Office Settings";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(10);

            TableLayoutPanel outer = new TableLayoutPanel();
            outer.AutoSize = true;
            outer.ColumnCount = 1;
            outer.Dock = DockStyle.Fill;
            Controls.Add(outer);

            FlowLayoutPanel scopeRow = Row();
            scopeRow.Controls.Add(Label("Settings for:"));
            _scope.DropDownStyle = ComboBoxStyle.DropDownList;
            _scope.Width = 140;
            _scope.Items.Add(AllApps);
            foreach (string app in _appNames)
            {
                _scope.Items.Add(app);
            }

            scopeRow.Controls.Add(_scope);
            _clearOverrides.Text = "Clear this app's overrides";
            _clearOverrides.AutoSize = true;
            _clearOverrides.Click += OnClearOverrides;
            scopeRow.Controls.Add(_clearOverrides);
            outer.Controls.Add(scopeRow);

            _pageWidth = Math.Max(MinimumPageWidth, TabRowWidth());
            _textWidth = _pageWidth - TextMargin;
            TabControl pages = new TabControl();
            pages.Width = _pageWidth;
            pages.Height = PageHeight;
            pages.TabPages.Add(AppsPage());
            pages.TabPages.Add(TabsPage());
            pages.TabPages.Add(WindowsPage());
            pages.TabPages.Add(LookPage());
            pages.TabPages.Add(LoremPage());
            pages.TabPages.Add(AboutPage());
            outer.Controls.Add(pages);

            Label note = Label("Changes show in each app the next time you switch documents there. Changes on "
                + "the Apps page take effect the next time that app starts.");
            note.MaximumSize = new Size(_pageWidth, 0);
            outer.Controls.Add(note);

            FlowLayoutPanel buttons = Row();
            buttons.Controls.Add(Button("Backup...", OnBackup));
            buttons.Controls.Add(Button("Restore...", OnRestore));
            Button ok = Button("OK", OnOk);
            Button cancel = Button("Cancel", delegate { Close(); });
            buttons.Controls.Add(ok);
            buttons.Controls.Add(cancel);
            outer.Controls.Add(buttons);
            AcceptButton = ok;
            CancelButton = cancel;

            int start = currentApp == null ? 0 : Array.IndexOf(_appNames, currentApp) + 1;
            _scope.SelectedIndex = start < 0 ? 0 : start;
            _scope.SelectedIndexChanged += OnScopeChanged;
            LoadScope();
        }

        // ----- Pages -----

        private TabPage AppsPage()
        {
            TabPage page = Page(PageNames[0]);
            TableLayoutPanel grid = new TableLayoutPanel();
            grid.AutoSize = true;
            grid.ColumnCount = AppColumns.Length + 1;
            grid.Padding = new Padding(6);
            grid.Controls.Add(Label(string.Empty), 0, 0);
            for (int column = 0; column < AppColumnTitles.Length; column++)
            {
                grid.Controls.Add(Label(AppColumnTitles[column]), column + 1, 0);
            }

            for (int row = 0; row < _appNames.Length; row++)
            {
                string app = _appNames[row];
                grid.Controls.Add(Label(app), 0, row + 1);
                bool toolbarApp = Array.IndexOf(ToolbarApps, app) >= 0;
                for (int column = 0; column < AppColumns.Length; column++)
                {
                    if (AppColumns[column] != SettingNames.TabsEnabled && !toolbarApp)
                    {
                        continue;
                    }

                    CheckBox box = new CheckBox();
                    box.AutoSize = true;
                    box.Anchor = AnchorStyles.None;
                    box.AccessibleName = app + " " + AppColumnTitles[column];
                    string key = AppKey(app, AppColumns[column]);
                    box.CheckedChanged += delegate
                    {
                        if (!_loading)
                        {
                            _appsChanged.Add(key);
                        }
                    };
                    _appBoxes[key] = box;
                    grid.Controls.Add(box, column + 1, row + 1);
                }
            }

            page.Controls.Add(grid);
            return page;
        }

        private static string AppKey(string app, string settingName)
        {
            return app + "|" + settingName;
        }

        private static void SplitAppKey(string key, out string app, out string settingName)
        {
            int split = key.IndexOf('|');
            app = key.Substring(0, split);
            settingName = key.Substring(split + 1);
        }

        private TabPage TabsPage()
        {
            TabPage page = Page(PageNames[1]);
            FlowLayoutPanel list = Column();
            list.Controls.Add(Check(SettingNames.DoubleClickClose, "Double-click a tab to close it"));
            list.Controls.Add(Check(SettingNames.MiddleClickClose, "Middle-click a tab to close it"));
            list.Controls.Add(Check(SettingNames.DragToReorder, "Drag tabs to reorder them"));
            list.Controls.Add(Check(SettingNames.BoldCurrentTab, "Bold text on the current tab"));
            list.Controls.Add(Check(SettingNames.FileIcons, "File icon on each tab"));
            list.Controls.Add(Check(SettingNames.NewDocumentButton, "New-document button (+) at the end of the tabs"));
            list.Controls.Add(Check(SettingNames.CloseAllButton, "Close-all button at the end of the tab bar"));
            page.Controls.Add(list);
            return page;
        }

        private TabPage WindowsPage()
        {
            TabPage page = Page(PageNames[2]);
            FlowLayoutPanel list = Column();
            CheckBox single = Check(SettingNames.SingleWindowMode,
                "Single-window mode: every document window gets the same size and place");
            CheckBox ribbon = Check(SettingNames.SyncRibbonTab,
                "Keep the selected ribbon tab in sync between documents");
            single.MaximumSize = new Size(_textWidth, 0);
            ribbon.MaximumSize = new Size(_textWidth, 0);
            list.Controls.Add(single);
            list.Controls.Add(ribbon);
            Label oneWindow = Label("Project and Visio keep every document in one main window, so these options "
                + "don't apply to them.");
            oneWindow.MaximumSize = new Size(_textWidth, 0);
            list.Controls.Add(oneWindow);
            _windowOptions.Add(single);
            _windowOptions.Add(ribbon);
            page.Controls.Add(list);
            return page;
        }

        private TabPage LookPage()
        {
            TabPage page = Page(PageNames[3]);
            FlowLayoutPanel list = Column();

            _autoWidth = new RadioButton();
            _autoWidth.Text = "Auto width max";
            _autoWidth.AutoSize = true;
            _fixedWidth = new RadioButton();
            _fixedWidth.Text = "Fixed width of";
            _fixedWidth.AutoSize = true;
            _autoWidth.CheckedChanged += delegate { Changed(SettingNames.AutoWidth); };
            _controls[SettingNames.AutoWidth] = _autoWidth;

            ComboBox fonts = new ComboBox();
            fonts.DropDownStyle = ComboBoxStyle.DropDown;
            fonts.Width = 200;
            foreach (FontFamily family in FontFamily.Families)
            {
                fonts.Items.Add(family.Name);
            }

            fonts.TextChanged += delegate { Changed(SettingNames.CustomFontName); };
            _controls[SettingNames.CustomFontName] = fonts;

            // One grid, so every field starts in the same column.
            TableLayoutPanel grid = new TableLayoutPanel();
            grid.AutoSize = true;
            grid.ColumnCount = 3;
            grid.Controls.Add(_autoWidth, 0, 0);
            grid.Controls.Add(Number(SettingNames.MaxTabChars, 10, 260), 1, 0);
            grid.Controls.Add(Label("characters"), 2, 0);
            grid.Controls.Add(_fixedWidth, 0, 1);
            grid.Controls.Add(Number(SettingNames.FixedTabWidth, 40, 600), 1, 1);
            grid.Controls.Add(Label("pixels"), 2, 1);
            CheckBox multiLine = Check(SettingNames.MultiLine, "Wrap tabs onto more rows when they don't fit");
            grid.Controls.Add(multiLine, 0, 2);
            grid.SetColumnSpan(multiLine, 3);
            grid.Controls.Add(Check(SettingNames.UseCustomFont, "Font:"), 0, 3);
            grid.Controls.Add(fonts, 1, 3);
            grid.SetColumnSpan(fonts, 2);
            grid.Controls.Add(Check(SettingNames.UseCustomFontSize, "Font size:"), 0, 4);
            grid.Controls.Add(Number(SettingNames.CustomFontSize, 6, 24), 1, 4);
            list.Controls.Add(grid);

            page.Controls.Add(list);
            return page;
        }

        private TabPage LoremPage()
        {
            TabPage page = Page(PageNames[4]);
            FlowLayoutPanel list = Column();

            TableLayoutPanel grid = new TableLayoutPanel();
            grid.AutoSize = true;
            grid.ColumnCount = 2;
            grid.Controls.Add(Label("Sections"), 0, 0);
            grid.Controls.Add(Number(SettingNames.LoremSections, 1, 20), 1, 0);
            grid.Controls.Add(Label("Paragraphs"), 0, 1);
            FlowLayoutPanel paragraphs = Row();
            paragraphs.Margin = Padding.Empty;
            paragraphs.Controls.Add(Number(SettingNames.LoremParagraphsLow, 1, 50));
            paragraphs.Controls.Add(Label("to"));
            paragraphs.Controls.Add(Number(SettingNames.LoremParagraphsHigh, 1, 50));
            grid.Controls.Add(paragraphs, 1, 1);
            grid.Controls.Add(Label("Sentences"), 0, 2);
            grid.Controls.Add(Number(SettingNames.LoremSentences, 1, 30), 1, 2);
            list.Controls.Add(grid);

            list.Controls.Add(Check(SettingNames.LoremTitles, "A title at the start of each section"));

            Label note = Label("Word uses all of these. Excel uses the sentence count for one cell. PowerPoint uses "
                + "the paragraph counts, with one sentence in each paragraph.");
            note.MaximumSize = new Size(_textWidth, 0);
            list.Controls.Add(note);

            page.Controls.Add(list);
            return page;
        }

        private TabPage AboutPage()
        {
            TabPage page = Page(PageNames[5]);
            FlowLayoutPanel list = Column();
            Label title = Label(AboutTitle);
            title.Font = new Font(Font, FontStyle.Bold);
            list.Controls.Add(title);
            Label about = Label(AboutText);
            about.MaximumSize = new Size(_textWidth, 0);
            list.Controls.Add(about);
            Label copyright = Label(CopyrightText);
            copyright.Margin = new Padding(copyright.Margin.Left, copyright.Margin.Top + 12, copyright.Margin.Right,
                copyright.Margin.Bottom);
            list.Controls.Add(copyright);
            list.Controls.Add(Label(LicenseText));
            LinkLabel link = new LinkLabel();
            link.Text = GitHubAddress;
            link.AutoSize = true;
            link.MaximumSize = new Size(_textWidth, 0);
            link.Padding = new Padding(0, 4, 0, 0);
            link.LinkClicked += OnGitHubClicked;
            list.Controls.Add(link);
            page.Controls.Add(list);
            return page;
        }

        private void OnGitHubClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                Process.Start(GitHubAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("SettingsWindow.OnGitHubClicked", ex);
            }
        }

        // The width the page names need on one row: each name's text plus the space Windows puts around it.
        private int TabRowWidth()
        {
            int width = TabRowSlack;
            foreach (string name in PageNames)
            {
                width += TextRenderer.MeasureText(name, Font).Width + TabPadding;
            }

            return width;
        }

        // ----- Loading and saving -----

        private void OnScopeChanged(object sender, EventArgs e)
        {
            string chosen = _scope.SelectedIndex <= 0 ? null : _appNames[_scope.SelectedIndex - 1];
            if (chosen == _scopeApp)
            {
                return;
            }

            if (_changed.Count > 0 || _appsChanged.Count > 0)
            {
                DialogResult answer = MessageBox.Show(this, "Save the changes made for " + ScopeName() + " first?",
                    Text, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (answer == DialogResult.Cancel)
                {
                    _loading = true;
                    _scope.SelectedIndex = _scopeApp == null ? 0 : Array.IndexOf(_appNames, _scopeApp) + 1;
                    _loading = false;
                    return;
                }

                if (answer == DialogResult.Yes)
                {
                    Save();
                }
            }

            LoadScope();
        }

        private void LoadScope()
        {
            _loading = true;
            try
            {
                _scopeApp = _scope.SelectedIndex <= 0 ? null : _appNames[_scope.SelectedIndex - 1];
                _clearOverrides.Visible = _scopeApp != null;
                foreach (KeyValuePair<string, CheckBox> pair in _appBoxes)
                {
                    string app;
                    string settingName;
                    SplitAppKey(pair.Key, out app, out settingName);
                    pair.Value.Checked = (bool)Effective(app, settingName);
                }

                foreach (KeyValuePair<string, Control> pair in _controls)
                {
                    object value = Effective(_scopeApp, pair.Key);
                    bool own = _scopeApp != null && Settings.ReadStored(_scopeApp, pair.Key) != null;
                    Show(pair.Key, pair.Value, value, own);
                }

                bool oneWindow = _scopeApp == "Project" || _scopeApp == "Visio";
                foreach (Control control in _windowOptions)
                {
                    control.Enabled = !oneWindow;
                }

                _changed.Clear();
                _appsChanged.Clear();
            }
            finally
            {
                _loading = false;
            }
        }

        private void Show(string name, Control control, object value, bool own)
        {
            control.Font = new Font(Font, own ? FontStyle.Bold : FontStyle.Regular);
            if (name == SettingNames.AutoWidth)
            {
                _autoWidth.Checked = (bool)value;
                _fixedWidth.Checked = !(bool)value;
                _fixedWidth.Font = control.Font;
                return;
            }

            CheckBox box = control as CheckBox;
            if (box != null)
            {
                box.Checked = (bool)value;
                return;
            }

            NumericUpDown number = control as NumericUpDown;
            if (number != null)
            {
                number.Value = Math.Max(number.Minimum, Math.Min(number.Maximum, (int)value));
                return;
            }

            control.Text = Convert.ToString(value);
        }

        private static object Effective(string appName, string name)
        {
            object value = appName == null ? null : Settings.ReadStored(appName, name);
            return value ?? Settings.ReadStored(null, name) ?? Settings.GetDefault(name);
        }

        private void Save()
        {
            foreach (string name in _changed)
            {
                Settings.WriteStored(_scopeApp, name, ValueOf(name, _controls[name]));
            }

            foreach (string key in _appsChanged)
            {
                string app;
                string settingName;
                SplitAppKey(key, out app, out settingName);
                Settings.WriteStored(app, settingName, _appBoxes[key].Checked);
            }

            _changed.Clear();
            _appsChanged.Clear();
            Applied();
        }

        private object ValueOf(string name, Control control)
        {
            if (name == SettingNames.AutoWidth)
            {
                return _autoWidth.Checked;
            }

            CheckBox box = control as CheckBox;
            if (box != null)
            {
                return box.Checked;
            }

            NumericUpDown number = control as NumericUpDown;
            if (number != null)
            {
                return (int)number.Value;
            }

            return control.Text;
        }

        private void Applied()
        {
            if (_saved == null)
            {
                return;
            }

            try
            {
                _saved();
            }
            catch (Exception ex)
            {
                Logger.Error("SettingsWindow.Applied", ex);
            }
        }

        private void OnOk(object sender, EventArgs e)
        {
            Save();
            Close();
        }

        private void OnClearOverrides(object sender, EventArgs e)
        {
            if (_scopeApp == null)
            {
                return;
            }

            DialogResult answer = MessageBox.Show(this, "Return " + _scopeApp + " to the global settings? Its own "
                + "values on the Tabs, Windows and Look pages are removed. Its on/off setting on the Apps page stays.",
                Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (answer != DialogResult.OK)
            {
                return;
            }

            foreach (string name in _controls.Keys)
            {
                Settings.DeleteStored(_scopeApp, name);
            }

            LoadScope();
            Applied();
        }

        // ----- Backup and Restore -----

        private void OnBackup(object sender, EventArgs e)
        {
            if ((_changed.Count > 0 || _appsChanged.Count > 0) && !SaveFirst())
            {
                return;
            }

            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Title = "Back up OpenTools settings";
                dialog.Filter = "Registry file (*.reg)|*.reg";
                dialog.DefaultExt = "reg";
                dialog.AddExtension = true;
                dialog.OverwritePrompt = true;
                dialog.FileName = "OpenTools Settings " + DateTime.Now.ToString("yyyy-MM-dd HHmm") + ".reg";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                // reg.exe can only export a key that exists; a user who never changed a setting has none yet.
                using (RegistryKey hive = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
                using (RegistryKey key = hive.CreateSubKey(Settings.KeyPath))
                {
                }

                string error = RunReg("export \"HKCU\\" + Settings.KeyPath + "\" \"" + dialog.FileName + "\" /y");
                if (error == null)
                {
                    MessageBox.Show(this, "Settings saved to " + dialog.FileName, Text, MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(this, "The backup failed. " + error, Text, MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
        }

        private void OnRestore(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "Restore OpenTools settings";
                dialog.Filter = "Registry file (*.reg)|*.reg";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                string problem = CheckBackup(dialog.FileName);
                if (problem != null)
                {
                    MessageBox.Show(this, "This file can't be restored. " + problem, Text, MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                DialogResult answer = MessageBox.Show(this, "Replace all current OpenTools settings with the settings "
                    + "in this file?", Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (answer != DialogResult.OK)
                {
                    return;
                }

                using (RegistryKey hive = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
                {
                    hive.DeleteSubKeyTree(Settings.KeyPath, false);
                }

                string error = RunReg("import \"" + dialog.FileName + "\"");
                LoadScope();
                Applied();
                if (error == null)
                {
                    MessageBox.Show(this, "Settings restored. Turning an app's tabs on or off takes effect the next "
                        + "time that app starts.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(this, "The restore failed, so the default settings are in use. " + error, Text,
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        // A backup may only add values under HKCU\Software\OpenTools. Any other key, and any line that
        // deletes a key, is refused before anything is changed.
        private static string CheckBackup(string path)
        {
            string text;
            try
            {
                text = File.ReadAllText(path);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            string[] lines = text.Replace("\r\n", "\n").Split('\n');
            bool header = false;
            bool anyKey = false;
            foreach (string raw in lines)
            {
                string line = raw.Trim();
                if (line.Length == 0)
                {
                    continue;
                }

                if (!header)
                {
                    if (line != RegistryHeader)
                    {
                        return "It isn't a registry file.";
                    }

                    header = true;
                    continue;
                }

                if (line.StartsWith("[", StringComparison.Ordinal))
                {
                    if (!line.StartsWith(KeyPrefix, StringComparison.OrdinalIgnoreCase)
                        || !(line.Length > KeyPrefix.Length && (line[KeyPrefix.Length] == ']'
                        || line[KeyPrefix.Length] == '\\')))
                    {
                        return "It contains settings for something other than OpenTools: " + line;
                    }

                    anyKey = true;
                }
            }

            return anyKey ? null : "It doesn't contain OpenTools settings.";
        }

        private static string RunReg(string arguments)
        {
            try
            {
                ProcessStartInfo start = new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "reg.exe"),
                    arguments);
                start.UseShellExecute = false;
                start.CreateNoWindow = true;
                start.RedirectStandardError = true;
                start.RedirectStandardOutput = true;
                using (Process process = Process.Start(start))
                {
                    string errors = process.StandardError.ReadToEnd();
                    process.StandardOutput.ReadToEnd();
                    if (!process.WaitForExit(RegTimeoutMilliseconds))
                    {
                        return "reg.exe did not finish.";
                    }

                    return process.ExitCode == 0 ? null : errors.Trim();
                }
            }
            catch (Exception ex)
            {
                Logger.Error("SettingsWindow.RunReg", ex);
                return ex.Message;
            }
        }

        private bool SaveFirst()
        {
            DialogResult answer = MessageBox.Show(this, "Save the changes in this window before the backup?", Text,
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (answer == DialogResult.Cancel)
            {
                return false;
            }

            if (answer == DialogResult.Yes)
            {
                Save();
            }

            return true;
        }

        // ----- Small builders -----

        private void Changed(string name)
        {
            if (!_loading)
            {
                _changed.Add(name);
            }
        }

        private string ScopeName()
        {
            return _scopeApp ?? "all apps";
        }

        private CheckBox Check(string name, string text)
        {
            CheckBox box = new CheckBox();
            box.Text = text;
            box.AutoSize = true;
            box.CheckedChanged += delegate { Changed(name); };
            _controls[name] = box;
            return box;
        }

        private NumericUpDown Number(string name, int minimum, int maximum)
        {
            NumericUpDown number = new NumericUpDown();
            number.Minimum = minimum;
            number.Maximum = maximum;
            number.Width = 60;
            number.ValueChanged += delegate { Changed(name); };
            _controls[name] = number;
            return number;
        }

        private static TabPage Page(string text)
        {
            TabPage page = new TabPage(text);
            page.Padding = new Padding(8);
            page.UseVisualStyleBackColor = true;
            return page;
        }

        private static FlowLayoutPanel Column()
        {
            FlowLayoutPanel panel = new FlowLayoutPanel();
            panel.FlowDirection = FlowDirection.TopDown;
            panel.WrapContents = false;
            panel.Dock = DockStyle.Fill;
            panel.AutoScroll = true;
            return panel;
        }

        private static FlowLayoutPanel Row()
        {
            FlowLayoutPanel panel = new FlowLayoutPanel();
            panel.FlowDirection = FlowDirection.LeftToRight;
            panel.WrapContents = false;
            panel.AutoSize = true;
            return panel;
        }

        private static Label Label(string text)
        {
            Label label = new Label();
            label.Text = text;
            label.AutoSize = true;
            label.Padding = new Padding(0, 4, 0, 0);
            return label;
        }

        private static Button Button(string text, EventHandler click)
        {
            Button button = new Button();
            button.Text = text;
            button.AutoSize = true;
            button.Click += click;
            return button;
        }
    }
}
