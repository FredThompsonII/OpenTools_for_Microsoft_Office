// ---------------------------------------------------------------------------------------------------------------
// Name:     Connect.cs
// Purpose:  The add-in itself, registered as OpenTools.Connect. Office loads it and calls the five
//           IDTExtensibility2 methods, plus CTPFactoryAvailable so panes can be created. It sets the window
//           class names for the host app (AppProfile) and starts the Word or Excel adapter with a tab
//           coordinator. PowerPoint uses the same coordinator with its own window adapter. Microsoft Project and
//           Visio are recognised by their process (WINPROJ, VISIO) and reached by late binding, because no
//           interop package is used for them. It also supplies the OpenTools button at the end of the View tab
//           (TabView) in every app (IRibbonExtensibility). The button opens the settings window and uses the
//           WindowsArrangeAll icon, which all five apps have. It works even when tabs are turned off for the app,
//           so they can be turned back on. The class is exposed through IDispatch (AutoDispatch) because Office
//           calls ribbon callbacks by name. In Word, Excel and PowerPoint it also adds the OpenTools tab with
//           the OpenToolbars menus and toolbars (ToolbarRibbonXml), when Toolbars is turned on for that app. It
//           selects that tab when the app starts and in each new window when Start on OpenTools tab is on, or
//           Home when it is off. Every Settings button uses OpenTools' own gear picture. The Insert menu's
//           Lorem Ipsum item and the Lorem Ipsum button at the far right of the Insert tab call OnLoremIpsum
//           (LoremInserter). The button uses OpenTools' own page picture (LoremIcon). No exception escapes.
// Revision: 23
// DTG:      261025050000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace OpenTools
{
    [ComVisible(true)]
    [Guid("65CA4A37-C8CA-40BA-A55B-01E1DD4559E0")]
    [ProgId("OpenTools.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class Connect : IDTExtensibility2, Office.ICustomTaskPaneConsumer, Office.IRibbonExtensibility
    {
        private IAppAdapter _adapter;
        private TabCoordinator _coordinator;
        private Office.ICTPFactory _factory;
        private object _application;
        private Office.IRibbonUI _ribbon;
        private bool _toolbarsShown;
        private StartTabWatcher _startTab;

        private const string SettingsTabXml =
            "<tab idMso=\"TabView\"><group id=\"OpenToolsGroup\" label=\"OpenTools\">"
            + "<button id=\"OpenToolsSettings\" label=\"OpenTools\" size=\"large\" getImage=\"GetSettingsImage\" "
            + "onAction=\"OnOpenSettings\" screentip=\"OpenTools settings\" /></group></tab>";
        private const char TagSeparator = '|';
        private const string MenuSettingsId = "ot_menu_view_settings";
        private const string InsertTabXml =
            "<tab idMso=\"TabInsert\"><group id=\"ot_insert_group\" label=\"OpenTools\">"
            + "<button id=\"ot_insert_lorem\" label=\"Lorem Ipsum\" size=\"large\" getImage=\"GetLoremImage\" "
            + "onAction=\"OnLoremIpsum\" screentip=\"Insert Lorem Ipsum filler text\" /></group></tab>";
        private const int LargePicture = 32;
        private const int SmallPicture = 16;
        private static object _largeGear;
        private static object _smallGear;
        private static object _loremPage;

        // ----- IDTExtensibility2 -----

        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            try
            {
                _application = application;
                Logger.Info("OnConnection reached.");
                Word.Application word = application as Word.Application;
                if (word == null)
                {
                    Excel.Application excel = application as Excel.Application;
                    if (excel != null)
                    {
                        StartExcel(excel);
                        return;
                    }

                    PowerPoint.Application powerPoint = application as PowerPoint.Application;
                    if (powerPoint != null)
                    {
                        StartPowerPoint(powerPoint);
                        return;
                    }

                    if (IsProcess("WINPROJ"))
                    {
                        StartProject(application);
                        return;
                    }

                    if (IsProcess("VISIO"))
                    {
                        StartVisio(application);
                        return;
                    }

                    Logger.Info("This host app is not supported by this build. The add-in stays idle.");
                    return;
                }

                Logger.SetContext("Word");
                AppProfile.Use("OpusApp", "_WwF");
                Settings settings = new Settings("Word");
                if (!settings.GetBool(SettingNames.TabsEnabled))
                {
                    Logger.Info("Tabs are turned off for Word in the settings.");
                    return;
                }

                _adapter = new WordAdapter(word);
                _coordinator = new TabCoordinator(_adapter, settings);
                if (_factory != null)
                {
                    _coordinator.SetPaneFactory(_factory);
                }

                _adapter.Start(_coordinator);
            }
            catch (Exception ex)
            {
                Logger.Error("Connect.OnConnection", ex);
            }
        }

        // Excel: each workbook has its own XLMAIN frame. The document area is XLDESK.
        private void StartExcel(Excel.Application excel)
        {
            Logger.SetContext("Excel");
            Settings settings = new Settings("Excel");
            if (!settings.GetBool(SettingNames.TabsEnabled))
            {
                Logger.Info("Tabs are turned off for Excel in the settings.");
                return;
            }

            AppProfile.Use("XLMAIN", "XLDESK");
            _adapter = new ExcelAdapter(excel);
            _coordinator = new TabCoordinator(_adapter, settings);
            if (_factory != null)
            {
                _coordinator.SetPaneFactory(_factory);
            }

            _adapter.Start(_coordinator);
        }

        // PowerPoint: separate PPTFrameClass frames with a visible MDIClient document area.
        // Closing asks about unsaved changes before each presentation is closed, so close is enabled.
        private void StartPowerPoint(PowerPoint.Application powerPoint)
        {
            Logger.SetContext("PowerPoint");
            Settings settings = new Settings("PowerPoint");
            if (!settings.GetBool(SettingNames.TabsEnabled))
            {
                Logger.Info("Tabs are turned off for PowerPoint in the settings.");
                return;
            }

            AppProfile.Use("PPTFrameClass", "MDIClient");
            _adapter = new PowerPointAdapter(powerPoint);
            _coordinator = new TabCoordinator(_adapter, settings);
            if (_factory != null)
            {
                _coordinator.SetPaneFactory(_factory);
            }

            _adapter.Start(_coordinator);
        }

        private static bool IsProcess(string name)
        {
            using (Process process = Process.GetCurrentProcess())
            {
                return string.Equals(process.ProcessName, name, StringComparison.OrdinalIgnoreCase);
            }
        }

        // Project: one JWinproj-WhimperMainClass main window holds every project in its MDIClient.

        private void StartProject(object application)
        {
            Logger.SetContext("Project");
            Settings settings = new Settings("Project");
            if (!settings.GetBool(SettingNames.TabsEnabled))
            {
                Logger.Info("Tabs are turned off for Project in the settings.");
                return;
            }

            AppProfile.Use("JWinproj-WhimperMainClass", "MDIClient");
            _adapter = new ProjectAdapter(application);
            _coordinator = new TabCoordinator(_adapter, settings);
            if (_factory != null)
            {
                _coordinator.SetPaneFactory(_factory);
            }

            _adapter.Start(_coordinator);
        }

        // Visio: one VISIOA main window holds every drawing in its MDIClient.
        private void StartVisio(object application)
        {
            Logger.SetContext("Visio");
            Settings settings = new Settings("Visio");
            if (!settings.GetBool(SettingNames.TabsEnabled))
            {
                Logger.Info("Tabs are turned off for Visio in the settings.");
                return;
            }

            AppProfile.Use("VISIOA", "MDIClient");
            _adapter = new VisioAdapter(application);
            _coordinator = new TabCoordinator(_adapter, settings);
            if (_factory != null)
            {
                _coordinator.SetPaneFactory(_factory);
            }

            _adapter.Start(_coordinator);
        }

        // ----- Ribbon -----

        public string GetCustomUI(string ribbonId)
        {
            string toolbarTab = string.Empty;
            string insertTab = string.Empty;
            try
            {
                string app = CurrentAppName();
                ToolbarSet set = ToolbarsFor(app);
                Settings settings = set == null ? null : new Settings(app);
                if (settings != null && settings.GetBool(SettingNames.ToolbarsEnabled))
                {
                    toolbarTab = ToolbarRibbonXml.Tab(set, settings.GetBool(SettingNames.LoremEnabled));
                    _toolbarsShown = true;
                }

                if (settings != null && settings.GetBool(SettingNames.LoremEnabled))
                {
                    insertTab = InsertTabXml;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Connect.GetCustomUI", ex);
                toolbarTab = string.Empty;
                insertTab = string.Empty;
                _toolbarsShown = false;
            }

            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"OnRibbonLoad\">"
                + "<ribbon><tabs>" + toolbarTab + insertTab + SettingsTabXml + "</tabs></ribbon></customUI>";
        }

        // The OpenToolbars commands for the app, or null when the app has none yet.
        private static ToolbarSet ToolbarsFor(string app)
        {
            switch (app)
            {
                case "Word":
                case "Excel":
                case "PowerPoint":
                    return ToolbarFile.Load(app);
                default:
                    return null;
            }
        }

        // Called by Office, by name, once the ribbon is built.
        public void OnRibbonLoad(Office.IRibbonUI ribbon)
        {
            try
            {
                _ribbon = ribbon;
                if (!_toolbarsShown)
                {
                    return;
                }

                if (_application is Word.Application || _application is Excel.Application
                    || _application is PowerPoint.Application)
                {
                    _startTab = new StartTabWatcher(_application, ribbon, CurrentAppName());
                    _startTab.SelectAtStart();
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Connect.OnRibbonLoad", ex);
            }
        }

        // Called by Office, by name, for the Insert menu's Lorem Ipsum item.
        public void OnLoremIpsum(Office.IRibbonControl control)
        {
            try
            {
                LoremInserter.Insert(_application, new Settings(CurrentAppName()));
            }
            catch (Exception ex)
            {
                Logger.Error("Connect.OnLoremIpsum", ex);
            }
        }

        // Called by Office, by name, for the Lorem Ipsum button on the Insert tab.
        public object GetLoremImage(Office.IRibbonControl control)
        {
            try
            {
                if (_loremPage == null)
                {
                    _loremPage = RibbonPicture.From(LoremIcon.Draw(LargePicture));
                }

                return _loremPage;
            }
            catch (Exception ex)
            {
                Logger.Error("Connect.GetLoremImage", ex);
                return null;
            }
        }

        // Called by Office, by name, for every Settings button. The menu item gets the small gear.
        public object GetSettingsImage(Office.IRibbonControl control)
        {
            try
            {
                if (control.Id == MenuSettingsId)
                {
                    if (_smallGear == null)
                    {
                        _smallGear = RibbonPicture.From(GearIcon.Draw(SmallPicture));
                    }

                    return _smallGear;
                }

                if (_largeGear == null)
                {
                    _largeGear = RibbonPicture.From(GearIcon.Draw(LargePicture));
                }

                return _largeGear;
            }
            catch (Exception ex)
            {
                Logger.Error("Connect.GetSettingsImage", ex);
                return null;
            }
        }

        // Called by Office, by name, for each OpenTools menu heading. The tag is "caption number|fallback".
        public string GetMenuLabel(Office.IRibbonControl control)
        {
            try
            {
                return MenuCaptions.Label(MenuCaption(control));
            }
            catch (Exception ex)
            {
                Logger.Error("Connect.GetMenuLabel", ex);
                return string.Empty;
            }
        }

        public string GetMenuKeytip(Office.IRibbonControl control)
        {
            try
            {
                return MenuCaptions.Keytip(MenuCaption(control));
            }
            catch (Exception ex)
            {
                Logger.Error("Connect.GetMenuKeytip", ex);
                return string.Empty;
            }
        }

        private string MenuCaption(Office.IRibbonControl control)
        {
            string tag = control.Tag ?? string.Empty;
            int split = tag.IndexOf(TagSeparator);
            if (split < 0)
            {
                return tag;
            }

            int captionId;
            int.TryParse(tag.Substring(0, split), out captionId);
            return MenuCaptions.Read(_application, captionId, tag.Substring(split + 1));
        }

        // Called by Office, by name, when the OpenTools button is clicked.
        public void OnOpenSettings(Office.IRibbonControl control)
        {
            try
            {
                TabCoordinator coordinator = _coordinator;
                Action saved = null;
                if (coordinator != null)
                {
                    saved = coordinator.ReloadSettings;
                }

                using (SettingsWindow window = new SettingsWindow(CurrentAppName(), saved))
                {
                    window.ShowDialog(OfficeWindow());
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Connect.OnOpenSettings", ex);
            }
        }

        private static string CurrentAppName()
        {
            using (Process process = Process.GetCurrentProcess())
            {
                switch (process.ProcessName.ToUpperInvariant())
                {
                    case "WINWORD":
                        return "Word";
                    case "EXCEL":
                        return "Excel";
                    case "POWERPNT":
                        return "PowerPoint";
                    case "WINPROJ":
                        return "Project";
                    case "VISIO":
                        return "Visio";
                    default:
                        return null;
                }
            }
        }

        // The Office window the user clicked in, so the settings window opens centered over it.
        private static IWin32Window OfficeWindow()
        {
            IntPtr foreground = NativeMethods.GetForegroundWindow();
            uint processId;
            NativeMethods.GetWindowThreadProcessId(foreground, out processId);
            using (Process process = Process.GetCurrentProcess())
            {
                return processId == (uint)process.Id ? new WindowHandle(foreground) : null;
            }
        }

        private sealed class WindowHandle : IWin32Window
        {
            private readonly IntPtr _handle;

            public WindowHandle(IntPtr handle)
            {
                _handle = handle;
            }

            public IntPtr Handle
            {
                get { return _handle; }
            }
        }

        public void OnDisconnection(int removeMode, ref Array custom)
        {
            Shutdown();
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Shutdown();
        }

        // ----- ICustomTaskPaneConsumer -----

        public void CTPFactoryAvailable(Office.ICTPFactory ctpFactoryInst)
        {
            try
            {
                _factory = ctpFactoryInst;
                if (_coordinator != null)
                {
                    _coordinator.SetPaneFactory(ctpFactoryInst);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Connect.CTPFactoryAvailable", ex);
            }
        }

        // ----- Shutdown -----

        private void Shutdown()
        {
            try
            {
                if (_adapter != null)
                {
                    _adapter.Stop();
                    _adapter = null;
                }

                if (_coordinator != null)
                {
                    _coordinator.Shutdown();
                    _coordinator = null;
                }

                if (_startTab != null)
                {
                    _startTab.Stop();
                    _startTab = null;
                }

                _factory = null;
            }
            catch (Exception ex)
            {
                Logger.Error("Connect.Shutdown", ex);
            }
        }
    }
}
