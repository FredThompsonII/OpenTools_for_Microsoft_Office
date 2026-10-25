// ---------------------------------------------------------------------------------------------------------------
// Name:     TabCoordinator.cs
// Purpose:  Connects one app adapter to the tab list, the tab bar panes, the window stacker and the settings.
//           The adapter reports documents here. Tab bar clicks come back here and are passed to the adapter.
//           Every public method catches its own exceptions, because Office calls them. By default the tab bar is
//           embedded in each Office main window (EmbeddedBarHost), and all documents in one main window share
//           one bar. With the EmbedTabBar setting off, each document window gets a floating bar over an Office
//           task pane. A tab click brings the document's main window to the front. New main windows are created
//           at the size and place of the current one and are kept unseen until their tab bar is ready
//           (NewWindowPlacer), so they do not jump into position. The bars of the other windows stop drawing while
//           a new window is kept unseen, so the new tab is not seen early. They draw again when it is shown.
//           When the SyncRibbonTab setting is on and the user moves to a document in another main window, that
//           window's ribbon is switched to the tab that was selected in the previous one (RibbonTabSync).
//           ReloadSettings lets the settings window apply saved changes at once.
// Revision: 13
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace OpenTools
{
    internal sealed class TabCoordinator
    {
        private readonly IAppAdapter _adapter;
        private readonly bool _allowClose;
        private readonly Settings _settings;
        private readonly WindowStacker _stacker = new WindowStacker();
        private readonly TabList _tabs = new TabList();
        private readonly Dictionary<string, ITabBarHost> _panes = new Dictionary<string, ITabBarHost>();
        private readonly Dictionary<IntPtr, FrameUse> _frames = new Dictionary<IntPtr, FrameUse>();
        private readonly Dictionary<string, Func<object>> _windowGetters = new Dictionary<string, Func<object>>();
        private readonly NewWindowPlacer _placer;
        private bool _barsHeld;
        private bool _closingAll;
        private Office.ICTPFactory _factory;

        // P2 PowerPoint cannot safely close unsaved presentations. Other apps retain the default true.
        public TabCoordinator(IAppAdapter adapter, Settings settings, bool allowClose = true)
        {
            _adapter = adapter;
            _settings = settings;
            _allowClose = allowClose;
            _placer = new NewWindowPlacer(CurrentFrame, StackingWanted);
            _placer.WindowShown += OnWindowShown;
        }

        public TabList Tabs
        {
            get { return _tabs; }
        }

        // Document windows are matched in size and place when single-window mode is on.
        private bool StackingWanted()
        {
            return _settings.GetBool(SettingNames.SingleWindowMode);
        }

        // The main window of the current tab. New main windows are created to match it.
        private IntPtr CurrentFrame()
        {
            TabItem current = _tabs.Find(_tabs.CurrentKey);
            if (current == null)
            {
                return IntPtr.Zero;
            }

            IntPtr frame = NativeMethods.GetAncestor(current.Hwnd, NativeMethods.GA_ROOT);
            return frame != IntPtr.Zero ? frame : current.Hwnd;
        }

        // ----- Pane factory from Office -----

        public void SetPaneFactory(Office.ICTPFactory factory)
        {
            try
            {
                _factory = factory;
                foreach (TabItem item in new List<TabItem>(_tabs.Items))
                {
                    EnsurePane(item.Key);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("TabCoordinator.SetPaneFactory", ex);
            }
        }

        // ----- Reports from the adapter -----

        public void DocumentOpened(string key, IntPtr hwnd, Func<object> getWindow, string fullPath, string name,
            bool isDirty)
        {
            try
            {
                if (_tabs.Find(key) != null)
                {
                    return;
                }

                FrameAnimation.Disable(NativeMethods.GetAncestor(hwnd, NativeMethods.GA_ROOT));
                TabItem item = new TabItem(key, hwnd, fullPath, name);
                item.IsDirty = isDirty;
                _windowGetters[key] = getWindow;
                IntPtr frame = NativeMethods.GetAncestor(hwnd, NativeMethods.GA_ROOT);
                if (_placer.IsHeld(frame))
                {
                    HoldOtherBars(true);
                }

                _tabs.Add(item);
                EnsurePane(key);
                _placer.Reveal(frame);
            }
            catch (Exception ex)
            {
                Logger.Error("TabCoordinator.DocumentOpened", ex);
            }
        }

        // ----- Holding the other bars -----

        private void HoldOtherBars(bool hold)
        {
            _barsHeld = hold;
            Logger.Info((hold ? "Bars held at " : "Bars released at ") + (Environment.TickCount % 1000000) + ".");
            foreach (ITabBarHost host in new List<ITabBarHost>(_panes.Values))
            {
                host.HoldPaint(hold);
            }
        }

        private void OnWindowShown()
        {
            if (_barsHeld)
            {
                HoldOtherBars(false);
            }
        }

        private void HoldFrame(IntPtr frame, bool hold)
        {
            FrameUse use;
            if (_frames.TryGetValue(frame, out use))
            {
                use.Host.HoldPaint(hold);
            }
        }

        public void DocumentClosed(string key)
        {
            try
            {
                ITabBarHost host;
                if (_panes.TryGetValue(key, out host))
                {
                    _panes.Remove(key);
                    host.Dispose();
                }

                _windowGetters.Remove(key);
                _tabs.Remove(key);
            }
            catch (Exception ex)
            {
                Logger.Error("TabCoordinator.DocumentClosed", ex);
            }
        }

        public void DocumentActivated(string key)
        {
            try
            {
                TabItem item = _tabs.Find(key);
                if (item == null)
                {
                    return;
                }

                RefreshSettings();
                TabItem previous = _tabs.Find(_tabs.CurrentKey);
                if (previous != null && previous.Key != key && StackingWanted())
                {
                    _stacker.ApplyBoundsFrom(previous.Hwnd, new IntPtr[] { item.Hwnd });
                }

                _tabs.SetCurrent(key);
                _placer.Reveal(NativeMethods.GetAncestor(item.Hwnd, NativeMethods.GA_ROOT));
                if (previous != null && previous.Key != key && _settings.GetBool(SettingNames.SyncRibbonTab))
                {
                    RibbonTabSync.Copy(NativeMethods.GetAncestor(previous.Hwnd, NativeMethods.GA_ROOT),
                        NativeMethods.GetAncestor(item.Hwnd, NativeMethods.GA_ROOT));
                }
            }
            catch (Exception ex)
            {
                Logger.Error("TabCoordinator.DocumentActivated", ex);
            }
        }

        public void DocumentRenamed(string key, string fullPath, string name)
        {
            try
            {
                _tabs.UpdateFile(key, fullPath, name);
            }
            catch (Exception ex)
            {
                Logger.Error("TabCoordinator.DocumentRenamed", ex);
            }
        }

        public void DocumentDirtyChanged(string key, bool isDirty)
        {
            try
            {
                _tabs.SetDirty(key, isDirty);
            }
            catch (Exception ex)
            {
                Logger.Error("TabCoordinator.DocumentDirtyChanged", ex);
            }
        }

        public void WindowMoved(IntPtr sourceHwnd)
        {
            try
            {
                if (!StackingWanted())
                {
                    return;
                }

                List<IntPtr> targets = new List<IntPtr>();
                foreach (TabItem item in _tabs.Items)
                {
                    targets.Add(item.Hwnd);
                }

                _stacker.ApplyBoundsFrom(sourceHwnd, targets);
            }
            catch (Exception ex)
            {
                Logger.Error("TabCoordinator.WindowMoved", ex);
            }
        }

        // ----- Requests from the tab bar -----

        private void OnActivateRequested(string key)
        {
            try
            {
                TabItem item = _tabs.Find(key);
                if (item == null)
                {
                    return;
                }

                TabItem previous = _tabs.Find(_tabs.CurrentKey);
                if (previous != null && previous.Key != key && StackingWanted())
                {
                    _stacker.ApplyBoundsFrom(previous.Hwnd, new IntPtr[] { item.Hwnd });
                }

                _adapter.ActivateDocument(key);

                // A document window inside one main window is not top level, so its main window is raised.
                IntPtr top = NativeMethods.GetAncestor(item.Hwnd, NativeMethods.GA_ROOT);
                _stacker.BringToFront(top != IntPtr.Zero ? top : item.Hwnd);
            }
            catch (Exception ex)
            {
                Logger.Error("TabCoordinator.OnActivateRequested", ex);
            }
        }

        private void OnCloseRequested(string key)
        {
            if (!_allowClose)
            {
                return;
            }

            try
            {
                _adapter.CloseDocument(key);
            }
            catch (Exception ex)
            {
                Logger.Error("TabCoordinator.OnCloseRequested", ex);
            }
        }

        private void OnNewDocumentRequested()
        {
            try
            {
                _adapter.NewDocument();
            }
            catch (Exception ex)
            {
                Logger.Error("TabCoordinator.OnNewDocumentRequested", ex);
            }
        }

        // Asks what to close, then closes every document, last tab first. Each tab is removed as soon as its
        // document is closed. Word still asks about unsaved documents. If the user cancels one of those prompts,
        // closing stops there, and the app is not closed even if that was asked for.
        private void OnCloseAllRequested()
        {
            if (!_allowClose || _closingAll)
            {
                return;
            }

            _closingAll = true;
            try
            {
                List<string> keys = new List<string>();
                foreach (TabItem item in _tabs.Items)
                {
                    keys.Add(item.Key);
                }

                if (keys.Count == 0)
                {
                    return;
                }

                CloseAllChoice choice = CloseAllDialog.Ask(new OwnerWindow(CurrentFrame()), keys.Count,
                    _adapter.AppName);
                if (choice == CloseAllChoice.Cancel)
                {
                    return;
                }

                for (int i = keys.Count - 1; i >= 0; i--)
                {
                    if (!_adapter.CloseDocument(keys[i]))
                    {
                        Logger.Info("Close all stopped. A document was kept.");
                        return;
                    }

                    DocumentClosed(keys[i]);
                }

                if (choice == CloseAllChoice.CloseDocumentsAndApp)
                {
                    QuitSoon();
                }
            }
            catch (Exception ex)
            {
                Logger.Error("TabCoordinator.OnCloseAllRequested", ex);
            }
            finally
            {
                _closingAll = false;
            }
        }

        // The app is closed from a short timer, so the click that asked for it has finished first.
        private void QuitSoon()
        {
            Timer timer = new Timer();
            timer.Interval = 100;
            timer.Tick += delegate
            {
                timer.Stop();
                timer.Dispose();
                try
                {
                    _adapter.QuitApplication();
                }
                catch (Exception ex)
                {
                    Logger.Error("TabCoordinator.QuitSoon", ex);
                }
            };
            timer.Start();
        }

        private sealed class OwnerWindow : IWin32Window
        {
            private readonly IntPtr _handle;

            public OwnerWindow(IntPtr handle)
            {
                _handle = handle;
            }

            public IntPtr Handle
            {
                get { return _handle; }
            }
        }

        // ----- Panes -----

        private void EnsurePane(string key)
        {
            if (_panes.ContainsKey(key))
            {
                return;
            }

            if (_settings.GetBool(SettingNames.EmbedTabBar))
            {
                EnsureEmbedded(key);
                return;
            }

            if (_factory == null)
            {
                return;
            }

            Func<object> getWindow;
            if (!_windowGetters.TryGetValue(key, out getWindow))
            {
                return;
            }

            object window = getWindow();
            try
            {
                _panes[key] = TaskPaneHost.Create(_factory, window, _tabs, _settings, OnActivateRequested,
                    OnCloseRequested, OnNewDocumentRequested, OnCloseAllRequested, _allowClose);
            }
            catch (Exception ex)
            {
                Logger.Error("TabCoordinator.EnsurePane " + key, ex);
            }
            finally
            {
                if (window != null && Marshal.IsComObject(window))
                {
                    Marshal.ReleaseComObject(window);
                }
            }
        }

        // ----- Embedded bars: one per Office main window, shared by its documents -----

        private sealed class FrameUse
        {
            public EmbeddedBarHost Host;
            public int Users;
        }

        private sealed class FrameLease : ITabBarHost
        {
            private readonly TabCoordinator _owner;
            private readonly IntPtr _frame;
            private bool _released;

            public FrameLease(TabCoordinator owner, IntPtr frame)
            {
                _owner = owner;
                _frame = frame;
            }

            public void Refresh()
            {
                _owner.RefreshFrame(_frame);
            }

            public void HoldPaint(bool hold)
            {
                _owner.HoldFrame(_frame, hold);
            }

            public void Dispose()
            {
                if (!_released)
                {
                    _released = true;
                    _owner.ReleaseFrame(_frame);
                }
            }
        }

        private void EnsureEmbedded(string key)
        {
            TabItem item = _tabs.Find(key);
            if (item == null)
            {
                return;
            }

            IntPtr frame = NativeMethods.GetAncestor(item.Hwnd, NativeMethods.GA_ROOT);
            if (frame == IntPtr.Zero)
            {
                frame = item.Hwnd;
            }

            try
            {
                FrameUse use;
                if (!_frames.TryGetValue(frame, out use))
                {
                    EmbeddedBarHost host = EmbeddedBarHost.Create(frame, _tabs, _settings, OnActivateRequested,
                        OnCloseRequested, OnNewDocumentRequested, OnCloseAllRequested, _allowClose);
                    if (host == null)
                    {
                        return;
                    }

                    use = new FrameUse();
                    use.Host = host;
                    _frames[frame] = use;
                }

                use.Users++;
                _panes[key] = new FrameLease(this, frame);
            }
            catch (Exception ex)
            {
                Logger.Error("TabCoordinator.EnsureEmbedded " + key, ex);
            }
        }

        private void RefreshFrame(IntPtr frame)
        {
            FrameUse use;
            if (_frames.TryGetValue(frame, out use))
            {
                use.Host.Refresh();
            }
        }

        private void ReleaseFrame(IntPtr frame)
        {
            FrameUse use;
            if (_frames.TryGetValue(frame, out use))
            {
                use.Users--;
                if (use.Users <= 0)
                {
                    _frames.Remove(frame);
                    use.Host.Dispose();
                }
            }
        }

        // Called by the settings window after it saves, so this app shows the change without a document switch.
        public void ReloadSettings()
        {
            try
            {
                RefreshSettings();
            }
            catch (Exception ex)
            {
                Logger.Error("TabCoordinator.ReloadSettings", ex);
            }
        }

        private void RefreshSettings()
        {
            _settings.Reload();
            foreach (ITabBarHost host in new List<ITabBarHost>(_panes.Values))
            {
                host.Refresh();
            }
        }

        // ----- Shutdown -----

        public void Shutdown()
        {
            try
            {
                foreach (ITabBarHost host in new List<ITabBarHost>(_panes.Values))
                {
                    host.Dispose();
                }

                _panes.Clear();
                foreach (FrameUse use in new List<FrameUse>(_frames.Values))
                {
                    use.Host.Dispose();
                }

                _frames.Clear();
                _placer.WindowShown -= OnWindowShown;
                _placer.Dispose();
                FrameAnimation.RestoreAll();
                _windowGetters.Clear();
            }
            catch (Exception ex)
            {
                Logger.Error("TabCoordinator.Shutdown", ex);
            }
        }
    }
}
