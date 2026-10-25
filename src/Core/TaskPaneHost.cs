// ---------------------------------------------------------------------------------------------------------------
// Name:     TaskPaneHost.cs
// Purpose:  The fallback tab bar host, used when the EmbedTabBar setting is off. Owns one Office custom task pane
//           and the floating window that holds the tab bar. Office always draws a "Tabs" title strip on a pane,
//           so the pane is kept only as tall as the tab bar, and the real tab bar floats over it in a borderless
//           window (TabOverlayWindow). The window starts to the right of the word "Tabs" and stops before
//           Office's own pane buttons. A short timer keeps the window on top of the pane and hides it when the
//           pane or the Office window is hidden or minimized. The dock position is set only when it changes,
//           because setting it again makes Office re-dock the pane. One host exists per document window.
// Revision: 7
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace OpenTools
{
    internal sealed class TaskPaneHost : ITabBarHost
    {
        private const string PaneProgId = "OpenTools.TabPane";
        private const string PaneTitle = "Tabs";
        private const int FollowMilliseconds = 60;
        private const int ControlSliver = 3;
        private const int TitleGap = 28;
        private const int MinOverlayWidth = 80;
        private const int MaxHeightAttempts = 4;
        private const int PlacementLogLimit = 6;

        private readonly Office.CustomTaskPane _pane;
        private readonly TabPaneControl _host;
        private readonly TabBarControl _bar;
        private readonly TabOverlayWindow _overlay;
        private readonly Settings _settings;
        private readonly Timer _timer;
        private int _lastTarget = -1;
        private int _attempts;
        private int _titleWidth;
        private int _placementLogs;
        private bool _gaveUpOnHeight;
        private bool _tickErrorLogged;
        private bool _placed;
        private bool _placedBelow;
        private bool _disposed;

        private TaskPaneHost(Office.CustomTaskPane pane, TabPaneControl host, TabBarControl bar, Settings settings)
        {
            _pane = pane;
            _host = host;
            _bar = bar;
            _settings = settings;
            _overlay = new TabOverlayWindow(bar);
            _bar.HeightWanted += OnHeightWanted;
            _titleWidth = TextRenderer.MeasureText(PaneTitle, SystemFonts.MessageBoxFont).Width;

            _timer = new Timer();
            _timer.Interval = FollowMilliseconds;
            _timer.Tick += OnTick;
        }

        // ----- Creating -----

        public static TaskPaneHost Create(Office.ICTPFactory factory, object window, TabList list,
            Settings settings, Action<string> onActivate, Action<string> onClose, Action onNewDocument,
            Action onCloseAll, bool allowClose)
        {
            Office.CustomTaskPane pane = factory.CreateCTP(PaneProgId, PaneTitle, window);
            TabBarControl bar = null;
            try
            {
                TabPaneControl host = pane.ContentControl as TabPaneControl;
                if (host == null)
                {
                    throw new InvalidOperationException("Office did not create the tab bar control.");
                }

                bar = new TabBarControl(list, settings, allowClose);
                bar.ActivateRequested += onActivate;
                bar.CloseRequested += onClose;
                bar.NewDocumentRequested += onNewDocument;
                bar.CloseAllRequested += onCloseAll;

                TaskPaneHost result = new TaskPaneHost(pane, host, bar, settings);
                bar = null;
                result.ApplyPlacement();
                pane.Visible = true;
                result._timer.Start();
                return result;
            }
            catch (Exception)
            {
                if (bar != null)
                {
                    bar.Dispose();
                }

                ReleasePane(pane);
                throw;
            }
        }

        // ----- Placement of the pane -----

        public void HoldPaint(bool hold)
        {
            if (!_disposed)
            {
                _bar.HoldPaint(hold);
            }
        }

        public void Refresh()
        {
            if (_disposed)
            {
                return;
            }

            _bar.ApplySettings();
            ApplyPlacement();
        }

        private void ApplyPlacement()
        {
            bool below = _settings.GetPosition() == TabBarPosition.Below;
            if (!_placed || below != _placedBelow)
            {
                _pane.DockPosition = below
                    ? Office.MsoCTPDockPosition.msoCTPDockPositionBottom
                    : Office.MsoCTPDockPosition.msoCTPDockPositionTop;
                _pane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                _placed = true;
                _placedBelow = below;
                Logger.Info("Pane docked " + (below ? "below" : "above") + ".");
            }

            ApplyHeight(_bar.WantedHeight);
        }

        private void OnHeightWanted(int height)
        {
            try
            {
                ApplyHeight(height);
            }
            catch (Exception ex)
            {
                Logger.Error("TaskPaneHost.OnHeightWanted", ex);
            }
        }

        // The pane is made as tall as the tab bar plus a few pixels, so the pane still has a real control area
        // whose position can be read. If Office will not take the height, a few tries are made and then it stops,
        // so a pane that Office will not resize cannot cause a loop.
        private void ApplyHeight(int wanted)
        {
            if (_disposed || wanted <= 0)
            {
                return;
            }

            int target = wanted + Scale(ControlSliver);
            if (target != _lastTarget)
            {
                _lastTarget = target;
                _attempts = 0;
                _gaveUpOnHeight = false;
            }

            if (_pane.Height == target)
            {
                _attempts = 0;
                return;
            }

            if (_attempts >= MaxHeightAttempts)
            {
                if (!_gaveUpOnHeight)
                {
                    _gaveUpOnHeight = true;
                    Logger.Info("Pane height stays at " + _pane.Height + " when " + target + " was asked for.");
                }

                return;
            }

            _attempts++;
            _pane.Height = target;
        }

        // ----- Following the pane -----

        private void OnTick(object sender, EventArgs e)
        {
            try
            {
                FollowPane();
            }
            catch (Exception ex)
            {
                if (!_tickErrorLogged)
                {
                    _tickErrorLogged = true;
                    Logger.Error("TaskPaneHost.OnTick", ex);
                }
            }
        }

        private void FollowPane()
        {
            if (_disposed)
            {
                return;
            }

            if (!_host.IsHandleCreated)
            {
                _overlay.HideWindow();
                return;
            }

            IntPtr control = _host.Handle;
            if (!NativeMethods.IsWindowVisible(control))
            {
                _overlay.HideWindow();
                return;
            }

            IntPtr root = NativeMethods.GetAncestor(control, NativeMethods.GA_ROOT);
            NativeMethods.RECT rect;
            if (root == IntPtr.Zero || NativeMethods.IsIconic(root) || !NativeMethods.GetWindowRect(control, out rect))
            {
                _overlay.HideWindow();
                return;
            }

            int paneHeight = _pane.Height;
            int left = rect.Left + LeftOffset();
            int right = rect.Right - Scale(_settings.GetInt(SettingNames.OverlayRightReserve));
            if (paneHeight <= 0 || right - left < Scale(MinOverlayWidth))
            {
                _overlay.HideWindow();
                return;
            }

            Rectangle bounds = new Rectangle(left, rect.Bottom - paneHeight, right - left, paneHeight);
            _overlay.ShowAt(bounds, root);
            if (_placementLogs < PlacementLogLimit)
            {
                _placementLogs++;
                Logger.Info("Tab window at " + bounds.X + "," + bounds.Y + " size " + bounds.Width + "x"
                    + bounds.Height + ". Pane control " + (rect.Right - rect.Left) + "x" + (rect.Bottom - rect.Top)
                    + ".");
            }
        }

        private int LeftOffset()
        {
            int fixedOffset = _settings.GetInt(SettingNames.OverlayLeftOffset);
            return fixedOffset > 0 ? Scale(fixedOffset) : _titleWidth + Scale(TitleGap);
        }

        private int Scale(int value)
        {
            return (int)Math.Round(value * _bar.DeviceDpi / 96.0);
        }

        // ----- Cleanup -----

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            try
            {
                _timer.Stop();
                _timer.Dispose();
                _bar.HeightWanted -= OnHeightWanted;
                _overlay.Dispose();
                _pane.Delete();
            }
            catch (Exception ex)
            {
                Logger.Error("TaskPaneHost.Dispose", ex);
            }
            finally
            {
                ReleasePane(_pane);
            }
        }

        private static void ReleasePane(Office.CustomTaskPane pane)
        {
            try
            {
                if (pane != null && Marshal.IsComObject(pane))
                {
                    Marshal.ReleaseComObject(pane);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("TaskPaneHost.ReleasePane", ex);
            }
        }
    }
}
