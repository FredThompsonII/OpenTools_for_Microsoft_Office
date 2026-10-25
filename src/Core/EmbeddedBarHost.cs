// ---------------------------------------------------------------------------------------------------------------
// Name:     EmbeddedBarHost.cs
// Purpose:  Puts the tab bar inside an Office main window, between the ribbon and the document area, with no task
//           pane. The tab bar is a child window of the Office main window. The document area (the child window
//           of class _WwF in Word) is moved down by the height of the tab bar. A hook on the document area edits
//           each move or resize Office asks for, so the shift survives Office's own layout. A timer re-finds the
//           document area if Office replaces it. An event brake stops runaway layout loops. The first run logs
//           the window tree of the main window, so a layout that differs from Word's can be diagnosed. One host
//           exists per Office main window and is shared by every document in that window. A move or resize that
//           equals where the document area already is was made by Office from a position it read back, so it
//           is left alone. The individual move and resize requests are not logged. When the host is removed, the
//           strip where the bar was is redrawn so no old tab edges are left on screen. A short gap in the
//           document area's visibility is Word repainting, so the bar is hidden only when the gap lasts. Each
//           hide and each move of the bar is logged, to find out what makes the tabs blink.
// Revision: 10
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace OpenTools
{
    internal sealed class EmbeddedBarHost : ITabBarHost
    {
        private const int PollMilliseconds = 250;
        private const int PollsBeforeTreeLog = 8;
        private const int BrakeLimit = 40;
        private const int BrakeWindowMilliseconds = 2000;
        private const int BrakePauseMilliseconds = 5000;
        private const int PlacementLogLimit = 6;
        private const int MoveLogLimit = 20;
        private const int HideDelayMilliseconds = 400;

        private readonly IntPtr _root;
        private readonly TabBarControl _bar;
        private readonly Timer _timer;
        private readonly System.Collections.Generic.Queue<int> _recent = new System.Collections.Generic.Queue<int>();
        private AreaHook _hook;
        private IntPtr _area = IntPtr.Zero;
        private Rectangle _barBounds = Rectangle.Empty;
        private bool _barVisible;
        private int _shift;
        private int _polls;
        private int _placementLogs;
        private int _moveLogs;
        private int _hideSince;
        private int _brakedAt;
        private bool _braked;
        private bool _applying;
        private bool _treeLogged;
        private bool _disposed;

        private EmbeddedBarHost(IntPtr root, TabBarControl bar)
        {
            _root = root;
            _bar = bar;
            _timer = new Timer();
            _timer.Interval = PollMilliseconds;
            _timer.Tick += OnPoll;
        }

        // ----- Creating -----

        public static EmbeddedBarHost Create(IntPtr root, TabList list, Settings settings,
            Action<string> onActivate, Action<string> onClose, Action onNewDocument, Action onCloseAll,
            bool allowClose)
        {
            if (root == IntPtr.Zero || !NativeMethods.IsWindow(root))
            {
                return null;
            }

            TabBarControl bar = new TabBarControl(list, settings, allowClose);
            try
            {
                bar.ActivateRequested += onActivate;
                bar.CloseRequested += onClose;
                bar.NewDocumentRequested += onNewDocument;
                bar.CloseAllRequested += onCloseAll;
                bar.UseHostParent(root);
                bar.SetBounds(0, 0, 800, 30);
                IntPtr created = bar.Handle;
                NativeMethods.ShowWindow(created, NativeMethods.SW_HIDE);

                EmbeddedBarHost host = new EmbeddedBarHost(root, bar);
                host.TryInstall();
                host._timer.Start();
                return host;
            }
            catch (Exception)
            {
                bar.Dispose();
                throw;
            }
        }

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
            UpdateBar();
        }

        // ----- Finding and hooking the document area -----

        private void TryInstall()
        {
            if (_disposed || _area != IntPtr.Zero)
            {
                return;
            }

            IntPtr area = FindDocumentArea(_root);
            NativeMethods.RECT rect;
            if (area == IntPtr.Zero || !NativeMethods.GetWindowRect(area, out rect))
            {
                return;
            }

            _bar.Width = Math.Max(100, rect.Right - rect.Left);
            _shift = Math.Max(1, _bar.WantedHeight);
            _area = area;
            _hook = new AreaHook(this);
            _hook.AssignHandle(area);
            MoveArea(_shift);
            _bar.HeightWanted += OnHeightWanted;
            UpdateBar();
            Logger.Info("Tab bar embedded in window " + _root.ToInt64().ToString("X") + ". Document area "
                + area.ToInt64().ToString("X") + " moved down " + _shift + " pixels.");
        }

        private static IntPtr FindDocumentArea(IntPtr root)
        {
            for (IntPtr child = NativeMethods.GetWindow(root, NativeMethods.GW_CHILD); child != IntPtr.Zero;
                child = NativeMethods.GetWindow(child, NativeMethods.GW_HWNDNEXT))
            {
                if (ClassOf(child) == AppProfile.DocumentAreaClass && NativeMethods.IsWindowVisible(child))
                {
                    return child;
                }
            }

            for (IntPtr child = NativeMethods.GetWindow(root, NativeMethods.GW_CHILD); child != IntPtr.Zero;
                child = NativeMethods.GetWindow(child, NativeMethods.GW_HWNDNEXT))
            {
                for (IntPtr inner = NativeMethods.GetWindow(child, NativeMethods.GW_CHILD); inner != IntPtr.Zero;
                    inner = NativeMethods.GetWindow(inner, NativeMethods.GW_HWNDNEXT))
                {
                    if (ClassOf(inner) == AppProfile.DocumentAreaClass && NativeMethods.IsWindowVisible(inner))
                    {
                        return inner;
                    }
                }
            }

            return IntPtr.Zero;
        }

        private static string ClassOf(IntPtr window)
        {
            StringBuilder text = new StringBuilder(128);
            return NativeMethods.GetClassName(window, text, text.Capacity) > 0 ? text.ToString() : string.Empty;
        }

        private void LogWindowTree()
        {
            _treeLogged = true;
            StringBuilder text = new StringBuilder("Window tree of " + _root.ToInt64().ToString("X") + ":");
            int shown = 0;
            for (IntPtr child = NativeMethods.GetWindow(_root, NativeMethods.GW_CHILD); child != IntPtr.Zero
                && shown < 40; child = NativeMethods.GetWindow(child, NativeMethods.GW_HWNDNEXT))
            {
                shown++;
                text.Append(" [").Append(Describe(child));
                int inner = 0;
                for (IntPtr sub = NativeMethods.GetWindow(child, NativeMethods.GW_CHILD); sub != IntPtr.Zero
                    && inner < 8; sub = NativeMethods.GetWindow(sub, NativeMethods.GW_HWNDNEXT))
                {
                    inner++;
                    text.Append(" > ").Append(Describe(sub));
                }

                text.Append("]");
            }

            Logger.Info(text.ToString());
        }

        private string Describe(IntPtr window)
        {
            NativeMethods.RECT rect;
            NativeMethods.GetWindowRect(window, out rect);
            return ClassOf(window) + " " + (rect.Right - rect.Left) + "x" + (rect.Bottom - rect.Top)
                + (NativeMethods.IsWindowVisible(window) ? "" : " hidden");
        }

        // ----- Moving the document area -----

        // The document area's place inside its parent window, in the parent's client pixels.
        private bool TryGetAreaBounds(out Rectangle bounds)
        {
            bounds = Rectangle.Empty;
            NativeMethods.RECT rect;
            if (_area == IntPtr.Zero || !NativeMethods.GetWindowRect(_area, out rect))
            {
                return false;
            }

            NativeMethods.POINT corner = new NativeMethods.POINT();
            corner.X = rect.Left;
            corner.Y = rect.Top;
            NativeMethods.ScreenToClient(NativeMethods.GetAncestor(_area, NativeMethods.GA_PARENT), ref corner);
            bounds = new Rectangle(corner.X, corner.Y, rect.Right - rect.Left, rect.Bottom - rect.Top);
            return true;
        }

        // Moves the document area down by delta pixels and makes it that much shorter. The hook ignores this
        // move, because it is the correction itself.
        private void MoveArea(int delta)
        {
            Rectangle now;
            if (delta == 0 || !TryGetAreaBounds(out now))
            {
                return;
            }

            _applying = true;
            try
            {
                NativeMethods.SetWindowPos(_area, IntPtr.Zero, now.X, now.Y + delta, now.Width,
                    Math.Max(0, now.Height - delta), NativeMethods.SWP_NOZORDER | NativeMethods.SWP_NOACTIVATE);
            }
            finally
            {
                _applying = false;
            }
        }

        private void OnAreaChanging(IntPtr lParam)
        {
            if (_applying || _shift <= 0 || _disposed)
            {
                return;
            }

            NativeMethods.WINDOWPOS pos = (NativeMethods.WINDOWPOS)Marshal.PtrToStructure(lParam,
                typeof(NativeMethods.WINDOWPOS));
            Rectangle now;
            bool known = TryGetAreaBounds(out now);

            if ((pos.Flags & NativeMethods.SWP_NOMOVE) == 0 && !(known && pos.Y == now.Y))
            {
                pos.Y += _shift;
            }

            if ((pos.Flags & NativeMethods.SWP_NOSIZE) == 0 && !(known && pos.Cy == now.Height))
            {
                pos.Cy = Math.Max(0, pos.Cy - _shift);
            }

            Marshal.StructureToPtr(pos, lParam, false);
        }

        private void OnAreaChanged()
        {
            if (_applying || _disposed || TooManyEvents())
            {
                return;
            }

            UpdateBar();
        }

        private void OnHeightWanted(int height)
        {
            try
            {
                if (_disposed || height <= 0 || height == _shift || TooManyEvents())
                {
                    return;
                }

                int delta = height - _shift;
                _shift = height;
                MoveArea(delta);
                UpdateBar();
            }
            catch (Exception ex)
            {
                Logger.Error("EmbeddedBarHost.OnHeightWanted", ex);
            }
        }

        // ----- Placing the tab bar -----

        // The bar sits directly above the document area, with the same left edge and width.
        private void UpdateBar()
        {
            if (_disposed || _area == IntPtr.Zero)
            {
                return;
            }

            Rectangle areaBounds;
            bool visible = NativeMethods.IsWindowVisible(_area) && !NativeMethods.IsIconic(_root);
            if (!visible || !TryGetAreaBounds(out areaBounds))
            {
                HideBarSoon();
                return;
            }

            _hideSince = 0;
            NativeMethods.RECT screen;
            NativeMethods.GetWindowRect(_area, out screen);
            NativeMethods.POINT corner = new NativeMethods.POINT();
            corner.X = screen.Left;
            corner.Y = screen.Top;
            NativeMethods.ScreenToClient(_root, ref corner);
            Rectangle bounds = new Rectangle(corner.X, corner.Y - _shift, areaBounds.Width, _shift);
            if (bounds.Width <= 0 || bounds.Height <= 0)
            {
                HideBar();
                return;
            }

            if (bounds != _barBounds)
            {
                NativeMethods.SetWindowPos(_bar.Handle, IntPtr.Zero, bounds.X, bounds.Y, bounds.Width, bounds.Height,
                    NativeMethods.SWP_NOACTIVATE);
                if (_barBounds != Rectangle.Empty && _moveLogs < MoveLogLimit)
                {
                    _moveLogs++;
                    Logger.Info("Tab bar moved at " + (Environment.TickCount % 1000000) + " from "
                        + _barBounds.X + "," + _barBounds.Y + " " + _barBounds.Width + "x" + _barBounds.Height
                        + " to " + bounds.X + "," + bounds.Y + " " + bounds.Width + "x" + bounds.Height + ".");
                }

                _barBounds = bounds;
                if (_placementLogs < PlacementLogLimit)
                {
                    _placementLogs++;
                    Logger.Info("Tab bar at " + bounds.X + "," + bounds.Y + " size " + bounds.Width + "x"
                        + bounds.Height + ".");
                }
            }

            if (!_barVisible)
            {
                NativeMethods.ShowWindow(_bar.Handle, NativeMethods.SW_SHOWNOACTIVATE);
                _barVisible = true;
            }
        }

        // A short gap in the document area's visibility is Word repainting. Only a gap that lasts is hidden.
        private void HideBarSoon()
        {
            int now = Environment.TickCount | 1;
            if (_hideSince == 0)
            {
                _hideSince = now;
            }
            else if (_barVisible && now - _hideSince > HideDelayMilliseconds)
            {
                Logger.Info("Tab bar hidden at " + (Environment.TickCount % 1000000) + ".");
                HideBar();
            }
        }

        private void HideBar()
        {
            if (_barVisible)
            {
                // Office may have destroyed the main window (and the bar with it) already. Asking a dead
                // control for its handle would try to rebuild the window, which fails.
                if (_bar.IsHandleCreated && NativeMethods.IsWindow(_bar.Handle))
                {
                    NativeMethods.ShowWindow(_bar.Handle, NativeMethods.SW_HIDE);
                }

                _barVisible = false;
            }
        }

        // ----- Timer -----

        private void OnPoll(object sender, EventArgs e)
        {
            try
            {
                if (_disposed)
                {
                    return;
                }

                if (_area != IntPtr.Zero && !NativeMethods.IsWindow(_area))
                {
                    Logger.Info("The document area was replaced by Office. Looking for the new one.");
                    DetachHook();
                    _area = IntPtr.Zero;
                    _barBounds = Rectangle.Empty;
                    _treeLogged = false;
                    _polls = 0;
                }

                if (_area == IntPtr.Zero)
                {
                    _polls++;
                    TryInstall();
                    if (_area == IntPtr.Zero && _polls >= PollsBeforeTreeLog && !_treeLogged)
                    {
                        Logger.Info("The document area (class " + AppProfile.DocumentAreaClass + ") was not found.");
                        LogWindowTree();
                    }

                    return;
                }

                UpdateBar();
            }
            catch (Exception ex)
            {
                Logger.Error("EmbeddedBarHost.OnPoll", ex);
            }
        }

        // Returns true when layout events should be ignored. More than BrakeLimit events inside
        // BrakeWindowMilliseconds switches the host off for BrakePauseMilliseconds and logs one line.
        private bool TooManyEvents()
        {
            int now = Environment.TickCount;
            if (_braked)
            {
                if (now - _brakedAt < BrakePauseMilliseconds)
                {
                    return true;
                }

                _braked = false;
                _recent.Clear();
                Logger.Info("Embedded bar event brake released.");
            }

            _recent.Enqueue(now);
            while (_recent.Count > 0 && now - _recent.Peek() > BrakeWindowMilliseconds)
            {
                _recent.Dequeue();
            }

            if (_recent.Count > BrakeLimit)
            {
                _braked = true;
                _brakedAt = now;
                Logger.Info("Embedded bar event brake on: more than " + BrakeLimit + " layout events in "
                    + BrakeWindowMilliseconds + " ms.");
                return true;
            }

            return false;
        }

        // ----- Cleanup -----

        private void DetachHook()
        {
            try
            {
                if (_hook != null)
                {
                    _hook.ReleaseHandle();
                    _hook = null;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("EmbeddedBarHost.DetachHook", ex);
            }
        }

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
                HideBar();
                if (_area != IntPtr.Zero && NativeMethods.IsWindow(_area))
                {
                    int shift = _shift;
                    _shift = 0;
                    MoveArea(-shift);
                }

                DetachHook();
                _bar.Dispose();
                if (NativeMethods.IsWindow(_root))
                {
                    NativeMethods.RedrawWindow(_root, IntPtr.Zero, IntPtr.Zero, NativeMethods.RDW_INVALIDATE
                        | NativeMethods.RDW_ERASE | NativeMethods.RDW_ALLCHILDREN | NativeMethods.RDW_UPDATENOW);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("EmbeddedBarHost.Dispose", ex);
            }
        }

        // ----- Hook on the document area -----

        private sealed class AreaHook : NativeWindow
        {
            private readonly EmbeddedBarHost _owner;

            public AreaHook(EmbeddedBarHost owner)
            {
                _owner = owner;
            }

            protected override void WndProc(ref Message m)
            {
                try
                {
                    if (m.Msg == NativeMethods.WM_WINDOWPOSCHANGING)
                    {
                        _owner.OnAreaChanging(m.LParam);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error("AreaHook.Changing", ex);
                }

                base.WndProc(ref m);

                try
                {
                    if (m.Msg == NativeMethods.WM_WINDOWPOSCHANGED)
                    {
                        _owner.OnAreaChanged();
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error("AreaHook.Changed", ex);
                }
            }
        }
    }
}
