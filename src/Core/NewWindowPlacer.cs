// ---------------------------------------------------------------------------------------------------------------
// Name:     NewWindowPlacer.cs
// Purpose:  Stops a new Office main window from being seen while it is still being set up. A window hook on the
//           Office user-interface thread sees each window as it is being created. When it is a new top-level
//           Office main window (class OpusApp), the position and size in the create request are changed to match
//           the main window of the current tab, and the window is created as a layered window. A layered window
//           shows nothing until its opacity is set, so the new window stays unseen while Office builds it and the
//           tab bar is added. A timer reveals any window that was not revealed within a second and a half, so
//           a window can never stay hidden. Word and Excel use an opaque stage before removing the layered
//           style. PowerPoint removes layering, requests one non-immediate redraw, then uncloaks the frame.
//           PowerPoint cloaking is idempotent within each creation hold. Its reveal waits until the editing pane
//           and slide child are visible with stable bounds across two checks; the existing timeout is retained.
//           Separate Office border-effect windows are observed but never modified or used as a release gate.
//           PowerPoint activates the new frame while it is still cloaked, so the frame the user is looking at
//           receives WM_NCACTIVATE(FALSE) and Office redraws its border in the inactive style, which looks like
//           the border vanishing. While a new frame is held, that frame is subclassed and WM_NCACTIVATE(FALSE)
//           is passed on as TRUE, so it keeps its active look until the new frame is shown. WM_ACTIVATE is not
//           touched, so Office's own activation tracking is unchanged. The subclass is removed at reveal. The old
//           frame then sits behind the new one and resumes normal drawing on its next activation change.
//           Bounded visibility snapshots record DWM cloak state and surrounding border-window visibility.
//           Windows refuses window attributes while a window is still being created,
//           so a second hook sees the message that is sent just before a new window is shown. At that point the
//           window animations are turned off and the window is cloaked (FrameAnimation). The window stays
//           cloaked while its bar is prepared. It is uncloaked when the layered style is removed, so the viewer
//           sees one step. Until then the window is click-through (WS_EX_TRANSPARENT), so
//           the mouse pointer does not react to a window that cannot be seen. For PowerPoint, the creation
//           request is checked for WS_CHILD rather than a nonzero owner, and class atoms are resolved by HWND.
//           Rejection reasons are logged with a strict limit and unverified windows are left untouched.
//           Animations are also turned off when a main window is first activated. A failure inside a hook is
//           logged once and then ignored. These are hooks on this thread only. They never reach other programs.
// Revision: 26 (rev33 code plus the active-look hold only)
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OpenTools
{
    internal sealed class NewWindowPlacer : IDisposable
    {
        private const int PlacementLogLimit = 6;
        private const int TimerMilliseconds = 100;
        private const int MaxHiddenMilliseconds = 1500;
        private const int LayeredStyleMilliseconds = 300;
        private const int StepLogLimit = 80;
        private const int PowerPointSkipLogLimit = 18;
        private const int PowerPointOtherClassLogLimit = 3;
        private const int PowerPointVisibilityLogLimit = 5;
        private const int PowerPointReadinessLogLimit = 12;
        private const uint GW_HWNDFIRST = 0;
        private const int WM_NCACTIVATE = 0x0086;
        private const int DWMWA_CLOAKED = 14;
        private const int DWM_CLOAKED_APP = 1;
        private const int WS_CHILD = 0x40000000;
        private const uint RDW_FRAME = 0x0400; // Winuser.h: invalidate nonclient frame for later repaint.

        private readonly Func<IntPtr> _referenceFrame;
        private readonly Func<bool> _isWanted;
        private readonly NativeMethods.HookProc _callback;
        private readonly Dictionary<IntPtr, int> _hidden = new Dictionary<IntPtr, int>();
        private readonly Dictionary<IntPtr, int> _revealed = new Dictionary<IntPtr, int>();
        private readonly HashSet<IntPtr> _powerPointCloaked = new HashSet<IntPtr>();
        private readonly HashSet<IntPtr> _powerPointRevealPending = new HashSet<IntPtr>();
        private readonly Dictionary<IntPtr, ActiveLook> _activeLooks = new Dictionary<IntPtr, ActiveLook>();
        private readonly Dictionary<IntPtr, string> _powerPointReadyLayouts =
            new Dictionary<IntPtr, string>();
        private readonly Dictionary<IntPtr, int> _inspectAfterReveal = new Dictionary<IntPtr, int>();
        private readonly Timer _timer;
        private readonly NativeMethods.HookProc _showCallback;
        private IntPtr _hook = IntPtr.Zero;
        private IntPtr _showHook = IntPtr.Zero;
        private int _placementLogs;
        private int _stepLogs;
        private int _powerPointSkipLogs;
        private int _powerPointOtherClassLogs;
        private int _powerPointVisibilityLogs;
        private int _powerPointReadinessLogs;
        private bool _powerPointModeLogged;
        private bool _errorLogged;
        private bool _atomLogged;

        public NewWindowPlacer(Func<IntPtr> referenceFrame, Func<bool> isWanted)
        {
            _referenceFrame = referenceFrame;
            _isWanted = isWanted;
            _callback = OnHook;
            _showCallback = OnShowHook;
            _timer = new Timer();
            _timer.Interval = TimerMilliseconds;
            _timer.Tick += OnTick;
            try
            {
                _hook = NativeMethods.SetWindowsHookEx(NativeMethods.WH_CBT, _callback, IntPtr.Zero,
                    NativeMethods.GetCurrentThreadId());
                if (_hook == IntPtr.Zero)
                {
                    Logger.Info("The new-window hook could not be installed. Error "
                        + Marshal.GetLastWin32Error() + ".");
                }

                _showHook = NativeMethods.SetWindowsHookEx(NativeMethods.WH_CALLWNDPROC, _showCallback, IntPtr.Zero,
                    NativeMethods.GetCurrentThreadId());
                if (_showHook == IntPtr.Zero)
                {
                    Logger.Info("The show-window hook could not be installed. Error "
                        + Marshal.GetLastWin32Error() + ".");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("NewWindowPlacer.Install", ex);
            }
        }

        // ----- The hook -----

        private IntPtr OnHook(int code, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if ((code == NativeMethods.HCBT_MINMAX || code == NativeMethods.HCBT_ACTIVATE)
                    && _hidden.ContainsKey(wParam))
                {
                    LogStep(wParam, "cbt " + code + " arg " + lParam.ToInt64());
                    if (code == NativeMethods.HCBT_MINMAX)
                    {
                        CloakHeldPowerPointOnce(wParam);
                    }
                }

                if (code == NativeMethods.HCBT_CREATEWND)
                {
                    Place(wParam, lParam);
                }
                else if (code == NativeMethods.HCBT_ACTIVATE && IsFrame(wParam))
                {
                    FrameAnimation.Disable(wParam);
                }
            }
            catch (Exception ex)
            {
                if (!_errorLogged)
                {
                    _errorLogged = true;
                    Logger.Error("NewWindowPlacer.OnHook", ex);
                }
            }

            return NativeMethods.CallNextHookEx(_hook, code, wParam, lParam);
        }

        // Runs just before a window is shown. Only windows kept unseen by Place are touched.
        private IntPtr OnShowHook(int code, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (code >= 0)
                {
                    NativeMethods.CWPSTRUCT message = (NativeMethods.CWPSTRUCT)Marshal.PtrToStructure(lParam,
                        typeof(NativeMethods.CWPSTRUCT));
                    if (_hidden.ContainsKey(message.Hwnd) && (message.Message == 0x18 || message.Message == 0x46
                        || message.Message == 0x47 || message.Message == 0x05 || message.Message == 0x06))
                    {
                        LogStep(message.Hwnd, "message " + message.Message.ToString("X") + " wParam "
                            + message.WParam.ToInt64());
                    }

                    if (message.Message == NativeMethods.WM_SHOWWINDOW && message.WParam != IntPtr.Zero
                        && _hidden.ContainsKey(message.Hwnd))
                    {
                        CloakHeldPowerPointOnce(message.Hwnd);
                    }
                }
            }
            catch (Exception ex)
            {
                if (!_errorLogged)
                {
                    _errorLogged = true;
                    Logger.Error("NewWindowPlacer.OnShowHook", ex);
                }
            }

            return NativeMethods.CallNextHookEx(_showHook, code, wParam, lParam);
        }

        // Word/Excel retain their established behavior. PowerPoint needs only one DWM cloak per hold.
        // WM_SHOWWINDOW and CBT_MINMAX can both report the same frame before the tab bar exists.
        private void CloakHeldPowerPointOnce(IntPtr window)
        {
            if (AppProfile.FrameClass != "PPTFrameClass")
            {
                FrameAnimation.Disable(window, true);
                FrameAnimation.Cloak(window, true);
                return;
            }

            if (_powerPointCloaked.Add(window))
            {
                FrameAnimation.Disable(window, true);
                FrameAnimation.Cloak(window, true);
            }
        }

        // Diagnostic: records what happens to a new window and in what order. Stops after StepLogLimit lines.
        private void LogStep(IntPtr window, string text)
        {
            if (_stepLogs >= StepLogLimit)
            {
                return;
            }

            _stepLogs++;
            Logger.Info("step " + (Environment.TickCount % 1000000) + " window " + window.ToInt64().ToString("X")
                + " " + text);
        }

        private static bool IsFrame(IntPtr window)
        {
            System.Text.StringBuilder text = new System.Text.StringBuilder(64);
            return NativeMethods.GetClassName(window, text, text.Capacity) > 0
                && text.ToString() == AppProfile.FrameClass;
        }

        // The creation hook may see an owned top-level frame. A nonzero Parent is not proof of WS_CHILD.
        // Word and Excel keep their established Parent check. PowerPoint uses the style bit instead.
        private void Place(IntPtr window, IntPtr lParam)
        {
            bool isPowerPoint = AppProfile.FrameClass == "PPTFrameClass";
            if (!_isWanted())
            {
                if (isPowerPoint && !_powerPointModeLogged)
                {
                    _powerPointModeLogged = true;
                    LogPowerPointSkip(window, "SingleWindowMode disabled; creation hold not used");
                }

                return;
            }

            NativeMethods.CBT_CREATEWND request = (NativeMethods.CBT_CREATEWND)Marshal.PtrToStructure(lParam,
                typeof(NativeMethods.CBT_CREATEWND));
            NativeMethods.CREATESTRUCT create = (NativeMethods.CREATESTRUCT)Marshal.PtrToStructure(
                request.CreateStruct, typeof(NativeMethods.CREATESTRUCT));
            if (!isPowerPoint && create.Parent != IntPtr.Zero)
            {
                return;
            }

            string className;
            if ((ulong)create.ClassName.ToInt64() <= 0xFFFF)
            {
                // An atom is not a string pointer. Only GetClassName on the supplied HWND can verify it.
                if (!isPowerPoint)
                {
                    if (!_atomLogged && create.ClassName != IntPtr.Zero)
                    {
                        _atomLogged = true;
                        Logger.Info("A top-level window was created with a class atom, so it was not checked.");
                    }

                    return;
                }

                System.Text.StringBuilder resolved = new System.Text.StringBuilder(64);
                if (NativeMethods.GetClassName(window, resolved, resolved.Capacity) <= 0)
                {
                    LogPowerPointSkip(window, "class atom could not be verified by GetClassName, parent="
                        + create.Parent.ToInt64().ToString("X") + " style=" + create.Style.ToString("X")
                        + " error=" + Marshal.GetLastWin32Error());
                    return;
                }

                className = resolved.ToString();
            }
            else
            {
                className = Marshal.PtrToStringUni(create.ClassName);
            }

            if (className != AppProfile.FrameClass)
            {
                if (isPowerPoint && _powerPointOtherClassLogs < PowerPointOtherClassLogLimit)
                {
                    _powerPointOtherClassLogs++;
                    LogPowerPointSkip(window, "non-frame class='" + className + "' parent="
                        + create.Parent.ToInt64().ToString("X") + " style=" + create.Style.ToString("X"));
                }

                return;
            }

            if (isPowerPoint && (create.Style & WS_CHILD) != 0)
            {
                LogPowerPointSkip(window, "PPTFrameClass rejected: WS_CHILD set, parent="
                    + create.Parent.ToInt64().ToString("X") + " style=" + create.Style.ToString("X"));
                return;
            }

            IntPtr reference = _referenceFrame();
            if (reference == IntPtr.Zero || !NativeMethods.IsWindow(reference) || NativeMethods.IsIconic(reference))
            {
                if (isPowerPoint)
                {
                    LogPowerPointSkip(window, "PPTFrameClass rejected: reference frame unavailable or minimized,"
                        + " reference=" + reference.ToInt64().ToString("X") + " parent="
                        + create.Parent.ToInt64().ToString("X") + " style=" + create.Style.ToString("X"));
                }

                return;
            }

            NativeMethods.RECT bounds = ReadBounds(reference);
            create.X = bounds.Left;
            create.Y = bounds.Top;
            create.Cx = bounds.Right - bounds.Left;
            create.Cy = bounds.Bottom - bounds.Top;
            create.ExStyle |= NativeMethods.WS_EX_LAYERED | NativeMethods.WS_EX_TRANSPARENT;
            Marshal.StructureToPtr(create, request.CreateStruct, false);
            _hidden[window] = Environment.TickCount;
            if (isPowerPoint)
            {
                HoldActiveLook(window, reference);
            }

            LogStep(window, "created style " + create.Style.ToString("X") + " exstyle " + create.ExStyle.ToString("X"));
            _timer.Start();
            if (_placementLogs < PlacementLogLimit)
            {
                _placementLogs++;
                Logger.Info("New main window " + window.ToInt64().ToString("X") + " placed at " + create.X + ","
                    + create.Y + " size " + create.Cx + "x" + create.Cy + " and kept unseen until it is ready.");
            }
        }

        private void LogPowerPointSkip(IntPtr window, string reason)
        {
            if (_powerPointSkipLogs >= PowerPointSkipLogLimit)
            {
                return;
            }

            _powerPointSkipLogs++;
            Logger.Info("PPT creation hook skipped window " + window.ToInt64().ToString("X") + ": " + reason + ".");
        }

        // A maximized window is copied as it is now. Any other window is copied at its restored size.
        private static NativeMethods.RECT ReadBounds(IntPtr window)
        {
            NativeMethods.RECT rect;
            if (!NativeMethods.IsZoomed(window))
            {
                NativeMethods.WINDOWPLACEMENT placement = new NativeMethods.WINDOWPLACEMENT();
                placement.Length = (uint)Marshal.SizeOf(typeof(NativeMethods.WINDOWPLACEMENT));
                if (NativeMethods.GetWindowPlacement(window, ref placement))
                {
                    return placement.NormalPosition;
                }
            }

            NativeMethods.GetWindowRect(window, out rect);
            return rect;
        }

        // Raised when a held window has been shown, or has gone away without being shown.
        public event Action WindowShown;

        // True while a window is kept unseen and has not been revealed yet.
        public bool IsHeld(IntPtr window)
        {
            return window != IntPtr.Zero && _hidden.ContainsKey(window);
        }

        private void RaiseWindowShown()
        {
            Action handler = WindowShown;
            if (handler != null)
            {
                handler();
            }
        }

        // ----- Revealing -----

        // PowerPoint can finish the tab bar before its slide window has settled. Keep that frame
        // cloaked until the visible editing pane and slide have stable bounds across checks.
        // Word and Excel still reveal immediately through their original two-stage path.
        public void Reveal(IntPtr window)
        {
            if (window == IntPtr.Zero || !_hidden.ContainsKey(window))
            {
                return;
            }

            if (AppProfile.FrameClass == "PPTFrameClass")
            {
                if (_powerPointRevealPending.Add(window))
                {
                    string layout;
                    string reason;
                    if (TryPowerPointLayout(window, out layout, out reason))
                    {
                        _powerPointReadyLayouts[window] = layout;
                    }
                    else
                    {
                        LogPowerPointReadiness(window, "held: " + reason);
                    }
                }

                _timer.Start();
                return;
            }

            FinishReveal(window);
        }

        // Completes the original reveal operation, after PowerPoint readiness or the safety timeout.
        private void FinishReveal(IntPtr window, bool disposing = false)
        {
            if (window == IntPtr.Zero || !_hidden.ContainsKey(window))
            {
                return;
            }

            _hidden.Remove(window);
            _powerPointCloaked.Remove(window);
            _powerPointRevealPending.Remove(window);
            _powerPointReadyLayouts.Remove(window);
            if (!NativeMethods.IsWindow(window))
            {
                ReleaseActiveLook(window);
                _inspectAfterReveal.Remove(window);
                return;
            }

            bool isPowerPoint = AppProfile.FrameClass == "PPTFrameClass";
            if (!isPowerPoint)
            {
                // Preserve the established Word/Excel immediate redraw behavior.
                NativeMethods.RedrawWindow(window, IntPtr.Zero, IntPtr.Zero, NativeMethods.RDW_INVALIDATE
                    | NativeMethods.RDW_ALLCHILDREN | NativeMethods.RDW_UPDATENOW);
            }

            FrameAnimation.Disable(window, true);
            if (isPowerPoint)
            {
                // Do not hide or move PowerPoint's independent border-effect windows.
                bool inspect = !disposing && _powerPointVisibilityLogs < PowerPointVisibilityLogLimit;
                if (inspect)
                {
                    _powerPointVisibilityLogs++;
                    LogPowerPointVisibility(window, "before uncloak");
                }

                RemoveLayered(window);
                ReleaseActiveLook(window);
                if (inspect)
                {
                    LogPowerPointVisibility(window, "after uncloak");
                    _inspectAfterReveal[window] = Environment.TickCount;
                    _timer.Start();
                }

                RaiseWindowShown();
                if (_hidden.Count == 0 && _revealed.Count == 0 && _inspectAfterReveal.Count == 0)
                {
                    _timer.Stop();
                }
            }
            else
            {
                // Retain the established Word/Excel two-stage reveal behavior.
                NativeMethods.SetLayeredWindowAttributes(window, 0, 255, NativeMethods.LWA_ALPHA);
                _revealed[window] = Environment.TickCount;
            }
        }

        // IsWindowVisible checks the WS_VISIBLE chain, even for DWM-cloaked frames. It is
        // suitable for testing PowerPoint's child layout, not a guarantee that pixels were painted.
        private static bool TryPowerPointLayout(IntPtr frame, out string layout, out string reason)
        {
            layout = null;
            reason = "editing pane not found";
            if (!NativeMethods.IsWindowVisible(frame) || NativeMethods.IsIconic(frame))
            {
                reason = "frame not shown";
                return false;
            }

            NativeMethods.RECT frameRect;
            if (!NativeMethods.GetWindowRect(frame, out frameRect))
            {
                reason = "frame bounds unavailable";
                return false;
            }

            IntPtr pane = IntPtr.Zero;
            for (IntPtr child = NativeMethods.GetWindow(frame, NativeMethods.GW_CHILD); child != IntPtr.Zero;
                child = NativeMethods.GetWindow(child, NativeMethods.GW_HWNDNEXT))
            {
                string name = ClassOf(child);
                if (name == "FullpageUIHost" && NativeMethods.IsWindowVisible(child))
                {
                    reason = "PowerPoint transition overlay visible";
                    return false;
                }

                if (name == "MDIClient" && NativeMethods.IsWindowVisible(child))
                {
                    pane = child;
                }
            }

            if (pane == IntPtr.Zero)
            {
                reason = "editing pane not visible";
                return false;
            }

            NativeMethods.RECT paneRect;
            if (!NativeMethods.GetWindowRect(pane, out paneRect) || !HasArea(paneRect))
            {
                reason = "editing pane has no usable bounds";
                return false;
            }

            for (IntPtr child = NativeMethods.GetWindow(pane, NativeMethods.GW_CHILD); child != IntPtr.Zero;
                child = NativeMethods.GetWindow(child, NativeMethods.GW_HWNDNEXT))
            {
                if (ClassOf(child) != "mdiClass" || !NativeMethods.IsWindowVisible(child))
                {
                    continue;
                }

                NativeMethods.RECT slideRect;
                if (!NativeMethods.GetWindowRect(child, out slideRect) || !HasArea(slideRect))
                {
                    continue;
                }

                layout = RectKey(frameRect) + ":" + RectKey(paneRect) + ":" + RectKey(slideRect);
                reason = "editing pane and slide visible";
                return true;
            }

            reason = "slide child not visible";
            return false;
        }

        private static bool HasArea(NativeMethods.RECT area)
        {
            return area.Right > area.Left && area.Bottom > area.Top;
        }

        private static string RectKey(NativeMethods.RECT area)
        {
            return area.Left + "," + area.Top + "," + area.Right + "," + area.Bottom;
        }

        private static string ClassOf(IntPtr window)
        {
            System.Text.StringBuilder name = new System.Text.StringBuilder(64);
            return NativeMethods.GetClassName(window, name, name.Capacity) > 0 ? name.ToString() : string.Empty;
        }

        private void LogPowerPointReadiness(IntPtr window, string text)
        {
            if (_powerPointReadinessLogs >= PowerPointReadinessLogLimit)
            {
                return;
            }

            _powerPointReadinessLogs++;
            Logger.Info("PPT readiness frame=" + window.ToInt64().ToString("X") + " " + text + ".");
        }

        private void OnTick(object sender, EventArgs e)
        {
            try
            {
                int now = Environment.TickCount;
                bool changed = false;
                foreach (KeyValuePair<IntPtr, int> pair in new List<KeyValuePair<IntPtr, int>>(_inspectAfterReveal))
                {
                    if (!NativeMethods.IsWindow(pair.Key))
                    {
                        _inspectAfterReveal.Remove(pair.Key);
                    }
                    else if (unchecked(now - pair.Value) >= TimerMilliseconds)
                    {
                        LogPowerPointVisibility(pair.Key, "next timer tick");
                        _inspectAfterReveal.Remove(pair.Key);
                    }
                }

                foreach (IntPtr frame in new List<IntPtr>(_powerPointRevealPending))
                {
                    if (!_hidden.ContainsKey(frame))
                    {
                        _powerPointRevealPending.Remove(frame);
                        _powerPointReadyLayouts.Remove(frame);
                        continue;
                    }

                    string layout;
                    string reason;
                    if (!TryPowerPointLayout(frame, out layout, out reason))
                    {
                        _powerPointReadyLayouts.Remove(frame);
                        continue;
                    }

                    string previous;
                    if (_powerPointReadyLayouts.TryGetValue(frame, out previous) && previous == layout)
                    {
                        LogPowerPointReadiness(frame, "stable editing pane and slide; revealing");
                        FinishReveal(frame);
                    }
                    else
                    {
                        _powerPointReadyLayouts[frame] = layout;
                    }
                }

                foreach (KeyValuePair<IntPtr, int> pair in new List<KeyValuePair<IntPtr, int>>(_hidden))
                {
                    if (!NativeMethods.IsWindow(pair.Key))
                    {
                        _hidden.Remove(pair.Key);
                        _powerPointCloaked.Remove(pair.Key);
                        _powerPointRevealPending.Remove(pair.Key);
                        _powerPointReadyLayouts.Remove(pair.Key);
                        ReleaseActiveLook(pair.Key);
                        changed = true;
                    }
                    else if (now - pair.Value > MaxHiddenMilliseconds)
                    {
                        Logger.Info("Window " + pair.Key.ToInt64().ToString("X") + " was revealed by the timer.");
                        if (_powerPointRevealPending.Contains(pair.Key))
                        {
                            LogPowerPointReadiness(pair.Key, "safety timeout; releasing without readiness");
                        }

                        FinishReveal(pair.Key);
                    }
                }

                foreach (KeyValuePair<IntPtr, int> pair in new List<KeyValuePair<IntPtr, int>>(_revealed))
                {
                    if (!NativeMethods.IsWindow(pair.Key))
                    {
                        _revealed.Remove(pair.Key);
                        changed = true;
                    }
                    else if (now - pair.Value > LayeredStyleMilliseconds)
                    {
                        RemoveLayered(pair.Key);
                        _revealed.Remove(pair.Key);
                        changed = true;
                    }
                }

                if (changed)
                {
                    RaiseWindowShown();
                }

                if (_hidden.Count == 0 && _revealed.Count == 0 && _inspectAfterReveal.Count == 0)
                {
                    _timer.Stop();
                }
            }
            catch (Exception ex)
            {
                Logger.Error("NewWindowPlacer.OnTick", ex);
            }
        }

        // Diagnostic only: Office draws its border-effect windows separately from PPTFrameClass.
        // Compare the DWM application-cloak bit with nearby visible border windows at reveal time.
        private static void LogPowerPointVisibility(IntPtr frame, string phase)
        {
            int cloakFlags = -1;
            int result = NativeMethods.DwmGetWindowAttribute(frame, DWMWA_CLOAKED, ref cloakFlags,
                sizeof(int));
            if (result != 0)
            {
                cloakFlags = -1;
            }

            NativeMethods.RECT bounds;
            bool hasBounds = NativeMethods.GetWindowRect(frame, out bounds);
            uint processId;
            uint frameThread = NativeMethods.GetWindowThreadProcessId(frame, out processId);
            int nearbyBorders = 0;
            int visibleBorders = 0;
            int seen = 0;
            for (IntPtr other = NativeMethods.GetWindow(frame, GW_HWNDFIRST); other != IntPtr.Zero && seen < 512;
                other = NativeMethods.GetWindow(other, NativeMethods.GW_HWNDNEXT))
            {
                seen++;
                uint otherProcess;
                if (other == frame || NativeMethods.GetWindowThreadProcessId(other, out otherProcess) != frameThread)
                {
                    continue;
                }

                System.Text.StringBuilder name = new System.Text.StringBuilder(64);
                if (NativeMethods.GetClassName(other, name, name.Capacity) <= 0
                    || name.ToString() != "MSO_BORDEREFFECT_WINDOW_CLASS")
                {
                    continue;
                }

                NativeMethods.RECT area;
                if (!hasBounds || !NativeMethods.GetWindowRect(other, out area)
                    || area.Left > bounds.Right + 16 || area.Right < bounds.Left - 16
                    || area.Top > bounds.Bottom + 16 || area.Bottom < bounds.Top - 16)
                {
                    continue;
                }

                nearbyBorders++;
                if (NativeMethods.IsWindowVisible(other))
                {
                    visibleBorders++;
                }
            }

            Logger.Info("PPT visibility " + phase + " frame=" + frame.ToInt64().ToString("X")
                + " visible=" + NativeMethods.IsWindowVisible(frame)
                + " appCloaked=" + (cloakFlags < 0 ? "unknown" : ((cloakFlags & DWM_CLOAKED_APP) != 0).ToString())
                + " cloakFlags=" + cloakFlags + " dwmResult=" + result
                + " layered=" + ((NativeMethods.GetExStyle(frame) & NativeMethods.WS_EX_LAYERED) != 0)
                + " nearBorderWindows=" + nearbyBorders + " visibleBorders=" + visibleBorders
                + " tick=" + (Environment.TickCount % 1000000) + ".");
        }

        // ----- PowerPoint active look during a hold -----

        // Keeps the frame the user is looking at drawn as active while the new frame is held. If that frame is
        // already held by an earlier creation, nothing more is added.
        private void HoldActiveLook(IntPtr held, IntPtr visible)
        {
            foreach (ActiveLook existing in _activeLooks.Values)
            {
                if (existing.Handle == visible)
                {
                    return;
                }
            }

            ActiveLook look = new ActiveLook();
            look.AssignHandle(visible);
            _activeLooks[held] = look;
        }

        private void ReleaseActiveLook(IntPtr held)
        {
            ActiveLook look;
            if (!_activeLooks.TryGetValue(held, out look))
            {
                return;
            }

            _activeLooks.Remove(held);
            if (look.Handle != IntPtr.Zero)
            {
                look.ReleaseHandle();
            }
        }

        private sealed class ActiveLook : NativeWindow
        {
            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_NCACTIVATE && m.WParam == IntPtr.Zero)
                {
                    m.WParam = new IntPtr(1);
                }

                base.WndProc(ref m);
            }
        }

        private static void RemoveLayered(IntPtr window)
        {
            int style = NativeMethods.GetExStyle(window);
            if ((style & NativeMethods.WS_EX_LAYERED) != 0)
            {
                NativeMethods.SetExStyle(window, style & ~(NativeMethods.WS_EX_LAYERED
                    | NativeMethods.WS_EX_TRANSPARENT));
                if (AppProfile.FrameClass == "PPTFrameClass")
                {
                    // Microsoft layered-window example: invalidate once after removing WS_EX_LAYERED.
                    // No RDW_UPDATENOW: Windows can paint when the newly uncloaked frame is ready.
                    NativeMethods.RedrawWindow(window, IntPtr.Zero, IntPtr.Zero, NativeMethods.RDW_INVALIDATE
                        | NativeMethods.RDW_ERASE | RDW_FRAME | NativeMethods.RDW_ALLCHILDREN);
                }
                else
                {
                    // Preserve the original Word/Excel cleanup redraw exactly.
                    NativeMethods.RedrawWindow(window, IntPtr.Zero, IntPtr.Zero, NativeMethods.RDW_INVALIDATE
                        | NativeMethods.RDW_ERASE | NativeMethods.RDW_ALLCHILDREN | NativeMethods.RDW_UPDATENOW);
                }
            }

            FrameAnimation.Cloak(window, false);
        }

        // ----- Cleanup -----

        public void Dispose()
        {
            try
            {
                _timer.Stop();
                _timer.Dispose();
                foreach (IntPtr window in new List<IntPtr>(_hidden.Keys))
                {
                    FinishReveal(window, true);
                }

                foreach (IntPtr window in new List<IntPtr>(_revealed.Keys))
                {
                    if (NativeMethods.IsWindow(window))
                    {
                        RemoveLayered(window);
                    }
                }

                foreach (IntPtr window in new List<IntPtr>(_activeLooks.Keys))
                {
                    ReleaseActiveLook(window);
                }

                _revealed.Clear();
                _powerPointCloaked.Clear();
                _powerPointRevealPending.Clear();
                _powerPointReadyLayouts.Clear();
                _inspectAfterReveal.Clear();
                RaiseWindowShown();
                if (_hook != IntPtr.Zero)
                {
                    NativeMethods.UnhookWindowsHookEx(_hook);
                    _hook = IntPtr.Zero;
                }

                if (_showHook != IntPtr.Zero)
                {
                    NativeMethods.UnhookWindowsHookEx(_showHook);
                    _showHook = IntPtr.Zero;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("NewWindowPlacer.Dispose", ex);
            }
        }
    }
}
