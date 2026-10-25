// ---------------------------------------------------------------------------------------------------------------
// Name:     FrameAnimation.cs
// Purpose:  Turns off the Windows open, close and minimize animations (fade and zoom) on Office main windows, so
//           that opening or closing a document does not look like a flash. This is a per-window Windows setting.
//           It lasts only as long as the window does. The caller can ask for the setting to be applied again
//           (force), for example when a window is created and again when it is revealed. RestoreAll puts the
//           animations back on every window that was changed, when the add-in stops. A failure is logged once
//           and then ignored.
// Revision: 4
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;

namespace OpenTools
{
    internal static class FrameAnimation
    {
        private static readonly HashSet<IntPtr> _changed = new HashSet<IntPtr>();
        private static bool _errorLogged;

        // Safe to call more than once for the same window. A window that was already changed is skipped.
        public static void Disable(IntPtr window)
        {
            Disable(window, false);
        }

        // With force set, the setting is applied again even if the window was changed before. The log line
        // is written only the first time a window is changed.
        public static void Disable(IntPtr window, bool force)
        {
            if (window == IntPtr.Zero || (!force && _changed.Contains(window)) || !NativeMethods.IsWindow(window))
            {
                return;
            }

            if (Set(window, 1) && _changed.Add(window))
            {
                Logger.Info("Window animations turned off for window " + window.ToInt64().ToString("X") + ".");
            }
        }

        // Hides a window from the Windows window manager, or shows it again. A cloaked window is not drawn and
        // is not animated. The window must already be known to Windows, so call this when it is about to be shown.
        public static void Cloak(IntPtr window, bool cloaked)
        {
            if (window == IntPtr.Zero || !NativeMethods.IsWindow(window))
            {
                return;
            }

            int value = cloaked ? 1 : 0;
            int result = NativeMethods.DwmSetWindowAttribute(window, NativeMethods.DWMWA_CLOAK, ref value,
                sizeof(int));
            Logger.Info("Window " + window.ToInt64().ToString("X") + (cloaked ? " cloaked" : " uncloaked")
                + ". Result " + result + ".");
        }

        public static void RestoreAll()
        {
            foreach (IntPtr window in new List<IntPtr>(_changed))
            {
                if (NativeMethods.IsWindow(window))
                {
                    Set(window, 0);
                }
            }

            _changed.Clear();
        }

        private static bool Set(IntPtr window, int value)
        {
            try
            {
                int result = NativeMethods.DwmSetWindowAttribute(window,
                    NativeMethods.DWMWA_TRANSITIONS_FORCEDISABLED, ref value, sizeof(int));
                if (result != 0 && !_errorLogged)
                {
                    _errorLogged = true;
                    Logger.Info("Windows refused to change the window animations. Result " + result + ".");
                }

                return result == 0;
            }
            catch (Exception ex)
            {
                if (!_errorLogged)
                {
                    _errorLogged = true;
                    Logger.Error("FrameAnimation.Set", ex);
                }

                return false;
            }
        }
    }
}
