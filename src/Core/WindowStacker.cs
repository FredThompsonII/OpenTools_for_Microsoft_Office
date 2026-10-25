// ---------------------------------------------------------------------------------------------------------------
// Name:     WindowStacker.cs
// Purpose:  Single-window mode. Gives every document window of an app the same size and position, and brings
//           one to the front on request. Office windows stay where they are in the window tree. Nothing is
//           pulled into a container. A re-entry guard stops two windows from answering each other's resize.
// Revision: 2
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace OpenTools
{
    internal sealed class WindowStacker
    {
        private bool _busy;

        // ----- Matching size and position -----

        public void ApplyBoundsFrom(IntPtr source, IEnumerable<IntPtr> targets)
        {
            if (_busy || source == IntPtr.Zero || !NativeMethods.IsWindow(source) || NativeMethods.IsIconic(source))
            {
                return;
            }

            _busy = true;
            try
            {
                bool maximized = NativeMethods.IsZoomed(source);
                NativeMethods.RECT bounds = ReadNormalBounds(source);
                foreach (IntPtr target in targets)
                {
                    if (target == source || !NativeMethods.IsWindow(target) || NativeMethods.IsIconic(target))
                    {
                        continue;
                    }

                    MatchWindow(target, maximized, bounds);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("WindowStacker.ApplyBoundsFrom", ex);
            }
            finally
            {
                _busy = false;
            }
        }

        private static NativeMethods.RECT ReadNormalBounds(IntPtr window)
        {
            NativeMethods.WINDOWPLACEMENT placement = new NativeMethods.WINDOWPLACEMENT();
            placement.Length = (uint)Marshal.SizeOf(typeof(NativeMethods.WINDOWPLACEMENT));
            if (NativeMethods.GetWindowPlacement(window, ref placement))
            {
                return placement.NormalPosition;
            }

            NativeMethods.RECT rect;
            NativeMethods.GetWindowRect(window, out rect);
            return rect;
        }

        private static void MatchWindow(IntPtr target, bool maximized, NativeMethods.RECT bounds)
        {
            bool targetMaximized = NativeMethods.IsZoomed(target);
            if (maximized)
            {
                if (!targetMaximized)
                {
                    NativeMethods.ShowWindow(target, NativeMethods.SW_SHOWMAXIMIZED);
                }

                return;
            }

            if (targetMaximized)
            {
                NativeMethods.ShowWindow(target, NativeMethods.SW_RESTORE);
            }

            NativeMethods.SetWindowPos(target, IntPtr.Zero, bounds.Left, bounds.Top,
                bounds.Right - bounds.Left, bounds.Bottom - bounds.Top,
                NativeMethods.SWP_NOZORDER | NativeMethods.SWP_NOACTIVATE);
        }

        // ----- Bringing a window to the front -----

        public void BringToFront(IntPtr window)
        {
            try
            {
                if (window == IntPtr.Zero || !NativeMethods.IsWindow(window))
                {
                    return;
                }

                if (NativeMethods.IsIconic(window))
                {
                    NativeMethods.ShowWindow(window, NativeMethods.SW_RESTORE);
                }

                NativeMethods.SetForegroundWindow(window);
            }
            catch (Exception ex)
            {
                Logger.Error("WindowStacker.BringToFront", ex);
            }
        }
    }
}
