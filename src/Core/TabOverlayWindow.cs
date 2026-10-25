// ---------------------------------------------------------------------------------------------------------------
// Name:     TabOverlayWindow.cs
// Purpose:  A small borderless window that floats over Office's own "Tabs" task pane strip and holds the real tab
//           bar. It is owned by the Office main window, so it stays above it and hides with it. It never takes
//           focus, is not listed on the taskbar, and is moved and shown by TaskPaneHost. It does no Office calls.
// Revision: 2
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Drawing;
using System.Windows.Forms;

namespace OpenTools
{
    internal sealed class TabOverlayWindow : Form
    {
        private bool _shown;
        private bool _visibleNow;
        private Rectangle _bounds = Rectangle.Empty;

        public TabOverlayWindow(TabBarControl bar)
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            AutoScaleMode = AutoScaleMode.None;
            Padding = Padding.Empty;
            ClientSize = new Size(600, 30);

            bar.Dock = DockStyle.Fill;
            Controls.Add(bar);
        }

        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= NativeMethods.WS_EX_NOACTIVATE | NativeMethods.WS_EX_TOOLWINDOW;
                return cp;
            }
        }

        // ----- Showing and placing -----

        // Moves and shows the window. The first call also gives it its owner, which is the Office main window.
        public void ShowAt(Rectangle bounds, IntPtr ownerHandle)
        {
            if (IsDisposed)
            {
                return;
            }

            if (!_shown)
            {
                Bounds = bounds;
                _bounds = bounds;
                Show(new OwnerWindow(ownerHandle));
                _shown = true;
                _visibleNow = true;
                return;
            }

            if (bounds != _bounds)
            {
                NativeMethods.SetWindowPos(Handle, IntPtr.Zero, bounds.X, bounds.Y, bounds.Width, bounds.Height,
                    NativeMethods.SWP_NOZORDER | NativeMethods.SWP_NOACTIVATE);
                _bounds = bounds;
            }

            if (!_visibleNow)
            {
                NativeMethods.ShowWindow(Handle, NativeMethods.SW_SHOWNOACTIVATE);
                _visibleNow = true;
            }
        }

        public void HideWindow()
        {
            if (IsDisposed || !_shown || !_visibleNow)
            {
                return;
            }

            NativeMethods.ShowWindow(Handle, NativeMethods.SW_HIDE);
            _visibleNow = false;
        }

        // ----- Owner wrapper -----

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
    }
}
