// ---------------------------------------------------------------------------------------------------------------
// Name:     TabBarControl.cs
// Purpose:  The tab bar itself: owner-drawn WinForms control shown inside an Office main window or a floating
//           window. Draws the tabs with clipped (cut) top corners, handles click, double-click, middle-click and
//           drag, and asks its host to activate or close a document, to start a new one, or to close all
//           documents (the X button at the far right). It never talks to Office directly. Drag
//           reordering changes the shared TabList. UseHostParent makes the control a child window of a window
//           it does not own, such as the Office main window. HoldPaint stops the bar from drawing list changes
//           until the hold ends, so a new tab is not seen in a window that is about to be covered. From the click
//           on the new-document button until the bar is held and released (or two seconds, if nothing is held),
//           the bar shows the busy cursor, so the pointer does not switch between Office's busy cursor and the
//           bar's arrow.
// Revision: 9
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OpenTools
{
    internal sealed class TabBarControl : Control
    {
        private const int TabGap = 2;
        private const int WaitCursorMilliseconds = 2000;
        private const int IconSize = 16;
        private const int DotSize = 6;
        private const int PadSide = 8;
        private const int PadVertical = 4;
        private const int AccentThickness = 2;
        private const int CornerClip = 6;

        private readonly bool _allowClose;
        private readonly TabList _list;
        private readonly Settings _settings;
        private readonly TabBarLayout _layout = new TabBarLayout();
        private readonly Dictionary<string, Bitmap> _icons = new Dictionary<string, Bitmap>();
        private readonly List<int> _widths = new List<int>();
        private readonly List<string> _labels = new List<string>();
        private readonly ToolTip _toolTip = new ToolTip();

        private Font _font;
        private Font _boldFont;
        private ContextMenuStrip _overflowMenu;
        private IntPtr _hostParent = IntPtr.Zero;
        private int _hoverIndex = -1;
        private bool _hoverNew;
        private bool _hoverCloseAll;
        private int _dragIndex = -1;
        private Point _dragStart = Point.Empty;
        private bool _dragging;
        private bool _inRelayout;
        private bool _held;
        private readonly Timer _waitTimer = new Timer();
        private bool _heldChange;
        private bool _paintErrorLogged;

        // An app can hide all closing actions until its adapter supports safe cancellation.
        public TabBarControl(TabList list, Settings settings, bool allowClose)
        {
            _list = list;
            _settings = settings;
            _allowClose = allowClose;

            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint
                | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            TabStop = false;
            BackColor = SystemColors.ControlLight;

            _list.Changed += OnListChanged;
            _waitTimer.Interval = WaitCursorMilliseconds;
            _waitTimer.Tick += OnWaitTick;
            ApplySettings();
        }

        public event Action<string> ActivateRequested;

        public event Action<string> CloseRequested;

        public event Action NewDocumentRequested;

        public event Action CloseAllRequested;

        public event Action<int> HeightWanted;

        public int WantedHeight
        {
            get { return _layout.Height; }
        }

        // Call before the window handle exists. The window is then created as a child of the given window.
        public void UseHostParent(IntPtr parent)
        {
            _hostParent = parent;
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                if (_hostParent != IntPtr.Zero)
                {
                    cp.Parent = _hostParent;
                    cp.Style |= NativeMethods.WS_CLIPSIBLINGS;
                }

                return cp;
            }
        }

        // ----- Settings and layout -----

        public void ApplySettings()
        {
            Font oldFont = _font;
            Font oldBold = _boldFont;
            BuildFonts();
            if (oldFont != null)
            {
                oldFont.Dispose();
            }

            if (oldBold != null)
            {
                oldBold.Dispose();
            }

            Relayout();
        }

        private void BuildFonts()
        {
            Font baseFont = SystemFonts.MessageBoxFont;
            string family = baseFont.FontFamily.Name;
            float size = baseFont.Size;

            if (_settings.GetBool(SettingNames.UseCustomFont))
            {
                string name = _settings.GetString(SettingNames.CustomFontName);
                if (!string.IsNullOrEmpty(name))
                {
                    family = name;
                }
            }

            if (_settings.GetBool(SettingNames.UseCustomFontSize))
            {
                size = _settings.GetInt(SettingNames.CustomFontSize);
            }

            try
            {
                _font = new Font(family, size, FontStyle.Regular, GraphicsUnit.Point);
                _boldFont = new Font(family, size, FontStyle.Bold, GraphicsUnit.Point);
            }
            catch (ArgumentException ex)
            {
                Logger.Error("TabBarControl.BuildFonts", ex);
                _font = new Font(baseFont, FontStyle.Regular);
                _boldFont = new Font(baseFont, FontStyle.Bold);
            }
        }

        private int Scale(int value)
        {
            return (int)Math.Round(value * DeviceDpi / 96.0);
        }

        private int RowHeight()
        {
            return Math.Max(_font.Height, Scale(IconSize)) + Scale(PadVertical * 2);
        }

        // A change to the list is drawn at once. This matters while a close-all is running, because the message
        // loop is busy then, and a repaint that only waited its turn would show every tab until the end.
        private void OnWaitTick(object sender, EventArgs e)
        {
            _waitTimer.Stop();
            if (!_held)
            {
                UseWaitCursor = false;
            }
        }

        public void HoldPaint(bool hold)
        {
            if (hold == _held)
            {
                return;
            }

            _held = hold;
            UseWaitCursor = hold;
            if (!hold)
            {
                _waitTimer.Stop();
            }

            if (!hold && _heldChange && !IsDisposed)
            {
                _heldChange = false;
                OnListChanged(this, EventArgs.Empty);
            }
        }

        private void OnListChanged(object sender, EventArgs e)
        {
            if (_held)
            {
                _heldChange = true;
                return;
            }

            Relayout();
            if (IsHandleCreated && Visible)
            {
                Update();
            }
        }

        private void Relayout()
        {
            if (_inRelayout || _font == null)
            {
                return;
            }

            _inRelayout = true;
            try
            {
                MeasureTabs();
                int rowHeight = RowHeight();
                _layout.Compute(ClientSize.Width, rowHeight, Scale(TabGap), _widths,
                    _list.IndexOf(_list.CurrentKey), _settings.GetBool(SettingNames.MultiLine),
                    _settings.GetBool(SettingNames.NewDocumentButton), rowHeight, CloseAllWidth(rowHeight));
                Invalidate();

                Action<int> handler = HeightWanted;
                if (handler != null && _layout.Height != ClientSize.Height)
                {
                    handler(_layout.Height);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("TabBarControl.Relayout", ex);
            }
            finally
            {
                _inRelayout = false;
            }
        }

        // The close-all button is as wide as a row is tall. It is shown when there are tabs and the setting is on.
        private int CloseAllWidth(int rowHeight)
        {
            return _allowClose && _list.Count > 0 && _settings.GetBool(SettingNames.CloseAllButton)
                ? rowHeight : 0;
        }

        private void MeasureTabs()
        {
            _widths.Clear();
            _labels.Clear();
            int maxChars = _settings.GetInt(SettingNames.MaxTabChars);
            bool auto = _settings.GetBool(SettingNames.AutoWidth);
            bool icons = _settings.GetBool(SettingNames.FileIcons);
            int fixedWidth = Scale(_settings.GetInt(SettingNames.FixedTabWidth));
            int extra = Scale(PadSide * 2) + (icons ? Scale(IconSize + 4) : 0) + Scale(DotSize + 6);

            foreach (TabItem item in _list.Items)
            {
                string label = TabLabeler.Display(item, maxChars);
                _labels.Add(label);
                if (auto)
                {
                    int textWidth = TextRenderer.MeasureText(label, _boldFont, new Size(10000, 1000),
                        TextFormatFlags.SingleLine | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix).Width;
                    _widths.Add(textWidth + extra);
                }
                else
                {
                    _widths.Add(fixedWidth);
                }
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            Relayout();
        }

        // ----- Painting -----

        protected override void OnPaint(PaintEventArgs e)
        {
            try
            {
                Graphics g = e.Graphics;
                g.Clear(BackColor);
                IList<TabItem> items = _list.Items;
                for (int i = 0; i < items.Count && i < _layout.TabRects.Count; i++)
                {
                    Rectangle rect = _layout.TabRects[i];
                    if (!rect.IsEmpty)
                    {
                        DrawTab(g, items[i], _labels[i], rect, i == _hoverIndex);
                    }
                }

                if (!_layout.NewButtonRect.IsEmpty)
                {
                    DrawNewButton(g, _layout.NewButtonRect);
                }

                if (!_settings.GetBool(SettingNames.MultiLine) && !_layout.OverflowRect.IsEmpty)
                {
                    DrawOverflow(g, _layout.OverflowRect);
                }

                if (!_layout.CloseAllRect.IsEmpty)
                {
                    DrawCloseAll(g, _layout.CloseAllRect);
                }

                DrawEdgeLine(g);
            }
            catch (Exception ex)
            {
                if (!_paintErrorLogged)
                {
                    _paintErrorLogged = true;
                    Logger.Error("TabBarControl.OnPaint", ex);
                }
            }
        }

        // The outline of a tab with its two outer corners cut off diagonally. The cut corners are at the top
        // when the bar is above the document and at the bottom when it is below.
        private GraphicsPath TabPath(Rectangle rect)
        {
            bool below = _settings.GetPosition() == TabBarPosition.Below;
            int clip = Math.Min(Scale(CornerClip), Math.Min(rect.Width, rect.Height) / 2);
            int left = rect.X;
            int right = rect.Right - 1;
            int top = rect.Y;
            int bottom = rect.Bottom - 1;
            Point[] points;
            if (below)
            {
                points = new Point[]
                {
                    new Point(left, top), new Point(right, top), new Point(right, bottom - clip),
                    new Point(right - clip, bottom), new Point(left + clip, bottom), new Point(left, bottom - clip)
                };
            }
            else
            {
                points = new Point[]
                {
                    new Point(left, bottom), new Point(left, top + clip), new Point(left + clip, top),
                    new Point(right - clip, top), new Point(right, top + clip), new Point(right, bottom)
                };
            }

            GraphicsPath path = new GraphicsPath();
            path.AddPolygon(points);
            return path;
        }

        private void DrawTab(Graphics g, TabItem item, string label, Rectangle rect, bool hover)
        {
            bool current = item.Key == _list.CurrentKey;
            Color back = hover ? SystemColors.ControlLightLight : SystemColors.Control;
            if (current)
            {
                back = SystemColors.Window;
            }

            using (GraphicsPath path = TabPath(rect))
            {
                Region oldClip = g.Clip;
                g.SetClip(path, CombineMode.Replace);
                using (SolidBrush brush = new SolidBrush(back))
                {
                    g.FillRectangle(brush, rect);
                }

                if (current)
                {
                    DrawAccent(g, rect);
                }

                g.Clip = oldClip;
                oldClip.Dispose();

                SmoothingMode oldSmoothing = g.SmoothingMode;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (Pen pen = new Pen(SystemColors.ControlDark))
                {
                    g.DrawPath(pen, path);
                }

                g.SmoothingMode = oldSmoothing;
            }

            int x = rect.X + Scale(PadSide);
            if (_settings.GetBool(SettingNames.FileIcons))
            {
                Bitmap icon = GetIcon(item);
                int size = Scale(IconSize);
                if (icon != null)
                {
                    g.DrawImage(icon, x, rect.Y + (rect.Height - size) / 2, size, size);
                }

                x += size + Scale(4);
            }

            int dotSpace = item.IsDirty ? Scale(DotSize + 4) : 0;
            Rectangle textRect = new Rectangle(x, rect.Y, rect.Right - x - Scale(PadSide) - dotSpace, rect.Height);
            Font font = current && _settings.GetBool(SettingNames.BoldCurrentTab) ? _boldFont : _font;
            TextRenderer.DrawText(g, label, font, textRect, SystemColors.WindowText,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis
                | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine | TextFormatFlags.NoPadding);

            if (item.IsDirty)
            {
                DrawDirtyDot(g, rect);
            }
        }

        private void DrawAccent(Graphics g, Rectangle rect)
        {
            int thickness = Scale(AccentThickness);
            int y = _settings.GetPosition() == TabBarPosition.Below ? rect.Bottom - thickness : rect.Y;
            using (SolidBrush brush = new SolidBrush(SystemColors.Highlight))
            {
                g.FillRectangle(brush, rect.X, y, rect.Width, thickness);
            }
        }

        private void DrawDirtyDot(Graphics g, Rectangle rect)
        {
            int size = Scale(DotSize);
            int x = rect.Right - Scale(PadSide) - size;
            int y = rect.Y + (rect.Height - size) / 2;
            SmoothingMode old = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (SolidBrush brush = new SolidBrush(SystemColors.Highlight))
            {
                g.FillEllipse(brush, x, y, size, size);
            }

            g.SmoothingMode = old;
        }

        private void DrawNewButton(Graphics g, Rectangle rect)
        {
            if (_hoverNew)
            {
                using (SolidBrush brush = new SolidBrush(SystemColors.ControlLightLight))
                {
                    g.FillRectangle(brush, rect);
                }
            }

            int arm = rect.Height / 4;
            int cx = rect.X + rect.Width / 2;
            int cy = rect.Y + rect.Height / 2;
            using (Pen pen = new Pen(SystemColors.ControlText, Math.Max(1, Scale(1))))
            {
                g.DrawLine(pen, cx - arm, cy, cx + arm, cy);
                g.DrawLine(pen, cx, cy - arm, cx, cy + arm);
            }
        }

        // A thin divider and a small X. The X turns red on hover.
        private void DrawCloseAll(Graphics g, Rectangle rect)
        {
            int inset = Scale(4);
            using (Pen divider = new Pen(SystemColors.ControlDark))
            {
                g.DrawLine(divider, rect.X, rect.Y + inset, rect.X, rect.Bottom - inset - 1);
            }

            if (_hoverCloseAll)
            {
                using (SolidBrush brush = new SolidBrush(Color.FromArgb(232, 17, 35)))
                {
                    g.FillRectangle(brush, rect.X + 1, rect.Y, rect.Width - 1, rect.Height - 1);
                }
            }

            int arm = Math.Max(3, rect.Height / 5);
            int cx = rect.X + rect.Width / 2;
            int cy = rect.Y + rect.Height / 2;
            SmoothingMode old = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (Pen pen = new Pen(_hoverCloseAll ? Color.White : SystemColors.ControlText, Math.Max(1, Scale(1))))
            {
                g.DrawLine(pen, cx - arm, cy - arm, cx + arm, cy + arm);
                g.DrawLine(pen, cx - arm, cy + arm, cx + arm, cy - arm);
            }

            g.SmoothingMode = old;
        }

        private void DrawOverflow(Graphics g, Rectangle rect)
        {
            int half = rect.Height / 6;
            int cx = rect.X + rect.Width / 2;
            int cy = rect.Y + rect.Height / 2;
            Point[] triangle = new Point[]
            {
                new Point(cx - half, cy - half / 2),
                new Point(cx + half, cy - half / 2),
                new Point(cx, cy + half / 2 + 1)
            };
            using (SolidBrush brush = new SolidBrush(SystemColors.ControlText))
            {
                g.FillPolygon(brush, triangle);
            }
        }

        private void DrawEdgeLine(Graphics g)
        {
            int y = _settings.GetPosition() == TabBarPosition.Below ? 0 : ClientSize.Height - 1;
            using (Pen pen = new Pen(SystemColors.ControlDark))
            {
                g.DrawLine(pen, 0, y, ClientSize.Width, y);
            }
        }

        // ----- File icons -----

        private Bitmap GetIcon(TabItem item)
        {
            string extension = Path.GetExtension(item.Name).ToLowerInvariant();
            Bitmap bitmap;
            if (_icons.TryGetValue(extension, out bitmap))
            {
                return bitmap;
            }

            bitmap = LoadShellIcon(extension.Length > 0 ? "file" + extension : "file");
            _icons[extension] = bitmap;
            return bitmap;
        }

        private static Bitmap LoadShellIcon(string fakeName)
        {
            NativeMethods.SHFILEINFO info = new NativeMethods.SHFILEINFO();
            IntPtr result = NativeMethods.SHGetFileInfo(fakeName, NativeMethods.FILE_ATTRIBUTE_NORMAL, ref info,
                (uint)Marshal.SizeOf(typeof(NativeMethods.SHFILEINFO)),
                NativeMethods.SHGFI_ICON | NativeMethods.SHGFI_SMALLICON | NativeMethods.SHGFI_USEFILEATTRIBUTES);
            if (result == IntPtr.Zero || info.HIcon == IntPtr.Zero)
            {
                return null;
            }

            try
            {
                using (Icon icon = Icon.FromHandle(info.HIcon))
                {
                    return icon.ToBitmap();
                }
            }
            finally
            {
                NativeMethods.DestroyIcon(info.HIcon);
            }
        }

        // ----- Mouse -----

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            try
            {
                int index = _layout.HitTestTab(e.Location);
                if (e.Button == MouseButtons.Left)
                {
                    if (!_layout.OverflowRect.IsEmpty && _layout.OverflowRect.Contains(e.Location)
                        && !_settings.GetBool(SettingNames.MultiLine))
                    {
                        ShowOverflowMenu();
                    }
                    else if (index >= 0)
                    {
                        _dragIndex = index;
                        _dragStart = e.Location;
                        _dragging = false;
                        RaiseActivate(_list.Items[index].Key);
                    }
                }
                else if (e.Button == MouseButtons.Middle && index >= 0
                    && _allowClose
                    && _settings.GetBool(SettingNames.MiddleClickClose))
                {
                    RaiseClose(_list.Items[index].Key);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("TabBarControl.OnMouseDown", ex);
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            try
            {
                if (e.Button == MouseButtons.Left && _dragIndex >= 0)
                {
                    HandleDrag(e.Location);
                }

                UpdateHover(e.Location);
            }
            catch (Exception ex)
            {
                Logger.Error("TabBarControl.OnMouseMove", ex);
            }
        }

        private void HandleDrag(Point location)
        {
            if (!_settings.GetBool(SettingNames.DragToReorder))
            {
                return;
            }

            Size dragSize = SystemInformation.DragSize;
            if (!_dragging && Math.Abs(location.X - _dragStart.X) < dragSize.Width
                && Math.Abs(location.Y - _dragStart.Y) < dragSize.Height)
            {
                return;
            }

            _dragging = true;
            int target = _layout.HitTestTab(location);
            if (target >= 0 && target != _dragIndex)
            {
                _list.Move(_dragIndex, target);
                _dragIndex = target;
            }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            try
            {
                bool wasDragging = _dragging;
                _dragIndex = -1;
                _dragging = false;
                if (e.Button == MouseButtons.Left && !wasDragging && !_layout.CloseAllRect.IsEmpty
                    && _layout.CloseAllRect.Contains(e.Location))
                {
                    Action closeAll = CloseAllRequested;
                    if (closeAll != null)
                    {
                        closeAll();
                    }

                    return;
                }

                if (e.Button == MouseButtons.Left && !wasDragging && !_layout.NewButtonRect.IsEmpty
                    && _layout.NewButtonRect.Contains(e.Location))
                {
                    Action handler = NewDocumentRequested;
                    if (handler != null)
                    {
                        UseWaitCursor = true;
                        _waitTimer.Stop();
                        _waitTimer.Start();
                        handler();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("TabBarControl.OnMouseUp", ex);
            }
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);
            try
            {
                int index = _layout.HitTestTab(e.Location);
                if (e.Button == MouseButtons.Left && index >= 0 && _allowClose
                    && _settings.GetBool(SettingNames.DoubleClickClose))
                {
                    RaiseClose(_list.Items[index].Key);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("TabBarControl.OnMouseDoubleClick", ex);
            }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            UpdateHover(new Point(-1, -1));
        }

        private void UpdateHover(Point location)
        {
            int index = _layout.HitTestTab(location);
            bool overNew = !_layout.NewButtonRect.IsEmpty && _layout.NewButtonRect.Contains(location);
            bool overClose = !_layout.CloseAllRect.IsEmpty && _layout.CloseAllRect.Contains(location);
            if (index == _hoverIndex && overNew == _hoverNew && overClose == _hoverCloseAll)
            {
                return;
            }

            _hoverIndex = index;
            _hoverNew = overNew;
            _hoverCloseAll = overClose;
            string tip = string.Empty;
            if (index >= 0 && index < _list.Count)
            {
                tip = _list.Items[index].FullPath;
            }
            else if (overNew)
            {
                tip = "New document";
            }
            else if (overClose)
            {
                tip = "Close all documents";
            }

            _toolTip.SetToolTip(this, tip);
            Invalidate();
        }

        // ----- Overflow menu -----

        private void ShowOverflowMenu()
        {
            if (_overflowMenu != null)
            {
                _overflowMenu.Dispose();
            }

            _overflowMenu = new ContextMenuStrip();
            int maxChars = _settings.GetInt(SettingNames.MaxTabChars);
            foreach (TabItem item in _list.Items)
            {
                string key = item.Key;
                ToolStripMenuItem entry = new ToolStripMenuItem(TabLabeler.Display(item, maxChars));
                entry.Checked = key == _list.CurrentKey;
                entry.Click += delegate { RaiseActivate(key); };
                _overflowMenu.Items.Add(entry);
            }

            _overflowMenu.Show(this, new Point(_layout.OverflowRect.Left, _layout.OverflowRect.Bottom));
        }

        // ----- Raising events -----

        private void RaiseActivate(string key)
        {
            Action<string> handler = ActivateRequested;
            if (handler != null)
            {
                handler(key);
            }
        }

        private void RaiseClose(string key)
        {
            if (!_allowClose)
            {
                return;
            }

            Action<string> handler = CloseRequested;
            if (handler != null)
            {
                handler(key);
            }
        }

        // ----- Cleanup -----

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _list.Changed -= OnListChanged;
                _waitTimer.Stop();
                _waitTimer.Dispose();
                _toolTip.Dispose();
                if (_overflowMenu != null)
                {
                    _overflowMenu.Dispose();
                }

                if (_font != null)
                {
                    _font.Dispose();
                }

                if (_boldFont != null)
                {
                    _boldFont.Dispose();
                }

                foreach (Bitmap bitmap in _icons.Values)
                {
                    if (bitmap != null)
                    {
                        bitmap.Dispose();
                    }
                }
            }

            base.Dispose(disposing);
        }
    }
}
