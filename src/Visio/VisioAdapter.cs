// ---------------------------------------------------------------------------------------------------------------
// Name:     VisioAdapter.cs
// Purpose:  Microsoft Visio adapter. Visio keeps every open drawing inside one main window (VISIOA), as a child
//           window (VISIOS) in its MDIClient, so all tabs share one bar in that window and no new main windows
//           are ever created. Visio is reached by late binding: properties and methods are called by name at
//           run time, with no interop package. A short timer reads the open drawings, their names, their saved
//           state and the active document, and reports changes. Only drawings (document type 1) that have a
//           drawing window get tabs; stencils and templates do not. A drawing is identified by its COM object,
//           which stays the same object for as long as the drawing is open. The document area (MDIClient) stops
//           drawing while a drawing is added or closed from the tab bar, then repaints once. Visio's Close
//           discards changes without asking, so closing first asks Save, Don't Save or Cancel. Yes saves in
//           place, or asks for a .vsdx file name when the drawing was never saved or is read-only. No marks the
//           drawing saved, so Visio closes it without asking again.
// Revision: 2
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace OpenTools
{
    internal sealed class VisioAdapter : IAppAdapter
    {
        private const int PollMilliseconds = 250;
        private const int WindowClassCapacity = 128;
        private const int DrawingType = 1;

        private readonly object _app;
        private readonly Dictionary<string, object> _drawings = new Dictionary<string, object>();
        private readonly Dictionary<string, string> _names = new Dictionary<string, string>();
        private readonly Dictionary<string, string> _paths = new Dictionary<string, string>();
        private readonly Dictionary<string, bool> _dirty = new Dictionary<string, bool>();
        private readonly List<string> _order = new List<string>();
        private TabCoordinator _coordinator;
        private Timer _pollTimer;
        private IntPtr _frame;
        private bool _refreshing;
        private int _nextId;

        public VisioAdapter(object app)
        {
            _app = app;
        }

        public string AppName
        {
            get { return "Visio"; }
        }

        public void Start(TabCoordinator coordinator)
        {
            try
            {
                _coordinator = coordinator;
                _pollTimer = new Timer();
                _pollTimer.Interval = PollMilliseconds;
                _pollTimer.Tick += OnPollTick;
                _pollTimer.Start();
                Logger.Info("Visio adapter started.");
                Refresh();
            }
            catch (Exception ex)
            {
                Logger.Error("VisioAdapter.Start", ex);
                Stop();
            }
        }

        public void Stop()
        {
            try
            {
                if (_pollTimer != null)
                {
                    _pollTimer.Stop();
                    _pollTimer.Tick -= OnPollTick;
                    _pollTimer.Dispose();
                    _pollTimer = null;
                }

                foreach (string key in new List<string>(_order))
                {
                    Forget(key);
                }

                Logger.Info("Visio adapter stopped.");
            }
            catch (Exception ex)
            {
                Logger.Error("VisioAdapter.Stop", ex);
            }
        }

        // ----- Reading Visio -----

        private void OnPollTick(object sender, EventArgs e)
        {
            Refresh();
        }

        private void Refresh()
        {
            if (_refreshing || _coordinator == null)
            {
                return;
            }

            _refreshing = true;
            object documents = null;
            try
            {
                if (_frame == IntPtr.Zero || !NativeMethods.IsWindow(_frame))
                {
                    _frame = FindFrame();
                }

                if (_frame == IntPtr.Zero)
                {
                    return;
                }

                documents = Get(_app, "Documents");
                int count = Convert.ToInt32(Get(documents, "Count"));
                HashSet<string> seen = new HashSet<string>();
                for (int i = 1; i <= count; i++)
                {
                    object document = Get(documents, "Item", i);
                    string key = FindKnown(document);
                    if (key == null)
                    {
                        // Stencils, templates and drawings without a drawing window yet are left out. A new drawing
                        // whose window is not open yet is picked up on a later tick.
                        if (Convert.ToInt32(Get(document, "Type")) != DrawingType || !HasDrawingWindow(document))
                        {
                            ReleaseTransient(document);
                            continue;
                        }

                        _nextId++;
                        key = "V" + _nextId.ToString("D6");
                        _drawings[key] = document;
                        _order.Add(key);
                        Logger.Info("Visio drawing registered " + key + ".");
                    }

                    seen.Add(key);
                    Report(key, document);
                }

                foreach (string key in new List<string>(_order))
                {
                    if (!seen.Contains(key))
                    {
                        Forget(key);
                    }
                }

                ReportActive();
            }
            catch (COMException)
            {
                // Visio refuses object model calls while text is being edited or a dialog is open. The next tick
                // reads again.
            }
            catch (Exception ex)
            {
                Logger.Error("VisioAdapter.Refresh", ex);
            }
            finally
            {
                Release(documents);
                _refreshing = false;
            }
        }

        private void Report(string key, object drawing)
        {
            string name = Convert.ToString(Get(drawing, "Name"));
            string path = DrawingPath(drawing, name);
            bool dirty = !Convert.ToBoolean(Get(drawing, "Saved"));
            if (_coordinator.Tabs.Find(key) == null)
            {
                _names[key] = name;
                _paths[key] = path;
                _dirty[key] = dirty;
                _coordinator.DocumentOpened(key, _frame, () => WindowOf(key), path, name, dirty);
                return;
            }

            string oldName;
            string oldPath;
            _names.TryGetValue(key, out oldName);
            _paths.TryGetValue(key, out oldPath);
            if (oldName != name || oldPath != path)
            {
                _names[key] = name;
                _paths[key] = path;
                _coordinator.DocumentRenamed(key, path, name);
            }

            bool oldDirty;
            if (!_dirty.TryGetValue(key, out oldDirty) || oldDirty != dirty)
            {
                _dirty[key] = dirty;
                _coordinator.DocumentDirtyChanged(key, dirty);
            }
        }

        private void ReportActive()
        {
            object active = null;
            try
            {
                active = Get(_app, "ActiveDocument");
                string key = FindKnown(active);
                if (key != null && _coordinator.Tabs.Find(key) != null && _coordinator.Tabs.CurrentKey != key)
                {
                    _coordinator.DocumentActivated(key);
                }
            }
            catch (COMException)
            {
                // There is no active document when the last drawing has just closed.
            }
            finally
            {
                ReleaseTransient(active);
            }
        }

        private static string DrawingPath(object drawing, string name)
        {
            try
            {
                string fullName = Convert.ToString(Get(drawing, "FullName"));
                return string.IsNullOrEmpty(fullName) ? name : fullName;
            }
            catch (COMException)
            {
                return name;
            }
        }

        private bool HasDrawingWindow(object drawing)
        {
            object window = FindDrawingWindow(drawing);
            Release(window);
            return window != null;
        }

        // The drawing window that shows this drawing, or null. The caller releases what is returned.
        private object FindDrawingWindow(object drawing)
        {
            object windows = null;
            try
            {
                windows = Get(_app, "Windows");
                int count = Convert.ToInt32(Get(windows, "Count"));
                for (int i = 1; i <= count; i++)
                {
                    object window = Get(windows, "Item", i);
                    object shown = null;
                    bool match = false;
                    try
                    {
                        if (Convert.ToInt32(Get(window, "Type")) == DrawingType)
                        {
                            shown = Get(window, "Document");
                            match = ReferenceEquals(shown, drawing);
                        }
                    }
                    catch (COMException)
                    {
                        match = false;
                    }
                    finally
                    {
                        if (!ReferenceEquals(shown, drawing))
                        {
                            ReleaseTransient(shown);
                        }
                    }

                    if (match)
                    {
                        return window;
                    }

                    Release(window);
                }

                return null;
            }
            finally
            {
                Release(windows);
            }
        }

        // The drawing's window, for the task pane bar used when EmbedTabBar is off.
        private object WindowOf(string key)
        {
            object drawing;
            if (!_drawings.TryGetValue(key, out drawing))
            {
                return null;
            }

            try
            {
                return FindDrawingWindow(drawing);
            }
            catch (Exception)
            {
                return null;
            }
        }

        // ----- Requests from the tab bar -----

        public void ActivateDocument(string key)
        {
            object drawing;
            if (!_drawings.TryGetValue(key, out drawing))
            {
                return;
            }

            try
            {
                ShowDrawing(drawing);
                ReportActive();
            }
            catch (Exception ex)
            {
                Logger.Error("VisioAdapter.ActivateDocument", ex);
            }
        }

        private void ShowDrawing(object drawing)
        {
            object window = FindDrawingWindow(drawing);
            try
            {
                if (window != null)
                {
                    Call(window, "Activate");
                }
            }
            finally
            {
                Release(window);
            }
        }

        // The drawing is brought forward so the user sees which one is being asked about. Unsaved changes are
        // asked about before drawing stops, so the question is never shown over a frozen window. A drawing that is
        // still open afterwards was kept by the user.
        public bool CloseDocument(string key)
        {
            object drawing;
            if (!_drawings.TryGetValue(key, out drawing))
            {
                return true;
            }

            ShowDrawing(drawing);
            if (!Convert.ToBoolean(Get(drawing, "Saved")) && !AskAndSave(drawing))
            {
                return false;
            }

            IntPtr frozen = Freeze();
            try
            {
                Call(drawing, "Close");
                if (StillOpen(drawing))
                {
                    return false;
                }

                Forget(key);
                return true;
            }
            finally
            {
                Thaw(frozen);
            }
        }

        // Asks Save, Don't Save or Cancel. Returns false when the user cancels this question or the Save As dialog.
        private bool AskAndSave(object drawing)
        {
            string name = Convert.ToString(Get(drawing, "Name"));
            DialogResult answer = MessageBox.Show(new OwnerWindow(_frame),
                "Do you want to save changes to '" + name + "'?", "Microsoft Visio", MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Warning);
            if (answer == DialogResult.Cancel)
            {
                return false;
            }

            if (answer == DialogResult.No)
            {
                // Marked saved so Visio closes the drawing without asking again. The changes are discarded.
                Put(drawing, "Saved", true);
                return true;
            }

            string path = Convert.ToString(Get(drawing, "Path"));
            if (!string.IsNullOrEmpty(path) && !Convert.ToBoolean(Get(drawing, "ReadOnly")))
            {
                Call(drawing, "Save");
                return true;
            }

            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Title = "Save As";
                dialog.FileName = name;
                dialog.Filter = "Visio Drawing (*.vsdx)|*.vsdx";
                dialog.DefaultExt = "vsdx";
                dialog.AddExtension = true;
                dialog.OverwritePrompt = true;
                if (dialog.ShowDialog(new OwnerWindow(_frame)) != DialogResult.OK)
                {
                    return false;
                }

                Call(drawing, "SaveAs", dialog.FileName);
                return true;
            }
        }

        private bool StillOpen(object drawing)
        {
            object documents = null;
            try
            {
                documents = Get(_app, "Documents");
                int count = Convert.ToInt32(Get(documents, "Count"));
                for (int i = 1; i <= count; i++)
                {
                    object candidate = Get(documents, "Item", i);
                    if (ReferenceEquals(candidate, drawing))
                    {
                        return true;
                    }

                    ReleaseTransient(candidate);
                }

                return false;
            }
            finally
            {
                Release(documents);
            }
        }

        public void NewDocument()
        {
            object documents = null;
            object created = null;
            IntPtr frozen = Freeze();
            try
            {
                documents = Get(_app, "Documents");
                created = Call(documents, "Add", string.Empty);

                // The new drawing is shown and its tab is selected in the same step as the tab is added, so the bar
                // shows the new tab already selected. The created object is kept, so Refresh files it under it.
                ShowDrawing(created);
                Refresh();
                string key = FindKnown(created);
                if (key != null && _coordinator.Tabs.Find(key) != null && _coordinator.Tabs.CurrentKey != key)
                {
                    _coordinator.DocumentActivated(key);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("VisioAdapter.NewDocument", ex);
            }
            finally
            {
                ReleaseTransient(created);
                Release(documents);
                Thaw(frozen);
            }
        }

        // Called after close-all has closed every drawing, when the user chose to close Visio as well.
        public void QuitApplication()
        {
            Call(_app, "Quit");
        }

        // ----- Bookkeeping -----

        private string FindKnown(object drawing)
        {
            if (drawing == null)
            {
                return null;
            }

            foreach (KeyValuePair<string, object> pair in _drawings)
            {
                if (ReferenceEquals(pair.Value, drawing))
                {
                    return pair.Key;
                }
            }

            return null;
        }

        private void Forget(string key)
        {
            object drawing;
            if (!_drawings.TryGetValue(key, out drawing))
            {
                return;
            }

            if (_coordinator != null && _coordinator.Tabs.Find(key) != null)
            {
                _coordinator.DocumentClosed(key);
            }

            _drawings.Remove(key);
            _names.Remove(key);
            _paths.Remove(key);
            _dirty.Remove(key);
            _order.Remove(key);
            Release(drawing);
            Logger.Info("Visio drawing removed " + key + ".");
        }

        // A COM call can return the same object already kept for an open drawing. Releasing that shared object
        // would break its tab, so only objects that are not kept are released.
        private void ReleaseTransient(object value)
        {
            if (value != null && FindKnown(value) == null)
            {
                Release(value);
            }
        }

        // The main Visio window of this process.
        private static IntPtr FindFrame()
        {
            IntPtr found = IntPtr.Zero;
            uint processId = (uint)Process.GetCurrentProcess().Id;
            NativeMethods.EnumWindows(delegate(IntPtr hWnd, IntPtr lParam)
            {
                uint owner;
                NativeMethods.GetWindowThreadProcessId(hWnd, out owner);
                if (owner == processId && NativeMethods.IsWindowVisible(hWnd) && ClassOf(hWnd) == AppProfile.FrameClass)
                {
                    found = hWnd;
                    return false;
                }

                return true;
            }, IntPtr.Zero);
            return found;
        }

        private static string ClassOf(IntPtr hWnd)
        {
            StringBuilder text = new StringBuilder(WindowClassCapacity);
            NativeMethods.GetClassName(hWnd, text, text.Capacity);
            return text.ToString();
        }

        // ----- Drawing -----

        // Stops Visio's document area and the drawing windows in it from drawing. Returns the area to pass to
        // Thaw.
        private IntPtr Freeze()
        {
            IntPtr area = DocumentArea();
            if (area != IntPtr.Zero)
            {
                NativeMethods.SendMessage(area, NativeMethods.WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);
            }

            return area;
        }

        // Lets the document area draw again and repaints it and every drawing window in it once.
        private static void Thaw(IntPtr area)
        {
            if (area == IntPtr.Zero || !NativeMethods.IsWindow(area))
            {
                return;
            }

            NativeMethods.SendMessage(area, NativeMethods.WM_SETREDRAW, new IntPtr(1), IntPtr.Zero);
            NativeMethods.RedrawWindow(area, IntPtr.Zero, IntPtr.Zero, NativeMethods.RDW_INVALIDATE
                | NativeMethods.RDW_ERASE | NativeMethods.RDW_FRAME | NativeMethods.RDW_ALLCHILDREN
                | NativeMethods.RDW_UPDATENOW);
        }

        // The MDIClient directly inside the main Visio window.
        private IntPtr DocumentArea()
        {
            if (_frame == IntPtr.Zero || !NativeMethods.IsWindow(_frame))
            {
                return IntPtr.Zero;
            }

            for (IntPtr child = NativeMethods.GetWindow(_frame, NativeMethods.GW_CHILD); child != IntPtr.Zero;
                child = NativeMethods.GetWindow(child, NativeMethods.GW_HWNDNEXT))
            {
                if (ClassOf(child) == AppProfile.DocumentAreaClass)
                {
                    return child;
                }
            }

            return IntPtr.Zero;
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

        // ----- Late binding -----

        private static object Get(object target, string name, params object[] arguments)
        {
            return Invoke(target, name, BindingFlags.GetProperty, arguments);
        }

        private static object Call(object target, string name, params object[] arguments)
        {
            return Invoke(target, name, BindingFlags.InvokeMethod, arguments);
        }

        private static void Put(object target, string name, object value)
        {
            Invoke(target, name, BindingFlags.SetProperty, new object[] { value });
        }

        // Reflection wraps errors raised by Visio. The original error is passed on, so a COMException from
        // Visio arrives as a COMException.
        private static object Invoke(object target, string name, BindingFlags flags, object[] arguments)
        {
            try
            {
                return target.GetType().InvokeMember(name, flags, null, target, arguments);
            }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
                throw;
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }
}
