// ---------------------------------------------------------------------------------------------------------------
// Name:     ProjectAdapter.cs
// Purpose:  Microsoft Project adapter. Project keeps every open project inside one main window
//           (JWinproj-WhimperMainClass), as a child window in its MDIClient, so all tabs share one bar in that
//           window and no new main windows are ever created. Microsoft publishes no Project interop package,
//           so Project is reached by late binding: properties and methods are called by name at run time. Late
//           binding gives no event hookup, so a short timer reads the open projects, their names, their saved
//           state and the active project, and reports changes. A project is identified by its COM object, which
//           stays the same object for as long as the project is open. Project builds and removes a project
//           window piece by piece, so the document area (MDIClient) stops drawing while a project is added or
//           closed from the tab bar, then repaints once. The ribbon and the tab bar keep drawing. Closing first
//           asks about unsaved changes; Yes runs Project's own Save, which shows Project's own Save As for a
//           project never saved. The project is then closed without a second prompt.
// Revision: 6
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
    internal sealed class ProjectAdapter : IAppAdapter
    {
        private const int PollMilliseconds = 250;
        private const int WindowClassCapacity = 128;
        private const int DoNotSave = 0;

        private readonly object _app;
        private readonly Dictionary<string, object> _projects = new Dictionary<string, object>();
        private readonly Dictionary<string, string> _names = new Dictionary<string, string>();
        private readonly Dictionary<string, string> _paths = new Dictionary<string, string>();
        private readonly Dictionary<string, bool> _dirty = new Dictionary<string, bool>();
        private readonly List<string> _order = new List<string>();
        private TabCoordinator _coordinator;
        private Timer _pollTimer;
        private IntPtr _frame;
        private bool _refreshing;
        private int _nextId;

        public ProjectAdapter(object app)
        {
            _app = app;
        }

        public string AppName
        {
            get { return "Project"; }
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
                Logger.Info("Project adapter started.");
                Refresh();
            }
            catch (Exception ex)
            {
                Logger.Error("ProjectAdapter.Start", ex);
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

                Logger.Info("Project adapter stopped.");
            }
            catch (Exception ex)
            {
                Logger.Error("ProjectAdapter.Stop", ex);
            }
        }

        // ----- Reading Project -----

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
            object projects = null;
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

                projects = Get(_app, "Projects");
                int count = Convert.ToInt32(Get(projects, "Count"));
                HashSet<string> seen = new HashSet<string>();
                for (int i = 1; i <= count; i++)
                {
                    object project = Get(projects, "Item", i);
                    string key = FindKnown(project);
                    if (key == null)
                    {
                        _nextId++;
                        key = "J" + _nextId.ToString("D6");
                        _projects[key] = project;
                        _order.Add(key);
                        Logger.Info("Project registered " + key + ".");
                    }

                    seen.Add(key);
                    Report(key, project);
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
                // Project refuses object model calls while a cell is being edited or a dialog is open. The next
                // tick reads again.
            }
            catch (Exception ex)
            {
                Logger.Error("ProjectAdapter.Refresh", ex);
            }
            finally
            {
                Release(projects);
                _refreshing = false;
            }
        }

        private void Report(string key, object project)
        {
            string name = Convert.ToString(Get(project, "Name"));
            string path = ProjectPath(project, name);
            bool dirty = !Convert.ToBoolean(Get(project, "Saved"));
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
                active = Get(_app, "ActiveProject");
                string key = FindKnown(active);
                if (key != null && _coordinator.Tabs.Find(key) != null && _coordinator.Tabs.CurrentKey != key)
                {
                    _coordinator.DocumentActivated(key);
                }
            }
            catch (COMException)
            {
                // There is no active project when the last one has just closed.
            }
            finally
            {
                ReleaseTransient(active);
            }
        }

        private static string ProjectPath(object project, string name)
        {
            try
            {
                string fullName = Convert.ToString(Get(project, "FullName"));
                return string.IsNullOrEmpty(fullName) ? name : fullName;
            }
            catch (COMException)
            {
                return name;
            }
        }

        // The project's first window, for the task pane bar used when EmbedTabBar is off.
        private object WindowOf(string key)
        {
            object project;
            if (!_projects.TryGetValue(key, out project))
            {
                return null;
            }

            object windows = null;
            try
            {
                windows = Get(project, "Windows");
                return Get(windows, "Item", 1);
            }
            catch (Exception)
            {
                return null;
            }
            finally
            {
                Release(windows);
            }
        }

        // ----- Requests from the tab bar -----

        public void ActivateDocument(string key)
        {
            object project;
            if (!_projects.TryGetValue(key, out project))
            {
                return;
            }

            try
            {
                Call(project, "Activate");
                ReportActive();
            }
            catch (Exception ex)
            {
                Logger.Error("ProjectAdapter.ActivateDocument", ex);
            }
        }

        // FileCloseEx closes the active project, so the project is activated first and checked to be the active
        // one, which makes sure the wrong project is never closed. Unsaved changes are asked about before drawing
        // stops, so the question is never shown over a frozen window. A project that is still open afterwards
        // was kept by the user.
        public bool CloseDocument(string key)
        {
            object project;
            if (!_projects.TryGetValue(key, out project))
            {
                return true;
            }

            Call(project, "Activate");
            object active = null;
            bool isActive;
            try
            {
                active = Get(_app, "ActiveProject");
                isActive = ReferenceEquals(active, project);
            }
            finally
            {
                ReleaseTransient(active);
            }

            if (!isActive)
            {
                Logger.Info("Close skipped for " + key + ": Project did not make it the active project.");
                return false;
            }

            if (!Convert.ToBoolean(Get(project, "Saved")) && !AskAndSave(project))
            {
                return false;
            }

            IntPtr frozen = Freeze();
            try
            {
                Call(_app, "FileCloseEx", DoNotSave);
                if (StillOpen(project))
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

        // Asks Save, Don't Save or Cancel for the active project. Yes runs Project's own Save, which shows Project's
        // Save As when the project has never been saved. Returns false when the user cancels either question.
        private bool AskAndSave(object project)
        {
            string name = Convert.ToString(Get(project, "Name"));
            DialogResult answer = MessageBox.Show(new OwnerWindow(_frame),
                "Do you want to save changes to '" + name + "'?", "Microsoft Project", MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Warning);
            if (answer == DialogResult.Cancel)
            {
                return false;
            }

            if (answer == DialogResult.No)
            {
                return true;
            }

            try
            {
                Call(_app, "FileSave");
            }
            catch (COMException)
            {
                // Project reports a cancelled Save As as an error.
                return false;
            }

            return Convert.ToBoolean(Get(project, "Saved"));
        }

        private bool StillOpen(object project)
        {
            object projects = null;
            try
            {
                projects = Get(_app, "Projects");
                int count = Convert.ToInt32(Get(projects, "Count"));
                for (int i = 1; i <= count; i++)
                {
                    object candidate = Get(projects, "Item", i);
                    if (ReferenceEquals(candidate, project))
                    {
                        return true;
                    }

                    ReleaseTransient(candidate);
                }

                return false;
            }
            finally
            {
                Release(projects);
            }
        }

        public void NewDocument()
        {
            object projects = null;
            object created = null;
            IntPtr frozen = Freeze();
            try
            {
                projects = Get(_app, "Projects");
                created = Call(projects, "Add");

                // The new project is made Project's active project and its tab is selected in the same step
                // as the tab is added, so the bar shows the new tab already selected instead of selecting it on a
                // later timer tick. The created object is kept, so Refresh files the project under it.
                try
                {
                    Call(created, "Activate");
                }
                catch (COMException)
                {
                    // Project rejects Activate on a project it has just created. It has already made that project
                    // the active one, so there is nothing to do.
                }

                Refresh();
                string key = FindKnown(created);
                if (key != null && _coordinator.Tabs.Find(key) != null && _coordinator.Tabs.CurrentKey != key)
                {
                    _coordinator.DocumentActivated(key);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("ProjectAdapter.NewDocument", ex);
            }
            finally
            {
                ReleaseTransient(created);
                Release(projects);
                Thaw(frozen);
            }
        }

        // Called after close-all has closed every project, when the user chose to close Project as well.
        public void QuitApplication()
        {
            Call(_app, "Quit");
        }

        // ----- Bookkeeping -----

        private string FindKnown(object project)
        {
            if (project == null)
            {
                return null;
            }

            foreach (KeyValuePair<string, object> pair in _projects)
            {
                if (ReferenceEquals(pair.Value, project))
                {
                    return pair.Key;
                }
            }

            return null;
        }

        private void Forget(string key)
        {
            object project;
            if (!_projects.TryGetValue(key, out project))
            {
                return;
            }

            if (_coordinator != null && _coordinator.Tabs.Find(key) != null)
            {
                _coordinator.DocumentClosed(key);
            }

            _projects.Remove(key);
            _names.Remove(key);
            _paths.Remove(key);
            _dirty.Remove(key);
            _order.Remove(key);
            Release(project);
            Logger.Info("Project removed " + key + ".");
        }

        // A COM call can return the same object already kept for an open project. Releasing that shared object
        // would break its tab, so only objects that are not kept are released.
        private void ReleaseTransient(object value)
        {
            if (value != null && FindKnown(value) == null)
            {
                Release(value);
            }
        }

        // The main Project window of this process.
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

        // Stops Project's document area and the project windows in it from drawing. Returns the area to pass to
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

        // Lets the document area draw again and repaints it and every project window in it once.
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

        // The MDIClient directly inside the main Project window.
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

        // Reflection wraps errors raised by Project. The original error is passed on, so a COMException from
        // Project arrives as a COMException.
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
