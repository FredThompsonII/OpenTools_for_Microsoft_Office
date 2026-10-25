// ---------------------------------------------------------------------------------------------------------------
// Name:     PowerPointAdapter.cs
// Purpose:  P2 PowerPoint adapter. Tabs identify presentation objects, not their changing frame captions.
//           A separate verified frame mapping supports shared/reused frames and stable tab creation order.
//           Slide shows remain excluded. Closing a tab asks about unsaved changes the way PowerPoint does
//           (Save, Don't Save, Cancel; Save As for a presentation never saved), then saves and closes the
//           presentation directly. Every step finishes before the close call returns, so close-all works.
// Revision: 9
// DTG:      261024233000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OpenTools
{
    internal sealed class PowerPointAdapter : IAppAdapter
    {
        private const int PollMilliseconds = 250;
        private const int WindowClassCapacity = 128;
        private const int WindowTitleCapacity = 512;
        private const int ActiveCandidateMaxAgeMilliseconds = 750;
        private const string FrameClass = "PPTFrameClass";
        private const string DocumentAreaClass = "MDIClient";

        private readonly PowerPoint.Application _app;
        private readonly Dictionary<string, PowerPoint.Presentation> _presentations =
            new Dictionary<string, PowerPoint.Presentation>();
        private readonly Dictionary<string, IntPtr> _frames = new Dictionary<string, IntPtr>();
        private string _closingKey;
        private readonly Dictionary<string, string> _names = new Dictionary<string, string>();
        private readonly Dictionary<string, string> _paths = new Dictionary<string, string>();
        private readonly List<string> _creationOrder = new List<string>();
        private PowerPoint.EApplication_Event _events;
        private TabCoordinator _coordinator;
        private Timer _pollTimer;
        private bool _refreshing;
        private int _nextId;
        private string _lastUnmatched = string.Empty;
        private string _activeCandidateKey;
        private IntPtr _activeCandidateFrame;
        private int _activeCandidateTick;

        public PowerPointAdapter(PowerPoint.Application app)
        {
            _app = app;
        }

        public string AppName
        {
            get { return "PowerPoint"; }
        }

        public void Start(TabCoordinator coordinator)
        {
            try
            {
                _coordinator = coordinator;
                _events = (PowerPoint.EApplication_Event)_app;
                _events.NewPresentation += OnNewPresentation;
                _events.PresentationOpen += OnPresentationOpen;
                _events.PresentationBeforeClose += OnPresentationBeforeClose;
                _events.PresentationClose += OnPresentationClose;
                _events.WindowActivate += OnWindowActivate;
                _events.WindowDeactivate += OnWindowDeactivate;
                _events.SlideShowBegin += OnSlideShowBegin;
                _events.SlideShowEnd += OnSlideShowEnd;

                _pollTimer = new Timer();
                _pollTimer.Interval = PollMilliseconds;
                _pollTimer.Tick += OnPollTick;
                _pollTimer.Start();
                Logger.Info("P2 PowerPoint adapter started. Close actions enabled.");
                RefreshWindows();
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointAdapter.Start", ex);
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

                if (_events != null)
                {
                    _events.NewPresentation -= OnNewPresentation;
                    _events.PresentationOpen -= OnPresentationOpen;
                    _events.PresentationBeforeClose -= OnPresentationBeforeClose;
                    _events.PresentationClose -= OnPresentationClose;
                    _events.WindowActivate -= OnWindowActivate;
                    _events.WindowDeactivate -= OnWindowDeactivate;
                    _events.SlideShowBegin -= OnSlideShowBegin;
                    _events.SlideShowEnd -= OnSlideShowEnd;
                    _events = null;
                }

                foreach (string key in new List<string>(_creationOrder))
                {
                    Forget(key);
                }

                Logger.Info("P2 PowerPoint adapter stopped.");
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointAdapter.Stop", ex);
            }
        }

        // ----- Events establish presentation creation order before frames become visible. -----

        private void OnNewPresentation(PowerPoint.Presentation presentation)
        {
            RememberEvent("NewPresentation", presentation);
        }

        private void OnPresentationOpen(PowerPoint.Presentation presentation)
        {
            RememberEvent("PresentationOpen", presentation);
        }

        private void OnPresentationBeforeClose(PowerPoint.Presentation presentation, ref bool cancel)
        {
            RememberEvent("PresentationBeforeClose", presentation);
        }

        private void OnPresentationClose(PowerPoint.Presentation presentation)
        {
            try
            {
                string key = FindKnown(presentation);
                Logger.Info("P2 event PresentationClose" + (key == null ? " (unmatched)." : " " + key + "."));
                // A presentation being closed by CloseDocument is forgotten there, after Close returns, so its
                // object is not released while Close is still running on it.
                if (key != null && key != _closingKey)
                {
                    Forget(key);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointAdapter.PresentationClose", ex);
            }
        }

        private void OnWindowActivate(PowerPoint.Presentation presentation, PowerPoint.DocumentWindow window)
        {
            RememberEvent("WindowActivate", presentation);
        }

        private void OnWindowDeactivate(PowerPoint.Presentation presentation, PowerPoint.DocumentWindow window)
        {
            RememberEvent("WindowDeactivate", presentation);
        }

        private void OnSlideShowBegin(PowerPoint.SlideShowWindow window)
        {
            LogEvent("SlideShowBegin (excluded)");
        }

        private void OnSlideShowEnd(PowerPoint.Presentation presentation)
        {
            LogEvent("SlideShowEnd");
        }

        private void RememberEvent(string eventName, PowerPoint.Presentation presentation)
        {
            try
            {
                string key = Remember(presentation);
                Logger.Info("P2 event " + eventName + (key == null ? " (unmatched)." : " " + key + "."));
                // Event callbacks often precede Win32 frame visibility. The poll waits for a verified match.
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointAdapter." + eventName, ex);
            }
        }

        private static void LogEvent(string eventName)
        {
            try
            {
                Logger.Info("P2 event " + eventName + ".");
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointAdapter." + eventName, ex);
            }
        }

        private void OnPollTick(object sender, EventArgs args)
        {
            try
            {
                RefreshWindows();
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointAdapter.OnPollTick", ex);
            }
        }

        // ----- Stable identity belongs to a live Presentation, never a window index or frame handle. -----

        private string Remember(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return null;
            }

            string known = FindKnown(presentation);
            if (known != null)
            {
                return known;
            }

            string name = presentation.Name;
            string path = PresentationPath(presentation, name);
            _nextId++;
            string key = "P" + _nextId.ToString("D6", System.Globalization.CultureInfo.InvariantCulture);
            _presentations.Add(key, presentation);
            _creationOrder.Add(key);
            _names.Add(key, name);
            _paths.Add(key, path);
            Logger.Info("P2 registered " + key + " presentation='" + name + "'.");
            return key;
        }

        private string FindKnown(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return null;
            }

            foreach (KeyValuePair<string, PowerPoint.Presentation> item in _presentations)
            {
                if (object.ReferenceEquals(item.Value, presentation))
                {
                    return item.Key;
                }
            }

            // PowerPoint can supply a fresh COM wrapper. The presentation's own live name and path,
            // not the changing Office Windows(index), must identify it. Ambiguous names are rejected.
            string name = presentation.Name;
            string path = PresentationPath(presentation, name);
            string found = null;
            foreach (KeyValuePair<string, PowerPoint.Presentation> item in _presentations)
            {
                string otherName;
                string otherPath;
                try
                {
                    otherName = item.Value.Name;
                    otherPath = PresentationPath(item.Value, otherName);
                }
                catch (COMException)
                {
                    continue;
                }

                if (!string.Equals(name, otherName, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(path, otherPath, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (found != null)
                {
                    return null;
                }

                found = item.Key;
            }

            return found;
        }

        private static string PresentationPath(PowerPoint.Presentation presentation, string name)
        {
            try
            {
                string fullName = presentation.FullName;
                return string.IsNullOrEmpty(fullName) ? name : fullName;
            }
            catch (COMException)
            {
                return name;
            }
        }

        private void DiscoverPresentations()
        {
            PowerPoint.Presentations presentations = null;
            try
            {
                presentations = _app.Presentations;
                int count = presentations.Count;
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Presentation presentation = null;
                    bool keep = false;
                    try
                    {
                        presentation = presentations[i];
                        string key = FindKnown(presentation);
                        if (key == null)
                        {
                            Remember(presentation);
                            keep = true;
                        }
                    }
                    catch (COMException)
                    {
                        // A presentation may be opening or closing while the collection is queried.
                    }
                    finally
                    {
                        if (!keep)
                        {
                            ReleaseTransientPresentation(presentation);
                        }
                    }
                }
            }
            finally
            {
                Release(presentations);
            }
        }

        // ----- A frame is a view location. More than one presentation may use the same frame. -----

        private void RefreshWindows()
        {
            if (_refreshing || _coordinator == null)
            {
                return;
            }

            _refreshing = true;
            try
            {
                DiscoverPresentations();
                Dictionary<IntPtr, string> frames = CollectEditingFrames();
                foreach (string key in new List<string>(_creationOrder))
                {
                    PowerPoint.Presentation presentation;
                    if (!_presentations.TryGetValue(key, out presentation))
                    {
                        continue;
                    }

                    string name;
                    string path;
                    bool dirty;
                    try
                    {
                        name = presentation.Name;
                        path = PresentationPath(presentation, name);
                        dirty = presentation.Saved == Office.MsoTriState.msoFalse;
                    }
                    catch (COMException)
                    {
                        continue;
                    }

                    _names[key] = name;
                    _paths[key] = path;
                    if (_coordinator.Tabs.Find(key) != null)
                    {
                        _coordinator.DocumentRenamed(key, path, name);
                        _coordinator.DocumentDirtyChanged(key, dirty);
                        continue;
                    }

                    PowerPoint.DocumentWindow window = null;
                    try
                    {
                        window = GetWindowObject(key);
                        if (window == null)
                        {
                            continue;
                        }

                        string caption = window.Caption;
                        IntPtr frame = FindUniqueFrame(caption, name, frames);
                        if (frame == IntPtr.Zero)
                        {
                            frame = FindVerifiedActiveFrame(key, frames);
                        }

                        if (frame == IntPtr.Zero)
                        {
                            string state = key + " caption='" + caption + "' name='" + name + "'";
                            if (_lastUnmatched != state)
                            {
                                _lastUnmatched = state;
                                Logger.Info("P2 waiting for a unique frame: " + state + ".");
                            }

                            continue;
                        }

                        _frames[key] = frame;
                        _coordinator.DocumentOpened(key, frame, delegate { return GetWindowObject(key); }, path,
                            name, dirty);
                        if (_coordinator.Tabs.Find(key) != null)
                        {
                            Logger.Info("P2 mapped " + key + " '" + name + "' to PPTFrameClass " + Hex(frame) + ".");
                            AlignTabOrder();
                            _lastUnmatched = string.Empty;
                        }
                        else
                        {
                            _frames.Remove(key);
                        }
                    }
                    finally
                    {
                        Release(window);
                    }
                }

                SetCurrentFromPowerPoint();
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointAdapter.RefreshWindows", ex);
            }
            finally
            {
                _refreshing = false;
            }
        }

        private void AlignTabOrder()
        {
            int destination = 0;
            foreach (string key in _creationOrder)
            {
                int current = _coordinator.Tabs.IndexOf(key);
                if (current < 0)
                {
                    continue;
                }

                if (current != destination)
                {
                    _coordinator.Tabs.Move(current, destination);
                }

                destination++;
            }
        }

        private static IntPtr FindUniqueFrame(string caption, string name, Dictionary<IntPtr, string> frames)
        {
            IntPtr found = IntPtr.Zero;
            foreach (KeyValuePair<IntPtr, string> frame in frames)
            {
                if (!Matches(frame.Value, caption) && !Matches(frame.Value, name))
                {
                    continue;
                }

                if (found != IntPtr.Zero)
                {
                    return IntPtr.Zero;
                }

                found = frame.Key;
            }

            return found;
        }

        // Only the foreground PowerPoint frame paired with the COM ActiveWindow is a valid title-free match.
        // Two consecutive polls are required, since Windows foreground and Office activation can transition
        // at slightly different times. This also permits two presentations to reuse one physical frame.
        private IntPtr FindVerifiedActiveFrame(string key, Dictionary<IntPtr, string> frames)
        {
            PowerPoint.DocumentWindow activeWindow = null;
            PowerPoint.Presentation activePresentation = null;
            try
            {
                IntPtr foreground = NativeMethods.GetForegroundWindow();
                IntPtr frame = NativeMethods.GetAncestor(foreground, NativeMethods.GA_ROOT);
                if (frame == IntPtr.Zero || !frames.ContainsKey(frame))
                {
                    ClearCandidateFor(key);
                    return IntPtr.Zero;
                }

                activeWindow = _app.ActiveWindow;
                if (activeWindow == null)
                {
                    ClearCandidateFor(key);
                    return IntPtr.Zero;
                }

                activePresentation = activeWindow.Presentation;
                if (FindKnown(activePresentation) != key)
                {
                    ClearCandidateFor(key);
                    return IntPtr.Zero;
                }

                int now = Environment.TickCount;
                if (_activeCandidateKey == key && _activeCandidateFrame == frame
                    && unchecked(now - _activeCandidateTick) <= ActiveCandidateMaxAgeMilliseconds)
                {
                    return frame;
                }

                _activeCandidateKey = key;
                _activeCandidateFrame = frame;
                _activeCandidateTick = now;
                return IntPtr.Zero;
            }
            catch (COMException)
            {
                ClearCandidateFor(key);
                return IntPtr.Zero;
            }
            finally
            {
                ReleaseTransientPresentation(activePresentation);
                Release(activeWindow);
            }
        }

        private void ClearCandidateFor(string key)
        {
            if (_activeCandidateKey == key)
            {
                _activeCandidateKey = null;
                _activeCandidateFrame = IntPtr.Zero;
            }
        }

        private Dictionary<IntPtr, string> CollectEditingFrames()
        {
            Dictionary<IntPtr, string> frames = new Dictionary<IntPtr, string>();
            uint processId;
            using (Process process = Process.GetCurrentProcess())
            {
                processId = (uint)process.Id;
            }

            NativeMethods.EnumWindowsProc callback = delegate(IntPtr hwnd, IntPtr unused)
            {
                try
                {
                    uint ownerId;
                    NativeMethods.GetWindowThreadProcessId(hwnd, out ownerId);
                    if (ownerId == processId && IsEditingFrame(hwnd))
                    {
                        frames[hwnd] = FrameTitle(hwnd);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error("PowerPointAdapter.CollectFrame", ex);
                }

                return true;
            };
            NativeMethods.EnumWindows(callback, IntPtr.Zero);
            GC.KeepAlive(callback);
            return frames;
        }

        private static bool IsEditingFrame(IntPtr hwnd)
        {
            if (!IsFrameAlive(hwnd) || !NativeMethods.IsWindowVisible(hwnd))
            {
                return false;
            }

            StringBuilder className = new StringBuilder(WindowClassCapacity);
            for (IntPtr child = NativeMethods.GetWindow(hwnd, NativeMethods.GW_CHILD);
                child != IntPtr.Zero; child = NativeMethods.GetWindow(child, NativeMethods.GW_HWNDNEXT))
            {
                className.Clear();
                if (NativeMethods.GetClassName(child, className, className.Capacity) > 0
                    && className.ToString() == DocumentAreaClass && NativeMethods.IsWindowVisible(child))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsFrameAlive(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || !NativeMethods.IsWindow(hwnd))
            {
                return false;
            }

            StringBuilder className = new StringBuilder(WindowClassCapacity);
            return NativeMethods.GetClassName(hwnd, className, className.Capacity) > 0
                && className.ToString() == FrameClass;
        }

        private static string FrameTitle(IntPtr frame)
        {
            if (!IsFrameAlive(frame))
            {
                return string.Empty;
            }

            StringBuilder title = new StringBuilder(WindowTitleCapacity);
            NativeMethods.GetWindowText(frame, title, title.Capacity);
            return title.ToString();
        }

        private static bool Matches(string title, string name)
        {
            if (string.IsNullOrEmpty(title) || string.IsNullOrEmpty(name))
            {
                return false;
            }

            if (string.Equals(title, name, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (title.Length > name.Length && title.StartsWith(name, StringComparison.OrdinalIgnoreCase))
            {
                char boundary = title[name.Length];
                return boundary == ' ' || boundary == '-' || boundary == '[' || boundary == '(';
            }

            return false;
        }

        private void SetCurrentFromPowerPoint()
        {
            PowerPoint.DocumentWindow window = null;
            PowerPoint.Presentation presentation = null;
            try
            {
                window = _app.ActiveWindow;
                if (window == null)
                {
                    return;
                }

                presentation = window.Presentation;
                string key = FindKnown(presentation);
                if (key != null && _coordinator.Tabs.Find(key) != null
                    && _coordinator.Tabs.CurrentKey != key)
                {
                    _coordinator.DocumentActivated(key);
                }
            }
            catch (COMException)
            {
                // PowerPoint has no active document window in Backstage or during some transitions.
            }
            finally
            {
                ReleaseTransientPresentation(presentation);
                Release(window);
            }
        }

        // ----- The COM Presentation owns its DocumentWindows, independent of any Win32 frame caption. -----

        private PowerPoint.DocumentWindow GetWindowObject(string key)
        {
            PowerPoint.Presentation presentation;
            if (!_presentations.TryGetValue(key, out presentation))
            {
                return null;
            }

            PowerPoint.DocumentWindows windows = null;
            PowerPoint.DocumentWindow window = null;
            PowerPoint.Presentation owner = null;
            bool keep = false;
            try
            {
                windows = presentation.Windows;
                if (windows == null || windows.Count < 1)
                {
                    return null;
                }

                window = windows[1];
                owner = window.Presentation;
                string name = presentation.Name;
                string path = PresentationPath(presentation, name);
                keep = string.Equals(owner.Name, name, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(PresentationPath(owner, owner.Name), path,
                        StringComparison.OrdinalIgnoreCase);
                return keep ? window : null;
            }
            finally
            {
                if (!keep)
                {
                    Release(window);
                }

                ReleaseTransientPresentation(owner);
                Release(windows);
            }
        }

        public void ActivateDocument(string key)
        {
            PowerPoint.DocumentWindow window = null;
            try
            {
                window = GetWindowObject(key);
                if (window != null)
                {
                    window.Activate();
                }
                else
                {
                    Logger.Info("P2 activation deferred for " + key + ": document window unavailable.");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointAdapter.ActivateDocument", ex);
            }
            finally
            {
                Release(window);
            }
        }

        // Presentation.Close never asks about unsaved changes, and PowerPoint's own Close command finishes after
        // the call that starts it, which stopped close-all after one tab. So the question is asked here, in
        // PowerPoint's words, and the presentation is saved and closed directly. The presentation's window is
        // brought to the front first, so the question isn't hidden behind another stacked PowerPoint window.
        public bool CloseDocument(string key)
        {
            PowerPoint.Presentation presentation;
            if (!_presentations.TryGetValue(key, out presentation))
            {
                return true;
            }

            IntPtr frame;
            _frames.TryGetValue(key, out frame);
            if (presentation.Saved == Office.MsoTriState.msoFalse)
            {
                ActivateDocument(key);
                if (frame != IntPtr.Zero)
                {
                    if (NativeMethods.IsIconic(frame))
                    {
                        NativeMethods.ShowWindow(frame, NativeMethods.SW_RESTORE);
                    }

                    NativeMethods.SetForegroundWindow(frame);
                }

                DialogResult answer = MessageBox.Show(new OwnerWindow(frame),
                    "Do you want to save changes to '" + presentation.Name + "'?", "Microsoft PowerPoint",
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (answer == DialogResult.Cancel)
                {
                    return false;
                }

                if (answer == DialogResult.Yes && !SavePresentation(presentation, frame))
                {
                    return false;
                }
            }

            _closingKey = key;
            try
            {
                presentation.Close();
            }
            finally
            {
                _closingKey = null;
            }

            Forget(key);
            return true;
        }

        // Saves in place, or asks for a file name when the presentation has never been saved or is read-only.
        // Returns false when the user cancels the Save As dialog.
        private static bool SavePresentation(PowerPoint.Presentation presentation, IntPtr frame)
        {
            if (!string.IsNullOrEmpty(presentation.Path) && presentation.ReadOnly != Office.MsoTriState.msoTrue)
            {
                presentation.Save();
                return true;
            }

            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Title = "Save As";
                dialog.FileName = presentation.Name;
                dialog.Filter = "PowerPoint Presentation (*.pptx)|*.pptx";
                dialog.DefaultExt = "pptx";
                dialog.AddExtension = true;
                dialog.OverwritePrompt = true;
                if (dialog.ShowDialog(new OwnerWindow(frame)) != DialogResult.OK)
                {
                    return false;
                }

                presentation.SaveAs(dialog.FileName, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                return true;
            }
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

        public void NewDocument()
        {
            PowerPoint.Presentations presentations = null;
            PowerPoint.Presentation created = null;
            try
            {
                presentations = _app.Presentations;
                created = presentations.Add(Office.MsoTriState.msoTrue);
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointAdapter.NewDocument", ex);
            }
            finally
            {
                ReleaseTransientPresentation(created);
                Release(presentations);
            }
        }

        // Called after close-all has closed every presentation, when the user chose to close PowerPoint as well.
        // If PowerPoint's main window is still there after Quit, it is closed the way the X button does.
        public void QuitApplication()
        {
            IntPtr main = new IntPtr(_app.HWND);
            _app.Quit();
            if (NativeMethods.IsWindow(main))
            {
                NativeMethods.PostMessage(main, NativeMethods.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            }
        }

        private void Forget(string key)
        {
            PowerPoint.Presentation presentation;
            if (!_presentations.TryGetValue(key, out presentation))
            {
                return;
            }

            if (_coordinator != null && _coordinator.Tabs.Find(key) != null)
            {
                _coordinator.DocumentClosed(key);
            }

            _presentations.Remove(key);
            _frames.Remove(key);
            _names.Remove(key);
            _paths.Remove(key);
            _creationOrder.Remove(key);
            Release(presentation);
            Logger.Info("P2 presentation removed: " + key + ".");
        }

        // A COM property may return the same RCW already retained for a live presentation. Releasing
        // that shared RCW would invalidate its tab's identity and object even while Office keeps it open.
        private void ReleaseTransientPresentation(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return;
            }

            foreach (PowerPoint.Presentation retained in _presentations.Values)
            {
                if (object.ReferenceEquals(retained, presentation))
                {
                    return;
                }
            }

            Release(presentation);
        }

        private static string Hex(IntPtr hwnd)
        {
            return "0x" + hwnd.ToInt64().ToString("X");
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
