// ---------------------------------------------------------------------------------------------------------------
// Name:     ExcelAdapter.cs
// Purpose:  The Excel side of the add-in. Listens for Excel's new, open, activate, close and window-resize events
//           and reports them to TabCoordinator, in the same way WordAdapter does for Word. Excel runs one main
//           window (class XLMAIN) per workbook, and the handle of the workbook window is that main window, so
//           the handle is also the key of the tab. A workbook with no visible window gets no tab. Excel has no
//           event for "unsaved changes" or "saved under a new name", so a half-second timer checks both. A short
//           timer also confirms a workbook is really gone after a close, because the user can still cancel the
//           save prompt. The main window of a saved workbook is cloaked as it closes, so the close does not
//           flash. Each Excel object is fetched into its own variable and released. An event brake logs the
//           first events of each kind and ignores window events for a few seconds if they arrive in a flood.
// Revision: 4
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OpenTools
{
    internal sealed class ExcelAdapter : IAppAdapter
    {
        private const int PollMilliseconds = 500;
        private const int CloseCheckMilliseconds = 300;
        private const int BrakeLimit = 40;
        private const int BrakeWindowMilliseconds = 2000;
        private const int BrakePauseMilliseconds = 5000;
        private const int EventsLoggedPerKind = 20;
        private const int CloakHoldMilliseconds = 2000;

        private readonly Excel.Application _app;
        private readonly Dictionary<string, Excel.Workbook> _books = new Dictionary<string, Excel.Workbook>();
        private readonly Dictionary<string, IntPtr> _hwnds = new Dictionary<string, IntPtr>();
        private readonly Dictionary<string, int> _eventCounts = new Dictionary<string, int>();
        private readonly Queue<int> _recentEvents = new Queue<int>();
        private readonly Dictionary<IntPtr, int> _cloakedFrames = new Dictionary<IntPtr, int>();
        private TabCoordinator _coordinator;
        private Excel.AppEvents_Event _events;
        private Timer _pollTimer;
        private Timer _closeTimer;
        private bool _braked;
        private int _brakedAt;

        public ExcelAdapter(Excel.Application app)
        {
            _app = app;
        }

        public string AppName
        {
            get { return "Excel"; }
        }

        // ----- Start and stop -----

        public void Start(TabCoordinator coordinator)
        {
            _coordinator = coordinator;

            _events = (Excel.AppEvents_Event)_app;
            _events.NewWorkbook += OnNewWorkbook;
            _events.WorkbookOpen += OnWorkbookOpen;
            _events.WorkbookBeforeClose += OnWorkbookBeforeClose;
            _events.WindowActivate += OnWindowActivate;
            _events.WindowResize += OnWindowResize;

            _pollTimer = new Timer();
            _pollTimer.Interval = PollMilliseconds;
            _pollTimer.Tick += OnPollTick;

            _closeTimer = new Timer();
            _closeTimer.Interval = CloseCheckMilliseconds;
            _closeTimer.Tick += OnCloseTick;

            SeedOpenWorkbooks();
            _pollTimer.Start();
            Logger.Info("Excel adapter started.");
        }

        public void Stop()
        {
            try
            {
                if (_pollTimer != null)
                {
                    _pollTimer.Stop();
                    _pollTimer.Dispose();
                }

                if (_closeTimer != null)
                {
                    _closeTimer.Stop();
                    _closeTimer.Dispose();
                }

                if (_events != null)
                {
                    _events.NewWorkbook -= OnNewWorkbook;
                    _events.WorkbookOpen -= OnWorkbookOpen;
                    _events.WorkbookBeforeClose -= OnWorkbookBeforeClose;
                    _events.WindowActivate -= OnWindowActivate;
                    _events.WindowResize -= OnWindowResize;
                }

                foreach (Excel.Workbook book in _books.Values)
                {
                    ReleaseFully(book);
                }

                _books.Clear();
                _hwnds.Clear();
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelAdapter.Stop", ex);
            }
        }

        // ----- Event brake -----

        // Returns true when the event should be ignored. The first events of each kind are logged, so a
        // runaway loop shows in the log. If more than BrakeLimit events arrive inside BrakeWindowMilliseconds,
        // events are ignored for BrakePauseMilliseconds and one line is logged.
        private bool TooManyEvents(string kind)
        {
            int now = Environment.TickCount;
            if (_braked)
            {
                if (now - _brakedAt < BrakePauseMilliseconds)
                {
                    return true;
                }

                _braked = false;
                _recentEvents.Clear();
                Logger.Info("Event brake released.");
            }

            int seen;
            _eventCounts.TryGetValue(kind, out seen);
            seen++;
            _eventCounts[kind] = seen;
            if (seen <= EventsLoggedPerKind)
            {
                Logger.Info("Event " + kind + " number " + seen + ".");
            }

            _recentEvents.Enqueue(now);
            while (_recentEvents.Count > 0 && now - _recentEvents.Peek() > BrakeWindowMilliseconds)
            {
                _recentEvents.Dequeue();
            }

            if (_recentEvents.Count > BrakeLimit)
            {
                _braked = true;
                _brakedAt = now;
                Logger.Info("Event brake on: more than " + BrakeLimit + " window events in "
                    + BrakeWindowMilliseconds + " ms. Last kind: " + kind + ". Ignoring window events for "
                    + BrakePauseMilliseconds + " ms.");
                return true;
            }

            return false;
        }

        // ----- Excel events -----

        private void OnNewWorkbook(Excel.Workbook book)
        {
            try
            {
                Register(book);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelAdapter.OnNewWorkbook", ex);
            }
        }

        private void OnWorkbookOpen(Excel.Workbook book)
        {
            try
            {
                Register(book);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelAdapter.OnWorkbookOpen", ex);
            }
        }

        private void OnWorkbookBeforeClose(Excel.Workbook book, ref bool cancel)
        {
            try
            {
                if (OpenFrameCount() > 1)
                {
                    CloakIfSaved(book);
                }

                RestartCloseTimer();
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelAdapter.OnWorkbookBeforeClose", ex);
            }
        }

        private void OnWindowActivate(Excel.Workbook book, Excel.Window window)
        {
            try
            {
                if (TooManyEvents("WindowActivate"))
                {
                    return;
                }

                string key = KeyFor(window.Hwnd);
                if (!_books.ContainsKey(key))
                {
                    Register(book);
                }

                _coordinator.DocumentActivated(key);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelAdapter.OnWindowActivate", ex);
            }
        }

        private void OnWindowResize(Excel.Workbook book, Excel.Window window)
        {
            try
            {
                if (TooManyEvents("WindowResize"))
                {
                    return;
                }

                _coordinator.WindowMoved(new IntPtr(window.Hwnd));
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelAdapter.OnWindowResize", ex);
            }
        }

        // ----- Registering workbooks -----

        private void SeedOpenWorkbooks()
        {
            Excel.Workbooks books = _app.Workbooks;
            try
            {
                int count = books.Count;
                for (int i = 1; i <= count; i++)
                {
                    Excel.Workbook book = books[i];
                    if (!Register(book))
                    {
                        Marshal.ReleaseComObject(book);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(books);
            }
        }

        // Returns true when the workbook was added and is now held by this adapter.
        private bool Register(Excel.Workbook book)
        {
            Excel.Windows windows = null;
            Excel.Window window = null;
            try
            {
                windows = book.Windows;
                if (windows.Count < 1)
                {
                    return false;
                }

                window = windows[1];
                if (!window.Visible)
                {
                    return false;
                }

                int hwnd = window.Hwnd;
                string key = KeyFor(hwnd);
                if (_books.ContainsKey(key))
                {
                    return false;
                }

                _books[key] = book;
                _hwnds[key] = new IntPtr(hwnd);
                _coordinator.DocumentOpened(key, new IntPtr(hwnd), delegate { return GetWindowObject(key); },
                    book.FullName, book.Name, !book.Saved);
                return true;
            }
            catch (COMException)
            {
                // A hidden workbook has no window, so it gets no tab. This is checked every poll, so it isn't logged.
                return false;
            }
            finally
            {
                Release(window);
                Release(windows);
            }
        }

        private object GetWindowObject(string key)
        {
            Excel.Workbook book;
            if (!_books.TryGetValue(key, out book))
            {
                return null;
            }

            Excel.Windows windows = book.Windows;
            try
            {
                return windows[1];
            }
            finally
            {
                Marshal.ReleaseComObject(windows);
            }
        }

        private static string KeyFor(int hwnd)
        {
            return "X" + hwnd.ToString("X");
        }

        // ----- Requests from the tab bar -----

        public void ActivateDocument(string key)
        {
            Excel.Workbook book;
            if (!_books.TryGetValue(key, out book))
            {
                return;
            }

            Excel.Windows windows = book.Windows;
            Excel.Window window = null;
            try
            {
                window = windows[1];
                window.Activate();
            }
            finally
            {
                Release(window);
                Marshal.ReleaseComObject(windows);
            }
        }

        public bool CloseDocument(string key)
        {
            Excel.Workbook book;
            if (!_books.TryGetValue(key, out book))
            {
                return true;
            }

            Excel._Workbook closable = book;
            closable.Close(Type.Missing, Type.Missing, Type.Missing);
            RestartCloseTimer();

            // A workbook that is still open after Close was kept by the user, who cancelled the save prompt.
            return IsWorkbookGone(key);
        }

        public void NewDocument()
        {
            Excel.Workbooks books = _app.Workbooks;
            try
            {
                Excel.Workbook created = books.Add(Type.Missing);
                Marshal.ReleaseComObject(created);
            }
            finally
            {
                Marshal.ReleaseComObject(books);
            }
        }

        public void QuitApplication()
        {
            Excel._Application app = _app;
            app.Quit();
        }

        // ----- Closing -----

        // Counts the main windows that are alive and not already cloaked for a close. A close-all closes the
        // workbooks quickly, so the list still holds ones that were just closed.
        private int OpenFrameCount()
        {
            int count = 0;
            foreach (KeyValuePair<string, IntPtr> pair in _hwnds)
            {
                IntPtr frame = NativeMethods.GetAncestor(pair.Value, NativeMethods.GA_ROOT);
                if (NativeMethods.IsWindow(pair.Value) && !_cloakedFrames.ContainsKey(frame))
                {
                    count++;
                }
            }

            return count;
        }

        // A saved workbook closes without a prompt, so its main window can be hidden at once.
        private void CloakIfSaved(Excel.Workbook book)
        {
            Excel.Windows windows = null;
            Excel.Window window = null;
            try
            {
                if (!book.Saved)
                {
                    return;
                }

                windows = book.Windows;
                window = windows[1];
                IntPtr frame = NativeMethods.GetAncestor(new IntPtr(window.Hwnd), NativeMethods.GA_ROOT);
                if (frame != IntPtr.Zero && !_cloakedFrames.ContainsKey(frame))
                {
                    _cloakedFrames[frame] = Environment.TickCount;
                    FrameAnimation.Cloak(frame, true);
                }
            }
            catch (COMException)
            {
                // The workbook has no window, so there is nothing to cloak.
            }
            finally
            {
                Release(window);
                Release(windows);
            }
        }

        // A main window that is still there after the hold time was not closed, so it is shown again.
        private void UncloakSurvivors()
        {
            foreach (KeyValuePair<IntPtr, int> pair in new List<KeyValuePair<IntPtr, int>>(_cloakedFrames))
            {
                if (!NativeMethods.IsWindow(pair.Key))
                {
                    _cloakedFrames.Remove(pair.Key);
                }
                else if (Environment.TickCount - pair.Value > CloakHoldMilliseconds)
                {
                    _cloakedFrames.Remove(pair.Key);
                    FrameAnimation.Cloak(pair.Key, false);
                }
            }
        }

        // ----- Timers -----

        private void RestartCloseTimer()
        {
            _closeTimer.Stop();
            _closeTimer.Start();
        }

        private void OnCloseTick(object sender, EventArgs e)
        {
            _closeTimer.Stop();
            try
            {
                RemoveGoneWorkbooks();
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelAdapter.OnCloseTick", ex);
            }
        }

        private void OnPollTick(object sender, EventArgs e)
        {
            try
            {
                RemoveGoneWorkbooks();
                UncloakSurvivors();
                SeedOpenWorkbooks();
                CheckSavedStateAndNames();
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelAdapter.OnPollTick", ex);
            }
        }

        private void RemoveGoneWorkbooks()
        {
            foreach (string key in new List<string>(_books.Keys))
            {
                if (IsWorkbookGone(key))
                {
                    Excel.Workbook book = _books[key];
                    _books.Remove(key);
                    _hwnds.Remove(key);
                    _coordinator.DocumentClosed(key);
                    ReleaseFully(book);
                }
            }
        }

        // A workbook is gone when its main window is gone, or when Excel no longer answers for it, or when its
        // window is no longer the one the tab was made for. The main window can stay open with no workbook.
        private bool IsWorkbookGone(string key)
        {
            Excel.Workbook book = _books[key];
            Excel.Windows windows = null;
            Excel.Window window = null;
            try
            {
                if (!NativeMethods.IsWindow(_hwnds[key]))
                {
                    return true;
                }

                windows = book.Windows;
                if (windows.Count < 1)
                {
                    return true;
                }

                window = windows[1];
                return window.Hwnd != _hwnds[key].ToInt32();
            }
            catch (COMException)
            {
                return true;
            }
            finally
            {
                Release(window);
                Release(windows);
            }
        }

        private void CheckSavedStateAndNames()
        {
            List<KeyValuePair<string, Excel.Workbook>> snapshot =
                new List<KeyValuePair<string, Excel.Workbook>>(_books);
            foreach (KeyValuePair<string, Excel.Workbook> pair in snapshot)
            {
                try
                {
                    _coordinator.DocumentDirtyChanged(pair.Key, !pair.Value.Saved);
                    _coordinator.DocumentRenamed(pair.Key, pair.Value.FullName, pair.Value.Name);
                }
                catch (COMException)
                {
                    // The workbook is closing. The next close check removes it.
                }
            }
        }

        // ----- Releasing -----

        private static void Release(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }

        private static void ReleaseFully(object comObject)
        {
            try
            {
                if (comObject != null && Marshal.IsComObject(comObject))
                {
                    Marshal.FinalReleaseComObject(comObject);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelAdapter.ReleaseFully", ex);
            }
        }
    }
}
