// ---------------------------------------------------------------------------------------------------------------
// Name:     WordAdapter.cs
// Purpose:  The Word side of the add-in. Listens for Word's open, new, activate, close and window-size events
//           and reports them to TabCoordinator. Word has no event for "unsaved changes" or "saved under a new
//           name", so a half-second timer checks both. A short timer also confirms a document is really gone
//           after a close, because the user can still cancel the save prompt. Each Word object is fetched into
//           its own variable. When a saved document starts to close, its frame is cloaked so the close does not
//           flash. The last open document is not cloaked, because Word keeps its frame. The frame is uncloaked if it
//           is still there two seconds later (the close was cancelled).
//           Word objects that are no longer needed are released. An event brake logs the
//           first events of each kind and ignores window events for a few seconds if they arrive in a flood.
// Revision: 10
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace OpenTools
{
    internal sealed class WordAdapter : IAppAdapter
    {
        private const int PollMilliseconds = 500;
        private const int CloseCheckMilliseconds = 300;
        private const int BrakeLimit = 40;
        private const int BrakeWindowMilliseconds = 2000;
        private const int BrakePauseMilliseconds = 5000;
        private const int EventsLoggedPerKind = 20;
        private const int CloakHoldMilliseconds = 2000;

        private readonly Word.Application _app;
        private readonly Dictionary<string, Word.Document> _docs = new Dictionary<string, Word.Document>();
        private readonly Dictionary<string, IntPtr> _hwnds = new Dictionary<string, IntPtr>();
        private readonly Dictionary<string, int> _eventCounts = new Dictionary<string, int>();
        private readonly Queue<int> _recentEvents = new Queue<int>();
        private readonly Dictionary<IntPtr, int> _cloakedFrames = new Dictionary<IntPtr, int>();
        private TabCoordinator _coordinator;
        private Word.ApplicationEvents4_Event _events;
        private Timer _pollTimer;
        private Timer _closeTimer;
        private bool _braked;
        private int _brakedAt;

        public WordAdapter(Word.Application app)
        {
            _app = app;
        }

        public string AppName
        {
            get { return "Word"; }
        }

        // ----- Start and stop -----

        public void Start(TabCoordinator coordinator)
        {
            _coordinator = coordinator;

            _events = (Word.ApplicationEvents4_Event)_app;
            _events.DocumentOpen += OnDocumentOpen;
            _events.NewDocument += OnNewDocument;
            _events.DocumentBeforeClose += OnDocumentBeforeClose;
            _events.WindowActivate += OnWindowActivate;
            _events.WindowSize += OnWindowSize;

            _pollTimer = new Timer();
            _pollTimer.Interval = PollMilliseconds;
            _pollTimer.Tick += OnPollTick;

            _closeTimer = new Timer();
            _closeTimer.Interval = CloseCheckMilliseconds;
            _closeTimer.Tick += OnCloseTick;

            SeedOpenDocuments();
            _pollTimer.Start();
            Logger.Info("Word adapter started.");
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
                    _events.DocumentOpen -= OnDocumentOpen;
                    _events.NewDocument -= OnNewDocument;
                    _events.DocumentBeforeClose -= OnDocumentBeforeClose;
                    _events.WindowActivate -= OnWindowActivate;
                    _events.WindowSize -= OnWindowSize;
                }

                foreach (Word.Document doc in _docs.Values)
                {
                    ReleaseFully(doc);
                }

                _docs.Clear();
                _hwnds.Clear();
            }
            catch (Exception ex)
            {
                Logger.Error("WordAdapter.Stop", ex);
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

        // ----- Word events -----

        private void OnDocumentOpen(Word.Document doc)
        {
            try
            {
                Register(doc);
            }
            catch (Exception ex)
            {
                Logger.Error("WordAdapter.OnDocumentOpen", ex);
            }
        }

        private void OnNewDocument(Word.Document doc)
        {
            try
            {
                Register(doc);
            }
            catch (Exception ex)
            {
                Logger.Error("WordAdapter.OnNewDocument", ex);
            }
        }

        private void OnDocumentBeforeClose(Word.Document doc, ref bool cancel)
        {
            try
            {
                if (OpenFrameCount() > 1)
                {
                    CloakIfSaved(doc);
                }

                RestartCloseTimer();
            }
            catch (Exception ex)
            {
                Logger.Error("WordAdapter.OnDocumentBeforeClose", ex);
            }
        }

        private void OnWindowActivate(Word.Document doc, Word.Window window)
        {
            try
            {
                if (TooManyEvents("WindowActivate"))
                {
                    return;
                }

                string key = KeyFor(window.Hwnd);
                if (!_docs.ContainsKey(key))
                {
                    Register(doc);
                }

                _coordinator.DocumentActivated(key);
            }
            catch (Exception ex)
            {
                Logger.Error("WordAdapter.OnWindowActivate", ex);
            }
        }

        private void OnWindowSize(Word.Document doc, Word.Window window)
        {
            try
            {
                if (TooManyEvents("WindowSize"))
                {
                    return;
                }

                _coordinator.WindowMoved(new IntPtr(window.Hwnd));
            }
            catch (Exception ex)
            {
                Logger.Error("WordAdapter.OnWindowSize", ex);
            }
        }

        // ----- Registering documents -----

        private void SeedOpenDocuments()
        {
            Word.Documents docs = _app.Documents;
            try
            {
                int count = docs.Count;
                for (int i = 1; i <= count; i++)
                {
                    Word.Document doc = docs[i];
                    if (!Register(doc))
                    {
                        Marshal.ReleaseComObject(doc);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(docs);
            }
        }

        // Returns true when the document was added and is now held by this adapter.
        private bool Register(Word.Document doc)
        {
            Word.Window window = null;
            try
            {
                window = doc.ActiveWindow;
                int hwnd = window.Hwnd;
                string key = KeyFor(hwnd);
                if (_docs.ContainsKey(key))
                {
                    return false;
                }

                _docs[key] = doc;
                _hwnds[key] = new IntPtr(hwnd);
                _coordinator.DocumentOpened(key, new IntPtr(hwnd), delegate { return GetWindowObject(key); },
                    doc.FullName, doc.Name, !doc.Saved);
                return true;
            }
            catch (COMException)
            {
                // A hidden document has no window, so it gets no tab. This is checked every poll, so it isn't logged.
                return false;
            }
            finally
            {
                if (window != null)
                {
                    Marshal.ReleaseComObject(window);
                }
            }
        }

        private object GetWindowObject(string key)
        {
            Word.Document doc;
            if (!_docs.TryGetValue(key, out doc))
            {
                return null;
            }

            return doc.ActiveWindow;
        }

        private static string KeyFor(int hwnd)
        {
            return "W" + hwnd.ToString("X");
        }

        // ----- Requests from the tab bar -----

        public void ActivateDocument(string key)
        {
            Word.Document doc;
            if (!_docs.TryGetValue(key, out doc))
            {
                return;
            }

            Word.Window window = doc.ActiveWindow;
            try
            {
                window.Activate();
            }
            finally
            {
                Marshal.ReleaseComObject(window);
            }
        }

        public bool CloseDocument(string key)
        {
            Word.Document doc;
            if (!_docs.TryGetValue(key, out doc))
            {
                return true;
            }

            // Document has both a Close method and a Close event, so the call goes through _Document.
            Word._Document closable = doc;
            closable.Close();
            RestartCloseTimer();

            // A document that is still open after Close was kept by the user, who cancelled the save prompt.
            try
            {
                string stillThere = doc.Name;
                return string.IsNullOrEmpty(stillThere);
            }
            catch (COMException)
            {
                return true;
            }
        }

        public void NewDocument()
        {
            Word.Documents docs = _app.Documents;
            try
            {
                Word.Document created = docs.Add();
                Marshal.ReleaseComObject(created);
            }
            finally
            {
                Marshal.ReleaseComObject(docs);
            }
        }

        public void QuitApplication()
        {
            // Application has both a Quit method and a Quit event, so the call goes through _Application.
            Word._Application app = _app;
            app.Quit();
        }

        // ----- Timers -----

        // Counts the main windows that are alive and not already cloaked for a close. A close-all closes the
        // documents quickly, so the document list still holds ones that were just closed.
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

        // A saved document closes without a prompt, so its frame can be hidden at once. Only the frame of the
        // document that is closing is cloaked.
        private void CloakIfSaved(Word.Document doc)
        {
            Word.Window window = null;
            try
            {
                if (!doc.Saved)
                {
                    return;
                }

                window = doc.ActiveWindow;
                IntPtr frame = NativeMethods.GetAncestor(new IntPtr(window.Hwnd), NativeMethods.GA_ROOT);
                if (frame != IntPtr.Zero && !_cloakedFrames.ContainsKey(frame))
                {
                    _cloakedFrames[frame] = Environment.TickCount;
                    FrameAnimation.Cloak(frame, true);
                }
            }
            catch (COMException)
            {
                // The document has no window, so there is nothing to cloak.
            }
            finally
            {
                if (window != null)
                {
                    Marshal.ReleaseComObject(window);
                }
            }
        }

        // A frame that is still there after the hold time was not closed, so it is shown again.
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
                RemoveGoneDocuments();
            }
            catch (Exception ex)
            {
                Logger.Error("WordAdapter.OnCloseTick", ex);
            }
        }

        private void OnPollTick(object sender, EventArgs e)
        {
            try
            {
                RemoveGoneDocuments();
                UncloakSurvivors();
                SeedOpenDocuments();
                CheckSavedStateAndNames();
            }
            catch (Exception ex)
            {
                Logger.Error("WordAdapter.OnPollTick", ex);
            }
        }

        private void RemoveGoneDocuments()
        {
            foreach (string key in new List<string>(_docs.Keys))
            {
                if (IsDocumentGone(key))
                {
                    Word.Document doc = _docs[key];
                    _docs.Remove(key);
                    _hwnds.Remove(key);
                    _coordinator.DocumentClosed(key);
                    ReleaseFully(doc);
                }
            }
        }

        private bool IsDocumentGone(string key)
        {
            Word.Document doc = _docs[key];
            Word.Window window = null;
            try
            {
                window = doc.ActiveWindow;
                return !NativeMethods.IsWindow(_hwnds[key]);
            }
            catch (COMException)
            {
                return true;
            }
            finally
            {
                if (window != null)
                {
                    Marshal.ReleaseComObject(window);
                }
            }
        }

        private void CheckSavedStateAndNames()
        {
            foreach (KeyValuePair<string, Word.Document> pair in new List<KeyValuePair<string, Word.Document>>(_docs))
            {
                try
                {
                    _coordinator.DocumentDirtyChanged(pair.Key, !pair.Value.Saved);
                    _coordinator.DocumentRenamed(pair.Key, pair.Value.FullName, pair.Value.Name);
                }
                catch (COMException)
                {
                    // The document is closing. The next close check removes it.
                }
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
                Logger.Error("WordAdapter.ReleaseFully", ex);
            }
        }
    }
}
