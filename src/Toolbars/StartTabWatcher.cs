// ---------------------------------------------------------------------------------------------------------------
// Name:     StartTabWatcher.cs
// Purpose:  Selects the starting ribbon tab in each new Word, Excel or PowerPoint document window: the
//           OpenTools tab when Start on OpenTools tab is on, Home when it is off. Office would otherwise open
//           on the first tab, which is OpenTools because it sits before Home. A Word or Excel window is handled
//           only the first time it is activated, so a tab the user picks later is left alone. PowerPoint is
//           handled when a presentation is created or opened, because each presentation has its own window.
//           The tab is selected at 150 ms, 0.5 s, 1 s and 2 s after that, because the new window's ribbon may
//           set Home after the first try. The same is done once when the app starts. The setting is read each
//           time, so a change takes effect with the next new window. No exception escapes into Office.
// Revision: 6
// DTG:      261025020000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace OpenTools
{
    internal sealed class StartTabWatcher
    {
        private static readonly int[] DelaysMilliseconds = { 150, 500, 1000, 2000 };

        private const string HomeTabId = "TabHome";

        private readonly Office.IRibbonUI _ribbon;
        private readonly string _appName;
        private readonly HashSet<int> _seen = new HashSet<int>();
        private readonly Timer _timer = new Timer();
        private Word.ApplicationEvents4_Event _wordEvents;
        private Excel.AppEvents_Event _excelEvents;
        private PowerPoint.EApplication_Event _powerPointEvents;
        private int _attempt;

        // application is Word's, Excel's or PowerPoint's Application object. appName is its settings name.
        public StartTabWatcher(object application, Office.IRibbonUI ribbon, string appName)
        {
            _ribbon = ribbon;
            _appName = appName;
            _timer.Interval = DelaysMilliseconds[0];
            _timer.Tick += OnTick;
            Word.Application word = application as Word.Application;
            if (word != null)
            {
                _wordEvents = (Word.ApplicationEvents4_Event)word;
                _wordEvents.WindowActivate += OnWordWindowActivate;
                return;
            }

            Excel.Application excel = application as Excel.Application;
            if (excel != null)
            {
                _excelEvents = (Excel.AppEvents_Event)excel;
                _excelEvents.WindowActivate += OnExcelWindowActivate;
                return;
            }

            PowerPoint.Application powerPoint = application as PowerPoint.Application;
            if (powerPoint != null)
            {
                _powerPointEvents = (PowerPoint.EApplication_Event)powerPoint;
                _powerPointEvents.NewPresentation += OnPresentationAdded;
                _powerPointEvents.AfterPresentationOpen += OnPresentationAdded;
            }
        }

        // ----- Events -----

        private void OnWordWindowActivate(Word.Document doc, Word.Window window)
        {
            try
            {
                WindowActivated(window.Hwnd);
            }
            catch (Exception ex)
            {
                Logger.Error("StartTabWatcher.OnWordWindowActivate", ex);
            }
        }

        private void OnExcelWindowActivate(Excel.Workbook workbook, Excel.Window window)
        {
            try
            {
                WindowActivated(window.Hwnd);
            }
            catch (Exception ex)
            {
                Logger.Error("StartTabWatcher.OnExcelWindowActivate", ex);
            }
        }

        private void OnPresentationAdded(PowerPoint.Presentation presentation)
        {
            try
            {
                StartSelecting();
            }
            catch (Exception ex)
            {
                Logger.Error("StartTabWatcher.OnPresentationAdded", ex);
            }
        }

        // Called from the event handlers above, inside their try blocks.
        private void WindowActivated(int hwnd)
        {
            if (_seen.Add(hwnd))
            {
                StartSelecting();
            }
        }

        // Called once by Connect when the ribbon is built, so the first window gets the same treatment.
        public void SelectAtStart()
        {
            StartSelecting();
        }

        private void StartSelecting()
        {
            _attempt = 0;
            _timer.Stop();
            _timer.Interval = DelaysMilliseconds[0];
            _timer.Start();
        }

        private void OnTick(object sender, EventArgs e)
        {
            try
            {
                _timer.Stop();
                SelectStartTab(_ribbon, new Settings(_appName).GetBool(SettingNames.StartOnOpenToolsTab));
                _attempt++;
                if (_attempt < DelaysMilliseconds.Length)
                {
                    _timer.Interval = DelaysMilliseconds[_attempt] - DelaysMilliseconds[_attempt - 1];
                    _timer.Start();
                }
            }
            catch (Exception ex)
            {
                Logger.Error("StartTabWatcher.OnTick", ex);
            }
        }

        private static void SelectStartTab(Office.IRibbonUI ribbon, bool openTools)
        {
            if (openTools)
            {
                ribbon.ActivateTab(ToolbarRibbonXml.TabId);
            }
            else
            {
                ribbon.ActivateTabMso(HomeTabId);
            }
        }

        // ----- Stop -----

        public void Stop()
        {
            try
            {
                _timer.Stop();
                _timer.Dispose();
                if (_wordEvents != null)
                {
                    _wordEvents.WindowActivate -= OnWordWindowActivate;
                    _wordEvents = null;
                }

                if (_excelEvents != null)
                {
                    _excelEvents.WindowActivate -= OnExcelWindowActivate;
                    _excelEvents = null;
                }

                if (_powerPointEvents != null)
                {
                    _powerPointEvents.NewPresentation -= OnPresentationAdded;
                    _powerPointEvents.AfterPresentationOpen -= OnPresentationAdded;
                    _powerPointEvents = null;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("StartTabWatcher.Stop", ex);
            }
        }
    }
}
