// ---------------------------------------------------------------------------------------------------------------
// Name:     MenuCaptions.cs
// Purpose:  Reads a menu heading from Office's hidden Office 2003 menu bar, so the OpenTools menus show Office's
//           own word for File, Edit and so on in the user's language. The caption carries an "&" before its Alt
//           letter, for example "&File". Label removes the "&"; Keytip returns the letter after it. When the
//           caption can't be read, the fallback caption is used. Office objects are released in finally blocks
//           and nothing is thrown.
// Revision: 1
// DTG:      261024210000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace OpenTools
{
    internal static class MenuCaptions
    {
        private const char AltMarker = '&';

        // ----- Reading -----

        public static string Read(object application, int captionId, string fallback)
        {
            if (application == null || captionId <= 0)
            {
                return fallback;
            }

            Office.CommandBars bars = null;
            Office.CommandBarControl control = null;
            try
            {
                bars = BarsOf(application);
                if (bars == null)
                {
                    return fallback;
                }

                control = bars.FindControl(Type.Missing, captionId, Type.Missing, Type.Missing);
                if (control == null)
                {
                    return fallback;
                }

                string caption = control.Caption;
                return string.IsNullOrEmpty(caption) ? fallback : caption;
            }
            catch (Exception ex)
            {
                Logger.Error("MenuCaptions.Read " + captionId, ex);
                return fallback;
            }
            finally
            {
                Release(control);
                Release(bars);
            }
        }

        private static Office.CommandBars BarsOf(object application)
        {
            Word.Application word = application as Word.Application;
            if (word != null)
            {
                return word.CommandBars;
            }

            Excel.Application excel = application as Excel.Application;
            if (excel != null)
            {
                return excel.CommandBars;
            }

            PowerPoint.Application powerPoint = application as PowerPoint.Application;
            if (powerPoint != null)
            {
                return powerPoint.CommandBars;
            }

            return null;
        }

        // ----- Caption parts -----

        // "&&" is a real ampersand. A single "&" marks the Alt letter and is dropped.
        public static string Label(string caption)
        {
            StringBuilder label = new StringBuilder();
            for (int i = 0; i < caption.Length; i++)
            {
                if (caption[i] == AltMarker)
                {
                    if (i + 1 < caption.Length && caption[i + 1] == AltMarker)
                    {
                        label.Append(AltMarker);
                        i++;
                    }

                    continue;
                }

                label.Append(caption[i]);
            }

            return label.ToString();
        }

        public static string Keytip(string caption)
        {
            for (int i = 0; i < caption.Length - 1; i++)
            {
                if (caption[i] != AltMarker)
                {
                    continue;
                }

                if (caption[i + 1] == AltMarker)
                {
                    i++;
                    continue;
                }

                return char.ToUpperInvariant(caption[i + 1]).ToString();
            }

            return string.Empty;
        }

        private static void Release(object comObject)
        {
            try
            {
                if (comObject != null && Marshal.IsComObject(comObject))
                {
                    Marshal.ReleaseComObject(comObject);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("MenuCaptions.Release", ex);
            }
        }
    }
}
