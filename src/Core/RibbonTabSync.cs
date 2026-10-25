// ---------------------------------------------------------------------------------------------------------------
// Name:     RibbonTabSync.cs
// Purpose:  Keeps the selected ribbon tab the same when the user switches between documents that have their own
//           main windows (Word, Excel, PowerPoint). Office has no call that reports the selected ribbon tab, so
//           the tab is read through Windows accessibility (MSAA): the ribbon's NetUIHWND window exposes its tabs
//           as page tabs, the selected one carrying the "selected" state. The tab with the same name in the
//           other window is then selected with its default action, the same thing a screen reader does. A
//           collapsed ribbon has no selected tab, and a tab that does not exist in the other window (a
//           contextual tab such as Table Design) is skipped, so nothing changes in those cases. The calls run
//           inside Office, on its own thread. Nothing here changes the ribbon's contents.
// Revision: 2
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace OpenTools
{
    internal static class RibbonTabSync
    {
        private const int WindowDepthLimit = 8;
        private const int NodeDepthLimit = 12;
        private const int NodeLimit = 3000;
        private const int ClassCapacity = 128;
        private const int SelfChild = 0;

        private sealed class RibbonTab
        {
            public IAccessible Container;
            public object Child;
            public string Name;
            public bool Selected;
        }

        // Selects, in toFrame's ribbon, the tab that is selected in fromFrame's ribbon.
        public static void Copy(IntPtr fromFrame, IntPtr toFrame)
        {
            if (fromFrame == IntPtr.Zero || toFrame == IntPtr.Zero || fromFrame == toFrame)
            {
                return;
            }

            List<object> held = new List<object>();
            try
            {
                string wanted = null;
                foreach (RibbonTab tab in ReadTabs(fromFrame, held))
                {
                    if (tab.Selected)
                    {
                        wanted = tab.Name;
                        break;
                    }
                }

                if (string.IsNullOrEmpty(wanted))
                {
                    return;
                }

                List<RibbonTab> targets = ReadTabs(toFrame, held);
                bool anySelected = false;
                RibbonTab match = null;
                foreach (RibbonTab tab in targets)
                {
                    anySelected |= tab.Selected;
                    if (tab.Name == wanted)
                    {
                        match = tab;
                    }
                }

                // A collapsed ribbon shows no selected tab. Selecting one there would drop the ribbon open over the
                // document, so the window is left alone.
                if (match == null || match.Selected || !anySelected)
                {
                    return;
                }

                match.Container.accDoDefaultAction(match.Child);
            }
            catch (Exception ex)
            {
                Logger.Error("RibbonTabSync.Copy", ex);
            }
            finally
            {
                foreach (object value in held)
                {
                    try
                    {
                        if (value != null && Marshal.IsComObject(value))
                        {
                            Marshal.ReleaseComObject(value);
                        }
                    }
                    catch (Exception)
                    {
                        // Already released through another reference. Nothing to do.
                    }
                }
            }
        }

        // The ribbon's page tabs in one main window, or an empty list.
        private static List<RibbonTab> ReadTabs(IntPtr frame, List<object> held)
        {
            List<RibbonTab> tabs = new List<RibbonTab>();
            IntPtr ribbonBar = FindChild(frame, "MsoCommandBar", "Ribbon", 0);
            if (ribbonBar == IntPtr.Zero)
            {
                return tabs;
            }

            IntPtr surface = FindChild(ribbonBar, "NetUIHWND", null, 0);
            if (surface == IntPtr.Zero)
            {
                return tabs;
            }

            Guid iid = NativeMethods.IID_IAccessible;
            IAccessible root;
            if (NativeMethods.AccessibleObjectFromWindow(surface, NativeMethods.OBJID_CLIENT, ref iid, out root) != 0
                || root == null)
            {
                return tabs;
            }

            held.Add(root);
            int visited = 0;
            Collect(root, 0, tabs, held, ref visited);
            return tabs;
        }

        // Walks the accessibility tree until it finds the container that holds the page tabs, then lists them.
        private static bool Collect(IAccessible node, int depth, List<RibbonTab> tabs, List<object> held,
            ref int visited)
        {
            if (depth > NodeDepthLimit || visited > NodeLimit)
            {
                return false;
            }

            int count = node.get_accChildCount();
            if (count <= 0)
            {
                return false;
            }

            object[] children = new object[count];
            int obtained;
            if (NativeMethods.AccessibleChildren(node, 0, count, children, out obtained) != 0)
            {
                return false;
            }

            bool found = false;
            List<IAccessible> deeper = new List<IAccessible>();
            for (int i = 0; i < obtained; i++)
            {
                visited++;
                IAccessible element = children[i] as IAccessible;
                IAccessible container = element ?? node;
                object child = element != null ? (object)SelfChild : children[i];
                if (element != null)
                {
                    held.Add(element);
                }

                int role = ToInt(container.get_accRole(child));
                if (role == NativeMethods.ROLE_SYSTEM_PAGETAB)
                {
                    int state = ToInt(container.get_accState(child));
                    if ((state & NativeMethods.STATE_SYSTEM_INVISIBLE) == 0)
                    {
                        RibbonTab tab = new RibbonTab();
                        tab.Container = container;
                        tab.Child = child;
                        tab.Name = container.get_accName(child);
                        tab.Selected = (state & NativeMethods.STATE_SYSTEM_SELECTED) != 0;
                        tabs.Add(tab);
                    }

                    found = true;
                }
                else if (element != null)
                {
                    deeper.Add(element);
                }
            }

            if (found)
            {
                return true;
            }

            foreach (IAccessible next in deeper)
            {
                if (Collect(next, depth + 1, tabs, held, ref visited))
                {
                    return true;
                }
            }

            return false;
        }

        // A child window, at any depth, with the given class and, when text is not null, the given title.
        private static IntPtr FindChild(IntPtr parent, string className, string text, int depth)
        {
            if (depth > WindowDepthLimit)
            {
                return IntPtr.Zero;
            }

            for (IntPtr child = NativeMethods.GetWindow(parent, NativeMethods.GW_CHILD); child != IntPtr.Zero;
                child = NativeMethods.GetWindow(child, NativeMethods.GW_HWNDNEXT))
            {
                if (Text(child, true) == className && (text == null || Text(child, false) == text))
                {
                    return child;
                }

                IntPtr found = FindChild(child, className, text, depth + 1);
                if (found != IntPtr.Zero)
                {
                    return found;
                }
            }

            return IntPtr.Zero;
        }

        private static string Text(IntPtr hWnd, bool classOnly)
        {
            StringBuilder value = new StringBuilder(ClassCapacity);
            if (classOnly)
            {
                NativeMethods.GetClassName(hWnd, value, value.Capacity);
            }
            else
            {
                NativeMethods.GetWindowText(hWnd, value, value.Capacity);
            }

            return value.ToString();
        }

        private static int ToInt(object value)
        {
            try
            {
                return value == null ? 0 : Convert.ToInt32(value);
            }
            catch (Exception)
            {
                return 0;
            }
        }
    }
}
