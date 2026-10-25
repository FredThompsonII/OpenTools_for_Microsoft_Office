// ---------------------------------------------------------------------------------------------------------------
// Name:     ToolbarRibbonXml.cs
// Purpose:  Builds the ribbon XML for the OpenTools tab from a ToolbarSet. The tab sits first, before Home. Group
//           one holds three rows: the menu bar, the Standard row and the Formatting row. Group two is the Drawing
//           group: the Draw menu and two rows of buttons. Every command is a built-in Office control placed with
//           <control idMso>, so Office supplies its own picture, label, state and action. Menu headings call back
//           to Connect (GetMenuLabel, GetMenuKeytip) so they can show Office's own caption in the user's
//           language. The menu's tag carries the caption number and the fallback caption, separated by "|".
//           The View menu ends with a separator and "OpenTools Settings...", which opens the settings window.
//           The Insert menu ends with "Lorem Ipsum" when Lorem Ipsum is turned on for the app.
//           Group three, at the far right, holds the large Settings button. Settings uses OpenTools' own gear
//           picture (GearIcon), supplied through the GetSettingsImage callback.
//           When the window narrows, Office shrinks the groups from the right and finally folds a whole group
//           into one drop-down button, so nothing is lost off the edge.
// Revision: 4
// DTG:      261025030000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System.Collections.Generic;
using System.Security;
using System.Text;

namespace OpenTools
{
    internal static class ToolbarRibbonXml
    {
        public const string TabId = "OpenToolsTab";
        private const string ViewMenuId = "ot_menu_view";
        private const string InsertMenuId = "ot_menu_insert";
        private const char ParentSeparator = '|';

        // ----- Tab -----

        public static string Tab(ToolbarSet set, bool lorem)
        {
            StringBuilder x = new StringBuilder();
            x.Append("<tab id=\"").Append(TabId).Append("\" label=\"OpenTools\" insertBeforeMso=\"TabHome\">");

            x.Append("<group id=\"ot_group_toolbars\" label=\"Menus and toolbars\">");
            x.Append("<box id=\"ot_row_menus\" boxStyle=\"horizontal\">");
            foreach (ToolbarMenu menu in set.Menus)
            {
                AppendMenu(x, menu, lorem);
            }

            x.Append("</box>");
            AppendRow(x, "ot_row_standard", set.StandardRow);
            AppendRow(x, "ot_row_formatting", set.FormattingRow);
            x.Append("</group>");

            x.Append("<group id=\"ot_group_drawing\" label=\"Drawing\">");
            x.Append("<box id=\"ot_row_drawing1\" boxStyle=\"horizontal\">");
            AppendMenu(x, set.DrawMenu, false);
            AppendControls(x, set.DrawingRow1);
            x.Append("</box>");
            AppendRow(x, "ot_row_drawing2", set.DrawingRow2);
            x.Append("</group>");

            x.Append("<group id=\"ot_group_opentools\" label=\"OpenTools\">");
            x.Append("<button id=\"ot_settings\" label=\"Settings\" size=\"large\" getImage=\"GetSettingsImage\" ")
                .Append("onAction=\"OnOpenSettings\" screentip=\"OpenTools settings\" />");
            x.Append("</group>");

            x.Append("</tab>");
            return x.ToString();
        }

        // ----- Rows -----

        private static void AppendRow(StringBuilder x, string id, string[] commands)
        {
            x.Append("<box id=\"").Append(id).Append("\" boxStyle=\"horizontal\">");
            AppendControls(x, commands);
            x.Append("</box>");
        }

        private static void AppendControls(StringBuilder x, string[] commands)
        {
            foreach (string command in commands)
            {
                x.Append("<control idMso=\"").Append(Escape(command)).Append("\" showLabel=\"false\" />");
            }
        }

        // ----- Menus -----

        // Items keep their table order. A submenu is placed where its first item appears and collects every
        // item with the same parent.
        private static void AppendMenu(StringBuilder x, ToolbarMenu menu, bool lorem)
        {
            x.Append("<menu id=\"").Append(menu.Id).Append("\" tag=\"").Append(menu.CaptionId).Append(ParentSeparator)
                .Append(Escape(menu.FallbackCaption)).Append("\"");
            x.Append(" getLabel=\"GetMenuLabel\" getKeytip=\"GetMenuKeytip\" itemSize=\"normal\">");

            Dictionary<string, List<string>> submenus = new Dictionary<string, List<string>>();
            List<string> order = new List<string>();
            foreach (string item in menu.Items)
            {
                int split = item.IndexOf(ParentSeparator);
                if (split < 0)
                {
                    order.Add(item);
                    continue;
                }

                string parent = item.Substring(0, split);
                List<string> children;
                if (!submenus.TryGetValue(parent, out children))
                {
                    children = new List<string>();
                    submenus[parent] = children;
                    order.Add(ParentSeparator + parent);
                }

                children.Add(item.Substring(split + 1));
            }

            int number = 0;
            foreach (string entry in order)
            {
                if (entry[0] != ParentSeparator)
                {
                    x.Append("<control idMso=\"").Append(Escape(entry)).Append("\" />");
                    continue;
                }

                string parent = entry.Substring(1);
                number++;
                x.Append("<menu id=\"").Append(menu.Id).Append("_sub").Append(number).Append("\" label=\"")
                    .Append(Escape(parent)).Append("\" itemSize=\"normal\">");
                foreach (string child in submenus[parent])
                {
                    x.Append("<control idMso=\"").Append(Escape(child)).Append("\" />");
                }

                x.Append("</menu>");
            }

            if (menu.Id == InsertMenuId && lorem)
            {
                x.Append("<button id=\"ot_menu_insert_lorem\" label=\"Lorem Ipsum\" onAction=\"OnLoremIpsum\" />");
            }

            if (menu.Id == ViewMenuId)
            {
                x.Append("<menuSeparator id=\"ot_menu_view_settings_line\" />");
                x.Append("<button id=\"ot_menu_view_settings\" label=\"OpenTools Settings...\" ")
                    .Append("getImage=\"GetSettingsImage\" onAction=\"OnOpenSettings\" />");
            }

            x.Append("</menu>");
        }

        private static string Escape(string text)
        {
            return SecurityElement.Escape(text);
        }
    }
}
