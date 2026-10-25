// ---------------------------------------------------------------------------------------------------------------
// Name:     ToolbarFile.cs
// Purpose:  Reads an app's OpenToolbars command list (Word.commands, Excel.commands, PowerPoint.commands). The
//           build compresses each list into the DLL as a resource named OpenTools.<App>.commands.gz, which keeps
//           the DLL small; it is read once, when the ribbon loads. Format: "menu <id> <caption number>
//           <fallback caption>" starts a menu, 0 meaning no caption number. The menu named ot_menu_draw is the
//           Drawing group's Draw menu. "standard", "formatting", "drawing1" and "drawing2" start the rows. Lines
//           starting with "#" and blank lines are skipped. Every other line is one Office command name, or
//           "Parent|Command" for a submenu item. Returns null when the app has no list.
// Revision: 1
// DTG:      261025050000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;

namespace OpenTools
{
    internal static class ToolbarFile
    {
        private const string DrawMenuId = "ot_menu_draw";
        private const string MenuWord = "menu";

        // ----- Loading -----

        public static ToolbarSet Load(string app)
        {
            Assembly assembly = typeof(ToolbarFile).Assembly;
            using (Stream packed = assembly.GetManifestResourceStream("OpenTools." + app + ".commands.gz"))
            {
                if (packed == null)
                {
                    return null;
                }

                using (GZipStream unpacked = new GZipStream(packed, CompressionMode.Decompress))
                using (StreamReader reader = new StreamReader(unpacked, Encoding.ASCII))
                {
                    return Parse(reader);
                }
            }
        }

        private static ToolbarSet Parse(StreamReader reader)
        {
            List<ToolbarMenu> menus = new List<ToolbarMenu>();
            ToolbarMenu drawMenu = null;
            Dictionary<string, List<string>> rows = new Dictionary<string, List<string>>();
            rows["standard"] = new List<string>();
            rows["formatting"] = new List<string>();
            rows["drawing1"] = new List<string>();
            rows["drawing2"] = new List<string>();

            string menuId = null;
            int captionId = 0;
            string fallback = null;
            List<string> current = null;
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                line = line.Trim();
                if (line.Length == 0 || line[0] == '#')
                {
                    continue;
                }

                List<string> row;
                if (rows.TryGetValue(line, out row))
                {
                    drawMenu = Finish(menus, menuId, captionId, fallback, current) ?? drawMenu;
                    menuId = null;
                    current = row;
                    continue;
                }

                if (line.StartsWith(MenuWord + " ", StringComparison.Ordinal))
                {
                    drawMenu = Finish(menus, menuId, captionId, fallback, current) ?? drawMenu;
                    string[] parts = line.Split(new[] { ' ' }, 4);
                    menuId = parts[1];
                    captionId = int.Parse(parts[2]);
                    fallback = parts[3];
                    current = new List<string>();
                    continue;
                }

                if (current != null)
                {
                    current.Add(line);
                }
            }

            drawMenu = Finish(menus, menuId, captionId, fallback, current) ?? drawMenu;
            return new ToolbarSet(menus.ToArray(), rows["standard"].ToArray(), rows["formatting"].ToArray(), drawMenu,
                rows["drawing1"].ToArray(), rows["drawing2"].ToArray());
        }

        // Closes the menu being read, if any. The Draw menu is returned instead of being added to the menu bar.
        private static ToolbarMenu Finish(List<ToolbarMenu> menus, string menuId, int captionId, string fallback,
            List<string> items)
        {
            if (menuId == null || items == null)
            {
                return null;
            }

            ToolbarMenu menu = new ToolbarMenu(menuId, captionId, fallback, items.ToArray());
            if (menuId == DrawMenuId)
            {
                return menu;
            }

            menus.Add(menu);
            return null;
        }
    }
}
