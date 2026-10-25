// ---------------------------------------------------------------------------------------------------------------
// Name:     TabLabeler.cs
// Purpose:  Builds the text shown on a tab. Files with the same name in different folders get the folder name
//           added, for example "Report.docx (Q3)". Long names are shortened in the middle so the extension
//           stays visible. Text is ASCII, so the shortened part is marked with three dots.
// Revision: 2
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;

namespace OpenTools
{
    internal static class TabLabeler
    {
        private const string Dots = "...";
        private const int MinimumNameChars = 8;

        // ----- Same-name folder suffix -----

        public static void AssignSuffixes(IList<TabItem> items)
        {
            Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (TabItem item in items)
            {
                int count;
                counts.TryGetValue(item.Name, out count);
                counts[item.Name] = count + 1;
            }

            foreach (TabItem item in items)
            {
                item.Suffix = string.Empty;
                if (counts[item.Name] < 2 || !item.HasSavedFile)
                {
                    continue;
                }

                string folderName = Path.GetFileName(item.Folder.TrimEnd('\\', '/'));
                if (string.IsNullOrEmpty(folderName))
                {
                    folderName = item.Folder;
                }

                item.Suffix = " (" + folderName + ")";
            }
        }

        // ----- Middle shortening -----

        public static string Display(TabItem item, int maxChars)
        {
            string name = item.Name;
            string suffix = item.Suffix;
            if (name.Length + suffix.Length <= maxChars)
            {
                return name + suffix;
            }

            int budget = maxChars - suffix.Length - Dots.Length;
            if (budget < MinimumNameChars)
            {
                budget = MinimumNameChars;
            }

            string extension = Path.GetExtension(name);
            string stem = name.Substring(0, name.Length - extension.Length);
            int stemChars = budget - extension.Length;
            if (stemChars < 1)
            {
                stemChars = 1;
            }

            if (stem.Length <= stemChars)
            {
                return name + suffix;
            }

            return stem.Substring(0, stemChars) + Dots + extension + suffix;
        }
    }
}
