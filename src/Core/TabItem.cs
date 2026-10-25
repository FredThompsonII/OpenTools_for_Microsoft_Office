// ---------------------------------------------------------------------------------------------------------------
// Name:     TabItem.cs
// Purpose:  One tab: one open document. Key is chosen by the app adapter and does not change when the document
//           is saved under a new name. Hwnd is the document window the stacker moves and raises.
// Revision: 2
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.IO;

namespace OpenTools
{
    internal sealed class TabItem
    {
        public TabItem(string key, IntPtr hwnd, string fullPath, string name)
        {
            Key = key;
            Hwnd = hwnd;
            FullPath = fullPath;
            Name = name;
            Suffix = string.Empty;
        }

        public string Key { get; private set; }

        public IntPtr Hwnd { get; set; }

        public string FullPath { get; set; }

        public string Name { get; set; }

        public bool IsDirty { get; set; }

        // Set by TabLabeler when another tab has the same file name, for example " (Q3)".
        public string Suffix { get; set; }

        // The folder that holds the file, or an empty string for a document that was never saved.
        public string Folder
        {
            get
            {
                try
                {
                    if (string.IsNullOrEmpty(FullPath) || !Path.IsPathRooted(FullPath))
                    {
                        return string.Empty;
                    }

                    return Path.GetDirectoryName(FullPath) ?? string.Empty;
                }
                catch (ArgumentException)
                {
                    return string.Empty;
                }
            }
        }

        public bool HasSavedFile
        {
            get { return Folder.Length > 0; }
        }
    }
}
