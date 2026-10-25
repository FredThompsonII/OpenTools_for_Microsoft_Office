// ---------------------------------------------------------------------------------------------------------------
// Name:     ToolbarMenu.cs
// Purpose:  One classic drop-down menu on the OpenTools tab. CaptionId is the number of the matching menu on
//           Office's hidden Office 2003 menu bar, whose caption gives the heading in the user's own language and its
//           Alt letter. FallbackCaption is used when that caption can't be read; its & marks the Alt letter. Each
//           item is a built-in Office command name (idMso). An item written as "Parent|Command" goes into a submenu
//           named Parent, which is placed where Parent first appears.
// Revision: 1
// DTG:      261024210000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

namespace OpenTools
{
    internal sealed class ToolbarMenu
    {
        private readonly string _id;
        private readonly int _captionId;
        private readonly string _fallbackCaption;
        private readonly string[] _items;

        public ToolbarMenu(string id, int captionId, string fallbackCaption, string[] items)
        {
            _id = id;
            _captionId = captionId;
            _fallbackCaption = fallbackCaption;
            _items = items;
        }

        public string Id
        {
            get { return _id; }
        }

        public int CaptionId
        {
            get { return _captionId; }
        }

        public string FallbackCaption
        {
            get { return _fallbackCaption; }
        }

        public string[] Items
        {
            get { return _items; }
        }
    }
}
