// ---------------------------------------------------------------------------------------------------------------
// Name:     ToolbarSet.cs
// Purpose:  Everything OpenToolbars shows for one app: the menus of the menu bar, the Standard and Formatting rows,
//           and the Drawing group (its Draw menu and two rows of buttons). Each row item is a built-in Office
//           command name (idMso).
// Revision: 1
// DTG:      261024210000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

namespace OpenTools
{
    internal sealed class ToolbarSet
    {
        private readonly ToolbarMenu[] _menus;
        private readonly string[] _standardRow;
        private readonly string[] _formattingRow;
        private readonly ToolbarMenu _drawMenu;
        private readonly string[] _drawingRow1;
        private readonly string[] _drawingRow2;

        public ToolbarSet(ToolbarMenu[] menus, string[] standardRow, string[] formattingRow, ToolbarMenu drawMenu,
            string[] drawingRow1, string[] drawingRow2)
        {
            _menus = menus;
            _standardRow = standardRow;
            _formattingRow = formattingRow;
            _drawMenu = drawMenu;
            _drawingRow1 = drawingRow1;
            _drawingRow2 = drawingRow2;
        }

        public ToolbarMenu[] Menus
        {
            get { return _menus; }
        }

        public string[] StandardRow
        {
            get { return _standardRow; }
        }

        public string[] FormattingRow
        {
            get { return _formattingRow; }
        }

        public ToolbarMenu DrawMenu
        {
            get { return _drawMenu; }
        }

        public string[] DrawingRow1
        {
            get { return _drawingRow1; }
        }

        public string[] DrawingRow2
        {
            get { return _drawingRow2; }
        }
    }
}
