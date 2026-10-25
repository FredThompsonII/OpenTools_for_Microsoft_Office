// ---------------------------------------------------------------------------------------------------------------
// Name:     TabBarLayout.cs
// Purpose:  Geometry only. Given tab widths and the bar width, works out where each tab, the new-document button,
//           the overflow arrow and the close-all button go. Multi-line wraps tabs onto more rows. Single-line
//           scrolls so the current tab stays visible and puts the overflow arrow at the end. The close-all button
//           is always at the far right of the first row, and tabs stop short of it. A tab with no room gets an
//           empty rectangle.
// Revision: 3
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System.Collections.Generic;
using System.Drawing;

namespace OpenTools
{
    internal sealed class TabBarLayout
    {
        private readonly List<Rectangle> _tabRects = new List<Rectangle>();

        public IList<Rectangle> TabRects
        {
            get { return _tabRects; }
        }

        public Rectangle NewButtonRect { get; private set; }

        public Rectangle OverflowRect { get; private set; }

        public Rectangle CloseAllRect { get; private set; }

        public int Height { get; private set; }

        // ----- Calculation -----

        // closeAllWidth is the width of the close-all button, or 0 when there is none.
        public void Compute(int barWidth, int rowHeight, int gap, IList<int> tabWidths, int currentIndex,
            bool multiLine, bool showNewButton, int buttonWidth, int closeAllWidth)
        {
            _tabRects.Clear();
            NewButtonRect = Rectangle.Empty;
            OverflowRect = Rectangle.Empty;
            CloseAllRect = Rectangle.Empty;
            if (barWidth < 1)
            {
                barWidth = 1;
            }

            if (closeAllWidth > 0)
            {
                CloseAllRect = new Rectangle(barWidth - closeAllWidth, 0, closeAllWidth, rowHeight);
            }

            if (multiLine)
            {
                ComputeMultiLine(barWidth, rowHeight, gap, tabWidths, showNewButton, buttonWidth, closeAllWidth);
            }
            else
            {
                ComputeSingleLine(barWidth, rowHeight, gap, tabWidths, currentIndex, showNewButton, buttonWidth,
                    closeAllWidth);
            }
        }

        // The first row gives up room for the close-all button. Later rows use the whole width.
        private static int RowLimit(int row, int barWidth, int closeAllWidth)
        {
            return row == 0 ? barWidth - closeAllWidth : barWidth;
        }

        private void ComputeMultiLine(int barWidth, int rowHeight, int gap, IList<int> tabWidths,
            bool showNewButton, int buttonWidth, int closeAllWidth)
        {
            int x = 0;
            int row = 0;
            for (int i = 0; i < tabWidths.Count; i++)
            {
                int width = tabWidths[i] > barWidth ? barWidth : tabWidths[i];
                if (x > 0 && x + width > RowLimit(row, barWidth, closeAllWidth))
                {
                    row++;
                    x = 0;
                }

                int limit = RowLimit(row, barWidth, closeAllWidth);
                if (width > limit)
                {
                    width = limit > 1 ? limit : 1;
                }

                _tabRects.Add(new Rectangle(x, row * rowHeight, width, rowHeight));
                x += width + gap;
            }

            if (showNewButton)
            {
                if (x > 0 && x + buttonWidth > RowLimit(row, barWidth, closeAllWidth))
                {
                    row++;
                    x = 0;
                }

                NewButtonRect = new Rectangle(x, row * rowHeight, buttonWidth, rowHeight);
            }

            Height = (row + 1) * rowHeight;
        }

        private void ComputeSingleLine(int barWidth, int rowHeight, int gap, IList<int> tabWidths,
            int currentIndex, bool showNewButton, int buttonWidth, int closeAllWidth)
        {
            int reserved = buttonWidth + (showNewButton ? buttonWidth : 0) + closeAllWidth;
            int available = barWidth - reserved;
            int first = 0;

            // Scroll forward until the current tab fits in the visible run.
            if (currentIndex >= 0 && currentIndex < tabWidths.Count)
            {
                while (first < currentIndex && RunWidth(tabWidths, first, currentIndex, gap) > available)
                {
                    first++;
                }
            }

            for (int i = 0; i < tabWidths.Count; i++)
            {
                _tabRects.Add(Rectangle.Empty);
            }

            int x = 0;
            for (int i = first; i < tabWidths.Count; i++)
            {
                int width = tabWidths[i];
                if (x + width > available && i > first)
                {
                    break;
                }

                if (width > available)
                {
                    width = available > 1 ? available : 1;
                }

                _tabRects[i] = new Rectangle(x, 0, width, rowHeight);
                x += width + gap;
            }

            if (showNewButton)
            {
                NewButtonRect = new Rectangle(x, 0, buttonWidth, rowHeight);
            }

            OverflowRect = new Rectangle(barWidth - buttonWidth - closeAllWidth, 0, buttonWidth, rowHeight);
            Height = rowHeight;
        }

        private static int RunWidth(IList<int> tabWidths, int first, int last, int gap)
        {
            int total = 0;
            for (int i = first; i <= last; i++)
            {
                total += tabWidths[i] + gap;
            }

            return total;
        }

        // ----- Hit testing -----

        public int HitTestTab(Point point)
        {
            for (int i = 0; i < _tabRects.Count; i++)
            {
                if (!_tabRects[i].IsEmpty && _tabRects[i].Contains(point))
                {
                    return i;
                }
            }

            return -1;
        }
    }
}
