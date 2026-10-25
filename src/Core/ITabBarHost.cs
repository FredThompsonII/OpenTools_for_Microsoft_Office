// ---------------------------------------------------------------------------------------------------------------
// Name:     ITabBarHost.cs
// Purpose:  What the coordinator needs from anything that shows a tab bar: refresh it after the settings change,
//           hold its drawing for a moment, and remove it. EmbeddedBarHost puts the bar inside the Office main
//           window. TaskPaneHost is the fallback that floats the bar over an Office task pane.
// Revision: 3
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;

namespace OpenTools
{
    internal interface ITabBarHost : IDisposable
    {
        void Refresh();

        // While held, the bar does not redraw. Changes are drawn all at once when the hold ends.
        void HoldPaint(bool hold);
    }
}
