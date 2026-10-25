// ---------------------------------------------------------------------------------------------------------------
// Name:     TabPaneControl.cs
// Purpose:  The ActiveX control Office puts inside the task pane. Office creates it by ProgId, so this class
//           must be COM visible and registered. It does nothing except hold the real tab bar.
// Revision: 2
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OpenTools
{
    [ComVisible(true)]
    [Guid("D2E0EDDB-0766-4293-BCD1-CFD0623F7038")]
    [ProgId("OpenTools.TabPane")]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class TabPaneControl : UserControl
    {
        public TabPaneControl()
        {
            Margin = Padding.Empty;
            Padding = Padding.Empty;
        }

        internal void Attach(Control bar)
        {
            Controls.Clear();
            bar.Dock = DockStyle.Fill;
            Controls.Add(bar);
        }
    }
}
