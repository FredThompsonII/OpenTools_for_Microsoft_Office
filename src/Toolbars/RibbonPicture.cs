// ---------------------------------------------------------------------------------------------------------------
// Name:     RibbonPicture.cs
// Purpose:  Turns a .NET picture into the picture object Office's ribbon expects from a getImage callback
//           (IPictureDisp). The conversion method is a protected member of AxHost, so this small class
//           inherits AxHost only to reach it. It is never created.
// Revision: 1
// DTG:      261025020000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System.Drawing;
using System.Windows.Forms;

namespace OpenTools
{
    internal sealed class RibbonPicture : AxHost
    {
        private RibbonPicture()
            : base(string.Empty)
        {
        }

        public static object From(Image image)
        {
            return GetIPictureDispFromPicture(image);
        }
    }
}
