// ---------------------------------------------------------------------------------------------------------------
// Name:     AppProfile.cs
// Purpose:  Holds the window class names that differ between Office apps: the class of the main window (the
//           frame) and the class of the document area inside it. There is one Office app per process, so the
//           values are static. Connect sets them once at start-up. The defaults are the Word values.
// Revision: 2
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

namespace OpenTools
{
    internal static class AppProfile
    {
        private const string WordFrameClass = "OpusApp";
        private const string WordDocumentAreaClass = "_WwF";

        private static string _frameClass = WordFrameClass;
        private static string _documentAreaClass = WordDocumentAreaClass;

        internal static string FrameClass
        {
            get { return _frameClass; }
        }

        internal static string DocumentAreaClass
        {
            get { return _documentAreaClass; }
        }

        internal static void Use(string frameClass, string documentAreaClass)
        {
            _frameClass = frameClass;
            _documentAreaClass = documentAreaClass;
        }
    }
}
