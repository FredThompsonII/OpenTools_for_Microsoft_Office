// ---------------------------------------------------------------------------------------------------------------
// Name:     IAppAdapter.cs
// Purpose:  What the shared core asks of an app adapter. The adapter listens to its own Office app and reports
//           documents to TabCoordinator. The coordinator calls back through this interface to activate, close
//           or create documents. Each Office app gets one small adapter.
// Revision: 6
// DTG:      261024200000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

namespace OpenTools
{
    internal interface IAppAdapter
    {
        string AppName { get; }

        void Start(TabCoordinator coordinator);

        void Stop();

        void ActivateDocument(string key);

        // Returns true when the document is closed. It returns false when the user cancelled a save prompt.
        bool CloseDocument(string key);

        void NewDocument();

        // Closes the Office app. Call it after every document is closed.
        void QuitApplication();
    }
}
