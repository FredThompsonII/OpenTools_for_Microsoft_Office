// ---------------------------------------------------------------------------------------------------------------
// Name:     CloseAllDialog.cs
// Purpose:  The small question window shown when the close-all button is clicked. It offers three answers: close
//           all documents, close all documents and then close the Office app, or do nothing. Cancel is the
//           default, so pressing Enter or Escape never closes anything.
// Revision: 3
// DTG:      261025060000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Drawing;
using System.Windows.Forms;

namespace OpenTools
{
    internal enum CloseAllChoice
    {
        Cancel,
        CloseDocuments,
        CloseDocumentsAndApp
    }

    internal sealed class CloseAllDialog : Form
    {
        private CloseAllChoice _choice = CloseAllChoice.Cancel;

        private CloseAllDialog(int count, string appName)
        {
            Text = "OpenTools for Microsoft Office";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Font;
            Font = SystemFonts.MessageBoxFont;

            Label question = new Label();
            question.AutoSize = true;
            question.MaximumSize = new Size(420, 0);
            question.Text = count == 1 ? "Close the open document?" : "Close all " + count + " open documents?";
            question.Location = new Point(16, 16);

            Button closeDocuments = MakeButton("Close documents", CloseAllChoice.CloseDocuments);
            Button closeAndApp = MakeButton("Close documents and " + appName, CloseAllChoice.CloseDocumentsAndApp);
            Button cancel = MakeButton("Cancel", CloseAllChoice.Cancel);

            Controls.Add(question);
            int top = question.Bottom + 16;
            int left = 16;
            foreach (Button button in new Button[] { closeDocuments, closeAndApp, cancel })
            {
                button.Location = new Point(left, top);
                Controls.Add(button);
                left = button.Right + 8;
            }

            ClientSize = new Size(Math.Max(left + 8, question.Right + 16), top + closeDocuments.Height + 16);
            AcceptButton = cancel;
            CancelButton = cancel;
            ActiveControl = cancel;
        }

        private Button MakeButton(string text, CloseAllChoice choice)
        {
            Button button = new Button();
            button.Text = text;
            button.AutoSize = true;
            button.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            button.Padding = new Padding(8, 2, 8, 2);
            button.UseVisualStyleBackColor = true;
            button.Click += delegate
            {
                _choice = choice;
                Close();
            };
            return button;
        }

        // Shows the question over the given window and returns the answer.
        public static CloseAllChoice Ask(IWin32Window owner, int count, string appName)
        {
            using (CloseAllDialog dialog = new CloseAllDialog(count, appName))
            {
                dialog.ShowDialog(owner);
                return dialog._choice;
            }
        }
    }
}
