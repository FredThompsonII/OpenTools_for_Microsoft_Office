// ---------------------------------------------------------------------------------------------------------------
// Name:     TrustPublisher.cs
// Purpose:  Windowless helper that the OpenTools installer runs to add or remove the publisher certificate for the
//           current Windows user. It replaces the PowerShell script the installer used before, which opened a
//           console window. This is a Windows application with no window and no console. Results go to
//           %LOCALAPPDATA%\OpenTools\TrustHelper.log. The exit code is 0 for success and 1 for failure.
//           Arguments: -C certificate file, -E expected thumbprint, -T transaction id, -O operation, which is
//           Install, Repair, Uninstall, Rollback, Commit or Cleanup. A journal file lets Rollback undo an
//           Install, Repair or Uninstall. Commit removes the journal once an install or repair has finished. An
//           ownership file records which certificate entries this helper added, so Uninstall removes only
//           those. Cleanup runs last on a real uninstall, never on an upgrade: it deletes the settings key
//           HKCU\Software\OpenTools and the folder %LOCALAPPDATA%\OpenTools with its logs. Nothing is logged
//           after Cleanup, because the log folder is what it deletes.
// Revision: 3
// DTG:      261025050000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace OpenTools.Installer
{
    internal static class TrustPublisher
    {
        private const int NetFramework48Release = 528040;
        private const string StateSchema = "1";

        private static string _appFolder;
        private static string _stateFolder;

        private static int Main(string[] args)
        {
            string certificatePath = null;
            string expectedThumbprint = null;
            string transactionId = null;
            string operation = null;
            X509Store rootStore = null;
            X509Store publisherStore = null;

            try
            {
                string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                _appFolder = Path.Combine(local, "OpenTools");
                _stateFolder = Path.Combine(_appFolder, "TrustTransactions");

                for (int i = 0; i + 1 < args.Length; i += 2)
                {
                    switch (args[i].ToUpperInvariant())
                    {
                        case "-C":
                            certificatePath = args[i + 1];
                            break;
                        case "-E":
                            expectedThumbprint = args[i + 1];
                            break;
                        case "-T":
                            transactionId = args[i + 1];
                            break;
                        case "-O":
                            operation = args[i + 1];
                            break;
                        default:
                            throw new ArgumentException("Unknown argument " + args[i] + ".");
                    }
                }

                Require(operation != null && Regex.IsMatch(operation,
                    "^(Install|Repair|Uninstall|Rollback|Commit|Cleanup)$"), "The operation is missing or not valid.");
                Require(transactionId != null && Regex.IsMatch(transactionId, "^[A-Fa-f0-9]{32}$"),
                    "The transaction id is missing or not valid.");

                string journal = Path.Combine(_stateFolder, transactionId + ".journal.state");
                string ownership = Path.Combine(_stateFolder, transactionId + ".owned.state");

                if (operation == "Cleanup")
                {
                    return Cleanup();
                }

                if (operation == "Commit")
                {
                    if (File.Exists(journal))
                    {
                        File.Delete(journal);
                    }

                    return 0;
                }

                Require(certificatePath != null && File.Exists(certificatePath), "The certificate file is missing.");
                Require(expectedThumbprint != null && Regex.IsMatch(expectedThumbprint, "^[A-Fa-f0-9]{40}$"),
                    "The expected thumbprint is missing or not valid.");

                X509Certificate2 cert = new X509Certificate2(certificatePath);
                Require(string.Equals(cert.Thumbprint, expectedThumbprint, StringComparison.OrdinalIgnoreCase),
                    "The public certificate differs from the approved installer signer.");
                Require(!cert.HasPrivateKey, "Private signing keys must not be installed.");
                bool selfIssued = string.Equals(cert.Subject, cert.Issuer, StringComparison.Ordinal);

                rootStore = new X509Store("Root", StoreLocation.CurrentUser);
                publisherStore = new X509Store("TrustedPublisher", StoreLocation.CurrentUser);
                rootStore.Open(OpenFlags.ReadWrite);
                publisherStore.Open(OpenFlags.ReadWrite);

                if (operation == "Rollback")
                {
                    RestoreTransaction(journal, ownership, cert, selfIssued, rootStore, publisherStore);
                    return 0;
                }

                Require(!File.Exists(journal),
                    "A trust transaction is pending. Keep it, and resolve the rollback before retrying.");

                bool hasOld = File.Exists(ownership);
                bool oldRootOwned = false;
                bool oldPublisherOwned = false;
                if (hasOld)
                {
                    Dictionary<string, string> old = ReadState(ownership, cert);
                    oldRootOwned = ReadBool(old, "RootOwned");
                    oldPublisherOwned = ReadBool(old, "PublisherOwned");
                }

                if (operation == "Repair" && !hasOld)
                {
                    throw new InvalidOperationException(
                        "Repair has no recorded publisher approval. Reinstall and give explicit consent.");
                }

                if (operation == "Uninstall" && !hasOld)
                {
                    return 0;
                }

                if (operation != "Uninstall")
                {
                    CheckDotNet48();
                    DateTime now = DateTime.Now;
                    Require(cert.NotBefore <= now && cert.NotAfter > now, "The certificate is not currently valid.");
                }

                Directory.CreateDirectory(_stateFolder);
                bool hadRoot = HasCertificate(rootStore, cert);
                bool hadPublisher = HasCertificate(publisherStore, cert);

                Dictionary<string, string> journalState = NewState(cert);
                journalState["HadRoot"] = hadRoot.ToString();
                journalState["HadPublisher"] = hadPublisher.ToString();
                journalState["HasPrevious"] = hasOld.ToString();
                journalState["PreviousRootOwned"] = oldRootOwned.ToString();
                journalState["PreviousPublisherOwned"] = oldPublisherOwned.ToString();
                SaveState(journal, journalState);

                try
                {
                    if (operation == "Uninstall")
                    {
                        if (oldPublisherOwned)
                        {
                            publisherStore.Remove(cert);
                        }

                        if (oldRootOwned && selfIssued)
                        {
                            rootStore.Remove(cert);
                        }

                        File.Delete(ownership);

                        // No commit step can run after an uninstall, because the files are gone by then. The
                        // journal is removed here so that installing the same package again is not blocked.
                        File.Delete(journal);
                    }
                    else
                    {
                        bool ownRoot = selfIssued && !hadRoot;
                        bool ownPublisher = !hadPublisher;
                        if (hasOld)
                        {
                            ownRoot = ownRoot || oldRootOwned;
                            ownPublisher = ownPublisher || oldPublisherOwned;
                        }

                        if (selfIssued && !hadRoot)
                        {
                            rootStore.Add(cert);
                        }

                        if (!hadPublisher)
                        {
                            publisherStore.Add(cert);
                        }

                        Dictionary<string, string> owned = NewState(cert);
                        owned["RootOwned"] = ownRoot.ToString();
                        owned["PublisherOwned"] = ownPublisher.ToString();
                        SaveState(ownership, owned);
                    }
                }
                catch (Exception)
                {
                    RestoreTransaction(journal, ownership, cert, selfIssued, rootStore, publisherStore);
                    throw;
                }

                Log("Publisher trust " + operation + " finished.");
                return 0;
            }
            catch (Exception ex)
            {
                Log("Publisher trust " + (operation ?? "(unknown)") + " failed: " + ex.Message);
                return 1;
            }
            finally
            {
                if (publisherStore != null)
                {
                    publisherStore.Close();
                }

                if (rootStore != null)
                {
                    rootStore.Close();
                }
            }
        }

        // ----- Checks -----

        private static void Require(bool condition, string message)
        {
            if (!condition)
            {
                throw new InvalidOperationException(message);
            }
        }

        private static void CheckDotNet48()
        {
            int release = 0;
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(
                @"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
            {
                if (key != null)
                {
                    object value = key.GetValue("Release");
                    if (value is int)
                    {
                        release = (int)value;
                    }
                }
            }

            Require(release >= NetFramework48Release,
                "Requires the Microsoft .NET Framework 4.8 or a newer 4.x release. The Release registry value is "
                + release + ". Install the .NET Framework 4.8 runtime, then run this installer again.");
        }

        private static bool HasCertificate(X509Store store, X509Certificate2 cert)
        {
            return store.Certificates.Find(X509FindType.FindByThumbprint, cert.Thumbprint, false).Count > 0;
        }

        // ----- Undoing a transaction -----

        private static void RestoreTransaction(string journal, string ownership, X509Certificate2 cert,
            bool selfIssued, X509Store rootStore, X509Store publisherStore)
        {
            if (!File.Exists(journal))
            {
                return;
            }

            Dictionary<string, string> j = ReadState(journal, cert);
            bool hadPublisher = ReadBool(j, "HadPublisher");
            bool hadRoot = ReadBool(j, "HadRoot");
            if (hadPublisher)
            {
                publisherStore.Add(cert);
            }
            else
            {
                publisherStore.Remove(cert);
            }

            if (selfIssued)
            {
                if (hadRoot)
                {
                    rootStore.Add(cert);
                }
                else
                {
                    rootStore.Remove(cert);
                }
            }

            if (ReadBool(j, "HasPrevious"))
            {
                Dictionary<string, string> previous = NewState(cert);
                previous["RootOwned"] = ReadBool(j, "PreviousRootOwned").ToString();
                previous["PublisherOwned"] = ReadBool(j, "PreviousPublisherOwned").ToString();
                SaveState(ownership, previous);
            }
            else if (File.Exists(ownership))
            {
                File.Delete(ownership);
            }

            File.Delete(journal);
        }

        // ----- State files: one Name=Value pair per line -----

        private static Dictionary<string, string> NewState(X509Certificate2 cert)
        {
            Dictionary<string, string> state = new Dictionary<string, string>();
            state["Schema"] = StateSchema;
            state["Thumbprint"] = cert.Thumbprint;
            return state;
        }

        private static Dictionary<string, string> ReadState(string path, X509Certificate2 cert)
        {
            Dictionary<string, string> state = new Dictionary<string, string>();
            foreach (string line in File.ReadAllLines(path, Encoding.ASCII))
            {
                int split = line.IndexOf('=');
                if (split > 0)
                {
                    state[line.Substring(0, split)] = line.Substring(split + 1);
                }
            }

            string schema;
            string thumbprint;
            if (!state.TryGetValue("Schema", out schema) || schema != StateSchema
                || !state.TryGetValue("Thumbprint", out thumbprint)
                || !string.Equals(thumbprint, cert.Thumbprint, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("Unrecognized trust ownership or journal.");
            }

            return state;
        }

        private static bool ReadBool(Dictionary<string, string> state, string name)
        {
            string text;
            bool value;
            if (!state.TryGetValue(name, out text) || !bool.TryParse(text, out value))
            {
                throw new InvalidOperationException("Invalid flag " + name + " in a trust state file.");
            }

            return value;
        }

        private static void SaveState(string path, Dictionary<string, string> state)
        {
            string pending = path + "." + Guid.NewGuid().ToString("N") + ".pending";
            try
            {
                StringBuilder text = new StringBuilder();
                foreach (KeyValuePair<string, string> pair in state)
                {
                    text.Append(pair.Key).Append('=').Append(pair.Value).Append("\r\n");
                }

                File.WriteAllText(pending, text.ToString(), new UTF8Encoding(false));
                if (File.Exists(path))
                {
                    File.Replace(pending, path, null);
                }
                else
                {
                    File.Move(pending, path);
                }
            }
            finally
            {
                if (File.Exists(pending))
                {
                    File.Delete(pending);
                }
            }
        }

        // ----- Cleanup -----

        // Deletes the settings key and the OpenTools folder. Either one may already be gone.
        private static int Cleanup()
        {
            int result = 0;
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(@"Software\OpenTools", false);
            }
            catch (Exception ex)
            {
                Log("Cleanup could not delete the settings key: " + ex.Message);
                result = 1;
            }

            try
            {
                if (Directory.Exists(_appFolder))
                {
                    Directory.Delete(_appFolder, true);
                }
            }
            catch (Exception)
            {
                // A file still in use is left behind. The uninstall carries on.
                result = 1;
            }

            return result;
        }

        // ----- Log -----

        private static void Log(string message)
        {
            try
            {
                Directory.CreateDirectory(_appFolder);
                string line = string.Format("{0:yyyy-MM-dd HH:mm:ss} {1}{2}", DateTime.Now, message,
                    Environment.NewLine);
                File.AppendAllText(Path.Combine(_appFolder, "TrustHelper.log"), line, new UTF8Encoding(false));
            }
            catch (Exception)
            {
                // The installer must not fail because a log line could not be written.
            }
        }
    }
}
