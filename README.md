# OpenTools for Microsoft Office

Classic toolbars, lorem ipsum, and tab support for Microsoft Office 2016-2024.

* **OpenTabs:** a tab bar. Each open document gets a tab. Click a tab to bring that document to the front.
* **OpenToolbars:** the classic menus and toolbars, on a ribbon tab named OpenTools.
* **OpenLorem:** Lorem Ipsum filler text, from the Insert menu or the Insert tab.

This is a new, independent project released under the MIT license. It contains no code from any other tab or
classic-menu product.

## Status

Public Release 1.00

## Supported apps

| App | OpenTabs | OpenToolbars | OpenLorem |
|---|---|---|---|
| Word, Excel, PowerPoint | Yes | Yes | Yes |
| Project, Visio | Yes | No | No |

In Word, Excel and PowerPoint the tab bar sits inside each document's main window, between the ribbon and the
document. In Project and Visio it sits inside the one main window, above the document area, because every document is
a window inside it.

Outlook, Publisher and Access aren't supported. Outlook's windows are single messages and appointments. Publisher
and Access open every document in its own copy of the app, with its own taskbar button, so tabs add little.

## OpenToolbars

The OpenTools ribbon tab sits first, before Home, in Word, Excel and PowerPoint. It holds:

* **Menus and toolbars:** the classic menu bar, then the Standard row, then the Formatting row.
* **Drawing:** the Draw menu and two rows of drawing buttons.
* **OpenTools:** a large Settings button.

Every command is one of Office's own, so Office supplies its picture, label and action. Commands an Office version
doesn't have are left out, and commands that can't act right now are grayed out. Menu headings come from Office's
hidden Office 2003 menu bar in your language, with their Alt letters. The Slide Show heading in PowerPoint is
English. The View menu ends with OpenTools Settings.

When the window narrows, Office shrinks the groups from the right, and finally folds a whole group into one
drop-down button. Nothing is lost off the edge.

Each app opens on the OpenTools tab, and so does each new window. Turn this off per app on the Apps page.

## OpenLorem

Lorem Ipsum is the last item on the classic Insert menu, and a button at the far right of the Insert tab.

| App | What it does |
|---|---|
| Word | Types at the cursor: sections, an optional title per section, then 3 to 5 paragraphs of 5 sentences |
| Excel | Types one paragraph into the active cell |
| PowerPoint | Fills the empty text placeholders on the current slide |

The counts are set on the Lorem Ipsum settings page.

## Build and install

Extract this package into a new folder and run `build.bat` as your normal Windows user. Internet access is required to
build. Open the resulting `OpenTools_for_Microsoft_Office_1.00.msi`. Its one window shows what OpenTools is, then the
publisher certificate question. Approve the certificate, then start Word, Excel, PowerPoint, Project or Visio. Close
those apps before installing, repairing, upgrading or uninstalling. Install using the Windows account that will run
Office. No administrator elevation is needed for this per-user MSI.

The add-in is registered as `OpenTools.Connect` and installs to
`%LOCALAPPDATA%\Programs\OpenTools`.

Uninstalling removes the add-in, the certificate trust it added, the settings in
`HKEY_CURRENT_USER\Software\OpenTools` and the log folder `%LOCALAPPDATA%\OpenTools`. Installing a newer build over
an older one keeps the settings.

## Contents and operation

`build.bat` is the one build file. It holds the PowerShell build logic, the WiX installer definition, and the
certificate trust and rollback steps. The C# source is in `src`. The OpenToolbars command lists are plain text files
in `src\Toolbars` (`Word.commands`, `Excel.commands`, `PowerPoint.commands`). The build compresses them into the DLL.
The small helper that adds and removes publisher trust is in `installer\TrustPublisher.cs`. The build has no tests,
probes or monitoring frameworks. The add-in keeps a plain log file at `%LOCALAPPDATA%\OpenTools\OpenTools.log`.

The batch downloads a private .NET 8 SDK (used only to run the C# compiler), WiX 3.14, the .NET Framework 4.8
reference files, and the Word, Excel, PowerPoint and Office core interop packages. Each download is checked against
a pinned checksum. The build stops if one doesn't match. Project and Visio need no interop package. The add-in
reaches them by late binding.

The downloads and build intermediates go into temporary folders under `%TEMP%`. The build also writes two helper
scripts beside `build.bat`. All of them are deleted on completion or on a handled failure. A forced termination or
power loss can leave them behind. The build doesn't uninstall system runtimes. Each build therefore downloads its
build tools again.

Each build stamps the add-in DLL with a newer file version, so the installer always replaces an installed DLL.

## Settings

Each app has an **OpenTools** button at the end of its View tab. The Settings button on the OpenTools tab and the
last item of the View menu open the same window:

* **Settings for:** all apps, or one app's own overrides. "Clear this app's overrides" returns an app to the
  settings for all apps.
* **Apps:** one row per installed Office app, with Tabs, Toolbars, Lorem Ipsum and Start on OpenTools tab. All are
  on by default. Project and Visio show Tabs only. Changes take effect the next time that app starts.
* **Tabs:** double-click or middle-click to close, drag to reorder, bold current tab, file icons, + button,
  close-all button.
* **Windows:** single-window mode, and keep the selected ribbon tab in sync. Word, Excel and PowerPoint only.
* **Look:** auto or fixed tab width, extra rows, font and font size.
* **Lorem Ipsum:** sections, paragraphs from and to, sentences, and a title at the start of each section.
* **About:** name and version, what it does, copyright, license and the GitHub address.
* **Backup and Restore:** export and import the settings as a .reg file. Restore refuses a file that touches any
  other registry key or deletes anything.

Settings live in `HKEY_CURRENT_USER\Software\OpenTools`. A value under `Apps\Word` overrides the global value
for Word. A value that is absent uses the built-in default.

## Signing

Choose one:

* Leave `Signing` empty: the build generates a reusable 3072-bit RSA code-signing certificate, valid for three
  years. It saves the encrypted private key as `Signing\CodeSigning.pfx` and its password as
  `Signing\CodeSigning.password.dpapi`, protected by the current Windows account on this computer. The key also stays
  in your Current User Personal store. Later builds reuse it.
* Supply `Signing\CodeSigning.pfx`: enter its password when prompted.
* Supply a thumbprint: one 40-character certificate thumbprint, for an RSA code-signing certificate with its
  private key in your Current User Personal store.

The key the build chooses is recorded per computer in `Signing\CertificateThumbprint_<computer>.txt`, because a key
lives in one computer's certificate store. If the recorded key isn't usable on a computer, the build finds that
computer's own OpenTools key in its store before it tries the PFX. It asks you to choose only when it can't tell
which key is yours. Nothing in `Signing` is ever deleted or regenerated by the build.

Generated certificates expire after three years. This local signing doesn't use an online timestamp service.
Preserve your key material securely. Generated DPAPI passwords don't work under another Windows account or on
another computer. Private keys are never included in the MSI and are ignored by Git. A supplied organization
certificate requires its issuing CA chain to be trusted on the destination machine.

The installer displays the exact certificate thumbprint and requires consent before adding the public certificate
to Current User Trusted Publishers. For a self-signed certificate, it also adds it to Current User Trusted Root. It
preserves pre-existing trust, rolls back its changes if installation fails, and removes only the trust it added on
uninstall. Trust applies to other code signed by this key too. Office security policies aren't changed.
Organization policy may prohibit local publisher trust or the add-in even with a valid signature.

## Requirements

Windows 10 or 11, Windows PowerShell 5.1, and desktop Office 2016-2024, 32-bit or 64-bit. The destination needs .NET
Framework 4.8, which is built into Windows 10 version 1903 and later and Windows 11. Testing so far is on Office 2021,
32-bit.

## Warranty

Absolutely no warranty or guarantee of any sort. It works for me.

## Licensing and source

OpenTools for Microsoft Office is MIT licensed; see LICENSE. The Microsoft interop packages, reference files, .NET SDK
and WiX are downloaded at build time under their own licenses. They aren't included in this repository.

Source: https://github.com/FredThompsonII/OpenTools_for_Microsoft_Office

GitHub holds source and `build.bat` only. No installers, certificates or keys. Each user builds and signs their own
MSI. Don't publish generated signing material or locally built installers to GitHub.
