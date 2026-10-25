// ---------------------------------------------------------------------------------------------------------------
// Name:     AssemblyInfo.cs
// Purpose:  Assembly-wide attributes. Nothing is visible to COM unless a class asks for it. Only Connect and
//           TabPaneControl do. The file version is not set here. build.bat stamps a newer one on every build,
//           so Windows Installer always replaces the installed DLL. AssemblyVersion stays 1.0.0.0 because the
//           COM registration names it.
// Revision: 4
// DTG:      261025060000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("OpenTools for Microsoft Office")]
[assembly: AssemblyDescription("Adds a tab bar to Microsoft Office desktop apps.")]
[assembly: AssemblyCompany("Fred Thompson")]
[assembly: AssemblyProduct("OpenTools for Microsoft Office")]
[assembly: AssemblyCopyright("Copyright (c) 2026 Fred Thompson. MIT license.")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: ComVisible(false)]
