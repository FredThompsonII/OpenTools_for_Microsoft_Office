@echo off
REM OpenTools for Microsoft Office - build.bat
REM (c) Fred Thompson. MIT license.
REM
REM Purpose: builds the OpenTools for Microsoft Office add-in from the source in the src folder and packs it into a
REM per-user Windows Installer (MSI) file that is signed with a local certificate. No Visual Studio is needed.
REM The installer also carries a small windowless helper program, built from installer\TrustPublisher.cs, that
REM adds and removes the publisher certificate without opening any console window.
REM Date: 2026-10-25, revision 43, DTG 261025120000.
REM The installer is written as OpenTools_for_Microsoft_Office_1.00.msi (the major version, then the minor
REM version as two digits), replacing the one from the previous build.
REM The chosen signing key is recorded per computer (Signing\CertificateThumbprint_<computer>.txt), because a key
REM lives in one computer's certificate store. When the recorded key is not usable in this computer's store, the
REM build recovers this computer's own OpenTools key from the store before it tries the PFX.
REM Microsoft Project and Visio need no interop package: the add-in reaches them by late binding.
REM Each build stamps the add-in DLL with a newer file version (1.0.YYDDD.N, N = seconds since midnight / 2),
REM so Windows Installer always replaces the installed DLL instead of keeping an older one of equal version.
REM Usage: double-click build.bat, or run it from a Command Prompt. It needs no arguments.
REM Dependencies: Windows PowerShell 5.1, an internet connection, and these downloads, each checked against a
REM pinned checksum: Word, Excel, PowerPoint and Office core interop packages, the
REM .NET Framework 4.8 reference files, WiX 3.14 and the Microsoft .NET SDK (used only to run the C# compiler).
REM Downloads are saved in
REM %TEMP%\OpenTools-build-cache while the build runs. Each one is checked against its checksum. At the end of
REM every build, finished or failed, the cache folder (including the .NET SDK) and the per-build working folder
REM are removed, so every build downloads its tools fresh.
REM Assumptions: existing Windows signing keys are reused; Signing may hold a PFX plus thumbprint.
REM Risk: LOW. It writes only the finished MSI beside this file and, on a first run, a certificate in Signing.
REM Nothing in the repo is deleted. The PowerShell build script is written out of this file as a temporary
REM helper file and deleted again.
REM Reset path: delete only the failed MSI. Never delete Signing keys to repair a build.
REM Exit codes: 0 success; 1 source check failed; 2 signing certificate problem; 3 download or checksum
REM failure; 4 compile failure; 5 installer build failure; 6 temporary files not removed; 7 unexpected error.
REM No elevation is needed. Nothing here changes machine-wide settings.

        cd /d %~dp0
        setlocal EnableExtensions DisableDelayedExpansion                                                               & REM Rule #135 deviation 1: the repo folder path can hold a "!" and delayed expansion would drop it
        set "EXITCODE=0"

REM -- Data section

        set "PS_HELPER=_OpenTools_write_helper.ps1"
        set "PS_BUILD=_OpenTools_build.ps1"

REM -- Write the build script out of this file

    REM The helper below reads the lines between the PS1 markers at the end of this file.

        > "%PS_HELPER%"  echo $bat = $args[0]
        >> "%PS_HELPER%" echo $out = $args[1]
        >> "%PS_HELPER%" echo $lines = Get-Content -LiteralPath $bat -Encoding UTF8
        >> "%PS_HELPER%" echo $copy = $false
        >> "%PS_HELPER%" echo $sb = [System.Text.StringBuilder]::new()
        >> "%PS_HELPER%" echo foreach ($line in $lines) {
        >> "%PS_HELPER%" echo     if ($line -eq '###PS1_START###') { $copy = $true; continue }
        >> "%PS_HELPER%" echo     if ($line -eq '###PS1_END###') { $copy = $false; break }
        >> "%PS_HELPER%" echo     if ($copy) { [void]$sb.Append($line + [Environment]::NewLine) }
        >> "%PS_HELPER%" echo }
        >> "%PS_HELPER%" echo [System.IO.File]::WriteAllText($out, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))
        powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_HELPER%" "%~f0" "%~dp0%PS_BUILD%"
        set "_PS_EC=%ERRORLEVEL%"
        del /F /Q "%PS_HELPER%" >nul 2>&1
        if %_PS_EC% neq 0 (
            echo Could not write the build script out of build.bat. Exit code %_PS_EC%.
            set "EXITCODE=%_PS_EC%"
            goto CLEANUP
        )

REM -- Run the build

        powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_BUILD%"
        set "_PS_EC=%ERRORLEVEL%"
        del /F /Q "%PS_BUILD%" >nul 2>&1
        if %_PS_EC% neq 0 (
            echo Build did not complete. Read the message above. Exit code %_PS_EC%.
            set "EXITCODE=%_PS_EC%"
            goto CLEANUP
        )

        goto CLEANUP

    :CLEANUP

        if exist "%PS_HELPER%" del /F /Q "%PS_HELPER%" >nul 2>&1
        if exist "%PS_BUILD%" del /F /Q "%PS_BUILD%" >nul 2>&1
        pause
        endlocal & exit /b %EXITCODE%

###PS1_START###
# OpenTools for Microsoft Office build script. Copyright (c) 2026 Fred Thompson. MIT license.
# SPDX-License-Identifier: MIT
# Written out and run by build.bat. Exit codes:
#   0 success
#   1 source check failed
#   2 signing certificate problem
#   3 download or checksum failure
#   4 compile failure
#   5 installer build failure
#   6 temporary files could not be removed
#   7 unexpected error
Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
$work = Join-Path ([IO.Path]::GetTempPath()) ('OpenTools-build-' + [Guid]::NewGuid().ToString('N'))
$cache = Join-Path ([IO.Path]::GetTempPath()) 'OpenTools-build-cache'
$src = Join-Path $root 'src'
$helperSource = Join-Path $root 'installer\TrustPublisher.cs'
$signing = Join-Path $root 'Signing'
$stage = Join-Path $work 'stage'
$productVersion = '1.0.0'
# Windows Installer replaces an installed file only when the new one has a higher version. Every DLL used to be
# 1.0.0.0, so a leftover copy was kept. Build part: two-digit year * 1000 + day of year. Revision part: seconds
# since midnight / 2. Both stay under the 65535 limit, and each later build is higher.
$buildTime = Get-Date
$fileVersion = '1.0.' + (($buildTime.Year % 100) * 1000 + $buildTime.DayOfYear) + '.' +
    [int][Math]::Floor($buildTime.TimeOfDay.TotalSeconds / 2)
$cert = $null
$imported = @()
$rc = 7

# ----- Pinned downloads -----
# Every NuGet package is pinned to an exact version. Sha256 is the hash of the .nupkg file.
# InnerPath and InnerSha256 pin one file inside the package as well. Empty means not pinned.
$packages = @(
    @{
        Id = 'Microsoft.Office.Interop.Word'; Version = '15.0.4797.1004'
        Sha256 = 'B3026AD69B9165086B016B145767F8C7688916FCDD4CC6A9BE301123051E66B4'
        InnerPath = 'lib\net20\Microsoft.Office.Interop.Word.dll'
        InnerSha256 = 'DAD6C302A2025BAF6F96AE12825A0FB7ABAAD0F9CD4AE84F374899E8E729D97D'
    },
    @{
        Id = 'Microsoft.Office.Interop.Excel'; Version = '15.0.4795.1001'
        Sha256 = '4AC556088F5440B64E241F2378D92C5933DDBD8A72D27AC3C69C053A48E4782C'
        InnerPath = 'lib\net20\Microsoft.Office.Interop.Excel.dll'
        InnerSha256 = 'D5FBF3F71C40CA63B27601B5275C1CF5DC0CFD187C972579E2100F9215A375FE'
    },
    @{
        Id = 'Microsoft.Office.Interop.PowerPoint'; Version = '15.0.4420.1018'
        Sha256 = '47AE1E3BD8EF6BE5560701BB3C88F58C571DB0130F7E26CD138DB746EAEDE253'
        InnerPath = 'lib\net20\Microsoft.Office.Interop.PowerPoint.dll'
        InnerSha256 = 'E252DB17CFD1BB58EEA20B5D8CA2B676236C8B015985B36398C569E5A67F0F0B'
    },
    @{
        Id = 'MSOfficeCore.Interop'; Version = '15.0.2'
        Sha256 = ''
        InnerPath = 'lib\net48\Office.dll'
        InnerSha256 = 'DC1C9337435FA37201DBB8C012E0397E0A1BAE7273305CA397FEED566BA0F9E9'
    },
    @{
        Id = 'Microsoft.NETFramework.ReferenceAssemblies.net48'; Version = '1.0.3'
        Sha256 = '8A7E348538E7EB91351696911689F49E3D4F63F8BAB517432BBE159B8B1104A2'
        InnerPath = ''
        InnerSha256 = ''
    }
)
$wixUrl = 'https://github.com/wixtoolset/wix3/releases/download/wix3141rtm/wix314-binaries.zip'
$wixSha256 = '6AC824E1642D6F7277D0ED7EA09411A508F6116BA6FAE0AA5F2C7DAA2FF43D31'

# ----- Source files that must exist -----
$expectedSources = @(
    'Com\IDTExtensibility2.cs', 'Com\TabPaneControl.cs',
    'Core\AppProfile.cs', 'Core\CloseAllDialog.cs',
    'Core\EmbeddedBarHost.cs', 'Core\FrameAnimation.cs', 'Core\IAppAdapter.cs', 'Core\ITabBarHost.cs',
    'Core\Logger.cs',
    'Core\NativeMethods.cs', 'Core\NewWindowPlacer.cs', 'Core\RibbonTabSync.cs', 'Core\SettingNames.cs',
    'Core\InstalledApps.cs',
    'Core\Settings.cs', 'Core\SettingsWindow.cs', 'Core\TabBarControl.cs',
    'Core\TabBarLayout.cs',
    'Core\TabBarPosition.cs', 'Core\TabCoordinator.cs', 'Core\TabItem.cs', 'Core\TabLabeler.cs', 'Core\TabList.cs',
    'Core\TabOverlayWindow.cs', 'Core\TaskPaneHost.cs', 'Core\WindowStacker.cs',
    'Excel\ExcelAdapter.cs', 'Lorem\LoremIcon.cs', 'Lorem\LoremInserter.cs',
    'Lorem\LoremText.cs',
    'PowerPoint\PowerPointAdapter.cs', 'Project\ProjectAdapter.cs',
    'Visio\VisioAdapter.cs',
    'Shell\AssemblyInfo.cs', 'Shell\Connect.cs',
    'Toolbars\GearIcon.cs', 'Toolbars\MenuCaptions.cs', 'Toolbars\RibbonPicture.cs',
    'Toolbars\StartTabWatcher.cs', 'Toolbars\ToolbarFile.cs', 'Toolbars\ToolbarMenu.cs',
    'Toolbars\ToolbarRibbonXml.cs', 'Toolbars\ToolbarSet.cs',
    'Toolbars\Excel.commands', 'Toolbars\PowerPoint.commands', 'Toolbars\Word.commands',
    'Word\WordAdapter.cs'
)

# ----- Helper functions -----
function Fail([int]$code, [string]$message) {
    $script:rc = $code
    throw $message
}

function Run([string]$exe, [string[]]$arguments) {
    & $exe @arguments | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw ((Split-Path $exe -Leaf) + ' failed with exit code ' + $LASTEXITCODE)
    }
}

function WriteFile([string]$path, [string]$text) {
    [IO.File]::WriteAllText($path, $text, [Text.UTF8Encoding]::new($false))
}

function Download([string]$url, [string]$path) {
    $wc = [Net.WebClient]::new()
    try {
        $wc.DownloadFile($url, $path)
    } finally {
        $wc.Dispose()
    }
}

function CheckHash([string]$path, [string]$expected, [string]$algorithm, [string]$what) {
    $actual = (Get-FileHash -LiteralPath $path -Algorithm $algorithm).Hash
    if ($actual -ine $expected) {
        Fail 3 ('Checksum mismatch for ' + $what + '. Expected ' + $expected + ' but the download is ' + $actual + '.')
    }
}

# ----- Step 1: check the source files -----
function CheckCsFile([IO.FileInfo]$f) {
    $bytes = [IO.File]::ReadAllBytes($f.FullName)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        Fail 1 ($f.Name + ' starts with a byte order mark. It must be UTF-8 without one.')
    }
    foreach ($b in $bytes) {
        if ($b -gt 127) { Fail 1 ($f.Name + ' contains a character that is not ASCII.') }
        if ($b -eq 9) { Fail 1 ($f.Name + ' contains a tab character.') }
    }
    $text = [Text.Encoding]::ASCII.GetString($bytes)
    if ([regex]::IsMatch($text, '(?<!\r)\n')) {
        Fail 1 ($f.Name + ' has a line ending that is not CRLF.')
    }
    if ([regex]::IsMatch($text, '(?m)^\s*#region')) {
        Fail 1 ($f.Name + ' contains a #region block.')
    }
}

function CheckSources {
    foreach ($rel in $expectedSources) {
        if (-not (Test-Path -LiteralPath (Join-Path $src $rel))) {
            Fail 1 ('Missing source file: src\' + $rel)
        }
    }
    if (-not (Test-Path -LiteralPath $helperSource)) {
        Fail 1 'Missing source file: installer\TrustPublisher.cs'
    }
    foreach ($f in @(Get-ChildItem -LiteralPath $src -Recurse -Filter '*.cs')) {
        CheckCsFile $f
    }
    foreach ($f in @(Get-ChildItem -LiteralPath $src -Recurse -Filter '*.commands')) {
        CheckCsFile $f
    }
    CheckCsFile (Get-Item -LiteralPath $helperSource)
    Write-Host ('Source check passed for ' + ($expectedSources.Count + 1) + ' required files.')
}

# ----- Step 2: choose the signing certificate -----
# Revision 16: unambiguous store-key reuse; prompt once only when necessary.
# Do not discard a signing identity merely because the saved PFX password cannot be decrypted.
function IsValidSigningCertificate($candidate) {
    if ($null -eq $candidate -or -not $candidate.HasPrivateKey) { return $false }
    if ($candidate.PublicKey.Oid.Value -ne '1.2.840.113549.1.1.1') { return $false }
    if ($candidate.NotBefore -gt (Get-Date) -or $candidate.NotAfter -le (Get-Date)) { return $false }
    foreach ($extension in $candidate.Extensions) {
        if ($extension.Oid.Value -eq '2.5.29.37') {
            $eku = [Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]::new($extension, $false)
            foreach ($oid in $eku.EnhancedKeyUsages) {
                if ($oid.Value -eq '1.3.6.1.5.5.7.3.3') { return $true }
            }
        }
    }
    return $false
}

function GetSigningCertificate {
    [void](New-Item -ItemType Directory -Path $signing -Force)
    $pfx = Join-Path $signing 'CodeSigning.pfx'
    $secret = Join-Path $signing 'CodeSigning.password.dpapi'
    $thumbName = 'CertificateThumbprint_' + $env:COMPUTERNAME + '.txt'
    $thumbFile = Join-Path $signing $thumbName
    $pfxPresent = Test-Path -LiteralPath $pfx
    $thumb = ''
    $found = $null
    $candidates = @()

    # The thumbprint is a durable public identity selector, NOT an encrypted password.
    # Both the PFX and thumbprint are allowed: first reuse the key already in the store.
    if (Test-Path -LiteralPath $thumbFile) {
        $thumb = (Get-Content -LiteralPath $thumbFile -Raw).Trim().Replace(' ', '').Replace('-', '')
        if ($thumb -notmatch '^[a-fA-F0-9]{40}$') { Fail 2 ('Invalid ' + $thumbName + '.') }
        $keyPath = 'Cert:\CurrentUser\My\' + $thumb
        if (Test-Path -LiteralPath $keyPath) {
            $candidate = Get-Item -LiteralPath $keyPath
            if (IsValidSigningCertificate $candidate) {
                $found = $candidate
                Write-Host ('Reusing existing Windows private key: ' + $thumb)
            } else {
                Write-Warning 'The selected Windows certificate exists but is not currently a usable signing key.'
            }
        }
        if ($null -eq $found) {
            # The recorded key is not in this store. Recover this computer's own key from the store below, the
            # same way a build with no record does, rather than reaching for the PFX first.
            Write-Host ('The recorded key ' + $thumb + ' is not usable on this computer. Looking for this ' +
                'computer''s OpenTools key in its certificate store.')
            $thumb = ''
        }
    }

    # No key is recorded for this computer. Look for this computer's own OpenTools key in its store before
    # touching the optional DPAPI password file.
    if ($null -eq $found -and $thumb -eq '') {
        $candidates = @(Get-ChildItem Cert:\CurrentUser\My | Where-Object {
            $_.Subject -eq 'CN=OpenTools Local Publisher' -and
            (IsValidSigningCertificate $_)
        })
        if ($candidates.Count -gt 1) {
            # Existing installer signatures disambiguate the original identity without PFX decryption.
            $previous = @(Get-ChildItem -LiteralPath $root -Filter '*.msi' -File | Where-Object {
                $_.Name -like 'OpenTools_for_Microsoft_Office_*.msi'
            } | Sort-Object LastWriteTime -Descending)
            foreach ($msi in $previous) {
                try {
                    $signed = Get-AuthenticodeSignature -LiteralPath $msi.FullName
                    if ($null -eq $signed.SignerCertificate) { continue }
                    $matching = @($candidates | Where-Object {
                        $_.Thumbprint -eq $signed.SignerCertificate.Thumbprint
                    })
                    if ($matching.Count -eq 1) {
                        $found = $matching[0]
                        Write-Host ('Recovered signing identity from existing installer: ' + $msi.Name)
                        break
                    }
                } catch {
                    Write-Warning ('Could not inspect installer signature: ' + $msi.Name)
                }
            }
            # A previously installed add-in carries the certificate used to sign it.
            # This is a stronger identity record than choosing the newest certificate.
            if ($null -eq $found) {
                $installedDll = Join-Path $env:LOCALAPPDATA 'Programs\OpenTools\OpenTools.dll'
                if (Test-Path -LiteralPath $installedDll) {
                    try {
                        $signedDll = Get-AuthenticodeSignature -LiteralPath $installedDll -ErrorAction Stop
                        if ($signedDll.Status -in @('Valid', 'NotTrusted') -and
                            $null -ne $signedDll.SignerCertificate) {
                            $matching = @($candidates | Where-Object {
                                $_.Thumbprint -eq $signedDll.SignerCertificate.Thumbprint
                            })
                            if ($matching.Count -eq 1) {
                                $found = $matching[0]
                                Write-Host 'Recovered signing identity from the installed OpenTools DLL.'
                            }
                        }
                    } catch {
                        Write-Warning 'Could not read the installed OpenTools DLL signing identity.'
                    }
                }
            }
            # Two usable private keys are present. Never guess, never fall back to a
            # machine-bound PFX password when the store keys can be selected directly.
            if ($null -eq $found) {
                $candidates = @($candidates | Sort-Object Thumbprint)
                Write-Host 'Multiple local OpenTools signing keys are installed.'
                Write-Host 'Choose the certificate to keep using on THIS computer:'
                for ($i = 0; $i -lt $candidates.Count; $i++) {
                    $item = $candidates[$i]
                    $line = '  ' + ($i + 1) + ') ' + $item.Thumbprint
                    $line += '  expires ' + $item.NotAfter.ToString('yyyy-MM-dd')
                    Write-Host $line
                }
                while ($null -eq $found) {
                    $choice = (Read-Host 'Certificate number (Q cancels without changes)').Trim()
                    if ($choice -ieq 'Q') {
                        Fail 2 'Certificate selection cancelled. No signing identity was changed.'
                    }
                    if ($choice -match '^[1-9][0-9]*$') {
                        $index = [int]$choice - 1
                        if ($index -ge 0 -and $index -lt $candidates.Count) {
                            $found = $candidates[$index]
                            Write-Host ('Selected existing certificate: ' + $found.Thumbprint)
                        }
                    }
                    if ($null -eq $found) { Write-Host 'Enter a listed number, or Q to cancel.' }
                }
            }
        } elseif ($candidates.Count -eq 1) {
            $found = $candidates[0]
            Write-Host 'Reusing the sole existing OpenTools private key in CurrentUser\My.'
        }
        if ($null -ne $found) {
            $thumb = $found.Thumbprint
            # Record the recovered identity. This is public data, not signing material.
            WriteFile $thumbFile ($thumb + [Environment]::NewLine)
            Write-Host ('Recovered certificate thumbprint saved: ' + $thumb)
        }
    }

    if ($null -eq $found -and $pfxPresent) {
        # A PFX is needed only if the local store does not already have the selected key.
        # An inaccessible DPAPI file must never be fatal to an otherwise usable local key.
        $password = $null
        if (Test-Path -LiteralPath $secret) {
            try {
                $saved = (Get-Content -LiteralPath $secret -Raw).Trim()
                $password = ConvertTo-SecureString -String $saved -ErrorAction Stop
            } catch {
                $message = 'Saved PFX password is inaccessible in this Windows security context: '
                $message += $_.Exception.Message
                Write-Warning $message
            }
        }
        if ($null -eq $password) {
            Write-Host 'The existing PFX and signing files will NOT be deleted or regenerated.'
            $password = Read-Host 'Enter the original CodeSigning.pfx password if known' -AsSecureString
        }
        try {
            $before = @(Get-ChildItem Cert:\CurrentUser\My | ForEach-Object { $_.Thumbprint })
            $loaded = @(Import-PfxCertificate -FilePath $pfx `
                -CertStoreLocation Cert:\CurrentUser\My -Password $password -ErrorAction Stop)
            $keys = @($loaded | Where-Object { $_.HasPrivateKey })
            if ($keys.Count -ne 1) { Fail 2 'The existing PFX does not contain exactly one signing private key.' }
            if ($thumb -ne '' -and $keys[0].Thumbprint -ine $thumb) {
                Fail 2 ('Existing PFX conflicts with ' + $thumbName + '. Neither identity was replaced.')
            }
            if (-not (IsValidSigningCertificate $keys[0])) {
                Fail 2 'PFX contains no currently usable RSA code-signing key.'
            }
            $found = $keys[0]
            # Keep the imported signing key in CurrentUser\My so the next build can reuse it.
            $thumb = $found.Thumbprint
            WriteFile $thumbFile ($thumb + [Environment]::NewLine)
            Write-Host ('Imported existing PFX; private key will remain installed: ' + $thumb)
        } catch {
            if ($script:rc -eq 2) { throw }
            $message = 'Existing PFX could not be unlocked or imported: ' + $_.Exception.Message
            $message += '. No key or signing file was deleted. If this PFX was generated with a '
            $message += 'random password whose DPAPI cache is no longer decryptable, recover the '
            $message += 'original private key from its Windows certificate store or backup.'
            Fail 2 $message
        }
    }

    if ($null -eq $found) {
        $unknown = @(Get-ChildItem -LiteralPath $signing -Force | Where-Object {
            $_.Name -notin @('readme.txt', '.gitkeep')
        })
        if ($unknown.Count) {
            $message = 'Existing Signing material is present but does not identify an accessible private key. '
            $message += 'No key was replaced. '
            if ($candidates.Count -gt 1) {
                $message += 'Multiple OpenTools signing certificates exist in CurrentUser\My; '
                $message += 'select the intended certificate in ' + $thumbName + '. '
                $message += 'Available thumbprints: ' + (($candidates | ForEach-Object { $_.Thumbprint }) -join ', ')
            }
            Fail 2 $message
        }
        if ($candidates.Count -gt 1) {
            $message = 'Multiple OpenTools signing keys exist. Select the intended thumbprint '
            $message += 'using ' + $thumbName + '; no new key was generated.'
            Fail 2 $message
        }
        Write-Host 'No existing signing identity found; generating the initial signing key.'
        $found = New-SelfSignedCertificate -Type CodeSigningCert `
            -Subject 'CN=OpenTools Local Publisher' `
            -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy Exportable `
            -KeyLength 3072 -HashAlgorithm SHA256 -NotAfter (Get-Date).AddYears(3)
        $random = New-Object byte[] 48
        $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
        try { $rng.GetBytes($random) } finally { $rng.Dispose() }
        $password = ConvertTo-SecureString ([Convert]::ToBase64String($random)) -AsPlainText -Force
        [Array]::Clear($random, 0, $random.Length)
        Export-PfxCertificate -Cert $found -FilePath $pfx -Password $password -ChainOption EndEntityCertOnly | Out-Null
        WriteFile $thumbFile ($found.Thumbprint + [Environment]::NewLine)
        try {
            WriteFile $secret (ConvertFrom-SecureString $password)
        } catch {
            Write-Warning 'Optional DPAPI password cache could not be written. Store key will be reused.'
        }
    }
    if (-not (IsValidSigningCertificate $found)) { Fail 2 'No usable RSA code-signing key was selected.' }
    Write-Host ('Signing publisher: ' + $found.Subject + '; thumbprint: ' + $found.Thumbprint)
    return $found
}

# ----- Step 3: download the build tools -----
# Downloads are saved in the cache folder and reused. A saved copy is checked against the same pinned checksums
# before it is used, and is downloaded again if the check fails. Extraction goes to a .partial folder first, so a
# stopped build cannot leave a half-extracted tool that looks complete.
function ExtractFresh([string]$zip, [string]$dir) {
    $partial = $dir + '.partial'
    if (Test-Path -LiteralPath $partial) { Remove-Item -LiteralPath $partial -Recurse -Force }
    if (Test-Path -LiteralPath $dir) { Remove-Item -LiteralPath $dir -Recurse -Force }
    [IO.Compression.ZipFile]::ExtractToDirectory($zip, $partial)
    Move-Item -LiteralPath $partial -Destination $dir
}

function PackageIsGood([hashtable]$pin, [string]$zip, [string]$dir) {
    if (-not (Test-Path -LiteralPath $zip)) { return $false }
    if (-not (Test-Path -LiteralPath $dir)) { return $false }
    if ($pin.Sha256 -ne '') {
        if ((Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash -ine $pin.Sha256) { return $false }
    }
    if ($pin.InnerPath -ne '') {
        $inner = Join-Path $dir $pin.InnerPath
        if (-not (Test-Path -LiteralPath $inner)) { return $false }
        if ((Get-FileHash -LiteralPath $inner -Algorithm SHA256).Hash -ine $pin.InnerSha256) { return $false }
    }
    return $true
}

function GetPackage([hashtable]$pin) {
    $name = $pin.Id + '.' + $pin.Version
    $zip = Join-Path $cache ($name + '.nupkg')
    $dir = Join-Path $cache ('pkg-' + $name)
    if (PackageIsGood $pin $zip $dir) {
        Write-Host ('Using the saved copy of ' + $name)
        return $dir
    }
    Write-Host ('Downloading ' + $name)
    if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
    Download ('https://www.nuget.org/api/v2/package/' + $pin.Id + '/' + $pin.Version) $zip
    if ($pin.Sha256 -ne '') {
        CheckHash $zip $pin.Sha256 'SHA256' ($name + '.nupkg')
    }
    ExtractFresh $zip $dir
    if ($pin.InnerPath -ne '') {
        $inner = Join-Path $dir $pin.InnerPath
        if (-not (Test-Path -LiteralPath $inner)) {
            Fail 3 ('The package ' + $name + ' does not contain ' + $pin.InnerPath)
        }
        CheckHash $inner $pin.InnerSha256 'SHA256' ($name + ' ' + $pin.InnerPath)
    }
    return $dir
}

# The SDK checksum comes from Microsoft's release information, so it is checked when the SDK is downloaded.
# A marker file records that check. A later build that finds the SDK and the marker reuses the SDK.
function GetDotNetSdk {
    $sdk = Join-Path $cache 'dotnet'
    $marker = Join-Path $cache 'dotnet.verified.txt'
    if ((Test-Path -LiteralPath (Join-Path $sdk 'dotnet.exe')) -and (Test-Path -LiteralPath $marker)) {
        Write-Host ('Using the saved .NET SDK: ' + (Get-Content -LiteralPath $marker -Raw).Trim())
        return $sdk
    }
    $wc = [Net.WebClient]::new()
    try {
        $meta = $wc.DownloadString('https://dotnetcli.blob.core.windows.net/dotnet/release-metadata/8.0/releases.json') | ConvertFrom-Json
    } finally {
        $wc.Dispose()
    }
    $version = [string]$meta.'latest-sdk'
    $rid = if ([Environment]::Is64BitOperatingSystem) { 'win-x64' } else { 'win-x86' }
    $release = @($meta.releases | Where-Object { $_.PSObject.Properties['sdk'] -and $_.sdk.version -eq $version })
    $artifact = @($release[0].sdk.files | Where-Object { $_.rid -eq $rid -and $_.name -eq ('dotnet-sdk-' + $rid + '.zip') })
    if ($artifact.Count -ne 1 -or $artifact[0].hash -notmatch '^[a-fA-F0-9]{128}$') {
        Fail 3 'The Microsoft .NET SDK download information is not valid.'
    }
    Write-Host ('Downloading the .NET SDK ' + $version + ' (used only to run the C# compiler)')
    $sdkZip = Join-Path $cache ('dotnet-sdk-' + $version + '-' + $rid + '.zip')
    if (Test-Path -LiteralPath $marker) { Remove-Item -LiteralPath $marker -Force }
    if (Test-Path -LiteralPath $sdkZip) { Remove-Item -LiteralPath $sdkZip -Force }
    Download $artifact[0].url $sdkZip
    CheckHash $sdkZip $artifact[0].hash 'SHA512' 'the .NET SDK'
    ExtractFresh $sdkZip $sdk
    WriteFile $marker ('SDK ' + $version + ' ' + $rid + ', SHA512 ' + $artifact[0].hash)
    return $sdk
}

function GetWix {
    $zip = Join-Path $cache 'wix314-binaries.zip'
    $dir = Join-Path $cache 'wix'
    $saved = (Test-Path -LiteralPath $zip) -and (Test-Path -LiteralPath (Join-Path $dir 'light.exe'))
    if ($saved) {
        $saved = ((Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash -ieq $wixSha256)
    }
    if ($saved) {
        Write-Host 'Using the saved copy of WiX 3.14'
        return $dir
    }
    Write-Host 'Downloading WiX 3.14'
    if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
    Download $wixUrl $zip
    CheckHash $zip $wixSha256 'SHA256' 'WiX 3.14'
    ExtractFresh $zip $dir
    return $dir
}

# ----- Step 4: compile the add-in and the installer helper -----
function CompileAddIn([string]$sdk, [string]$wordDir, [string]$excelDir, [string]$pptDir, [string]$officeDir, [string]$refsDir, [string]$outDll) {
    $dotnet = Join-Path $sdk 'dotnet.exe'
    $csc = @(Get-ChildItem -LiteralPath (Join-Path $sdk 'sdk') -Recurse -Filter 'csc.dll' | Where-Object { $_.FullName -like '*Roslyn\bincore\csc.dll' })
    if ($csc.Count -lt 1) { Fail 4 'The C# compiler was not found in the .NET SDK.' }
    $refs = Join-Path $refsDir 'build\.NETFramework\v4.8'
    if (-not (Test-Path -LiteralPath (Join-Path $refs 'mscorlib.dll'))) {
        Fail 4 'The .NET Framework 4.8 reference files were not found where expected (build\.NETFramework\v4.8).'
    }
    $lines = New-Object Collections.Generic.List[string]
    $lines.Add('/nologo')
    $lines.Add('/noconfig')
    $lines.Add('/nostdlib+')
    $lines.Add('/target:library')
    $lines.Add('/platform:anycpu')
    $lines.Add('/langversion:7.3')
    $lines.Add('/optimize+')
    $lines.Add('/debug-')
    $lines.Add('/warn:4')
    $lines.Add('/out:"' + $outDll + '"')
    foreach ($ref in @('mscorlib', 'System', 'System.Core', 'System.Drawing', 'System.Windows.Forms')) {
        $lines.Add('/reference:"' + (Join-Path $refs ($ref + '.dll')) + '"')
    }
    $lines.Add('/link:"' + (Join-Path $officeDir 'lib\net48\Office.dll') + '"')
    $lines.Add('/link:"' + (Join-Path $wordDir 'lib\net20\Microsoft.Office.Interop.Word.dll') + '"')
    $lines.Add('/link:"' + (Join-Path $excelDir 'lib\net20\Microsoft.Office.Interop.Excel.dll') + '"')
    $lines.Add('/link:"' + (Join-Path $pptDir 'lib\net20\Microsoft.Office.Interop.PowerPoint.dll') + '"')
    foreach ($f in @(Get-ChildItem -LiteralPath $src -Recurse -Filter '*.cs' | Sort-Object FullName)) {
        $lines.Add('"' + $f.FullName + '"')
    }
    # Each OpenToolbars command list goes into the DLL compressed, as OpenTools.<App>.commands.gz.
    foreach ($f in @(Get-ChildItem -LiteralPath $src -Recurse -Filter '*.commands' | Sort-Object FullName)) {
        $packed = Join-Path $work ($f.Name + '.gz')
        $plain = [IO.File]::OpenRead($f.FullName)
        try {
            $output = [IO.File]::Create($packed)
            try {
                $zip = New-Object IO.Compression.GZipStream($output, [IO.Compression.CompressionLevel]::Optimal)
                try { $plain.CopyTo($zip) } finally { $zip.Dispose() }
            } finally { $output.Dispose() }
        } finally { $plain.Dispose() }
        $lines.Add('/resource:"' + $packed + '",OpenTools.' + $f.Name + '.gz')
    }
    $versionSource = Join-Path $work 'BuildFileVersion.cs'
    WriteFile $versionSource ('[assembly: System.Reflection.AssemblyFileVersion("' + $fileVersion + '")]' +
        [Environment]::NewLine)
    $lines.Add('"' + $versionSource + '"')
    $rsp = Join-Path $work 'compile.rsp'
    WriteFile $rsp ($lines -join [Environment]::NewLine)
    $env:DOTNET_ROOT = $sdk
    $env:DOTNET_CLI_HOME = Join-Path $work 'dotnet-home'
    $env:DOTNET_NOLOGO = 'true'
    $env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
    Write-Host ('Compiling OpenTools.dll, file version ' + $fileVersion)
    & $dotnet $csc[0].FullName ('@' + $rsp) | Out-Host
    if ($LASTEXITCODE -ne 0) { Fail 4 ('The C# compiler reported errors. Exit code ' + $LASTEXITCODE + '.') }
    if (-not (Test-Path -LiteralPath $outDll)) { Fail 4 'The compiler finished but produced no DLL.' }
}

# The helper is a Windows application (winexe), so it has no console window. The installer runs it to add or
# remove the publisher certificate.
function CompileTrustHelper([string]$sdk, [string]$refsDir, [string]$outExe) {
    $dotnet = Join-Path $sdk 'dotnet.exe'
    $csc = @(Get-ChildItem -LiteralPath (Join-Path $sdk 'sdk') -Recurse -Filter 'csc.dll' | Where-Object { $_.FullName -like '*Roslyn\bincore\csc.dll' })
    if ($csc.Count -lt 1) { Fail 4 'The C# compiler was not found in the .NET SDK.' }
    $refs = Join-Path $refsDir 'build\.NETFramework\v4.8'
    if (-not (Test-Path -LiteralPath (Join-Path $refs 'mscorlib.dll'))) {
        Fail 4 'The .NET Framework 4.8 reference files were not found where expected (build\.NETFramework\v4.8).'
    }
    $lines = New-Object Collections.Generic.List[string]
    $lines.Add('/nologo')
    $lines.Add('/noconfig')
    $lines.Add('/nostdlib+')
    $lines.Add('/target:winexe')
    $lines.Add('/platform:anycpu')
    $lines.Add('/langversion:7.3')
    $lines.Add('/optimize+')
    $lines.Add('/debug-')
    $lines.Add('/warn:4')
    $lines.Add('/out:"' + $outExe + '"')
    foreach ($ref in @('mscorlib', 'System', 'System.Core')) {
        $lines.Add('/reference:"' + (Join-Path $refs ($ref + '.dll')) + '"')
    }
    $lines.Add('"' + $helperSource + '"')
    $rsp = Join-Path $work 'helper.rsp'
    WriteFile $rsp ($lines -join [Environment]::NewLine)
    $env:DOTNET_ROOT = $sdk
    $env:DOTNET_CLI_HOME = Join-Path $work 'dotnet-home'
    $env:DOTNET_NOLOGO = 'true'
    $env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
    Write-Host 'Compiling TrustPublisher.exe'
    & $dotnet $csc[0].FullName ('@' + $rsp) | Out-Host
    if ($LASTEXITCODE -ne 0) { Fail 4 ('The C# compiler reported errors in the installer helper. Exit code ' + $LASTEXITCODE + '.') }
    if (-not (Test-Path -LiteralPath $outExe)) { Fail 4 'The compiler finished but produced no installer helper.' }
}

# ----- Step 5: sign a file -----
function SignFile([string]$path) {
    $result = Set-AuthenticodeSignature -LiteralPath $path -Certificate $cert -HashAlgorithm SHA256 -IncludeChain NotRoot
    $okStatus = @('Valid', 'NotTrusted', 'UnknownError')
    if ($null -eq $result.SignerCertificate -or $result.SignerCertificate.Thumbprint -ne $cert.Thumbprint -or $result.Status -notin $okStatus) {
        Fail 2 ('Signing failed for ' + (Split-Path $path -Leaf) + ': ' + $result.StatusMessage)
    }
}

# ----- Step 6: registry entries for the two COM classes -----
# Each class is registered in the 32-bit and the 64-bit registry view, under the current user only.
# A 32-bit Office reads the 32-bit view and a 64-bit Office reads the 64-bit view.
function ClassRegistry([string]$clsid, [string]$progId, [string]$className, [bool]$isControl) {
    $assembly = 'OpenTools, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'
    $codeBase = 'file:///[INSTALLFOLDER]OpenTools.dll'
    $base = 'Software\Classes\CLSID\' + $clsid
    $x = New-Object Text.StringBuilder
    function Value([string]$key, [string]$name, [string]$value, [bool]$keyPath) {
        $text = '<RegistryValue Root="HKCU" Key="' + $key + '"'
        if ($name -ne '') { $text += ' Name="' + $name + '"' }
        $text += ' Type="string" Value="' + $value + '"'
        if ($keyPath) { $text += ' KeyPath="yes"' }
        [void]$x.AppendLine($text + '/>')
    }
    function EmptyKey([string]$key) {
        [void]$x.AppendLine('<RegistryKey Root="HKCU" Key="' + $key + '" Action="createAndRemoveOnUninstall"/>')
    }
    Value ('Software\Classes\' + $progId) '' $progId $true
    Value ('Software\Classes\' + $progId + '\CLSID') '' $clsid $false
    Value $base '' $progId $false
    Value ($base + '\ProgId') '' $progId $false
    EmptyKey ($base + '\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}')
    Value ($base + '\InprocServer32') '' 'mscoree.dll' $false
    foreach ($key in @(($base + '\InprocServer32'), ($base + '\InprocServer32\1.0.0.0'))) {
        if ($key.EndsWith('1.0.0.0')) {
            Value $key 'Class' $className $false
        } else {
            Value $key 'ThreadingModel' 'Both' $false
            Value $key 'Class' $className $false
        }
        Value $key 'Assembly' $assembly $false
        Value $key 'RuntimeVersion' 'v4.0.30319' $false
        Value $key 'CodeBase' $codeBase $false
    }
    if ($isControl) {
        EmptyKey ($base + '\Control')
        EmptyKey ($base + '\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}')
        Value ($base + '\MiscStatus') '' '0' $false
        Value ($base + '\MiscStatus\1') '' '131473' $false
        Value ($base + '\Version') '' '1.0' $false
    }
    return $x.ToString()
}

# ----- Step 7: installer source -----
# The publisher certificate steps run TrustPublisher.exe, the windowless helper, not PowerShell.
function GetWixSource([string]$thumbprint) {
    $connect64 = ClassRegistry '{65CA4A37-C8CA-40BA-A55B-01E1DD4559E0}' 'OpenTools.Connect' 'OpenTools.Connect' $false
    $pane64 = ClassRegistry '{D2E0EDDB-0766-4293-BCD1-CFD0623F7038}' 'OpenTools.TabPane' 'OpenTools.TabPaneControl' $true
    $transaction = [Guid]::NewGuid().ToString('N')
    $action = 'ExeCommand="&quot;[#F_TrustHelper]&quot; -C &quot;[#F_PublicCert]&quot; -E __THUMB__ -T __TXN__ -O '
    $wxs = @'
<?xml version="1.0" encoding="UTF-8"?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
<Product Id="*" Name="OpenTools for Microsoft Office" Language="1033" Version="__VERSION__" Manufacturer="Fred Thompson"
UpgradeCode="{6F5FA55A-7947-40C7-A32B-24B08F9AB51A}">
<Package InstallerVersion="500" Platform="x64" Compressed="yes" InstallScope="perUser" InstallPrivileges="limited"
Description="Adds a tab bar to Microsoft Office desktop apps. MIT license." />
<MajorUpgrade Schedule="afterInstallInitialize" AllowSameVersionUpgrades="yes" DowngradeErrorMessage="A later version of OpenTools for Microsoft Office is already installed." />
<MediaTemplate EmbedCab="yes" />
<Condition Message="Rollback must be enabled to install or remove publisher trust safely.">NOT RollbackDisabled</Condition>
<Directory Id="TARGETDIR" Name="SourceDir">
<Directory Id="LocalAppDataFolder">
<Directory Id="OTDir" Name="Programs">
<Directory Id="INSTALLFOLDER" Name="OpenTools">
</Directory>
</Directory>
</Directory>
</Directory>
<DirectoryRef Id="INSTALLFOLDER">
<Component Id="C_Dll" Guid="{3555E911-B544-49BE-840B-D43CC8E53848}"><File Id="F_Dll" Name="OpenTools.dll" Source="OpenTools.dll"/><RegistryValue Root="HKCU" Key="Software\OpenTools\Installer\ComponentMarkers" Name="C_Dll" Type="integer" Value="1" KeyPath="yes"/></Component>
<Component Id="C_PublicCert" Guid="{15890F17-4882-48DB-9684-00422654AF85}"><File Id="F_PublicCert" Name="OpenTools_PublicCertificate.cer" Source="OpenTools_PublicCertificate.cer"/><RegistryValue Root="HKCU" Key="Software\OpenTools\Installer\ComponentMarkers" Name="C_PublicCert" Type="integer" Value="1" KeyPath="yes"/></Component>
<Component Id="C_TrustHelper" Guid="{89F250EC-B1FC-4E26-B4C8-877E6867B696}"><File Id="F_TrustHelper" Name="TrustPublisher.exe" Source="TrustPublisher.exe"/><RegistryValue Root="HKCU" Key="Software\OpenTools\Installer\ComponentMarkers" Name="C_TrustHelper" Type="integer" Value="1" KeyPath="yes"/></Component>
</DirectoryRef>
<DirectoryRef Id="TARGETDIR">
<Component Id="C_ConnectReg64" Guid="*" Win64="yes">__CONNECT__</Component>
<Component Id="C_ConnectReg32" Guid="*" Win64="no">__CONNECT__</Component>
<Component Id="C_PaneReg64" Guid="*" Win64="yes">__PANE__</Component>
<Component Id="C_PaneReg32" Guid="*" Win64="no">__PANE__</Component>
<Component Id="C_ExcelAddin" Guid="*" Win64="yes">
<RegistryValue Root="HKCU" Key="Software\Microsoft\Office\Excel\Addins\OpenTools.Connect" Name="FriendlyName" Type="string" Value="OpenTools for Microsoft Office" KeyPath="yes" />
<RegistryValue Root="HKCU" Key="Software\Microsoft\Office\Excel\Addins\OpenTools.Connect" Name="Description" Type="string" Value="Adds a tab bar to Excel." />
<RegistryValue Root="HKCU" Key="Software\Microsoft\Office\Excel\Addins\OpenTools.Connect" Name="LoadBehavior" Type="integer" Value="3" />
</Component>
<Component Id="C_WordAddin" Guid="*" Win64="yes">
<RemoveFolder Id="Cleanup_INSTALLFOLDER" Directory="INSTALLFOLDER" On="uninstall" />
<RemoveFolder Id="Cleanup_OTDir" Directory="OTDir" On="uninstall" />
<RegistryValue Root="HKCU" Key="Software\Microsoft\Office\Word\Addins\OpenTools.Connect" Name="FriendlyName" Type="string" Value="OpenTools for Microsoft Office" KeyPath="yes" />
<RegistryValue Root="HKCU" Key="Software\Microsoft\Office\Word\Addins\OpenTools.Connect" Name="Description" Type="string" Value="Adds a tab bar to Word." />
<RegistryValue Root="HKCU" Key="Software\Microsoft\Office\Word\Addins\OpenTools.Connect" Name="LoadBehavior" Type="integer" Value="3" />
</Component>
<Component Id="C_PowerPointAddin" Guid="*" Win64="yes">
<RegistryValue Root="HKCU" Key="Software\Microsoft\Office\PowerPoint\Addins\OpenTools.Connect" Name="FriendlyName" Type="string" Value="OpenTools for Microsoft Office" KeyPath="yes" />
<RegistryValue Root="HKCU" Key="Software\Microsoft\Office\PowerPoint\Addins\OpenTools.Connect" Name="Description" Type="string" Value="Adds a tab bar to PowerPoint." />
<RegistryValue Root="HKCU" Key="Software\Microsoft\Office\PowerPoint\Addins\OpenTools.Connect" Name="LoadBehavior" Type="integer" Value="3" />
</Component>
<Component Id="C_ProjectAddin" Guid="*" Win64="yes">
<RegistryValue Root="HKCU" Key="Software\Microsoft\Office\MS Project\Addins\OpenTools.Connect" Name="FriendlyName" Type="string" Value="OpenTools for Microsoft Office" KeyPath="yes" />
<RegistryValue Root="HKCU" Key="Software\Microsoft\Office\MS Project\Addins\OpenTools.Connect" Name="Description" Type="string" Value="Adds a tab bar to Project." />
<RegistryValue Root="HKCU" Key="Software\Microsoft\Office\MS Project\Addins\OpenTools.Connect" Name="LoadBehavior" Type="integer" Value="3" />
</Component>
<Component Id="C_VisioAddin" Guid="*" Win64="yes">
<RegistryValue Root="HKCU" Key="Software\Microsoft\Visio\Addins\OpenTools.Connect" Name="FriendlyName" Type="string" Value="OpenTools for Microsoft Office" KeyPath="yes" />
<RegistryValue Root="HKCU" Key="Software\Microsoft\Visio\Addins\OpenTools.Connect" Name="Description" Type="string" Value="Adds a tab bar to Visio." />
<RegistryValue Root="HKCU" Key="Software\Microsoft\Visio\Addins\OpenTools.Connect" Name="LoadBehavior" Type="integer" Value="3" />
</Component>
</DirectoryRef>
<Property Id="OT_CERT_CONSENT" Secure="yes" />
<WixVariable Id="WixUIDialogBmp" Value="dialog.bmp" />
<UI>
<UIRef Id="WixUI_Common" />
<DialogRef Id="ErrorDlg" />
<DialogRef Id="FatalError" />
<DialogRef Id="FilesInUse" />
<DialogRef Id="MsiRMFilesInUse" />
<DialogRef Id="UserExit" />
<DialogRef Id="ExitDialog" />
<DialogRef Id="ProgressDlg" />
<DialogRef Id="PrepareDlg" />
<DialogRef Id="ResumeDlg" />
<Publish Dialog="ExitDialog" Control="Finish" Event="EndDialog" Value="Return">1</Publish>
<TextStyle Id="WixUI_Font_Normal" FaceName="Tahoma" Size="8" />
<TextStyle Id="WixUI_Font_Bigger" FaceName="Tahoma" Size="12" />
<TextStyle Id="WixUI_Font_Title" FaceName="Tahoma" Size="9" Bold="yes" />
<TextStyle Id="OTFont" FaceName="Tahoma" Size="9" />
<Property Id="DefaultUIFont" Value="OTFont" />
<Dialog Id="PublisherConsentDlg" Width="370" Height="350" Title="OpenTools for Microsoft Office - Publisher Approval">
<Control Id="Title" Type="Text" X="18" Y="15" Width="334" Height="18" Text="{\WixUI_Font_Bigger}OpenTools for Microsoft Office 1.00" />
<Control Id="About1" Type="Text" X="18" Y="39" Width="334" Height="12" Text="Classic toolbars, lorem ipsum, and tab support" />
<Control Id="About2" Type="Text" X="18" Y="51" Width="334" Height="12" Text="for Microsoft Office 2016-2024" />
<Control Id="Copyright" Type="Text" X="18" Y="75" Width="334" Height="12" Text="(c) Fred Thompson 2026" />
<Control Id="License" Type="Text" X="18" Y="87" Width="334" Height="12" Text="MIT License" />
<Control Id="GitHub" Type="Hyperlink" X="18" Y="99" Width="334" Height="12" Text="&lt;a href=&quot;https://github.com/FredThompsonII/OpenTools_for_Microsoft_Office&quot;&gt;https://github.com/FredThompsonII/OpenTools_for_Microsoft_Office&lt;/a&gt;" />
<Control Id="Heading" Type="Text" X="18" Y="125" Width="334" Height="16" Text="Approve the OpenTools for Microsoft Office publisher certificate" />
<Control Id="Explanation" Type="Text" X="18" Y="145" Width="334" Height="40" Text="This installer is signed with a certificate made on the computer that built it. Certificate SHA-1 thumbprint: __THUMB__. The public certificate may not yet be trusted." />
<Control Id="Scope" Type="Text" X="18" Y="187" Width="334" Height="44" Text="If you approve, this certificate is added to your Windows user's Trusted Publishers store. A self-signed certificate is also added to your Root store." />
<Control Id="Risks" Type="Text" X="18" Y="233" Width="334" Height="26" Text="Trust affects other programs signed with the same certificate. Office security settings will NOT be disabled." />
<Control Id="Consent" Type="CheckBox" X="18" Y="263" Width="330" Height="22" Property="OT_CERT_CONSENT" CheckBoxValue="1" Text="I explicitly approve trusting this certificate." />
<Control Id="Cancel" Type="PushButton" X="235" Y="318" Width="56" Height="18" Cancel="yes" Text="Cancel">
<Publish Event="EndDialog" Value="Exit">1</Publish>
</Control>
<Control Id="Install" Type="PushButton" X="297" Y="318" Width="56" Height="18" Default="yes" Text="Install">
<Publish Event="EndDialog" Value="Return">OT_CERT_CONSENT="1"</Publish>
</Control>
</Dialog>
<InstallUISequence><Show Dialog="PublisherConsentDlg" After="CostFinalize">NOT Installed</Show></InstallUISequence>
</UI>
<CustomAction Id="RejectMissingApproval" Error="This installer requires explicit publisher certificate approval in its interactive dialog. Administrative deployments must deliberately set OT_CERT_CONSENT=1." />
<CustomAction Id="RollbackUserPublisherTrust" Directory="INSTALLFOLDER" __ACTION__Rollback" Execute="rollback" Impersonate="yes" Return="check" />
<CustomAction Id="ApplyUserPublisherTrust" Directory="INSTALLFOLDER" __ACTION__Install" Execute="deferred" Impersonate="yes" Return="check" />
<CustomAction Id="CommitUserPublisherTrust" Directory="INSTALLFOLDER" __ACTION__Commit" Execute="commit" Impersonate="yes" Return="ignore" />
<CustomAction Id="RollbackRemovedPublisherTrust" Directory="INSTALLFOLDER" __ACTION__Rollback" Execute="rollback" Impersonate="yes" Return="check" />
<CustomAction Id="RemoveUserPublisherTrust" Directory="INSTALLFOLDER" __ACTION__Uninstall" Execute="deferred" Impersonate="yes" Return="check" />
<CustomAction Id="RollbackRepairPublisherTrust" Directory="INSTALLFOLDER" __ACTION__Rollback" Execute="rollback" Impersonate="yes" Return="check" />
<CustomAction Id="RepairUserPublisherTrust" Directory="INSTALLFOLDER" __ACTION__Repair" Execute="deferred" Impersonate="yes" Return="check" />
<CustomAction Id="CleanupUserData" Directory="INSTALLFOLDER" __ACTION__Cleanup" Execute="deferred" Impersonate="yes" Return="ignore" />
<InstallExecuteSequence>
<Custom Action="RejectMissingApproval" Before="InstallInitialize">NOT Installed AND OT_CERT_CONSENT &lt;&gt; "1"</Custom>
<Custom Action="RollbackUserPublisherTrust" After="InstallFiles">NOT Installed AND OT_CERT_CONSENT = "1"</Custom>
<Custom Action="ApplyUserPublisherTrust" After="RollbackUserPublisherTrust">NOT Installed AND OT_CERT_CONSENT = "1"</Custom>
<Custom Action="RollbackRepairPublisherTrust" After="ApplyUserPublisherTrust">Installed AND REINSTALL AND NOT REMOVE</Custom>
<Custom Action="RepairUserPublisherTrust" After="RollbackRepairPublisherTrust">Installed AND REINSTALL AND NOT REMOVE</Custom>
<Custom Action="CommitUserPublisherTrust" After="RepairUserPublisherTrust">(NOT Installed AND OT_CERT_CONSENT = "1") OR (Installed AND REINSTALL AND NOT REMOVE)</Custom>
<Custom Action="RollbackRemovedPublisherTrust" Before="RemoveUserPublisherTrust">Installed AND REMOVE = "ALL"</Custom>
<Custom Action="RemoveUserPublisherTrust" Before="RemoveFiles">Installed AND REMOVE = "ALL"</Custom>
<Custom Action="CleanupUserData" After="RemoveUserPublisherTrust">REMOVE~="ALL" AND NOT UPGRADINGPRODUCTCODE</Custom>
</InstallExecuteSequence>
<Feature Id="Complete" Title="OpenTools for Microsoft Office" Level="1">
<ComponentRef Id="C_Dll" />
<ComponentRef Id="C_PublicCert" />
<ComponentRef Id="C_TrustHelper" />
<ComponentRef Id="C_ConnectReg64" />
<ComponentRef Id="C_ConnectReg32" />
<ComponentRef Id="C_PaneReg64" />
<ComponentRef Id="C_PaneReg32" />
<ComponentRef Id="C_WordAddin" />
<ComponentRef Id="C_ExcelAddin" />
<ComponentRef Id="C_PowerPointAddin" />
<ComponentRef Id="C_ProjectAddin" />
<ComponentRef Id="C_VisioAddin" />
</Feature>
</Product>
</Wix>
'@
    $wxs = $wxs.Replace('__CONNECT__', $connect64).Replace('__PANE__', $pane64).Replace('__ACTION__', $action)
    return $wxs.Replace('__THUMB__', $thumbprint).Replace('__TXN__', $transaction).Replace('__VERSION__', $productVersion)
}

# ----- Main flow -----
try {
    [void](New-Item -ItemType Directory -Path $stage -Force)
    [void](New-Item -ItemType Directory -Path $cache -Force)
    $env:TEMP = $work
    $env:TMP = $work
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    CheckSources
    $cert = GetSigningCertificate

    $dirs = @{}
    foreach ($pin in $packages) {
        $dirs[$pin.Id] = GetPackage $pin
    }
    $officeDll = Join-Path $dirs['MSOfficeCore.Interop'] 'lib\net48\Office.dll'
    $officeSignature = Get-AuthenticodeSignature -LiteralPath $officeDll
    if ($officeSignature.Status -ne 'Valid' -or $null -eq $officeSignature.SignerCertificate -or $officeSignature.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
        Fail 3 ('Office.dll is not validly signed by Microsoft Corporation. Status: ' + $officeSignature.Status)
    }
    Write-Host 'Office.dll is signed by Microsoft Corporation.'

    $sdk = GetDotNetSdk
    $wix = GetWix

    $dll = Join-Path $stage 'OpenTools.dll'
    CompileAddIn $sdk $dirs['Microsoft.Office.Interop.Word'] $dirs['Microsoft.Office.Interop.Excel'] $dirs['Microsoft.Office.Interop.PowerPoint'] $dirs['MSOfficeCore.Interop'] $dirs['Microsoft.NETFramework.ReferenceAssemblies.net48'] $dll
    SignFile $dll

    $helperExe = Join-Path $stage 'TrustPublisher.exe'
    CompileTrustHelper $sdk $dirs['Microsoft.NETFramework.ReferenceAssemblies.net48'] $helperExe
    SignFile $helperExe

    Export-Certificate -Cert $cert -FilePath (Join-Path $stage 'OpenTools_PublicCertificate.cer') | Out-Null
    WriteFile (Join-Path $stage 'installer.wxs') (GetWixSource $cert.Thumbprint)

    # A one-pixel white picture replaces WiX's 68 KB dialog artwork. Windows stretches it to fill the panel.
    [byte[]]$whitePixel = @(0x42, 0x4D, 0x3A, 0, 0, 0, 0, 0, 0, 0, 0x36, 0, 0, 0, 0x28, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0,
        1, 0, 0x18, 0, 0, 0, 0, 0, 4, 0, 0, 0, 0x13, 0x0B, 0, 0, 0x13, 0x0B, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0xFF, 0xFF, 0xFF, 0)
    [IO.File]::WriteAllBytes((Join-Path $stage 'dialog.bmp'), $whitePixel)

    $msi = Join-Path $work 'OpenTools.msi'
    Push-Location -LiteralPath $stage
    try {
        $script:rc = 5
        Run (Join-Path $wix 'candle.exe') @('-nologo', '-arch', 'x64', '-out', (Join-Path $work 'installer.wixobj'), 'installer.wxs')
        Run (Join-Path $wix 'light.exe') @('-nologo', '-ext', 'WixUIExtension', '-b', '.', '-out', $msi, (Join-Path $work 'installer.wixobj'))
    } finally {
        Pop-Location
    }
    SignFile $msi

    $version = [Version]$productVersion
    $output = Join-Path $root ('OpenTools_for_Microsoft_Office_' + $version.Major + '.' + $version.Minor.ToString('D2') +
        '.msi')
    Copy-Item -LiteralPath $msi -Destination $output -Force
    Write-Host ('Created: ' + $output)
    $rc = 0
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
} finally {
    # Keep imported signing identities in CurrentUser\My for future builds.
    # The working folder holds this build's staged files. The cache folder holds the downloads, including the
    # .NET SDK. Both are temporary and removed here.
    foreach ($temporary in @($work, $cache)) {
        try {
            if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
        } catch {
            Write-Warning ('Could not remove the temporary folder ' + $temporary + ': ' + $_.Exception.Message)
            if ($rc -eq 0) { $rc = 6 }
        }
    }
}
exit $rc
###PS1_END###
