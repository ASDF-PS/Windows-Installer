<#
.SYNOPSIS
    Restores missing Windows Installer cache files (.msi) for a broken server by
    matching them against a donor server's cache, and staging the matches RENAMED
    to the broken server's expected filenames.

.DESCRIPTION
    Companion to Test-MsiCache.ps1.

    MATCHING (important): ProductCode alone is NOT enough - Windows Installer
    keeps the SAME ProductCode across minor updates/patches, so the same code can
    mean different builds. This script therefore requires the BUILD to agree too:

      * If both sides have a PackageCode (the MSI's unique per-build Revision
        GUID), it must match exactly  -> guarantees an identical package.
      * Otherwise it falls back to DisplayVersion having to match.
      * -IgnoreVersion disables that guard (ProductCode only) - not recommended.

    The cache filename (e.g. 119918f.msi) is random per machine, so matched donor
    files are copied to the broken server's expected target filename into a
    STAGING folder you choose (never the live cache).

    Two ways to point at the donor's files:

    * SCAN MODE (no -LabCsv): you copied the donor's C:\Windows\Installer to a
      local folder (e.g. C:\_PKG\InstallerCache). The script opens every .msi in
      -SourcePath and reads its ProductCode / ProductVersion / PackageCode.
    * CSV MODE (-LabCsv): you have the donor's full audit CSV and reach its files
      over a share. (PackageCode matching works only if that CSV has a PackageCode
      column - i.e. produced by the updated Test-MsiCache.ps1.)

    Typical SCAN-MODE round trip:
      1. Broken server:  .\Test-MsiCache.ps1 -MissingOnly  -> Missing-Only.csv
      2. Copy a donor's C:\Windows\Installer -> C:\_PKG\InstallerCache (robocopy).
      3. .\Restore-MsiCache.ps1 -MissingCsv Missing-Only.csv `
             -SourcePath C:\_PKG\InstallerCache -DestinationPath C:\_PKG\Staging
         (dry run, then add -Execute)
      4. Copy C:\_PKG\Staging\*.msi back to the broken server's C:\Windows\Installer.
      5. Re-run .\Test-MsiCache.ps1 -MissingOnly to confirm.

    NOTE: PackageCode matching needs PackageCode in Missing-Only.csv. Re-run the
    updated Test-MsiCache.ps1 on the broken server so that column is present;
    otherwise the script falls back to DisplayVersion matching.

    SAFE BY DEFAULT: reports the match table and does nothing. Add -Execute to
    copy. Supports -WhatIf. After each copy it re-opens the file and confirms the
    MSI's ProductCode matches what was expected (unless -SkipVerify).

.PARAMETER MissingCsv
    CSV of missing products from the broken server (Test-MsiCache.ps1 -MissingOnly).

.PARAMETER SourcePath
    Folder holding the donor's cache files: a local copy of its C:\Windows\Installer
    (scan mode) or a UNC path to it (e.g. \\DONOR\C$\Windows\Installer).

.PARAMETER LabCsv
    OPTIONAL. Donor's full audit CSV. Omit to scan -SourcePath.

.PARAMETER DestinationPath
    Staging folder where matched files are written, each RENAMED to the broken
    server's expected filename. Never the live cache. Created if missing.

.PARAMETER IgnoreVersion
    Disable the build guard and match on ProductCode only. NOT recommended - can
    stage a same-ProductCode file of a different build.

.PARAMETER SkipVerify
    Skip the post-copy ProductCode verification (faster, less safe).

.PARAMETER Execute
    Actually perform the copies. Without it the script is a dry run / report only.

.EXAMPLE
    # SCAN MODE - donor's installer dir copied locally; dry run
    .\Restore-MsiCache.ps1 -MissingCsv .\Missing-Only.csv `
        -SourcePath C:\_PKG\InstallerCache -DestinationPath C:\_PKG\Staging

.EXAMPLE
    # SCAN MODE - stage the matched, renamed, verified files
    .\Restore-MsiCache.ps1 -MissingCsv .\Missing-Only.csv `
        -SourcePath C:\_PKG\InstallerCache -DestinationPath C:\_PKG\Staging -Execute
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory)][string]$MissingCsv,
    [Parameter(Mandatory)][string]$SourcePath,
    [string]$LabCsv,
    [Parameter(Mandatory)][string]$DestinationPath,
    [switch]$IgnoreVersion,
    [switch]$SkipVerify,
    [switch]$Execute
)

# --- Read ProductCode, ProductVersion and PackageCode from an .msi via the
# Windows Installer COM API. PackageCode is the per-build Revision GUID from the
# summary information stream (PID_REVNUMBER = 9) and uniquely identifies a build.
function Get-MsiProperties {
    param([string]$Path)
    $installer = $null; $db = $null; $sum = $null
    $out = [ordered]@{ ProductCode = $null; ProductVersion = $null; PackageCode = $null }
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.GetType().InvokeMember('OpenDatabase','InvokeMethod',$null,$installer,@($Path,0))
        foreach ($name in 'ProductCode','ProductVersion') {
            $view = $null; $rec = $null
            try {
                $view = $db.GetType().InvokeMember('OpenView','InvokeMethod',$null,$db,
                            @("SELECT Value FROM Property WHERE Property='$name'"))
                $view.GetType().InvokeMember('Execute','InvokeMethod',$null,$view,$null) | Out-Null
                $rec = $view.GetType().InvokeMember('Fetch','InvokeMethod',$null,$view,$null)
                if ($rec) { $out[$name] = $rec.GetType().InvokeMember('StringData','GetProperty',$null,$rec,@(1)) }
            } finally {
                foreach ($o in @($rec,$view)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
            }
        }
        try {
            $sum = $installer.GetType().InvokeMember('SummaryInformation','GetProperty',$null,$installer,@($Path,0))
            $out['PackageCode'] = $sum.GetType().InvokeMember('Property','GetProperty',$null,$sum,@(9))
        } catch { }
    } catch {
    } finally {
        foreach ($o in @($sum,$db,$installer)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    }
    [pscustomobject]$out
}

# Safe property read (handles rows that may not have the column).
function Get-Prop {
    param($Object, [string]$Name)
    if ($Object -and ($Object.PSObject.Properties.Name -contains $Name)) { $Object.$Name } else { $null }
}

if (-not (Test-Path -LiteralPath $MissingCsv)) { throw "CSV not found: $MissingCsv" }
if (-not (Test-Path -LiteralPath $SourcePath)) { throw "SourcePath not found: $SourcePath" }

$missing = Import-Csv -LiteralPath $MissingCsv | Where-Object { $_.ProductCode }

# --- Build donor lookups, by ProductCode AND by PackageCode -----------------
# In SCAN mode the ProductCode read from an .msi is the file's *internal* code,
# which for SQL Server differs from the registered ProductCode the missing CSV
# holds. The PackageCode is identical on both sides, so it's the reliable key for
# those - we fall back to it when ProductCode doesn't line up.
$donorMap       = @{}   # keyed by ProductCode
$donorByPackage = @{}   # keyed by PackageCode
if ($LabCsv) {
    if (-not (Test-Path -LiteralPath $LabCsv)) { throw "CSV not found: $LabCsv" }
    foreach ($row in (Import-Csv -LiteralPath $LabCsv | Where-Object { $_.ProductCode })) {
        $donorMap[$row.ProductCode.ToUpper()] = $row
        $rp = Get-Prop $row 'PackageCode'
        if ($rp) { $donorByPackage[$rp.ToUpper()] = $row }
    }
    $sourceDesc = "donor CSV ($LabCsv)"
}
else {
    # SCAN MODE - read identity from every .msi in SourcePath.
    $files = @(Get-ChildItem -LiteralPath $SourcePath -Filter *.msi -File -ErrorAction SilentlyContinue)
    Write-Host "Scanning $($files.Count) MSIs in $SourcePath ..." -ForegroundColor Cyan
    $i = 0
    foreach ($f in $files) {
        $i++
        Write-Progress -Activity 'Scanning source MSIs' -Status "$i / $($files.Count): $($f.Name)" `
            -PercentComplete (($i / [math]::Max($files.Count,1)) * 100)
        $props = Get-MsiProperties -Path $f.FullName
        if ($props.ProductCode) {
            $donorRowObj = [pscustomobject]@{
                CacheFileName  = $f.Name
                DisplayVersion = $props.ProductVersion
                PackageCode    = $props.PackageCode
            }
            $donorMap[$props.ProductCode.ToUpper()] = $donorRowObj
            if ($props.PackageCode) { $donorByPackage[$props.PackageCode.ToUpper()] = $donorRowObj }
        }
    }
    Write-Progress -Activity 'Scanning source MSIs' -Completed
    $sourceDesc = "scanned $($files.Count) MSIs in $SourcePath"
}

if (-not (Test-Path -LiteralPath $DestinationPath)) {
    if ($Execute -and $PSCmdlet.ShouldProcess($DestinationPath, 'Create staging folder')) {
        New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
    }
}

$plan = foreach ($m in $missing) {
    $code     = $m.ProductCode.ToUpper()
    $missPkg  = Get-Prop $m 'PackageCode'   # the build the broken server requires
    $donorRow = $donorMap[$code]

    $status = $null; $source = $null
    # Target = the broken server's expected cache filename, staged into YOUR folder.
    $targetName = if (-not [string]::IsNullOrWhiteSpace($m.LocalPackage)) {
        Split-Path $m.LocalPackage -Leaf
    } else {
        $m.CacheFileName
    }
    $target = if ($targetName) { Join-Path $DestinationPath $targetName } else { $null }

    # Decide whether the donor candidate is the SAME BUILD, not just same product.
    $buildOk = $true; $buildNote = ''; $matchedOn = ''; $verifyBy = 'ProductCode'
    if ($donorRow) {
        $donorPkg = Get-Prop $donorRow 'PackageCode'
        $donorVer = Get-Prop $donorRow 'DisplayVersion'
        if ($IgnoreVersion) {
            $matchedOn = 'ProductCode (version ignored)'
        }
        elseif ($missPkg -and $donorPkg) {
            $matchedOn = 'ProductCode+PackageCode'
            if ($missPkg.ToUpper() -ne $donorPkg.ToUpper()) {
                $buildOk = $false; $buildNote = 'PackageMismatch'
            }
        }
        else {
            $matchedOn = 'ProductCode+Version'
            if ($m.DisplayVersion -ne $donorVer) {
                $buildOk = $false; $buildNote = "VersionMismatch ($($m.DisplayVersion) vs $donorVer)"
            }
        }
    }
    elseif ($missPkg -and $donorByPackage.ContainsKey($missPkg.ToUpper())) {
        # ProductCode differs (SQL registers a per-instance code while the cached
        # package keeps a base code); the PackageCode IS the package identity.
        # Match on it and verify on it after copy.
        $donorRow  = $donorByPackage[$missPkg.ToUpper()]
        $matchedOn = 'PackageCode'
        $verifyBy  = 'PackageCode'
    }

    if (-not $donorRow) {
        $status = 'NoDonorMatch'
    }
    else {
        $source = Join-Path $SourcePath $donorRow.CacheFileName
        if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
            $status = 'SourceFileMissing'
        }
        elseif (-not $buildOk) {
            $status = $buildNote
        }
        elseif ([string]::IsNullOrWhiteSpace($target)) {
            $status = 'NoTargetPath'
        }
        elseif (Test-Path -LiteralPath $target -PathType Leaf) {
            $status = 'TargetAlreadyRestored'
        }
        else {
            $status = 'Ready'
        }
    }

    [pscustomobject]@{
        DisplayName    = $m.DisplayName
        DisplayVersion = $m.DisplayVersion
        ProductCode    = $m.ProductCode
        PackageCode    = $missPkg
        MatchedOn      = $matchedOn
        VerifyBy       = $verifyBy
        SourceFile     = $source
        TargetFile     = $target
        Status         = $status
        Result         = ''
    }
}

# --- Report the plan ---------------------------------------------------------
$plan | Format-Table DisplayName, DisplayVersion, ProductCode, PackageCode, MatchedOn, Status -AutoSize

$ready   = @($plan | Where-Object Status -eq 'Ready')
$blocked = @($plan | Where-Object Status -ne 'Ready')

Write-Host ""
Write-Host "Source           : $sourceDesc"
Write-Host "Build guard      : $(if ($IgnoreVersion) {'OFF (ProductCode only)'} else {'ON (PackageCode/Version must match)'})"
Write-Host "Missing products : $($missing.Count)"
Write-Host "Ready to restore : $($ready.Count)"   -ForegroundColor Green
if ($blocked.Count) {
    # Break the blocked total down by status (strip any parenthetical detail so
    # "VersionMismatch (a vs b)" groups under "VersionMismatch").
    $blocked |
        Group-Object { ($_.Status -split ' \(')[0] } |
        Sort-Object Count -Descending |
        ForEach-Object {
            $color = if ($_.Name -eq 'TargetAlreadyRestored') { 'Green' } else { 'Yellow' }
            Write-Host ("{0,-16} : {1}" -f $_.Name, $_.Count) -ForegroundColor $color
        }
}

# --- Execute -----------------------------------------------------------------
if (-not $Execute) {
    Write-Host "`nDRY RUN - no files copied. Re-run with -Execute to perform the restore." -ForegroundColor Cyan
}
else {
    foreach ($item in $ready) {
        if ($PSCmdlet.ShouldProcess($item.TargetFile, "Copy from $($item.SourceFile)")) {
            try {
                Copy-Item -LiteralPath $item.SourceFile -Destination $item.TargetFile -Force -ErrorAction Stop

                if ($SkipVerify) {
                    $item.Result = 'Copied'
                }
                else {
                    # Verify on whatever we matched on: PackageCode for SQL-style
                    # per-instance products, otherwise ProductCode.
                    $props = Get-MsiProperties -Path $item.TargetFile
                    if ($item.VerifyBy -eq 'PackageCode') {
                        $actual = $props.PackageCode; $expected = $item.PackageCode
                    } else {
                        $actual = $props.ProductCode; $expected = $item.ProductCode
                    }
                    if ($actual -and $expected -and $actual.ToUpper() -eq $expected.ToUpper()) {
                        $item.Result = "Copied+Verified ($($item.VerifyBy))"
                    }
                    else {
                        $item.Result = "VERIFY-FAILED ($($item.VerifyBy): file=$actual) - removed"
                        Remove-Item -LiteralPath $item.TargetFile -Force -ErrorAction SilentlyContinue
                    }
                }
            }
            catch {
                $item.Result = "ERROR: $($_.Exception.Message)"
            }
        }
    }

    Write-Host "`n--- Restore results ---`n"
    $ready | Format-Table DisplayName, ProductCode, Result -AutoSize

    $ok   = @($ready | Where-Object Result -like 'Copied*').Count
    $fail = $ready.Count - $ok
    Write-Host ""
    Write-Host "Restored OK : $ok"   -ForegroundColor Green
    Write-Host "Failed      : $fail" -ForegroundColor ($(if ($fail) {'Red'} else {'Green'}))
}

# --- Log the plan/results for the record ------------------------------------
$logDir  = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$logPath = Join-Path $logDir "MsiCacheRestore_${env:COMPUTERNAME}_$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
$plan | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding UTF8
Write-Host "`nLog written : $logPath" -ForegroundColor Cyan
