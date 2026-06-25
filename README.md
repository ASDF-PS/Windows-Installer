# MSI Cache Tools

Three PowerShell utilities for auditing and repairing the **Windows Installer
cache** (`C:\Windows\Installer`) and inspecting MSI files in bulk.

Windows Installer keeps a private copy of every installed `.msi` in
`C:\Windows\Installer`, renamed to a random name (e.g. `119918f.msi`). If those
cached files go missing, repairs, patches, and uninstalls break. These scripts
let you inspect product/cache state and restore missing cache files from a known
good "donor" machine — matched on **build identity**, not just ProductCode.

All three are read-only by default; the only script that writes files
(`Restore-MsiCache.ps1`) is a dry run unless you pass `-Execute`.

---

## Scripts

| Script | Purpose | Writes? |
|---|---|---|
| `Get-MsiCacheInfo.ps1` | Look up one product by ProductCode and report its cache location + whether the cached file still exists. Built for the ConfigMgr **Run Scripts** feature (emits JSON). | No |
| `Show-MsiGrid.ps1` | Recursively scan a folder for `.msi` files and show their identity (name, version, manufacturer, ProductCode, PackageCode, size, path) in a sortable/filterable grid. | No |
| `Restore-MsiCache.ps1` | Restore missing cache files for a broken server by matching them against a donor's cache and staging the matches renamed to the broken server's expected filenames. | Only with `-Execute` |

---

## Why build identity matters

Windows Installer keeps the **same ProductCode** across minor updates and
patches, so a matching ProductCode does **not** guarantee an identical package.
`Restore-MsiCache.ps1` therefore requires the **build** to agree as well:

* If both sides expose a **PackageCode** (the MSI's unique per-build revision
  GUID), it must match exactly — this guarantees an identical package.
* Otherwise it falls back to requiring **DisplayVersion** to match.
* `-IgnoreVersion` disables that guard (ProductCode only) — **not recommended**.

After each copy the script re-opens the file and verifies the identity it
matched on before keeping it.

---

## `Get-MsiCacheInfo.ps1`

Given a ProductCode, returns a single compact JSON object with the ProductName,
version, PackageCode, cache path (`LocalPackage`), and whether that cached file
exists. Runs as SYSTEM on the target device when invoked via ConfigMgr Run
Scripts.

```powershell
.\Get-MsiCacheInfo.ps1 -ProductCode '{1088A5C5-4543-4706-B1D7-111E0FBDC0BC}'
# braces optional, case-insensitive:
.\Get-MsiCacheInfo.ps1 -ProductCode 1088A5C5-4543-4706-B1D7-111E0FBDC0BC
```

---

## `Show-MsiGrid.ps1` — scan a specific folder

Recursively reads every `.msi` under a folder and opens an interactive
`Out-GridView`. Defaults to the **current directory**; pass `-Path` to scan a
specific folder.

```powershell
# Scan the current directory
.\Show-MsiGrid.ps1

# Scan a specific folder (recurses every subfolder)
.\Show-MsiGrid.ps1 -Path 'D:\SomeCache'

# Inspect the live Windows Installer cache
.\Show-MsiGrid.ps1 -Path 'C:\Windows\Installer'
```

---

## `Restore-MsiCache.ps1` — scan a source, stage to a destination

Two modes for pointing at the donor's files:

* **Scan mode** (default): you copied the donor's `C:\Windows\Installer` to a
  local folder. The script reads identity from every `.msi` in `-SourcePath`.
* **CSV mode** (`-LabCsv`): you have the donor's full audit CSV and reach its
  files over a share. (PackageCode matching needs a `PackageCode` column in the
  CSV.)

Matched donor files are copied to a **staging folder** you choose
(`-DestinationPath`) — never the live cache — each renamed to the broken
server's expected filename.

### Typical scan-mode round trip

```powershell
# 1. On the BROKEN server, produce the list of missing products
#    (use your audit script, e.g. Test-MsiCache.ps1 -MissingOnly -> Missing-Only.csv)

# 2. Copy a DONOR's installer cache to a local folder
robocopy '\\DONOR\C$\Windows\Installer' 'C:\_PKG\MSICache' /E

# 3. Scan that source folder, stage matches to a destination — DRY RUN first
.\Restore-MsiCache.ps1 -MissingCsv .\Missing-Only.csv `
    -SourcePath 'C:\_PKG\MSICache' `
    -DestinationPath 'C:\_PKG\Staging'

# 4. Same command with -Execute to copy + verify the matched, renamed files
.\Restore-MsiCache.ps1 -MissingCsv .\Missing-Only.csv `
    -SourcePath 'C:\_PKG\MSICache' `
    -DestinationPath 'C:\_PKG\Staging' -Execute

# 5. Copy C:\_PKG\Staging\*.msi back to the broken server's C:\Windows\Installer
#    and re-run your audit to confirm.
```

### Key parameters

| Parameter | Meaning |
|---|---|
| `-MissingCsv` | CSV of missing products from the broken server (must have `ProductCode`; `PackageCode`/`LocalPackage`/`DisplayVersion` improve matching). |
| `-SourcePath` | Folder holding the donor's cache files (local copy, or a UNC path to its `C:\Windows\Installer`). |
| `-DestinationPath` | Staging folder for the matched, renamed files. Created if missing. Never the live cache. |
| `-LabCsv` | Optional. Donor's full audit CSV — omit to scan `-SourcePath`. |
| `-IgnoreVersion` | Match on ProductCode only (disables the build guard). Not recommended. |
| `-SkipVerify` | Skip the post-copy identity verification (faster, less safe). |
| `-Execute` | Actually perform the copies. Without it the script is a dry run / report only. Supports `-WhatIf`. |

Every run also writes a CSV log of the plan/results next to the script:
`MsiCacheRestore_<COMPUTERNAME>_<timestamp>.csv`.

### Example output

A dry run reports the match plan; `-Execute` adds the per-file restore results:

.\Restore-MsiCache.ps1 -MissingCsv "C:\_PKG\MSICache\MsiCacheAudit_LSSCCM_20260622-111402.csv" -SourcePath C:\_PKG\MSICache -DestinationPath C:\_PKG\MSIStaging -Execute

<img width="1749" height="701" alt="restore-msi-cache" src="https://github.com/user-attachments/assets/032e6b4c-6240-4e3a-935a-3bbfa4801ad7" />

In the run above, 116 products were missing on the broken server. Most
(`TargetAlreadyRestored`) were already staged from a prior pass; 9 had no donor
match and 1 was a `PackageMismatch` (same ProductCode, different build — correctly
refused). The 2 that were `Ready` (Visual C++ 2013 runtimes) were copied and
`Copied+Verified (ProductCode)`.

---

## Requirements

* Windows PowerShell 5.1+ (uses the `WindowsInstaller.Installer` COM API).
* `Show-MsiGrid.ps1` uses `Out-GridView`, which needs the desktop PowerShell
  ISE/console (not available in PowerShell remoting sessions).
* `Get-MsiCacheInfo.ps1` reports machine-context products; run it as the same
  context (e.g. SYSTEM) that installed the product.
