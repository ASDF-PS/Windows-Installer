<#
.SYNOPSIS
    SCCM Run Script: given an MSI ProductCode, return its PackageCode, the Windows
    Installer cache location (LocalPackage), and whether that cached file exists.

.DESCRIPTION
    Built for the Configuration Manager "Run Scripts" feature. Runs as SYSTEM on
    the target device, queries Windows Installer for the supplied product, and
    writes a single JSON object to stdout (captured as the script output in CM).

    Read-only. Makes no changes.

.PARAMETER ProductCode
    The product's ProductCode GUID. Braces optional, case-insensitive. Examples:
      {1088A5C5-4543-4706-B1D7-111E0FBDC0BC}
      1088A5C5-4543-4706-B1D7-111E0FBDC0BC
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$ProductCode
)

# Normalise to {UPPERCASE-GUID}
$pc = '{' + ($ProductCode.Trim().Trim('{', '}').ToUpper()) + '}'

$result = [ordered]@{
    ProductCode  = $pc
    Found        = $false
    ProductName  = $null
    Version      = $null
    PackageCode  = $null
    LocalPackage = $null
    CacheExists  = $false
    Error        = $null
}

try {
    $installer = New-Object -ComObject WindowsInstaller.Installer

    # ProductInfo throws if the product isn't installed in this context; swallow
    # per-property so a single missing value doesn't abort the whole lookup.
    $get = {
        param($prop)
        try { $installer.ProductInfo($pc, $prop) } catch { $null }
    }

    $name = & $get 'InstalledProductName'
    if (-not $name) { $name = & $get 'ProductName' }

    if ($name) {
        $result.Found        = $true
        $result.ProductName  = $name
        $result.Version      = & $get 'VersionString'
        $result.PackageCode  = & $get 'PackageCode'
        $result.LocalPackage = & $get 'LocalPackage'
        if ($result.LocalPackage) {
            $result.CacheExists = Test-Path -LiteralPath $result.LocalPackage -PathType Leaf
        }
    }
    else {
        $result.Error = 'Product not found / not installed (machine context).'
    }
}
catch {
    $result.Error = $_.Exception.Message
}
finally {
    if ($installer) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($installer) }
}

[pscustomobject]$result | ConvertTo-Json -Compress
