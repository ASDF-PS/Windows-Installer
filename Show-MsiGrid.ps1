<#
.SYNOPSIS
    Recursively scan a folder for .msi files and show their details
    (ProductName, version, manufacturer, ProductCode, PackageCode, size, full
    path) in an interactive Out-GridView you can sort and filter.

.PARAMETER Path
    Root folder to search (recurses every subfolder). Defaults to the current
    directory.

.EXAMPLE
    .\Show-MsiGrid.ps1
    .\Show-MsiGrid.ps1 -Path 'D:\SomeOtherCache'
    .\Show-MsiGrid.ps1 -Path 'C:\Windows\Installer'
#>
param(
    [string]$Path = (Get-Location).Path
)

# Read identity fields from an .msi via the Windows Installer COM API.
function Get-MsiId {
    param([string]$File)
    $i = $null; $db = $null; $sum = $null
    $o = [ordered]@{ ProductName = $null; ProductVersion = $null; Manufacturer = $null; ProductCode = $null; PackageCode = $null }
    try {
        $i  = New-Object -ComObject WindowsInstaller.Installer
        $db = $i.GetType().InvokeMember('OpenDatabase','InvokeMethod',$null,$i,@($File,0))
        foreach ($n in 'ProductName','ProductVersion','Manufacturer','ProductCode') {
            $v = $null; $r = $null
            try {
                $v = $db.GetType().InvokeMember('OpenView','InvokeMethod',$null,$db,
                        @("SELECT Value FROM Property WHERE Property='$n'"))
                $v.GetType().InvokeMember('Execute','InvokeMethod',$null,$v,$null) | Out-Null
                $r = $v.GetType().InvokeMember('Fetch','InvokeMethod',$null,$v,$null)
                if ($r) { $o[$n] = $r.GetType().InvokeMember('StringData','GetProperty',$null,$r,@(1)) }
            } finally {
                foreach ($x in @($r,$v)) { if ($x) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) } }
            }
        }
        try {
            $sum = $i.GetType().InvokeMember('SummaryInformation','GetProperty',$null,$i,@($File,0))
            $o['PackageCode'] = $sum.GetType().InvokeMember('Property','GetProperty',$null,$sum,@(9))
        } catch { }
    } catch {
    } finally {
        foreach ($x in @($sum,$db,$i)) { if ($x) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) } }
    }
    [pscustomobject]$o
}

if (-not (Test-Path -LiteralPath $Path)) { throw "Path not found: $Path" }

$files = @(Get-ChildItem -LiteralPath $Path -Recurse -Filter *.msi -File -ErrorAction SilentlyContinue)
Write-Host "Scanning $($files.Count) MSIs under $Path ..." -ForegroundColor Cyan

$results = for ($n = 0; $n -lt $files.Count; $n++) {
    $f = $files[$n]
    Write-Progress -Activity 'Reading MSIs' -Status "$($n+1)/$($files.Count): $($f.Name)" `
        -PercentComplete (($n + 1) / [math]::Max($files.Count,1) * 100)
    $id = Get-MsiId $f.FullName
    [pscustomobject]@{
        ProductName    = $id.ProductName
        ProductVersion = $id.ProductVersion
        Manufacturer   = $id.Manufacturer
        ProductCode    = $id.ProductCode
        PackageCode    = $id.PackageCode
        SizeMB         = [math]::Round($f.Length / 1MB, 2)
        FileName       = $f.Name
        Path           = $f.FullName
    }
}
Write-Progress -Activity 'Reading MSIs' -Completed

Write-Host "Found $($results.Count) MSIs - opening grid..." -ForegroundColor Cyan
$results | Sort-Object ProductName | Out-GridView -Title "MSIs under $Path  ($($results.Count) found)"
