<#
.SYNOPSIS
    Finds the top 10 folders consuming the most disk space across an entire drive.
.EXAMPLE
    .\Get-TopDiskByDrive.ps1
    .\Get-TopDiskByDrive.ps1 -Drive D
    .\Get-TopDiskByDrive.ps1 -Drive C -TopN 20
#>
[CmdletBinding()]
param(
    [string]$Drive,
    [int]$TopN = 10
)

# If no drive specified, list available drives and prompt
if (-not $Drive) {
    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -match '^[A-Z]:\\$' }

    if ($drives.Count -eq 0) {
        Write-Error "No drives found."
        exit 1
    }

    Write-Host "`nAvailable drives:`n" -ForegroundColor Cyan
    for ($i = 0; $i -lt $drives.Count; $i++) {
        $d = $drives[$i]
        $usedGB  = [math]::Round(($d.Used  / 1GB), 1)
        $freeGB  = [math]::Round(($d.Free  / 1GB), 1)
        Write-Host ("  [{0}] {1}  Used: {2} GB   Free: {3} GB" -f ($i + 1), $d.Root, $usedGB, $freeGB)
    }

    Write-Host ""
    do {
        $selection = Read-Host "Select a drive number (1-$($drives.Count))"
        $index = $selection -as [int]
    } while (-not $index -or $index -lt 1 -or $index -gt $drives.Count)

    $scanRoot = $drives[$index - 1].Root
} else {
    $scanRoot = $Drive.TrimEnd(':\') + ':\'
    if (-not (Test-Path $scanRoot)) {
        Write-Error "Drive not found: $scanRoot"
        exit 1
    }
}

Write-Host "`nScanning: $scanRoot  (top $TopN folders by size)`n" -ForegroundColor Cyan
Write-Host "This may take several minutes on large drives..." -ForegroundColor DarkGray

$folders = Get-ChildItem -Path $scanRoot -Directory -ErrorAction SilentlyContinue

$results = foreach ($folder in $folders) {
    Write-Host "  Calculating: $($folder.FullName)" -ForegroundColor DarkGray
    $size = (Get-ChildItem -Path $folder.FullName -Recurse -File -ErrorAction SilentlyContinue |
        Measure-Object -Property Length -Sum).Sum

    [PSCustomObject]@{
        Path      = $folder.FullName
        SizeBytes = if ($size) { $size } else { 0 }
        SizeGB    = if ($size) { [math]::Round($size / 1GB, 3) } else { 0 }
        SizeMB    = if ($size) { [math]::Round($size / 1MB, 1) } else { 0 }
    }
}

# Also count loose files directly in the root (not in any subfolder)
$rootFiles = (Get-ChildItem -Path $scanRoot -File -ErrorAction SilentlyContinue |
    Measure-Object -Property Length -Sum).Sum
if ($rootFiles -gt 0) {
    $results += [PSCustomObject]@{
        Path      = "$scanRoot [root files]"
        SizeBytes = $rootFiles
        SizeGB    = [math]::Round($rootFiles / 1GB, 3)
        SizeMB    = [math]::Round($rootFiles / 1MB, 1)
    }
}

$topN = $results | Sort-Object SizeBytes -Descending | Select-Object -First $TopN

if ($topN.Count -eq 0) {
    Write-Warning "No folders found or no access to drive contents."
    exit 0
}

$totalBytes = ($results | Measure-Object -Property SizeBytes -Sum).Sum
$topBytes   = ($topN    | Measure-Object -Property SizeBytes -Sum).Sum

Write-Host "`nTop $TopN largest folders on ${scanRoot}:`n" -ForegroundColor Green
Write-Host ("{0,-6} {1,-12} {2}" -f "Rank", "Size", "Path")
Write-Host ("-" * 75)

$rank = 1
foreach ($item in $topN) {
    $sizeDisplay = if ($item.SizeGB -ge 1) {
        "$($item.SizeGB) GB"
    } elseif ($item.SizeMB -ge 1) {
        "$($item.SizeMB) MB"
    } else {
        "$([math]::Round($item.SizeBytes / 1KB, 1)) KB"
    }
    Write-Host ("{0,-6} {1,-12} {2}" -f "#$rank", $sizeDisplay, $item.Path)
    $rank++
}

Write-Host ("-" * 75)

$totalDisplay = if ($totalBytes / 1GB -ge 1) { "$([math]::Round($totalBytes / 1GB, 3)) GB" } else { "$([math]::Round($totalBytes / 1MB, 1)) MB" }
$topDisplay   = if ($topBytes   / 1GB -ge 1) { "$([math]::Round($topBytes   / 1GB, 3)) GB" } else { "$([math]::Round($topBytes   / 1MB, 1)) MB" }

Write-Host ("       {0,-12} Total of top $TopN" -f $topDisplay) -ForegroundColor Yellow
Write-Host ("       {0,-12} Total scanned (all folders)" -f $totalDisplay) -ForegroundColor DarkGray
Write-Host ""
