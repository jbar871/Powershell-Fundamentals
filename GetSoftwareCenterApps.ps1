<#
.SYNOPSIS
    Lists all applications available to the current user in Software Center.
    Optionally presents an interactive picker to install a selected application.

.PARAMETER ComputerName
    Remote computer to query/install on. Defaults to localhost.

.PARAMETER ShowAll
    Include already-installed applications in the output (default: available only).

.PARAMETER Install
    After listing, open a selection menu so you can pick one or more apps to install.
    Uses Out-GridView (GUI) when available; falls back to a numbered console menu.
#>

param(
    [string]$ComputerName = $env:COMPUTERNAME,
    [switch]$ShowAll,
    [switch]$Install
)

$isRemote = $ComputerName -ne $env:COMPUTERNAME

# ── Query Applications ────────────────────────────────────────────────────────
$cimParams = @{ Namespace = 'root\ccm\clientSDK'; ClassName = 'CCM_Application' }
if ($isRemote) { $cimParams['ComputerName'] = $ComputerName }

try {
    $apps = Get-CimInstance @cimParams -ErrorAction Stop
} catch {
    Write-Warning "Could not query Software Center applications on ${ComputerName}: $_"
    return
}

# InstallState: 0=NotInstalled 1=Unknown 2=Error 3=Installed 4=NotEvaluated 5=NotUpdated 6=Obsolete
$stateMap = @{ 0='Not Installed'; 1='Unknown'; 2='Error'; 3='Installed'; 4='Not Evaluated'; 5='Not Updated'; 6='Obsolete' }

$results = $apps | Where-Object {
    $ShowAll -or $_.InstallState -ne 3
} | Select-Object `
    Name,
    @{ N='Version';      E={ $_.SoftwareVersion } },
    @{ N='Publisher';    E={ $_.Publisher } },
    @{ N='InstallState'; E={ $stateMap[[int]$_.InstallState] } },
    @{ N='Deadline';     E={ if ($_.Deadline) { [datetime]$_.Deadline } else { 'None' } } },
    @{ N='UserTargeted'; E={ $_.UserUI } },
    @{ N='HighImpact';   E={ $_.HighImpact } } |
    Sort-Object InstallState, Name

if (-not $results) {
    Write-Host "No available applications found for this user/machine in Software Center." -ForegroundColor Yellow
    return
}

$results | Format-Table -AutoSize
Write-Host "Total: $($results.Count) application(s)" -ForegroundColor Cyan

# ── Legacy Packages / Programs ────────────────────────────────────────────────
try {
    $pkgParams = @{ Namespace = 'root\ccm\clientSDK'; ClassName = 'CCM_Program' }
    if ($isRemote) { $pkgParams['ComputerName'] = $ComputerName }

    $programs = Get-CimInstance @pkgParams -ErrorAction Stop |
        Where-Object { $ShowAll -or $_.CurrentState -ne 'Succeeded' } |
        Select-Object PackageName, ProgramName,
            @{ N='State';         E={ $_.CurrentState } },
            @{ N='LastRunStatus'; E={ $_.LastRunStatus } } |
        Sort-Object PackageName

    if ($programs) {
        Write-Host "`nLegacy Packages / Programs:" -ForegroundColor Cyan
        $programs | Format-Table -AutoSize
    }
} catch { <# CCM_Program absent on some client versions #> }

# ── Interactive Install ───────────────────────────────────────────────────────
if (-not $Install) { return }

# Pick selection method: Out-GridView only works in a GUI/interactive session
$useGridView = (-not $isRemote) -and (Get-Command Out-GridView -ErrorAction SilentlyContinue)

if ($useGridView) {
    $selected = $results | Out-GridView -Title "Select application(s) to install — hold Ctrl for multi-select" -PassThru
} else {
    # Numbered console menu (works over remote sessions / no-GUI hosts)
    Write-Host "`nEnter the number(s) to install (comma-separated), or 0 to cancel:" -ForegroundColor Cyan
    $i = 1
    $indexed = $results | ForEach-Object { [PSCustomObject]@{ '#' = $i++; App = $_ } }
    $indexed | ForEach-Object { Write-Host "  $($_.'#')  $($_.App.Name)  $($_.App.Version)" }
    $raw = Read-Host "Selection"
    if ($raw -eq '0' -or [string]::IsNullOrWhiteSpace($raw)) { return }
    $chosen = $raw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' }
    $selected = $indexed | Where-Object { $_.'#' -in $chosen } | ForEach-Object { $_.App }
}

if (-not $selected) { Write-Host "No selection made. Exiting." -ForegroundColor Yellow; return }

foreach ($pick in $selected) {
    # Match back to the raw CCM_Application instance to get Id/Revision/IsMachineTarget
    $appObj = $apps | Where-Object {
        $_.Name -eq $pick.Name -and
        ([string]$_.SoftwareVersion -eq [string]$pick.Version)
    } | Select-Object -First 1

    if (-not $appObj) {
        Write-Warning "Could not find matching CCM_Application object for '$($pick.Name)'. Skipping."
        continue
    }

    Write-Host "`nInstalling: $($appObj.Name) $($appObj.SoftwareVersion)..." -ForegroundColor Cyan

    try {
        $installParams = @{ Namespace = 'root\ccm\clientSDK'; ClassName = 'CCM_Application' }
        if ($isRemote) { $installParams['ComputerName'] = $ComputerName }

        Invoke-CimMethod @installParams -MethodName Install -ErrorAction Stop -Arguments @{
            Id                = $appObj.Id
            Revision          = $appObj.Revision
            IsMachineTarget   = [bool]$appObj.IsMachineTarget
            EnforcePreference = [uint32]0
            Priority          = 'High'
        } | Out-Null

        Write-Host "  Install triggered. Monitor: C:\Windows\CCM\Logs\AppEnforce.log" -ForegroundColor Green
    } catch {
        Write-Warning "  Install failed for '$($appObj.Name)': $_"
        if ($_ -match '0x80070005|Access is denied') {
            Write-Host "  Tip: For user-targeted apps run without 'Run as Administrator'." -ForegroundColor Yellow
            Write-Host "       For machine-targeted apps run as SYSTEM: psexec -s powershell.exe" -ForegroundColor Yellow
        }
    }
}
