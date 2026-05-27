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

# ── Input Validation ──────────────────────────────────────────────────────────
# Reject ComputerName values that contain characters outside the valid hostname /
# FQDN / IPv4 character set.  This prevents CIM-parameter injection.
if ($ComputerName -notmatch '^[\w][\w\-\.]{0,253}[\w]$|^\w$') {
    Write-Error "Invalid ComputerName '$ComputerName'. Only letters, digits, hyphens, underscores and dots are allowed."
    return
}

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
    if ($raw -eq '0' -or [string]::IsNullOrWhiteSpace($raw)) { exit 0 }

    # Accept only digits and commas; reject everything else to prevent injection
    if ($raw -notmatch '^[\d,\s]+$') {
        Write-Warning "Invalid input. Only numbers and commas are accepted."
        return
    }

    $chosen = $raw -split ',' |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -match '^\d+$' } |
        ForEach-Object { [int]$_ } |
        Where-Object { $_ -ge 1 -and $_ -le $indexed.Count }   # bounds-check

    $selected = $indexed | Where-Object { $_.'#' -in $chosen } | ForEach-Object { $_.App }
}

if (-not $selected) { Write-Host "No selection made. Exiting." -ForegroundColor Yellow; return }

# ── Trigger Installation via CCM_Application.Install() ───────────────────────
foreach ($pick in $selected) {
    # Match back to the raw CCM_Application instance to get Id/Revision/IsMachineTarget
    $rawApp = $apps | Where-Object { $_.Name -eq $pick.Name } | Select-Object -First 1

    if (-not $rawApp) {
        Write-Warning "Could not locate source CIM instance for '$($pick.Name)'. Skipping."
        continue
    }

    Write-Host "Queuing install: $($pick.Name) $($pick.Version) …" -ForegroundColor Cyan

    try {
        $methodArgs = @{
            Id               = $rawApp.Id
            Revision         = $rawApp.Revision
            IsMachineTarget  = [bool]$rawApp.IsMachineTarget
            IsEnforced       = $true
            Priority         = 'High'
            IsRebootIfNeeded = $false
        }

        $invokeParams = @{
            Namespace  = 'root\ccm\clientSDK'
            ClassName  = 'CCM_Application'
            MethodName = 'Install'
            Arguments  = $methodArgs
        }
        if ($isRemote) { $invokeParams['ComputerName'] = $ComputerName }

        $result = Invoke-CimMethod @invokeParams -ErrorAction Stop

        switch ($result.ReturnValue) {
            0       { Write-Host "  ✓ Install queued successfully." -ForegroundColor Green }
            default { Write-Warning "  Install returned code $($result.ReturnValue) for '$($pick.Name)'." }
        }
    }
    catch {
        Write-Warning "  Failed to queue install for '$($pick.Name)': $_"
    }
}
