<#
.SYNOPSIS
    Lists all applications available to the current user in Software Center,
    including install state and whether it is user-targeted or machine-targeted.

.PARAMETER ComputerName
    Remote computer to query. Defaults to localhost.

.PARAMETER ShowAll
    Include already-installed applications in the output (default: available only).
#>

param(
    [string]$ComputerName = $env:COMPUTERNAME,
    [switch]$ShowAll
)

$cimParams = @{ Namespace = 'root\ccm\clientSDK'; ClassName = 'CCM_Application' }
if ($ComputerName -ne $env:COMPUTERNAME) {
    $cimParams['ComputerName'] = $ComputerName
}

try {
    $apps = Get-CimInstance @cimParams -ErrorAction Stop
} catch {
    Write-Warning "Could not query Software Center applications on $ComputerName`: $_"
    exit 1
}

# InstallState values: NotInstalled=0, Unknown=1, Error=2, Installed=3, NotEvaluated=4, NotUpdated=5, Obsoleted=6
$stateMap = @{ 0='Not Installed'; 1='Unknown'; 2='Error'; 3='Installed'; 4='Not Evaluated'; 5='Not Updated'; 6='Obsolete' }

$results = $apps | Where-Object {
    $ShowAll -or $_.InstallState -ne 3
} | Select-Object `
    Name,
    @{ N='Version';       E={ $_.SoftwareVersion } },
    @{ N='Publisher';     E={ $_.Publisher } },
    @{ N='InstallState';  E={ $stateMap[[int]$_.InstallState] } },
    @{ N='Deadline';      E={ if ($_.Deadline) { [datetime]$_.Deadline } else { 'None' } } },
    @{ N='UserTargeted';  E={ $_.UserUI } },
    @{ N='HighImpact';    E={ $_.HighImpact } } |
    Sort-Object InstallState, Name

if (-not $results) {
    Write-Host "No available applications found for this user/machine in Software Center." -ForegroundColor Yellow
} else {
    $results | Format-Table -AutoSize
    Write-Host "Total: $($results.Count) application(s)" -ForegroundColor Cyan
}

# Also show available legacy packages/programs
try {
    $pkgParams = @{ Namespace = 'root\ccm\clientSDK'; ClassName = 'CCM_Program' }
    if ($ComputerName -ne $env:COMPUTERNAME) { $pkgParams['ComputerName'] = $ComputerName }

    $programs = Get-CimInstance @pkgParams -ErrorAction Stop |
        Where-Object { $ShowAll -or $_.CurrentState -ne 'Succeeded' } |
        Select-Object PackageName, ProgramName,
            @{ N='State'; E={ $_.CurrentState } },
            @{ N='LastRunStatus'; E={ $_.LastRunStatus } } |
        Sort-Object PackageName

    if ($programs) {
        Write-Host "`nLegacy Packages / Programs:" -ForegroundColor Cyan
        $programs | Format-Table -AutoSize
    }
} catch {
    # CCM_Program may not exist on all client versions; silently skip
}
