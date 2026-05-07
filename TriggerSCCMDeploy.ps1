<#
.SYNOPSIS
    Forces the SCCM/ConfigMgr client to immediately download policy and evaluate deployments.
    Run this on the target machine after pushing a package from SCCM to kick off installation
    without waiting for the next scheduled client cycle (default is every 60 minutes).

.PARAMETER ComputerName
    Remote computer to trigger. Defaults to localhost.
#>

param(
    [string]$ComputerName = $env:COMPUTERNAME
)

$schedules = @(
    @{ ID = '{00000000-0000-0000-0000-000000000021}'; Name = 'Machine Policy Retrieval & Evaluation' },
    @{ ID = '{00000000-0000-0000-0000-000000000121}'; Name = 'Application Deployment Evaluation'     }
)

$cimParams = @{ Namespace = 'root\ccm'; ClassName = 'SMS_Client' }
if ($ComputerName -ne $env:COMPUTERNAME) {
    $cimParams['ComputerName'] = $ComputerName
}

# Verify the SCCM client service is running
$svc = Get-Service -Name CcmExec -ComputerName $ComputerName -ErrorAction SilentlyContinue
if (-not $svc -or $svc.Status -ne 'Running') {
    Write-Warning "CcmExec service is not running on $ComputerName. Is the SCCM client installed?"
    exit 1
}

foreach ($schedule in $schedules) {
    Write-Host "Triggering: $($schedule.Name)..." -ForegroundColor Cyan
    try {
        Invoke-CimMethod @cimParams -MethodName TriggerSchedule `
            -Arguments @{ sScheduleID = $schedule.ID } | Out-Null
        Write-Host "  OK" -ForegroundColor Green
    } catch {
        Write-Warning "  Failed: $_"
    }
    Start-Sleep -Seconds 5
}

Write-Host "`nDone. The client will now pull the latest policy and begin evaluating deployments." -ForegroundColor Yellow
Write-Host "Monitor progress in: C:\Windows\CCM\Logs\AppEnforce.log or use CMTrace."
