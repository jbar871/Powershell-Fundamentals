#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [string]$ComputerName,
    [string]$PrinterPath = '\\prdprnmgmt01\8802-parts'
)

function Add-PrinterToRemoteMachine {
    param(
        [string]$Computer,
        [string]$Printer
    )

    Write-Host "`nConnecting to $Computer..." -ForegroundColor Cyan

    if (-not (Test-Connection -ComputerName $Computer -Count 1 -Quiet)) {
        Write-Warning "Cannot reach $Computer. Check the name or network connectivity."
        return $false
    }

    try {
        $result = Invoke-Command -ComputerName $Computer -ScriptBlock {
            param($PrinterPath)

            $existing = Get-Printer -Name $PrinterPath -ErrorAction SilentlyContinue
            if ($existing) {
                return [PSCustomObject]@{ Status = 'AlreadyExists'; Message = "Printer '$PrinterPath' is already installed." }
            }

            try {
                Add-Printer -ConnectionName $PrinterPath -ErrorAction Stop
                return [PSCustomObject]@{ Status = 'Success'; Message = "Printer '$PrinterPath' added successfully." }
            }
            catch {
                return [PSCustomObject]@{ Status = 'Error'; Message = $_.Exception.Message }
            }
        } -ArgumentList $Printer -ErrorAction Stop

        switch ($result.Status) {
            'Success'       { Write-Host "[OK] $($result.Message)" -ForegroundColor Green }
            'AlreadyExists' { Write-Host "[INFO] $($result.Message)" -ForegroundColor Yellow }
            'Error'         { Write-Warning "[FAIL] $($result.Message)" }
        }

        return $result.Status -in 'Success', 'AlreadyExists'
    }
    catch {
        Write-Warning "Remote session failed on ${Computer}: $($_.Exception.Message)"
        return $false
    }
}

# --- Main ---

Write-Host "==========================================" -ForegroundColor DarkCyan
Write-Host "  Remote Printer Deployment" -ForegroundColor DarkCyan
Write-Host "  Printer: $PrinterPath" -ForegroundColor DarkCyan
Write-Host "==========================================" -ForegroundColor DarkCyan

if (-not $ComputerName) {
    $ComputerName = Read-Host "`nEnter the target computer name"
}

if ([string]::IsNullOrWhiteSpace($ComputerName)) {
    Write-Warning "No computer name provided. Exiting."
    exit 1
}

$success = Add-PrinterToRemoteMachine -Computer $ComputerName.Trim() -Printer $PrinterPath

if ($success) {
    Write-Host "`nDone." -ForegroundColor Green
} else {
    Write-Host "`nDeployment failed. Review warnings above." -ForegroundColor Red
    exit 1
}
