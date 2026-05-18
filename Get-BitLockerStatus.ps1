# Check BitLocker encryption status for all drives on a local or remote device
# Requires elevation (Run as Administrator) for full details
# Usage: .\Get-BitLockerStatus.ps1
#        .\Get-BitLockerStatus.ps1 -ComputerName PC01
#        .\Get-BitLockerStatus.ps1 -ComputerName PC01,PC02,PC03

param(
    [string[]]$ComputerName = $env:COMPUTERNAME
)

$queryBlock = {
    try {
        $volumes = Get-BitLockerVolume -ErrorAction Stop
        foreach ($vol in $volumes) {
            [PSCustomObject]@{
                MountPoint          = $vol.MountPoint
                VolumeStatus        = $vol.VolumeStatus
                ProtectionStatus    = $vol.ProtectionStatus
                EncryptionPct       = $vol.EncryptionPercentage
                KeyProtectors       = ($vol.KeyProtector.KeyProtectorType -join ', ')
            }
        }
    }
    catch {
        # manage-bde fallback for systems where Get-BitLockerVolume is unavailable
        manage-bde -status
    }
}

foreach ($Computer in $ComputerName) {
    Write-Host "`n=== $Computer ===" -ForegroundColor Cyan

    try {
        $isLocal = ($Computer -eq $env:COMPUTERNAME) -or ($Computer -eq 'localhost') -or ($Computer -eq '.')
        $results = if ($isLocal) {
            & $queryBlock
        } else {
            Invoke-Command -ComputerName $Computer -ScriptBlock $queryBlock -ErrorAction Stop
        }

        foreach ($item in $results) {
            if ($item -is [string]) {
                # manage-bde raw text output
                Write-Host $item
            } else {
                $protected = $item.ProtectionStatus -eq 'On'
                $color = if ($protected) { 'Green' } else { 'Red' }

                Write-Host "  Drive $($item.MountPoint)" -ForegroundColor White
                Write-Host "    Encryption Status : $($item.VolumeStatus)" -ForegroundColor $color
                Write-Host "    Protection Status : $($item.ProtectionStatus)" -ForegroundColor $color
                Write-Host "    Encryption %      : $($item.EncryptionPct)%"
                Write-Host "    Key Protectors    : $($item.KeyProtectors)"
            }
        }
    }
    catch {
        Write-Warning "  Could not query $Computer`: $_"
    }
}
