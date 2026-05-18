# Check BitLocker encryption status for all drives on a local or remote device
# Requires elevation (Run as Administrator) for full details
# Usage: .\Get-BitLockerStatus.ps1
#        .\Get-BitLockerStatus.ps1 -ComputerName PC01
#        .\Get-BitLockerStatus.ps1 -ComputerName PC01,PC02,PC03

param(
    [string[]]$ComputerName = $env:COMPUTERNAME
)

foreach ($Computer in $ComputerName) {
    Write-Host "`n=== $Computer ===" -ForegroundColor Cyan

    try {
        $volumes = Get-BitLockerVolume -MountPoint * -CimSession $Computer -ErrorAction Stop

        foreach ($vol in $volumes) {
            $protected = $vol.ProtectionStatus -eq 'On'
            $color = if ($protected) { 'Green' } else { 'Red' }

            Write-Host "  Drive $($vol.MountPoint)" -ForegroundColor White
            Write-Host "    Encryption Status : $($vol.VolumeStatus)" -ForegroundColor $color
            Write-Host "    Protection Status : $($vol.ProtectionStatus)" -ForegroundColor $color
            Write-Host "    Encryption %      : $($vol.EncryptionPercentage)%"
            Write-Host "    Key Protectors    : $($vol.KeyProtector.KeyProtectorType -join ', ')"
        }
    }
    catch [Microsoft.Management.Infrastructure.CimException] {
        # Fall back to manage-bde for older OS or no WMI access
        Write-Warning "WMI unavailable, falling back to manage-bde"
        try {
            $result = Invoke-Command -ComputerName $Computer -ScriptBlock {
                manage-bde -status
            } -ErrorAction Stop
            Write-Host $result
        }
        catch {
            Write-Warning "  Could not connect to $Computer`: $_"
        }
    }
    catch {
        Write-Warning "  Error on $Computer`: $_"
    }
}
