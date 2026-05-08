<#
.SYNOPSIS
    Check if a remote server is listening on port 445 (SMB).
.PARAMETER ComputerName
    Hostname or IP address of the target server.
.PARAMETER TimeoutMs
    Connection timeout in milliseconds (default: 3000).
.EXAMPLE
    .\CheckPort445.ps1 -ComputerName 192.168.1.10
    .\CheckPort445.ps1 -ComputerName fileserver.domain.local -TimeoutMs 5000
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ComputerName,

    [int]$TimeoutMs = 3000
)

$port = 445

Write-Host "Checking $ComputerName port $port..." -ForegroundColor Cyan

$tcp = New-Object System.Net.Sockets.TcpClient
try {
    $connect = $tcp.BeginConnect($ComputerName, $port, $null, $null)
    $waited  = $connect.AsyncWaitHandle.WaitOne($TimeoutMs, $false)

    if ($waited -and $tcp.Connected) {
        $tcp.EndConnect($connect)
        Write-Host "OPEN   - $ComputerName`:$port is listening." -ForegroundColor Green
        exit 0
    } else {
        Write-Host "CLOSED - $ComputerName`:$port is not reachable (timeout or refused)." -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "ERROR  - $ComputerName`:$port : $_" -ForegroundColor Yellow
    exit 2
} finally {
    $tcp.Close()
}
