# Install-RSU.ps1
# Silent UI-automation installer for Right-Suite Universal (RSUSetupWSF26001.exe)

param(
    [string]$ExePath = "C:\temp\RSUSetupWSF26001.exe"
)

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$LogDir  = "C:\Logs"
$LogFile = "$LogDir\RSU_Install.log"
$WinTitle = "Right-Suite"   # partial match — handles ® encoding differences

if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }

function Write-Log {
    param([string]$Message)
    $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    Write-Output $entry
    Add-Content -Path $LogFile -Value $entry
}

function Wait-InstallerWindow {
    param([int]$TimeoutSec = 60)
    Write-Log "Waiting for installer window..."
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        $desktop = [System.Windows.Automation.AutomationElement]::RootElement
        $all = $desktop.FindAll(
            [System.Windows.Automation.TreeScope]::Children,
            [System.Windows.Automation.Condition]::TrueCondition
        )
        foreach ($w in $all) {
            if ($w.Current.Name -like "*$WinTitle*") { return $w }
        }
        Start-Sleep -Milliseconds 500
    }
    return $null
}

function Find-Descendant {
    param($Root, [string]$Name)
    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty, $Name)
    return $Root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond)
}

function Find-DescendantLike {
    param($Root, [string]$Partial)
    $all = $Root.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition
    )
    foreach ($el in $all) {
        if ($el.Current.Name -like "*$Partial*") { return $el }
    }
    return $null
}

function Invoke-Element {
    param($Element)
    try {
        $p = $Element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
        $p.Invoke()
        return $true
    } catch {
        return $false
    }
}

function Select-RadioButton {
    param($Element)
    try {
        $p = $Element.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
        $p.Select()
        return $true
    } catch {
        # Fallback: invoke it like a button
        return (Invoke-Element $Element)
    }
}

function Click-Named {
    param($Window, [string]$Name)
    $el = Find-Descendant $Window $Name
    if (-not $el) { $el = Find-DescendantLike $Window $Name }
    if ($el) {
        $result = Select-RadioButton $el
        Start-Sleep -Milliseconds 300
        return $result
    }
    Write-Log "WARNING: Could not find control '$Name'"
    return $false
}

function Click-Next {
    param($Window)
    $el = Find-DescendantLike $Window "Next"
    if ($el) {
        Invoke-Element $el | Out-Null
        Start-Sleep -Milliseconds 800
        return $true
    }
    Write-Log "WARNING: Next button not found"
    return $false
}

function Click-Finish {
    param($Window)
    $el = Find-DescendantLike $Window "Finish"
    if ($el) {
        Invoke-Element $el | Out-Null
        return $true
    }
    return $false
}

# ── Main ──────────────────────────────────────────────────────────────────────

if (-not (Test-Path $ExePath)) {
    Write-Log "ERROR: Installer not found at $ExePath"
    exit 1
}

Write-Log "Launching installer: $ExePath"
Start-Process -FilePath $ExePath

$win = Wait-InstallerWindow -TimeoutSec 60
if (-not $win) {
    Write-Log "ERROR: Installer window never appeared"
    exit 1
}
Write-Log "Installer window found: $($win.Current.Name)"

Start-Sleep -Seconds 1

# Screen 1 — License Agreement
Write-Log "Screen 1: Accepting license agreement"
Click-Named $win "I ACCEPT the terms of this License Agreement" | Out-Null
Start-Sleep -Milliseconds 500
Click-Next $win | Out-Null

Start-Sleep -Seconds 1

# Screen 2 — Installation Configuration (Standalone already selected)
Write-Log "Screen 2: Selecting Standalone"
Click-Named $win "Standalone" | Out-Null
Start-Sleep -Milliseconds 300
Click-Next $win | Out-Null

Start-Sleep -Seconds 1

# Screen 3 — Availability (Anyone already selected)
Write-Log "Screen 3: Selecting Anyone"
Click-Named $win "Anyone" | Out-Null
Start-Sleep -Milliseconds 300
Click-Next $win | Out-Null

Start-Sleep -Seconds 1

# Screen 4 — Installation type (Typical already selected)
Write-Log "Screen 4: Selecting Typical"
Click-Named $win "Typical" | Out-Null
Start-Sleep -Milliseconds 300
Click-Next $win | Out-Null

Start-Sleep -Seconds 1

# Screen 5 — Install Folders (accept defaults)
Write-Log "Screen 5: Accepting default install folders"
Click-Next $win | Out-Null

Start-Sleep -Seconds 1

# Screen 6 — Likely a confirmation / Begin Install screen
Write-Log "Screen 6: Starting installation"
$installBtn = Find-DescendantLike $win "Install"
if ($installBtn) {
    Write-Log "Clicking Install button"
    Invoke-Element $installBtn | Out-Null
} else {
    Click-Next $win | Out-Null
}

# Wait for installation to complete (up to 10 minutes)
Write-Log "Waiting for installation to complete..."
$deadline = (Get-Date).AddMinutes(10)
$finished = $false
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Seconds 3
    # Re-acquire window reference (it may have changed)
    $win = Wait-InstallerWindow -TimeoutSec 5
    if (-not $win) {
        Write-Log "Installer window closed — install likely complete"
        $finished = $true
        break
    }
    if (Click-Finish $win) {
        Write-Log "Clicked Finish button"
        $finished = $true
        break
    }
    # Handle any unexpected Next buttons that may appear
    $nextEl = Find-DescendantLike $win "Next"
    if ($nextEl) {
        Write-Log "Unexpected Next button found — clicking"
        Invoke-Element $nextEl | Out-Null
    }
}

if ($finished) {
    Write-Log "Installation complete"
    exit 0
} else {
    Write-Log "ERROR: Timed out waiting for installation to finish"
    exit 1
}
