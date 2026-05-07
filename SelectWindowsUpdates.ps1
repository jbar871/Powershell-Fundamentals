#Requires -RunAsAdministrator

function Search-AvailableUpdates {
    Write-Host "`nSearching for available updates, please wait..." -ForegroundColor Cyan

    $session  = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $searcher.Online = $true   # force live check, skip local cache

    $allUpdates = [System.Collections.Generic.List[object]]::new()

    foreach ($query in @("IsInstalled=0 and Type='Software'", "IsInstalled=0 and Type='Driver'")) {
        try {
            $result = $searcher.Search($query)
            foreach ($u in $result.Updates) { $allUpdates.Add($u) }
        } catch {
            Write-Host "Search failed for query '$query': $_" -ForegroundColor Yellow
        }
    }

    # Deduplicate by update identity
    $seen = @{}
    $unique = [System.Collections.Generic.List[object]]::new()
    foreach ($u in $allUpdates) {
        $id = $u.Identity.UpdateID
        if (-not $seen.ContainsKey($id)) {
            $seen[$id] = $true
            $unique.Add($u)
        }
    }

    Write-Host "Online search complete. Found $($unique.Count) update(s)." -ForegroundColor Cyan
    return $unique
}

function Show-UpdateList {
    param([object]$Updates)

    $required = [System.Collections.Generic.List[object]]::new()
    $optional = [System.Collections.Generic.List[object]]::new()

    foreach ($u in $Updates) {
        # AutoSelectOnWebSites maps to the "Important/Recommended" flag
        if ($u.AutoSelectOnWebSites) { $required.Add($u) } else { $optional.Add($u) }
    }

    $index = 1
    $map   = @{}   # display number -> update object

    if ($required.Count -gt 0) {
        Write-Host "`n--- Required / Recommended Updates ---" -ForegroundColor Yellow
        foreach ($u in $required) {
            $kb   = if ($u.KBArticleIDs.Count) { "(KB$($u.KBArticleIDs[0]))" } else { "" }
            $size = Format-Size $u.MaxDownloadSize
            Write-Host ("  [{0,2}] {1} {2}  [{3}]" -f $index, $u.Title, $kb, $size) -ForegroundColor White
            $map[$index] = $u
            $index++
        }
    }

    if ($optional.Count -gt 0) {
        Write-Host "`n--- Optional Updates ---" -ForegroundColor Cyan
        foreach ($u in $optional) {
            $kb   = if ($u.KBArticleIDs.Count) { "(KB$($u.KBArticleIDs[0]))" } else { "" }
            $size = Format-Size $u.MaxDownloadSize
            Write-Host ("  [{0,2}] {1} {2}  [{3}]" -f $index, $u.Title, $kb, $size) -ForegroundColor Gray
            $map[$index] = $u
            $index++
        }
    }

    Write-Host ""
    return $map
}

function Format-Size {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N1} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N1} MB" -f ($Bytes / 1MB) }
    return "{0:N0} KB" -f ($Bytes / 1KB)
}

function Read-Selection {
    param([int]$Max)

    while ($true) {
        Write-Host "Enter update numbers to install (e.g. 1,3,5), 'all', or 'q' to quit:" -ForegroundColor Cyan
        $raw = Read-Host "Selection"

        if ($raw.Trim() -eq 'q') { return $null }

        if ($raw.Trim() -eq 'all') { return 1..$Max }

        $parts   = $raw -split '[,\s]+' | Where-Object { $_ -ne '' }
        $numbers = foreach ($p in $parts) {
            if ($p -match '^\d+$') { [int]$p }
            else {
                Write-Host "  '$p' is not a valid number." -ForegroundColor Red
                break
            }
        }

        if ($numbers.Count -ne $parts.Count) { continue }

        $bad = $numbers | Where-Object { $_ -lt 1 -or $_ -gt $Max }
        if ($bad) {
            Write-Host "  Out-of-range: $($bad -join ', ')  (valid: 1-$Max)" -ForegroundColor Red
            continue
        }

        return $numbers | Select-Object -Unique | Sort-Object
    }
}

function Install-Updates {
    param(
        [hashtable]$Map,
        [int[]]$Selected
    )

    $session    = New-Object -ComObject Microsoft.Update.Session
    $collection = New-Object -ComObject Microsoft.Update.UpdateColl

    foreach ($n in $Selected) {
        $u = $Map[$n]
        if (-not $u.EulaAccepted) { $u.AcceptEula() }
        $collection.Add($u) | Out-Null
    }

    # Download
    Write-Host "`nDownloading $($collection.Count) update(s)..." -ForegroundColor Cyan
    $dl          = $session.CreateUpdateDownloader()
    $dl.Updates  = $collection
    $dl.Download() | Out-Null

    # Install
    Write-Host "Installing..." -ForegroundColor Cyan
    $inst          = $session.CreateUpdateInstaller()
    $inst.Updates  = $collection
    $installResult = $inst.Install()

    # Results
    Write-Host "`n=== Results ===" -ForegroundColor Green
    for ($i = 0; $i -lt $collection.Count; $i++) {
        $code   = $installResult.GetUpdateResult($i).ResultCode
        $status = switch ($code) {
            0 { "Not Started" }
            1 { "In Progress" }
            2 { "Succeeded"   }
            3 { "Succeeded (with errors)" }
            4 { "Failed"      }
            5 { "Aborted"     }
            default { "Unknown ($code)" }
        }
        $color = if ($code -eq 2) { "Green" } elseif ($code -eq 3) { "Yellow" } else { "Red" }
        Write-Host ("  {0}: {1}" -f $collection.Item($i).Title, $status) -ForegroundColor $color
    }

    if ($installResult.RebootRequired) {
        Write-Host "`nReboot required to complete installation." -ForegroundColor Yellow
    } else {
        Write-Host "`nInstallation complete. No reboot needed." -ForegroundColor Green
    }
}

# ── Main ──────────────────────────────────────────────────────────────────────

$updates = Search-AvailableUpdates
if (-not $updates -or $updates.Count -eq 0) {
    Write-Host "No updates available." -ForegroundColor Green
    return
}

Write-Host "Found $($updates.Count) update(s)." -ForegroundColor Cyan

$map = Show-UpdateList -Updates $updates

$selected = Read-Selection -Max $map.Count
if (-not $selected) {
    Write-Host "Exiting without installing." -ForegroundColor Gray
    return
}

Write-Host "`nSelected: $($selected -join ', ')" -ForegroundColor White
Install-Updates -Map $map -Selected $selected
