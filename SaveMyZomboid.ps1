<#
PowerShell v5.1 - Project Zomboid Save Manager
Features:
- Arrow-key interactive TUI (A3)
- Start game (steam://rungameid/108600) only when not running
- Load saves (flattened list) -- NOT available when game running
- Restores: backs up current active save to "latest recovered (auto-backup before last restore)" then restores chosen slot
- Background autosave job monitors game process and rotates backup slots only when tracked files changed:
  map.bin, statistic.bin, vehicles.db, players.db
- Multi-threaded Robocopy (/MT) used for copying
#>

# ===========================
# CONFIG. Edit config.json
# ===========================

$script:configFile = "$PSScriptRoot\config.json"

function get-config{
    param($file)

    $script:config = Get-Content $file | ConvertFrom-Json

    $Script:ZFolder        = $ExecutionContext.InvokeCommand.ExpandString($config.ZFolder)
    $Script:SaveFolder     = (Join-Path $Script:ZFolder $config.SaveFolder)
    $Script:BackupFolder   = (Join-Path $Script:ZFolder $config.BackupFolder)
    $Script:SaveCount      = $config.SaveCount
    $Script:SaveFrequency  = $config.SaveFrequency
    $Script:RobocopyThreads= $config.RobocopyThreads
    $Script:RobocopyOptions= $config.RobocopyOptions
    $Script:TrackedFiles   = $config.TrackedFiles
    $Script:SteamURI       = $config.SteamURI
    $Script:AutoJobName    = $config.AutoJobName   
}

function set-config{
    param($file)

    $script:config.SaveCount       = $Script:SaveCount
    $script:config.SaveFrequency   = $Script:SaveFrequency
    
    $config | ConvertTo-Json | Out-File $script:configFile
    
    get-config -file $script:configFile
}

get-config -file $script:configFile

# ===========================
# Basic validation & create folders
# ===========================
if (-not (Test-Path $SaveFolder -PathType Container)) {
    Write-Host "SaveFolder '$SaveFolder' does not exist. Please check $SaveFolder." -ForegroundColor Red
    pause
    exit 1
}
if (-not (Test-Path $BackupFolder -PathType Container)) {
    New-Item -Path $BackupFolder -ItemType Directory -Force | Out-Null
}

# Fingerprint store path
$Script:FingerprintFile = Join-Path $BackupFolder "fingerprints.json"
if (-not (Test-Path $FingerprintFile)) { '{}' | Out-File $FingerprintFile }

# ===========================
# Utility functions
# ===========================
function Log {
    param([string]$m)
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Write-Output "[$ts] $m"
    "[$ts] $m" | Out-file -FilePath "$BackupFolder\log.log" -Append
}

function Run-Robocopy {
    param([string]$src,[string]$dst)
    #$opts = $Script:RobocopyOptions
    # Ensure target folder exists
    if (-not (Test-Path $dst)) { New-Item -Path $dst -ItemType Directory -Force | Out-Null }
    & Robocopy.exe $src $dst @RoboCopyOptions | Out-Null
    return $LASTEXITCODE
}

function Get-ProcessRunning {
    # returns $true if any ProjectZomboid process is running (uses wildcard)
    $p = Get-Process -Name "ProjectZomboid*" -ErrorAction SilentlyContinue
    return $p -ne $null
}

function Load-Fingerprint {
    try {
        return (Get-Content $FingerprintFile -Raw | ConvertFrom-Json -ErrorAction Stop)
    } catch {
        return @{}
    }
}

function Save-Fingerprint($obj) {
    $obj | ConvertTo-Json -Depth 10 | Out-File $FingerprintFile -Force
}

function Get-SaveFingerprint($SavePath) {
    $result = @{}
    foreach ($f in $Script:TrackedFiles) {
        $file = Join-Path $SavePath $f
        if (Test-Path $file) {
            $ts = (Get-Item $file).LastWriteTimeUtc.Ticks
            $result[$f] = $ts
        } else {
            $result[$f] = 0
        }
    }
    return $result
}

function Has-Changed($Old, $New) {
    if (-not $Old) { return $true }
    foreach ($k in $New.Keys) {
        if (-not $Old.ContainsKey($k)) { return $true }
        if ($New[$k] -ne $Old[$k]) { return $true }
    }
    return $false
}

function Ensure-Initial-Slots-And-Fingerprints {
    $fp = Load-Fingerprint
    $worldDirs = Get-ChildItem -Path $SaveFolder -Directory -ErrorAction SilentlyContinue
    foreach ($world in $worldDirs) {
        $saves = Get-ChildItem -Path (Join-Path $SaveFolder $world.Name) -Directory -ErrorAction SilentlyContinue
        foreach ($save in $saves) {
            $target = Join-Path $BackupFolder "$($world.Name)\$($save.Name)\1"
            if (-not (Test-Path $target)) {
                # ensure parent
                $parent = Split-Path $target -Parent
                if (-not (Test-Path $parent)) { New-Item -Path $parent -ItemType Directory -Force | Out-Null }
                Log "Initial copy -> $($world.Name)/$($save.Name) to slot 1"
                Run-Robocopy $save.FullName $target | Out-Null
            }
            $key = "$($world.Name)/$($save.Name)"
            if (-not ($fp.$key -ne $null)) {
                $fp | Add-Member -MemberType NoteProperty -Name $key -Value (Get-SaveFingerprint $save.FullName)
            }
        }
    }
    Save-Fingerprint $fp
}

# ===========================
# Background job scriptblock (autosave engine)
# ===========================
$jobScript = {
    param($SaveFolder,$BackupFolder,$SaveCount,$SaveFrequency,$RobocopyOptions,$TrackedFiles,$FingerprintFile)

    function LogJ { param($m) $ts=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Write-Output "Jo[$ts] $m"; "JOB: [$ts] $m" | Out-file -FilePath "$BackupFolder\log.log" -Append }

    function Run-RobocopyJ {
        param($src,$dst)
        #$opts = $RobocopyOptions + " /MT:$RobocopyThreads" + " /TEE"
        if (-not (Test-Path $dst)) { New-Item -Path $dst -ItemType Directory -Force | Out-Null }
        & Robocopy.exe $src $dst @RoboCopyOptions | Out-Null
        return $LASTEXITCODE
    }

    function Load-FingerprintJ {
        try { return (Get-Content $FingerprintFile -Raw | ConvertFrom-Json -ErrorAction Stop) } catch { return @{} }
    }
    function Save-FingerprintJ($obj) { $obj | ConvertTo-Json -Depth 10 | Out-File $FingerprintFile -Force }

    function Get-SaveFingerprintJ($SavePath) {
        $result = [pscustomobject]@{}
        foreach ($f in $TrackedFiles) {
            $file = Join-Path $SavePath $f
            if (Test-Path $file) { $ts=(Get-Item $file).LastWriteTimeUtc.Ticks; $result | Add-Member -MemberType NoteProperty -Name $f -Value $ts } else { $result | Add-Member -MemberType NoteProperty -Name $f -Value 0 }
        }
        return $result
    }

    # ensure fingerprint file exists
    if (-not (Test-Path $FingerprintFile)) { '{}' | Out-File $FingerprintFile -Force }
    # initial run: create slot1 if missing
    $fp = Load-FingerprintJ
    $worldDirs = Get-ChildItem -Path $SaveFolder -Directory -ErrorAction SilentlyContinue
    foreach ($world in $worldDirs) {
        $saves = Get-ChildItem -Path (Join-Path $SaveFolder $world.Name) -Directory -ErrorAction SilentlyContinue
        foreach ($save in $saves) {
            $target = Join-Path $BackupFolder "$($world.Name)\$($save.Name)\1"
            if (-not (Test-Path $target)) {
                $parent = Split-Path $target -Parent
                if (-not (Test-Path $parent)) { New-Item -Path $parent -ItemType Directory -Force | Out-Null }
                LogJ "Initial copy (job) -> $($world.Name)/$($save.Name) to slot 1"
                Run-RobocopyJ $save.FullName $target
            }
            $key = "$($world.Name)/$($save.Name)"
            if (-not ($fp.$key -ne $null)) {
                $fp | Add-Member -MemberType NoteProperty -Name $key -Value (Get-SaveFingerprint $save.FullName)
            }
        }
    }
    Save-FingerprintJ $fp

    while ($true) {
        $isRunning = (Get-Process -Name "ProjectZomboid*" -ErrorAction SilentlyContinue) -ne $null
        if ($isRunning) {
            $fp = Load-FingerprintJ
            $worldDirs = Get-ChildItem -Path $SaveFolder -Directory -ErrorAction SilentlyContinue
            foreach ($world in $worldDirs) {
                $saves = Get-ChildItem -Path (Join-Path $SaveFolder $world.Name) -Directory -ErrorAction SilentlyContinue
                foreach ($save in $saves) {
                    $key = "$($world.Name)/$($save.Name)"
                    $current = Get-SaveFingerprintJ $save.FullName
                    $old = $null
                    if ($fp.$key -ne $null) { $old = $fp.$key }
                    #if (Has-ChangedJ $old $current) {
                    $hasChanged = $false
                    $TrackedFiles | ForEach-Object{If($old.$_ -ne $current.$_){$hasChanged = $true}}
                    
                    if($hasChanged) {
                        # rotate into next numeric slot in BackupFolder\<world>\<save>\n
                        $saveBase = Join-Path $BackupFolder "$($world.Name)\$($save.Name)"
                        if (-not (Test-Path $saveBase)) { New-Item -Path $saveBase -ItemType Directory -Force | Out-Null }

                        # get the existing saves up to $savecount
                        $existing = Get-ChildItem -Path $saveBase -Directory -ErrorAction SilentlyContinue | Where-Object{$_.Name -in (1..$SaveCount)} | Sort-Object LastWriteTimeUtc
                        $newest = $existing | select -Last 1
                        $oldest = $existing | select -First 1
                        If($existing.Count -lt $SaveCount){$nextIndex = [int]$existing.Count +1}
                        Else{ $nextIndex = [int]$oldest.Name}
                        $target = Join-Path $saveBase $nextIndex
                        New-Item -Path $target -ItemType Directory -Force | Out-Null
                        LogJ "Job backup: $($world.Name)/$($save.Name) -> slot $nextIndex"
                        #ensure backups are not done too often
                        if((New-TimeSpan -Start $newest.LastWriteTime -End (Get-Date)).TotalSeconds -gt $SaveFrequency*60){
                            $rc = Run-RobocopyJ $save.FullName $target
                            if ($rc -ge 8) { LogJ "Robocopy error code $rc while backing up $key" } else { LogJ "Robocopy ok -> slot $nextIndex" }
                        }

                        # update fingerprint
                        $fp[$key] = $current
                        Save-FingerprintJ $fp
                    }
                }
            }
        }
        Start-Sleep -Seconds ($SaveFrequency * 60)
    }
}

# ===========================
# Start background autosave job if not running
# ===========================
$existingJob = Get-Job -Name $Script:AutoJobName  -ErrorAction SilentlyContinue
if (-not $existingJob) {
    Start-Job -Name $Script:AutoJobName -ScriptBlock $jobScript -ArgumentList $SaveFolder,$BackupFolder,$SaveCount,$SaveFrequency,$Script:RobocopyOptions,$Script:TrackedFiles,$FingerprintFile | Out-Null
    Start-Sleep -Milliseconds 300
    Log "Started background autosave job '$($Script:AutoJobName)'."
} else {
    Log "Background autosave job already present."
}

# Ensure initial slots & fingerprints in foreground as well
Ensure-Initial-Slots-And-Fingerprints

# ===========================
# Interactive arrow-key TUI
# ===========================
# Helper: read single key
function Read-Key {
    $k = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    return $k
}

# Helper: build flattened save list under SaveFolder, returns array of objects:
# @{ Display = "World/Save (slot X...)"; World="..."; SaveName="..."; Path=...; BackupSlots=@(slotnames...) }
function Get-FlattenedSaves {
    $out = @()
    $worlds = Get-ChildItem -Path $SaveFolder -Directory -ErrorAction SilentlyContinue
    foreach ($w in $worlds) {
        $saves = Get-ChildItem -Path (Join-Path $SaveFolder $w.Name) -Directory -ErrorAction SilentlyContinue
        foreach ($s in $saves) {
            $backupBase = Join-Path $BackupFolder "$($w.Name)\$($s.Name)"
            $slots = @()
            if (Test-Path $backupBase) {
                $slotsFolders = Get-ChildItem -Path $backupBase -Directory -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc -Descending
                $slots = $slotsFolders | ForEach-Object { $_.Name }
                $slots
            }
            # include latest recovered if present
            if (Test-Path (Join-Path $backupBase "latest recovered")) {
                # ensure it's listed first by convention
                $slots = @("latest recovered (auto-backup before last restore)") + ($slots | Where-Object { $_ -ne "latest recovered" })
            }
            $display = "$($w.Name)/$($s.Name)"
            $out += [PSCustomObject]@{ Display=$display; World=$w.Name; SaveName=$s.Name; SavePath=$s.FullName; BackupSlots=$slots; BackupBase=$backupBase }
        }
    }
    return $out
}

# Draw menu with arrow navigation
function Show-Menu {
    param([string[]]$items, [int]$selected)
    Clear-Host
    Write-Host "╔════════════════════════════════════════════════════════════════╗"
    Write-Host "║                 PROJECT ZOMBOID SAVE MANAGER                   ║"
    Write-Host "╠════════════════════════════════════════════════════════════════╣"
    for ($i=0; $i -lt $items.Length; $i++) {
        $prefix = "  "
        if ($i -eq $selected) { $prefix = "➤ " }
        Write-Host ("{0}{1}" -f $prefix, $items[$i])
    }
    Write-Host "╚════════════════════════════════════════════════════════════════╝"
    Write-Host ""
    Write-Host "Use ↑↓ to navigate, Enter to select, Esc to exit."
}

# Menu loop
$menuItems = @()  # will be built each iteration because Start Game may appear/disappear
$selected = 0

while ($true) {
    # Rebuild menu dynamic items
    $isRunning = Get-ProcessRunning
    $menuItems = @()

    if (-not $isRunning) {
        $menuItems += "Start Project Zomboid"
    } else {
        $menuItems += "ProjectZomboid process running...."  # non-selectable
    }
    $menuItems += "Load Save"
    $menuItems += "Configure (SaveCount / SaveFrequency)"
    $menuItems += "Exit"

    # Ensure selection index valid
    if ($selected -ge $menuItems.Length) { $selected = 0 }

    Show-Menu -items $menuItems -selected $selected

    # Handle keypress navigation
    $key = Read-Key
    if ($key.VirtualKeyCode -eq 27) {  # ESC
        break
    } elseif ($key.VirtualKeyCode -eq 38) { # Up
        $selected = if ($selected -le 0) { $menuItems.Length - 1 } else { $selected - 1 }
        continue
    } elseif ($key.VirtualKeyCode -eq 40) { # Down
        $selected = if ($selected -ge ($menuItems.Length - 1)) { 0 } else { $selected + 1 }
        continue
    } elseif ($key.VirtualKeyCode -eq 13) { # Enter
        $choice = $menuItems[$selected]
        # Handle non-selectable "running..." line
        if ($choice -eq "ProjectZomboid process running....") {
            # show simple message then continue
            Write-Host "`nProjectZomboid is currently running. Return to menu when game has exited." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            continue
        }

        switch ($choice) {
            "Start Project Zomboid" {
                if (Get-ProcessRunning) {
                    Write-Host "Game already running." -ForegroundColor Yellow
                    Start-Sleep -Seconds 1
                } else {
                    Try {
                        Start-Process $Script:SteamURI
                        Write-Host "Starting Project Zomboid via Steam..." -ForegroundColor Green
                    } catch {
                        Write-Host "Failed to start via Steam URI: $_" -ForegroundColor Red
                    }
                    Start-Sleep -Seconds 2
                }
            }
            "Load Save" {
                # disallow if game running
                if (Get-ProcessRunning) {
                    Write-Host "`nCannot load saves while ProjectZomboid is running." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    continue
                }
                # Show flattened list and let user pick
                $flat = Get-FlattenedSaves
                if (-not $flat -or $flat.Count -eq 0) {
                    Write-Host "No saves found." -ForegroundColor Yellow
                    Start-Sleep -Seconds 2
                    continue
                }
                # Build list of choices: each BackupSlot becomes an item "World/Save :: slot X"
                $choices = @()
                foreach ($entry in $flat) {
                    $base = $entry.BackupBase
                    $slots = @()
                    if (Test-Path $base) {
                        $slots = Get-ChildItem -Path $base -Directory -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc -Descending | ForEach-Object { $_.Name }
                    }
                    # move "latest recovered" to descriptive name if exists
                    if (Test-Path (Join-Path $base "latest recovered")) {
                        $slots = @("latest recovered (auto-backup before last restore)") + ($slots | Where-Object { $_ -ne "latest recovered" })
                    }
                    foreach ($slot in $slots) {
                        $label = "$($entry.Display) :: $slot"
                        $choices += [PSCustomObject]@{ Label=$label; World=$entry.World; SaveName=$entry.SaveName; Slot=$slot; SlotPath=(Join-Path $base $slot) ; SavePath=$entry.SavePath }
                    }
                }
                if ($choices.Count -eq 0) {
                    Write-Host "No backup slots available." -ForegroundColor Yellow
                    Start-Sleep -Seconds 2
                    continue
                }

                # Arrow-key select among choices
                $sel = 0
                while ($true) {
                    Clear-Host
                    Write-Host "Select a backup slot to restore (Esc to cancel):`n"
                    for ($i=0; $i -lt $choices.Count; $i++) {
                        $pfx = "  "
                        if ($i -eq $sel) { $pfx = "➤ " }
                        Write-Host "$pfx$($choices[$i].Label)"
                    }
                    $k = Read-Key
                    if ($k.VirtualKeyCode -eq 27) { break } # cancel
                    elseif ($k.VirtualKeyCode -eq 38) { $sel = if ($sel -le 0) { $choices.Count - 1 } else { $sel - 1 } ; continue }
                    elseif ($k.VirtualKeyCode -eq 40) { $sel = if ($sel -ge ($choices.Count - 1)) { 0 } else { $sel + 1 } ; continue }
                    elseif ($k.VirtualKeyCode -eq 13) {
                        $chosen = $choices[$sel]
                        # Confirm restore
                        Clear-Host
                        Write-Host "You chose to restore: $($chosen.Label)`n" -ForegroundColor Cyan
                        $confirm = Read-Host "Type YES to confirm restore (this will overwrite live save)."
                        if ($confirm -ne "YES") {
                            Write-Host "Cancelled." -ForegroundColor Yellow
                            Start-Sleep -Seconds 1
                            break
                        }

                        # Perform restore sequence:
                        # 1) Backup current active save -> Backup\<world>\<save>\latest recovered (overwrite)
                        $activeLive = $chosen.SavePath
                        If($chosen.SlotPath.endswith('(auto-backup before last restore)')){
                            Log -m "Restoring Latest recovered, placing latest recovered in slot x to enable backup of current"
                            $parent = split-path $chosen.slotPath
                            $chosen.slotPath = "$(split-path $chosen.slotPath)\latest recovered"

                            Rename-Item $chosen.slotPath "$parent\x"
                            $chosen.slotPath = "$parent\x"

                        }
                        Log -m "restoring to $($activeLive)"
                        $backupLatest = Join-Path (Join-Path $Script:BackupFolder "$($chosen.World)\$($chosen.SaveName)") "latest recovered"
                        if (-not (Test-Path (Split-Path $backupLatest -Parent))) { New-Item -Path (Split-Path $backupLatest -Parent) -ItemType Directory -Force | Out-Null }
                        if (Test-Path $activeLive) {
                            Write-Host "Backing up current live save to 'latest recovered'..." -ForegroundColor Yellow
                            # remove existing latest recovered then copy
                            if (Test-Path $backupLatest) { Remove-Item -Path $backupLatest -Recurse -Force -ErrorAction SilentlyContinue }
                            Run-Robocopy $activeLive $backupLatest | Out-Null
                            Write-Host "Backed up live save -> $backupLatest" -ForegroundColor Green
                        } else {
                            Write-Host "Live save folder missing: $activeLive" -ForegroundColor Yellow
                        }

                        # 2) Copy selected backup slot -> live save folder (overwrite)
                        $src = $chosen.SlotPath
                        if (-not (Test-Path $src)) {
                            Write-Host "Chosen slot path not found: $src" -ForegroundColor Red
                            Start-Sleep -Seconds 2
                            break 2
                        }
                        # ensure live save parent exists
                        if (-not (Test-Path (Split-Path $activeLive -Parent))) { New-Item -Path (Split-Path $activeLive -Parent) -ItemType Directory -Force | Out-Null }
                        # remove live content safely, then copy
                        if (Test-Path $activeLive) {
                            # move to temp then delete to avoid partial overwrite issues
                            $tmp = "$activeLive.__tmp_restore__"
                            if (Test-Path $tmp) { Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue }
                            Rename-Item -Path $activeLive -NewName (Split-Path $tmp -Leaf) -ErrorAction SilentlyContinue
                            # ensure parent recreated
                            New-Item -Path $activeLive -ItemType Directory -Force | Out-Null
                        } else {
                            New-Item -Path $activeLive -ItemType Directory -Force | Out-Null
                        }
                        Write-Host "Restoring selected slot to live save folder..." -ForegroundColor Yellow
                        $rc = Run-Robocopy $src $activeLive
                        if ($rc -ge 8) {
                            Write-Host "Robocopy returned error code $rc. Restore may be incomplete." -ForegroundColor Red
                        } else {
                            Write-Host "Restore completed (robocopy $rc)." -ForegroundColor Green
                            # cleanup any tmp moved folder
                            $tmp = "$activeLive.__tmp_restore__"
                            if (Test-Path $tmp) { Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue }
                        }
                        Start-Sleep -Seconds 2

                        If($chosen.slotPath.EndsWith('x')){
                            Remove-Item $chosen.slotPath -Force -Recurse
                        }

                        break
                    }
                } # end slot selection
            }
            "Configure (SaveCount / SaveFrequency)" {
                Clear-Host
                Write-Host "Current SaveCount: $Script:SaveCount"
                $newCount = Read-Host "Enter new SaveCount (or press Enter to keep)"
                if ($newCount -match '^\d+$') {
                    $Script:SaveCount = [int]$newCount
                    Write-Host "SaveCount updated to $Script:SaveCount"
                } elseif ($newCount -ne "") {
                    Write-Host "Invalid input. SaveCount unchanged." -ForegroundColor Yellow
                }

                Write-Host "Current SaveFrequency (minutes): $Script:SaveFrequency"
                $newFreq = Read-Host "Enter new SaveFrequency in minutes (or press Enter to keep)"
                if ($newFreq -match '^\d+$') {
                    $Script:SaveFrequency = [int]$newFreq
                    Write-Host "SaveFrequency updated to $Script:SaveFrequency minutes"
                } elseif ($newFreq -ne "") {
                    Write-Host "Invalid input. SaveFrequency unchanged." -ForegroundColor Yellow
                }
                set-config -file $script:configFile

                # Propagate changes to background job by stopping and restarting job
                $job = Get-Job -Name $Script:AutoJobName -ErrorAction SilentlyContinue
                if ($job) {
                    Write-Host "Restarting autosave background job to apply config..." -ForegroundColor Cyan
                    # Stop job and start new one with updated args
                    Stop-Job -Job $job -ErrorAction SilentlyContinue
                    Remove-Job -Job $job -ErrorAction SilentlyContinue
                    Start-Sleep -Milliseconds 200
                    Start-Job -Name $Script:AutoJobName -ScriptBlock $jobScript -ArgumentList $SaveFolder,$BackupFolder,$Script:SaveCount,$Script:SaveFrequency,$Script:RobocopyOptions,$Script:TrackedFiles,$FingerprintFile | Out-Null
                    Start-Sleep -Milliseconds 200
                    Write-Host "Autosave job restarted." -ForegroundColor Green
                } else {
                    Start-Job -Name $Script:AutoJobName -ScriptBlock $jobScript -ArgumentList $SaveFolder,$BackupFolder,$Script:SaveCount,$Script:SaveFrequency,$Script:RobocopyOptions,$Script:TrackedFiles,$FingerprintFile | Out-Null
                    Start-Sleep -Milliseconds 200
                    Write-Host "Autosave job started." -ForegroundColor Green
                }
                Start-Sleep -Seconds 2
            }
            "Exit" {
                Write-Host "Exiting..." -ForegroundColor Cyan
                break 2
            }
        } # end switch
    } # end Enter
    # Small delay to avoid tight loop
    Start-Sleep -Milliseconds 100
} # end while menu

# Cleanup note: Do not remove background job on exit so the service can persist if desired.
# If you want the job removed when exiting, uncomment below:
# $j = Get-Job -Name $Script:AutoJobName -ErrorAction SilentlyContinue
# if ($j) { Stop-Job -Job $j -Force; Remove-Job -Job $j -Force }

Log "Exiting."