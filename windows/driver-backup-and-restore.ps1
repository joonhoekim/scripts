<#
.SYNOPSIS
    Windows 11 Driver Backup/Restore Tool (DISM-based)
.DESCRIPTION
    Online:     irm https://raw.githubusercontent.com/joonhoekim/scripts/main/windows/driver-backup-and-restore.ps1 | iex
    Standalone: Run restore.ps1 from the backup folder (no network required)
.LINK
    https://github.com/joonhoekim/scripts
#>

# Auto-elevate to Administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $scriptPath = if ($PSCommandPath) { $PSCommandPath } else {
        # irm | iex: save to temp and re-run
        $tmpFile = Join-Path $env:TEMP "driver-backup-and-restore.ps1"
        $MyInvocation.MyCommand.ScriptBlock.ToString() | Set-Content -Path $tmpFile -Encoding UTF8
        $tmpFile
    }
    Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`""
    exit
}

$ErrorActionPreference = 'Stop'

function Read-Path {
    param(
        [string]$Prompt,
        [string]$Default,
        [switch]$MustExist
    )
    Write-Host "  Tip: You can drag & drop a folder here to paste its path." -ForegroundColor DarkGray
    Write-Host "  Example: C:\Users\me\DriverBackup, E:\Backup" -ForegroundColor DarkGray

    while ($true) {
        $input_ = Read-Host "$Prompt (default: $Default)"
        $path = if ($input_.Trim()) { $input_.Trim() } else { $Default }

        # Strip quotes from drag-and-drop
        $path = $path.Trim('"').Trim("'")

        if ($MustExist) {
            if (Test-Path $path) { return $path }
            Write-Host "  Path not found: $path" -ForegroundColor Red
        } else {
            # New path — just verify the drive exists
            $root = [System.IO.Path]::GetPathRoot($path)
            if ($root -and (Test-Path $root)) { return $path }
            Write-Host "  Drive not found: $root" -ForegroundColor Red
        }
        Write-Host "  Please try again." -ForegroundColor Yellow
    }
}

function Show-Menu {
    Clear-Host
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "  Windows Driver Backup/Restore Tool (DISM)" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [1] Backup drivers"
    Write-Host "  [2] Restore drivers (from existing backup)"
    Write-Host "  [Q] Quit"
    Write-Host ""
    $choice = Read-Host "Select"
    return $choice.Trim().ToUpper()
}

function Get-ThirdPartyDrivers {
    Write-Host "`nScanning installed 3rd-party drivers..." -ForegroundColor Yellow
    $drivers = Get-WindowsDriver -Online |
        Select-Object Driver, ClassName, ProviderName, Version, Date
    return $drivers
}

function Invoke-Backup {
    # Let user choose backup destination
    $defaultRoot = "$env:USERPROFILE\DriverBackup"
    Write-Host ""
    Write-Host "Available drives:" -ForegroundColor Yellow
    Get-Volume | Where-Object { $_.DriveLetter -and $_.DriveType -ne 'CD-ROM' } |
        Format-Table DriveLetter, FileSystemLabel, @{N='FreeGB';E={[math]::Round($_.SizeRemaining/1GB,1)}}, DriveType -AutoSize |
        Out-String | Write-Host

    $backupRoot = Read-Path -Prompt "Backup destination folder" -Default $defaultRoot

    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $dest = Join-Path $backupRoot "${hostname}_${timestamp}"
    $driverDir = Join-Path $dest "drivers"

    New-Item -Path $driverDir -ItemType Directory -Force | Out-Null

    # Check disk space on target drive
    $resolved = Resolve-Path $backupRoot -ErrorAction SilentlyContinue
    $driveLetter = if ($resolved) { $resolved.Drive.Name } else { $backupRoot.Substring(0, 1) }
    $driveInfo = Get-PSDrive -Name $driveLetter -ErrorAction SilentlyContinue
    if ($driveInfo -and $driveInfo.Free -lt 1GB) {
        $freeMB = [math]::Round($driveInfo.Free / 1MB, 0)
        Write-Host "`nWARNING: Drive ${driveLetter}: has only ${freeMB} MB free!" -ForegroundColor Red
        $proceed = Read-Host "Continue anyway? [y/N]"
        if (-not $proceed -or $proceed.ToUpper() -ne 'Y') {
            Write-Host "Cancelled." -ForegroundColor Yellow
            return
        }
    }

    # Show driver list
    $drivers = Get-ThirdPartyDrivers
    Write-Host "`nFound $($drivers.Count) drivers:" -ForegroundColor Green
    $drivers | Format-Table -AutoSize | Out-String | Write-Host

    # Confirm before export
    Write-Host "Destination: $dest" -ForegroundColor Cyan
    $confirm = Read-Host "Proceed with backup? [Y/n]"
    if ($confirm -and $confirm.ToUpper() -ne 'Y') {
        Write-Host "Cancelled." -ForegroundColor Yellow
        return
    }

    # Export drivers (with log)
    $logPath = Join-Path $dest "backup-log.txt"
    Write-Host "`nExporting drivers to: $driverDir" -ForegroundColor Yellow
    Write-Host "Log file: $logPath" -ForegroundColor DarkGray
    dism /online /export-driver /destination:"$driverDir" | Tee-Object -FilePath $logPath

    if ($LASTEXITCODE -ne 0) {
        Write-Host "DISM export failed with exit code $LASTEXITCODE (see $logPath)" -ForegroundColor Red
        return
    }

    # Save driver inventory
    $inventoryPath = Join-Path $dest "driver-inventory.txt"
    @(
        "Driver Backup Inventory"
        "Computer: $hostname"
        "Date:     $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "OS:       $(Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty Caption)"
        ""
        ($drivers | Format-Table -AutoSize | Out-String)
    ) | Set-Content -Path $inventoryPath -Encoding UTF8

    # Generate standalone restore script
    $restoreScript = @'
<#
.SYNOPSIS
    Standalone driver restore - no network required
    Auto-generated by Driver Backup Tool
    https://github.com/joonhoekim/scripts
#>

# Auto-elevate to Administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`""
    exit
}

$ErrorActionPreference = 'Stop'

try {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    $driverDir = Join-Path $scriptDir "drivers"

    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "  Driver Restore (Standalone)" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan

    # Show inventory if available
    $inventoryPath = Join-Path $scriptDir "driver-inventory.txt"
    if (Test-Path $inventoryPath) {
        Write-Host ""
        Get-Content $inventoryPath | Select-Object -First 4 | Write-Host
    }

    if (-not (Test-Path $driverDir)) {
        Write-Host "`nERROR: drivers folder not found at $driverDir" -ForegroundColor Red
        Write-Host "Make sure this script is in the same folder as the 'drivers' directory." -ForegroundColor Yellow
        pause
        exit 1
    }

    $infFiles = Get-ChildItem -Path $driverDir -Filter "*.inf" -Recurse
    Write-Host "`nFound $($infFiles.Count) driver packages (.inf)" -ForegroundColor Green
    Write-Host "Source: $driverDir"
    Write-Host ""

    $confirm = Read-Host "Proceed with restore? [Y/n]"
    if ($confirm -and $confirm.ToUpper() -ne 'Y') {
        Write-Host "Cancelled." -ForegroundColor Yellow
        pause
        exit 0
    }

    Write-Host "`nRestoring drivers..." -ForegroundColor Yellow
    dism /online /add-driver /driver:"$driverDir" /recurse /forceunsigned

    if ($LASTEXITCODE -eq 0) {
        Write-Host "`nAll drivers restored successfully!" -ForegroundColor Green
    } else {
        Write-Host "`nDISM finished with exit code $LASTEXITCODE" -ForegroundColor Yellow
        Write-Host "Some drivers may have been skipped (already installed or not applicable)." -ForegroundColor Yellow
    }
} catch {
    Write-Host ""
    Write-Host "====== ERROR ======" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Location: $($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor DarkGray
    Write-Host "Command:  $($_.InvocationInfo.Line.Trim())" -ForegroundColor DarkGray
    Write-Host "===================" -ForegroundColor Red
}

pause
'@

    $restorePath = Join-Path $dest "restore.ps1"
    $restoreScript | Set-Content -Path $restorePath -Encoding UTF8

    # Generate .bat launcher (bypasses execution policy)
    $batPath = Join-Path $dest "restore.bat"
    '@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0restore.ps1"' | Set-Content -Path $batPath -Encoding ASCII

    # Summary
    $size = (Get-ChildItem -Path $dest -Recurse | Measure-Object -Property Length -Sum).Sum
    $sizeMB = [math]::Round($size / 1MB, 1)

    Write-Host ""
    Write-Host "=== Backup Complete ===" -ForegroundColor Green
    Write-Host "  Location:    $dest"
    Write-Host "  Total size:  $sizeMB MB"
    Write-Host "  Drivers:     $($drivers.Count)"
    Write-Host ""
    Write-Host "To restore on a clean install:" -ForegroundColor Cyan
    Write-Host "  1. Copy '${hostname}_${timestamp}' folder to USB" -ForegroundColor Cyan
    Write-Host "  2. Double-click restore.bat (auto-elevates to Admin)" -ForegroundColor Cyan
    Write-Host ""

    $open = Read-Host "Open backup folder in Explorer? [Y/n]"
    if (-not $open -or $open.ToUpper() -eq 'Y') {
        explorer.exe $dest
    }
}

function Invoke-Restore {
    Write-Host ""
    Write-Host "Enter the path to a backup folder or a folder containing backups." -ForegroundColor Yellow

    $defaultRoot = "$env:USERPROFILE\DriverBackup"
    $restoreRoot = Read-Path -Prompt "Path" -Default $defaultRoot -MustExist

    # Check if the given path itself is a backup (has drivers/ subfolder)
    $directDriverDir = Join-Path $restoreRoot "drivers"
    if (Test-Path $directDriverDir) {
        Write-Host "`nFound drivers folder directly in: $restoreRoot" -ForegroundColor Green
        $restorePath = Join-Path $restoreRoot "restore.ps1"
        if (Test-Path $restorePath) {
            & $restorePath
        } else {
            Write-Host "Restoring from $directDriverDir ..." -ForegroundColor Yellow
            dism /online /add-driver /driver:"$directDriverDir" /recurse /forceunsigned
            if ($LASTEXITCODE -eq 0) {
                Write-Host "`nAll drivers restored successfully!" -ForegroundColor Green
            } else {
                Write-Host "`nDISM finished with exit code $LASTEXITCODE" -ForegroundColor Yellow
                Write-Host "Some drivers may have been skipped (already installed or not applicable)." -ForegroundColor Yellow
            }
        }
        return
    }

    # Otherwise list available backups in the folder
    $backups = Get-ChildItem -Path $restoreRoot -Directory | Sort-Object Name -Descending
    $backups = $backups | Where-Object { Test-Path (Join-Path $_.FullName "drivers") }

    if ($backups.Count -eq 0) {
        Write-Host "No backup folders (with 'drivers' subfolder) found in: $restoreRoot" -ForegroundColor Red
        return
    }

    Write-Host "`nAvailable backups:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $backups.Count; $i++) {
        $infCount = (Get-ChildItem -Path (Join-Path $backups[$i].FullName "drivers") -Filter "*.inf" -Recurse -ErrorAction SilentlyContinue).Count
        Write-Host "  [$($i+1)] $($backups[$i].Name)  ($infCount drivers)"
    }

    $sel = Read-Host "`nSelect backup number"
    $idx = [int]$sel - 1
    if ($idx -lt 0 -or $idx -ge $backups.Count) {
        Write-Host "Invalid selection." -ForegroundColor Red
        return
    }

    $selectedDir = $backups[$idx].FullName
    $restorePath = Join-Path $selectedDir "restore.ps1"
    if (Test-Path $restorePath) {
        & $restorePath
    } else {
        $driverDir = Join-Path $selectedDir "drivers"
        Write-Host "Restoring from $driverDir ..." -ForegroundColor Yellow
        dism /online /add-driver /driver:"$driverDir" /recurse /forceunsigned
        if ($LASTEXITCODE -eq 0) {
            Write-Host "`nAll drivers restored successfully!" -ForegroundColor Green
        } else {
            Write-Host "`nDISM finished with exit code $LASTEXITCODE" -ForegroundColor Yellow
            Write-Host "Some drivers may have been skipped (already installed or not applicable)." -ForegroundColor Yellow
        }
    }
}

# --- Main ---
try {
    do {
        $choice = Show-Menu
        switch ($choice) {
            '1' { Invoke-Backup; pause }
            '2' { Invoke-Restore; pause }
            'Q' { break }
            default { Write-Host "Invalid option." -ForegroundColor Red; Start-Sleep 1 }
        }
    } while ($choice -ne 'Q')
} catch {
    Write-Host ""
    Write-Host "====== ERROR ======" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Location: $($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor DarkGray
    Write-Host "Command:  $($_.InvocationInfo.Line.Trim())" -ForegroundColor DarkGray
    Write-Host "===================" -ForegroundColor Red
    pause
}