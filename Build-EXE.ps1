<#
.SYNOPSIS
    PC Cleaner - Safe Windows cleanup and optimization script.
.DESCRIPTION
    Cleans temp files, browser caches, Windows Update cache, Recycle Bin,
    error logs, flushes DNS, and optimizes drives. Saves a log to the Desktop.
#>

$ErrorActionPreference = 'SilentlyContinue'

# --- Self-elevate if not admin ---
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Start-Process PowerShell -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

try {

# --- Helpers ---

function Write-Header([string]$text) {
    Write-Host "`n================================================" -ForegroundColor Cyan
    Write-Host "  $text" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
}

function Write-OK([string]$msg)   { Write-Host "  [OK]      $msg" -ForegroundColor Green }
function Write-Skip([string]$msg) { Write-Host "  [SKIPPED] $msg" -ForegroundColor Yellow }
function Write-Fail([string]$msg) { Write-Host "  [ERROR]   $msg" -ForegroundColor Red }
function Write-Info([string]$msg) { Write-Host "  [INFO]    $msg" -ForegroundColor White }

function Remove-Items([string]$path, [string]$label) {
    if (-not (Test-Path $path)) {
        Write-Skip "$label - path not found"
        return
    }
    try {
        $items = Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue
        $count = ($items | Measure-Object).Count
        Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-OK "$label - $count item(s) removed"
    } catch {
        Write-Fail "$label - $_"
    }
}

function Get-DiskFreeGB([string]$driveLetter) {
    $disk = Get-PSDrive -Name $driveLetter -ErrorAction SilentlyContinue
    if ($disk) { return [math]::Round($disk.Free / 1GB, 2) }
    return 0
}

function Clean-BrowserProfiles([string]$userDataRoot, [string]$browser) {
    if (-not (Test-Path $userDataRoot)) {
        Write-Skip "$browser - not installed"
        return
    }
    $profiles = Get-ChildItem -Path $userDataRoot -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^(Default|Profile \d+)$' }
    if (-not $profiles) {
        Write-Skip "$browser - no profiles found"
        return
    }
    foreach ($profile in $profiles) {
        $label = if ($profile.Name -eq 'Default') { $browser } else { "$browser [$($profile.Name)]" }
        Remove-Items -path "$($profile.FullName)\Cache"      -label "$label Cache"
        Remove-Items -path "$($profile.FullName)\Code Cache" -label "$label Code Cache"
    }
}

# --- Startup ---

$winDrive    = $env:SystemDrive.TrimEnd('\')
$driveLetter = $winDrive.TrimEnd(':')

$logPath = "$env:TEMP\PC-Clean-Log.txt"
try { Start-Transcript -Path $logPath -Force | Out-Null } catch { $logPath = "unavailable" }

Write-Host ""
Write-Host "  PC CLEANER" -ForegroundColor Magenta
Write-Host "  Safe Windows cleanup and optimization" -ForegroundColor Magenta
Write-Host ""
Write-Host "  Date  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White
Write-Host "  PC    : $env:COMPUTERNAME" -ForegroundColor White
Write-Host "  User  : $env:USERNAME" -ForegroundColor White
Write-Host "  Drive : $winDrive" -ForegroundColor White
Write-Host "  Log   : $logPath" -ForegroundColor DarkGray
Write-Host ""

$freeBefore = Get-DiskFreeGB $driveLetter
Write-Info "Disk $winDrive free space before: $freeBefore GB"

# --- 1. User Temp ---

Write-Header "1. User Temp Files"
Remove-Items -path $env:TEMP -label "User TEMP"

# --- 2. System Temp ---

Write-Header "2. System Temp Files"
Remove-Items -path "$env:SystemRoot\Temp" -label "Windows Temp"

# --- 3. Prefetch ---

Write-Header "3. Prefetch Cache"
Remove-Items -path "$env:SystemRoot\Prefetch" -label "Windows Prefetch"

# --- 4. Windows Update cache ---

Write-Header "4. Windows Update Download Cache"
try {
    Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
    Write-Info "Windows Update service stopped"
    Remove-Items -path "$env:SystemRoot\SoftwareDistribution\Download" -label "WU Download Cache"
    Start-Service -Name wuauserv -ErrorAction SilentlyContinue
    Write-Info "Windows Update service restarted"
} catch {
    Write-Fail "Windows Update cache - $_"
    Start-Service -Name wuauserv -ErrorAction SilentlyContinue
}

# --- 5. Recycle Bin ---

Write-Header "5. Recycle Bin"
try {
    Clear-RecycleBin -Force -ErrorAction SilentlyContinue
    Write-OK "Recycle Bin emptied"
} catch {
    Write-Fail "Recycle Bin - $_"
}

# --- 6. Browser Caches ---

Write-Header "6. Browser Caches"

Clean-BrowserProfiles "$env:LOCALAPPDATA\Google\Chrome\User Data"               "Chrome"
Clean-BrowserProfiles "$env:LOCALAPPDATA\Microsoft\Edge\User Data"              "Edge"
Clean-BrowserProfiles "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data" "Brave"

$ffProfilesRoot = "$env:APPDATA\Mozilla\Firefox\Profiles"
if (Test-Path $ffProfilesRoot) {
    $ffProfiles = Get-ChildItem -Path $ffProfilesRoot -Directory -ErrorAction SilentlyContinue
    if ($ffProfiles) {
        foreach ($profile in $ffProfiles) {
            Remove-Items -path "$($profile.FullName)\cache2"       -label "Firefox cache2 [$($profile.Name)]"
            Remove-Items -path "$($profile.FullName)\OfflineCache" -label "Firefox OfflineCache [$($profile.Name)]"
            Remove-Items -path "$($profile.FullName)\thumbnails"   -label "Firefox thumbnails [$($profile.Name)]"
        }
    } else {
        Write-Skip "Firefox - no profiles found"
    }
} else {
    Write-Skip "Firefox - not installed"
}

# --- 7. Windows Error Reporting logs ---

Write-Header "7. Windows Error Reporting Logs"
Remove-Items -path "$env:ProgramData\Microsoft\Windows\WER\ReportArchive" -label "WER Archive"
Remove-Items -path "$env:ProgramData\Microsoft\Windows\WER\ReportQueue"   -label "WER Queue"

# --- 8. DNS Flush ---

Write-Header "8. DNS Cache Flush"
try {
    & ipconfig /flushdns | Out-Null
    Write-OK "DNS cache flushed"
} catch {
    Write-Fail "DNS flush - $_"
}

# --- 9. Drive Optimization ---

Write-Header "9. Drive Optimization ($winDrive)"
try {
    $partition  = Get-Partition -DriveLetter $driveLetter -ErrorAction Stop
    $diskNumber = $partition.DiskNumber
    $physDisk   = Get-PhysicalDisk | Where-Object { $_.DeviceID -eq $diskNumber } | Select-Object -First 1

    if ($physDisk -and $physDisk.MediaType -eq 'SSD') {
        Write-Info "SSD detected - running TRIM"
        Optimize-Volume -DriveLetter $driveLetter -ReTrim -Verbose:$false
        Write-OK "SSD TRIM completed"
    } elseif ($physDisk -and $physDisk.MediaType -eq 'HDD') {
        Write-Info "HDD detected - running Defrag"
        Optimize-Volume -DriveLetter $driveLetter -Defrag -Verbose:$false
        Write-OK "HDD Defrag completed"
    } else {
        Write-Info "Drive type unknown - attempting TRIM"
        Optimize-Volume -DriveLetter $driveLetter -ReTrim -Verbose:$false
        Write-OK "Optimization completed"
    }
} catch {
    Write-Fail "Drive optimization - $_"
}

# --- Final Report ---

Write-Header "DONE - Summary"

$freeAfter = Get-DiskFreeGB $driveLetter
$freedGB   = [math]::Round($freeAfter - $freeBefore, 2)
$freedMB   = [math]::Round($freedGB * 1024, 0)

Write-Host ""
Write-Host "  PC     : $env:COMPUTERNAME ($env:USERNAME)" -ForegroundColor White
Write-Host "  Before : $freeBefore GB free" -ForegroundColor White
Write-Host "  After  : $freeAfter GB free"  -ForegroundColor White

if ($freedGB -gt 0) {
    Write-Host "  Freed  : +$freedGB GB ($freedMB MB)" -ForegroundColor Green
} elseif ($freedGB -eq 0) {
    Write-Host "  Freed  : ~0 GB (locked files skipped or already clean)" -ForegroundColor Yellow
} else {
    Write-Host "  Freed  : $freedGB GB (some space used by logs/optimization)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "  Log saved to: $logPath" -ForegroundColor DarkGray
Write-Host ""

try { Stop-Transcript | Out-Null } catch {}

Read-Host "  Press Enter to exit"

} catch {
    $errLines = @(
        "PC Cleaner - FATAL ERROR",
        "========================",
        "Date : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "PC   : $env:COMPUTERNAME",
        "User : $env:USERNAME",
        "",
        "Error:",
        "$_",
        "",
        "Stack trace:",
        "$($_.ScriptStackTrace)"
    )
    $errMsg  = $errLines -join "`r`n"
    $errFile = "$env:TEMP\PC-Clean-ERROR.txt"
    $errMsg | Set-Content -Path $errFile -Encoding UTF8
    try { $errMsg | Set-Content -Path "$env:USERPROFILE\Desktop\PC-Clean-ERROR.txt" -Encoding UTF8 } catch {}
    try { Stop-Transcript | Out-Null } catch {}

    Write-Host ""
    Write-Host "  [FATAL ERROR] $_" -ForegroundColor Red
    Write-Host "  Error log saved to: $errFile" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "  Press Enter to exit"
    exit 1
}
