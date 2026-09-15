<#
  Small automatic PC cleanup app - all in one file.
  Runs basic, non-destructive cleanup tasks (temp files, recycle bin, caches).
  Tries to elevate to Administrator for deeper cleanup; falls back to
  user-level-only cleanup if elevation isn't available.
  On first run it also creates a "PC Clean" Desktop shortcut (broom icon,
  always requests admin) so future runs are a simple double-click.
#>

# ---- Self-install: create Desktop shortcut on first run ----
function Install-DesktopShortcut {
    $desktop = [Environment]::GetFolderPath('Desktop')
    $lnkPath = Join-Path $desktop 'PC Clean.lnk'
    if (Test-Path -LiteralPath $lnkPath) { return }

    $powershellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

    $broomIcon = Join-Path $env:SystemRoot 'System32\cleanmgr.exe'
    if (Test-Path -LiteralPath $broomIcon) {
        $iconLocation = "$broomIcon,0"
    } else {
        $iconLocation = Join-Path $env:SystemRoot 'System32\shell32.dll,-46'
    }

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($lnkPath)
    $shortcut.TargetPath = $powershellExe
    $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $shortcut.WorkingDirectory = Split-Path -Parent $PSCommandPath
    $shortcut.IconLocation = $iconLocation
    $shortcut.Description = 'Spustí základní čištění PC'
    $shortcut.Save()

    # Force "Run as administrator" on the shortcut so double-clicking it
    # always prompts UAC directly, without a non-elevated relaunch hop.
    $bytes = [System.IO.File]::ReadAllBytes($lnkPath)
    $bytes[0x15] = $bytes[0x15] -bor 0x20
    [System.IO.File]::WriteAllBytes($lnkPath, $bytes)

    Write-Host "Zástupce na ploše vytvořen: $lnkPath" -ForegroundColor Green
}

Install-DesktopShortcut

# ---- Elevation bootstrap ----
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    try {
        Start-Process -FilePath 'powershell.exe' `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"") `
            -Verb RunAs -ErrorAction Stop
        exit
    } catch {
        Write-Host "Zvýšení oprávnění odmítnuto nebo nedostupné - pokračuji pouze s čištěním na úrovni uživatele." -ForegroundColor Yellow
    }
}

$results = New-Object System.Collections.Generic.List[object]

function Remove-CleanItems {
    param(
        [string]$TaskName,
        [string]$Path,
        [switch]$RequiresAdmin
    )

    if ($RequiresAdmin -and -not $isAdmin) {
        $results.Add([pscustomobject]@{ Task = $TaskName; Items = 0; FreedMB = 0; Status = 'Přeskočeno (vyžaduje admin)' })
        return
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        $results.Add([pscustomobject]@{ Task = $TaskName; Items = 0; FreedMB = 0; Status = 'Nenalezeno' })
        return
    }

    $items = Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
    $itemCount = 0
    $bytesFreed = 0L

    foreach ($item in $items) {
        if ($item.PSIsContainer) { continue }
        try {
            $size = $item.Length
            Remove-Item -LiteralPath $item.FullName -Force -ErrorAction Stop
            $bytesFreed += $size
            $itemCount++
        } catch {
            # locked/in-use file - skip silently
        }
    }

    # best-effort cleanup of now-empty subfolders
    Get-ChildItem -LiteralPath $Path -Recurse -Force -Directory -ErrorAction SilentlyContinue |
        Sort-Object { $_.FullName.Length } -Descending |
        ForEach-Object {
            try {
                if (-not (Get-ChildItem -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue)) {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
                }
            } catch {}
        }

    $results.Add([pscustomobject]@{
        Task    = $TaskName
        Items   = $itemCount
        FreedMB = [math]::Round($bytesFreed / 1MB, 2)
        Status  = 'OK'
    })
}

function Remove-CleanItemsByFilter {
    param(
        [string]$TaskName,
        [string]$Path,
        [string]$Filter
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        $results.Add([pscustomobject]@{ Task = $TaskName; Items = 0; FreedMB = 0; Status = 'Nenalezeno' })
        return
    }

    $files = Get-ChildItem -LiteralPath $Path -Filter $Filter -Force -ErrorAction SilentlyContinue
    $itemCount = 0
    $bytesFreed = 0L

    foreach ($file in $files) {
        try {
            $size = $file.Length
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
            $bytesFreed += $size
            $itemCount++
        } catch {
            # locked/in-use file - skip silently
        }
    }

    $results.Add([pscustomobject]@{
        Task    = $TaskName
        Items   = $itemCount
        FreedMB = [math]::Round($bytesFreed / 1MB, 2)
        Status  = 'OK'
    })
}

Write-Host "Spouštím čištění PC $(if ($isAdmin) { '(se zvýšenými oprávněními)' } else { '(pouze uživatelská úroveň)' })..." -ForegroundColor Cyan
Write-Host ""

# ---- User-level tasks (always run) ----
Remove-CleanItems -TaskName 'Dočasné soubory uživatele' -Path $env:TEMP

try {
    Clear-RecycleBin -Force -ErrorAction Stop
    $results.Add([pscustomobject]@{ Task = 'Koš'; Items = 0; FreedMB = 0; Status = 'Vyprázdněno' })
} catch {
    $results.Add([pscustomobject]@{ Task = 'Koš'; Items = 0; FreedMB = 0; Status = 'Prázdný/Přeskočeno' })
}

Remove-CleanItemsByFilter -TaskName 'Mezipaměť miniatur' `
    -Path "$env:LOCALAPPDATA\Microsoft\Windows\Explorer" -Filter 'thumbcache_*.db'

Remove-CleanItems -TaskName 'Hlášení chyb Windows (uživatel)' -Path "$env:LOCALAPPDATA\Microsoft\Windows\WER"

try {
    ipconfig /flushdns | Out-Null
    $results.Add([pscustomobject]@{ Task = 'Mezipaměť DNS'; Items = 0; FreedMB = 0; Status = 'Vymazáno' })
} catch {
    $results.Add([pscustomobject]@{ Task = 'Mezipaměť DNS'; Items = 0; FreedMB = 0; Status = 'Selhalo' })
}

# ---- Admin-only tasks ----
Remove-CleanItems -TaskName 'Systémové dočasné soubory' -Path "$env:WINDIR\Temp" -RequiresAdmin

if ($isAdmin) {
    $wuService = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
    try {
        if ($wuService) { Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue }
        Remove-CleanItems -TaskName 'Mezipaměť Windows Update' -Path "$env:WINDIR\SoftwareDistribution\Download" -RequiresAdmin
    } finally {
        if ($wuService) { Start-Service -Name wuauserv -ErrorAction SilentlyContinue }
    }
} else {
    $results.Add([pscustomobject]@{ Task = 'Mezipaměť Windows Update'; Items = 0; FreedMB = 0; Status = 'Přeskočeno (vyžaduje admin)' })
}

Remove-CleanItems -TaskName 'Prefetch' -Path "$env:WINDIR\Prefetch" -RequiresAdmin
Remove-CleanItems -TaskName 'Mezipaměť optimalizace doručování' -Path "$env:WINDIR\SoftwareDistribution\DeliveryOptimization" -RequiresAdmin

# ---- Summary ----
Write-Host ""
Write-Host "Souhrn čištění" -ForegroundColor Cyan
Write-Host "--------------"
$results | Format-Table @{Label='Úkol'; Expression={$_.Task}}, @{Label='Položky'; Expression={$_.Items}}, @{Label='Uvolněno (MB)'; Expression={$_.FreedMB}}, @{Label='Stav'; Expression={$_.Status}} -AutoSize

$totalMB = [math]::Round(($results | Measure-Object -Property FreedMB -Sum).Sum, 2)
Write-Host ""
Write-Host "Celkem uvolněno místa: $totalMB MB" -ForegroundColor Green
Write-Host "Režim spuštění: $(if ($isAdmin) { 'Správce (úplné čištění)' } else { 'Pouze uživatelská úroveň (některé úkoly přeskočeny)' })"

Write-Host ""
Read-Host "Stiskněte Enter pro zavření"
