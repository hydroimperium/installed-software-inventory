#Requires -Version 5.1
<#
.SYNOPSIS
    Инвентаризация установленного ПО Windows: реестр, Store, Steam/Epic/GOG, Scoop, WSL, portable.

.DESCRIPTION
    Собирает установленные программы из независимых источников:
      * реестр Uninstall (HKLM/HKCU, 32/64 бита) + App Paths;
      * Microsoft Store (Get-AppxPackage, с человеческими именами из манифеста);
      * Steam (все библиотеки), Epic Games, GOG;
      * Scoop, WSL, Android SDK, Edge WebView2;
      * полный обход дисков: все .exe, сгруппированные по папкам.

    Каждая папка проверяется: существует ли, сколько весит, сколько файлов.
    Пустые и почти пустые папки, апдейтеры, кэши и дампы установщиков
    отбрасываются и попадают в reports/excluded.csv с причиной.

.PARAMETER OutputDir
    Куда складывать отчёты. По умолчанию .\reports рядом со скриптом.

.PARAMETER MinGameSizeMB
    Минимальный размер папки игры в МБ. По умолчанию 100.

.PARAMETER MinPortableSizeMB
    Минимальный размер папки portable-программы в МБ. По умолчанию 3.

.PARAMETER IncludeSystem
    Не скрывать служебное ПО (VC++, .NET, SDK, драйверы) в HTML-отчёте.

.PARAMETER NoStore
    Не собирать приложения Microsoft Store.

.PARAMETER NoGames
    Не собирать игры.

.PARAMETER NoPortable
    Не обходить диски (только реестр, Store, Steam/Epic/GOG, Scoop, WSL).

.PARAMETER ExtraRoots
    Дополнительные папки для поиска portable-программ.

.PARAMETER OverridesPath
    Файл ручных правок (по умолчанию .\overrides.psd1, если есть).

.PARAMETER Open
    Открыть HTML-отчёт после сборки.

.PARAMETER Audit
    Показать папки с exe, которые не попали в отчёт.

.EXAMPLE
    .\Get-InstalledSoftware.ps1

.EXAMPLE
    .\Get-InstalledSoftware.ps1 -MinGameSizeMB 50 -Open

.EXAMPLE
    .\Get-InstalledSoftware.ps1 -NoPortable -Audit
#>
[CmdletBinding()]
param(
    [string]$OutputDir,
    [int]$MinGameSizeMB = 100,
    [int]$MinPortableSizeMB = 3,
    [switch]$IncludeSystem,
    [switch]$NoStore,
    [switch]$NoGames,
    [switch]$NoPortable,
    [string[]]$ExtraRoots = @(),
    [string]$OverridesPath,
    [switch]$SkipShortcuts,
    [switch]$Audit,
    [switch]$Open,
    [switch]$Quiet
)

$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'

if (-not $OutputDir) { $OutputDir = Join-Path $PSScriptRoot 'reports' }
$null = New-Item -ItemType Directory -Force -Path $OutputDir

function Write-Step([string]$text) {
    if (-not $Quiet) { Write-Host "  $text" -ForegroundColor DarkGray }
}

# ---------------------------------------------------------------- helpers
function Get-Norm([string]$s) {
    if (-not $s) { return '' }
    return ($s -replace '[^\p{L}\p{N}]', '').ToLowerInvariant()
}

function Get-CleanName([string]$name) {
    if (-not $name) { return '' }
    $n = $name.Trim()
    $n = $n -replace '\s*[\[\(](RePack|Repack|GOG|Portable|Rip)[\]\)]\s*$', ''
    $n = $n -replace "[\u2018\u2019]", "'"
    $n = $n -replace '^(.+?)[,\s]+версия\s+[\d\.\+\w\-]+$', '$1'
    $n = $n -replace '^(.+?)\s+version\s+[\d\.\+\w\-]+$', '$1'
    $n = $n -replace '\s+\((User|x64|x86|64-bit|32-bit)\)\s*$', ''
    $n = $n -replace '\s+for\s+\S+$', ''
    $n = $n -replace '^Sid Meiers ', "Sid Meier's "
    return $n.Trim()
}

$script:StatsCache = @{}
function Get-FolderStats([string]$path) {
    if (-not $path) { return $null }
    if ($script:StatsCache.ContainsKey($path)) { return $script:StatsCache[$path] }
    $r = [pscustomobject]@{ SizeMB = 0; Files = 0 }
    if ($path -match '^[A-Za-z]:\\' -and (Test-Path -LiteralPath $path -PathType Leaf)) {
        $f = Get-Item -LiteralPath $path
        $r.SizeMB = [math]::Round($f.Length / 1MB, 1); $r.Files = 1
        $script:StatsCache[$path] = $r
        return $r
    }
    if ($path -match '^[A-Za-z]:\\' -and (Test-Path -LiteralPath $path -PathType Container)) {
        $sum = 0L; $cnt = 0
        foreach ($f in (Get-ChildItem -LiteralPath $path -Recurse -File -Force -ErrorAction SilentlyContinue)) {
            $sum += $f.Length; $cnt++
        }
        $r.SizeMB = [math]::Round($sum / 1MB, 1); $r.Files = $cnt
    }
    $script:StatsCache[$path] = $r
    return $r
}

function Add-Skipped([string]$name, [string]$path, [double]$mb, [int]$files, [string]$reason) {
    [void]$script:Skipped.Add([pscustomobject]@{
            Name = $name; Path = $path; SizeMB = $mb; Files = $files; Reason = $reason
        })
}

function Test-AppFolder([string]$path, [double]$minMB, [int]$minFiles, [string]$name, [string]$why) {
    $st = Get-FolderStats $path
    if (-not $st) { return $false }
    if ($st.SizeMB -lt $minMB) { Add-Skipped $name $path $st.SizeMB $st.Files "${why}: меньше $minMB МБ"; return $false }
    if ($st.Files -lt $minFiles) { Add-Skipped $name $path $st.SizeMB $st.Files "${why}: меньше $minFiles файлов"; return $false }
    return $true
}

$SystemRx = '(?i)Visual C\+\+|\.NET (SDK|Runtime|Desktop Runtime)|Windows Desktop Runtime|Windows SDK|DirectX|Edge WebView2|Edge Update|Epic Online Services|Mozilla Maintenance|AMD[ _](Chipset|Install)|Realtek Card Reader|EpsonNet|Драйвер принтера|Application Compatibility|Application Virtualization|Windows Subsystem for Linux|^WSL|WinAppRuntime|ShellExtension|ContextMenu|SparseApp|Winget\.Source|DesktopAppInstaller|StorePurchaseApp|LanguageExperiencePack|VCLibs|Microsoft\.UI\.Xaml|HostRuntime|AppRuntime|Windows App Runtime|Windows Media Player|Windows Photo Viewer|Windows Defender|Microsoft Office$|Razer S3|Razer.*Wizard'
$ToolRx   = '(?i)Visual Studio Code|Android Studio|Android SDK|IntelliJ|PyCharm|WebStorm|^Git$|GitHub CLI|Node\.js|Python|OpenJDK|Temurin|JDK|Docker|Postman|Scoop|Maestro|Cursor|Devin|Antigravity|OpenCode|DeepSeek|Codex|Blender|Godot|NekoBox|LockHunter|WinRAR|7-Zip|PowerToys|zapret|Ubuntu|Bun$|Chocolatey|spicetify'
$GamePathRx = '(?i)steamapps\\common|\\Games\\|\\torrent\\|GOG Games'

function Get-Category([string]$name, [string]$path, [string]$source, [string]$kind) {
    if ($kind -eq 'Игра') { return 'Игра' }
    if ($source -in @('Steam', 'Epic', 'GOG')) { return 'Игра' }
    if ($path -and $path -match $GamePathRx) { return 'Игра' }
    if ($source -eq 'WSL') { return 'Инструмент' }
    if ($source -eq 'Microsoft Store') {
        if ($name -match $SystemRx) { return 'Системное' }
        return 'Store'
    }
    if ($name -match $SystemRx) { return 'Системное' }
    if ($name -match $ToolRx) { return 'Инструмент' }
    return 'Приложение'
}

function New-App {
    param(
        [string]$Name, [string]$Version, [string]$Publisher, [string]$Path,
        [string]$Source, [string]$Scope = '', [string]$Uninstall = '', [string]$Kind = 'Установлено'
    )
    $clean = Get-CleanName $Name
    if (-not $clean) { $clean = $Name }
    [pscustomobject]@{
        Name      = $clean
        Version   = "$Version".Trim()
        Publisher = "$Publisher".Trim()
        Path      = if ($Path) { $pp = ($Path -replace '/', '\').TrimEnd('\'); if ($pp -match '^[a-z]:') { $pp = $pp.Substring(0, 1).ToUpper() + $pp.Substring(1) }; $pp } else { '' }
        Category  = (Get-Category $clean $Path $Source $Kind)
        Source    = $Source
        Scope     = $Scope
        Kind      = $Kind
        Uninstall = $Uninstall
        SizeMB    = $null
        Files     = $null
        Exists    = ''
        LastRun   = $null
        Extra     = ''
    }
}

# ---------------------------------------------------------------- 1. registry
function Get-RegistryApps {
    Write-Step 'Реестр (HKLM/HKCU, 32/64 бита)...'
    $keys = @(
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'; Scope = 'Система' },
        @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'; Scope = 'Система' },
        @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'; Scope = 'Пользователь' },
        @{ Path = 'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'; Scope = 'Пользователь' }
    )

    $appPaths = @{}
    foreach ($hive in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\*',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\*')) {
        Get-ItemProperty $hive -ErrorAction SilentlyContinue | ForEach-Object {
            $exe = "$($_.PSChildName)" -replace '\.exe$', ''
            $p = $_.'(default)'
            if ($exe -and $p -and $p -match '^[A-Za-z]:\\') { $appPaths[(Get-Norm $exe)] = $p }
        }
    }

    $out = New-Object System.Collections.ArrayList
    foreach ($k in $keys) {
        Get-ItemProperty $k.Path | ForEach-Object {
            if (-not $_.DisplayName) { return }
            if ($_.SystemComponent -eq 1) { return }
            if ($_.ParentKeyName) { return }

            $loc = $_.InstallLocation
            if ($loc) { $loc = $loc.Trim('"').TrimEnd('\') }

            if (-not $loc -and $_.DisplayIcon) {
                $ico = ($_.DisplayIcon -split ',')[0].Trim('"').Trim()
                if ($ico -match '^[A-Za-z]:\\' -and (Test-Path -LiteralPath $ico)) { $loc = Split-Path -Parent $ico }
            }
            if (-not $loc -and $_.UninstallString) {
                $us = $_.UninstallString.Trim()
                if ($us -match '^"([^"]+)"') { $us = $matches[1] }
                elseif ($us -match '^(\S+\.exe)') { $us = $matches[1] }
                if ($us -match '^[A-Za-z]:\\' -and (Test-Path -LiteralPath $us)) { $loc = Split-Path -Parent $us }
            }
            if (-not $loc) {
                $nn = Get-Norm $_.DisplayName
                foreach ($key in $appPaths.Keys) {
                    if ($key.Length -ge 3 -and ($nn -like "*$key*" -or $key -like "*$nn*")) {
                        $loc = Split-Path -Parent $appPaths[$key]
                        break
                    }
                }
            }
            if ($loc -match '(?i)Package Cache|\\Installer\\|\\Windows\\Installer') { $loc = '' }
            if ($loc -match '(?i)\\bin$' -and (Split-Path -Parent $loc) -and (Test-Path (Split-Path -Parent $loc))) {
                $loc = Split-Path -Parent $loc
            }

            [void]$out.Add((New-App -Name $_.DisplayName -Version "$($_.DisplayVersion)" -Publisher "$($_.Publisher)" `
                        -Path $loc -Source 'Реестр' -Scope $k.Scope -Uninstall "$($_.UninstallString)"))
        }
    }
    return $out
}

# ---------------------------------------------------------------- 2. store
function Get-StoreApps {
    Write-Step 'Приложения Microsoft Store...'
    $out = New-Object System.Collections.ArrayList
    Get-AppxPackage | Where-Object {
        -not $_.IsFramework -and -not $_.IsResourcePackage -and $_.SignatureKind -ne 'System'
    } | ForEach-Object {
        $display = $null
        try {
            $manifest = Get-AppxPackageManifest $_
            $display = $manifest.Package.Properties.DisplayName
        } catch {}
        if (-not $display -or $display -match '^ms-resource:') { $display = $_.Name }
        [void]$out.Add((New-App -Name $display -Version "$($_.Version)" -Publisher "$($_.PublisherDisplayName)" `
                    -Path $_.InstallLocation -Source 'Microsoft Store' `
                    -Scope $(if ($_.InstallLocation -like "$env:LOCALAPPDATA*") { 'Пользователь' } else { 'Система' }) `
                    -Uninstall "Get-AppxPackage -Name '$($_.Name)' | Remove-AppxPackage" -Kind 'Store'))
        $out[$out.Count - 1].Extra = $_.Name
    }
    return $out
}

# ---------------------------------------------------------------- 3. steam
function Get-SteamApps {
    Write-Step 'Steam (все библиотеки)...'
    $out = New-Object System.Collections.ArrayList
    $roots = @()
    foreach ($base in @('C:\Program Files (x86)\Steam', 'C:\Program Files\Steam', 'D:\steam', 'D:\Steam', 'E:\steam')) {
        if (Test-Path (Join-Path $base 'steamapps')) { $roots += $base }
    }
    $steamReg = Get-ItemProperty 'HKCU:\Software\Valve\Steam' -ErrorAction SilentlyContinue
    if ($steamReg.SteamPath -and (Test-Path (Join-Path $steamReg.SteamPath 'steamapps'))) { $roots += $steamReg.SteamPath }

    $roots = $roots | ForEach-Object { $_.TrimEnd('\') } | Sort-Object -Unique
    if (-not $roots) { return $out }

    $libs = @()
    foreach ($r in $roots) {
        $libs += $r
        $vdf = Join-Path $r 'steamapps\libraryfolders.vdf'
        if (Test-Path $vdf) {
            Get-Content $vdf | ForEach-Object { if ($_ -match '"path"\s+"(.+?)"') { $libs += ($matches[1] -replace '\\\\', '\') } }
        }
    }
    $libs = $libs | ForEach-Object { $_.TrimEnd('\') } | Sort-Object -Unique
    $manifestDirs = @()
    $manifestNames = @()

    foreach ($lib in $libs) {
        Get-ChildItem (Join-Path $lib 'steamapps') -Filter 'appmanifest_*.acf' -ErrorAction SilentlyContinue | ForEach-Object {
            $txt = Get-Content $_.FullName -Raw
            $appid = if ($txt -match '"appid"\s+"(\d+)"') { $matches[1] } else { '' }
            $nm    = if ($txt -match '"name"\s+"(.+?)"') { $matches[1] } else { '' }
            $inst  = if ($txt -match '"installdir"\s+"(.+?)"') { $matches[1] } else { '' }
            if (-not $nm -or $nm -match 'Steamworks Common Redistributables') { return }
            $p = Join-Path (Join-Path $lib 'steamapps\common') $inst
            $manifestDirs += $p.ToLower()
            $manifestNames += (Get-Norm $nm)
            $st = Get-FolderStats $p
            if (-not (Test-AppFolder $p $MinGameSizeMB 10 $nm 'игра')) { return }
            $a = New-App -Name $nm -Version "AppID $appid" -Publisher 'Steam' -Path $p -Source 'Steam' -Scope 'Система' -Uninstall "steam://uninstall/$appid" -Kind 'Игра'
            $a.SizeMB = $st.SizeMB; $a.Files = $st.Files
            [void]$out.Add($a)
        }
    }
    foreach ($lib in $libs) {
        $common = Join-Path $lib 'steamapps\common'
        Get-ChildItem -LiteralPath $common -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            if ($manifestDirs -contains $_.FullName.ToLower()) { return }
            if ($_.Name -eq 'Steamworks Shared') { return }
            if ($manifestNames -contains (Get-Norm $_.Name)) { return }
            $st = Get-FolderStats $_.FullName
            if (-not (Test-AppFolder $_.FullName $MinGameSizeMB 10 $_.Name 'игра без манифеста')) { return }
            $a = New-App -Name $_.Name -Version '' -Publisher 'Steam' -Path $_.FullName -Source 'Steam' -Scope 'Система' -Kind 'Игра'
            $a.SizeMB = $st.SizeMB; $a.Files = $st.Files
            [void]$out.Add($a)
        }
    }
    return $out
}

# ---------------------------------------------------------------- 4. epic / gog
function Get-EpicApps {
    Write-Step 'Epic Games...'
    $out = New-Object System.Collections.ArrayList
    foreach ($d in @('C:\ProgramData\Epic\EpicGamesLauncher\Data\Manifests', "$env:ProgramData\Epic\EpicGamesLauncher\Data\Manifests") | Sort-Object -Unique) {
        Get-ChildItem $d -Filter '*.item' -ErrorAction SilentlyContinue | ForEach-Object {
            try { $j = Get-Content $_.FullName -Raw | ConvertFrom-Json } catch { return }
            if (-not $j.DisplayName -or -not $j.InstallLocation) { return }
            $st = Get-FolderStats $j.InstallLocation
            if (-not (Test-AppFolder $j.InstallLocation $MinGameSizeMB 10 $j.DisplayName 'игра Epic')) { return }
            $a = New-App -Name $j.DisplayName -Version "$($j.AppVersion)" -Publisher 'Epic Games' -Path $j.InstallLocation -Source 'Epic' -Kind 'Игра'
            $a.SizeMB = $st.SizeMB; $a.Files = $st.Files
            [void]$out.Add($a)
        }
    }
    return $out
}

function Get-GogApps {
    Write-Step 'GOG...'
    $out = New-Object System.Collections.ArrayList
    foreach ($hive in @('HKLM:\SOFTWARE\WOW6432Node\GOG.com\Games\*', 'HKLM:\SOFTWARE\GOG.com\Games\*')) {
        Get-ItemProperty $hive -ErrorAction SilentlyContinue | ForEach-Object {
            if (-not $_.path) { return }
            $st = Get-FolderStats $_.path
            if (-not (Test-AppFolder $_.path $MinGameSizeMB 10 $_.gameName 'игра GOG')) { return }
            $a = New-App -Name $_.gameName -Version "$($_.ver)" -Publisher 'GOG' -Path $_.path -Source 'GOG' -Kind 'Игра'
            $a.SizeMB = $st.SizeMB; $a.Files = $st.Files
            [void]$out.Add($a)
        }
    }
    return $out
}

# ---------------------------------------------------------------- 5. scoop / wsl
function Get-ScoopApps {
    Write-Step 'Scoop...'
    $out = New-Object System.Collections.ArrayList
    $root = Join-Path $env:USERPROFILE 'scoop\apps'
    if (-not (Test-Path $root)) { return $out }
    Get-ChildItem $root -Directory | ForEach-Object {
        if ($_.Name -eq 'scoop') { return }
        $cur = Join-Path $_.FullName 'current'
        $p = if (Test-Path $cur) { $cur } else { $_.FullName }
        $v = ''
        $verDir = Get-ChildItem $_.FullName -Directory | Where-Object Name -ne 'current' | Sort-Object Name -Descending | Select-Object -First 1
        if ($verDir) { $v = $verDir.Name }
        $appName = $_.Name
        if ($appName -cmatch '^[a-z0-9\.\-]+$') { $appName = (Get-Culture).TextInfo.ToTitleCase($appName) }
        $a = New-App -Name $appName -Version $v -Publisher 'Scoop' -Path $p -Source 'Scoop' -Scope 'Пользователь' -Kind 'Портативно'
        $st = Get-FolderStats $p
        $a.SizeMB = $st.SizeMB; $a.Files = $st.Files
        [void]$out.Add($a)
    }
    return $out
}

function Get-WslApps {
    Write-Step 'WSL...'
    $out = New-Object System.Collections.ArrayList
    $raw = & wsl.exe -l -q 2>$null
    if (-not $raw) { return $out }
    foreach ($line in $raw) {
        $n = ("$line" -replace "`0", '').Trim()
        if (-not $n) { continue }
        if ($n -match '(?i)^docker-desktop') { continue }
        $a = New-App -Name $n -Version '' -Publisher 'WSL' -Path "\\wsl$\$n" -Source 'WSL' -Scope 'Пользователь' -Kind 'Портативно'
        $a.Exists = 'да'
        [void]$out.Add($a)
    }
    return $out
}

# ---------------------------------------------------------------- 6. toolchains
function Get-ToolchainApps {
    Write-Step 'SDK и тулчейны...'
    $out = New-Object System.Collections.ArrayList

    foreach ($p in (@($env:ANDROID_HOME, $env:ANDROID_SDK_ROOT, "$env:LOCALAPPDATA\Android\Sdk") | Where-Object { $_ })) {
        if (-not (Test-Path $p)) { continue }
        $a = New-App -Name 'Android SDK' -Version '' -Publisher 'Google' -Path $p -Source 'Toolchain' -Scope 'Пользователь' -Kind 'Портативно'
        $st = Get-FolderStats $p; $a.SizeMB = $st.SizeMB; $a.Files = $st.Files
        [void]$out.Add($a)
    }

    $scoopRoot = Join-Path $env:USERPROFILE 'scoop'
    if (Test-Path (Join-Path $scoopRoot 'shims\scoop.cmd')) {
        $a = New-App -Name 'Scoop' -Version '' -Publisher 'Scoop' -Path $scoopRoot -Source 'Toolchain' -Scope 'Пользователь' -Kind 'Портативно'
        $st = Get-FolderStats $scoopRoot; $a.SizeMB = $st.SizeMB; $a.Files = $st.Files
        [void]$out.Add($a)
    }

    $wvRoot = 'C:\Program Files (x86)\Microsoft\EdgeWebView\Application'
    if (Test-Path $wvRoot) {
        $ver = Get-ChildItem $wvRoot -Directory | Where-Object { $_.Name -match '^\d+\.' } | Sort-Object Name -Descending | Select-Object -First 1
        $a = New-App -Name 'Microsoft Edge WebView2 Runtime' -Version $(if ($ver) { $ver.Name } else { '' }) -Publisher 'Microsoft' -Path $wvRoot -Source 'Toolchain' -Kind 'Установлено'
        $st = Get-FolderStats $wvRoot; $a.SizeMB = $st.SizeMB; $a.Files = $st.Files
        [void]$out.Add($a)
    }
    return $out
}

# ---------------------------------------------------------------- 7. disk index
$script:DiskIndex = $null

function Get-DiskRoots {
    $roots = @([System.IO.DriveInfo]::GetDrives() |
        Where-Object { $_.DriveType -eq 'Fixed' -and $_.IsReady } |
        ForEach-Object { $_.RootDirectory.FullName })
    $roots += $ExtraRoots
    return $roots | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Sort-Object -Unique
}

$WalkSkipRx = '(?i)\\(Windows|WinSxS|\$Recycle\.Bin|System Volume Information|WindowsApps|Installer|Package Cache|Temp|temp|node_modules|\.git|\.cache|__pycache__|site-packages|\.gradle|\.docker|\.nuget|\.m2|\.cargo|\.rustup|\.android|\.kotlin|\.vscode|\.idea|globalStorage|JetBrains|ClickToRun|_components|OneDrive|\.tlauncher|build-tools|platform-tools|chocolatey\\lib)(\\|$)'

function Get-DiskIndex {
    if ($script:DiskIndex) { return $script:DiskIndex }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $exes = New-Object System.Collections.Generic.List[object]
    $sizes = New-Object 'System.Collections.Generic.Dictionary[string,long]'
    $counts = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    $dlls = New-Object 'System.Collections.Generic.Dictionary[string,int]'

    $stack = New-Object System.Collections.Generic.Stack[string]
    foreach ($r in (Get-DiskRoots)) { $stack.Push($r) }

    $processed = 0
    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()
        if ($dir -match $WalkSkipRx) { continue }
        $di = $null
        try { $di = New-Object System.IO.DirectoryInfo $dir } catch { continue }
        if ($di.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }
        $processed++
        if (-not $Quiet -and ($processed % 2000 -eq 0)) {
            Write-Progress -Activity 'Обход дисков' -Status ("{0:N0} папок" -f $processed) -CurrentOperation $dir
        }
        $sum = 0L; $cnt = 0; $dll = 0
        try {
            foreach ($f in $di.EnumerateFiles()) {
                $sum += $f.Length; $cnt++
                $ext = $f.Extension
                if ($ext -ieq '.exe') { $exes.Add([pscustomobject]@{ Path = $f.FullName; Size = $f.Length; Dir = $dir }) }
                elseif ($ext -ieq '.dll') { $dll++ }
            }
        } catch {}
        $sizes[$dir] = $sum; $counts[$dir] = $cnt; $dlls[$dir] = $dll
        try { foreach ($d in $di.EnumerateDirectories()) { $stack.Push($d.FullName) } } catch {}
    }

    # свернуть размеры вверх по дереву
    $dirs = @($sizes.Keys) | Sort-Object { $_.Length } -Descending
    foreach ($d in $dirs) {
        $parent = Split-Path -Parent $d
        if (-not $parent) { continue }
        if (-not $sizes.ContainsKey($parent)) { $sizes[$parent] = 0L; $counts[$parent] = 0; $dlls[$parent] = 0 }
        $sizes[$parent] += $sizes[$d]
        $counts[$parent] += $counts[$d]
        $dlls[$parent] += $dlls[$d]
    }

    $sw.Stop()
    if (-not $Quiet) { Write-Progress -Activity 'Обход дисков' -Completed }
    Write-Step ("  просканировано: {0} exe, {1} папок, {2:N0} с" -f $exes.Count, $sizes.Count, $sw.Elapsed.TotalSeconds)

    $script:DiskIndex = @{
        Exes  = $exes
        Sizes = $sizes
        Counts = $counts
        Dlls  = $dlls
    }
    return $script:DiskIndex
}

# ---------------------------------------------------------------- 8. disk apps
$JunkDirRx = '(?i)\\(updater|updaters|update|updates|pending|download|downloads|installer|installers|setup|crashpad|crash|crashreporter|logs?|cache|temp|tmp|backup|backups|old|previous|obsolete|node_modules|__pycache__|site-packages|\.venv|venv|dist|build|obj|target|out|bin|bin64|lib|libs|lib64|include|resources|locales|plugins|extensions|drivers|runtime|jre|jdk|jbr|redist|redistributable|edgecore|edgeupdate|cnext|qemu|windows defender|windows media player|windows photo viewer|wsl|common files|epicgameslauncher|microsoft office[^\\]*|battlenet|_components|agunlocker|ai\.opencode\.desktop)(\\|$)'
$InstallerExeRx = '(?i)(setup|install|installer|unins|uninst|update|updater|redist|vc_redist|dxsetup|directx|dotnet|runtime)'

$ContainerRx = '(?i)^([A-Z]:\\|C:\\Program Files|C:\\Program Files \(x86\)|C:\\ProgramData|C:\\Users|C:\\Users\\[^\\]+|C:\\Users\\[^\\]+\\AppData|C:\\Users\\[^\\]+\\AppData\\(Local|Roaming|LocalLow)|C:\\Users\\[^\\]+\\Desktop|C:\\Users\\[^\\]+\\Downloads|C:\\Users\\[^\\]+\\Documents|C:\\Users\\Public|C:\\Users\\All Users)$'

function Get-DiskApps($known) {
    Write-Step 'Обход дисков...'
    $idx = Get-DiskIndex
    $out = New-Object System.Collections.ArrayList

    $knownDirs = @()
    $knownNames = @{}
    foreach ($k in $known) {
        if (-not $k.Path) { continue }
        $p = $k.Path.ToLower()
        if ($p -match '\\[^\\]+\.[a-z0-9]+$') { $p = (Split-Path -Parent $k.Path).ToLower() }
        $knownDirs += $p
        $knownNames[(Get-Norm $k.Name)] = 1
    }

    $exeByDir = @{}
    foreach ($e in $idx.Exes) {
        if (-not $exeByDir.ContainsKey($e.Dir)) { $exeByDir[$e.Dir] = New-Object System.Collections.ArrayList }
        [void]$exeByDir[$e.Dir].Add($e)
    }

    # папка -> главный exe (с подъёмом из bin/bin64)
    $candDirs = @{}
    foreach ($g in ($idx.Exes | Group-Object Dir)) {
        $dir = $g.Name
        $leaf = Split-Path -Leaf $dir
        if ($leaf -match '(?i)^(bin|bin64|x64)$') {
            $parent = Split-Path -Parent $dir
            if ($parent) { $dir = $parent }
        }
        if (-not $candDirs.ContainsKey($dir)) {
            $best = $g.Group | Sort-Object Size -Descending | Select-Object -First 1
            $candDirs[$dir] = $best.Path
        }
    }

    $cands = @()
    foreach ($dir in $candDirs.Keys) {
        $exePath = $candDirs[$dir]
        $vi = $null
        try { $vi = (Get-Item -LiteralPath $exePath).VersionInfo } catch {}

        $product = "$($vi.ProductName)".Trim()
        $company = "$($vi.CompanyName)".Trim()
        $descr   = "$($vi.FileDescription)".Trim()
        $version = "$($vi.ProductVersion)".Trim()
        $leaf    = Split-Path -Leaf $dir

        $name = $product
        if (-not $name -or $name -match '(?i)^(Microsoft® Windows® Operating System|Microsoft Windows|Windows)$' -or $name -match '^\d') {
            $name = $leaf -replace '^v?\d+[\d\.\-_]*\s*', ''
            if (-not $name) { $name = $leaf }
        }
        $name = Get-CleanName $name

        $exesHere = @()
        if ($exeByDir.ContainsKey($dir)) { $exesHere = @($exeByDir[$dir]) }
        if (-not $exesHere) { $exesHere = @($idx.Exes | Where-Object { $_.Path -eq $exePath }) }

        $cands += [pscustomobject]@{
            Dir      = $dir
            Name     = $name
            Version  = $version
            SizeMB   = [math]::Round($idx.Sizes[$dir] / 1MB, 1)
            Files    = $idx.Counts[$dir]
            Dlls     = $idx.Dlls[$dir]
            ExeCount = $exesHere.Count
            Exe      = $exePath
            Product  = $product
            Company  = $company
            Descr    = $descr
            Leaf     = $leaf
        }
    }

    # сначала отсеиваем контейнеры, мусор и уже известное
    $valid = New-Object System.Collections.ArrayList
    foreach ($c in $cands) {
        $d = $c.Dir
        $dl = $d.ToLower()
        $leaf = $c.Leaf

        if ($d -match $ContainerRx) {
            Add-Skipped $c.Name $d $c.SizeMB $c.Files 'контейнер (рабочий стол, корень диска, Program Files)'
            continue
        }
        if ($d -match $JunkDirRx) {
            Add-Skipped $c.Name $d $c.SizeMB $c.Files 'служебная папка (апдейтер, кэш, компонент)'
            continue
        }
        if ($leaf -match '(?i)^(bin|bin64|x64|win64|lib|libs|resources|locales|plugins|extensions|drivers|runtime|jre|jdk|jbr|redist|setup|installer|uninstall|update|updater|pending|crashpad|crash|helper|service|agent|cli-plugins|platform-tools|emulator|build-tools)$') {
            Add-Skipped $c.Name $d $c.SizeMB $c.Files 'служебная папка (компонент)'
            continue
        }
        if ($c.Name.Trim() -match '^[\d\.]+$') {
            Add-Skipped $c.Name $d $c.SizeMB $c.Files 'имя — только версия'
            continue
        }
        if ($c.Name -match '(?i)(Admin Agent|Update Agent)$' -or $c.Name -match '(?i)Updater\s*$' -or $c.Name -match '^(Qt6|vulkan|windows-x86_64|platform-tools|emulator|cli-plugins)$') {
            Add-Skipped $c.Name $d $c.SizeMB $c.Files 'компонент, а не программа'
            continue
        }
        if ($c.Name -match '(?i)Операционная система|Microsoft \(R\) Windows') {
            $c.Name = $leaf
        }

        $covered = $false
        foreach ($kd in $knownDirs) {
            if ($dl -eq $kd -or $dl.StartsWith($kd + '\') -or $kd.StartsWith($dl + '\')) { $covered = $true; break }
        }
        if ($covered) {
            Add-Skipped $c.Name $d $c.SizeMB $c.Files 'уже есть в списке (реестр/Store/Steam)'
            continue
        }
        if ($knownNames.ContainsKey((Get-Norm $c.Name))) {
            Add-Skipped $c.Name $d $c.SizeMB $c.Files 'уже есть в списке по имени'
            continue
        }

        if ($c.Dlls -eq 0 -and ($c.Descr -match '(?i)setup|install' -or (Split-Path -Leaf $c.Exe) -match '(?i)setup|install')) {
            Add-Skipped $c.Name $d $c.SizeMB $c.Files 'дамп установщиков'
            continue
        }

        [void]$valid.Add($c)
    }

    # теперь оставляем только верхние папки, вложенные отбрасываем
    $ordered = $valid | Sort-Object { $_.Dir.Length }
    $kept = New-Object System.Collections.ArrayList
    foreach ($c in $ordered) {
        $cl = $c.Dir.ToLower()
        $nested = $false
        foreach ($k in $kept) {
            $kl = $k.Dir.ToLower()
            if ($cl.StartsWith($kl + '\')) { $nested = $true; break }
        }
        if ($nested) {
            Add-Skipped $c.Name $c.Dir $c.SizeMB $c.Files 'вложено в другую программу'
            continue
        }
        [void]$kept.Add($c)
    }

    foreach ($c in $kept) {
        $d = $c.Dir
        $leaf = $c.Leaf

        $isGame = ($d -match $GamePathRx) -or
                  (Test-Path (Join-Path $d 'UnityPlayer.dll')) -or
                  (Test-Path (Join-Path $d 'Engine')) -or
                  ($c.Exe -match '(?i)-Win64-Shipping\.exe$') -or
                  ($c.Product -match '(?i)BootstrapPackagedGame')

        if ($isGame) {
            if ($c.SizeMB -lt $MinGameSizeMB) { Add-Skipped $c.Name $d $c.SizeMB $c.Files "игра: меньше $MinGameSizeMB МБ"; continue }
            $name = Get-CleanName $leaf
            if (-not $name) { $name = $leaf }
        }
        else {
            if ($c.SizeMB -lt $MinPortableSizeMB) { Add-Skipped $c.Name $d $c.SizeMB $c.Files "portable: меньше $MinPortableSizeMB МБ"; continue }
            $name = $c.Name
        }

        $kind = if ($isGame) { 'Игра' } else { 'Портативно' }
        $a = New-App -Name $name -Version $c.Version -Publisher $c.Company -Path $d -Source 'Диск' -Scope 'Пользователь' -Kind $kind
        $a.SizeMB = $c.SizeMB
        $a.Files = $c.Files
        [void]$out.Add($a)
    }
    return $out
}

# ---------------------------------------------------------------- 9. shortcuts
function Get-ShortcutApps($known) {
    Write-Step 'Ярлыки «Пуск» и рабочий стол...'
    $out = New-Object System.Collections.ArrayList
    try { $shell = New-Object -ComObject WScript.Shell } catch { return $out }

    $knownDirs = @()
    foreach ($k in $known) {
        if (-not $k.Path) { continue }
        if ($k.Path -match '\\[^\\]+\.[a-zA-Z0-9]+$') { $knownDirs += (Split-Path -Parent $k.Path).ToLower() }
        else { $knownDirs += $k.Path.ToLower() }
    }

    $dirs = @(
        'C:\ProgramData\Microsoft\Windows\Start Menu\Programs',
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs",
        "$env:USERPROFILE\Desktop",
        'C:\Users\Public\Desktop'
    ) | Where-Object { Test-Path $_ }

    $seen = @{}
    foreach ($d in $dirs) {
        Get-ChildItem $d -Filter '*.lnk' -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $t = ''
            try { $t = $shell.CreateShortcut($_.FullName).TargetPath } catch { return }
            if (-not $t -or $t -notmatch '(?i)\.exe$') { return }
            if (-not (Test-Path -LiteralPath $t)) { return }
            if ($t -match '(?i)\\Windows\\|\\System32\\|\\SysWOW64\\|unins|uninstall|setup|update|crash|helper') { return }
            if ($t -like 'C:\Program Files\*' -or $t -like 'C:\Program Files (x86)\*' -or $t -like "$env:LOCALAPPDATA\Programs\*") { return }
            if ($seen.ContainsKey($t.ToLower())) { return }
            $seen[$t.ToLower()] = 1

            $dirT = Split-Path -Parent $t
            $covered = $false
            foreach ($kd in $knownDirs) {
                if ($dirT.ToLower() -eq $kd -or $dirT.ToLower().StartsWith($kd + '\') -or $kd.StartsWith($dirT.ToLower() + '\')) { $covered = $true; break }
            }
            if ($covered) { return }

            $st = Get-FolderStats $dirT
            if ($st.SizeMB -lt $MinPortableSizeMB) { return }
            $a = New-App -Name $_.BaseName -Version '' -Publisher '' -Path $t -Source 'Ярлык' -Scope 'Пользователь' -Kind 'Портативно'
            $a.SizeMB = $st.SizeMB; $a.Files = $st.Files
            [void]$out.Add($a)
        }
    }
    return $out
}

# ---------------------------------------------------------------- 10. merge
function Merge-Apps($apps) {
    $prio = @{ 'Steam' = 0; 'Epic' = 0; 'GOG' = 0; 'Реестр' = 1; 'Toolchain' = 1; 'Microsoft Store' = 2; 'Scoop' = 3; 'Диск' = 4; 'Ярлык' = 5 }
    $sorted = $apps | Sort-Object @{ Expression = { if ($prio.ContainsKey($_.Source)) { $prio[$_.Source] } else { 9 } } }

    $result = New-Object System.Collections.ArrayList
    $byName = @{}
    $byPath = @{}
    $dupCount = @{}

    foreach ($a in $sorted) {
        $n = Get-Norm $a.Name
        $p = "$($a.Path)".ToLower()

        if ($byName.ContainsKey("$n|$p")) { continue }

        $dup = $false
        if ($n -and $byName.ContainsKey($n)) {
            $prev = $byName[$n]
            if (-not $p -or -not $prev.Path -or $p -eq $prev.Path.ToLower() -or
                $p.StartsWith($prev.Path.ToLower() + '\') -or $prev.Path.ToLower().StartsWith($p + '\')) { $dup = $true }
        }
        if ($dup) { continue }
        if ($p -and $byPath.ContainsKey($p)) { continue }

        if ($n -and $byName.ContainsKey($n)) {
            $dupCount[$n] = 1 + $(if ($dupCount.ContainsKey($n)) { $dupCount[$n] } else { 0 })
            $a.Name = "$($a.Name) (копия $($dupCount[$n]))"
        }

        [void]$result.Add($a)
        if ($n) { $byName[$n] = $a }
        if ($p) { $byPath[$p] = $a }
    }
    return $result
}

# ---------------------------------------------------------------- 11. validate
function Apply-Overrides($apps, [string]$file) {
    if (-not $file -or -not (Test-Path $file)) { return $apps }
    Write-Step "Правки из $([IO.Path]::GetFileName($file))..."
    $ov = Import-PowerShellDataFile $file

    $map = @{}
    if ($ov.Names) { foreach ($k in $ov.Names.Keys) { $map[(Get-Norm $k)] = $ov.Names[$k] } }

    foreach ($a in $apps) {
        $key = Get-Norm $a.Name
        if (-not $map.ContainsKey($key)) { continue }
        $fix = $map[$key]
        if ($fix.Path) { $a.Path = $fix.Path.TrimEnd('\'); $a.SizeMB = $null; $a.Files = $null }
        if ($fix.Name) { $a.Name = $fix.Name }
        if ($fix.Category) { $a.Category = $fix.Category }
        if ($fix.Version) { $a.Version = $fix.Version }
    }

    if ($ov.Extra) {
        foreach ($e in $ov.Extra) {
            if (-not $e.Path -or -not (Test-Path $e.Path)) { continue }
            [void]$apps.Add((New-App -Name $e.Name -Version "$($e.Version)" -Publisher "$($e.Publisher)" -Path $e.Path `
                        -Source 'Override' -Scope 'Пользователь' -Kind 'Портативно'))
        }
    }
    return $apps
}

function Test-Apps($apps) {
    Write-Step 'Валидация: пустые папки, мёртвые пути, размеры...'
    $out = New-Object System.Collections.ArrayList
    foreach ($a in $apps) {
        if ($a.Name -match '(?i)Application Compatibility Fix Database') { continue }
        if ($a.Name -match '^Отмена установки принтера (.+)$') { $a.Name = "Драйвер принтера $($matches[1])" }

        $exists = $false
        if ($a.Path -and $a.Path -match '^[A-Za-z]:\\') { $exists = Test-Path -LiteralPath $a.Path }

        if ($a.Path -and $a.Path -match '^[A-Za-z]:\\' -and -not $exists) {
            Add-Skipped $a.Name $a.Path 0 0 'путь не найден (остаток удаления)'
            continue
        }

        if ($exists) {
            $st = Get-FolderStats $a.Path
            # у Steam папка steamapps — это игры, их считаем отдельно
            if ($a.Name -eq 'Steam' -and (Test-Path (Join-Path $a.Path 'steamapps'))) {
                $sa = Get-FolderStats (Join-Path $a.Path 'steamapps')
                if ($sa) {
                    $st = [pscustomobject]@{
                        SizeMB = [math]::Max(0, $st.SizeMB - $sa.SizeMB)
                        Files  = [math]::Max(0, $st.Files - $sa.Files)
                    }
                }
            }
            if ($null -eq $a.SizeMB) { $a.SizeMB = $st.SizeMB }
            if ($null -eq $a.Files) { $a.Files = $st.Files }

            if ($a.Category -eq 'Игра') {
                if ($st.SizeMB -lt $MinGameSizeMB) { Add-Skipped $a.Name $a.Path $st.SizeMB $st.Files "игра: меньше $MinGameSizeMB МБ"; continue }
                if ($st.Files -lt 10) { Add-Skipped $a.Name $a.Path $st.SizeMB $st.Files 'игра: меньше 10 файлов'; continue }
            }
            elseif ($a.Kind -eq 'Портативно' -and $st.SizeMB -lt $MinPortableSizeMB) {
                Add-Skipped $a.Name $a.Path $st.SizeMB $st.Files "portable: меньше $MinPortableSizeMB МБ"; continue
            }
        }

        if (-not $a.Path) { $a.Exists = '' }
        elseif ($a.Path -match '^\\\\') { $a.Exists = 'да' }
        elseif ($exists) { $a.Exists = 'да' }
        else { $a.Exists = 'нет' }

        [void]$out.Add($a)
    }
    return $out
}

# ---------------------------------------------------------------- 11b. history
function Get-ChangeSet($apps, [string]$historyPath) {
    $res = @{ HasPrev = $false; New = @{}; Updated = @{}; Removed = @() }
    if (-not (Test-Path $historyPath)) { return $res }
    try { $prev = Get-Content $historyPath -Raw | ConvertFrom-Json } catch { return $res }
    if (-not $prev.items) { return $res }
    $res.HasPrev = $true

    $prevByKey = @{}
    $prevByName = @{}
    foreach ($it in $prev.items) {
        $prevByKey[$it.k] = $it
        if (-not $prevByName.ContainsKey($it.n)) { $prevByName[$it.n] = New-Object System.Collections.ArrayList }
        [void]$prevByName[$it.n].Add($it)
    }

    $used = @{}
    foreach ($a in $apps) {
        $norm = Get-Norm $a.Name
        $k = "$norm|$("$($a.Path)".ToLower())"
        $hit = $null
        if ($prevByKey.ContainsKey($k)) { $hit = $prevByKey[$k] }
        else {
            # путь мог поменяться (версия Store-пакета, Discord app-1.0.x) — ищем по имени
            if ($prevByName.ContainsKey($norm)) {
                foreach ($cand in $prevByName[$norm]) {
                    if (-not $used.ContainsKey($cand.k)) { $hit = $cand; break }
                }
            }
        }
        if (-not $hit) { $res.New[$k] = 1; continue }
        $used[$hit.k] = 1
        $verChanged = "$($hit.v)" -ne "$($a.Version)"
        $sizeChanged = ([math]::Abs([double]$hit.s - [double]$a.SizeMB) -gt 1)
        if ($verChanged -or $sizeChanged) { $res.Updated[$k] = 1 }
    }
    $res.Removed = @($prev.items | Where-Object { -not $used.ContainsKey($_.k) })
    return $res
}

function Save-Snapshot($apps, [string]$historyPath) {
    $items = @($apps | ForEach-Object {
            [pscustomobject]@{
                k = (Get-Norm $_.Name) + '|' + "$($_.Path)".ToLower()
                n = Get-Norm $_.Name
                v = "$($_.Version)"
                s = $_.SizeMB
                p = $_.Path
            }
        })
    $doc = [pscustomobject]@{ generated = (Get-Date -Format 'yyyy-MM-dd HH:mm'); items = $items }
    $doc | ConvertTo-Json -Depth 4 | Out-File -FilePath $historyPath -Encoding UTF8
}

# ---------------------------------------------------------------- 11c. usage
function Rot13([string]$s) {
    $chars = $s.ToCharArray()
    for ($i = 0; $i -lt $chars.Length; $i++) {
        $c = $chars[$i]
        if ($c -ge 'a' -and $c -le 'z') { $chars[$i] = [char]((([int]$c - 97 + 13) % 26) + 97) }
        elseif ($c -ge 'A' -and $c -le 'Z') { $chars[$i] = [char]((([int]$c - 65 + 13) % 26) + 65) }
    }
    return -join $chars
}

function Get-UsageIndex {
    $res = New-Object System.Collections.ArrayList
    $root = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist'
    $known = @{
        '{6D809377-6AF0-444B-8957-A3773F02200E}' = 'C:\Program Files'
        '{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}' = 'C:\Program Files (x86)'
        '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}' = 'C:\Windows\System32'
        '{F38BF404-1D43-42F2-9305-67DE0B28FC23}' = 'C:\Windows'
    }
    foreach ($k in (Get-ChildItem $root -ErrorAction SilentlyContinue)) {
        $item = Get-Item (Join-Path $k.PSPath 'Count') -ErrorAction SilentlyContinue
        if (-not $item) { continue }
        foreach ($name in $item.GetValueNames()) {
            $decoded = Rot13 $name
            if ($decoded -match '^\{[0-9A-Fa-f\-]+\}\\(.+)$') {
                $g = $decoded.Substring(0, $decoded.IndexOf('}') + 1)
                $rest = $matches[1]
                if ($known.ContainsKey($g)) { $decoded = Join-Path $known[$g] $rest }
                else { $decoded = $rest }
            }
            if ($decoded -notmatch '^[A-Za-z]:\\') { continue }
            $raw = $item.GetValue($name)
            if ($raw.Length -lt 68) { continue }
            $ft = [BitConverter]::ToInt64($raw, 60)
            if ($ft -le 0) { continue }
            try { $last = [DateTime]::FromFileTime($ft) } catch { continue }
            if ($last.Year -lt 2000) { continue }
            [void]$res.Add([pscustomobject]@{ Path = $decoded; Last = $last })
        }
    }
    return $res
}

function Add-Usage($apps) {
    $usage = Get-UsageIndex
    if (-not $usage.Count) { return $apps }
    Write-Step ("  данные о запусках: {0} записей" -f $usage.Count)
    foreach ($a in $apps) {
        if (-not $a.Path) { continue }
        $p = $a.Path.ToLower()
        $best = $null
        foreach ($u in $usage) {
            $up = $u.Path.ToLower()
            if ($up -eq $p -or $up.StartsWith($p + '\')) {
                if (-not $best -or $u.Last -gt $best) { $best = $u.Last }
            }
        }
        if (-not $best) {
            $leaf = Split-Path -Leaf $a.Path
            foreach ($u in $usage) {
                if ((Split-Path -Leaf $u.Path) -ieq $leaf) { if (-not $best -or $u.Last -gt $best) { $best = $u.Last } }
            }
        }
        if ($best) { $a.LastRun = $best }
    }
    return $apps
}

# ---------------------------------------------------------------- 12. report
function Esc([string]$s) { if (-not $s) { return '' }; return [System.Net.WebUtility]::HtmlEncode($s) }

function Format-Size($mb) {
    if ($null -eq $mb -or $mb -le 0) { return '—' }
    if ($mb -ge 1024) { return ('{0:N1} ГБ' -f ($mb / 1024)) }
    return ('{0:N0} МБ' -f $mb)
}

function Export-Report($apps, $skipped, $changes) {
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $catOrder = @{ 'Приложение' = 0; 'Инструмент' = 1; 'Игра' = 2; 'Store' = 3; 'Системное' = 4 }
    $all = $apps | Sort-Object @{ Expression = { if ($catOrder.ContainsKey($_.Category)) { $catOrder[$_.Category] } else { 9 } } }, Name

    $all | Select-Object Category, Name, Version, Publisher, SizeMB, Files, Path, Source, Scope, Kind, Exists, Uninstall,
        @{ n = 'Change'; e = { $k = (Get-Norm $_.Name) + '|' + "$($_.Path)".ToLower(); if ($changes.New.ContainsKey($k)) { 'новое' } elseif ($changes.Updated.ContainsKey($k)) { 'обновлено' } else { '' } } },
        @{ n = 'LastRun'; e = { if ($_.LastRun) { $_.LastRun.ToString('yyyy-MM-dd HH:mm') } else { '' } } } |
        Export-Csv -Path (Join-Path $OutputDir 'software-inventory.csv') -NoTypeInformation -Encoding UTF8

    if ($changes.HasPrev) {
        $rows = @()
        $rows += $all | Where-Object { $changes.New.ContainsKey((Get-Norm $_.Name) + '|' + "$($_.Path)".ToLower()) } |
            ForEach-Object { [pscustomobject]@{ Change = 'новое'; Name = $_.Name; Version = $_.Version; SizeMB = $_.SizeMB; Path = $_.Path } }
        $rows += $all | Where-Object { $changes.Updated.ContainsKey((Get-Norm $_.Name) + '|' + "$($_.Path)".ToLower()) } |
            ForEach-Object { [pscustomobject]@{ Change = 'обновлено'; Name = $_.Name; Version = $_.Version; SizeMB = $_.SizeMB; Path = $_.Path } }
        $rows += $changes.Removed | ForEach-Object { [pscustomobject]@{ Change = 'пропало'; Name = $_.n; Version = $_.v; SizeMB = $_.s; Path = $_.p } }
        $rows | Export-Csv -Path (Join-Path $OutputDir 'changes.csv') -NoTypeInformation -Encoding UTF8
    }

    $skipped | Sort-Object Name | Export-Csv -Path (Join-Path $OutputDir 'excluded.csv') -NoTypeInformation -Encoding UTF8

    $all | Select-Object Category, Name, Version, Publisher, SizeMB, Files, Path, Source, Scope, Kind, Exists |
        ConvertTo-Json -Depth 3 | Out-File -FilePath (Join-Path $OutputDir 'software-inventory.json') -Encoding UTF8

    $md = New-Object System.Text.StringBuilder
    [void]$md.AppendLine('# Установленное ПО')
    [void]$md.AppendLine('')
    [void]$md.AppendLine("Сформировано: $stamp · всего записей: **$($all.Count)** · исключено папок: **$($skipped.Count)**")
    [void]$md.AppendLine('')
    foreach ($cat in @('Приложение', 'Инструмент', 'Игра', 'Store', 'Системное')) {
        $grp = @($all | Where-Object Category -eq $cat)
        if (-not $grp.Count) { continue }
        [void]$md.AppendLine("## $cat ($($grp.Count))")
        [void]$md.AppendLine('')
        [void]$md.AppendLine('| Название | Версия | Издатель | Размер | Путь |')
        [void]$md.AppendLine('|---|---|---|---|---|')
        foreach ($a in $grp) {
            $p = if ($a.Path) { $a.Path } else { '—' }
            if ($a.Exists -eq 'нет') { $p = "$p *(нет папки)*" }
            $v = if ($a.Version) { $a.Version } else { '—' }
            $pub = if ($a.Publisher) { $a.Publisher } else { '—' }
            [void]$md.AppendLine("| $($a.Name -replace '\|','/') | $($v -replace '\|','/') | $($pub -replace '\|','/') | $(Format-Size $a.SizeMB) | ``$($p -replace '\|','/')`` |")
        }
        [void]$md.AppendLine('')
    }
    [System.IO.File]::WriteAllText((Join-Path $OutputDir 'software-inventory.md'), $md.ToString(), (New-Object System.Text.UTF8Encoding $true))

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8">')
    [void]$sb.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1">')
    [void]$sb.AppendLine('<title>Установленное ПО</title><style>')
    [void]$sb.AppendLine(@'
:root{--bg:#0f1115;--panel:#171a21;--line:#262b36;--txt:#e6e9ef;--dim:#8b93a7;--acc:#5b9dff}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--txt);font:14px/1.45 "Segoe UI",system-ui,sans-serif}
header{padding:24px 32px 14px;border-bottom:1px solid var(--line);position:sticky;top:0;background:rgba(15,17,21,.97);backdrop-filter:blur(8px);z-index:5}
h1{margin:0 0 4px;font-size:21px;font-weight:600}
.sub{color:var(--dim);font-size:13px}
.controls{display:flex;gap:9px;flex-wrap:wrap;align-items:center;margin-top:12px}
input[type=search]{flex:1;min-width:230px;padding:8px 12px;border-radius:8px;border:1px solid var(--line);background:#0b0d11;color:var(--txt);font-size:14px;outline:none}
input[type=search]:focus{border-color:var(--acc)}
.chip{padding:6px 11px;border-radius:999px;border:1px solid var(--line);background:var(--panel);color:var(--dim);cursor:pointer;font-size:13px;user-select:none;white-space:nowrap}
.chip.on{border-color:var(--acc);color:#fff;background:#1b2b45}
.chip b{color:var(--txt);font-weight:600;margin-left:6px}
.sw{display:flex;align-items:center;gap:7px;font-size:13px;color:var(--dim);cursor:pointer;user-select:none;padding:6px 11px;border:1px solid var(--line);border-radius:8px;background:var(--panel)}
.sw input{accent-color:var(--acc);width:15px;height:15px;cursor:pointer}
main{padding:4px 32px 60px}
table{width:100%;border-collapse:collapse;margin-top:6px}
th{position:sticky;top:136px;background:var(--bg);text-align:left;font-size:12px;letter-spacing:.04em;text-transform:uppercase;color:var(--dim);padding:9px 12px;border-bottom:1px solid var(--line);cursor:pointer;white-space:nowrap}
td{padding:8px 12px;border-bottom:1px solid #1d212b;vertical-align:top}
tr:hover td{background:#141821}
.n{color:var(--dim);font-variant-numeric:tabular-nums;width:44px}
.name{font-weight:600}
.ver{color:var(--dim);font-variant-numeric:tabular-nums;white-space:nowrap}
.pub{color:#b9c1d1}
.size{color:#9fe0b0;font-variant-numeric:tabular-nums;white-space:nowrap}
.path{font-family:Consolas,"Cascadia Mono",monospace;font-size:12px;color:#9fb3d1;word-break:break-all}
.tag{display:inline-block;padding:2px 8px;border-radius:5px;font-size:11px;font-weight:600;white-space:nowrap}
.t-app{background:#1d3a2a;color:#7ee2a8}
.t-tool{background:#2a2f3a;color:#9fb3d1}
.t-game{background:#3a1e2f;color:#ff8fc7}
.t-store{background:#1e2f4a;color:#7fb3ff}
.t-sys{background:#3a2a1e;color:#ffb87f}
.badge{display:inline-block;padding:1px 6px;border-radius:4px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.03em;vertical-align:middle}
.b-new{background:#204a35;color:#6ee7a8}
.b-upd{background:#3b3a1e;color:#e7d96e}
.plink{color:#9fb3d1;text-decoration:none;border-bottom:1px dotted #4a5568}
.plink:hover{color:#cfe0ff;border-bottom-color:#5b9dff}
.run{color:#9fb3d1;font-variant-numeric:tabular-nums;white-space:nowrap}
.run.never{color:#5f6675}
.empty{padding:40px;text-align:center;color:var(--dim)}
@media(max-width:900px){header,main{padding-left:14px;padding-right:14px}th:nth-child(3),td:nth-child(3){display:none}}
'@)
    [void]$sb.AppendLine('</style></head><body>')
    [void]$sb.AppendLine('<header><h1>Установленное ПО</h1>')
    $hostname = $env:COMPUTERNAME
    $os = (Get-CimInstance Win32_OperatingSystem).Caption
    $chg = ''
    if ($changes.HasPrev) {
        $chg = " · новых: $($changes.New.Count) · обновлено: $($changes.Updated.Count) · пропало: $($changes.Removed.Count)"
    }
    [void]$sb.AppendLine("<div class=`"sub`">$hostname · $os · $stamp · показано: <b id=`"total`">$($all.Count)</b> из $($all.Count) · исключено: $($skipped.Count)$chg</div>")
    [void]$sb.AppendLine('<div class="controls"><input type="search" id="q" placeholder="Поиск: название, издатель, путь...">')
    foreach ($cat in @('Приложение', 'Инструмент', 'Игра', 'Store', 'Системное')) {
        $cnt = @($all | Where-Object Category -eq $cat).Count
        if ($cnt -gt 0) { [void]$sb.AppendLine("<div class=`"chip on`" data-c=`"$cat`">$cat<b>$cnt</b></div>") }
    }
    [void]$sb.AppendLine("<label class=`"sw`"><input type=`"checkbox`" id=`"hideSys`" $(if (-not $IncludeSystem) { 'checked' })> скрыть служебное</label>")
    [void]$sb.AppendLine("<label class=`"sw`"><input type=`"checkbox`" id=`"hideMs`"> скрыть Microsoft</label>")
    [void]$sb.AppendLine("<label class=`"sw`" title=`"журнал проводника: программы из автозагрузки и консольные могут не иметь записи`"><input type=`"checkbox`" id=`"onlyUnused`"> нет записи о запуске</label>")
    if ($changes.HasPrev) {
        [void]$sb.AppendLine("<label class=`"sw`"><input type=`"checkbox`" id=`"onlyNew`"> только изменения</label>")
    }
    [void]$sb.AppendLine('</div></header><main><table id="t"><thead><tr><th class="n">#</th><th data-s="name">Название</th><th data-s="ver">Версия</th><th data-s="pub">Издатель</th><th data-s="size">Размер</th><th data-s="run">Запуск</th><th data-s="path">Путь</th><th data-s="cat">Категория</th></tr></thead><tbody>')

    $i = 0
    foreach ($a in $all) {
        $i++
        $cls = switch ($a.Category) { 'Приложение' { 't-app' } 'Инструмент' { 't-tool' } 'Игра' { 't-game' } 'Store' { 't-store' } 'Системное' { 't-sys' } }
        $p = if ($a.Path) { $a.Path } else { '—' }
        $v = if ($a.Version) { $a.Version } else { '—' }
        $pub = if ($a.Publisher) { $a.Publisher } else { '—' }
        $key = (Get-Norm $a.Name) + '|' + "$($a.Path)".ToLower()
        $isNew = if ($changes.New.ContainsKey($key)) { '1' } else { '0' }
        $isUpd = if ($changes.Updated.ContainsKey($key)) { '1' } else { '0' }
        $badge = ''
        if ($isNew -eq '1') { $badge = ' <span class="badge b-new">новое</span>' }
        elseif ($isUpd -eq '1') { $badge = ' <span class="badge b-upd">обновлено</span>' }

        $pathCell = Esc $p
        if ($a.Exists -eq 'нет') { $pathCell = "$pathCell <span class=`"gone`">(папки нет)</span>" }
        elseif ($a.Path -and $a.Path -match '^[A-Za-z]:\\') {
            $target = $a.Path
            if ($target -match '\\[^\\]+\.[A-Za-z0-9]+$') { $target = Split-Path -Parent $target }
            $clean = $target -replace '\\', '/'
            $href = 'file:///' + [uri]::EscapeUriString($clean)
            $pathCell = "<a class=`"plink`" href=`"$href`" title=`"Открыть папку`">$pathCell</a>"
        }
        $sys = if ($a.Category -eq 'Системное') { '1' } else { '0' }
        $ms = if ($a.Publisher -match '(?i)Microsoft|Майкрософт') { '1' } else { '0' }
        $run = if ($a.LastRun) { $a.LastRun.ToString('dd.MM.yy') } else { '—' }
        $runV = if ($a.LastRun) { $a.LastRun.ToString('yyyyMMdd') } else { '0' }
        $runCls = if ($a.LastRun) { 'run' } else { 'run never' }
        [void]$sb.AppendLine("<tr data-c=`"$($a.Category)`" data-sys=`"$sys`" data-ms=`"$ms`" data-new=`"$isNew`" data-upd=`"$isUpd`" data-run=`"$(if ($a.LastRun) { 1 } else { 0 })`" data-mb=`"$($a.SizeMB)`"><td class=`"n`">$i</td><td class=`"name`" data-v=`"$(Esc $a.Name)`">$(Esc $a.Name)$badge</td><td class=`"ver`" data-v=`"$(Esc $v)`">$(Esc $v)</td><td class=`"pub`" data-v=`"$(Esc $pub)`">$(Esc $pub)</td><td class=`"size`" data-v=`"$($a.SizeMB)`">$(Format-Size $a.SizeMB)</td><td class=`"$runCls`" data-v=`"$runV`" title=`"последний запуск по данным проводника`">$run</td><td class=`"path`" data-v=`"$(Esc $p)`">$pathCell</td><td><span class=`"tag $cls`">$($a.Category)</span></td></tr>")
    }
    [void]$sb.AppendLine('</tbody></table><div class="empty" id="none" style="display:none">Ничего не найдено</div></main>')
    [void]$sb.AppendLine(@'
<script>
const q=document.getElementById('q'),rows=[...document.querySelectorAll('#t tbody tr')],none=document.getElementById('none');
const chips=[...document.querySelectorAll('.chip')],hideSys=document.getElementById('hideSys'),hideMs=document.getElementById('hideMs'),onlyNew=document.getElementById('onlyNew'),onlyUnused=document.getElementById('onlyUnused');
let active=new Set(chips.filter(c=>c.classList.contains('on')).map(c=>c.dataset.c));
function apply(){const s=q.value.toLowerCase().trim();let n=0;
 rows.forEach(r=>{const okC=active.has(r.dataset.c);const okS=!s||r.textContent.toLowerCase().includes(s);
  const okY=!(hideSys.checked&&r.dataset.sys==='1');const okM=!(hideMs&&hideMs.checked&&r.dataset.ms==='1');
  const okN=!(onlyNew&&onlyNew.checked&&r.dataset.new!=='1'&&r.dataset.upd!=='1');
  const okU=!(onlyUnused&&onlyUnused.checked&&r.dataset.run!=='0');
  const show=okC&&okS&&okY&&okM&&okN&&okU;r.style.display=show?'':'none';if(show)n++;});
 none.style.display=n?'none':'block';document.getElementById('total').textContent=n;}
q.addEventListener('input',apply);hideSys.addEventListener('change',apply);if(hideMs)hideMs.addEventListener('change',apply);if(onlyNew)onlyNew.addEventListener('change',apply);if(onlyUnused)onlyUnused.addEventListener('change',apply);
chips.forEach(c=>c.addEventListener('click',()=>{const k=c.dataset.c;if(active.has(k)){active.delete(k);c.classList.remove('on')}else{active.add(k);c.classList.add('on')}apply();}));
document.querySelectorAll('th[data-s]').forEach(th=>{let asc=true;th.addEventListener('click',()=>{const tb=document.querySelector('#t tbody');const rs=[...tb.rows];const idx=[...th.parentNode.children].indexOf(th);
 rs.sort((a,b)=>{const x=a.cells[idx].dataset.v||a.cells[idx].textContent,y=b.cells[idx].dataset.v||b.cells[idx].textContent;return x.localeCompare(y,'ru',{numeric:true})*(asc?1:-1)});
 asc=!asc;rs.forEach(r=>tb.appendChild(r));});});
apply();
</script></body></html>
'@)
    $htmlPath = Join-Path $OutputDir 'software-inventory.html'
    [System.IO.File]::WriteAllText($htmlPath, $sb.ToString(), (New-Object System.Text.UTF8Encoding $true))
    return $htmlPath
}

# ---------------------------------------------------------------- 13. audit
function Invoke-Audit($apps, $skipped, [string]$file) {
    Write-Step 'Аудит покрытия...'
    $paths = @()
    foreach ($a in $apps) {
        if (-not $a.Path) { continue }
        $p = $a.Path.ToLower()
        if ($p -match '\\[^\\]+\.[a-z0-9]+$') { $p = (Split-Path -Parent $a.Path).ToLower() }
        $paths += $p
    }
    $idx = Get-DiskIndex
    $lines = @()
    foreach ($g in ($idx.Exes | Group-Object Dir)) {
        $d = $g.Name.ToLower()
        $hit = $false
        foreach ($x in $paths) { if ($d -eq $x -or $d.StartsWith($x + '\') -or $x.StartsWith($d + '\')) { $hit = $true; break } }
        if ($hit) { continue }
        $knownSkip = $skipped | Where-Object { $_.Path -and $_.Path.ToLower() -eq $d }
        $reason = if ($knownSkip) { ($knownSkip | Select-Object -First 1).Reason } else { 'не классифицировано' }
        $lines += ('{0,9:N1} МБ  {1}  [{2}]' -f [math]::Round($idx.Sizes[$g.Name] / 1MB, 1), $g.Name, $reason)
    }
    $lines = $lines | Sort-Object -Unique
    $lines | Out-File -FilePath $file -Encoding UTF8
    return $lines
}

# ---------------------------------------------------------------- main
$script:Skipped = New-Object System.Collections.ArrayList
$apps = New-Object System.Collections.ArrayList

Write-Host 'Сбор данных...' -ForegroundColor Cyan
foreach ($a in (Get-RegistryApps)) { [void]$apps.Add($a) }
if (-not $NoStore) { foreach ($a in (Get-StoreApps)) { [void]$apps.Add($a) } }
if (-not $NoGames) {
    foreach ($a in (Get-SteamApps)) { [void]$apps.Add($a) }
    foreach ($a in (Get-EpicApps)) { [void]$apps.Add($a) }
    foreach ($a in (Get-GOGApps)) { [void]$apps.Add($a) }
}
foreach ($a in (Get-ScoopApps)) { [void]$apps.Add($a) }
foreach ($a in (Get-WslApps)) { [void]$apps.Add($a) }
foreach ($a in (Get-ToolchainApps)) { [void]$apps.Add($a) }

if (-not $NoPortable) {
    foreach ($a in (Get-DiskApps $apps)) { [void]$apps.Add($a) }
    if (-not $SkipShortcuts) { foreach ($a in (Get-ShortcutApps $apps)) { [void]$apps.Add($a) } }
}

Write-Step 'Склейка дублей...'
$apps = Merge-Apps $apps

if (-not $OverridesPath) { $OverridesPath = Join-Path $PSScriptRoot 'overrides.psd1' }
$apps = Apply-Overrides $apps $OverridesPath
$apps = Test-Apps $apps
$apps = Add-Usage $apps

$historyPath = Join-Path $OutputDir 'history.json'
$changes = Get-ChangeSet $apps $historyPath

$html = Export-Report $apps $script:Skipped $changes
Save-Snapshot $apps $historyPath

if ($Audit) {
    $auditLines = Invoke-Audit $apps $script:Skipped (Join-Path $OutputDir 'audit.txt')
    Write-Host ''
    if ($auditLines.Count) {
        Write-Host "Не покрыто отчётом: $($auditLines.Count)" -ForegroundColor Yellow
        $auditLines | Select-Object -First 25 | ForEach-Object { Write-Host "  $_" }
        Write-Host "  ...полный список: $OutputDir\audit.txt" -ForegroundColor DarkGray
    }
    else {
        Write-Host 'Аудит: каждая папка с exe либо в отчёте, либо в excluded.csv.' -ForegroundColor Green
    }
}

Write-Host ''
Write-Host 'Готово.' -ForegroundColor Green
foreach ($g in ($apps | Group-Object Category | Sort-Object Name)) { Write-Host ("  {0,-12} {1}" -f $g.Name, $g.Count) }
Write-Host ("  Всего:       {0}" -f $apps.Count)
if ($script:Skipped.Count) { Write-Host ("  Исключено:   {0} (см. excluded.csv)" -f $script:Skipped.Count) -ForegroundColor DarkYellow }
if ($changes.HasPrev) {
    Write-Host ("  Новых: {0} · обновлено: {1} · пропало: {2} (changes.csv)" -f $changes.New.Count, $changes.Updated.Count, $changes.Removed.Count) -ForegroundColor Cyan
}
Write-Host ''
Write-Host "Отчёты: $OutputDir" -ForegroundColor Cyan

if ($Open) { Start-Process $html }
