#Requires -Version 5.1
<#
.SYNOPSIS
    Инвентаризация установленного ПО Windows: реестр, Store, Steam/Epic/GOG, Scoop, portable.

.DESCRIPTION
    Собирает установленные программы из нескольких независимых источников,
    проверяет, что папки реально существуют и не пустые, и формирует отчёт
    (HTML + CSV + Markdown + JSON) с путями и размерами.

    Папки-пустышки и огрызки игр отбрасываются: для игр по умолчанию нужно
    не меньше 100 МБ и 10 файлов, для portable — 3 МБ. Всё отброшенное
    пишется в reports/skipped.csv с причиной.

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
    Не собирать игры (Steam/Epic/GOG/папки).

.PARAMETER NoPortable
    Не сканировать файловую систему на portable-программы.

.PARAMETER ExtraRoots
    Дополнительные папки для поиска portable-программ.

.PARAMETER SkipShortcuts
    Не разбирать ярлыки меню «Пуск» и рабочего стола.

.PARAMETER Open
    Открыть HTML-отчёт в браузере после сборки.

.EXAMPLE
    .\Get-InstalledSoftware.ps1

.EXAMPLE
    .\Get-InstalledSoftware.ps1 -MinGameSizeMB 50 -Open

.EXAMPLE
    .\Get-InstalledSoftware.ps1 -OutputDir D:\tmp\inventory -IncludeSystem
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
    return $n.Trim()
}

$script:SizeCache = @{}
function Get-FolderStats([string]$path) {
    if (-not $path) { return $null }
    if ($script:SizeCache.ContainsKey($path)) { return $script:SizeCache[$path] }
    $result = [pscustomobject]@{ SizeMB = 0; Files = 0 }
    if ($path -match '^[A-Za-z]:\\' -and (Test-Path -LiteralPath $path -PathType Leaf)) {
        $f = Get-Item -LiteralPath $path
        $result.SizeMB = [math]::Round($f.Length / 1MB, 1)
        $result.Files = 1
        $script:SizeCache[$path] = $result
        return $result
    }
    if ($path -match '^[A-Za-z]:\\' -and (Test-Path -LiteralPath $path -PathType Container)) {
        $sum = 0L; $cnt = 0
        foreach ($f in (Get-ChildItem -LiteralPath $path -Recurse -File -Force -ErrorAction SilentlyContinue)) {
            $sum += $f.Length; $cnt++
        }
        $result.SizeMB = [math]::Round($sum / 1MB, 1)
        $result.Files = $cnt
    }
    $script:SizeCache[$path] = $result
    return $result
}

function Add-Skipped([string]$name, [string]$path, [double]$mb, [int]$files, [string]$reason) {
    [void]$script:Skipped.Add([pscustomobject]@{
            Name = $name; Path = $path; SizeMB = $mb; Files = $files; Reason = $reason
        })
}

$SystemRx = '(?i)Visual C\+\+|\.NET (SDK|Runtime|Desktop Runtime)|Windows Desktop Runtime|Windows SDK|DirectX|Edge WebView2|Edge Update|Epic Online Services|Mozilla Maintenance|AMD[ _]Chipset|AMD[ _]Install|Realtek Card Reader|EpsonNet|Драйвер принтера|Application Compatibility|Application Virtualization|Windows Subsystem for Linux|^WSL|WinAppRuntime|ShellExtension|ContextMenu|SparseApp|Winget\.Source|DesktopAppInstaller|StorePurchaseApp|LanguageExperiencePack|VCLibs|Microsoft\.UI\.Xaml|HostRuntime|AppRuntime'
$ToolRx   = '(?i)Visual Studio Code|Android Studio|IntelliJ|PyCharm|WebStorm|^Git$|GitHub CLI|Node\.js|Python|OpenJDK|Temurin|JDK|Docker|Postman|Scoop|Maestro|Cursor|Devin|Antigravity|OpenCode|DeepSeek|Blender|Godot|NekoBox|LockHunter|WinRAR|7-Zip|PowerToys|zapret|Ubuntu'
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
        Path      = if ($Path) { ($Path -replace '/', '\').TrimEnd('\') } else { '' }
        Category  = (Get-Category $clean $Path $Source $Kind)
        Source    = $Source
        Scope     = $Scope
        Kind      = $Kind
        Uninstall = $Uninstall
        SizeMB    = $null
        Files     = $null
        Exists    = ''
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
    $out = New-Object System.Collections.ArrayList
    # карта «exe -> путь» из App Paths: помогает достать путь там, где реестр его не даёт
    $appPaths = @{}
    foreach ($hive in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\*',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\*')) {
        Get-ItemProperty $hive -ErrorAction SilentlyContinue | ForEach-Object {
            $exe = "$($_.PSChildName)" -replace '\.exe$', ''
            $p = $_.'(default)'
            if ($exe -and $p -and $p -match '^[A-Za-z]:\\') { $appPaths[(Get-Norm $exe)] = $p }
        }
    }

    foreach ($k in $keys) {
        Get-ItemProperty $k.Path | ForEach-Object {
            if (-not $_.DisplayName) { return }
            if ($_.SystemComponent -eq 1) { return }
            if ($_.ParentKeyName) { return }

            $loc = $_.InstallLocation
            if ($loc) { $loc = $loc.Trim('"').TrimEnd('\') }
            if ($loc -match '(?i)Package Cache|\\Installer\\|\\Windows\\Installer') { $loc = '' }

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
        [void]$out.Add((New-App -Name $_.Name -Version "$($_.Version)" -Publisher "$($_.PublisherDisplayName)" `
                    -Path $_.InstallLocation -Source 'Microsoft Store' `
                    -Scope $(if ($_.InstallLocation -like "$env:LOCALAPPDATA*") { 'Пользователь' } else { 'Система' }) `
                    -Uninstall "Get-AppxPackage -Name '$($_.Name)' | Remove-AppxPackage" -Kind 'Store'))
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
            Get-Content $vdf | ForEach-Object {
                if ($_ -match '"path"\s+"(.+?)"') { $libs += ($matches[1] -replace '\\\\', '\') }
            }
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
    # папки без манифеста: берём только те, где реально лежит игра, а не пустой огрызок
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

    $sdkPaths = @($env:ANDROID_HOME, $env:ANDROID_SDK_ROOT, "$env:LOCALAPPDATA\Android\Sdk") | Where-Object { $_ }
    foreach ($p in $sdkPaths) {
        if (-not (Test-Path $p)) { continue }
        $a = New-App -Name 'Android SDK' -Version '' -Publisher 'Google' -Path $p -Source 'Toolchain' -Scope 'Пользователь' -Kind 'Портативно'
        $st = Get-FolderStats $p
        $a.SizeMB = $st.SizeMB; $a.Files = $st.Files
        [void]$out.Add($a)
    }

    # Scoop (сам менеджер, не пакеты)
    $scoopRoot = Join-Path $env:USERPROFILE 'scoop'
    if (Test-Path (Join-Path $scoopRoot 'shims\scoop.cmd')) {
        $a = New-App -Name 'Scoop' -Version '' -Publisher 'Scoop' -Path $scoopRoot -Source 'Toolchain' -Scope 'Пользователь' -Kind 'Портативно'
        $st = Get-FolderStats $scoopRoot
        $a.SizeMB = $st.SizeMB; $a.Files = $st.Files
        [void]$out.Add($a)
    }

    # Microsoft Edge WebView2 Runtime
    $wvRoot = 'C:\Program Files (x86)\Microsoft\EdgeWebView\Application'
    if (Test-Path $wvRoot) {
        $ver = Get-ChildItem $wvRoot -Directory | Where-Object { $_.Name -match '^\d+\.' } | Sort-Object Name -Descending | Select-Object -First 1
        $a = New-App -Name 'Microsoft Edge WebView2 Runtime' -Version $(if ($ver) { $ver.Name } else { '' }) -Publisher 'Microsoft' -Path $wvRoot -Source 'Toolchain' -Kind 'Установлено'
        $st = Get-FolderStats $wvRoot
        $a.SizeMB = $st.SizeMB; $a.Files = $st.Files
        [void]$out.Add($a)
    }

    # папки-наборы: zapret (bin\winws.exe), копии Android SDK (platform-tools + build-tools)
    foreach ($drv in (Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -match '^[A-Z]:\\$' })) {
        Get-ChildItem $drv.Root -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $d = $_.FullName
            if (Test-Path (Join-Path $d 'bin\winws.exe')) {
                $a = New-App -Name 'zapret (обход блокировок)' -Version '' -Publisher '' -Path $d -Source 'Toolchain' -Scope 'Пользователь' -Kind 'Портативно'
                $st = Get-FolderStats $d
                $a.SizeMB = $st.SizeMB; $a.Files = $st.Files
                [void]$out.Add($a)
            }
            elseif ((Test-Path (Join-Path $d 'platform-tools\adb.exe')) -and (Test-Path (Join-Path $d 'build-tools'))) {
                $parts = @('platform-tools', 'build-tools', 'emulator', 'platforms', 'sources', 'extras', 'system-images') |
                    ForEach-Object { Join-Path $d $_ } | Where-Object { Test-Path $_ }
                $sizeMB = 0.0; $files = 0
                foreach ($part in $parts) {
                    $st = Get-FolderStats $part
                    if ($st) { $sizeMB += $st.SizeMB; $files += $st.Files }
                }
                $a = New-App -Name 'Android SDK (копия)' -Version '' -Publisher 'Google' -Path $d -Source 'Toolchain' -Scope 'Пользователь' -Kind 'Портативно'
                $a.SizeMB = [math]::Round($sizeMB, 1); $a.Files = $files
                [void]$out.Add($a)
            }
        }
    }
    return $out
}

# ---------------------------------------------------------------- 7. portable
$SkipDirRx = '(?i)\\(Windows|WinSxS|Installer|Temp|temp|Cache|cache|node_modules|\.git|locales|resources|plugins|Extensions|packages|Packages|logs|Logs|Crashpad|crashpad|Installers|redist|Redist|DirectX|DirectXRedist|dotnet|DotNet|Common Files|WindowsApps|ModifiableWindowsApps|Windows Defender|Windows Defender Advanced Threat Protection|Windows Mail|Windows Media Player|Windows Photo Viewer|Internet Explorer|Microsoft Office( 15)?|EdgeCore|EdgeUpdate|CNext|Epic Online Services|WSL|Razer|AMD|tcl|Scripts|\.tlauncher|\.minecraft|Tick_Tick|Battle\.net|SDK|platforms|archives|licenses|sources|\.temp|jbr|jre|jdk[\w\.\-]*|runtime|lib|lib64|bin|bin64|drivers|include|qemu|__pycache__|\.venv|venv)(\\|$)'
$SkipExeRx = '(?i)^(unins\d*|uninstall|uninst|setup|install|installer|update|updater|upgrade|crash|crashpad|crashreporter|reporter|helper|service|svc|elevate|vcredist|dxsetup|dotnet|node|npm|npx|python|pythonw|pip|ffmpeg|ffplay|ffprobe|7z|7za|rar|unrar|sqlite3|curl|wget|git|gitk|git-bash|bash|sh|cmd|powershell|pwsh|wsl|wslconfig|wslg|docker|dockerd|com\.docker|electron|chrome_crashpad_handler|nacl64|notification_helper|elevation_service|software_reporter_tool|msedgewebview2|msedge_proxy|msedge|chrome_proxy|chrome_pwa|googleupdate|updater_bootstrap|aapt2|lld|adb|fastboot|emulator-check|winws|sing-box|tun2proxy-bin|xray|antifilter|jabswitch|jps|jcmd|jstat|keytool|java|javaw|javac|javadoc|jar|jarsigner|jdeps|jlink|jmap|jstack|jconsole|jdb|jfr|jhsdb|jshell|jweblauncher|jaccessinspector|jaccesswalker|orbd|pack200|policytool|unpack200|wsgen|wsimport|xjc|vswhere|VsDevCmd|msbuild|nuget)\.exe$'

function Get-PortableRoots {
    $roots = @(
        'C:\Program Files',
        'C:\Program Files (x86)',
        "$env:LOCALAPPDATA\Programs",
        "$env:USERPROFILE\scoop\apps",
        'D:\apps', 'D:\soft', 'D:\desktop', 'D:\ai', 'D:\Program Files',
        'D:\Games', 'D:\torrent', 'D:\Epic Games', 'D:\GOG Games'
    )
    foreach ($d in (Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -match '^[A-Z]:\\$' })) { $roots += $d.Root }
    $roots += $ExtraRoots
    return ($roots | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | ForEach-Object { $_.TrimEnd('\') } | Sort-Object -Unique)
}

function Get-PortableApps($known) {
    Write-Step 'Файловая система (portable)...'
    $out = New-Object System.Collections.ArrayList
    $roots = Get-PortableRoots

    # папки уже найденных приложений — чтобы не тащить их внутренности как отдельные программы
    $knownDirs = @()
    foreach ($k in $known) {
        if (-not $k.Path) { continue }
        if ($k.Path -match '\\[^\\]+\.[a-zA-Z0-9]+$') { $knownDirs += (Split-Path -Parent $k.Path).ToLower() }
        else { $knownDirs += $k.Path.ToLower() }
    }

    $exes = @()
    foreach ($root in $roots) {
        $depth = if ($root -match '^[A-Z]:\\?$') { 2 } else { 3 }
        $exes += Get-ChildItem -LiteralPath $root -Filter '*.exe' -File -Recurse -Depth $depth |
            Where-Object { $_.FullName -notmatch $SkipDirRx -and $_.Name -notmatch $SkipExeRx -and $_.Length -gt 30000 }
    }

    $cands = @()
    foreach ($g in ($exes | Group-Object { Split-Path -Parent $_.FullName })) {
        $dir = $g.Name
        $scored = foreach ($f in $g.Group) {
            $vi = $f.VersionInfo; $score = 0
            $fn = (Split-Path -Leaf $dir) -replace '[^a-zA-Z0-9]', ''
            $en = [IO.Path]::GetFileNameWithoutExtension($f.Name) -replace '[^a-zA-Z0-9]', ''
            if ($fn -and $en -and ($fn -like "*$en*" -or $en -like "*$fn*")) { $score += 100 }
            if ($vi.CompanyName) { $score += 20 }
            if ($vi.FileDescription -and $vi.FileDescription -notmatch '(?i)uninstall|update|setup|installer|crash|helper|service') { $score += 15 }
            if ($vi.ProductName -and $vi.ProductName -notmatch '(?i)Windows') { $score += 15 }
            $score += [math]::Min(20, [int]($f.Length / 1MB))
            [pscustomobject]@{ File = $f; VI = $vi; Score = $score }
        }
        $best = $scored | Sort-Object Score -Descending | Select-Object -First 1
        if (-not $best) { continue }
        $cands += [pscustomobject]@{ Dir = $dir; Exe = $best.File.FullName; VI = $best.VI; Folder = (Split-Path -Leaf $dir) }
    }

    $kept = New-Object System.Collections.ArrayList
    foreach ($c in ($cands | Sort-Object { $_.Dir.Length })) {
        $nested = $false
        foreach ($k in $kept) { if ($c.Dir.StartsWith($k.Dir + '\', [StringComparison]::OrdinalIgnoreCase)) { $nested = $true; break } }
        if (-not $nested) { [void]$kept.Add($c) }
    }

    foreach ($c in $kept) {
        $cl = $c.Dir.ToLower()
        $covered = $false
        foreach ($kd in $knownDirs) {
            if ($cl -eq $kd -or $cl.StartsWith($kd + '\')) { $covered = $true; break }
        }
        if ($covered) { continue }

        $isGame = $c.Dir -match '(?i)\\torrent\\|\\Games\\|steamapps\\common|GOG Games'
        if ($isGame) {
            $name = $c.Folder; $minMB = $MinGameSizeMB; $minFiles = 10; $why = 'игра'
        } else {
            $name = $c.VI.ProductName
            if (-not $name -or $name -match '(?i)^(Microsoft® Windows® Operating System|Microsoft Windows|Windows)' -or $name -match '^[0-9\.]+$') { $name = $c.Folder }
            $minMB = $MinPortableSizeMB; $minFiles = 1; $why = 'portable'
        }
        if ($name -match '(?i)^(bin|bin64|Win64|windows-x86_64|core|tun2?|Scripts|nmake|vulkan|lld-bin|platform-tools|emulator|36\.0\.0|jbr)$') { continue }
        if (-not (Test-AppFolder $c.Dir $minMB $minFiles $name $why)) { continue }
        $st = Get-FolderStats $c.Dir
        $a = New-App -Name $name -Version "$($c.VI.ProductVersion)" -Publisher "$($c.VI.CompanyName)" -Path $c.Dir `
            -Source 'Portable' -Scope $(if ($c.Dir -like 'C:\Program Files*') { 'Система' } else { 'Пользователь' }) -Kind 'Портативно'
        $a.SizeMB = $st.SizeMB; $a.Files = $st.Files
        [void]$out.Add($a)
    }
    return $out
}

# ---------------------------------------------------------------- 8. shortcuts
function Get-ShortcutApps($known) {
    Write-Step 'Ярлыки «Пуск» и рабочий стол...'
    $out = New-Object System.Collections.ArrayList
    try { $shell = New-Object -ComObject WScript.Shell } catch { return $out }

    # куда уже указывают найденные приложения — чтобы не плодить дубли
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
            if ($t -match '(?i)^C:\\Program Files( \(x86\))?\\|^C:\\Users\\[^\\]+\\AppData\\Local\\Programs\\') { return }
            if ($seen.ContainsKey($t.ToLower())) { return }
            $seen[$t.ToLower()] = 1

            $dirT = Split-Path -Parent $t
            $covered = $false
            foreach ($kd in $knownDirs) {
                if ($dirT.ToLower() -eq $kd -or $dirT.ToLower().StartsWith($kd + '\') -or $kd.StartsWith($dirT.ToLower() + '\')) { $covered = $true; break }
            }
            if ($covered) { return }

            $st = Get-FolderStats $dirT
            if ($st.SizeMB -lt $MinPortableSizeMB) {
                Add-Skipped $_.BaseName $dirT $st.SizeMB $st.Files 'ярлык: папка слишком маленькая'
                return
            }
            $a = New-App -Name $_.BaseName -Version '' -Publisher '' -Path $t -Source 'Ярлык' -Scope 'Пользователь' -Kind 'Портативно'
            $a.SizeMB = $st.SizeMB; $a.Files = $st.Files
            [void]$out.Add($a)
        }
    }
    return $out
}

# ---------------------------------------------------------------- 9. merge
function Merge-Apps($apps) {
    $prio = @{ 'Steam' = 0; 'Epic' = 0; 'GOG' = 0; 'Реестр' = 1; 'Toolchain' = 1; 'Microsoft Store' = 2; 'Scoop' = 3; 'Portable' = 4; 'Ярлык' = 5 }
    $sorted = $apps | Sort-Object @{ Expression = { if ($prio.ContainsKey($_.Source)) { $prio[$_.Source] } else { 9 } } }

    $result = New-Object System.Collections.ArrayList
    $byName = @{}
    $byPath = @{}

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

        [void]$result.Add($a)
        if ($n) { $byName[$n] = $a }
        if ($p) { $byPath[$p] = $a }
    }
    return $result
}

# ---------------------------------------------------------------- 10. validate
function Apply-Overrides($apps, [string]$file) {
    if (-not $file -or -not (Test-Path $file)) { return $apps }
    Write-Step "Правки из $([IO.Path]::GetFileName($file))..."
    $ov = Import-PowerShellDataFile $file

    $map = @{}
    if ($ov.Names) {
        foreach ($k in $ov.Names.Keys) { $map[(Get-Norm $k)] = $ov.Names[$k] }
    }

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
            $a = New-App -Name $e.Name -Version "$($e.Version)" -Publisher "$($e.Publisher)" -Path $e.Path `
                -Source 'Override' -Scope 'Пользователь' -Kind 'Портативно'
            [void]$apps.Add($a)
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

# ---------------------------------------------------------------- 11. audit
function Invoke-Audit($apps, [string]$file) {
    Write-Step 'Аудит покрытия папок...'
    $paths = @()
    foreach ($a in $apps) {
        if (-not $a.Path) { continue }
        $p = $a.Path.ToLower()
        if ($p -match '\\[^\\]+\.[a-z0-9]+$') { $p = (Split-Path -Parent $a.Path).ToLower() }
        $paths += $p
    }

    $skipRx = '(?i)\\(Windows|WinSxS|Installer|Temp|temp|Cache|cache|node_modules|\.git|locales|resources|plugins|Extensions|packages|Packages|logs|Logs|Common Files|WindowsApps|Windows Defender|Windows Defender Advanced Threat Protection|Windows Mail|Windows Media Player|Windows Photo Viewer|Internet Explorer|WSL|Razer|AMD|Microsoft Office|Microsoft Office 15|Microsoft\.NET|EdgeCore|EdgeUpdate|CNext|WindowsPowerShell|jbr|jre|jdk|bin|bin64|lib|lib64|include|drivers|tcl|Scripts|__pycache__|\.venv|venv|Projects|FinalWire)(\\|$)'
    $lines = @()
    foreach ($root in (Get-PortableRoots)) {
        Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $p = $_.FullName.ToLower()
            if ($p -match $skipRx) { return }
            $hit = $false
            foreach ($x in $paths) {
                if ($p -eq $x -or $p.StartsWith($x + '\') -or $x.StartsWith($p + '\')) { $hit = $true; break }
            }
            if ($hit) { return }
            $exe = Get-ChildItem -LiteralPath $_.FullName -Filter '*.exe' -File -Recurse -Depth 4 -ErrorAction SilentlyContinue | Select-Object -First 1
            if (-not $exe) { return }
            $st = Get-FolderStats $_.FullName
            if ($st.SizeMB -lt $MinPortableSizeMB) { return }
            $lines += ('{0,9:N1} МБ  {1}' -f $st.SizeMB, $_.FullName)
        }
    }
    $lines = $lines | Sort-Object -Unique
    $lines | Out-File -FilePath $file -Encoding UTF8
    return $lines
}

# ---------------------------------------------------------------- 12. report
function Esc([string]$s) { if (-not $s) { return '' }; return [System.Net.WebUtility]::HtmlEncode($s) }

function Format-Size($mb) {
    if ($null -eq $mb -or $mb -le 0) { return '—' }
    if ($mb -ge 1024) { return ('{0:N1} ГБ' -f ($mb / 1024)) }
    return ('{0:N0} МБ' -f $mb)
}

function Export-Report($apps, $skipped) {
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $catOrder = @{ 'Приложение' = 0; 'Инструмент' = 1; 'Игра' = 2; 'Store' = 3; 'Системное' = 4 }
    $all = $apps | Sort-Object @{ Expression = { if ($catOrder.ContainsKey($_.Category)) { $catOrder[$_.Category] } else { 9 } } }, Name

    $all | Select-Object Category, Name, Version, Publisher, SizeMB, Files, Path, Source, Scope, Kind, Exists, Uninstall |
        Export-Csv -Path (Join-Path $OutputDir 'software-inventory.csv') -NoTypeInformation -Encoding UTF8

    $skipped | Sort-Object Name | Export-Csv -Path (Join-Path $OutputDir 'skipped.csv') -NoTypeInformation -Encoding UTF8

    $all | Select-Object Category, Name, Version, Publisher, SizeMB, Files, Path, Source, Scope, Kind, Exists |
        ConvertTo-Json -Depth 3 | Out-File -FilePath (Join-Path $OutputDir 'software-inventory.json') -Encoding UTF8

    # ---- markdown
    $md = New-Object System.Text.StringBuilder
    [void]$md.AppendLine('# Установленное ПО')
    [void]$md.AppendLine('')
    [void]$md.AppendLine("Сформировано: $stamp · всего записей: **$($all.Count)** · отброшено: **$($skipped.Count)**")
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

    # ---- html
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
.path .gone{color:#ff8f8f}
.tag{display:inline-block;padding:2px 8px;border-radius:5px;font-size:11px;font-weight:600;white-space:nowrap}
.t-app{background:#1d3a2a;color:#7ee2a8}
.t-tool{background:#2a2f3a;color:#9fb3d1}
.t-game{background:#3a1e2f;color:#ff8fc7}
.t-store{background:#1e2f4a;color:#7fb3ff}
.t-sys{background:#3a2a1e;color:#ffb87f}
.empty{padding:40px;text-align:center;color:var(--dim)}
@media(max-width:900px){header,main{padding-left:14px;padding-right:14px}th:nth-child(3),td:nth-child(3){display:none}}
'@)
    [void]$sb.AppendLine('</style></head><body>')
    [void]$sb.AppendLine('<header><h1>Установленное ПО</h1>')
    $hostname = $env:COMPUTERNAME
    $os = (Get-CimInstance Win32_OperatingSystem).Caption
    [void]$sb.AppendLine("<div class=`"sub`">$hostname · $os · $stamp · показано: <b id=`"total`">$($all.Count)</b> из $($all.Count) · отброшено: $($skipped.Count)</div>")
    [void]$sb.AppendLine('<div class="controls"><input type="search" id="q" placeholder="Поиск: название, издатель, путь...">')
    foreach ($cat in @('Приложение', 'Инструмент', 'Игра', 'Store', 'Системное')) {
        $cnt = @($all | Where-Object Category -eq $cat).Count
        if ($cnt -gt 0) {
            [void]$sb.AppendLine("<div class=`"chip on`" data-c=`"$cat`">$cat<b>$cnt</b></div>")
        }
    }
    [void]$sb.AppendLine("<label class=`"sw`"><input type=`"checkbox`" id=`"hideSys`" $(if (-not $IncludeSystem) { 'checked' })> скрыть служебное</label>")
    [void]$sb.AppendLine("<label class=`"sw`"><input type=`"checkbox`" id=`"hideMs`"> скрыть Microsoft</label>")
    [void]$sb.AppendLine('</div></header><main><table id="t"><thead><tr><th class="n">#</th><th data-s="name">Название</th><th data-s="ver">Версия</th><th data-s="pub">Издатель</th><th data-s="size">Размер</th><th data-s="path">Путь</th><th data-s="cat">Категория</th></tr></thead><tbody>')

    $i = 0
    foreach ($a in $all) {
        $i++
        $cls = switch ($a.Category) { 'Приложение' { 't-app' } 'Инструмент' { 't-tool' } 'Игра' { 't-game' } 'Store' { 't-store' } 'Системное' { 't-sys' } }
        $p = if ($a.Path) { $a.Path } else { '—' }
        $v = if ($a.Version) { $a.Version } else { '—' }
        $pub = if ($a.Publisher) { $a.Publisher } else { '—' }
        $pathCell = Esc $p
        if ($a.Exists -eq 'нет') { $pathCell = "$pathCell <span class=`"gone`">(папки нет)</span>" }
        $sys = if ($a.Category -eq 'Системное') { '1' } else { '0' }
        $ms = if ($a.Publisher -match '(?i)Microsoft|Майкрософт') { '1' } else { '0' }
        [void]$sb.AppendLine("<tr data-c=`"$($a.Category)`" data-sys=`"$sys`" data-ms=`"$ms`"><td class=`"n`">$i</td><td class=`"name`" data-v=`"$(Esc $a.Name)`">$(Esc $a.Name)</td><td class=`"ver`" data-v=`"$(Esc $v)`">$(Esc $v)</td><td class=`"pub`" data-v=`"$(Esc $pub)`">$(Esc $pub)</td><td class=`"size`" data-v=`"$($a.SizeMB)`">$(Format-Size $a.SizeMB)</td><td class=`"path`" data-v=`"$(Esc $p)`">$pathCell</td><td><span class=`"tag $cls`">$($a.Category)</span></td></tr>")
    }
    [void]$sb.AppendLine('</tbody></table><div class="empty" id="none" style="display:none">Ничего не найдено</div></main>')
    [void]$sb.AppendLine(@'
<script>
const q=document.getElementById('q'),rows=[...document.querySelectorAll('#t tbody tr')],none=document.getElementById('none');
const chips=[...document.querySelectorAll('.chip')],hideSys=document.getElementById('hideSys'),hideMs=document.getElementById('hideMs');
let active=new Set(chips.filter(c=>c.classList.contains('on')).map(c=>c.dataset.c));
function apply(){const s=q.value.toLowerCase().trim();let n=0;
 rows.forEach(r=>{const okC=active.has(r.dataset.c);const okS=!s||r.textContent.toLowerCase().includes(s);
  const okY=!(hideSys.checked&&r.dataset.sys==='1');const okM=!(hideMs&&hideMs.checked&&r.dataset.ms==='1');
  const show=okC&&okS&&okY&&okM;r.style.display=show?'':'none';if(show)n++;});
 none.style.display=n?'none':'block';document.getElementById('total').textContent=n;}
q.addEventListener('input',apply);hideSys.addEventListener('change',apply);if(hideMs)hideMs.addEventListener('change',apply);
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
if (-not $NoPortable) { foreach ($a in (Get-PortableApps $apps)) { [void]$apps.Add($a) } }
if (-not $SkipShortcuts) { foreach ($a in (Get-ShortcutApps $apps)) { [void]$apps.Add($a) } }

Write-Step 'Склейка дублей...'
$apps = Merge-Apps $apps

if (-not $OverridesPath) { $OverridesPath = Join-Path $PSScriptRoot 'overrides.psd1' }
$apps = Apply-Overrides $apps $OverridesPath
$apps = Test-Apps $apps

$html = Export-Report $apps $script:Skipped

if ($Audit) {
    $auditLines = Invoke-Audit $apps (Join-Path $OutputDir 'audit.txt')
    Write-Host ''
    if ($auditLines.Count) {
        Write-Host "Не покрыто отчётом (проверь глазами): $($auditLines.Count)" -ForegroundColor Yellow
        $auditLines | Select-Object -First 25 | ForEach-Object { Write-Host "  $_" }
        Write-Host "  ...полный список: $OutputDir\audit.txt" -ForegroundColor DarkGray
    }
    else {
        Write-Host 'Аудит: все найденные папки покрыты отчётом.' -ForegroundColor Green
    }
}

Write-Host ''
Write-Host 'Готово.' -ForegroundColor Green
foreach ($g in ($apps | Group-Object Category | Sort-Object Name)) { Write-Host ("  {0,-12} {1}" -f $g.Name, $g.Count) }
Write-Host ("  Всего:       {0}" -f $apps.Count)
if ($script:Skipped.Count) { Write-Host ("  Отброшено:   {0} (см. skipped.csv)" -f $script:Skipped.Count) -ForegroundColor DarkYellow }
Write-Host ''
Write-Host "Отчёты: $OutputDir" -ForegroundColor Cyan
Write-Host '  software-inventory.html' -ForegroundColor Cyan

if ($Open) { Start-Process $html }
