# Регулярки классификации
$SystemRx = '(?i)Visual C\+\+|\.NET (SDK|Runtime|Desktop Runtime)|Windows Desktop Runtime|Windows SDK|DirectX|Edge WebView2|Edge Update|Epic Online Services|Mozilla Maintenance|AMD[ _](Chipset|Install)|Realtek Card Reader|EpsonNet|Драйвер принтера|Application Compatibility|Application Virtualization|Windows Subsystem for Linux|^WSL|WinAppRuntime|ShellExtension|ContextMenu|SparseApp|Winget\.Source|DesktopAppInstaller|StorePurchaseApp|LanguageExperiencePack|VCLibs|Microsoft\.UI\.Xaml|HostRuntime|AppRuntime|PlatformRuntime|Windows App Runtime|Windows Media Player|Windows Photo Viewer|Windows Defender|Microsoft Office$|Razer S3|Razer.*Wizard'
$ToolRx   = '(?i)Visual Studio Code|Android Studio|Android SDK|IntelliJ|PyCharm|WebStorm|^Git$|GitHub CLI|Node\.js|Python|OpenJDK|Temurin|JDK|Docker|Postman|Scoop|Maestro|Cursor|Devin|Antigravity|OpenCode|DeepSeek|Codex|Blender|Godot|NekoBox|LockHunter|WinRAR|7-Zip|PowerToys|zapret|Ubuntu|Bun$|Chocolatey|spicetify'
$GamePathRx = '(?i)steamapps\\common|\\Games\\|\\torrent\\|GOG Games'
# Pure helpers: no filesystem access, safe to unit-test.

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

function Format-Size($mb) {
    if ($null -eq $mb -or $mb -le 0) { return '—' }
    if ($mb -ge 1024) { return ('{0:N1} ГБ' -f ($mb / 1024)) }
    return ('{0:N0} МБ' -f $mb)
}

function Format-Inv($v) {
    if ($null -eq $v) { return '' }
    return ([double]$v).ToString([System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-ContainerPath([string]$path) {
    if (-not $path) { return '' }
    if (Test-Path -LiteralPath $path -PathType Leaf) { return (Split-Path -Parent $path).ToLower() }
    return $path.ToLower()
}

function Rot13([string]$s) {
    $chars = $s.ToCharArray()
    for ($i = 0; $i -lt $chars.Length; $i++) {
        $c = $chars[$i]
        if ($c -ge 'a' -and $c -le 'z') { $chars[$i] = [char]((([int]$c - 97 + 13) % 26) + 97) }
        elseif ($c -ge 'A' -and $c -le 'Z') { $chars[$i] = [char]((([int]$c - 65 + 13) % 26) + 65) }
    }
    return -join $chars
}

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
        Copies    = 1
        OtherPaths = ''
        Extra     = ''
    }
}

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

function Group-Copies($apps) {
    $byName = @{}
    $out = New-Object System.Collections.ArrayList
    foreach ($a in $apps) {
        $base = $a.Name -replace '\s*\(копия \d+\)\s*$', ''
        $n = Get-Norm $base
        if ($base -match '^ms-resource') { [void]$out.Add($a); continue }
        if (-not $n) { [void]$out.Add($a); continue }
        if (-not $byName.ContainsKey($n)) {
            $byName[$n] = $a
            [void]$out.Add($a)
            continue
        }
        $primary = $byName[$n]
        # копией считаем только тот же продукт: совпадает категория или источник
        $primary.Copies++
        if ($a.Path) {
            $primary.OtherPaths = if ($primary.OtherPaths) { "$($primary.OtherPaths) | $($a.Path)" } else { $a.Path }
        }
        if ($a.SizeMB -and $primary.SizeMB) { $primary.SizeMB = $primary.SizeMB + $a.SizeMB }
        elseif ($a.SizeMB) { $primary.SizeMB = $a.SizeMB }
        if ($a.LastRun -and (-not $primary.LastRun -or $a.LastRun -gt $primary.LastRun)) { $primary.LastRun = $a.LastRun }
    }
    return $out
}



function Esc([string]$s) { if (-not $s) { return '' }; return [System.Net.WebUtility]::HtmlEncode($s) }

