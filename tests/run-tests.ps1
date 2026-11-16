#Requires -Version 5.1
<#
    Тесты без внешних зависимостей. Запуск:
        powershell -NoProfile -ExecutionPolicy Bypass -File tests\run-tests.ps1
    Код выхода: 0 — всё зелёное, 1 — есть падения.
#>
param(
    [switch]$IncludeSmoke
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'lib\Core.ps1')

$script:Passed = 0
$script:Failed = 0
$script:Failures = @()

function Assert-Equal($expected, $actual, [string]$name) {
    if ("$expected" -ceq "$actual") { $script:Passed++; return }
    $script:Failed++
    $script:Failures += "$name`n      ожидалось: [$expected]`n      получено:  [$actual]"
}

function Assert-True($condition, [string]$name) {
    if ($condition) { $script:Passed++; return }
    $script:Failed++
    $script:Failures += "$name — условие ложно"
}

function Assert-Contains([string]$haystack, [string]$needle, [string]$name) {
    if ($haystack -like "*$needle*") { $script:Passed++; return }
    $script:Failed++
    $script:Failures += "$name — нет '$needle' в '$haystack'"
}

# ---------------------------------------------------------------- Get-Norm
Assert-Equal 'microsoftvisualc' (Get-Norm 'Microsoft Visual C++') 'Norm: пунктуация и регистр'
Assert-Equal '' (Get-Norm '') 'Norm: пустая строка'
Assert-Equal 'python3119' (Get-Norm 'Python 3.11.9') 'Norm: версия'

# ---------------------------------------------------------------- Get-CleanName
Assert-Equal "Sid Meier's Civilization VI" (Get-CleanName ('Sid Meier' + [char]0x2019 + 's Civilization VI [RePack]')) 'Clean: RePack и кривой апостроф'
Assert-Equal 'Cursor' (Get-CleanName 'Cursor (User)') 'Clean: (User)'
Assert-Equal 'TickTick' (Get-CleanName 'TickTick version 8.2.5.0') 'Clean: version'
Assert-Equal 'Quattro VPN' (Get-CleanName 'Quattro VPN, версия 0.28.3+53') 'Clean: версия'
Assert-Equal 'Roblox Player' (Get-CleanName 'Roblox Player for bychk') 'Clean: for user'
Assert-Equal 'Foo' (Get-CleanName 'Foo (x64)') 'Clean: (x64)'
Assert-Equal '' (Get-CleanName '') 'Clean: пустая строка'

# ---------------------------------------------------------------- Get-Category
Assert-Equal 'Игра' (Get-Category 'Dota 2' 'D:\steam\steamapps\common\dota 2 beta' 'Steam' 'Игра') 'Category: steam'
Assert-Equal 'Игра' (Get-Category 'Tiny Eden' 'D:\torrent\Tiny Eden (2026)' 'Диск' 'Портативно') 'Category: torrent'
Assert-Equal 'Системное' (Get-Category 'Microsoft Visual C++ 2015' '' 'Реестр' 'Установлено') 'Category: redist'
Assert-Equal 'Системное' (Get-Category 'Microsoft.DesktopAppInstaller' '' 'Microsoft Store' 'Store') 'Category: store system'
Assert-Equal 'Инструмент' (Get-Category 'Android Studio' 'D:\apps\AnSt' 'Реестр' 'Установлено') 'Category: tool'
Assert-Equal 'Инструмент' (Get-Category 'Ubuntu' '\\wsl$\Ubuntu' 'WSL' 'Портативно') 'Category: wsl'
Assert-Equal 'Приложение' (Get-Category 'Some App' 'D:\apps\some' 'Реестр' 'Установлено') 'Category: app'
Assert-Equal 'Store' (Get-Category 'DuckDuckGo' 'C:\Program Files\WindowsApps\DuckDuckGo' 'Microsoft Store' 'Store') 'Category: store app'

# ---------------------------------------------------------------- числа и размеры
$old = [System.Globalization.CultureInfo]::CurrentCulture
try {
    [System.Globalization.CultureInfo]::CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('ru-RU')
    Assert-Equal '305.6' (Format-Inv 305.6) 'Format-Inv: точка при ru-RU'
    Assert-Equal '' (Format-Inv $null) 'Format-Inv: null'
    Assert-Contains (Format-Size 512) 'МБ' 'Format-Size: мегабайты'
    Assert-Contains (Format-Size 2048) 'ГБ' 'Format-Size: гигабайты'
    Assert-Equal '—' (Format-Size 0) 'Format-Size: ноль'
}
finally {
    [System.Globalization.CultureInfo]::CurrentCulture = $old
}

# ---------------------------------------------------------------- Rot13
Assert-Equal 'P:\Grfg' (Rot13 'C:\Test') 'Rot13: строка'
Assert-Equal 'C:\Test' (Rot13 (Rot13 'C:\Test')) 'Rot13: обратимость'

# ---------------------------------------------------------------- пути
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('isi-test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
try {
    $dir = Join-Path $tmp 'App'
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    $exe = Join-Path $dir 'app.exe'
    Set-Content -Path $exe -Value 'x'
    $verDir = Join-Path $tmp 'zapret-1.9.7b'
    New-Item -ItemType Directory -Path $verDir -Force | Out-Null

    Assert-Equal $dir.ToLower() (Get-ContainerPath $dir) 'ContainerPath: папка'
    Assert-Equal $dir.ToLower() (Get-ContainerPath $exe) 'ContainerPath: файл'
    Assert-Equal $verDir.ToLower() (Get-ContainerPath $verDir) 'ContainerPath: папка с версией в имени'
}
finally {
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------- New-App
$a = New-App -Name 'Foo' -Version '1.0' -Publisher 'Bar' -Path 'C:/Foo/Bar/' -Source 'Реестр'
Assert-Equal 'C:\Foo\Bar' $a.Path 'New-App: нормализация слэшей'
Assert-Equal 'Foo' $a.Name 'New-App: имя'
Assert-Equal 1 $a.Copies 'New-App: Copies'

# ---------------------------------------------------------------- Merge-Apps
$list = @(
    (New-App -Name 'Foo' -Path 'D:\apps\foo' -Source 'Реестр'),
    (New-App -Name 'Foo' -Path 'D:\apps\foo' -Source 'Диск'),
    (New-App -Name 'Foo' -Path 'D:\other\foo' -Source 'Диск'),
    (New-App -Name 'Bar' -Path 'D:\apps\bar' -Source 'Диск')
)
$merged = Merge-Apps $list
Assert-Equal 3 (@($merged).Count) 'Merge: одинаковый путь схлопнут'

# ---------------------------------------------------------------- Group-Copies
$g = @(
    [pscustomobject]@{ Name = 'SHADE'; Copies = 1; OtherPaths = ''; SizeMB = 300; LastRun = $null; Path = 'D:\a' },
    [pscustomobject]@{ Name = 'SHADE (копия 1)'; Copies = 1; OtherPaths = ''; SizeMB = 200; LastRun = $null; Path = 'D:\b' }
)
$grouped = @(Group-Copies $g)
Assert-Equal 1 $grouped.Count 'Group-Copies: одна строка'
Assert-Equal 2 $grouped[0].Copies 'Group-Copies: счётчик'
Assert-Equal 'D:\b' $grouped[0].OtherPaths 'Group-Copies: второй путь'
Assert-Equal 500 $grouped[0].SizeMB 'Group-Copies: сумма размеров'

# ---------------------------------------------------------------- отчёт: HTML-эскейп
Assert-Equal '&lt;b&gt;&amp;' (Esc '<b>&') 'Esc: html'

# ---------------------------------------------------------------- smoke
if ($IncludeSmoke) {
    $out = Join-Path $env:TEMP ('isi-smoke-' + [guid]::NewGuid().ToString('N'))
    $script = Join-Path $root 'Get-InstalledSoftware.ps1'
    & powershell -NoProfile -ExecutionPolicy Bypass -File $script -NoStore -NoPortable -OutputDir $out -Quiet | Out-Null
    Assert-Equal 0 $LASTEXITCODE 'Smoke: код выхода'
    foreach ($f in @('software-inventory.csv', 'software-inventory.html', 'excluded.csv')) {
        Assert-True (Test-Path (Join-Path $out $f)) "Smoke: есть $f"
    }
    $csv = Import-Csv (Join-Path $out 'software-inventory.csv')
    Assert-True ($csv.Count -gt 0) 'Smoke: отчёт не пустой'
    Assert-True ("$($csv[0].SizeMB)" -notmatch ',') 'Smoke: точка в SizeMB'
    Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------- итог
Write-Host ''
if ($script:Failed) {
    Write-Host "Провалено: $($script:Failed), пройдено: $($script:Passed)" -ForegroundColor Red
    foreach ($f in $script:Failures) { Write-Host "  - $f" -ForegroundColor Red }
    exit 1
}
Write-Host "Все тесты пройдены: $($script:Passed)" -ForegroundColor Green
exit 0
