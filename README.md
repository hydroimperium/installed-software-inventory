# installed-software-inventory

Скрипт для Windows, который собирает полный список установленного ПО: обычные программы, portable, Microsoft Store, Steam/Epic/GOG, Scoop, WSL и SDK. На выходе — интерактивный HTML-отчёт с поиском и фильтрами, CSV, Markdown и JSON.

Главное отличие от «просто перечислить реестр»: скрипт проверяет, что папка реально существует и не пустая. Огрызки игр на 0 МБ и папки-пустышки отбрасываются и попадают в отдельный `skipped.csv` с причиной.

## Требования

- Windows 10/11
- PowerShell 5.1 (встроенный) или PowerShell 7
- Права пользователя достаточно; часть путей может быть недоступна без админа

## Быстрый старт

```powershell
git clone https://github.com/<you>/installed-software-inventory.git
cd installed-software-inventory
.\Get-InstalledSoftware.ps1 -Open
```

Или двойным кликом: `run.cmd`.

Если PowerShell блокирует запуск:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Get-InstalledSoftware.ps1
```

## Параметры

| Параметр | По умолчанию | Что делает |
|---|---|---|
| `-OutputDir` | `.\reports` | Куда писать отчёты |
| `-MinGameSizeMB` | `100` | Минимальный размер папки игры |
| `-MinPortableSizeMB` | `3` | Минимальный размер portable-программы |
| `-IncludeSystem` | выкл. | Показать служебное ПО сразу, без фильтра |
| `-NoStore` | — | Не собирать Microsoft Store |
| `-NoGames` | — | Не собирать игры |
| `-NoPortable` | — | Не сканировать диск на portable |
| `-ExtraRoots` | — | Дополнительные папки для поиска |
| `-SkipShortcuts` | — | Не разбирать ярлыки «Пуск» и рабочего стола |
| `-OverridesPath` | `.\overrides.psd1` | Файл ручных правок |
| `-Audit` | — | Проверить, что все папки с exe попали в отчёт |
| `-Open` | — | Открыть HTML после сборки |

Примеры:

```powershell
# только игры, порог 50 ГБ-папок
.\Get-InstalledSoftware.ps1 -NoStore -NoPortable -MinGameSizeMB 50

# добавить свой диск с софтом
.\Get-InstalledSoftware.ps1 -ExtraRoots 'E:\portable','F:\tools'

# проверить полноту и открыть отчёт
.\Get-InstalledSoftware.ps1 -Audit -Open
```

## Откуда берутся данные

1. **Реестр** — `HKLM`/`HKCU`, 32- и 64-битные ветки `Uninstall`, плюс `App Paths` для восстановления путей.
2. **Microsoft Store** — `Get-AppxPackage` без фреймворков и ресурсных пакетов.
3. **Steam** — все библиотеки из `libraryfolders.vdf`, манифесты и папки без манифеста.
4. **Epic Games** — манифесты лаунчера.
5. **GOG** — ветка реестра `GOG.com\Games`.
6. **Scoop** — `%USERPROFILE%\scoop\apps`.
7. **WSL** — установленные дистрибутивы.
8. **SDK/тулчейны** — Android SDK, Edge WebView2, Scoop, `zapret`.
9. **Файловая система** — поиск exe в стандартных местах и корнях дисков, с выбором главного exe в папке.
10. **Ярлыки** — «Пуск» и рабочий стол, как страховка от нестандартных мест установки.

## Валидация

Правила отбраковки:

- путь не существует → `путь не найден (остаток удаления)`;
- игра меньше `MinGameSizeMB` или в ней меньше 10 файлов → `игра: меньше ...`;
- portable меньше `MinPortableSizeMB` → `portable: меньше ...`;
- ярлык ведёт в папку меньше порога → `ярлык: папка слишком маленькая`.

Всё отброшенное — в `reports/skipped.csv`, чтобы можно было проверить руками.

## Фильтры в HTML

- поиск по названию, издателю и пути;
- чипы по категориям: Приложение, Инструмент, Игра, Store, Системное;
- галочка **«скрыть служебное»** — убирает Visual C++, .NET, SDK, драйверы и прочие рантаймы (включена по умолчанию);
- галочка **«скрыть Microsoft»** — убирает всё от Microsoft;
- сортировка по любому столбцу.

## Ручные правки

Если путь не определяется (например, установка без записи `InstallLocation`), создай `overrides.psd1` рядом со скриптом:

```powershell
@{
    Names = @{
        'Python 3.11.9' = @{ Path = 'C:\Users\me\AppData\Local\Programs\Python\Python311' }
        'Старое имя'    = @{ Name = 'Новое имя'; Category = 'Инструмент' }
    }
    Extra = @(
        @{ Name = 'Мой тул'; Path = 'D:\tools\mytool'; Version = '1.0' }
    )
}
```

Файл в `.gitignore` — личные пути не попадут в репозиторий. Шаблон: `overrides.example.psd1`.

## Выходные файлы

| Файл | Что внутри |
|---|---|
| `software-inventory.html` | Интерактивная таблица с фильтрами |
| `software-inventory.csv` | Для Excel (UTF-8) |
| `software-inventory.md` | Текстовая версия по категориям |
| `software-inventory.json` | Для скриптов |
| `skipped.csv` | Что и почему отброшено |
| `audit.txt` | Папки с exe, не попавшие в отчёт (при `-Audit`) |

## Ограничения

- Не видит portable-программы, у которых нет ни ярлыка, ни папки в стандартном месте.
- Не разбирает содержимое MSI-кэша и не считает «программы», установленные только для другого пользователя.
- Размер считается рекурсивно по папке и может занимать время на больших каталогах.
- Служебные компоненты Windows (VC++, .NET, UI.Xaml) в отчёт попадают, но скрыты фильтром — это сделано намеренно.

## Лицензия

MIT
