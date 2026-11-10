@{
    # Скопируй в overrides.psd1 и поправь под свою машину.
    # Ключ — имя программы в отчёте (регистр и пробелы не важны).
    Names = @{
        'Python 3.11.9' = @{ Path = 'C:\Users\me\AppData\Local\Programs\Python\Python311' }
        'Node.js'       = @{ Path = 'C:\Program Files\nodejs' }
        # 'Старое имя'  = @{ Name = 'Новое имя'; Category = 'Инструмент' }
    }

    # Программы, которые не находятся автоматически.
    Extra = @(
        # @{ Name = 'Мой тул'; Path = 'D:\tools\mytool'; Version = '1.0'; Publisher = '' }
    )
}
