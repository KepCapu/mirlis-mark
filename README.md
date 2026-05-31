# Mark - Система маркировки

Десктоп-приложение для пищевого производства: печать этикеток на товары и учёт
статистики. Работает на Windows, написано на Python + PyQt5.

Продукт выпускается под брендом **LDN Tech**. (C) 2026 LDN Tech.

## Требования
- Windows 10/11 (x64), Python 3.10+
- PyQt5 (НЕ PyQt6 - целевое оборудование Celeron J1900 без AVX)
- Термопринтер с TSPL (4B-2054L, 203 dpi) - только для реальной печати

## Быстрый старт (разработка)
1. git clone <URL>   (или Code -> Download ZIP)
2. py -m venv .venv ; .\.venv\Scripts\Activate.ps1
3. python -m pip install -r requirements.txt
4. python main.py

## Сборка .exe
py -m PyInstaller --noconfirm --clean --name Mark --icon=assets\mark_app.ico --version-file version_info.txt --onedir --noconsole --add-data "assets;assets" --add-data "data_sources;data_sources" main.py

Результат: dist\Mark\Mark.exe

## Установщик (Inno Setup)
Открыть installer.iss -> F9. Результат: installer\MarkSetup.exe

## Данные пользователя
%LOCALAPPDATA%\MirlisMark\ - настройки, архив этикеток (365 дней),
журнал статистики, кэш весов (scale_link.json).
Внутреннее имя папки MirlisMark не видно пользователю; менять не нужно.

## Ключевые файлы
- main.py - главное окно, UI, печать
- excel_loader.py - чтение Excel (листы продукт/изготовил/цех)
- label_logic.py - формирование этикетки
- printer.py - отправка TSPL на термопринтер
- scale_reader.py - чтение веса (автоопределение порта/скорости/протокола)
- statistics_*.py, stats_*.py - модуль статистики
- installer.iss - установщик; Mark.spec - сборка; version_info.txt - метаданные .exe

## Заметки для разработчиков
- Только PyQt5 (Qt.AlignCenter, без AlignmentFlag).
- Excel: три листа продукт/изготовил/цех, автообновление ~60 c.
- Печать через виртуальный холст фиксированного размера (стабильно на разных DPI).
- Весы: tEnSo (Тензо-М), 6.43, ASCII/MT-SICS; приём по CRC, кэш комбинации.
- Два режима интерфейса: ПК и планшет.
