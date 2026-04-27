# CLASSiC Automation — контекст проекта

## Что это
Автоматизация лабораторных работ по ТАУ. Скрипты управляют программой CLASSiC 3.2
через pyautogui, делают скриншоты и заполняют отчёты .docx.

## Структура
- config.py       — пути и таймауты
- parser.py       — читает номер варианта из .doc, параметры из VARIANT_TABLE
- calculator.py   — расчёты через SymPy (WP, Ф, Фe, Ккр, eуст)
- mdl_writer.py   — генерирует .mdl файл для CLASSiC
- classic_gui.py  — управление CLASSiC через pyautogui
- report.py       — вставляет скриншоты и расчёты в .docx шаблон
- main.py         — точка входа

## Следующий шаг
Нужно переделать в MCP сервер чтобы Claude Code мог вызывать
отдельные инструменты адаптивно, а не запускать весь пайплайн сразу.
Пример: агент читает .doc задание и сам решает какие инструменты вызвать.

## Пути на этом ПК
CLASSIC_EXE = C:\Users\lud50\Desktop\TU\classic\CLASSiC32.exe
TASKS_DIR   = C:\Users\lud50\Desktop\TU\tasks
REPORTS_DIR = C:\Users\lud50\Desktop\TU\reports