# CLASSiC Automation

Автоматизация лабораторных работ по ТАУ в программе CLASSiC 3.2.

## Установка

```bash
cd classic_automation
pip install -r requirements.txt
```

## Первый запуск — проверка

```bash
python main.py --check-deps
```

Должны стоять все галочки ✓. Если нет — установите недостающие пакеты.

## Настройка

Откройте `config.py` и проверьте пути:

```python
CLASSIC_EXE = r"C:\Users\lud50\Desktop\TU\classic\CLASSiC32.exe"
TASKS_DIR   = r"C:\Users\lud50\Desktop\TU\tasks"      # сюда кладём .doc задания
REPORTS_DIR = r"C:\Users\lud50\Desktop\TU\reports"    # сюда сохраняются отчёты
MDL_DIR     = r"C:\Users\lud50\Desktop\TU\models"     # временные .mdl файлы
```

## Добавление вариантов

В `parser.py` → `VARIANT_TABLE` добавьте строку:

```python
VARIANT_TABLE = {
    ...
    23: (120, 3.0, 0.6, 0.25, 0.06, 0.06),  # (K1, K3, T3, K4, T4, K5)
    ...
}
```

## Использование

### Один файл
```bash
python main.py LP_v17.doc
```

### Все файлы из папки tasks
```bash
python main.py
```

### Только расчёты (без CLASSiC, для проверки)
```bash
python main.py --calc-only LP_v17.doc
```

## Структура проекта

```
classic_automation/
├── main.py         ← точка входа
├── config.py       ← настройки путей и таймаутов
├── parser.py       ← извлекает номер варианта из .doc
├── calculator.py   ← все теоретические расчёты (SymPy)
├── mdl_writer.py   ← генерирует .mdl файл для CLASSiC
├── classic_gui.py  ← управление CLASSiC через pyautogui
├── report.py       ← заполняет .docx отчёт
└── requirements.txt
```

## Как это работает

1. `parser.py` читает `.doc` → находит номер варианта → берёт параметры из `VARIANT_TABLE`
2. `calculator.py` считает WP(s), Ф(s), Фe(s), Ккр, eуст через SymPy
3. `mdl_writer.py` создаёт `.mdl` файл с нужными параметрами
4. `classic_gui.py` запускает CLASSiC, загружает `.mdl`, делает скриншоты всех задач
5. `report.py` вставляет скриншоты и расчёты в шаблон `.docx`

## Настройка CLASSiC GUI (если что-то не кликается)

Если скрипт не попадает в нужные пункты меню, откройте `classic_gui.py`
и добавьте точные названия пунктов меню из вашей версии CLASSiC в словари `menu_map`.

Также можно увеличить таймауты в `config.py`:
```python
DIALOG_TIMEOUT = 5    # если CLASSiC медленно открывает диалоги
RENDER_TIMEOUT = 3    # если графики долго строятся
```

## Использование с Claude Code в VS Code

В `.vscode/tasks.json` добавьте:

```json
{
  "version": "2.0.0",
  "tasks": [
    {
      "label": "CLASSiC: обработать все задания",
      "type": "shell",
      "command": "python ${workspaceFolder}/classic_automation/main.py",
      "group": "build"
    },
    {
      "label": "CLASSiC: обработать текущий файл",
      "type": "shell",
      "command": "python ${workspaceFolder}/classic_automation/main.py ${file}",
      "group": "build"
    }
  ]
}
```

Или просто попросите агента Claude Code:
> "Запусти main.py для файла LP_v23.doc"
