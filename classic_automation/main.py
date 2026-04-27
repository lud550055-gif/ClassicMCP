#!/usr/bin/env python3
# ============================================================
#  main.py — точка входа для агента Claude в VS Code
#
#  Использование:
#    python main.py                        # обработать все .doc в TASKS_DIR
#    python main.py LP_v17.doc             # один файл
#    python main.py --list                 # показать доступные варианты
#    python main.py --check-deps           # проверить зависимости
# ============================================================
import sys
import os
import time
import argparse
import traceback
from pathlib import Path

import config


def check_dependencies() -> bool:
    """Проверяет все зависимости перед запуском."""
    ok = True
    print("=== Проверка зависимостей ===")

    # Python пакеты
    packages = {
        "pyautogui":    "pip install pyautogui",
        "pygetwindow":  "pip install pygetwindow",
        "PIL":          "pip install Pillow",
        "psutil":       "pip install psutil",
        "sympy":        "pip install sympy",
        "docx":         "pip install python-docx",
    }
    for pkg, install_cmd in packages.items():
        try:
            __import__(pkg)
            print(f"  ✓ {pkg}")
        except ImportError:
            print(f"  ✗ {pkg}  →  {install_cmd}")
            ok = False

    # CLASSiC
    exe = Path(config.CLASSIC_EXE)
    if exe.exists():
        print(f"  ✓ CLASSiC: {exe}")
    else:
        print(f"  ✗ CLASSiC не найден: {exe}")
        print(f"    Проверьте CLASSIC_EXE в config.py")
        ok = False

    # Папки
    for name, path in [("TASKS_DIR", config.TASKS_DIR),
                        ("REPORTS_DIR", config.REPORTS_DIR),
                        ("MDL_DIR", config.MDL_DIR)]:
        p = Path(path)
        if not p.exists():
            p.mkdir(parents=True, exist_ok=True)
            print(f"  ✓ {name}: создана папка {p}")
        else:
            print(f"  ✓ {name}: {p}")

    print()
    return ok


def process_file(doc_path: str) -> bool:
    """
    Полный цикл обработки одного файла задания.
    Возвращает True при успехе.
    """
    from parser     import parse_variant
    from calculator import calculate, format_results
    from mdl_writer import write_mdl
    from classic_gui import ClassicController
    from report     import fill_report

    doc_path = str(Path(doc_path).resolve())
    print(f"\n{'='*60}")
    print(f" Обработка: {Path(doc_path).name}")
    print(f"{'='*60}")

    try:
        # 1. Парсинг варианта
        print("\n[1/5] Парсинг задания...")
        params = parse_variant(doc_path)
        print(f"  Вариант {params.variant}: K1={params.K1}, "
              f"K3={params.K3}, T3={params.T3}, "
              f"K4={params.K4}, T4={params.T4}, K5={params.K5}")

        # 2. Теоретические расчёты
        print("\n[2/5] Расчёты...")
        calc = calculate(params)
        print(format_results(calc))

        # 3. Генерация MDL
        print("\n[3/5] Генерация MDL файла...")
        mdl_path = str(Path(config.MDL_DIR) / f"var{params.variant:02d}.mdl")
        write_mdl(params, mdl_path)
        print(f"  MDL сохранён: {mdl_path}")

        # 4. CLASSiC — снимаем скриншоты
        print("\n[4/5] Запуск CLASSiC и съёмка скриншотов...")
        shots_dir = str(Path(config.REPORTS_DIR) / f"v{params.variant:02d}_shots")
        ctrl = ClassicController(
            mdl_path=mdl_path,
            output_dir=shots_dir,
            variant=params.variant,
            K1_critical=calc.K1_critical,
        )
        shots = ctrl.run_all()
        print(f"  Скриншоты в: {shots_dir}")

        # 5. Заполнение отчёта
        print("\n[5/5] Заполнение отчёта...")
        report_path = fill_report(
            template_path=doc_path,
            output_dir=config.REPORTS_DIR,
            calc=calc,
            shots=shots,
        )
        print(f"\n✓ Отчёт готов: {report_path}")
        return True

    except Exception as e:
        print(f"\n✗ Ошибка при обработке {Path(doc_path).name}:")
        traceback.print_exc()
        return False


def process_all():
    """Обрабатывает все .doc/.docx файлы из TASKS_DIR."""
    tasks_dir = Path(config.TASKS_DIR)
    files = sorted(list(tasks_dir.glob("*.doc")) + list(tasks_dir.glob("*.docx")))

    if not files:
        print(f"Нет файлов заданий в {tasks_dir}")
        return

    print(f"Найдено {len(files)} файл(ов) заданий:")
    for f in files:
        print(f"  {f.name}")

    success, failed = 0, 0
    for f in files:
        if process_file(str(f)):
            success += 1
        else:
            failed += 1
        # Небольшая пауза между файлами
        time.sleep(2)

    print(f"\n{'='*60}")
    print(f" Итог: {success} успешно, {failed} ошибок")
    print(f" Отчёты в: {config.REPORTS_DIR}")
    print(f"{'='*60}")


def main():
    parser = argparse.ArgumentParser(
        description="Автоматизация лабораторных работ CLASSiC"
    )
    parser.add_argument(
        "file", nargs="?",
        help="Путь к файлу задания (если не указан — обрабатываются все из TASKS_DIR)"
    )
    parser.add_argument(
        "--check-deps", action="store_true",
        help="Проверить зависимости и выйти"
    )
    parser.add_argument(
        "--list", action="store_true",
        help="Показать список файлов в TASKS_DIR"
    )
    parser.add_argument(
        "--calc-only", action="store_true",
        help="Только расчёты, без запуска CLASSiC"
    )
    args = parser.parse_args()

    if args.check_deps:
        ok = check_dependencies()
        sys.exit(0 if ok else 1)

    if args.list:
        tasks_dir = Path(config.TASKS_DIR)
        files = sorted(list(tasks_dir.glob("*.doc")) + list(tasks_dir.glob("*.docx")))
        print(f"Файлы заданий в {tasks_dir}:")
        for f in files:
            print(f"  {f.name}")
        return

    if args.calc_only:
        # Только расчёты без GUI — для отладки
        from parser import parse_variant
        from calculator import calculate, format_results
        target = args.file or str(Path(config.TASKS_DIR) / "*.doc")
        files = [args.file] if args.file else \
                sorted(Path(config.TASKS_DIR).glob("*.doc"))
        for f in files:
            try:
                params = parse_variant(str(f))
                calc = calculate(params)
                print(format_results(calc))
            except Exception as e:
                print(f"Ошибка: {e}")
        return

    if args.file:
        process_file(args.file)
    else:
        process_all()


if __name__ == "__main__":
    main()
