#!/usr/bin/env python3
"""
MCP-сервер для автоматизации лабораторных работ CLASSiC.
Каждый модуль проекта — отдельный инструмент.

Запуск напрямую:
    python server.py

Регистрация в Claude Code:
    claude mcp add classic-automation python C:/Users/lud50/Desktop/TU/classic_automation/server.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import config
from mcp.server.fastmcp import FastMCP

mcp = FastMCP(
    "classic-automation",
    instructions=(
        "Автоматизация лабораторных по ТАУ (CLASSiC 3.2). "
        "Типичный порядок вызовов: get_variant_params → calculate_transfer_functions "
        "→ write_mdl_file → run_classic_gui → fill_report_docx. "
        "Для полного пайплайна используй run_full_pipeline."
    ),
)


# ── 1. list_task_files ────────────────────────────────────────────────────────

@mcp.tool()
def list_task_files(tasks_dir: str = "") -> dict:
    """
    Перечисляет .doc/.docx файлы заданий.
    tasks_dir: папка с заданиями (по умолчанию TASKS_DIR из config.py).
    """
    p = Path(tasks_dir or config.TASKS_DIR)
    if not p.exists():
        return {"error": f"Папка не найдена: {p}", "files": []}
    files = sorted(
        [f.name for f in p.glob("*.doc")] +
        [f.name for f in p.glob("*.docx")]
    )
    return {"tasks_dir": str(p), "files": files, "count": len(files)}


# ── 2. parse_task_file ────────────────────────────────────────────────────────

@mcp.tool()
def parse_task_file(doc_path: str) -> dict:
    """
    Читает .doc/.docx файл задания и извлекает параметры варианта.
    doc_path: полный путь к файлу задания.
    Возвращает: variant, K1, K3, T3, K4, T4, K5.
    """
    from parser import parse_variant
    params = parse_variant(doc_path)
    return {
        "variant": params.variant,
        "K1": params.K1,
        "K3": params.K3,
        "T3": params.T3,
        "K4": params.K4,
        "T4": params.T4,
        "K5": params.K5,
    }


# ── 3. get_variant_params ─────────────────────────────────────────────────────

@mcp.tool()
def get_variant_params(variant: int) -> dict:
    """
    Возвращает параметры варианта напрямую из VARIANT_TABLE без чтения файла.
    variant: номер варианта (1–20).
    Возвращает: variant, K1, K3, T3, K4, T4, K5.
    """
    from parser import VARIANT_TABLE
    if variant not in VARIANT_TABLE:
        available = sorted(VARIANT_TABLE.keys())
        return {"error": f"Вариант {variant} не найден. Доступны: {available}"}
    K1, K3, T3, K4, T4, K5 = VARIANT_TABLE[variant]
    return {
        "variant": variant,
        "K1": K1, "K3": K3, "T3": T3,
        "K4": K4, "T4": T4, "K5": K5,
    }


# ── 4. calculate_transfer_functions ──────────────────────────────────────────

@mcp.tool()
def calculate_transfer_functions(
    variant: int,
    K1: float,
    K3: float,
    T3: float,
    K4: float,
    T4: float,
    K5: float,
) -> dict:
    """
    Вычисляет передаточные функции и анализирует устойчивость системы.
    Принимает параметры из parse_task_file или get_variant_params.

    Возвращает:
    - WP_str, Phi_str, PhiE_str — формулы передаточных функций
    - char_poly_str, char_coeffs — характеристический полином
    - hurwitz_stable, hurwitz_detail — критерий Гурвица
    - K1_critical, K_loop_critical — критические коэффициенты усиления
    - e_ust_step, e_ust_ramp — установившиеся ошибки
    - WP_classic, Phi_classic, PhiE_classic — коэффициенты для CLASSiC
    - summary — текстовый отчёт
    """
    from parser import VariantParams
    from calculator import calculate, format_results

    params = VariantParams(
        variant=variant,
        K1=K1, K3=K3, T3=T3,
        K4=K4, T4=T4, K5=K5,
    )
    r = calculate(params)
    return {
        "variant": r.variant,
        "WP_str": r.WP_str,
        "Phi_str": r.Phi_str,
        "PhiE_str": r.PhiE_str,
        "char_poly_str": r.char_poly_str,
        "char_coeffs": r.char_coeffs,
        "hurwitz_stable": r.hurwitz_stable,
        "hurwitz_detail": r.hurwitz_detail,
        "K1_critical": r.K1_critical,
        "K_loop_critical": r.K_loop_critical,
        "e_ust_step": r.e_ust_step,
        "e_ust_ramp": r.e_ust_ramp,
        "WP_classic": r.WP_classic,
        "Phi_classic": r.Phi_classic,
        "PhiE_classic": r.PhiE_classic,
        "K1": r.K1, "K3": r.K3, "T3": r.T3,
        "K4": r.K4, "T4": r.T4, "K5": r.K5,
        "summary": format_results(r),
    }


# ── 5. write_mdl_file ─────────────────────────────────────────────────────────

@mcp.tool()
def write_mdl_file(
    variant: int,
    K1: float,
    K3: float,
    T3: float,
    K4: float,
    T4: float,
    K5: float,
    output_path: str = "",
) -> dict:
    """
    Генерирует .mdl файл структурной схемы для CLASSiC 3.2.
    Принимает параметры из parse_task_file или get_variant_params.
    output_path: путь для сохранения (по умолчанию MDL_DIR/varXX.mdl).
    Возвращает: mdl_path — путь к созданному файлу.
    """
    from parser import VariantParams
    from mdl_writer import write_mdl

    params = VariantParams(
        variant=variant,
        K1=K1, K3=K3, T3=T3,
        K4=K4, T4=T4, K5=K5,
    )
    if not output_path:
        Path(config.MDL_DIR).mkdir(parents=True, exist_ok=True)
        output_path = str(Path(config.MDL_DIR) / f"var{variant:02d}.mdl")

    result_path = write_mdl(params, output_path)
    return {"mdl_path": result_path, "variant": variant}


# ── 6. run_classic_gui ────────────────────────────────────────────────────────

@mcp.tool()
def run_classic_gui(
    mdl_path: str,
    variant: int,
    k1_critical: float = 0.0,
    output_dir: str = "",
) -> dict:
    """
    Запускает CLASSiC 3.2, загружает .mdl модель и делает скриншоты всех задач.
    mdl_path: путь к .mdl файлу (из write_mdl_file).
    k1_critical: критическое усиление для задачи 11 (из calculate_transfer_functions).
    output_dir: папка для скриншотов (по умолчанию REPORTS_DIR/vXX_shots).

    Возвращает словарь screenshots с путями к PNG-файлам:
    schema, text_form, WP, Phi, PhiE, step_response, ramp_response, critical, bode.
    """
    from classic_gui import ClassicController

    if not output_dir:
        output_dir = str(Path(config.REPORTS_DIR) / f"v{variant:02d}_shots")

    ctrl = ClassicController(
        mdl_path=mdl_path,
        output_dir=output_dir,
        variant=variant,
        K1_critical=k1_critical,
    )
    shots = ctrl.run_all()
    return {
        "output_dir": output_dir,
        "screenshots": {
            "schema":          shots.schema,
            "text_form":       shots.text_form,
            "characteristics": shots.characteristics,
            "root_locus":      shots.root_locus,
            "step_response":   shots.step_response,
            "ramp_response":   shots.ramp_response,
            "bode":            shots.bode,
            "tf_panel":        shots.tf_panel,
            "critical":        shots.critical,
        },
    }


# ── 7. fill_report_docx ───────────────────────────────────────────────────────

@mcp.tool()
def fill_report_docx(
    template_path: str,
    calc_results: dict,
    screenshots: dict = None,
    output_dir: str = "",
) -> dict:
    """
    Заполняет шаблон .docx/.doc результатами расчётов и скриншотами.

    template_path: путь к исходному файлу задания (шаблону).
    calc_results: словарь из calculate_transfer_functions.
    screenshots: словарь screenshots из run_classic_gui (или пустой/None).
    output_dir: папка для готового отчёта (по умолчанию REPORTS_DIR).

    Возвращает: report_path — путь к готовому .docx отчёту.
    """
    from calculator import CalcResults
    from classic_gui import Screenshots
    from report import fill_report

    r = calc_results
    calc = CalcResults(
        variant=r.get("variant", 0),
        WP_str=r.get("WP_str", ""),
        Phi_str=r.get("Phi_str", ""),
        PhiE_str=r.get("PhiE_str", ""),
        char_poly_str=r.get("char_poly_str", ""),
        char_coeffs=r.get("char_coeffs", []),
        hurwitz_stable=r.get("hurwitz_stable", False),
        hurwitz_detail=r.get("hurwitz_detail", ""),
        K1_critical=r.get("K1_critical", 0.0),
        K_loop_critical=r.get("K_loop_critical", 0.0),
        e_ust_step=r.get("e_ust_step", ""),
        e_ust_ramp=r.get("e_ust_ramp", ""),
        WP_classic=r.get("WP_classic", {}),
        Phi_classic=r.get("Phi_classic", {}),
        PhiE_classic=r.get("PhiE_classic", {}),
        K1=r.get("K1", 0.0), K3=r.get("K3", 0.0), T3=r.get("T3", 0.0),
        K4=r.get("K4", 0.0), T4=r.get("T4", 0.0), K5=r.get("K5", 0.0),
    )

    s = screenshots or {}
    shots = Screenshots(
        schema=s.get("schema", ""),
        text_form=s.get("text_form", ""),
        characteristics=s.get("characteristics", ""),
        root_locus=s.get("root_locus", ""),
        step_response=s.get("step_response", ""),
        ramp_response=s.get("ramp_response", ""),
        bode=s.get("bode", ""),
        tf_panel=s.get("tf_panel", ""),
        critical=s.get("critical", ""),
    )

    report_path = fill_report(
        template_path=template_path,
        output_dir=output_dir or config.REPORTS_DIR,
        calc=calc,
        shots=shots,
    )
    return {"report_path": report_path}


# ── 8. check_dependencies ─────────────────────────────────────────────────────

@mcp.tool()
def check_dependencies() -> dict:
    """
    Проверяет все зависимости: Python-пакеты, CLASSiC.exe, рабочие папки.
    Возвращает статус каждого компонента и общий флаг all_ok.
    """
    pkg_status = {}
    for pkg in ["pyautogui", "pygetwindow", "PIL", "psutil", "sympy", "docx", "mcp"]:
        try:
            __import__(pkg)
            pkg_status[pkg] = "ok"
        except ImportError:
            pkg_status[pkg] = "missing"

    exe = Path(config.CLASSIC_EXE)
    dirs_status = {}
    for name, path in [
        ("TASKS_DIR",   config.TASKS_DIR),
        ("REPORTS_DIR", config.REPORTS_DIR),
        ("MDL_DIR",     config.MDL_DIR),
    ]:
        dirs_status[name] = {"path": path, "exists": Path(path).exists()}

    return {
        "packages":    pkg_status,
        "classic_exe": {"path": str(exe), "exists": exe.exists()},
        "directories": dirs_status,
        "all_ok": all(v == "ok" for v in pkg_status.values()) and exe.exists(),
    }


# ── 9. run_full_pipeline ──────────────────────────────────────────────────────

@mcp.tool()
def run_full_pipeline(doc_path: str) -> dict:
    """
    Запускает полный пайплайн обработки одного файла задания:
    парсинг → расчёты → MDL → CLASSiC (скриншоты) → отчёт .docx.

    doc_path: путь к файлу задания (.doc/.docx).
    Возвращает пути ко всем артефактам и краткий итог расчётов.
    """
    from parser import parse_variant
    from calculator import calculate
    from mdl_writer import write_mdl
    from classic_gui import ClassicController
    from report import fill_report

    doc_path = str(Path(doc_path).resolve())

    params = parse_variant(doc_path)
    calc = calculate(params)

    Path(config.MDL_DIR).mkdir(parents=True, exist_ok=True)
    mdl_path = str(Path(config.MDL_DIR) / f"var{params.variant:02d}.mdl")
    write_mdl(params, mdl_path)

    shots_dir = str(Path(config.REPORTS_DIR) / f"v{params.variant:02d}_shots")
    ctrl = ClassicController(
        mdl_path=mdl_path,
        output_dir=shots_dir,
        variant=params.variant,
        K1_critical=calc.K1_critical,
    )
    shots = ctrl.run_all()

    report_path = fill_report(
        template_path=doc_path,
        output_dir=config.REPORTS_DIR,
        calc=calc,
        shots=shots,
    )

    return {
        "variant":     params.variant,
        "mdl_path":    mdl_path,
        "shots_dir":   shots_dir,
        "report_path": report_path,
        "calc": {
            "WP_str":          calc.WP_str,
            "hurwitz_stable":  calc.hurwitz_stable,
            "K1_critical":     calc.K1_critical,
            "K_loop_critical": calc.K_loop_critical,
            "e_ust_step":      calc.e_ust_step,
            "e_ust_ramp":      calc.e_ust_ramp,
        },
    }


if __name__ == "__main__":
    mcp.run()
