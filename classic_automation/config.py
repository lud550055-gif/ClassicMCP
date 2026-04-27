# ============================================================
#  config.py — пути вычисляются относительно проекта
# ============================================================
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent.parent

CLASSIC_EXE = str(_PROJECT_ROOT / "classic" / "CLASSiC32.exe")

SCREEN_WIDTH  = 3440
SCREEN_HEIGHT = 1440

TASKS_DIR   = str(_PROJECT_ROOT / "tasks")
REPORTS_DIR = str(_PROJECT_ROOT / "reports")
MDL_DIR     = str(_PROJECT_ROOT / "models")

# Таймауты (секунды)
CLASSIC_LAUNCH_TIMEOUT = 8    # ожидание запуска CLASSiC
DIALOG_TIMEOUT         = 3    # ожидание появления диалога
RENDER_TIMEOUT         = 2    # ожидание отрисовки графика
ACTION_DELAY           = 0.4  # пауза между действиями pyautogui
