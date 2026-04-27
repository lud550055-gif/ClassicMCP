"""Отладка открытия файла — снимаем только окно CLASSiC."""
import time, subprocess, shutil, ctypes
from pathlib import Path
from PIL import ImageGrab
import pyautogui, pygetwindow as gw
import config

def find_win():
    wins = [w for w in gw.getAllWindows()
            if "classic" in w.title.lower() and w.width > 50]
    return max(wins, key=lambda w: w.width * w.height) if wins else None

def grab_win(name):
    w = find_win()
    if not w:
        ImageGrab.grab().save(f"debug_{name}.png")
        print(f"  [debug] {name} (full screen — CLASSiC не найден)")
        return
    img = ImageGrab.grab(bbox=(w.left, w.top, w.left + w.width, w.top + w.height))
    img.save(f"debug_{name}.png")
    print(f"  [debug] {name}  ({w.width}x{w.height} @ {w.left},{w.top})")

def focus_win():
    w = find_win()
    if not w:
        return
    # win32 SetForegroundWindow — надёжнее, чем activate()
    hwnd = w._hWnd
    ctypes.windll.user32.ShowWindow(hwnd, 9)   # SW_RESTORE
    ctypes.windll.user32.SetForegroundWindow(hwnd)
    time.sleep(0.5)
    # Клик по центру body окна (ниже menu bar)
    cx = w.left + w.width // 2
    cy = w.top + 60
    pyautogui.click(cx, cy)
    time.sleep(0.5)

# ── Запуск ────────────────────────────────────────────────────────────────────
proc = subprocess.Popen([config.CLASSIC_EXE])
print("Запущен, ждём 4 сек...")
time.sleep(4)

w = find_win()
if not w:
    print("Окно не найдено!"); proc.terminate(); exit()
print(f"Окно: {w.title!r}  {w.width}x{w.height}  @ {w.left},{w.top}")

# Копируем MDL
src = r"C:\Users\lud50\Desktop\TU\models\var17.mdl"
dst = Path(config.CLASSIC_EXE).parent / "var17.mdl"
shutil.copy2(src, dst)

# ── Шаг 1: фокус ─────────────────────────────────────────────────────────────
focus_win()
grab_win("1_focused")

# ── Шаг 2: Ctrl+O ────────────────────────────────────────────────────────────
print("Ctrl+O...")
pyautogui.hotkey('ctrl', 'o')
time.sleep(5)
grab_win("2_after_ctrlO")

# ── Шаг 3: вводим имя ────────────────────────────────────────────────────────
print("Вводим имя файла...")
pyautogui.press('home')
time.sleep(0.3)
pyautogui.hotkey('shift', 'end')
time.sleep(0.3)
grab_win("3_selected")

pyautogui.write("var17.mdl", interval=0.1)
time.sleep(0.5)
grab_win("4_typed")

# ── Шаг 4: Enter ─────────────────────────────────────────────────────────────
print("Enter...")
pyautogui.press('enter')
time.sleep(4)
grab_win("5_after_enter")

dst.unlink(missing_ok=True)
proc.terminate()
print("Готово.")
