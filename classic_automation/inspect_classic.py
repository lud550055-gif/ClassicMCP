"""
inspect_classic.py — запускает CLASSiC и дампит UI-дерево.
"""
import time
import sys
import subprocess
import warnings

try:
    from pywinauto import Application, Desktop
    import pygetwindow as gw
    import psutil
except ImportError as e:
    print(f"Нет зависимости: {e}\npip install pywinauto pygetwindow psutil")
    sys.exit(1)

import config

EXE = config.CLASSIC_EXE

print(f"[1] Запускаем {EXE} ...")
proc = subprocess.Popen([EXE])
print(f"    PID={proc.pid}")

# Ждём до 10 сек пока появится хоть какое-то окно с "CLASSiC" или "classic" в заголовке
print("[2] Ждём окно CLASSiC (до 10 сек)...")
found_win = None
for _ in range(20):
    time.sleep(0.5)
    for w in gw.getAllWindows():
        try:
            if "classic" in w.title.lower() and w.width > 50:
                found_win = w
                break
        except Exception:
            pass
    if found_win:
        break

if not found_win:
    print("  Окно не найдено! Список всех окон:")
    for w in gw.getAllWindows():
        try:
            if w.title.strip():
                print(f"    {w.title!r}  ({w.width}x{w.height})")
        except Exception:
            pass
    sys.exit(1)

print(f"    Найдено окно: {found_win.title!r}  {found_win.width}x{found_win.height}")

# Ищем реальный PID окна через psutil + hwnd
import ctypes
hwnd = found_win._hWnd
pid_buf = ctypes.c_ulong()
ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid_buf))
real_pid = pid_buf.value
print(f"    HWND={hwnd}  реальный PID={real_pid}")

# Проверяем живой ли процесс
try:
    p = psutil.Process(real_pid)
    print(f"    Процесс: {p.name()}  exe={p.exe()}")
except Exception as e:
    print(f"    psutil: {e}")

time.sleep(1)

# ── UIA backend ───────────────────────────────────────────────────────────
print("\n[3] UIA backend ...")
try:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        app_uia = Application(backend="uia").connect(handle=hwnd, timeout=5)
    dlg = app_uia.top_window()
    print(f"    OK: '{dlg.window_text()}'  class='{dlg.class_name()}'")
    print("\n" + "=" * 60)
    print("ДЕРЕВО — UIA (глубина 4):")
    print("=" * 60)
    dlg.print_control_identifiers(depth=4)
except Exception as e:
    print(f"    UIA ошибка: {e}")

# ── win32 backend ─────────────────────────────────────────────────────────
print("\n[4] win32 backend ...")
try:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        app_w32 = Application(backend="win32").connect(handle=hwnd, timeout=5)
    dlg = app_w32.top_window()
    print(f"    OK: '{dlg.window_text()}'  class='{dlg.class_name()}'")
    print("\n" + "=" * 60)
    print("ДЕРЕВО — win32 (глубина 3):")
    print("=" * 60)
    dlg.print_control_identifiers(depth=3)
    print("\n" + "=" * 60)
    print("МЕНЮ — win32:")
    print("=" * 60)
    try:
        menu = dlg.menu()
        for i in range(menu.item_count()):
            item = menu.item(i)
            label = item.text()
            print(f"  [{i}] {label!r}")
            try:
                sub = item.sub_menu()
                for j in range(sub.item_count()):
                    print(f"       [{j}] {sub.item(j).text()!r}")
            except Exception:
                pass
    except Exception as e:
        print(f"  меню ошибка: {e}")
except Exception as e:
    print(f"    win32 ошибка: {e}")

print("\n[готово] CLASSiC остаётся открытым — закрой вручную.")
