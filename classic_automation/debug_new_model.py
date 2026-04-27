"""
Разведка диалога Параметры блока: снимаем ПЕРЕД каждым действием,
чтобы видеть реальное состояние полей.
"""
import time
import subprocess
from pathlib import Path
from PIL import ImageGrab
import pyautogui

import config
from classic_gui import _wait_for_window, _get_win_rect, _find_windows

OUT = Path(r"C:\Users\lud50\Desktop\TU\reports\test_shots")
OUT.mkdir(parents=True, exist_ok=True)
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.25

WIN = "CLASSiC"
DLG = "Параметры блока"

_dlg_bbox = None  # кэшируем bbox диалога при первом обнаружении


def cache_dlg():
    """Запоминаем bbox диалога пока он ещё открыт."""
    global _dlg_bbox
    wins = _find_windows(DLG)
    if wins:
        w = wins[0]
        _dlg_bbox = (w.left, w.top, w.left + w.width, w.top + w.height)
    return _dlg_bbox


def shot(name):
    """Снимает диалог (по кэшированному bbox) или весь экран."""
    bbox = _dlg_bbox
    if bbox:
        img = ImageGrab.grab(bbox=bbox)
    else:
        img = ImageGrab.grab()
    p = str(OUT / f"dlg_{name}.png")
    img.save(p)
    print(f"  [скрин] {p}")


def dlg_rect():
    wins = _find_windows(DLG)
    if not wins:
        return None
    w = wins[0]
    return w.left, w.top, w.width, w.height


# ── Запуск и создание новой модели ───────────────────────────────────────────
proc = subprocess.Popen([config.CLASSIC_EXE])
_wait_for_window(WIN, timeout=10)
time.sleep(3)
pyautogui.press('enter'); time.sleep(0.5)
pyautogui.press('escape'); time.sleep(0.5)

# Файл -> Новый
pyautogui.press('alt');   time.sleep(0.5)
pyautogui.press('down');  time.sleep(0.3)
pyautogui.press('enter'); time.sleep(1.5)   # "Новый" -> диалог выбора типа
pyautogui.press('enter'); time.sleep(2)     # OK -> пустой холст

# Снимаем состояние сразу после "Новый"
def shot_win(name):
    r = _get_win_rect(WIN)
    if r:
        img = ImageGrab.grab(bbox=(r['x'], r['y'], r['x']+r['w'], r['y']+r['h']))
    else:
        img = ImageGrab.grab()
    p = str(OUT / f"step_{name}.png")
    img.save(p)
    print(f"  [скрин] {p}")

shot_win("01_after_new")

# Ставим один блок: F4 -> клик в центр -> Esc
rect = _get_win_rect(WIN)
cx = rect['x'] + rect['w'] // 2
cy = rect['y'] + rect['h'] // 2
print(f"Центр окна: ({cx}, {cy})")

pyautogui.press('f4'); time.sleep(0.5)
shot_win("02_f4_mode")

pyautogui.click(cx, cy); time.sleep(1.0)
shot_win("03_after_place")

pyautogui.press('escape'); time.sleep(0.5)
shot_win("04_after_esc")

# Открываем параметры блока: двойной клик
pyautogui.doubleClick(cx, cy); time.sleep(2.0)
shot_win("05_after_dblclick")

# Проверяем диалог
info = dlg_rect()
if not info:
    print("Диалог не найден! Смотри step_*.png для диагностики.")
    proc.terminate()
    exit()

dx, dy, dw, dh = info
print(f"Диалог: {dw}x{dh} @ ({dx},{dy})")

# Кэшируем bbox пока диалог открыт
cache_dlg()

# ── СНИМОК 0: начальное состояние диалога ────────────────────────────────────
shot("00_initial")

# BackSpace×10 надёжно очищает поле в Win16 spin-edit
def set_field(x, y, value):
    pyautogui.click(x, y); time.sleep(0.25)
    pyautogui.press('end'); time.sleep(0.05)
    for _ in range(12):          # удаляем до 12 символов назад
        pyautogui.press('backspace')
    time.sleep(0.1)
    pyautogui.write(str(value), interval=0.06)
    time.sleep(0.2)

# ── Числитель: set_field (click + End + BackSpace×12 + write) ────────────────
num_x = dx + int(dw * 0.80)
num_y = dy + int(dh * 0.37)
shot("01_before_num")
set_field(num_x, num_y, "2.5")
shot("02_num_typed")

# Нормализуем состояние: простой клик по числителю сбрасывает "text-modified" флаг
pyautogui.click(num_x, num_y); time.sleep(0.3)

# ── Знаменатель s^0 ───────────────────────────────────────────────────────────
# Поле знаменателя на ~46% Y (ниже горизонтальной линии)
den_x = dx + int(dw * 0.80)
den_y = dy + int(dh * 0.46)
print(f"Кликаем знаменатель s^0 @ ({den_x}, {den_y})  [80% x, 46% y]")
pyautogui.click(den_x, den_y); time.sleep(0.5)
shot("03_den_focused")
pyautogui.write("1", interval=0.06)
time.sleep(0.3)
shot("04_den0_typed")

# ── ▼ знаменателя: добавляет степень s^1 (знаменатель растёт вниз) ────────────
# Кнопка ▼ находится ниже поля знаменателя — 63% X, 47% Y
dn_x = dx + int(dw * 0.63)
dn_y = dy + int(dh * 0.47)
print(f"Кликаем [DOWN] знаменателя @ ({dn_x}, {dn_y})  [63% x, 47% y]")
shot("05_before_dn")
pyautogui.click(dn_x, dn_y); time.sleep(0.5)
shot("06_after_dn")

# ── Знаменатель s^1: новое поле появилось ниже (примерно 54% Y) ──────────────
den1_y = dy + int(dh * 0.54)
print(f"Кликаем знаменатель s^1 @ ({den_x}, {den1_y})  [80% x, 54% y]")
pyautogui.click(den_x, den1_y); time.sleep(0.5)
shot("07_den1_focused")
pyautogui.write("0.5", interval=0.06)
time.sleep(0.3)
shot("08_den1_typed")

# ── Нажимаем OK ───────────────────────────────────────────────────────────────
shot("10_before_ok")
pyautogui.press('enter'); time.sleep(0.8)
if _find_windows(DLG):
    shot("11_still_open")
    print("Диалог ещё открыт после Enter - нажимаем Tab+Enter")
    pyautogui.press('tab'); time.sleep(0.1)
    pyautogui.press('enter'); time.sleep(0.5)
else:
    print("Диалог закрыт (OK)")

# ── Верификация: снова открываем диалог блока ────────────────────────────────
time.sleep(0.5)
pyautogui.doubleClick(cx, cy); time.sleep(1.5)
info2 = dlg_rect()
if info2:
    dx2, dy2, dw2, dh2 = info2
    # обновляем кэш для новой позиции диалога
    wins = _find_windows(DLG)
    if wins:
        w2 = wins[0]
        _dlg_bbox = (w2.left, w2.top, w2.left + w2.width, w2.top + w2.height)

    shot("12_reopen_initial")         # что показывает знаменатель при открытии?

    # Нажимаем ▲ (UP) для знаменателя — смотрим предыдущий коэффициент
    up_x2 = dx2 + int(dw2 * 0.63)
    up_y2 = dy2 + int(dh2 * 0.44)
    print(f"[проверка] [UP] знаменатель @ ({up_x2}, {up_y2})")
    pyautogui.click(up_x2, up_y2); time.sleep(0.4)
    shot("13_reopen_up")              # s^0 коэффициент?

    # Нажимаем ▼ (DOWN) — обратно к следующему
    dn_x2 = dx2 + int(dw2 * 0.63)
    dn_y2 = dy2 + int(dh2 * 0.47)
    pyautogui.click(dn_x2, dn_y2); time.sleep(0.4)
    shot("14_reopen_down")            # s^1 коэффициент?

    pyautogui.press('escape'); time.sleep(0.3)
else:
    print("Диалог не открылся при повторном клике")

# ── Текстовая форма для финальной верификации ────────────────────────────────
time.sleep(0.3)
# Вид -> Модель - текстовая форма (сводка)
rect3 = _get_win_rect(WIN)
if rect3:
    cx3 = rect3['x'] + rect3['w'] // 2
    cy3 = rect3['y'] + rect3['h'] - 12
    pyautogui.click(cx3, cy3); time.sleep(0.3)
pyautogui.press('alt'); time.sleep(0.4)
pyautogui.press('right')             # Редактирование
pyautogui.press('right')             # Вид
time.sleep(0.2)
pyautogui.press('down'); time.sleep(0.2)   # открываем Вид
pyautogui.press('down'); time.sleep(0.1)   # Модель - текстовая форма
pyautogui.press('enter'); time.sleep(2.0)

rect4 = _get_win_rect(WIN)
if rect4:
    img = ImageGrab.grab(bbox=(rect4['x'], rect4['y'],
                               rect4['x'] + rect4['w'], rect4['y'] + rect4['h']))
    p = str(OUT / "dlg_15_text_form.png")
    img.save(p)
    print(f"  [скрин] {p}")
pyautogui.press('escape'); time.sleep(0.3)

# ── Финальный снимок схемы ───────────────────────────────────────────────────
time.sleep(0.3)
rect2 = _get_win_rect(WIN)
if rect2:
    img = ImageGrab.grab(bbox=(rect2['x'], rect2['y'],
                               rect2['x'] + rect2['w'], rect2['y'] + rect2['h']))
    p = str(OUT / "dlg_12_schema_after.png")
    img.save(p)
    print(f"  [скрин] {p}")

# ── Закрываем ────────────────────────────────────────────────────────────────
time.sleep(0.5)
pyautogui.hotkey('alt', 'f4'); time.sleep(1)
pyautogui.press('escape'); time.sleep(0.3)
proc.terminate()
print("Готово.")
