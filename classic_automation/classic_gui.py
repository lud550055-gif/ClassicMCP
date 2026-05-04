"""
classic_gui.py — управление CLASSiC 3.2 через pyautogui.

CLASSiC 3.2 — Win16-приложение под otvdm.exe, UI-дерева нет.
Навигация: клавиатурные шоткаты + поиск окон по заголовку.

Меню: Файл | Редактирование | Вид | Расчеты | Окно | Справка
Расчеты → item 1 → окно «Характеристики» (4 панели: корневая, переходные, частотные, ПФ)

Стратегия загрузки файла:
  1. До запуска — записываем mdl_path как File1 в [MRU] CLASSiC.ini
  2. Запускаем CLASSiC без аргументов
  3. Закрываем splash (Enter/Esc)
  4. Файл → File1 (Alt, Down, Down×11, Enter) — минуя Win16-диалог
"""
import re
import time
import subprocess
import ctypes
import ctypes.wintypes
from pathlib import Path
from typing import Optional
from dataclasses import dataclass

try:
    import pyautogui
    import pygetwindow as gw
    from PIL import ImageGrab
except ImportError as e:
    raise ImportError(f"Установи зависимости: pip install pyautogui pygetwindow Pillow\n{e}")

import config

pyautogui.FAILSAFE = True
pyautogui.PAUSE = config.ACTION_DELAY

# ── Заголовки окон CLASSiC ────────────────────────────────────────────────────
_WIN_MAIN  = "CLASSiC"        # главное окно (содержит подстроку)
_INI_PATH  = Path(config.CLASSIC_EXE).parent / "CLASSiC.ini"

# Win32-поведение подтверждено: Down на меню-баре открывает меню и сразу
# ставит курсор на item 1 (Новый). Чтобы попасть на item N, нужно N-1 Down.
# File1 = item 11 → 10 Down-нажатий от Новый.
_MRU_DOWNS = 10


@dataclass
class Screenshots:
    schema:          str = ""  # структурная схема (главное окно)
    text_form:       str = ""  # Вид → Модель - текстовая форма (сводка)
    characteristics: str = ""  # весь F9-экран целиком
    root_locus:      str = ""  # Корневая плоскость    (верх-лево)
    step_response:   str = ""  # Переходные процессы   (верх-право)
    ramp_response:   str = ""  # Переходные процессы при Линейном входе 0.1t
    bode:            str = ""  # Частотные характеристики (низ-лево)
    tf_panel:        str = ""  # Передаточные функции  (низ-право)
    critical:        str = ""  # весь F9-экран при K1=K1кр


class ClassicController:

    def __init__(self, mdl_path: str, output_dir: str,
                 variant: int, K1_critical: float = 0.0):
        self.mdl_path    = str(mdl_path)
        self.output_dir  = Path(output_dir)
        self.variant     = variant
        self.K1_critical = K1_critical
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self._proc: Optional[subprocess.Popen] = None

    # ── Публичный API ─────────────────────────────────────────────────────────

    def run_all(self) -> Screenshots:
        shots = Screenshots()
        try:
            self.launch()

            # 1. Структурная схема
            shots.schema = self._shot_win(_WIN_MAIN, "schema")

            # 2. Текстовая форма: Вид → Модель - текстовая форма (сводка)
            shots.text_form = self._text_form()

            # 3. Расчеты → Характеристики (все 4 панели)
            self._open_characteristics()
            self._wait_for_render()

            shots.characteristics = self._shot_win(_WIN_MAIN, "characteristics")
            crops = self._quad_crops(_WIN_MAIN)
            shots.root_locus    = self._save_crop(crops.get('tl'), "root_locus")
            shots.step_response = self._save_crop(crops.get('tr'), "step_response")
            shots.bode          = self._save_crop(crops.get('bl'), "bode")
            shots.tf_panel      = self._save_crop(crops.get('br'), "tf_panel")

            # 4. Рамповый вход f(t)=0.1t → рис.4
            shots.ramp_response = self._ramp_shot(crops.get('tr'))

            # 5. K1кр — меняем параметр блока, снова F9
            if self.K1_critical > 0:
                shots.critical = self._critical_shot()

        except Exception as e:
            print(f"[CLASSiC] run_all ошибка: {e}")
            import traceback
            traceback.print_exc()
        finally:
            self.close()

        return shots

    def launch(self):
        """
        1. Копирует MDL в папку CLASSiC (Win16/otvdm открывает только оттуда).
        2. Записывает путь к копии как File1 в CLASSiC.ini ([MRU]).
        3. Запускает CLASSiC без аргументов.
        4. Ждёт splash и закрывает его.
        5. Открывает модель через Файл → первый MRU-файл.
        """
        import shutil
        print("[CLASSiC] Запуск...")

        # Копируем MDL в папку CLASSiC — Win16/otvdm читает файлы только отсюда
        classic_dir = Path(config.CLASSIC_EXE).parent
        mdl_name = Path(self.mdl_path).name
        self._tmp_mdl = classic_dir / mdl_name
        shutil.copy2(self.mdl_path, self._tmp_mdl)
        print(f"[MDL] Скопирован: {self._tmp_mdl}")

        self._set_mru_file1(str(self._tmp_mdl))

        self._proc = subprocess.Popen([config.CLASSIC_EXE])
        _wait_for_window(_WIN_MAIN, timeout=config.CLASSIC_LAUNCH_TIMEOUT)

        # Перемещаем окно на основной монитор сразу после появления
        _move_to_primary(_WIN_MAIN)

        print("[CLASSiC] Ждём закрытия splash...")
        time.sleep(3)
        self._dismiss_any_dialog()
        time.sleep(2)
        self._dismiss_any_dialog()

        self._open_from_mru()
        self._dismiss_any_dialog()   # закрываем диалог ошибки если файл не открылся
        print("[CLASSiC] Запущен, модель загружена.")

    def close(self):
        for w in _find_windows(_WIN_MAIN):
            try:
                w.activate()
                time.sleep(0.3)
                pyautogui.hotkey('alt', 'f4')
                time.sleep(1.0)
            except Exception:
                pass
        time.sleep(0.5)
        for label in ["Нет", "No"]:
            for w in _find_windows(label):
                try:
                    w.activate()
                    pyautogui.press('enter')
                except Exception:
                    pass
        if self._proc:
            try:
                self._proc.terminate()
            except Exception:
                pass
        # Удаляем временную копию MDL из папки CLASSiC
        if hasattr(self, '_tmp_mdl'):
            try:
                Path(self._tmp_mdl).unlink(missing_ok=True)
            except Exception:
                pass

    # ── Загрузка модели через MRU ─────────────────────────────────────────────

    def _set_mru_file1(self, mdl_path: str):
        """
        Записывает mdl_path как File1 в [MRU] раздел CLASSiC.ini.
        Файл читается/пишется побайтово чтобы не испортить Кириллицу в других
        строках (они хранятся в CP1251).
        """
        try:
            raw = _INI_PATH.read_bytes()
            new_val = b'File1=' + mdl_path.encode('ascii')

            # Удаляем ВСЕ существующие File1= строки (включая дубли),
            # затем добавляем одну актуальную после [MRU]
            cleaned = re.sub(rb'(?m)^File1=[^\r\n]*\r?\n?', b'', raw)
            new_raw = re.sub(
                rb'(\[MRU\]\r?\n)',
                lambda m: m.group(0) + new_val + b'\r\n',
                cleaned,
            )
            _INI_PATH.write_bytes(new_raw)
            print(f"[INI] File1 -> {mdl_path}")
        except Exception as e:
            print(f"[INI] Предупреждение: не удалось изменить CLASSiC.ini: {e}")

    def _open_from_mru(self):
        """
        Открывает File1 через меню Файл без использования диалога Ctrl+O.

        Схема навигации (разделители пропускаются Win16-менеджером меню):
          Alt          → активируем меню-бар, фокус на «Файл»
          Down         → открываем меню Файл, курсор на «Новый» (item 1)
          Down × 10    → Открыть→Сохранить→Сохранить как→Закрыть→
                         Экспорт→Импорт→Печать→Настройки→Выход→File1
          Enter        → открываем File1
        """
        print("[CLASSiC] Открываем File1 из меню Файл...")
        self._focus(_WIN_MAIN)
        time.sleep(0.5)

        pyautogui.press('alt')        # активируем меню-бар → Файл в фокусе
        time.sleep(0.5)
        pyautogui.press('down')       # открываем Файл, курсор на «Новый»
        time.sleep(0.3)

        for _ in range(_MRU_DOWNS):
            pyautogui.press('down')
            time.sleep(0.15)

        pyautogui.press('enter')
        print("[CLASSiC] Ждём загрузки модели (6 сек)...")
        time.sleep(6)

    # ── Характеристики ────────────────────────────────────────────────────────

    def _open_characteristics(self):
        """
        Расчеты (4-й пункт меню) → первый пункт = Характеристики.

        Важно: кликаем в ЦЕНТР холста (не статус-бар), чтобы снять фокус
        с меню-бара если он остался активным после предыдущей операции.
        Затем Alt активирует меню-бар с Файл в фокусе → Right×3 → Расчеты.
        """
        print("[CLASSiC] Открываем Характеристики через меню Расчеты...")
        # Кликаем в центр холста — снимает любой фокус с меню-бара
        rect = _get_win_rect(_WIN_MAIN)
        if rect:
            cx = rect['x'] + rect['w'] // 2
            cy = rect['y'] + rect['h'] // 2
            pyautogui.click(cx, cy)
        time.sleep(0.5)
        pyautogui.press('alt')
        time.sleep(0.5)
        pyautogui.press('right')   # Редактирование
        time.sleep(0.1)
        pyautogui.press('right')   # Вид
        time.sleep(0.1)
        pyautogui.press('right')   # Расчеты
        time.sleep(0.4)
        pyautogui.press('down')    # открыть Расчеты → курсор на item 1
        time.sleep(0.3)
        pyautogui.press('enter')   # Характеристики

    # ── Текстовая форма ───────────────────────────────────────────────────────

    def _text_form(self) -> str:
        """
        Вид (3-й пункт меню) → Модель - текстовая форма (сводка) (2-й пункт).
        Alt → Right×2 → вид подсвечен; Down → меню открыто; Down×2 → нужный пункт.
        """
        print("[CLASSiC] Текстовая форма...")
        self._focus(_WIN_MAIN)
        pyautogui.press('alt')
        time.sleep(0.4)
        pyautogui.press('right')      # → Редактирование
        pyautogui.press('right')      # → Вид
        time.sleep(0.2)
        pyautogui.press('down')       # открыть Вид → курсор на «Табличная форма» (item 1)
        time.sleep(0.2)
        pyautogui.press('down')       # курсор на «Модель - текстовая форма (сводка)» (item 2)
        pyautogui.press('enter')
        time.sleep(config.RENDER_TIMEOUT + 0.5)

        path = self._shot_win(_WIN_MAIN, "text_form")
        # Закрываем панель текстовой формы — Esc возвращает к схеме
        self._focus(_WIN_MAIN)
        pyautogui.press('escape')
        time.sleep(0.3)
        return path

    # ── Рамповый вход (задача 8) ─────────────────────────────────────────────

    def _ramp_shot(self, tr_bbox: Optional[tuple]) -> str:
        """
        Меняет тип входного воздействия «Переходных процессов» на Линейное (0.1),
        снимает tr-квадрант → ramp_response, восстанавливает Ступенчатое (1.0).

        Маршрут: правый клик по tr-панели → Down×2 (→ Тип) → Enter →
          диалог «Переходные процессы: Тип»: Down (→ Линейное) → Tab → коэф → Enter.
        """
        if not tr_bbox:
            return ""
        print("[CLASSiC] Рамповый отклик (Линейное, 0.1)...")

        cx = (tr_bbox[0] + tr_bbox[2]) // 2
        cy = (tr_bbox[1] + tr_bbox[3]) // 2

        def _open_type_dlg():
            self._focus(_WIN_MAIN)
            pyautogui.rightClick(cx, cy)
            time.sleep(0.4)
            pyautogui.press('down')   # → «Параметры» (1-й пункт)
            time.sleep(0.1)
            pyautogui.press('down')   # → «Тип»       (2-й пункт)
            time.sleep(0.1)
            pyautogui.press('enter')
            time.sleep(0.5)

        def _choose(arrow: str, coeff: str):
            """Сдвигает radio на один шаг (arrow='down'/'up'), ставит коэф., OK."""
            pyautogui.press(arrow)
            time.sleep(0.1)
            pyautogui.press('tab')        # → поле «Коэффициент»
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write(coeff, interval=0.05)
            pyautogui.press('enter')      # OK (default button)
            time.sleep(0.3)

        # Ступенчатое → Линейное
        _open_type_dlg()
        _choose('down', '0.1')
        self._wait_for_render()
        path = self._save_crop(tr_bbox, "ramp_response")

        # Восстановить: Линейное → Ступенчатое
        _open_type_dlg()
        _choose('up', '1.0')
        self._wait_for_render()

        return path

    # ── Симуляция при K1кр ───────────────────────────────────────────────────

    def _critical_shot(self) -> str:
        """
        Закрывает окно Характеристик → двойной клик по блоку K1 →
        диалог «Параметры блока» → меняем числитель → переоткрываем
        Характеристики → скриншот → восстанавливаем исходный .mdl.

        Координаты блока K1: (0.26, 0.32) — откалиброваны по схеме
        1375×994 на разрешении 3440×1440 @ 125% DPI.
        """
        print(f"[CLASSiC] K1кр = {self.K1_critical}...")

        # 1. Закрываем окно Характеристик (MDI-дочернее) — Ctrl+F4
        self._focus(_WIN_MAIN)
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'f4')
        time.sleep(0.8)

        # 2. Двойной клик по блоку K1 (второй блок на схеме)
        self._focus(_WIN_MAIN)
        rect = _get_win_rect(_WIN_MAIN)
        if rect:
            cx = rect['x'] + int(rect['w'] * 0.26)
            cy = rect['y'] + int(rect['h'] * 0.32)
            print(f"  [K1 клик] ({cx}, {cy})")
            pyautogui.doubleClick(cx, cy)
            time.sleep(config.DIALOG_TIMEOUT)

            # 3. В диалоге «Параметры блока» меняем значение числителя
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write(str(round(self.K1_critical, 4)), interval=0.05)
            pyautogui.press('enter')
            time.sleep(0.5)

        # 4. Открываем Характеристики и снимаем скриншот
        self._open_characteristics()
        self._wait_for_render()
        path = self._shot_win(_WIN_MAIN, "critical")

        # 5. Восстанавливаем исходную модель
        self._set_mru_file1(str(self._tmp_mdl))
        self._open_from_mru()
        return path

    # ── Ожидание отрисовки ───────────────────────────────────────────────────

    def _wait_for_render(self, timeout: float = 30.0, poll: float = 0.4):
        """
        Ждёт пока содержимое окна CLASSiC изменится, затем стабилизируется.
        Фаза 1: ждём первого изменения (окно начало открываться).
        Фаза 2: ждём стабилизации (отрисовка завершена).
        """
        CHANGE_THRESH = 500
        STABLE_THRESH = 200
        STABLE_FRAMES = 3

        def _snap():
            rect = _get_win_rect(_WIN_MAIN)
            if not rect:
                return None
            img = ImageGrab.grab(bbox=(
                rect['x'], rect['y'],
                rect['x'] + rect['w'], rect['y'] + rect['h']
            ))
            return list(img.resize((40, 30)).getdata())

        def _diff(a, b):
            if a is None or b is None:
                return 0
            return sum(abs(x[0]-y[0]) + abs(x[1]-y[1]) + abs(x[2]-y[2])
                       for x, y in zip(a, b))

        deadline = time.time() + timeout

        # Фаза 1: ждём первого изменения
        prev = _snap()
        changed = False
        while time.time() < deadline:
            time.sleep(poll)
            curr = _snap()
            if _diff(prev, curr) > CHANGE_THRESH:
                changed = True
                break
            prev = curr

        if not changed:
            print("  [render] таймаут ожидания изменения")
            return

        # Фаза 2: ждём стабилизации (N подряд одинаковых кадров)
        prev = _snap()
        stable = 0
        while time.time() < deadline:
            time.sleep(poll)
            curr = _snap()
            if _diff(prev, curr) < STABLE_THRESH:
                stable += 1
                if stable >= STABLE_FRAMES:
                    time.sleep(0.5)  # небольшой запас после стабилизации
                    print("  [render] окно стабилизировалось")
                    return
            else:
                stable = 0
            prev = curr

        print("  [render] таймаут ожидания стабилизации")

    # ── Скриншоты ────────────────────────────────────────────────────────────

    def _shot_win(self, title_sub: str, name: str) -> str:
        time.sleep(0.5)
        rect = _get_win_rect(title_sub)
        if rect:
            img = ImageGrab.grab(bbox=(
                rect['x'], rect['y'],
                rect['x'] + rect['w'], rect['y'] + rect['h']
            ))
        else:
            img = ImageGrab.grab()
        path = str(self.output_dir / f"v{self.variant:02d}_{name}.png")
        img.save(path)
        print(f"  [скриншот] {path}")
        return path

    def _quad_crops(self, title_sub: str) -> dict:
        """
        Режет внешнее окно CLASSiC на 4 квадранта панелей «Характеристики».

        MDI-дочерние окна не перечисляются pygetwindow, поэтому используем
        внешний rect и пропорции, найденные пиксельным сканированием
        скриншота 1375×994 px @ 125% DPI:
          x1/x2 — левый/правый край панелей
          mx    — вертикальный разделитель (x=590/1375 ≈ 0.4291)
          y1/y2 — верхний/нижний край панелей
          my    — горизонтальный разделитель (y=431/994 ≈ 0.4336)
        Разделители замерены напрямую, а не как середина x1..x2.
        """
        rect = _get_win_rect(title_sub)
        if not rect:
            return {}
        x, y, w, h = rect['x'], rect['y'], rect['w'], rect['h']

        # Пропорции получены пиксельным сканом v17_characteristics.png
        x1 = x + int(w * 0.0305)  # ≈ 42px   — левый край панелей
        x2 = x + int(w * 0.8000)  # ≈ 1100px — правый край панелей
        mx = x + int(w * 0.4291)  # ≈ 590px  — вертикальный разделитель
        y1 = y + int(h * 0.1006)  # ≈ 100px  — верх панелей
        y2 = y + int(h * 0.6922)  # ≈ 688px  — низ панелей
        my = y + int(h * 0.4336)  # ≈ 431px  — горизонтальный разделитель

        return {
            'tl': (x1, y1, mx, my),  # Корневая плоскость
            'tr': (mx, y1, x2, my),  # Переходные процессы
            'bl': (x1, my, mx, y2),  # Частотные характеристики
            'br': (mx, my, x2, y2),  # Передаточные функции
        }

    def _save_crop(self, bbox: Optional[tuple], name: str) -> str:
        if not bbox:
            return ""
        img = ImageGrab.grab(bbox=bbox)
        path = str(self.output_dir / f"v{self.variant:02d}_{name}.png")
        img.save(path)
        print(f"  [кроп] {path}")
        return path

    def _dismiss_any_dialog(self):
        """Закрывает любой модальный диалог (splash и др.) нажатием Enter/Esc.
        Не пытается перевести фокус — диалог уже в фокусе (он модальный)."""
        pyautogui.press('enter')
        time.sleep(0.4)
        pyautogui.press('escape')
        time.sleep(0.3)

    def _focus(self, title_sub: str):
        """
        Кликает по статус-бару CLASSiC.
        pyautogui.click() — единственный надёжный способ переключить фокус
        на Win16/otvdm окно; activate() не работает.
        Статус-бар безопасен: кликая туда CLASSiC не открывает меню.
        """
        rect = _get_win_rect(title_sub)
        if rect:
            cx = rect['x'] + rect['w'] // 2
            cy = rect['y'] + rect['h'] - 12   # статус-бар
            pyautogui.click(cx, cy)
            time.sleep(0.3)


# ── Модульные утилиты ────────────────────────────────────────────────────────

def _find_windows(substr: str) -> list:
    result = []
    for w in gw.getAllWindows():
        try:
            if substr.lower() in w.title.lower() and w.width > 10:
                result.append(w)
        except Exception:
            pass
    return result


def _get_win_rect(title_sub: str) -> Optional[dict]:
    wins = _find_windows(title_sub)
    if not wins:
        return None
    main = max(wins, key=lambda w: w.width * w.height)
    try:
        return {'x': main.left, 'y': main.top, 'w': main.width, 'h': main.height}
    except Exception:
        return None


def _wait_for_window(title_sub: str, timeout: int = 30):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _find_windows(title_sub):
            return
        time.sleep(0.5)
    raise TimeoutError(f"Окно '{title_sub}' не появилось за {timeout} сек.")


def _move_to_primary(title_sub: str, x: int = 200, y: int = 150):
    """
    Перемещает окно CLASSiC на основной монитор в заданную позицию.
    Использует SetWindowPos с физическими координатами (процесс DPI-aware через pyautogui).
    """
    import ctypes
    wins = _find_windows(title_sub)
    if not wins:
        return
    hwnd = wins[0]._hWnd
    rect = ctypes.wintypes.RECT()
    ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
    w = rect.right - rect.left
    h = rect.bottom - rect.top
    # SWP_NOZORDER=0x0004, SWP_NOACTIVATE=0x0010
    ctypes.windll.user32.SetWindowPos(hwnd, None, x, y, w, h, 0x0004 | 0x0010)
    time.sleep(0.3)
    print(f"[CLASSiC] Окно перемещено на основной монитор: ({x},{y}) {w}x{h}")
