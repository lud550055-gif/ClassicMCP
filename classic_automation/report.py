# report.py — заполняет задание LP_vXX.doc результатами и скриншотами
import shutil
from pathlib import Path

from docx import Document
from docx.shared import Inches

from calculator import CalcResults
from classic_gui import Screenshots

ELLIPSIS = '…'  # …
_COL_W = 10          # ширина колонки таблицы (9 пробелов + символ)
_DEG_W = 7           # ширина колонки степени (3 + digit + 3)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _convert_doc_to_docx(doc_path: str, out_path: str) -> str:
    """Конвертирует .doc → .docx через Word COM."""
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(Path(doc_path).resolve()))
        doc.SaveAs2(str(Path(out_path).resolve()), FileFormat=16)
        doc.Close(False)
    finally:
        word.Quit()
    return out_path


def _replace_ellipsis(para, replacements: list[str]):
    """Заменяет … в прогоне параграфа значениями из списка по порядку."""
    i = 0
    for run in para.runs:
        while ELLIPSIS in run.text and i < len(replacements):
            run.text = run.text.replace(ELLIPSIS, replacements[i], 1)
            i += 1


def _fmt_cell(val: float) -> str:
    """Значение для 10-символьной колонки таблицы (правое выравнивание)."""
    if val == 0.0:
        return ' ' * _COL_W
    return f'{val:.4g}'.rjust(_COL_W)


def _make_table_row(num: float, den: float, deg: int, first: bool = False) -> str:
    """Формирует строку ASCII-таблицы CLASSiC с заданными значениями."""
    col1 = '| Ном.Система  |' if first else '|              |'
    return f"{col1}{_fmt_cell(num)} |{_fmt_cell(den)} |   {deg}   |"


def _rewrite_table_rows(rows: list, classic: dict):
    """
    Перезаписывает уже заполненные строки ASCII-таблицы (без «…»).
    Используется для WP-таблицы с образцовыми значениями другого варианта.
    """
    num = classic.get('num', [0.0])
    den = classic.get('den', [0.0])
    max_deg = max(len(num), len(den)) - 1

    def get(lst, i):
        return lst[i] if i < len(lst) else 0.0

    for row_i, para in enumerate(rows):
        if not para.runs:
            continue
        first = row_i == 0
        new_text = _make_table_row(get(num, row_i), get(den, row_i), row_i, first)
        para.runs[0].text = new_text


def _fill_table_rows(paras: list, classic: dict):
    """
    Заполняет строки ASCII-таблицы CLASSiC коэффициентами ПФ.
    classic: {"num": [c0,c1,...], "den": [c0,c1,...]} от s^0.
    Если степень > 2, вставляет дополнительные строки перед последней.
    """
    import copy

    num = classic.get('num', [0.0])
    den = classic.get('den', [0.0])

    def get(lst, i):
        return lst[i] if i < len(lst) else 0.0

    max_deg = max(len(num), len(den)) - 1

    def _write_row(run, deg, is_last=False):
        text = run.text
        n_val = _fmt_cell(get(num, deg))
        d_val = _fmt_cell(get(den, deg))
        text = text.replace(' ' * 9 + ELLIPSIS, n_val, 1)
        text = text.replace(' ' * 9 + ELLIPSIS, d_val, 1)
        if is_last:
            text = text.replace(f'   {ELLIPSIS}   ', f'   {deg}   ', 1)
        run.text = text

    # Строки 0 и 1 (фиксированные степени)
    if len(paras) > 0 and paras[0].runs:
        _write_row(paras[0].runs[0], 0)
    if len(paras) > 1 and paras[1].runs:
        _write_row(paras[1].runs[0], 1)

    # Для степеней 2..max_deg-1 вставляем дополнительные строки перед шаблонной
    last_para = paras[2] if len(paras) > 2 else None
    if last_para:
        for deg in range(2, max_deg):
            new_elem = copy.deepcopy(last_para._element)
            last_para._element.addprevious(new_elem)
            # Найдём <w:t> внутри скопированного элемента и перезапишем
            ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            for wt in new_elem.findall(f'.//{{{ns}}}t'):
                text = wt.text or ''
                n_val = _fmt_cell(get(num, deg))
                d_val = _fmt_cell(get(den, deg))
                text = text.replace(' ' * 9 + ELLIPSIS, n_val, 1)
                text = text.replace(' ' * 9 + ELLIPSIS, d_val, 1)
                text = text.replace(f'   {ELLIPSIS}   ', f'   {deg}   ', 1)
                wt.text = text

        # Заполняем последнюю (шаблонную) строку — максимальная степень
        if last_para.runs:
            _write_row(last_para.runs[0], max_deg, is_last=True)


def _insert_image_in_para(para, img_path: str, width_in: float = 5.5):
    """Добавляет изображение в существующий (пустой) параграф."""
    if not img_path or not Path(img_path).exists():
        return
    run = para.add_run()
    run.add_picture(img_path, width=Inches(width_in))


def _insert_image_before(doc, target_para, img_path: str, width_in: float = 5.5):
    """Вставляет новый параграф с изображением перед target_para."""
    if not img_path or not Path(img_path).exists():
        return
    tmp = doc.add_paragraph()
    tmp.alignment = 1  # CENTER
    tmp.add_run().add_picture(img_path, width=Inches(width_in))
    target_para._element.addprevious(tmp._element)


# ── Main ──────────────────────────────────────────────────────────────────────

def fill_report(
    template_path: str,
    output_dir:    str,
    calc:          CalcResults,
    shots:         Screenshots,
) -> str:
    """
    Заполняет задание LP_vXX результатами расчётов и скриншотами.
    Возвращает путь к готовому .docx файлу.
    """
    src = Path(template_path)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    variant_str = f'{calc.variant:02d}'
    out_path = out_dir / f'LP_v{variant_str}_done.docx'

    # 1. .doc → .docx если нужно
    if src.suffix.lower() == '.doc':
        docx_src = src.with_suffix('.docx')
        if not docx_src.exists():
            print('[report] Конвертация .doc → .docx...')
            _convert_doc_to_docx(str(src), str(docx_src))
        src = docx_src

    shutil.copy(src, out_path)
    doc = Document(str(out_path))
    paras = doc.paragraphs

    # Извлекаем только правую часть формул (убираем "X(s) = " префикс)
    def _rhs(formula: str) -> str:
        return formula.split(' = ', 1)[-1] if ' = ' in formula else formula

    wp_rhs   = _rhs(calc.WP_str)
    phi_rhs  = _rhs(calc.Phi_str)
    phie_rhs = _rhs(calc.PhiE_str)

    # 2. ── WP-формула и WP-таблица ────────────────────────────────────────
    # Строка вида "WP(s)= 0.1/(...)" — образец другого варианта; заменяем
    for p in paras:
        txt = p.text
        if txt.startswith('WP(s)=') and ELLIPSIS not in txt and len(txt) > 7:
            # Находим run с '= ' и перезаписываем с этой позиции
            for i, run in enumerate(p.runs):
                if '= ' in run.text or run.text == '=':
                    run.text = f'= {wp_rhs}'
                    for r2 in p.runs[i + 1:]:
                        r2.text = ''
                    break
            break  # только первое совпадение

    # Строки WP-таблицы: заканчиваются на |   N   | (колонка степени),
    # без «…», до первой Phi-строки с «…»
    import re as _re
    wp_rows: list = []
    for p in paras:
        if p.text.startswith('|') and ELLIPSIS in p.text:
            break  # дошли до Phi-таблицы
        if ELLIPSIS not in p.text and _re.search(r'\|   \d+   \|$', p.text):
            wp_rows.append(p)

    # Перезаписываем WP-строки значениями WP_classic
    _rewrite_table_rows(wp_rows, calc.WP_classic)

    # 3. ── Текстовые замены ────────────────────────────────────────────────
    phie_count = 0

    for p in paras:
        txt = p.text

        # Ф(s)= … — один параграф с числовым ответом (задача 5)
        if txt.startswith('Ф(s)=') and ELLIPSIS in txt:
            _replace_ellipsis(p, [phi_rhs])

        # Фe(s)= … — первый (формула), второй (числовой ответ задачи 6)
        elif txt.startswith('Фe(s)=') and ELLIPSIS in txt:
            phie_count += 1
            if phie_count == 1:
                _replace_ellipsis(p, ['1 / (1 + WP(s))'])
            elif phie_count == 2:
                _replace_ellipsis(p, [phie_rhs])

        # eуст = lim … = … (задача 7, ступенчатое)
        elif txt.startswith('eуст=lim') and ELLIPSIS in txt:
            expr = 'Фe(0)'
            _replace_ellipsis(p, [expr, calc.e_ust_step])

        # eуст = … (задача 8, рампа)
        elif txt.startswith('eуст=') and ELLIPSIS in txt and 'lim' not in txt:
            _replace_ellipsis(p, [calc.e_ust_ramp])

        # Kкр = …
        elif txt.startswith('Kкр=') and ELLIPSIS in txt:
            _replace_ellipsis(p, [str(round(calc.K_loop_critical, 4))])

        # Имя MDL-файла: 'Модель сохранена в файле … .mdl.'
        # run с «…» идёт перед run ' .mdl.' — убираем пробел после замены
        elif 'Модель сохранена в файле' in txt and ELLIPSIS in txt:
            for j, run in enumerate(p.runs):
                if ELLIPSIS in run.text:
                    run.text = run.text.replace(ELLIPSIS, f'var{variant_str}')
                    if j + 1 < len(p.runs):
                        p.runs[j + 1].text = p.runs[j + 1].text.lstrip()
                    break

    # 3. ── Таблицы коэффициентов ПФ ───────────────────────────────────────
    # Находим тройки строк таблицы по шаблону (строка начинается с '|' и содержит …)
    table_groups: list[list] = []
    buf: list = []
    for p in paras:
        if p.text.startswith('|') and ELLIPSIS in p.text:
            buf.append(p)
            if len(buf) == 3:
                table_groups.append(buf)
                buf = []
        else:
            buf = []

    # Первая группа — Ф(s), вторая — Фe(s)
    if len(table_groups) >= 1:
        _fill_table_rows(table_groups[0], calc.Phi_classic)
    if len(table_groups) >= 2:
        _fill_table_rows(table_groups[1], calc.PhiE_classic)

    # 4. ── Тестовые вопросы (выделение правильного ответа жирным) ────────
    # Q10: устойчивость по Гурвицу — варианты на отдельных строках
    stability_answers = {
        True:   '1:  система устойчива,',
        False:  None,   # определяем ниже
    }
    # Если неустойчива — проверяем чем именно
    if not calc.hurwitz_stable:
        coeffs = calc.char_coeffs
        if all(c > 0 for c in coeffs):
            # Все коэффициенты > 0, но определитель ≤ 0 → граница устойчивости
            q10_answer = '2:  система нейтральна (находится на нейтральной границе устойчивости),'
        else:
            q10_answer = '4:  система неустойчива.'
    else:
        q10_answer = '1:  система устойчива,'

    for p in paras:
        if p.text.strip() == q10_answer.strip():
            for run in p.runs:
                run.bold = True
            break

    # 5. ── Скриншоты ──────────────────────────────────────────────────────
    # Находим параграфы с подписями к рисункам
    ris3 = ris5 = ris6 = None
    for p in paras:
        t = p.text.strip()
        if t == 'Рис. 3':
            ris3 = p
        elif 'Рис. 5' in t:
            ris5 = p
        elif t == 'Рис. 6':
            ris6 = p

    # Рис.3: пустой параграф перед подписью
    if ris3:
        idx = paras.index(ris3)
        if idx > 0 and paras[idx - 1].text == '':
            _insert_image_in_para(paras[idx - 1], shots.step_response)

    # Рис.5: вставляем перед подписью
    if ris5 and shots.critical:
        _insert_image_before(doc, ris5, shots.critical)

    # Рис.6: вставляем перед подписью
    if ris6 and shots.bode:
        _insert_image_before(doc, ris6, shots.bode)

    doc.save(str(out_path))
    print(f'[report] Готово: {out_path}')
    return str(out_path)
