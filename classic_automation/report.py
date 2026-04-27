# ============================================================
#  report.py — заполняет шаблон .docx результатами
# ============================================================
"""
Принимает путь к шаблону LP_vXX.doc, результаты расчётов
и скриншоты, и создаёт заполненный отчёт в output_dir.
"""
import os
import re
import subprocess
import shutil
from pathlib import Path
from PIL import Image

from calculator import CalcResults
from classic_gui import Screenshots


def _convert_to_docx(doc_path: str, work_dir: str) -> str:
    """Конвертирует .doc → .docx если нужно."""
    src = Path(doc_path)
    if src.suffix.lower() == ".docx":
        dst = Path(work_dir) / src.name
        shutil.copy(src, dst)
        return str(dst)

    dst = Path(work_dir) / (src.stem + ".docx")
    # Используем LibreOffice
    cmd = ["soffice", "--headless", "--convert-to", "docx",
           str(src), "--outdir", work_dir]
    subprocess.run(cmd, capture_output=True, timeout=30)
    if dst.exists():
        return str(dst)
    raise FileNotFoundError(f"Не удалось конвертировать {doc_path}")


def _unpack(docx_path: str, out_dir: str):
    """Распаковывает .docx."""
    import zipfile
    with zipfile.ZipFile(docx_path, 'r') as z:
        z.extractall(out_dir)


def _pack(unpacked_dir: str, output_docx: str):
    """Упаковывает директорию обратно в .docx."""
    import zipfile
    with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(unpacked_dir):
            for file in files:
                fp = os.path.join(root, file)
                arcname = os.path.relpath(fp, unpacked_dir)
                z.write(fp, arcname)


def _img_emu(img_path: str, max_width_cm: float = 15.0):
    """Возвращает (cx, cy) в EMU для вставки изображения."""
    EMU_PER_CM = 360000
    MAX_W = int(max_width_cm * EMU_PER_CM)
    img = Image.open(img_path)
    w, h = img.size
    # Масштаб по 96 dpi
    emu_w = int(w * 914400 / 96)
    emu_h = int(h * 914400 / 96)
    if emu_w > MAX_W:
        scale = MAX_W / emu_w
        emu_w = MAX_W
        emu_h = int(emu_h * scale)
    return emu_w, emu_h


def _make_inline_img_xml(r_id: str, img_path: str, pic_id: int, name: str) -> str:
    """Генерирует XML для вставки inline-изображения."""
    cx, cy = _img_emu(img_path)
    return (
        f'<w:r><w:drawing>'
        f'<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        f' distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{pic_id}" name="{name}"/>'
        f'<wp:cNvGraphicFramePr>'
        f'<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        f'</wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="{pic_id}" name="{name}"/>'
        f'<pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{r_id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        f'</pic:pic></a:graphicData></a:graphic>'
        f'</wp:inline></w:drawing></w:r>'
    )


def fill_report(
    template_path: str,
    output_dir:    str,
    calc:          CalcResults,
    shots:         Screenshots,
) -> str:
    """
    Главная функция.
    Возвращает путь к готовому отчёту.
    """
    import tempfile
    work = tempfile.mkdtemp(prefix=f"report_v{calc.variant}_")
    variant_str = f"{calc.variant:02d}"

    # 1. Конвертируем шаблон
    docx_path = _convert_to_docx(template_path, work)

    # 2. Распаковываем
    unpacked = os.path.join(work, "unpacked")
    _unpack(docx_path, unpacked)

    doc_xml_path = os.path.join(unpacked, "word", "document.xml")
    rels_path    = os.path.join(unpacked, "word", "_rels", "document.xml.rels")
    media_dir    = os.path.join(unpacked, "word", "media")
    os.makedirs(media_dir, exist_ok=True)

    doc_xml = Path(doc_xml_path).read_text(encoding="utf-8")
    rels_xml = Path(rels_path).read_text(encoding="utf-8")

    # 3. Добавляем изображения и rId
    image_inserts: dict[str, tuple[str, str]] = {}  # name → (r_id, img_path)
    pic_id = 100
    r_id_counter = _max_rid(rels_xml) + 1

    img_tasks = [
        ("schema",          shots.schema),
        ("text_form",       shots.text_form),
        ("characteristics", shots.characteristics),
        ("root_locus",      shots.root_locus),
        ("step_response",   shots.step_response),
        ("bode",            shots.bode),
        ("tf_panel",        shots.tf_panel),
        ("critical",        shots.critical),
    ]

    for img_name, img_path in img_tasks:
        if not img_path or not Path(img_path).exists():
            continue
        ext = Path(img_path).suffix.lower()
        media_filename = f"auto_{img_name}{ext}"
        shutil.copy(img_path, os.path.join(media_dir, media_filename))
        r_id = f"rIdAuto{r_id_counter}"
        r_id_counter += 1
        rels_xml = rels_xml.replace(
            "</Relationships>",
            f'<Relationship Id="{r_id}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="media/{media_filename}"/>\n</Relationships>'
        )
        image_inserts[img_name] = (r_id, img_path)
        pic_id += 1

    # 4. Текстовые замены в document.xml
    replacements = _build_text_replacements(calc)
    for old, new in replacements.items():
        doc_xml = doc_xml.replace(old, new)

    # 5. Вставка скриншотов вместо плейсхолдеров вида [IMG:schema]
    for img_name, (r_id, img_path) in image_inserts.items():
        placeholder = f"[IMG:{img_name}]"
        img_xml = _make_inline_img_xml(r_id, img_path, pic_id, img_name)
        # Оборачиваем в параграф
        para_xml = f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr>{img_xml}</w:p>'
        doc_xml = doc_xml.replace(
            f'<w:t>{placeholder}</w:t>', img_xml
        )
        pic_id += 1

    # 6. Записываем изменённые файлы
    Path(doc_xml_path).write_text(doc_xml, encoding="utf-8")
    Path(rels_path).write_text(rels_xml, encoding="utf-8")

    # 7. Упаковываем
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    output_path = os.path.join(output_dir, f"LP_v{variant_str}_report.docx")
    _pack(unpacked, output_path)

    # Очистка
    shutil.rmtree(work, ignore_errors=True)

    print(f"[report] Отчёт сохранён: {output_path}")
    return output_path


def _max_rid(rels_xml: str) -> int:
    """Находит максимальный числовой Id в .rels файле."""
    ids = re.findall(r'Id="rId(\d+)"', rels_xml)
    return max((int(x) for x in ids), default=10)


def _build_text_replacements(calc: CalcResults) -> dict:
    """
    Строит словарь замен плейсхолдер→значение для document.xml.
    Плейсхолдеры — те что уже есть в шаблоне (…) или именованные [TASK4_WP] и т.д.
    """
    # Дробь для eуст
    e_ramp = calc.e_ust_ramp
    e_step = calc.e_ust_step

    # Коэффициенты хар. полинома
    c = calc.char_coeffs
    char_str = " + ".join(
        f"{v:.4g}·s^{i}" if i > 0 else f"{v:.4g}"
        for i, v in enumerate(reversed(c))
    )

    return {
        # Явные плейсхолдеры (если используются в шаблоне)
        "[TASK4_WP]":     calc.WP_str,
        "[TASK5_PHI]":    calc.Phi_str,
        "[TASK6_PHIE]":   calc.PhiE_str,
        "[CHAR_POLY]":    calc.char_poly_str,
        "[HURWITZ]":      calc.hurwitz_detail,
        "[K1CR]":         str(calc.K1_critical),
        "[KCR]":          str(calc.K_loop_critical),
        "[EUST_STEP]":    e_step,
        "[EUST_RAMP]":    e_ramp,
        "[STABLE]":       "устойчива" if calc.hurwitz_stable else "неустойчива",

        # Варианты с кириллицей (на случай разных шаблонов)
        "[ЗАД4_WP]":      calc.WP_str,
        "[ЗАД5_Ф]":       calc.Phi_str,
        "[ЗАД6_Фe]":      calc.PhiE_str,
        "[К1кр]":         str(calc.K1_critical),
        "[Ккр]":          str(calc.K_loop_critical),
        "[eуст_step]":    e_step,
        "[eуст_ramp]":    e_ramp,
    }
