# ============================================================
#  mdl_writer.py — генерирует .mdl файл для CLASSiC 3.2
# ============================================================
"""
Формат MDL (текстовый, определён по реверс-инжинирингу CLASSiC 3.2):

[ClBlocks]
Count=5

[Block1]
Name=1
Type=0          ; 0=обычный блок
Inputs=-        ; список входных связей через запятую (минус = инверсия)
Num=1           ; числитель (степень 0)
Den0=1          ; знаменатель степень 0
Den1=...        ; знаменатель степень 1 (если есть)
...

[Block5]
...
IsOutput=1      ; признак блока-выхода

[ClLinks]
Count=5
Link1=1,2       ; связь от блока 1 к блоку 2
...
"""
from pathlib import Path


def write_mdl(params, output_path: str) -> str:
    """
    Генерирует MDL файл для варианта params.
    Возвращает путь к созданному файлу.

    Если в папке CLASSiC есть template.mdl — использует mdl_patcher (бинарный патч).
    Иначе — записывает текстовый MDL (CLASSiC не откроет, только для совместимости).

    Структура:
      Блок 1 (вход/сумматор): W=1, минус на входе обратной связи
      Блок 2: W2 = K1
      Блок 3: W3 = K3/(T3*s+1)
      Блок 4: W4 = K4/(T4*s+1)
      Блок 5 (выход): W5 = K5/s
    """
    import config
    template = Path(config.CLASSIC_EXE).parent / "template.mdl"
    if template.exists():
        from mdl_patcher import write_mdl as patcher_write
        return patcher_write(params, output_path)

    lines = []

    lines.append("[ClMain]")
    lines.append("Window=100,100,900,650")
    lines.append("")

    # ── блоки ────────────────────────────────────────────────────────────────
    lines.append("[ClBlocks]")
    lines.append("Count=5")
    lines.append("")

    # Блок 1 — сумматор (вход)
    lines.append("[Block1]")
    lines.append("Name=1")
    lines.append("Type=0")
    lines.append("IsInput=1")
    lines.append("Inputs=5-")     # обратная связь от блока 5 со знаком минус
    lines.append("Num=1")
    lines.append("Den0=1")
    lines.append("PosX=80")
    lines.append("PosY=250")
    lines.append("")

    # Блок 2 — W2 = K1
    lines.append("[Block2]")
    lines.append("Name=2")
    lines.append("Type=0")
    lines.append(f"Num={_fmt(params.K1)}")
    lines.append("Den0=1")
    lines.append("PosX=220")
    lines.append("PosY=250")
    lines.append("")

    # Блок 3 — W3 = K3/(T3*s+1)
    lines.append("[Block3]")
    lines.append("Name=3")
    lines.append("Type=0")
    lines.append(f"Num={_fmt(params.K3)}")
    lines.append("Den0=1")
    lines.append(f"Den1={_fmt(params.T3)}")
    lines.append("PosX=380")
    lines.append("PosY=250")
    lines.append("")

    # Блок 4 — W4 = K4/(T4*s+1)
    lines.append("[Block4]")
    lines.append("Name=4")
    lines.append("Type=0")
    lines.append(f"Num={_fmt(params.K4)}")
    lines.append("Den0=1")
    lines.append(f"Den1={_fmt(params.T4)}")
    lines.append("PosX=540")
    lines.append("PosY=250")
    lines.append("")

    # Блок 5 — W5 = K5/s (выход)
    lines.append("[Block5]")
    lines.append("Name=5")
    lines.append("Type=0")
    lines.append("IsOutput=1")
    lines.append(f"Num={_fmt(params.K5)}")
    lines.append("Den0=1e-30")    # CLASSiC не принимает 0, используем ~0
    lines.append("Den1=1")
    lines.append("PosX=700")
    lines.append("PosY=250")
    lines.append("")

    # ── связи ────────────────────────────────────────────────────────────────
    lines.append("[ClLinks]")
    lines.append("Count=5")
    lines.append("Link1=1,2")
    lines.append("Link2=2,3")
    lines.append("Link3=3,4")
    lines.append("Link4=4,5")
    lines.append("Link5=5,1")     # обратная связь
    lines.append("")

    content = "\r\n".join(lines)
    Path(output_path).write_text(content, encoding="cp1251")
    return output_path


def _fmt(val: float) -> str:
    """Форматирует число для MDL (без лишних нулей)."""
    if val == int(val):
        return str(int(val))
    return str(val)


if __name__ == "__main__":
    from parser import parse_variant, VariantParams
    import sys, os
    p = VariantParams(variant=17, K1=100, K3=2.5, T3=0.5, K4=0.2, T4=0.05, K5=0.05)
    path = write_mdl(p, "test_v17.mdl")
    print(f"Создан MDL: {path}")
    print(Path(path).read_text(encoding="cp1251"))
