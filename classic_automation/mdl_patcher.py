# ============================================================
#  mdl_patcher.py — патчит параметры блоков в бинарном MDL
# ============================================================
"""
Стратегия: берём template.mdl (созданный вручную в CLASSiC с параметрами
варианта 17), ищем известные float32/float64 значения и заменяем их.

Workflow:
  1. Пользователь создаёт в CLASSiC модель с параметрами варианта 17 и
     сохраняет как template.mdl в папку CLASSiC.
  2. Вызываем locate_params(template_mdl) — ищет все оффсеты параметров.
  3. Вызываем patch_mdl(template_mdl, new_params, output_path) для каждого варианта.
"""
import struct
import shutil
from pathlib import Path
from typing import Optional


# Параметры варианта-шаблона (вариант 17)
TEMPLATE_PARAMS = {
    "K1": 100.0,
    "K3": 2.5,
    "T3": 0.5,
    "K4": 0.2,
    "T4": 0.05,
    "K5": 0.05,
}


def _pack32(v: float) -> bytes:
    return struct.pack('<f', v)


def _pack64(v: float) -> bytes:
    return struct.pack('<d', v)


def _find_all(data: bytes, pattern: bytes) -> list[int]:
    """Возвращает список всех позиций pattern в data."""
    offsets = []
    start = 0
    while True:
        idx = data.find(pattern, start)
        if idx == -1:
            break
        offsets.append(idx)
        start = idx + 1
    return offsets


def locate_params(template_path: str, template_params: dict = None) -> dict:
    """
    Ищет все оффсеты параметров template_params в бинарном MDL.
    Возвращает: {param_name: [offset, ...]} для float32 и float64.

    Вызовите эту функцию ОДИН РАЗ после создания template.mdl,
    затем сохраните offsets и используйте их в patch_mdl.
    """
    if template_params is None:
        template_params = TEMPLATE_PARAMS

    data = Path(template_path).read_bytes()
    result = {}

    for name, val in template_params.items():
        f32 = _pack32(val)
        f64 = _pack64(val)
        offs32 = _find_all(data, f32)
        offs64 = _find_all(data, f64)
        result[name] = {
            "value":  val,
            "float32": offs32,
            "float64": offs64,
        }
        print(f"  {name}={val}: float32 {f32.hex()} at {offs32}; float64 at {offs64}")

    return result


def patch_mdl(
    template_path: str,
    new_params: dict,
    output_path: str,
    offsets: Optional[dict] = None,
    template_params: dict = None,
) -> str:
    """
    Создаёт новый MDL из template_path, заменяя параметры.

    template_path: путь к template.mdl (бинарный OLE MDL CLASSiC).
    new_params:    {param_name: new_value} — новые значения.
    output_path:   куда сохранить результат.
    offsets:       кэш из locate_params (если None — вычисляется автоматически).
    """
    if template_params is None:
        template_params = TEMPLATE_PARAMS

    if offsets is None:
        offsets = locate_params(template_path, template_params)

    data = bytearray(Path(template_path).read_bytes())

    for name, new_val in new_params.items():
        if name not in offsets:
            print(f"  [warn] {name} не найден в offsets, пропускаем")
            continue
        old_val = template_params[name]
        info = offsets[name]

        # Заменяем float32
        old32 = _pack32(old_val)
        new32 = _pack32(new_val)
        for off in info["float32"]:
            data[off:off + 4] = new32
            print(f"  patch float32 {name}: off=0x{off:x} {old32.hex()} -> {new32.hex()}")

        # Заменяем float64
        old64 = _pack64(old_val)
        new64 = _pack64(new_val)
        for off in info["float64"]:
            data[off:off + 8] = new64
            print(f"  patch float64 {name}: off=0x{off:x} {old64.hex()} -> {new64.hex()}")

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    Path(output_path).write_bytes(bytes(data))
    print(f"[patch] Записан: {output_path}")
    return output_path


def write_mdl(params, output_path: str) -> str:
    """
    Главная функция — API-совместима с оригинальным mdl_writer.write_mdl.
    Требует наличия template.mdl рядом с CLASSiC.exe.

    params: VariantParams с полями K1, K3, T3, K4, T4, K5, variant.
    """
    import config

    template = Path(config.CLASSIC_EXE).parent / "template.mdl"
    if not template.exists():
        raise FileNotFoundError(
            f"Не найден {template}\n"
            "Создайте шаблон вручную:\n"
            "  1. Откройте CLASSiC → Файл → Новый\n"
            "  2. Создайте модель с 5 блоками (параметры варианта 17):\n"
            "       K1=100, K3=2.5, T3=0.5, K4=0.2, T4=0.05, K5=0.05\n"
            "  3. Сохраните как template.mdl в папку CLASSiC\n"
            "  4. Снова запустите скрипт"
        )

    new_params = {
        "K1": params.K1,
        "K3": params.K3,
        "T3": params.T3,
        "K4": params.K4,
        "T4": params.T4,
        "K5": params.K5,
    }

    print(f"[mdl_patcher] Патчим template.mdl для варианта {params.variant}...")
    return patch_mdl(str(template), new_params, output_path)


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python mdl_patcher.py <template.mdl>")
        sys.exit(1)
    print("Поиск параметров в шаблоне...")
    offsets = locate_params(sys.argv[1])
    print("\nРезультат:")
    for name, info in offsets.items():
        print(f"  {name}: float32={info['float32']}, float64={info['float64']}")
