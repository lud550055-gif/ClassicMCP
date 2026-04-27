"""Проверяем запуск CLASSiC с MDL-файлом в аргументе, ждём 12 сек."""
import time, subprocess
from PIL import ImageGrab
import pygetwindow as gw
import config

MDL = r"C:\Users\lud50\Desktop\TU\models\var17.mdl"

def find_classic():
    wins = [w for w in gw.getAllWindows()
            if "classic" in w.title.lower() and w.width > 50]
    return max(wins, key=lambda w: w.width * w.height) if wins else None

def shot(name):
    w = find_classic()
    if w:
        img = ImageGrab.grab(bbox=(w.left, w.top, w.left+w.width, w.top+w.height))
        img.save(f"dbg_{name}.png")
        print(f"  [shot] {name}  title={w.title!r}")
    else:
        print(f"  [!] окно не найдено для {name}")

print(f"Запускаем с аргументом: {MDL}")
proc = subprocess.Popen([config.CLASSIC_EXE, MDL])

for i in [3, 6, 9, 12]:
    time.sleep(3)
    shot(f"t{i}s")

proc.terminate()
print("Готово.")
