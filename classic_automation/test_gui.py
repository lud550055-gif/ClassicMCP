"""Тест classic_gui.py на заведомо рабочем pr_stan.mdl из examples."""
from classic_gui import ClassicController

# Используем готовый бинарный MDL из папки CLASSiC — его CLASSiC точно откроет
MDL = r"C:\Users\lud50\Desktop\TU\classic\examples\pr_stan.mdl"

ctrl = ClassicController(
    mdl_path=MDL,
    output_dir=r"C:\Users\lud50\Desktop\TU\reports\test_shots",
    variant=0,
    K1_critical=0.0,
)

shots = ctrl.run_all()

print("\n=== Результат ===")
for field, path in shots.__dict__.items():
    status = "OK" if path else "пусто"
    print(f"  {field:20s}: {status}  {path}")
