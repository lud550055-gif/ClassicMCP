# ============================================================
#  calculator.py — все теоретические расчёты
# ============================================================
from dataclasses import dataclass, field
from typing import List
import sympy as sp


@dataclass
class CalcResults:
    variant:    int   = 0

    # Передаточные функции (строки для вставки в отчёт)
    WP_str:    str = ""   # WP(s) разомкнутой
    Phi_str:   str = ""   # Ф(s) замкнутой по управлению
    PhiE_str:  str = ""   # Фe(s) по ошибке

    # Характеристический полином
    char_poly_str: str = ""
    char_coeffs:   List[float] = field(default_factory=list)  # [a0, a1, a2, a3]

    # Устойчивость
    hurwitz_stable:     bool  = False
    hurwitz_detail:     str   = ""    # строка с пояснением
    K1_critical:        float = 0.0  # критическое значение K1
    K_loop_critical:    float = 0.0  # Ккр = K1кр * K3 * K4 * K5

    # Ошибки
    e_ust_step:  str = ""   # при f(t)=1(t)
    e_ust_ramp:  str = ""   # при f(t)=0.1t

    # CLASSiC-таблицы (коэффициенты ПФ для вставки в отчёт)
    WP_classic:  dict = field(default_factory=dict)   # {num: [...], den: [...]}
    Phi_classic: dict = field(default_factory=dict)
    PhiE_classic: dict = field(default_factory=dict)

    # Параметры блоков (для заполнения текстовой таблицы блоков в отчёте)
    K1: float = 0.0
    K3: float = 0.0
    T3: float = 0.0
    K4: float = 0.0
    T4: float = 0.0
    K5: float = 0.0


def calculate(params) -> CalcResults:
    """
    Принимает VariantParams, возвращает CalcResults.
    Все вычисления через SymPy — точная алгебра.
    """
    s = sp.Symbol('s')
    r = CalcResults(variant=params.variant)

    # ── Передаточные функции блоков ───────────────────────────────────────────
    W2 = sp.Rational(params.K1).limit_denominator(10000) \
         if hasattr(params.K1, 'as_integer_ratio') \
         else sp.nsimplify(params.K1, rational=True)

    W3 = sp.nsimplify(params.K3, rational=True) / \
         (sp.nsimplify(params.T3, rational=True) * s + 1)

    W4 = sp.nsimplify(params.K4, rational=True) / \
         (sp.nsimplify(params.T4, rational=True) * s + 1)

    W5 = sp.nsimplify(params.K5, rational=True) / s

    # ── WP(s) — разомкнутая ПФ ───────────────────────────────────────────────
    WP = sp.together(W2 * W3 * W4 * W5)
    WP_num, WP_den = sp.fraction(WP)
    WP_num = sp.expand(WP_num)
    WP_den = sp.expand(WP_den)
    r.WP_str = f"WP(s) = {WP_num} / ({WP_den})"

    # ── Ф(s) — замкнутая по управлению ───────────────────────────────────────
    Phi = sp.together(WP / (1 + WP))
    Phi_num, Phi_den = sp.fraction(sp.simplify(Phi))
    Phi_num = sp.expand(Phi_num)
    Phi_den = sp.expand(Phi_den)
    r.Phi_str = f"Ф(s) = {Phi_num} / ({Phi_den})"

    # ── Фe(s) — ПФ по ошибке ─────────────────────────────────────────────────
    PhiE = sp.together(1 / (1 + WP))
    PhiE_num, PhiE_den = sp.fraction(sp.simplify(PhiE))
    PhiE_num = sp.expand(PhiE_num)
    PhiE_den = sp.expand(PhiE_den)
    r.PhiE_str = f"Фe(s) = ({PhiE_num}) / ({PhiE_den})"

    # ── Характеристический полином ────────────────────────────────────────────
    char_poly = sp.Poly(Phi_den, s)
    r.char_poly_str = str(Phi_den)
    r.char_coeffs = [float(c) for c in char_poly.all_coeffs()]

    # ── Критерий Гурвица ──────────────────────────────────────────────────────
    coeffs = r.char_coeffs
    n = len(coeffs)
    all_positive = all(c > 0 for c in coeffs)

    if n == 4:
        # a0*s^3 + a1*s^2 + a2*s + a3
        # Условие: a1*a2 > a0*a3
        a0, a1, a2, a3 = coeffs
        delta = a1 * a2 - a0 * a3
        r.hurwitz_stable = all_positive and delta > 0
        r.hurwitz_detail = (
            f"Коэффициенты: a0={a0}, a1={a1}, a2={a2}, a3={a3}\n"
            f"Все > 0: {all_positive}\n"
            f"Δ = a1·a2 - a0·a3 = {a1}·{a2} - {a0}·{a3} = {delta} {'> 0 ✓' if delta > 0 else '≤ 0 ✗'}\n"
            f"Система {'УСТОЙЧИВА' if r.hurwitz_stable else 'НЕУСТОЙЧИВА'}"
        )
    else:
        r.hurwitz_stable = all_positive
        r.hurwitz_detail = f"Все коэффициенты > 0: {all_positive}"

    # ── K1_кр (критическое усиление) ─────────────────────────────────────────
    K1_sym = sp.Symbol('K1', positive=True)
    WP_var = K1_sym * W3 * W4 * W5
    Phi_var = WP_var / (1 + WP_var)
    _, den_var = sp.fraction(sp.together(sp.simplify(Phi_var)))
    den_var = sp.expand(den_var)
    poly_var = sp.Poly(den_var, s)
    cv = poly_var.all_coeffs()

    if len(cv) == 4:
        a0v, a1v, a2v, a3v = cv
        # Граница: a1*a2 = a0*a3
        eq = sp.Eq(a1v * a2v, a0v * a3v)
        K1cr_list = sp.solve(eq, K1_sym)
        if K1cr_list:
            r.K1_critical = float(K1cr_list[0])
            r.K_loop_critical = r.K1_critical * float(
                sp.nsimplify(params.K3 * params.K4 * params.K5, rational=True)
            )

    # ── Установившиеся ошибки ─────────────────────────────────────────────────
    # e_уст = lim(s→0) s·Фe(s)·F(s)
    PhiE_expr = sp.together(1 / (1 + WP))

    # f(t)=1(t) → F(s)=1/s
    e_step = sp.limit(s * PhiE_expr * (1/s), s, 0)
    r.e_ust_step = str(sp.simplify(e_step))

    # f(t)=0.1t → F(s)=0.1/s²
    e_ramp = sp.limit(s * PhiE_expr * (sp.Rational(1,10) / s**2), s, 0)
    r.e_ust_ramp = str(sp.simplify(e_ramp))

    # ── CLASSiC-совместимые таблицы коэффициентов ────────────────────────────
    def poly_to_classic(num_expr, den_expr):
        """Раскладывает полиномы в список коэффициентов для таблицы CLASSiC."""
        num_poly = sp.Poly(sp.expand(num_expr), s) if sp.degree(num_expr, s) >= 0 \
                   else sp.Poly(num_expr, s)
        den_poly = sp.Poly(sp.expand(den_expr), s)
        # CLASSiC хранит коэффициенты от s^0 до s^n
        num_coeffs = list(reversed([float(c) for c in num_poly.all_coeffs()]))
        den_coeffs = list(reversed([float(c) for c in den_poly.all_coeffs()]))
        return {"num": num_coeffs, "den": den_coeffs}

    r.WP_classic  = poly_to_classic(WP_num, WP_den)
    r.Phi_classic = poly_to_classic(Phi_num, Phi_den)
    r.PhiE_classic = poly_to_classic(PhiE_num, PhiE_den)

    r.K1 = float(params.K1)
    r.K3 = float(params.K3); r.T3 = float(params.T3)
    r.K4 = float(params.K4); r.T4 = float(params.T4)
    r.K5 = float(params.K5)

    return r


def format_results(r: CalcResults) -> str:
    """Красивый вывод результатов для логирования."""
    lines = [
        f"{'='*60}",
        f" Вариант {r.variant}",
        f"{'='*60}",
        f"Задача 4: {r.WP_str}",
        f"Задача 5: {r.Phi_str}",
        f"Задача 6: {r.PhiE_str}",
        f"",
        f"Характеристический полином: {r.char_poly_str}",
        f"",
        f"Критерий Гурвица:",
        r.hurwitz_detail,
        f"",
        f"Критическое усиление: K1кр = {r.K1_critical}",
        f"Ккр (контура) = {r.K_loop_critical}",
        f"",
        f"Ошибки:",
        f"  f(t)=1(t)   → eуст = {r.e_ust_step}",
        f"  f(t)=0.1t   → eуст = {r.e_ust_ramp}",
        f"{'='*60}",
    ]
    return "\n".join(lines)


if __name__ == "__main__":
    from parser import parse_variant
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else "LP_v17.doc"
    params = parse_variant(path)
    results = calculate(params)
    print(format_results(results))
