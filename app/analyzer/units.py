from __future__ import annotations

EMU_PER_MM = 36_000.0
TWIPS_PER_CM = 566.9291338582677
TWIPS_PER_MM = 56.69291338582677
PT_PER_CM = 28.346456692913385


def emu_to_mm(value: float | None) -> float | None:
    if value is None:
        return None
    return float(value) / EMU_PER_MM


def twips_to_cm(value: float | None) -> float | None:
    if value is None:
        return None
    return float(value) / TWIPS_PER_CM


def twips_to_pt(value: float | None) -> float | None:
    if value is None:
        return None
    return float(value) / 20.0


def half_points_to_pt(value: float | None) -> float | None:
    if value is None:
        return None
    return float(value) / 2.0


def fmt_num(value: float, digits: int = 2) -> str:
    text = f"{value:.{digits}f}".rstrip("0").rstrip(".")
    return text.replace(".", ",")


def fmt_mm(value: float) -> str:
    return f"{fmt_num(value, 1)} мм"


def fmt_cm(value: float) -> str:
    return f"{fmt_num(value, 2)} см"


def fmt_pt(value: float) -> str:
    return f"{fmt_num(value, 1)} пт"


def normalize_ws(text: str) -> str:
    return " ".join(text.replace("\u00a0", " ").replace("\u202f", " ").split())


def shorten(text: str, limit: int = 120) -> str:
    text = normalize_ws(text)
    if len(text) <= limit:
        return text
    return text[: limit - 1].rstrip() + "…"
