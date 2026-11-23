from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class GostProfile:
    """Проверяемый профиль оформления.

    База — ГОСТ 7.32-2017. Кегль 14 пт и жёсткие поля взяты из ТЗ заказчика
    (вузовская методичка может отличаться; тогда меняются значения здесь).
    """

    name: str = "ГОСТ 7.32-2017 (профиль вуза: 14 пт)"
    standard_refs: str = "ГОСТ 7.32-2017; ГОСТ Р 7.0.100-2018; ГОСТ Р 7.0.5-2008"

    page_width_mm: float = 210.0
    page_height_mm: float = 297.0
    margin_left_mm: float = 30.0
    margin_right_mm: float = 15.0
    margin_top_mm: float = 20.0
    margin_bottom_mm: float = 20.0
    margin_tolerance_mm: float = 0.6

    body_font: str = "Times New Roman"
    body_font_pt: float = 14.0
    table_min_font_pt: float = 12.0
    font_tolerance_pt: float = 0.51

    line_spacing: float = 1.5
    line_spacing_tolerance: float = 0.03
    extra_spacing_tolerance_pt: float = 0.5

    first_line_indent_cm: float = 1.25
    indent_tolerance_cm: float = 0.10
    side_indent_tolerance_cm: float = 0.10

    body_alignment: str = "both"

    accepted_font_names: tuple[str, ...] = (
        "Times New Roman",
        "Times New Roman PSMT",
        "TimesNewRoman",
    )

    @property
    def normalized_font_names(self) -> set[str]:
        return {_norm_font_name(name) for name in self.accepted_font_names}


def _norm_font_name(name: str | None) -> str:
    if not name:
        return ""
    return "".join(ch for ch in name.casefold() if ch.isalnum())


PROFILE = GostProfile()
