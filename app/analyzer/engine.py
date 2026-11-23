from __future__ import annotations

import re
import time
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from docx.opc.exceptions import PackageNotFoundError
from docx.oxml.ns import qn

from .document import Block, DocxModel
from .models import (
    AnalysisError,
    AnalysisResult,
    Issue,
    Severity,
    SkippedCheck,
)
from .profile import PROFILE, GostProfile
from .units import (
    fmt_cm,
    fmt_mm,
    fmt_num,
    fmt_pt,
    normalize_ws,
    shorten,
)

STRUCTURAL_RE = re.compile(
    r"^(?:"
    r"СОДЕРЖАНИЕ|ОГЛАВЛЕНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|РЕФЕРАТ|"
    r"СПИСОК\s+(?:ИСПОЛЬЗОВАННЫХ\s+ИСТОЧНИКОВ(?:\s+И\s+ЛИТЕРАТУРЫ)?|"
    r"ИСПОЛЬЗОВАННОЙ\s+ЛИТЕРАТУРЫ|ИСПОЛЬЗУЕМОЙ\s+ЛИТЕРАТУРЫ|"
    r"ИСПОЛЬЗУЕМЫХ\s+ИСТОЧНИКОВ(?:\s+И\s+ЛИТЕРАТУРЫ)?|"
    r"ЛИТЕРАТУРНЫХ\s+ИСТОЧНИКОВ|ИСТОЧНИКОВ\s+И\s+ЛИТЕРАТУРЫ|ЛИТЕРАТУРЫ|ИСТОЧНИКОВ)|"
    r"БИБЛИОГРАФИЧЕСКИЙ\s+СПИСОК|БИБЛИОГРАФИЯ|"
    r"ПРИЛОЖЕНИЕ(?:\s+[А-ЯЁA-Z0-9])?|"
    r"ПЕРЕЧЕНЬ\s+СОКРАЩЕНИЙ(?:\s+И\s+ОБОЗНАЧЕНИЙ)?|"
    r"ТЕРМИНЫ\s+И\s+ОПРЕДЕЛЕНИЯ"
    r")$"
)
BIBLIOGRAPHY_RE = re.compile(
    r"^(?:СПИСОК\s+(?:ИСПОЛЬЗОВАННЫХ\s+ИСТОЧНИКОВ(?:\s+И\s+ЛИТЕРАТУРЫ)?|"
    r"ИСПОЛЬЗОВАННОЙ\s+ЛИТЕРАТУРЫ|ИСПОЛЬЗУЕМОЙ\s+ЛИТЕРАТУРЫ|"
    r"ИСПОЛЬЗУЕМЫХ\s+ИСТОЧНИКОВ(?:\s+И\s+ЛИТЕРАТУРЫ)?|"
    r"ЛИТЕРАТУРНЫХ\s+ИСТОЧНИКОВ|ИСТОЧНИКОВ\s+И\s+ЛИТЕРАТУРЫ|ЛИТЕРАТУРЫ|ИСТОЧНИКОВ)|"
    r"БИБЛИОГРАФИЧЕСКИЙ\s+СПИСОК|БИБЛИОГРАФИЯ)$"
)
NUMBERED_HEADING_RE = re.compile(r"^(\d+(?:\.\d+)*)\.?\s+([A-ZА-ЯЁ][^\n]{0,190})$")
FIGURE_CAPTION_RE = re.compile(
    r"^Рисунок\s+([А-ЯЁA-Z]\.)?(\d+(?:\.\d+)?)\s*([—–-])\s*(.+)$"
)
TABLE_CAPTION_RE = re.compile(
    r"^Таблица\s+([А-ЯЁA-Z]\.)?(\d+(?:\.\d+)?)\s*([—–-])\s*(.+)$"
)
CITATION_RE = re.compile(r"\[([^\]\n]{1,80})\]")
YEAR_RE = re.compile(r"\b(?:19|20)\d{2}\b")
ACCESS_DATE_RE = re.compile(r"дата\s+обращения\s*:\s*\d{1,2}\.\d{1,2}\.\d{4}", re.IGNORECASE)


def _norm_font(name: str | None) -> str:
    if not name:
        return ""
    return "".join(ch for ch in name.casefold() if ch.isalnum())


def _is_toc_entry(text: str) -> bool:
    if re.search(r"\.{4,}\s*\d+\s*$", text):
        return True
    return bool("\t" in text and re.search(r"\d+\s*$", text))


def _alignment_label(value: str | None) -> str:
    return {
        "both": "по ширине",
        "center": "по центру",
        "left": "по левому краю",
        "start": "по левому краю",
        "right": "по правому краю",
        "end": "по правому краю",
        "distribute": "распределённое",
        None: "по левому краю (по умолчанию)",
    }.get(value, value or "не задано")


@dataclass
class HeadingInfo:
    index: int
    text: str
    clean: str
    level: int
    structural: bool
    numbered: bool
    style_based: bool
    outline_level: int | None
    all_caps: bool


@dataclass
class CaptionInfo:
    index: int
    number: str
    numeric: int | None
    structure_ok: bool
    separator: str
    text: str


class GostAnalyzer:
    def __init__(
        self,
        path: str | Path,
        *,
        file_name: str | None = None,
        profile: GostProfile = PROFILE,
        max_issues: int = 1000,
    ) -> None:
        self.path = Path(path)
        self.file_name = file_name or self.path.name
        self.profile = profile
        self.max_issues = max_issues
        self.model = DocxModel(self.path)
        self.issues: list[Issue] = []

        self.headings: list[HeadingInfo] = []
        self.heading_by_index: dict[int, HeadingInfo] = {}
        self.heading_like: list[tuple[int, str]] = []
        self.unstyled_structural: list[tuple[int, str]] = []
        self.body_start = 0

        self.table_caption_indices: set[int] = set()
        self.figure_caption_indices: set[int] = set()
        self.table_captions: list[CaptionInfo] = []
        self.figure_captions: list[CaptionInfo] = []

        self.table_count = 0
        self.figure_count = 0
        self.source_count = 0
        self.citation_count = 0
        self.bibliography_heading_index: int | None = None
        self.bibliography_indices: set[int] = set()
        self.bibliography_entries: list[Block] = []

    # ------------------------------------------------------------------ public

    def analyze(self) -> AnalysisResult:
        started = time.perf_counter()
        self._detect_structure()
        self._scan_figures()
        self._scan_tables()
        self._scan_bibliography()
        self._check_page_setup()
        self._check_page_numbers()
        self._check_headings()
        self._check_fonts()
        self._check_paragraphs()
        self._check_citations()
        self._check_required_sections()

        ordered = sorted(self.issues, key=Issue.sort_key)
        truncated = 0
        if len(ordered) > self.max_issues:
            truncated = len(ordered) - self.max_issues
            ordered = ordered[: self.max_issues]

        skipped = [
            SkippedCheck(
                "Точные номера страниц",
                "DOCX не хранит готовое разбиение на страницы. Место ошибки указано номером абзаца и цитатой.",
            ),
            SkippedCheck(
                "Сноски, надписи и колонтитулы",
                "Проверяется основной текст и нижний колонтитул с номером страницы. Текст внутри фигур, надписей и сносок не анализируется.",
            ),
            SkippedCheck(
                "Формулы и приложения",
                "Их нумерация и оформление не входят в согласованный список проверок.",
            ),
            SkippedCheck(
                "Смысл, орфография, пунктуация, антиплагиат",
                "Это не входит в проверку оформления по ГОСТ.",
            ),
            SkippedCheck(
                "Индивидуальные требования методички",
                "Использован профиль ГОСТ 7.32-2017. Методичка кафедры может задавать другие поля, кегль или правила заголовков.",
            ),
        ]
        if self.body_start > 0:
            skipped.insert(
                0,
                SkippedCheck(
                    "Титульный лист и содержание",
                    "Не проверяются как основной текст, чтобы не выдавать ложные ошибки на титуле и автособираемом содержании.",
                ),
            )

        paragraphs = sum(
            1
            for block in self.model.blocks
            if block.paragraph is not None and normalize_ws(block.text)
        )

        return AnalysisResult(
            file_name=self.file_name,
            standard=self.profile.name,
            issues=ordered,
            skipped=skipped,
            paragraph_count=paragraphs,
            table_count=self.table_count,
            figure_count=self.figure_count,
            source_count=self.source_count,
            citation_count=self.citation_count,
            duration_sec=time.perf_counter() - started,
            truncated_count=truncated,
        )

    # --------------------------------------------------------------- structure

    def _detect_structure(self) -> None:
        for block in self.model.blocks:
            if block.paragraph is None:
                continue
            info = self._heading_info(block)
            if info is not None:
                self.headings.append(info)
                self.heading_by_index[block.index] = info

        # Начало основного текста: введение. Ищем по тексту, а не только по
        # стилю: иначе криво оформленное введение сдвинет начало проверки.
        intro = next(
            (
                block.index
                for block in self.model.blocks
                if block.paragraph is not None
                and normalize_ws(block.text).rstrip(":").strip() == "ВВЕДЕНИЕ"
            ),
            None,
        )
        if intro is not None:
            self.body_start = intro
        else:
            toc = next(
                (
                    h.index
                    for h in self.headings
                    if h.clean in {"СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ"}
                ),
                None,
            )
            if toc is not None:
                self.body_start = self._first_heading_after_toc(toc)
            if self.body_start == 0:
                self.body_start = self._first_content_after_break()

        # Структурные заголовки, оформленные обычным текстом.
        for block in self.model.blocks:
            if block.paragraph is None:
                continue
            if block.index in self.heading_by_index:
                continue
            clean = normalize_ws(block.text).rstrip(":").strip()
            if clean and STRUCTURAL_RE.match(clean):
                self.unstyled_structural.append((block.index, clean))

        # Заголовки, оформленные обычным текстом, но похожие на нумерованные
        # разделы. Это только предупреждение, чтобы не ломать демо.
        bibliography = next(
            (h.index for h in self.headings if BIBLIOGRAPHY_RE.match(h.clean)),
            None,
        )
        for block in self.model.blocks:
            if block.paragraph is None or block.index < self.body_start:
                continue
            if block.index in self.heading_by_index:
                continue
            if bibliography is not None and block.index > bibliography:
                continue
            text = normalize_ws(block.text)
            match = NUMBERED_HEADING_RE.match(text)
            if not match or len(text) > 120:
                continue
            if re.search(r"[.;!?]$", text):
                continue
            if _is_toc_entry(text):
                continue
            ppr = self.model.styles.paragraph_format(block.paragraph)
            if ppr.has_numbering or len(text.split()) > 12:
                continue
            next_block = self._next_nonempty(block.index)
            if next_block is None or len(normalize_ws(next_block.text)) < 80:
                continue
            self.heading_like.append((block.index, text))

    def _first_content_after_break(self) -> int:
        for block in self.model.blocks:
            if block.paragraph is None:
                continue
            if self.model.paragraph_has_page_break(block.paragraph) or self.model.paragraph_has_section_break(
                block.paragraph
            ):
                for next_block in self.model.blocks[block.index + 1 :]:
                    if normalize_ws(next_block.text):
                        return next_block.index
        return 0

    def _first_heading_after_toc(self, toc_index: int) -> int:
        for block in self.model.blocks[toc_index + 1 :]:
            if block.paragraph is None:
                continue
            text = normalize_ws(block.text)
            if not text or _is_toc_entry(text):
                continue
            info = self.heading_by_index.get(block.index)
            if info is not None:
                return block.index
        return 0

    def _heading_info(self, block: Block) -> HeadingInfo | None:
        if block.paragraph is None:
            return None
        text = normalize_ws(block.text)
        if not text:
            return None
        clean = text.rstrip(":").strip()
        ppr = self.model.styles.paragraph_format(block.paragraph)
        style_name = ppr.style_name.casefold()
        style_id = ppr.style_id.casefold()
        has_outline = ppr.outline_level is not None and ppr.outline_level < 9
        style_based = (
            has_outline
            or "heading" in style_name
            or "заголовок" in style_name
            or bool(re.fullmatch(r"h[1-6]", style_id))
        )
        structural = bool(STRUCTURAL_RE.match(clean))
        numbered_match = NUMBERED_HEADING_RE.match(text)
        numbered = bool(numbered_match)
        if not (structural or style_based):
            return None
        if len(text) > 260:
            return None

        if structural:
            level = 1
        elif numbered_match:
            level = min(numbered_match.group(1).count(".") + 1, 6)
        elif has_outline:
            level = min(ppr.outline_level + 1, 9)
        else:
            level = 1

        letters = [ch for ch in text if ch.isalpha()]
        all_caps = bool(letters) and all(ch.isupper() for ch in letters)
        return HeadingInfo(
            index=block.index,
            text=text,
            clean=clean,
            level=level,
            structural=structural,
            numbered=numbered,
            style_based=style_based,
            outline_level=ppr.outline_level,
            all_caps=all_caps,
        )

    # ------------------------------------------------------------------ helpers

    def _next_nonempty(self, start: int, limit: int = 3) -> Block | None:
        for block in self.model.blocks[start + 1 : start + 1 + limit]:
            if normalize_ws(block.text):
                return block
        return None

    def _previous_nonempty(self, start: int, limit: int = 3) -> Block | None:
        for block in reversed(self.model.blocks[max(0, start - limit) : start]):
            if normalize_ws(block.text):
                return block
        return None

    def _add(
        self,
        *,
        rule_id: str,
        category: str,
        severity: Severity,
        title: str,
        location: str,
        found: str,
        expected: str,
        quote: str = "",
        hint: str = "",
        rule_ref: str = "",
    ) -> None:
        self.issues.append(
            Issue(
                rule_id=rule_id,
                category=category,
                severity=severity,
                title=title,
                location=location,
                found=found,
                expected=expected,
                quote=shorten(quote, 140) if quote else "",
                hint=hint,
                rule_ref=rule_ref,
            )
        )

    def _font_ok(self, name: str | None) -> bool:
        return _norm_font(name) in self.profile.normalized_font_names

    def _starts_new_page(self, index: int) -> bool:
        if index == 0:
            return True
        block = self.model.blocks[index]
        if block.paragraph is not None:
            ppr = self.model.styles.paragraph_format(block.paragraph)
            if ppr.page_break_before:
                return True
        for back in range(index - 1, max(-1, index - 4), -1):
            prev = self.model.blocks[back]
            if prev.paragraph is None:
                return False
            if self.model.paragraph_has_page_break(prev.paragraph):
                return True
            if self.model.paragraph_has_section_break(prev.paragraph):
                return True
            if normalize_ws(prev.text):
                return False
        return False

    def _is_body_prose(self, block: Block) -> bool:
        if block.paragraph is None or block.index < self.body_start:
            return False
        if block.index in self.heading_by_index:
            return False
        if block.index in self.table_caption_indices or block.index in self.figure_caption_indices:
            return False
        if block.index in self.bibliography_indices:
            return False
        text = normalize_ws(block.text)
        if len(text) < 30 or _is_toc_entry(text):
            return False
        if re.match(r"^(?:[-–—•*]|\d+[.)])\s+", text):
            return False
        if re.match(r"^(?:Источник|Примечание)\s*:", text, re.IGNORECASE):
            return False
        ppr = self.model.styles.paragraph_format(block.paragraph)
        return not ppr.has_numbering

    # --------------------------------------------------------------- page setup

    def _check_page_setup(self) -> None:
        sections = list(self.model.document.sections)
        for index, section in enumerate(sections, start=1):
            width_mm = section.page_width.mm if section.page_width is not None else None
            height_mm = section.page_height.mm if section.page_height is not None else None
            left_mm = section.left_margin.mm if section.left_margin is not None else None
            right_mm = section.right_margin.mm if section.right_margin is not None else None
            top_mm = section.top_margin.mm if section.top_margin is not None else None
            bottom_mm = section.bottom_margin.mm if section.bottom_margin is not None else None
            gutter_mm = section.gutter.mm if section.gutter is not None else 0.0

            wrong: list[str] = []
            expected: list[str] = []
            profile = self.profile
            tolerance = profile.margin_tolerance_mm

            if width_mm is None or height_mm is None:
                wrong.append("размер страницы не задан")
                expected.append("A4, 210 × 297 мм")
            else:
                is_a4 = (
                    abs(width_mm - profile.page_width_mm) <= 1.0
                    and abs(height_mm - profile.page_height_mm) <= 1.0
                ) or (
                    abs(width_mm - profile.page_height_mm) <= 1.0
                    and abs(height_mm - profile.page_width_mm) <= 1.0
                )
                if not is_a4:
                    wrong.append(f"формат {fmt_num(width_mm, 0)} × {fmt_num(height_mm, 0)} мм")
                    expected.append("A4, 210 × 297 мм")

            for label, value, reference in (
                ("левое", left_mm, profile.margin_left_mm),
                ("правое", right_mm, profile.margin_right_mm),
                ("верхнее", top_mm, profile.margin_top_mm),
                ("нижнее", bottom_mm, profile.margin_bottom_mm),
            ):
                if value is None or abs(value - reference) > tolerance:
                    actual = "не задано" if value is None else fmt_mm(value)
                    wrong.append(f"{label} поле {actual}")
                    expected.append(f"{label} {fmt_mm(reference)}")

            if gutter_mm and abs(gutter_mm) > 0.1:
                wrong.append(f"переплётный отступ {fmt_mm(gutter_mm)}")
                expected.append("переплётный отступ 0 мм")

            if wrong:
                self._add(
                    rule_id="page.margins",
                    category="Поля страницы",
                    severity=Severity.ERROR,
                    title="Неверные поля или формат страницы",
                    location=f"Секция {index}" if len(sections) > 1 else "Параметры страницы",
                    found="; ".join(wrong),
                    expected="; ".join(dict.fromkeys(expected)),
                    hint="Word: Макет → Поля → Настраиваемые поля. Для всего документа нужны одинаковые поля.",
                    rule_ref="ГОСТ 7.32-2017, п. 6.1.1",
                )

    def _check_page_numbers(self) -> None:
        try:
            footers = self.model.footer_definitions()
        except Exception:
            self._add(
                rule_id="pagination.unreadable",
                category="Нумерация страниц",
                severity=Severity.INFO,
                title="Не удалось прочитать колонтитулы",
                location="Колонтитулы",
                found="структура колонтитулов не распознана",
                expected="номер страницы внизу по центру, на титульном листе без номера",
                hint="Проверьте нумерацию вручную: Вставка → Номер страницы → Внизу страницы → По центру.",
                rule_ref="ГОСТ 7.32-2017, п. 6.3.2",
            )
            return

        if not footers:
            return

        first_section, first_effective, _ = footers[0]
        first_default = first_effective.get("default")
        title_pg = first_section._sectPr.find(qn("w:titlePg")) is not None
        any_page_number = any(
            self.model.footer_has_page_field(effective.get("default"))
            for _section, effective, _refs in footers
        )

        if not any_page_number:
            self._add(
                rule_id="pagination.missing",
                category="Нумерация страниц",
                severity=Severity.ERROR,
                title="Нет номера страницы в нижнем колонтитуле",
                location="Нижний колонтитул",
                found="поле PAGE не найдено",
                expected="арабский номер внизу страницы по центру, без точки",
                hint="Вставка → Номер страницы → Внизу страницы → По центру.",
                rule_ref="ГОСТ 7.32-2017, п. 6.3.2",
            )
        else:
            if not title_pg and self.model.footer_has_page_field(first_default):
                self._add(
                    rule_id="pagination.title_page",
                    category="Нумерация страниц",
                    severity=Severity.WARNING,
                    title="Не отключён номер на титульном листе",
                    location="Секция 1",
                    found="в параметрах секции не включён «особый колонтитул первой страницы»",
                    expected="титульный лист входит в нумерацию, но номер на нём не печатается",
                    hint="Word: Макет → Параметры страницы → Особый колонтитул для первой страницы.",
                    rule_ref="ГОСТ 7.32-2017, п. 6.3.2",
                )

            checked_footers: set[int] = set()
            for section_index, (_section, effective, _refs) in enumerate(footers, start=1):
                default_footer = effective.get("default")
                location = (
                    f"Нижний колонтитул, секция {section_index}"
                    if len(footers) > 1
                    else "Нижний колонтитул"
                )
                if not self.model.footer_has_page_field(default_footer):
                    # Пустая первая секция — это может быть титульный лист,
                    # если в следующих секциях нумерация уже есть.
                    if section_index == 1:
                        continue
                    self._add(
                        rule_id="pagination.missing",
                        category="Нумерация страниц",
                        severity=Severity.ERROR,
                        title="Нет номера страницы в нижнем колонтитуле",
                        location=location,
                        found="поле PAGE не найдено",
                        expected="арабский номер внизу страницы по центру, без точки",
                        hint="Вставка → Номер страницы → Внизу страницы → По центру.",
                        rule_ref="ГОСТ 7.32-2017, п. 6.3.2",
                    )
                    continue
                key = id(default_footer)
                if key in checked_footers:
                    continue
                checked_footers.add(key)
                self._check_footer_alignment(default_footer, location)
                self._check_footer_font(default_footer)

        if title_pg:
            first_footer = first_effective.get("first")
            if self.model.footer_has_page_field(first_footer):
                self._add(
                    rule_id="pagination.title_page.visible",
                    category="Нумерация страниц",
                    severity=Severity.ERROR,
                    title="Номер печатается на титульном листе",
                    location="Колонтитул первой страницы",
                    found="поле PAGE найдено в колонтитуле первой страницы",
                    expected="на титульном листе номер отсутствует",
                    hint="Очистите колонтитул первой страницы или включите особый колонтитул.",
                    rule_ref="ГОСТ 7.32-2017, п. 6.3.2",
                )

    def _check_footer_alignment(self, part: object | None, label: str) -> None:
        for paragraph in self.model.footer_page_paragraphs(part):
            ppr = self.model.styles.ppr_for_element(paragraph)
            # Word часто ставит номер не через выравнивание, а после
            # центрального табулятора в стиле нижнего колонтитула.
            centered_by_tab = (
                self.model.styles.ppr_has_center_tab(ppr.style_id, paragraph.find(qn("w:pPr")))
                and self._tabs_before_page_field(paragraph) == 1
            )
            if ppr.alignment not in {"center"} and not centered_by_tab:
                self._add(
                    rule_id="pagination.alignment",
                    category="Нумерация страниц",
                    severity=Severity.ERROR,
                    title="Номер страницы выровнен не по центру",
                    location=label.capitalize(),
                    found=f"выравнивание {_alignment_label(ppr.alignment)}",
                    expected="по центру нижней части листа",
                    hint="Выделите абзац с номером и нажмите «Выровнять по центру».",
                    rule_ref="ГОСТ 7.32-2017, п. 6.3.2",
                )

    @staticmethod
    def _tabs_before_page_field(paragraph: object) -> int:
        tabs = 0
        pattern = re.compile(r"(?<![A-Z])PAGE\b", re.IGNORECASE)
        for node in paragraph.iter():  # type: ignore[attr-defined]
            if node.tag == qn("w:tab") and node.getparent() is not None and node.getparent().tag == qn("w:r"):
                tabs += 1
            elif node.tag == qn("w:instrText") and pattern.search(node.text or ""):
                return tabs
        return tabs

    def _check_footer_font(self, part: object | None) -> None:
        for paragraph in self.model.footer_page_paragraphs(part):
            ppr = self.model.styles.ppr_for_element(paragraph)
            direct_rpr = None
            for run_element in paragraph.iter(qn("w:r")):
                has_visible_or_field = any(
                    (node.text or "") for node in run_element.iter(qn("w:t"))
                ) or any(
                    "PAGE" in (node.text or "").upper()
                    for node in run_element.iter(qn("w:instrText"))
                )
                if has_visible_or_field and run_element.rPr is not None:
                    direct_rpr = run_element.rPr
                    break
            if direct_rpr is None:
                ppr_element = paragraph.find(qn("w:pPr"))
                direct_rpr = ppr_element.find(qn("w:rPr")) if ppr_element is not None else None
            fmt = self.model.styles.resolve_rpr(ppr.style_id, None, direct_rpr)
            if not self._font_ok(fmt.hansi_font) or abs(fmt.size_pt - self.profile.body_font_pt) > 0.51:
                self._add(
                    rule_id="pagination.font",
                    category="Нумерация страниц",
                    severity=Severity.WARNING,
                    title="Шрифт номера страницы отличается от основного",
                    location="Нижний колонтитул",
                    found=f"{fmt.hansi_font}, {fmt_pt(fmt.size_pt)}",
                    expected=f"{self.profile.body_font}, {fmt_pt(self.profile.body_font_pt)}",
                    hint="Номер страницы оформляют тем же шрифтом и кеглем, что и основной текст.",
                    rule_ref="ГОСТ 7.32-2017, п. 6.3.2",
                )

    # ----------------------------------------------------------------- headings

    def _check_headings(self) -> None:
        for heading in self.headings:
            if heading.index < self.body_start and heading.clean not in {
                "СОДЕРЖАНИЕ",
                "ОГЛАВЛЕНИЕ",
            }:
                # Содержание проверяем, титульные надписи — нет.
                continue
            if not heading.style_based:
                self._add(
                    rule_id="heading.style",
                    category="Заголовки",
                    severity=Severity.ERROR,
                    title="Заголовок оформлен не стилем",
                    location=f"Абзац {heading.index + 1}",
                    found=f"«{shorten(heading.text, 100)}» — обычный текст или ручное форматирование",
                    expected="стиль «Заголовок 1» / «Заголовок 2» и структура документа",
                    quote=heading.text,
                    hint="Назначьте заголовку соответствующий стиль Word, чтобы работали содержание и навигация.",
                    rule_ref="Требование ТЗ; ГОСТ 7.32-2017, п. 6.2",
                )
            if heading.text.rstrip().endswith("."):
                self._add(
                    rule_id="heading.period",
                    category="Заголовки",
                    severity=Severity.ERROR,
                    title="Точка в конце заголовка",
                    location=f"Абзац {heading.index + 1}",
                    found=f"«{shorten(heading.text, 100)}» заканчивается точкой",
                    expected="без точки в конце заголовка",
                    quote=heading.text,
                    hint="Уберите точку. Двоеточие в конце также нежелательно.",
                    rule_ref="ГОСТ 7.32-2017, п. 6.2.2",
                )
            if heading.structural and not heading.all_caps:
                self._add(
                    rule_id="heading.uppercase",
                    category="Заголовки",
                    severity=Severity.WARNING,
                    title="Структурный заголовок написан не заглавными",
                    location=f"Абзац {heading.index + 1}",
                    found=f"«{shorten(heading.text, 100)}»",
                    expected=f"«{heading.clean.upper()}» — прописными буквами",
                    quote=heading.text,
                    hint="Структурные элементы (СОДЕРЖАНИЕ, ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, СПИСОК...) пишут заглавными.",
                    rule_ref="ГОСТ 7.32-2017, п. 6.2",
                )
            if (heading.level == 1 or heading.structural) and not self._starts_new_page(
                heading.index
            ):
                self._add(
                    rule_id="heading.new_page",
                    category="Заголовки",
                    severity=Severity.ERROR,
                    title="Раздел не начинается с новой страницы",
                    location=f"Абзац {heading.index + 1}",
                    found=f"перед заголовком «{shorten(heading.text, 80)}» нет разрыва страницы",
                    expected="каждый раздел и структурный элемент начинается с новой страницы",
                    quote=heading.text,
                    hint="Поставьте разрыв страницы перед заголовком (Ctrl+Enter) или включите «с новой страницы» в свойствах абзаца.",
                    rule_ref="ГОСТ 7.32-2017, п. 6.2; ТЗ",
                )

        for index, text in self.unstyled_structural:
            self._add(
                rule_id="heading.style",
                category="Заголовки",
                severity=Severity.ERROR,
                title="Структурный заголовок оформлен не стилем",
                location=f"Абзац {index + 1}",
                found=f"«{shorten(text, 100)}» — обычный текст",
                expected="стиль «Заголовок 1» и прописные буквы",
                quote=text,
                hint="Примените стиль заголовка, чтобы раздел попал в навигацию и содержание.",
                rule_ref="Требование ТЗ; ГОСТ 7.32-2017, п. 6.2",
            )

        for index, text in self.heading_like:
            self._add(
                rule_id="heading.unstyled",
                category="Заголовки",
                severity=Severity.WARNING,
                title="Похоже на заголовок, но оформлено обычным текстом",
                location=f"Абзац {index + 1}",
                found=f"«{shorten(text, 100)}» без стиля заголовка",
                expected="стиль «Заголовок 1» или «Заголовок 2»",
                quote=text,
                hint="Если это раздел работы, примените стиль заголовка и начните его с новой страницы.",
                rule_ref="Требование ТЗ",
            )

    # -------------------------------------------------------------------- fonts

    def _check_fonts(self) -> None:
        bad_by_paragraph: dict[int, dict[str, Counter[str]]] = {}

        for block in self.model.blocks:
            if block.paragraph is None or block.index < self.body_start:
                continue
            paragraph = block.paragraph
            paragraph_style_id = self.model.styles.paragraph_style_id(paragraph)
            for run in self.model.iter_runs(paragraph):
                text = run.text
                if not re.search(r"[A-Za-zА-Яа-яЁё0-9]", text or ""):
                    continue
                fmt = self.model.styles.run_format(run, paragraph_style_id)
                has_cyr = bool(re.search(r"[А-Яа-яЁё]", text))
                has_lat = bool(re.search(r"[A-Za-z]", text))

                bad_fonts: list[str] = []
                if (has_cyr or not has_lat) and not self._font_ok(fmt.hansi_font):
                    bad_fonts.append(fmt.hansi_font)
                if (
                    has_lat
                    and not self._font_ok(fmt.ascii_font)
                    and fmt.ascii_font not in bad_fonts
                ):
                    bad_fonts.append(fmt.ascii_font)
                bad_size = abs(fmt.size_pt - self.profile.body_font_pt) > self.profile.font_tolerance_pt

                if not bad_fonts and not bad_size:
                    continue

                entry = bad_by_paragraph.setdefault(
                    block.index,
                    {"fonts": Counter(), "sizes": Counter()},
                )
                for font_name in bad_fonts:
                    entry["fonts"][font_name or "не задан"] += 1
                if bad_size:
                    entry["sizes"][fmt_pt(fmt.size_pt)] += 1

        for index, entry in sorted(bad_by_paragraph.items()):
            block = self.model.blocks[index]
            heading = self.heading_by_index.get(index)
            found_parts: list[str] = []
            for name, count in entry["fonts"].most_common():
                suffix = f" ({count} фрагм.)" if count > 1 else ""
                found_parts.append(f"шрифт {name}{suffix}")
            for size, count in entry["sizes"].most_common():
                suffix = f" ({count} фрагм.)" if count > 1 else ""
                found_parts.append(f"размер {size}{suffix}")
            location = f"Абзац {index + 1}"
            if heading is not None:
                location = f"Заголовок, абзац {index + 1}"
            self._add(
                rule_id="font.body",
                category="Шрифт",
                severity=Severity.ERROR,
                title=(
                    "Неверный шрифт или размер в заголовке"
                    if heading is not None
                    else "Неверный шрифт или размер"
                ),
                location=location,
                found="; ".join(found_parts),
                expected=f"{self.profile.body_font}, {fmt_pt(self.profile.body_font_pt)}",
                quote=block.text,
                hint="Проверьте и прямой формат фрагмента, и стиль абзаца: Word может наследовать шрифт из стиля.",
                rule_ref="ГОСТ 7.32-2017, п. 6.1.1; ТЗ",
            )

        self._check_table_fonts()

    def _check_table_fonts(self) -> None:
        for block in self.model.blocks:
            if block.table is None:
                continue
            bad_fonts: Counter[str] = Counter()
            bad_sizes: Counter[str] = Counter()
            for paragraph in self.model.table_paragraphs(block.table):
                style_id = self.model.styles.paragraph_style_id(paragraph)
                for run in self.model.iter_runs(paragraph):
                    text = run.text
                    if not re.search(r"[A-Za-zА-Яа-яЁё0-9]", text or ""):
                        continue
                    fmt = self.model.styles.run_format(run, style_id)
                    has_cyr = bool(re.search(r"[А-Яа-яЁё]", text))
                    has_lat = bool(re.search(r"[A-Za-z]", text))
                    if (has_cyr or not has_lat) and not self._font_ok(fmt.hansi_font):
                        bad_fonts[fmt.hansi_font] += 1
                    if has_lat and not self._font_ok(fmt.ascii_font):
                        bad_fonts[fmt.ascii_font] += 1
                    if fmt.size_pt + 0.01 < self.profile.table_min_font_pt:
                        bad_sizes[fmt_pt(fmt.size_pt)] += 1

            if not bad_fonts and not bad_sizes:
                continue
            found_parts = [f"шрифт {name}" for name, _ in bad_fonts.most_common()]
            found_parts += [f"размер {size}" for size, _ in bad_sizes.most_common()]
            self._add(
                rule_id="font.table",
                category="Шрифт",
                severity=Severity.ERROR,
                title="Неверный шрифт или слишком мелкий текст в таблице",
                location=f"Таблица, абзац {block.index + 1}",
                found="; ".join(found_parts),
                expected=(
                    f"{self.profile.body_font}, не менее "
                    f"{fmt_pt(self.profile.table_min_font_pt)} (в таблице допустим меньший кегль)"
                ),
                quote=block.text,
                hint="В таблицах ГОСТ допускает кегль меньше основного, но не произвольный шрифт.",
                rule_ref="ГОСТ 7.32-2017, п. 6.6.6",
            )

    # --------------------------------------------------------------- paragraphs

    def _check_paragraphs(self) -> None:
        for block in self.model.blocks:
            if not self._is_body_prose(block) or block.paragraph is None:
                continue
            ppr = self.model.styles.paragraph_format(block.paragraph)
            found: list[str] = []
            expected: list[str] = []

            if ppr.alignment != self.profile.body_alignment:
                found.append(f"выравнивание {_alignment_label(ppr.alignment)}")
                expected.append("по ширине")

            multiple = ppr.spacing_multiple
            if ppr.line_rule in {"exact", "atLeast"}:
                line_pt = (ppr.line_value or 0) / 20.0
                found.append(f"точный интервал {fmt_pt(line_pt)}")
                expected.append("интервал 1,5 строки")
            elif multiple is None:
                found.append("одинарный интервал (не задан)")
                expected.append("интервал 1,5 строки")
            elif abs(multiple - self.profile.line_spacing) > self.profile.line_spacing_tolerance:
                found.append(f"интервал {fmt_num(multiple, 2)}")
                expected.append("интервал 1,5 строки")

            if (
                ppr.before_pt > self.profile.extra_spacing_tolerance_pt
                or ppr.after_pt > self.profile.extra_spacing_tolerance_pt
            ):
                extras: list[str] = []
                if ppr.before_pt > self.profile.extra_spacing_tolerance_pt:
                    extras.append(f"до {fmt_pt(ppr.before_pt)}")
                if ppr.after_pt > self.profile.extra_spacing_tolerance_pt:
                    extras.append(f"после {fmt_pt(ppr.after_pt)}")
                found.append("интервалы " + ", ".join(extras))
                expected.append("без дополнительных интервалов до/после абзаца")

            first_line_cm = ppr.first_line_cm(self.profile.body_font_pt)
            if ppr.hanging_twips:
                found.append(f"висячий отступ {fmt_cm(abs(first_line_cm))}")
                expected.append("отступ первой строки 1,25 см")
            elif abs(first_line_cm - self.profile.first_line_indent_cm) > self.profile.indent_tolerance_cm:
                found.append(f"отступ первой строки {fmt_cm(first_line_cm)}")
                expected.append("отступ первой строки 1,25 см")

            if ppr.left_cm > self.profile.side_indent_tolerance_cm:
                found.append(f"левый отступ {fmt_cm(ppr.left_cm)}")
                expected.append("без боковых отступов")
            if ppr.right_cm > self.profile.side_indent_tolerance_cm:
                found.append(f"правый отступ {fmt_cm(ppr.right_cm)}")
                expected.append("без боковых отступов")

            if not found:
                continue

            self._add(
                rule_id="paragraph.format",
                category="Абзац и интервал",
                severity=Severity.ERROR,
                title="Нарушено оформление абзаца",
                location=f"Абзац {block.index + 1}",
                found="; ".join(found),
                expected="; ".join(dict.fromkeys(expected)),
                quote=block.text,
                hint="Удобнее исправить стиль «Обычный», а не каждый абзац по отдельности.",
                rule_ref="ГОСТ 7.32-2017, п. 6.1.1; ТЗ",
            )

    # ------------------------------------------------------------------ figures

    def _scan_figures(self) -> None:
        numbers: list[CaptionInfo] = []
        for block in self.model.blocks:
            if block.paragraph is None:
                continue
            if not self.model.element_has_image(block.element):
                continue
            self.figure_count += 1
            ppr = self.model.styles.paragraph_format(block.paragraph)
            if ppr.alignment not in {"center"}:
                self._add(
                    rule_id="figure.image.center",
                    category="Рисунки",
                    severity=Severity.WARNING,
                    title="Рисунок не выровнен по центру",
                    location=f"Абзац {block.index + 1}",
                    found=f"выравнивание {_alignment_label(ppr.alignment)}",
                    expected="рисунок по центру",
                    hint="Выделите абзац с рисунком и выровняйте по центру.",
                    rule_ref="ГОСТ 7.32-2017, п. 6.5",
                )

            caption_block = self._next_nonempty(block.index, limit=2)
            if caption_block is None or caption_block.paragraph is None:
                self._add(
                    rule_id="figure.caption.missing",
                    category="Рисунки",
                    severity=Severity.ERROR,
                    title="Под рисунком нет подписи",
                    location=f"Абзац {block.index + 1}",
                    found="после рисунка не найден абзац с подписью",
                    expected="Рисунок N — Название (под рисунком, по центру)",
                    hint="Добавьте подпись сразу под рисунком.",
                    rule_ref="ГОСТ 7.32-2017, п. 6.5.6",
                )
                continue

            text = normalize_ws(caption_block.text)
            match = FIGURE_CAPTION_RE.match(text)
            if match is None:
                if text.startswith(("Рис.", "Рис ")):
                    self._add(
                        rule_id="figure.caption.word",
                        category="Рисунки",
                        severity=Severity.ERROR,
                        title="Сокращение «Рис.» в подписи",
                        location=f"Абзац {caption_block.index + 1}",
                        found=f"«{shorten(text, 100)}»",
                        expected="слово «Рисунок» пишется полностью: «Рисунок 1 — Название»",
                        quote=text,
                        hint="Замените «Рис. 1» на «Рисунок 1».",
                        rule_ref="ГОСТ 7.32-2017, п. 6.5.6",
                    )
                else:
                    self._add(
                        rule_id="figure.caption.format",
                        category="Рисунки",
                        severity=Severity.WARNING,
                        title="Подпись рисунка не по шаблону",
                        location=f"Абзац {caption_block.index + 1}",
                        found=f"«{shorten(text, 100)}»",
                        expected="Рисунок N — Название",
                        quote=text,
                        hint="Проверьте номер, длинное тире и название с прописной буквы.",
                        rule_ref="ГОСТ 7.32-2017, п. 6.5.6",
                    )
                loose = re.match(
                    r"^Рис(?:унок|\.)?\s*([А-ЯЁA-Z]\.)?(\d+(?:\.\d+)?)\s*([—–-])?\s*(.*)$",
                    text,
                )
                if loose:
                    number = loose.group(2)
                    caption = CaptionInfo(
                        index=caption_block.index,
                        number=number,
                        numeric=int(number) if number.isdigit() else None,
                        structure_ok=False,
                        separator=loose.group(3) or "",
                        text=text,
                    )
                    numbers.append(caption)
                    self.figure_captions.append(caption)
                    self._check_caption_common(
                        caption_block,
                        caption,
                        expected="Рисунок N — Название, по центру, без точки в конце",
                        kind="рисунка",
                        center=True,
                    )
                self.figure_caption_indices.add(caption_block.index)
                continue

            self.figure_caption_indices.add(caption_block.index)
            number = match.group(2)
            separator = match.group(3)
            caption = CaptionInfo(
                index=caption_block.index,
                number=number,
                numeric=int(number) if number.isdigit() else None,
                structure_ok=True,
                separator=separator,
                text=text,
            )
            numbers.append(caption)
            self.figure_captions.append(caption)
            self._check_caption_common(
                caption_block,
                caption,
                expected="Рисунок N — Название, по центру, без точки в конце",
                kind="рисунка",
                center=True,
            )

        self._check_number_sequence(numbers, "Рисунки", "рисунка")

    def _check_caption_common(
        self,
        block: Block,
        caption: CaptionInfo,
        *,
        expected: str,
        kind: str,
        center: bool,
    ) -> None:
        if block.paragraph is None:
            return
        ppr = self.model.styles.paragraph_format(block.paragraph)
        if center:
            if ppr.alignment not in {"center"}:
                self._add(
                    rule_id="caption.alignment",
                    category="Рисунки" if kind == "рисунка" else "Таблицы",
                    severity=Severity.ERROR,
                    title=f"Подпись {kind} выровнена не по центру",
                    location=f"Абзац {block.index + 1}",
                    found=f"выравнивание {_alignment_label(ppr.alignment)}",
                    expected="по центру",
                    quote=block.text,
                    rule_ref="ГОСТ 7.32-2017",
                )
        else:
            if ppr.alignment not in {"left", "start", None}:
                self._add(
                    rule_id="caption.alignment",
                    category="Таблицы",
                    severity=Severity.ERROR,
                    title=f"Подпись {kind} выровнена не влево",
                    location=f"Абзац {block.index + 1}",
                    found=f"выравнивание {_alignment_label(ppr.alignment)}",
                    expected="слева, без абзацного отступа",
                    quote=block.text,
                    rule_ref="ГОСТ 7.32-2017, п. 6.6.3",
                )

        if abs(ppr.first_line_cm(self.profile.body_font_pt)) > self.profile.indent_tolerance_cm:
            self._add(
                rule_id="caption.indent",
                category="Рисунки" if kind == "рисунка" else "Таблицы",
                severity=Severity.ERROR,
                title=f"У подписи {kind} есть абзацный отступ",
                location=f"Абзац {block.index + 1}",
                found=f"отступ первой строки {fmt_cm(ppr.first_line_cm(self.profile.body_font_pt))}",
                expected="без абзацного отступа",
                quote=block.text,
                rule_ref="ГОСТ 7.32-2017",
            )

        if caption.text.rstrip().endswith("."):
            self._add(
                rule_id="caption.period",
                category="Рисунки" if kind == "рисунка" else "Таблицы",
                severity=Severity.WARNING,
                title=f"Точка в конце подписи {kind}",
                location=f"Абзац {block.index + 1}",
                found=f"«{shorten(caption.text, 100)}»",
                expected="без точки в конце",
                quote=caption.text,
                rule_ref="ГОСТ 7.32-2017",
            )

        if caption.separator == "-":
            self._add(
                rule_id="caption.dash",
                category="Рисунки" if kind == "рисунка" else "Таблицы",
                severity=Severity.WARNING,
                title=f"В подписи {kind} стоит дефис вместо тире",
                location=f"Абзац {block.index + 1}",
                found=f"«{shorten(caption.text, 100)}»",
                expected="длинное тире: «—»",
                quote=caption.text,
                hint="По ТЗ используется длинное тире (—), не дефис (-).",
                rule_ref="Требование ТЗ; ГОСТ 7.32-2017",
            )

    def _check_number_sequence(
        self,
        captions: list[CaptionInfo],
        category: str,
        kind: str,
    ) -> None:
        simple = [c for c in captions if c.numeric is not None]
        if not simple:
            return
        seen: set[int] = set()
        duplicates: list[int] = []
        broken: list[str] = []
        for position, caption in enumerate(simple, start=1):
            number = caption.numeric
            if number is None:
                continue
            if number in seen:
                duplicates.append(number)
            seen.add(number)
            if number != position:
                broken.append(f"вместо {position} — {number}")

        if broken or duplicates:
            details: list[str] = []
            if broken:
                details.append("нарушен порядок: " + ", ".join(broken[:5]))
            if duplicates:
                details.append(
                    "повторы номеров: " + ", ".join(str(n) for n in sorted(set(duplicates))[:5])
                )
            self._add(
                rule_id=f"{category.lower()}.numbering",
                category=category,
                severity=Severity.WARNING,
                title=f"Нумерация {kind} сбита",
                location="По документу",
                found="; ".join(details),
                expected=f"сквозная нумерация {kind} арабскими цифрами: 1, 2, 3…",
                hint="Если используется нумерация в пределах раздела, проверьте номера вида 1.1, 1.2.",
                rule_ref="ГОСТ 7.32-2017",
            )

    # ------------------------------------------------------------------- tables

    def _scan_tables(self) -> None:
        numbers: list[CaptionInfo] = []
        for block in self.model.blocks:
            if block.table is None:
                continue
            self.table_count += 1
            caption_block = self._previous_nonempty(block.index, limit=2)
            if caption_block is None or caption_block.paragraph is None:
                self._add(
                    rule_id="table.caption.missing",
                    category="Таблицы",
                    severity=Severity.ERROR,
                    title="У таблицы нет подписи",
                    location=f"Таблица, абзац {block.index + 1}",
                    found="перед таблицей не найден абзац «Таблица N — Название»",
                    expected="Таблица N — Название (слева над таблицей, без абзацного отступа)",
                    hint="Добавьте подпись непосредственно перед таблицей.",
                    rule_ref="ГОСТ 7.32-2017, п. 6.6.3",
                )
                continue

            text = normalize_ws(caption_block.text)
            if re.match(r"^(?:Продолжение|Окончание)\s+таблицы\b", text, re.IGNORECASE):
                self.table_caption_indices.add(caption_block.index)
                ppr = self.model.styles.paragraph_format(caption_block.paragraph)
                if ppr.alignment not in {"left", "start", None}:
                    self._add(
                        rule_id="table.continuation.alignment",
                        category="Таблицы",
                        severity=Severity.WARNING,
                        title="Подпись продолжения таблицы выровнена не влево",
                        location=f"Абзац {caption_block.index + 1}",
                        found=f"выравнивание {_alignment_label(ppr.alignment)}",
                        expected="слева, без абзацного отступа",
                        quote=text,
                        rule_ref="ГОСТ 7.32-2017, п. 6.6",
                    )
                continue

            match = TABLE_CAPTION_RE.match(text)
            if match is None:
                if text.startswith("Таблица"):
                    self._add(
                        rule_id="table.caption.format",
                        category="Таблицы",
                        severity=Severity.WARNING,
                        title="Подпись таблицы не по шаблону",
                        location=f"Абзац {caption_block.index + 1}",
                        found=f"«{shorten(text, 100)}»",
                        expected="Таблица N — Название",
                        quote=text,
                        hint="После номера ставят длинное тире, затем название с прописной буквы.",
                        rule_ref="ГОСТ 7.32-2017, п. 6.6.3",
                    )
                    self.table_caption_indices.add(caption_block.index)
                else:
                    self._add(
                        rule_id="table.caption.missing",
                        category="Таблицы",
                        severity=Severity.ERROR,
                        title="У таблицы нет подписи",
                        location=f"Таблица, абзац {block.index + 1}",
                        found=f"перед таблицей: «{shorten(text, 100)}»",
                        expected="Таблица N — Название",
                        quote=text,
                        rule_ref="ГОСТ 7.32-2017, п. 6.6.3",
                    )
                continue

            self.table_caption_indices.add(caption_block.index)
            number = match.group(2)
            caption = CaptionInfo(
                index=caption_block.index,
                number=number,
                numeric=int(number) if number.isdigit() else None,
                structure_ok=True,
                separator=match.group(3),
                text=text,
            )
            numbers.append(caption)
            self.table_captions.append(caption)
            self._check_caption_common(
                caption_block,
                caption,
                expected="Таблица N — Название, слева, без отступа",
                kind="таблицы",
                center=False,
            )

        self._check_number_sequence(numbers, "Таблицы", "таблицы")

    # ------------------------------------------------------------- bibliography

    def _scan_bibliography(self) -> None:
        heading = next(
            (h for h in self.headings if BIBLIOGRAPHY_RE.match(h.clean)),
            None,
        )
        if heading is None:
            self._add(
                rule_id="bibliography.missing",
                category="Список литературы",
                severity=Severity.ERROR,
                title="Не найден список источников",
                location="Документ",
                found="раздел «Список использованных источников» / «Список литературы» не найден",
                expected="отдельный раздел со списком использованных источников",
                hint="Добавьте раздел «СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ» после заключения.",
                rule_ref="ГОСТ 7.32-2017, п. 6.10; ГОСТ Р 7.0.100-2018",
            )
            return

        self.bibliography_heading_index = heading.index
        stop = len(self.model.blocks)
        for block in self.model.blocks[heading.index + 1 :]:
            info = self.heading_by_index.get(block.index)
            if info is not None and (info.level <= 1 or info.structural):
                stop = block.index
                break
            # Даже если приложение не оформлено стилем, его текст не должен
            # попадать в список литературы.
            clean = normalize_ws(block.text).rstrip(":").strip()
            if clean and STRUCTURAL_RE.match(clean) and not BIBLIOGRAPHY_RE.match(clean):
                stop = block.index
                break

        for block in self.model.blocks[heading.index + 1 : stop]:
            if block.paragraph is None:
                continue
            self.bibliography_indices.add(block.index)
            if not normalize_ws(block.text):
                continue
            if block.index in self.heading_by_index:
                continue
            # Подзаголовки вида «Источники на русском языке:» не являются
            # библиографическими записями.
            if normalize_ws(block.text).endswith(":") and len(normalize_ws(block.text)) < 90:
                continue
            self.bibliography_entries.append(block)

        self.source_count = len(self.bibliography_entries)
        if not self.bibliography_entries:
            self._add(
                rule_id="bibliography.empty",
                category="Список литературы",
                severity=Severity.ERROR,
                title="Список источников пуст",
                location=f"Абзац {heading.index + 1}",
                found="после заголовка нет библиографических записей",
                expected="пронумерованный список источников",
                quote=heading.text,
                rule_ref="ГОСТ 7.32-2017, п. 6.10",
            )
            return

        manual_numbers: list[int] = []
        missing_number: list[int] = []
        for block in self.bibliography_entries:
            assert block.paragraph is not None
            ppr = self.model.styles.paragraph_format(block.paragraph)
            text = normalize_ws(block.text)
            match = re.match(r"^(\d+)[.)]\s*(.+)$", text)
            if match:
                manual_numbers.append(int(match.group(1)))
                body = match.group(2)
            else:
                if not ppr.has_numbering:
                    missing_number.append(block.index)
                body = text

            if len(body) < 30:
                self._add(
                    rule_id="bibliography.short",
                    category="Список литературы",
                    severity=Severity.WARNING,
                    title="Библиографическая запись слишком короткая",
                    location=f"Абзац {block.index + 1}",
                    found=f"«{shorten(text, 100)}»",
                    expected="автор, название, город, издательство, год, страницы (или URL и дата обращения)",
                    quote=text,
                    rule_ref="ГОСТ Р 7.0.100-2018",
                )

            content_problems = self._bibliography_content_problems(body)
            if content_problems:
                self._add(
                    rule_id="bibliography.content",
                    category="Список литературы",
                    severity=Severity.WARNING,
                    title="Проверьте структуру библиографической записи",
                    location=f"Абзац {block.index + 1}",
                    found="; ".join(content_problems),
                    expected="автор. название. – город : издательство, год. – N с. (для статьи — // журнал, год, №, страницы)",
                    quote=text,
                    hint="Сверьте запись с ГОСТ Р 7.0.100-2018 и примерами для своего типа источника.",
                    rule_ref="ГОСТ Р 7.0.100-2018; ГОСТ Р 7.0.5-2008",
                )

            format_problems: list[str] = []
            if ppr.alignment != self.profile.body_alignment:
                format_problems.append(f"выравнивание {_alignment_label(ppr.alignment)}")
            multiple = ppr.spacing_multiple
            if ppr.line_rule in {"exact", "atLeast"}:
                format_problems.append("точный межстрочный интервал")
            elif multiple is None:
                format_problems.append("одинарный интервал")
            elif abs(multiple - self.profile.line_spacing) > self.profile.line_spacing_tolerance:
                format_problems.append(f"интервал {fmt_num(multiple, 2)}")
            if not ppr.has_numbering:
                first_line = ppr.first_line_cm(self.profile.body_font_pt)
                if abs(first_line - self.profile.first_line_indent_cm) > self.profile.indent_tolerance_cm:
                    format_problems.append(f"отступ первой строки {fmt_cm(first_line)}")
            if format_problems:
                self._add(
                    rule_id="bibliography.format",
                    category="Список литературы",
                    severity=Severity.WARNING,
                    title="Нарушено оформление записи списка",
                    location=f"Абзац {block.index + 1}",
                    found="; ".join(format_problems),
                    expected="по ширине, интервал 1,5, отступ первой строки 1,25 см",
                    quote=text,
                    rule_ref="ГОСТ 7.32-2017, п. 6.1.1",
                )

        if missing_number:
            locations = ", ".join(str(i + 1) for i in missing_number[:5])
            self._add(
                rule_id="bibliography.numbering.missing",
                category="Список литературы",
                severity=Severity.ERROR,
                title="Не все источники пронумерованы",
                location=f"Абзацы {locations}" if len(missing_number) > 1 else f"Абзац {locations}",
                found=f"без номера: {len(missing_number)} записей",
                expected="сквозная нумерация арабскими цифрами: 1, 2, 3…",
                hint="Нумеруйте вручную или используйте нумерованный список Word.",
                rule_ref="ГОСТ 7.32-2017, п. 6.10",
            )

        if manual_numbers:
            expected_numbers = list(range(1, len(manual_numbers) + 1))
            if manual_numbers != expected_numbers:
                self._add(
                    rule_id="bibliography.numbering.order",
                    category="Список литературы",
                    severity=Severity.WARNING,
                    title="Нумерация списка источников сбита",
                    location="Список литературы",
                    found="номера: " + ", ".join(str(n) for n in manual_numbers[:15]),
                    expected="1, 2, 3… без пропусков и повторов",
                    rule_ref="ГОСТ 7.32-2017, п. 6.10",
                )

    def _bibliography_content_problems(self, text: str) -> list[str]:
        problems: list[str] = []
        lower = text.casefold()
        if not YEAR_RE.search(text):
            problems.append("не указан год издания")
        if re.search(r"\b(?:19|20)\d{2}\s*г\.", text):
            problems.append("после года стоит «г.»")

        is_web = bool(re.search(r"https?://|\bURL\s*:", text, re.IGNORECASE))
        is_article = "//" in text

        if is_web:
            if not re.search(r"https?://", text, re.IGNORECASE):
                problems.append("не указан полный URL")
            if not ACCESS_DATE_RE.search(text):
                problems.append("не указана дата обращения")
        elif is_article:
            if not re.search(r"\bС\.\s*\d+", text, re.IGNORECASE):
                problems.append("не указаны страницы статьи")
            if "//" not in text:
                problems.append("не указано издание после //")
        else:
            if not re.search(r"\b\d+\s*с\.", lower):
                problems.append("не указано количество страниц")
            has_place = re.search(
                r"(?:^|[.\s])(?:М|СПб|Санкт-Петербург|Москва|Киев|Минск|Екатеринбург|"
                r"Новосибирск|Казань|Ростов-на-Дону|Томск|Самара|Омск)\s*[.:]",
                text,
                re.IGNORECASE,
            ) or re.search(
                r"(?:^|[.\s–-])[А-ЯЁ][А-Яа-яЁё./-]*"
                r"(?:\s+[А-Яа-яЁё./-]+)?\s*:\s*[А-ЯЁ]",
                text,
            )
            if not has_place:
                problems.append("не указано место издания")
        return problems

    # ---------------------------------------------------------------- citations

    def _check_citations(self) -> None:
        if self.source_count <= 0:
            return
        bad: list[tuple[int, list[int]]] = []
        any_citation = False
        for block in self.model.blocks:
            if block.paragraph is None or block.index < self.body_start:
                continue
            if block.index in self.bibliography_indices:
                continue
            text = normalize_ws(block.text)
            for match in CITATION_RE.finditer(text):
                inner = match.group(1).strip()
                if not re.match(r"^\d", inner):
                    continue
                any_citation = True
                self.citation_count += 1
                numbers = self._citation_numbers(inner)
                invalid = [n for n in numbers if n < 1 or n > self.source_count]
                if invalid:
                    bad.append((block.index, invalid))

        for index, numbers in bad:
            block = self.model.blocks[index]
            unique = sorted(set(numbers))
            self._add(
                rule_id="citation.range",
                category="Ссылки в тексте",
                severity=Severity.ERROR,
                title="Ссылка на несуществующий источник",
                location=f"Абзац {index + 1}",
                found=f"номер(а) {', '.join(str(n) for n in unique)} при {self.source_count} источниках",
                expected=f"номера от 1 до {self.source_count}",
                quote=block.text,
                hint="Проверьте номер в квадратных скобках и количество записей в списке литературы.",
                rule_ref="ГОСТ Р 7.0.5-2008",
            )

        if not any_citation:
            self._add(
                rule_id="citation.missing",
                category="Ссылки в тексте",
                severity=Severity.WARNING,
                title="В тексте нет ссылок на источники",
                location="Основной текст",
                found="не найдено ссылок вида [1], [1, с. 25]",
                expected="ссылки на источники из списка литературы",
                hint="Если используются подстрочные сноски, проверьте их вручную.",
                rule_ref="ГОСТ Р 7.0.5-2008",
            )

    @staticmethod
    def _citation_numbers(inner: str) -> list[int]:
        head = re.split(r"\bс\.\s*", inner, maxsplit=1, flags=re.IGNORECASE)[0]
        numbers: list[int] = []
        for start, end in re.findall(r"(\d+)\s*[–—-]\s*(\d+)", head):
            left, right = int(start), int(end)
            if left <= right and right - left <= 100:
                numbers.extend(range(left, right + 1))
        cleaned = re.sub(r"\d+\s*[–—-]\s*\d+", "", head)
        numbers.extend(int(n) for n in re.findall(r"\d+", cleaned))
        return numbers

    # --------------------------------------------------------------- structure

    def _check_required_sections(self) -> None:
        present = {h.clean for h in self.headings}
        if not ({"СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ"} & present):
            self._add(
                rule_id="structure.contents",
                category="Структура",
                severity=Severity.WARNING,
                title="Не найден раздел «Содержание»",
                location="Документ",
                found="заголовок «СОДЕРЖАНИЕ» или «ОГЛАВЛЕНИЕ» не найден",
                expected="отдельный структурный элемент «СОДЕРЖАНИЕ»",
                hint="Проверьте, что заголовок написан заглавными и оформлен стилем заголовка.",
                rule_ref="ГОСТ 7.32-2017, п. 6.2",
            )
        required = (
            ("ВВЕДЕНИЕ", "Введение"),
            ("ЗАКЛЮЧЕНИЕ", "Заключение"),
        )
        for key, title in required:
            if key not in present:
                self._add(
                    rule_id=f"structure.{key.lower()}",
                    category="Структура",
                    severity=Severity.WARNING,
                    title=f"Не найден раздел «{title}»",
                    location="Документ",
                    found=f"заголовок «{key}» не найден",
                    expected=f"отдельный структурный элемент «{key}»",
                    hint="Проверьте, что заголовок написан заглавными и оформлен стилем заголовка.",
                    rule_ref="ГОСТ 7.32-2017, п. 6.2",
                )


def analyze_docx(
    path: str | Path,
    *,
    file_name: str | None = None,
    max_issues: int = 1000,
) -> AnalysisResult:
    try:
        analyzer = GostAnalyzer(path, file_name=file_name, max_issues=max_issues)
        return analyzer.analyze()
    except PackageNotFoundError as exc:
        raise AnalysisError(
            "Не удалось открыть файл как .docx. Возможно, он повреждён, защищён паролем "
            "или на самом деле является .doc/.pdf.",
            detail=str(exc),
        ) from exc
    except AnalysisError:
        raise
    except Exception as exc:
        raise AnalysisError(
            "Не удалось разобрать документ. Убедитесь, что это обычный .docx без пароля "
            "и повреждений.",
            detail=f"{type(exc).__name__}: {exc}",
        ) from exc
