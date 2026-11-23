from __future__ import annotations

import re
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path

from docx import Document as open_docx
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .units import PT_PER_CM, TWIPS_PER_CM, half_points_to_pt

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


@dataclass(frozen=True)
class EffectiveParagraph:
    style_id: str
    style_name: str
    alignment: str | None
    line_rule: str | None
    line_value: int | None
    before_twips: int
    after_twips: int
    first_line_twips: int | None
    first_line_chars: int | None
    hanging_twips: int | None
    left_twips: int
    right_twips: int
    page_break_before: bool
    outline_level: int | None
    has_numbering: bool

    @property
    def spacing_multiple(self) -> float | None:
        # В OOXML отсутствующий lineRule означает auto, а не exact.
        if self.line_rule in {None, "auto"} and self.line_value:
            return self.line_value / 240.0
        return None

    @property
    def before_pt(self) -> float:
        return self.before_twips / 20.0

    @property
    def after_pt(self) -> float:
        return self.after_twips / 20.0

    def first_line_cm(self, font_pt: float) -> float:
        if self.hanging_twips:
            return -self.hanging_twips / TWIPS_PER_CM
        if self.first_line_twips:
            return self.first_line_twips / TWIPS_PER_CM
        if self.first_line_chars is not None:
            # Символьный отступ Word считает от кегля, а не в сантиметрах.
            return (self.first_line_chars / 100.0) * font_pt / PT_PER_CM
        return 0.0

    @property
    def left_cm(self) -> float:
        return self.left_twips / TWIPS_PER_CM

    @property
    def right_cm(self) -> float:
        return self.right_twips / TWIPS_PER_CM


@dataclass(frozen=True)
class EffectiveFont:
    ascii_font: str
    hansi_font: str
    size_pt: float
    bold: bool
    italic: bool
    color: str
    underline: bool


@dataclass
class Block:
    index: int
    kind: str
    element: object
    text: str
    paragraph: Paragraph | None = None
    table: Table | None = None
    section_index: int = 0

    @property
    def is_paragraph(self) -> bool:
        return self.kind == "paragraph"

    @property
    def is_table(self) -> bool:
        return self.kind == "table"


def _style_id_from_ppr(ppr: object | None) -> str | None:
    if ppr is None:
        return None
    node = ppr.find(qn("w:pStyle"))  # type: ignore[attr-defined]
    if node is None:
        return None
    return node.get(qn("w:val"))


def _style_id_from_rpr(rpr: object | None) -> str | None:
    if rpr is None:
        return None
    node = rpr.find(qn("w:rStyle"))  # type: ignore[attr-defined]
    if node is None:
        return None
    return node.get(qn("w:val"))


class StyleResolver:
    def __init__(self, document) -> None:
        self.document = document
        root = document.styles.element
        self.styles: dict[str, object] = {}
        self.default_paragraph_style_id: str | None = None
        for style in root.findall(qn("w:style")):
            style_id = style.get(qn("w:styleId"))
            if not style_id:
                continue
            self.styles[style_id] = style
            if (
                style.get(qn("w:type")) == "paragraph"
                and style.get(qn("w:default")) in {"1", "true", "on"}
            ):
                self.default_paragraph_style_id = style_id

        doc_defaults = root.find(qn("w:docDefaults"))
        self.default_ppr = None
        self.default_rpr = None
        if doc_defaults is not None:
            ppr_default = doc_defaults.find(qn("w:pPrDefault"))
            if ppr_default is not None:
                self.default_ppr = ppr_default.find(qn("w:pPr"))
            rpr_default = doc_defaults.find(qn("w:rPrDefault"))
            if rpr_default is not None:
                self.default_rpr = rpr_default.find(qn("w:rPr"))

        self.theme_fonts = self._read_theme_fonts(document)
        self._chain_cache: dict[str, tuple[object, ...]] = {}
        self._style_ppr_cache: dict[str, tuple[object, ...]] = {}
        self._style_rpr_cache: dict[str, tuple[object, ...]] = {}
        self._ppr_cache: dict[str, EffectiveParagraph] = {}
        self._rpr_cache: dict[str, EffectiveFont] = {}

    # ------------------------------------------------------------------ styles

    @staticmethod
    def _read_theme_fonts(document) -> dict[str, str]:
        result: dict[str, str] = {}
        for part in document.part.package.iter_parts():
            name = str(getattr(part, "partname", ""))
            if not name.endswith("theme1.xml") and "theme" not in name:
                continue
            try:
                root = parse_xml(part.blob)
            except Exception:
                continue
            for kind in ("major", "minor"):
                font = root.find(f".//{{{A_NS}}}{kind}Font/{{{A_NS}}}latin")
                if font is not None:
                    face = (font.get("typeface") or "").strip()
                    if face:
                        result[kind] = face
            if result:
                break
        return result

    def style_name(self, style_id: str | None) -> str:
        if not style_id:
            return ""
        element = self.styles.get(style_id)
        if element is None:
            return style_id
        name = element.find(qn("w:name"))  # type: ignore[attr-defined]
        return (name.get(qn("w:val")) if name is not None else "") or style_id

    def paragraph_style_id(self, paragraph: Paragraph) -> str:
        ppr = paragraph._p.pPr
        return (
            _style_id_from_ppr(ppr)
            or self.default_paragraph_style_id
            or "Normal"
        )

    def run_style_id(self, run: Run) -> str | None:
        rpr = run._r.rPr
        return _style_id_from_rpr(rpr)

    def _style_chain(self, style_id: str | None) -> tuple[object, ...]:
        if not style_id:
            return ()
        cached = self._chain_cache.get(style_id)
        if cached is not None:
            return cached
        chain: list[object] = []
        seen: set[str] = set()
        current = style_id
        while current and current not in seen:
            seen.add(current)
            element = self.styles.get(current)
            if element is None:
                break
            chain.append(element)
            based_on = element.find(qn("w:basedOn"))  # type: ignore[attr-defined]
            current = based_on.get(qn("w:val")) if based_on is not None else None
        result = tuple(chain)
        self._chain_cache[style_id] = result
        return result

    def _style_ppr_elements(self, style_id: str | None) -> tuple[object, ...]:
        if not style_id:
            return ()
        cached = self._style_ppr_cache.get(style_id)
        if cached is not None:
            return cached
        elements: list[object] = []
        for style in reversed(self._style_chain(style_id)):
            ppr = style.find(qn("w:pPr"))  # type: ignore[attr-defined]
            if ppr is not None:
                elements.append(ppr)
        result = tuple(elements)
        self._style_ppr_cache[style_id] = result
        return result

    def _style_rpr_elements(self, style_id: str | None) -> tuple[object, ...]:
        if not style_id:
            return ()
        cached = self._style_rpr_cache.get(style_id)
        if cached is not None:
            return cached
        elements: list[object] = []
        for style in reversed(self._style_chain(style_id)):
            rpr = style.find(qn("w:rPr"))  # type: ignore[attr-defined]
            if rpr is not None:
                elements.append(rpr)
        result = tuple(elements)
        self._style_rpr_cache[style_id] = result
        return result

    # ------------------------------------------------------------------ helpers

    @staticmethod
    def _elements(*items: object | None) -> list[object]:
        return [item for item in items if item is not None]

    @staticmethod
    def _collect_attrs(elements: list[object], tag: str) -> dict[str, str]:
        result: dict[str, str] = {}
        for element in elements:
            for node in element.findall(qn(tag)):  # type: ignore[attr-defined]
                for key, value in node.attrib.items():
                    result[key] = value
        return result

    @staticmethod
    def _last_int_attr(elements: list[object], tag: str, attr: str = "w:val") -> int | None:
        value: int | None = None
        for element in elements:
            for node in element.findall(qn(tag)):  # type: ignore[attr-defined]
                raw = node.get(qn(attr))
                if raw is None:
                    continue
                try:
                    value = int(raw)
                except (TypeError, ValueError):
                    continue
        return value

    @staticmethod
    def _resolve_bool(elements: list[object], tag: str) -> bool:
        value = False
        for element in elements:
            for node in element.findall(qn(tag)):  # type: ignore[attr-defined]
                raw = node.get(qn("w:val"))
                value = raw not in {"0", "false", "off"}
        return value

    @staticmethod
    def _has_element(elements: list[object], tag: str) -> bool:
        return any(element.find(qn(tag)) is not None for element in elements)  # type: ignore[attr-defined]

    # -------------------------------------------------------------- paragraphs

    def resolve_ppr(self, style_id: str | None, direct_ppr: object | None) -> EffectiveParagraph:
        cache_key = style_id or ""
        base = self._ppr_cache.get(cache_key)
        if direct_ppr is None and base is not None:
            return base

        elements = self._elements(
            self.default_ppr,
            *self._style_ppr_elements(style_id),
            direct_ppr,
        )
        jc = self._collect_attrs(elements, "w:jc").get(qn("w:val"))
        spacing = self._collect_attrs(elements, "w:spacing")
        ind = self._collect_attrs(elements, "w:ind")
        outline = self._last_int_attr(elements, "w:outlineLvl")

        result = EffectiveParagraph(
            style_id=style_id or "",
            style_name=self.style_name(style_id),
            alignment=jc,
            line_rule=spacing.get(qn("w:lineRule")),
            line_value=_int_or_none(spacing.get(qn("w:line"))),
            before_twips=_int_or_none(spacing.get(qn("w:before"))) or 0,
            after_twips=_int_or_none(spacing.get(qn("w:after"))) or 0,
            first_line_twips=_int_or_none(ind.get(qn("w:firstLine"))),
            first_line_chars=_int_or_none(ind.get(qn("w:firstLineChars"))),
            hanging_twips=_int_or_none(ind.get(qn("w:hanging"))),
            left_twips=_int_or_none(ind.get(qn("w:left")))
            or _int_or_none(ind.get(qn("w:start")))
            or 0,
            right_twips=_int_or_none(ind.get(qn("w:right")))
            or _int_or_none(ind.get(qn("w:end")))
            or 0,
            page_break_before=self._resolve_bool(elements, "w:pageBreakBefore"),
            outline_level=outline,
            has_numbering=self._has_element(elements, "w:numPr"),
        )
        if direct_ppr is None:
            self._ppr_cache[cache_key] = result
        return result

    def paragraph_format(self, paragraph: Paragraph) -> EffectiveParagraph:
        style_id = self.paragraph_style_id(paragraph)
        return self.resolve_ppr(style_id, paragraph._p.pPr)

    # -------------------------------------------------------------------- runs

    def resolve_rpr(
        self,
        paragraph_style_id: str | None,
        character_style_id: str | None,
        direct_rpr: object | None,
    ) -> EffectiveFont:
        cache_key = f"{paragraph_style_id or ''}|{character_style_id or ''}"
        base = self._rpr_cache.get(cache_key)
        if direct_rpr is None and base is not None:
            return base

        elements = self._elements(
            self.default_rpr,
            *self._style_rpr_elements(paragraph_style_id),
            *self._style_rpr_elements(character_style_id),
            direct_rpr,
        )
        fonts = self._collect_attrs(elements, "w:rFonts")
        theme_ascii = self._theme_font(fonts.get(qn("w:asciiTheme")))
        theme_hansi = self._theme_font(fonts.get(qn("w:hAnsiTheme")))
        ascii_font = (
            fonts.get(qn("w:ascii"))
            or theme_ascii
            or fonts.get(qn("w:hAnsi"))
            or theme_hansi
            or "Calibri"
        )
        hansi_font = (
            fonts.get(qn("w:hAnsi"))
            or theme_hansi
            or fonts.get(qn("w:ascii"))
            or theme_ascii
            or "Calibri"
        )

        sz = self._collect_attrs(elements, "w:sz").get(qn("w:val"))
        if sz is None:
            sz = self._collect_attrs(elements, "w:szCs").get(qn("w:val"))
        size_pt = half_points_to_pt(_int_or_none(sz)) or 11.0

        color = self._collect_attrs(elements, "w:color").get(qn("w:val")) or "auto"
        underline = self._resolve_bool(elements, "w:u")
        if underline:
            u_attrs = self._collect_attrs(elements, "w:u")
            underline = (u_attrs.get(qn("w:val")) or "single") not in {"none", "0"}

        result = EffectiveFont(
            ascii_font=ascii_font.strip(),
            hansi_font=hansi_font.strip(),
            size_pt=size_pt,
            bold=self._resolve_bool(elements, "w:b")
            or self._resolve_bool(elements, "w:bCs"),
            italic=self._resolve_bool(elements, "w:i")
            or self._resolve_bool(elements, "w:iCs"),
            color=color,
            underline=underline,
        )
        if direct_rpr is None:
            self._rpr_cache[cache_key] = result
        return result

    def run_format(self, run: Run, paragraph_style_id: str | None) -> EffectiveFont:
        return self.resolve_rpr(
            paragraph_style_id,
            self.run_style_id(run),
            run._r.rPr,
        )

    def ppr_has_center_tab(self, style_id: str | None, direct_ppr: object | None) -> bool:
        elements = self._elements(
            self.default_ppr,
            *self._style_ppr_elements(style_id),
            direct_ppr,
        )
        for element in elements:
            for tabs in element.findall(qn("w:tabs")):  # type: ignore[attr-defined]
                for tab in tabs.findall(qn("w:tab")):
                    if tab.get(qn("w:val")) == "center":
                        return True
        return False

    def _theme_font(self, value: str | None) -> str | None:
        if not value:
            return None
        if "major" in value:
            return self.theme_fonts.get("major")
        if "minor" in value:
            return self.theme_fonts.get("minor")
        return self.theme_fonts.get("minor")

    def ppr_for_element(self, paragraph_element: object) -> EffectiveParagraph:
        ppr = paragraph_element.find(qn("w:pPr"))  # type: ignore[attr-defined]
        style_id = _style_id_from_ppr(ppr) or self.default_paragraph_style_id or "Normal"
        return self.resolve_ppr(style_id, ppr)


def _int_or_none(value: str | None) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _element_text(element: object) -> str:
    parts: list[str] = []
    for node in element.iter():  # type: ignore[attr-defined]
        if node.tag == qn("w:t") and node.text:
            parts.append(node.text)
        elif node.tag == qn("w:tab"):
            parts.append("\t")
        elif node.tag in {qn("w:br"), qn("w:cr")}:
            parts.append("\n")
    return "".join(parts)


class DocxModel:
    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        self.document = open_docx(str(self.path))
        self.styles = StyleResolver(self.document)
        self.blocks = self._build_blocks()
        self._assign_sections()

    # ------------------------------------------------------------------ blocks

    def _build_blocks(self) -> list[Block]:
        blocks: list[Block] = []
        for element in self._iter_block_elements(self.document.element.body):
            if element.tag == qn("w:p"):
                paragraph = Paragraph(element, self.document)
                text = _element_text(element)
                blocks.append(
                    Block(
                        index=len(blocks),
                        kind="paragraph",
                        element=element,
                        text=text,
                        paragraph=paragraph,
                    )
                )
            elif element.tag == qn("w:tbl"):
                table = Table(element, self.document)
                blocks.append(
                    Block(
                        index=len(blocks),
                        kind="table",
                        element=element,
                        text=_element_text(element),
                        table=table,
                    )
                )
        return blocks

    @classmethod
    def _iter_block_elements(cls, parent: object) -> Iterator[object]:
        for child in parent.iterchildren():  # type: ignore[attr-defined]
            if child.tag in {qn("w:p"), qn("w:tbl")}:
                yield child
                continue
            if child.tag in {
                qn("w:sdt"),
                qn("w:customXml"),
                qn("w:smartTag"),
            }:
                content = child.find(qn("w:sdtContent"))
                yield from cls._iter_block_elements(content if content is not None else child)

    def _assign_sections(self) -> None:
        current = 0
        for block in self.blocks:
            block.section_index = current
            if block.paragraph is not None:
                ppr = block.paragraph._p.pPr
                if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
                    current += 1

    # ------------------------------------------------------------------- texts

    def iter_runs(self, paragraph: Paragraph) -> Iterator[Run]:
        for run_element in paragraph._p.iter(qn("w:r")):
            if any(
                ancestor.tag in {qn("w:del"), qn("w:txbxContent")}
                for ancestor in run_element.iterancestors()
            ):
                continue
            yield Run(run_element, paragraph)

    def table_paragraphs(self, table: Table) -> Iterator[Paragraph]:
        seen: set[int] = set()
        for row in table.rows:
            for cell in row.cells:
                key = id(cell._tc)
                if key in seen:
                    continue
                seen.add(key)
                yield from cell.paragraphs

    def paragraph_font(self, paragraph: Paragraph, run: Run) -> EffectiveFont:
        return self.styles.run_format(run, self.styles.paragraph_style_id(paragraph))

    # ------------------------------------------------------------------ footer

    def footer_definitions(self) -> list[tuple[object, dict[str, object | None], dict[str, object | None]]]:
        """Возвращает (section, effective_footers, explicit_refs) по секциям."""
        related = self.document.part.related_parts
        effective: dict[str, object | None] = {"default": None, "first": None, "even": None}
        result: list[tuple[object, dict[str, object | None], dict[str, object | None]]] = []
        for section in self.document.sections:
            refs: dict[str, object | None] = {}
            for ref in section._sectPr.findall(qn("w:footerReference")):
                kind = ref.get(qn("w:type")) or "default"
                rid = ref.get(qn("r:id"))
                refs[kind] = related.get(rid) if rid else None
            for kind in effective:
                if kind in refs:
                    effective[kind] = refs[kind]
            result.append((section, dict(effective), refs))
        return result

    @staticmethod
    def footer_has_page_field(part: object | None) -> bool:
        if part is None:
            return False
        element = getattr(part, "element", None)
        if element is None:
            return False
        pattern = re.compile(r"(?<![A-Z])PAGE\b", re.IGNORECASE)
        for node in element.iter(qn("w:instrText")):
            if pattern.search(node.text or ""):
                return True
        for node in element.iter(qn("w:fldSimple")):
            if pattern.search(node.get(qn("w:instr")) or ""):
                return True
        return False

    @staticmethod
    def footer_page_paragraphs(part: object | None) -> list[object]:
        if part is None:
            return []
        element = getattr(part, "element", None)
        if element is None:
            return []
        pattern = re.compile(r"(?<![A-Z])PAGE\b", re.IGNORECASE)
        result: list[object] = []
        for paragraph in element.iter(qn("w:p")):
            has_page = any(
                pattern.search(node.text or "")
                for node in paragraph.iter(qn("w:instrText"))
            ) or any(
                pattern.search(node.get(qn("w:instr")) or "")
                for node in paragraph.iter(qn("w:fldSimple"))
            )
            if has_page:
                result.append(paragraph)
        return result

    @staticmethod
    def paragraph_has_page_break(paragraph: Paragraph) -> bool:
        ppr = paragraph._p.pPr
        if ppr is not None and ppr.find(qn("w:pageBreakBefore")) is not None:
            return True
        for node in paragraph._p.iter(qn("w:br")):
            if node.get(qn("w:type")) == "page":
                return True
        return False

    @staticmethod
    def paragraph_has_section_break(paragraph: Paragraph) -> bool:
        ppr = paragraph._p.pPr
        return ppr is not None and ppr.find(qn("w:sectPr")) is not None

    @staticmethod
    def element_has_image(element: object) -> bool:
        return (
            element.find(".//" + qn("w:drawing")) is not None  # type: ignore[attr-defined]
            or element.find(".//" + qn("w:pict")) is not None  # type: ignore[attr-defined]
            or element.find(".//" + qn("w:object")) is not None  # type: ignore[attr-defined]
        )
