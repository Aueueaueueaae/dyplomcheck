from __future__ import annotations

import base64
from io import BytesIO
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt

TINY_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)


def _set_style_font(style, name: str, size_pt: float, *, bold: bool | None = None) -> None:
    style.font.name = name
    style.font.size = Pt(size_pt)
    if bold is not None:
        style.font.bold = bold
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        rfonts.attrib.pop(qn(f"w:{attr}"), None)
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{attr}"), name)


def _set_normal(
    doc: Document,
    *,
    font: str = "Times New Roman",
    size: float = 14,
    line_spacing: float = 1.5,
    indent_cm: float = 1.25,
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
) -> None:
    style = doc.styles["Normal"]
    _set_style_font(style, font, size)
    pf = style.paragraph_format
    pf.alignment = alignment
    pf.line_spacing = line_spacing
    pf.first_line_indent = Cm(indent_cm)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.left_indent = Cm(0)
    pf.right_indent = Cm(0)

    heading = doc.styles["Heading 1"]
    _set_style_font(heading, "Times New Roman", 14, bold=True)
    heading.paragraph_format.space_before = Pt(0)
    heading.paragraph_format.space_after = Pt(0)
    heading.paragraph_format.keep_with_next = True


def _add_page_field(paragraph, *, tab_center: bool = False) -> None:
    if tab_center:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.tab_stops.add_tab_stop(
            Cm(8.25), WD_TAB_ALIGNMENT.CENTER
        )
        tab_run = paragraph.add_run()
        tab_run._r.append(OxmlElement("w:tab"))
    else:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.line_spacing = 1.0
    run = paragraph.add_run()
    run.font.name = "Times New Roman"
    run.font.size = Pt(14)

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for node in (begin, instr, separate, text, end):
        run._r.append(node)


def _add_heading(doc: Document, text: str, *, page_break: bool = True):
    paragraph = doc.add_paragraph(text, style="Heading 1")
    paragraph.paragraph_format.page_break_before = page_break
    for run in paragraph.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(14)
        run.font.bold = True
    return paragraph


def _add_body(doc: Document, text: str):
    paragraph = doc.add_paragraph(text)
    return paragraph


def make_document(
    path: Path,
    *,
    correct: bool = True,
    with_figure: bool = False,
    with_uncaptioned_table: bool = False,
    bad_citation: bool = False,
    manual_headings: bool = False,
    footer_tab_center: bool = False,
    with_direct_errors: bool = False,
    table_font_size: float = 14,
) -> Path:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Mm(30 if correct else 20)
    section.right_margin = Mm(15)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)
    section.different_first_page_header_footer = True

    _set_normal(
        doc,
        font="Times New Roman" if correct else "Calibri",
        size=14 if correct else 11,
        line_spacing=1.5 if correct else 1.0,
        indent_cm=1.25 if correct else 0,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY if correct else WD_ALIGN_PARAGRAPH.LEFT,
    )

    footer_p = section.footer.paragraphs[0]
    _add_page_field(footer_p, tab_center=footer_tab_center)

    title = doc.add_paragraph("МИНИСТЕРСТВО ОБРАЗОВАНИЯ И НАУКИ")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = doc.add_paragraph("ДИПЛОМНАЯ РАБОТА")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()

    toc = _add_heading(doc, "СОДЕРЖАНИЕ", page_break=False)
    toc.paragraph_format.first_line_indent = Cm(0)
    doc.add_paragraph("ВВЕДЕНИЕ … 3")
    doc.add_paragraph("1 Теоретическая часть … 4")

    if manual_headings:
        intro = doc.add_paragraph("ВВЕДЕНИЕ")
        intro.paragraph_format.page_break_before = True
        for run in intro.runs:
            run.font.bold = True
    else:
        _add_heading(doc, "ВВЕДЕНИЕ", page_break=True)

    _add_body(
        doc,
        "Актуальность темы исследования обусловлена необходимостью проверки "
        "оформления выпускных квалификационных работ [1].",
    )
    _add_body(
        doc,
        "В работе рассматриваются требования к тексту, таблицам, рисункам и "
        "библиографическому списку. Это достаточно длинный абзац для проверки "
        "основного текста.",
    )

    if with_figure:
        image_p = doc.add_paragraph()
        image_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        image_p.paragraph_format.first_line_indent = Cm(0)
        image_p.add_run().add_picture(BytesIO(TINY_PNG), width=Mm(20))
        caption = doc.add_paragraph("Рис. 1 - Схема работы")
        caption.alignment = WD_ALIGN_PARAGRAPH.LEFT
        caption.paragraph_format.first_line_indent = Cm(0)

    _add_body(
        doc,
        "Для примера приведём таблицу с результатами расчёта и поясним "
        "используемые показатели.",
    )
    if with_uncaptioned_table:
        table = doc.add_table(rows=2, cols=2)
    else:
        caption = doc.add_paragraph("Таблица 1 — Результаты расчёта")
        caption.alignment = WD_ALIGN_PARAGRAPH.LEFT
        caption.paragraph_format.first_line_indent = Cm(0)
        table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Показатель"
    table.cell(0, 1).text = "Значение"
    table.cell(1, 0).text = "Итог"
    table.cell(1, 1).text = "100"
    if table_font_size != 14:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(table_font_size)

    if with_direct_errors:
        broken = doc.add_paragraph(
            "Этот абзац имеет правильный стиль, но испорчен прямым форматированием."
        )
        broken.alignment = WD_ALIGN_PARAGRAPH.LEFT
        broken.paragraph_format.first_line_indent = Cm(0)
        broken.paragraph_format.line_spacing = 1.0
        for run in broken.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(11)

    if not correct:
        # Заголовок обычным текстом, без стиля.
        doc.add_paragraph("1 Теоретическая часть")
        doc.add_paragraph(
            "В этом разделе описаны основные подходы к оценке качества "
            "оформления документов и приведены дополнительные пояснения."
        )

    if bad_citation:
        _add_body(
            doc,
            "Дополнительно используется источник [9], которого нет в списке литературы.",
        )

    _add_heading(doc, "ЗАКЛЮЧЕНИЕ", page_break=True)
    _add_body(
        doc,
        "В результате работы были сформулированы требования к оформлению "
        "и подготовлен чек-лист для самопроверки.",
    )

    _add_heading(doc, "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", page_break=True)
    if correct:
        doc.add_paragraph(
            "1. Иванов И. И. Теория управления : учебник. – М. : Наука, 2020. – 256 с."
        )
    else:
        doc.add_paragraph(
            "1. Иванов И. И. Теория управления. – М. : Наука. – 256 с."
        )
        doc.add_paragraph(
            "3. Петров П. П. Управление качеством : учебник. – СПб. : Питер, 2021. – 120 с."
        )

    doc.save(path)
    return path


def make_correct_document(path: Path) -> Path:
    return make_document(path, correct=True)
