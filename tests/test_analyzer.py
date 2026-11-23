from __future__ import annotations

from pathlib import Path

from app.analyzer import Severity, analyze_docx
from app.analyzer.report import render_full_text, render_summary_html, report_filename

from .docx_factory import make_correct_document, make_document


def _ids(result) -> set[str]:
    return {issue.rule_id for issue in result.issues}


def test_correct_document_has_no_checked_violations(tmp_path: Path) -> None:
    path = make_correct_document(tmp_path / "correct.docx")
    result = analyze_docx(path, file_name="correct.docx")

    if result.issues:
        print("\n".join(f"{i.rule_id}: {i.found} -> {i.expected}" for i in result.issues))
    assert result.errors == 0
    assert result.warnings == 0
    assert result.source_count == 1
    assert result.citation_count == 1


def test_broken_document_reports_core_categories(tmp_path: Path) -> None:
    path = make_document(
        tmp_path / "broken.docx",
        correct=False,
        with_uncaptioned_table=True,
        bad_citation=True,
        manual_headings=True,
    )
    result = analyze_docx(path, file_name="broken.docx")
    ids = _ids(result)

    assert "page.margins" in ids
    assert "font.body" in ids
    assert "paragraph.format" in ids
    assert "heading.style" in ids
    assert "table.caption.missing" in ids
    assert "bibliography.numbering.order" in ids
    assert "citation.range" in ids
    assert result.errors > 0


def test_direct_formatting_overrides_style(tmp_path: Path) -> None:
    path = make_document(
        tmp_path / "direct.docx", correct=True, with_direct_errors=True
    )
    result = analyze_docx(path, file_name="direct.docx")
    ids = _ids(result)

    assert "font.body" in ids
    assert "paragraph.format" in ids


def test_table_font_12_is_allowed_but_10_is_not(tmp_path: Path) -> None:
    allowed = analyze_docx(
        make_document(tmp_path / "table12.docx", correct=True, table_font_size=12),
        file_name="table12.docx",
    )
    assert "font.table" not in _ids(allowed)

    too_small = analyze_docx(
        make_document(tmp_path / "table10.docx", correct=True, table_font_size=10),
        file_name="table10.docx",
    )
    assert "font.table" in _ids(too_small)


def test_page_number_centered_by_tab_is_not_an_error(tmp_path: Path) -> None:
    path = make_document(
        tmp_path / "footer-tab.docx", correct=True, footer_tab_center=True
    )
    result = analyze_docx(path, file_name="footer-tab.docx")

    assert "pagination.alignment" not in _ids(result)
    assert "pagination.missing" not in _ids(result)


def test_figure_caption_problems(tmp_path: Path) -> None:
    path = make_document(tmp_path / "figure.docx", correct=True, with_figure=True)
    result = analyze_docx(path, file_name="figure.docx")
    ids = _ids(result)

    assert "figure.caption.word" in ids
    assert "caption.dash" in ids
    assert result.figure_count == 1


def test_summary_and_full_report_are_safe_html_and_plain_text(tmp_path: Path) -> None:
    path = make_document(tmp_path / "broken.docx", correct=False, bad_citation=True)
    result = analyze_docx(path, file_name="<diplom>.docx")

    summary = render_summary_html(result)
    assert "<diplom>" not in summary
    assert "&lt;diplom&gt;" in summary
    assert len(summary) <= 3900

    full = render_full_text(result)
    assert "ОТЧЁТ О ПРОВЕРКЕ ОФОРМЛЕНИЯ" in full
    assert "Найдено:" in full
    assert "<diplom>.docx" in full
    assert report_filename("Моя работа.docx").startswith("otchet_")


def test_severity_values_are_stable() -> None:
    assert Severity.ERROR.label == "Ошибка"
    assert Severity.WARNING.icon == "🟠"
