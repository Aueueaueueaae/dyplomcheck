from __future__ import annotations

from pathlib import Path

import pytest

from app.analyzer import AnalysisError
from app.handlers import (
    HELP,
    RULES,
    UPLOAD,
    WELCOME,
    RateLimiter,
    _validate_archive,
    main_keyboard,
    submenu_keyboard,
)

from .docx_factory import make_correct_document


def test_rate_limiter_blocks_after_limit() -> None:
    limiter = RateLimiter(limit=2, window_sec=60)
    assert limiter.allow(1) == (True, 0)
    assert limiter.allow(1) == (True, 0)
    allowed, retry = limiter.allow(1)
    assert not allowed
    assert retry >= 1
    assert limiter.allow(2)[0] is True


def test_validate_archive_accepts_docx(tmp_path: Path) -> None:
    path = make_correct_document(tmp_path / "ok.docx")
    _validate_archive(path)


def test_validate_archive_rejects_plain_file(tmp_path: Path) -> None:
    path = tmp_path / "fake.docx"
    path.write_text("not a zip", encoding="utf-8")
    with pytest.raises(AnalysisError):
        _validate_archive(path)


def test_menu_texts_fit_telegram_and_have_buttons() -> None:
    for text in (WELCOME, UPLOAD, RULES, HELP):
        assert len(text) < 4096
        assert "<b>" in text

    buttons = [
        button
        for row in main_keyboard().inline_keyboard
        for button in row
    ]
    assert {button.callback_data for button in buttons} == {
        "menu:upload",
        "menu:rules",
        "menu:help",
    }
    button_texts = " ".join(button.text for button in buttons)
    assert "Отправить работу" in button_texts
    assert "ГОСТ и проверки" in button_texts
    assert "Инструкция" in button_texts
    assert len(submenu_keyboard().inline_keyboard) == 2
