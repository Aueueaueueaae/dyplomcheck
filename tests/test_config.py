from __future__ import annotations

import pytest

from app.config import ConfigError, Settings


def test_settings_reads_proxy(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("BOT_TOKEN", "123456:TEST")
    monkeypatch.delenv("BOT_PROXY", raising=False)
    assert Settings.load().bot_proxy is None

    monkeypatch.setenv("BOT_PROXY", "socks5://127.0.0.1:1080")
    assert Settings.load().bot_proxy == "socks5://127.0.0.1:1080"


def test_settings_rejects_bad_proxy(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("BOT_TOKEN", "123456:TEST")
    monkeypatch.setenv("BOT_PROXY", "127.0.0.1:1080")
    with pytest.raises(ConfigError):
        Settings.load()
