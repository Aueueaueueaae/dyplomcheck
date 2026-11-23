from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")


class ConfigError(RuntimeError):
    """Понятная ошибка конфигурации, которую можно показать в консоли."""


def _env_int(name: str, default: int, minimum: int, maximum: int) -> int:
    raw = os.getenv(name, "").strip()
    if not raw:
        return default
    try:
        value = int(raw)
    except ValueError as exc:
        raise ConfigError(f"{name} должно быть целым числом, получено: {raw!r}") from exc
    if not minimum <= value <= maximum:
        raise ConfigError(f"{name} должно быть в диапазоне {minimum}..{maximum}, получено: {value}")
    return value


@dataclass(frozen=True)
class Settings:
    bot_token: str
    bot_proxy: str | None = None
    max_file_mb: int = 20
    rate_limit_checks: int = 8
    rate_limit_window_sec: int = 600
    max_report_issues: int = 1000
    log_level: str = "INFO"

    @classmethod
    def load(cls) -> Settings:
        token = os.getenv("BOT_TOKEN", "").strip()
        if not token:
            raise ConfigError(
                "BOT_TOKEN не задан. Откройте .env и вставьте токен от @BotFather: BOT_TOKEN=123456:ABC..."
            )
        if ":" not in token:
            raise ConfigError("BOT_TOKEN выглядит некорректно. Нужен токен вида 123456789:AA...")

        proxy = os.getenv("BOT_PROXY", "").strip() or None
        if proxy and not proxy.startswith(("http://", "https://", "socks4://", "socks5://")):
            raise ConfigError(
                "BOT_PROXY должен начинаться с http://, https://, socks4:// или socks5://"
            )

        level = os.getenv("LOG_LEVEL", "INFO").strip().upper() or "INFO"
        if level not in {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}:
            level = "INFO"

        return cls(
            bot_token=token,
            bot_proxy=proxy,
            max_file_mb=_env_int("MAX_FILE_MB", 20, 1, 20),
            rate_limit_checks=_env_int("RATE_LIMIT_CHECKS", 8, 1, 1000),
            rate_limit_window_sec=_env_int("RATE_LIMIT_WINDOW_SEC", 600, 10, 86_400),
            max_report_issues=_env_int("MAX_REPORT_ISSUES", 1000, 50, 10_000),
            log_level=level,
        )
