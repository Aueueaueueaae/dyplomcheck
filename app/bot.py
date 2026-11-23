from __future__ import annotations

import asyncio
import logging
import sys

from aiogram import Bot, Dispatcher
from aiogram.client.default import DefaultBotProperties
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.enums import ParseMode
from aiogram.exceptions import (
    TelegramAPIError,
    TelegramConflictError,
    TelegramNetworkError,
    TelegramUnauthorizedError,
)
from aiogram.types import BotCommand

from .config import ConfigError, Settings
from .handlers import RateLimiter, router

logger = logging.getLogger(__name__)


def _setup_logging(level: str) -> None:
    logging.basicConfig(
        level=getattr(logging, level, logging.INFO),
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%d.%m.%Y %H:%M:%S",
    )
    logging.getLogger("aiogram.event").setLevel(logging.WARNING)


async def _set_commands(bot: Bot) -> None:
    await bot.set_my_commands(
        [
            BotCommand(command="start", description="Открыть меню проверки"),
        ]
    )
    try:
        await bot.set_my_description(
            description=(
                "📄 DyplomCheck — проверка оформления диплома по ГОСТ.\n\n"
                "Пришлите .docx: покажу ошибки в шрифте, полях, интервалах, "
                "заголовках, нумерации, рисунках, таблицах и списке литературы. "
                "Файл не исправляю."
            )
        )
        await bot.set_my_short_description(
            short_description=(
                "Проверка .docx по ГОСТ: шрифт, поля, интервалы, заголовки, литература"
            )
        )
    except TelegramUnauthorizedError:
        raise
    except TelegramAPIError as exc:
        logger.warning("Не удалось обновить описание бота: %s", exc)


async def main() -> None:
    try:
        settings = Settings.load()
    except ConfigError as exc:
        print(f"Ошибка конфигурации: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc

    _setup_logging(settings.log_level)
    logger.info("Запуск DyplomCheck")

    try:
        session = AiohttpSession(proxy=settings.bot_proxy)
    except Exception as exc:
        logger.error("Не удалось настроить прокси BOT_PROXY: %s", exc)
        raise SystemExit(1) from exc

    bot = Bot(
        token=settings.bot_token,
        session=session,
        default=DefaultBotProperties(parse_mode=ParseMode.HTML),
    )
    if settings.bot_proxy:
        logger.info("Используется прокси для Telegram API: %s", settings.bot_proxy)
    dispatcher = Dispatcher()
    dispatcher["settings"] = settings
    dispatcher["limiter"] = RateLimiter(
        limit=settings.rate_limit_checks,
        window_sec=settings.rate_limit_window_sec,
    )
    dispatcher.include_router(router)

    commands_set = False
    try:
        while True:
            try:
                if not commands_set:
                    await _set_commands(bot)
                    commands_set = True
                await dispatcher.start_polling(
                    bot,
                    allowed_updates=dispatcher.resolve_used_update_types(),
                    close_bot_session=False,
                )
                break
            except TelegramUnauthorizedError as exc:
                logger.error(
                    "Telegram не принял BOT_TOKEN. Проверьте токен в .env. Ответ: %s",
                    exc,
                )
                raise SystemExit(2) from exc
            except TelegramConflictError as exc:
                logger.error(
                    "Конфликт: этот бот уже запущен в другом процессе. "
                    "Остановите вторую копию. Ответ: %s",
                    exc,
                )
                raise SystemExit(2) from exc
            except TelegramNetworkError as exc:
                logger.error(
                    "Нет связи с api.telegram.org: %s\n"
                    "Включите VPN или укажите BOT_PROXY в .env. "
                    "Повтор через 15 секунд…",
                    exc,
                )
                await asyncio.sleep(15)
    finally:
        await bot.session.close()
        logger.info("Бот остановлен")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        pass
