from __future__ import annotations

import asyncio
import html
import logging
import shutil
import tempfile
import time
import zipfile
from collections import deque
from math import ceil
from pathlib import Path

from aiogram import Bot, F, Router
from aiogram.enums import ChatAction
from aiogram.exceptions import TelegramBadRequest
from aiogram.filters import CommandStart
from aiogram.types import (
    BufferedInputFile,
    CallbackQuery,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
)
from aiogram.utils.chat_action import ChatActionSender

from .analyzer import AnalysisError, analyze_docx
from .analyzer.report import render_full_text, render_summary_html, report_filename
from .config import Settings

logger = logging.getLogger(__name__)
router = Router(name="main")

DIVIDER = "━━━━━━━━━━━━━━━━━━━━"

WELCOME = (
    "<b>📄 DyplomCheck</b>\n"
    "<i>Проверка оформления диплома по ГОСТ</i>\n"
    f"{DIVIDER}\n\n"
    "<b>✨ Что умею</b>\n"
    "📐 поля и формат страницы\n"
    "🔤 шрифт и размер текста\n"
    "📏 интервалы и отступы\n"
    "📑 заголовки и разрывы страниц\n"
    "🔢 нумерацию страниц\n"
    "🖼 подписи рисунков и таблиц\n"
    "📚 список литературы и ссылки [1]\n\n"
    "<b>🚀 Как проверить</b>\n"
    "1️⃣ Нажмите «📎 Отправить работу»\n"
    "2️⃣ Прикрепите <b>.docx</b> до 20 МБ\n"
    "3️⃣ Получите краткий разбор и полный отчёт\n\n"
    "<b>⚠️ Важно</b>\n"
    "🚫 Я не исправляю файл — только показываю, где ошибка\n"
    "📄 Работаю только с .docx: в PDF и .doc нет стилей и шрифтов\n\n"
    f"{DIVIDER}\n"
    "<b>Выберите действие ниже</b> 👇"
)

UPLOAD = (
    "<b>📎 Отправить работу</b>\n"
    f"{DIVIDER}\n\n"
    "<b>Как прикрепить файл</b>\n"
    "1️⃣ Нажмите скрепку 📎 справа от поля ввода\n"
    "2️⃣ Выберите «Файл»\n"
    "3️⃣ Отправьте <b>.docx</b> до 20 МБ\n\n"
    "<b>Что подходит</b>\n"
    "✅ Документ Word <b>.docx</b>\n"
    "✅ Обычный файл без пароля\n\n"
    "<b>Что не подходит</b>\n"
    "❌ PDF и старый .doc\n"
    "❌ Архивы и файлы, защищённые паролем\n"
    "❌ Текст, отправленный сообщением\n\n"
    "💡 Кнопка меню не может открыть проводник — файл нужно прикрепить "
    "скрепкой в Telegram.\n\n"
    "<i>После отправки проверка начнётся автоматически.</i>"
)

RULES = (
    "<b>📋 Что проверяется</b>\n"
    "<i>Профиль: ГОСТ 7.32-2017 + кегль 14 пт из ТЗ</i>\n"
    f"{DIVIDER}\n\n"
    "<b>📐 Страница</b>\n"
    "• A4\n"
    "• поля: левое 30 мм, правое 15, верхнее и нижнее 20\n\n"
    "<b>🔤 Основной текст</b>\n"
    "• Times New Roman, 14 пт\n"
    "• интервал 1,5\n"
    "• отступ первой строки 1,25 см\n"
    "• выравнивание по ширине\n"
    "• без лишних интервалов до/после абзаца\n\n"
    "<b>📑 Заголовки и структура</b>\n"
    "• стили Word\n"
    "• без точки в конце\n"
    "• разделы с новой страницы\n"
    "• СОДЕРЖАНИЕ, ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ\n\n"
    "<b>🔢 Нумерация страниц</b>\n"
    "• внизу по центру\n"
    "• на титульном листе номера нет\n\n"
    "<b>🖼 Рисунки</b>\n"
    "• «Рисунок N — Название»\n"
    "• под рисунком, по центру\n\n"
    "<b>📊 Таблицы</b>\n"
    "• «Таблица N — Название»\n"
    "• над таблицей, слева, без отступа\n\n"
    "<b>📚 Литература и ссылки</b>\n"
    "• отдельный нумерованный раздел\n"
    "• базовая структура записи по ГОСТ Р 7.0.100-2018 / 7.0.5-2008\n"
    "• ссылки [N] не выходят за пределы списка\n\n"
    f"{DIVIDER}\n"
    "⚠️ Методичка кафедры может отличаться — тогда часть замечаний "
    "трактуется по ней."
)

HELP = (
    "<b>📖 Как пользоваться ботом</b>\n"
    f"{DIVIDER}\n\n"
    "<b>1️⃣ Подготовьте файл</b>\n"
    "Сохраните работу в <b>.docx</b>. PDF и .doc не подойдут.\n"
    "Размер — до 20 МБ, без пароля.\n\n"
    "<b>2️⃣ Отправьте боту</b>\n"
    "Нажмите «📎 Отправить работу» и прикрепите файл скрепкой 📎.\n\n"
    "<b>3️⃣ Заберите отчёт</b>\n"
    "Сначала придёт краткий список ошибок, затем полный отчёт .txt.\n\n"
    "<b>📊 Как читать отчёт</b>\n"
    "🔴 Ошибка — нужно исправить\n"
    "🟠 Предупреждение — проверьте вручную\n"
    "📍 Номер абзаца — место в документе\n"
    "💬 Цитата — фрагмент, где проблема\n\n"
    "<b>⚠️ Что важно</b>\n"
    "• Бот не исправляет файл, только показывает ошибки\n"
    "• Точные страницы не определяет — указывает абзац\n"
    "• Не проверяет смысл, орфографию и антиплагиат\n"
    "• Файл удаляется после проверки\n\n"
    f"{DIVIDER}\n"
    "💡 Если файл не отправляется — проверьте формат, размер и пароль."
)


def main_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📎 Отправить работу", callback_data="menu:upload"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📚 ГОСТ и проверки", callback_data="menu:rules"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📖 Инструкция", callback_data="menu:help"
                )
            ],
        ]
    )


def submenu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📎 Отправить работу", callback_data="menu:upload"
                )
            ],
            [
                InlineKeyboardButton(text="⬅️ В меню", callback_data="menu:main"),
            ],
        ]
    )


def after_check_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📎 Проверить ещё", callback_data="menu:upload"
                )
            ],
            [
                InlineKeyboardButton(text="🏠 В меню", callback_data="menu:main"),
            ],
        ]
    )


class RateLimiter:
    def __init__(self, limit: int, window_sec: int) -> None:
        self.limit = limit
        self.window_sec = window_sec
        self._events: dict[int, deque[float]] = {}

    def allow(self, user_id: int) -> tuple[bool, int]:
        now = time.monotonic()
        events = self._events.setdefault(user_id, deque())
        while events and now - events[0] >= self.window_sec:
            events.popleft()
        if len(events) >= self.limit:
            retry = max(1, ceil(self.window_sec - (now - events[0])))
            return False, retry
        events.append(now)
        return True, 0


def _validate_archive(path: Path, max_uncompressed_mb: int = 250) -> None:
    if not zipfile.is_zipfile(path):
        raise AnalysisError(
            "Файл не похож на .docx. Похоже, это старый .doc, PDF или повреждённый файл."
        )
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            if "[Content_Types].xml" not in names or "word/document.xml" not in names:
                raise AnalysisError(
                    "В файле нет структуры Word .docx. Сохраните документ как «Документ Word (*.docx)»."
                )
            total = sum(item.file_size for item in archive.infolist())
            if total > max_uncompressed_mb * 1024 * 1024:
                raise AnalysisError(
                    "Документ слишком большой после распаковки. Упростите файл или удалите лишние изображения."
                )
    except zipfile.BadZipFile as exc:
        raise AnalysisError("Не удалось прочитать .docx: архив повреждён.") from exc


@router.message(CommandStart())
async def cmd_start(message: Message) -> None:
    await message.answer(
        WELCOME,
        reply_markup=main_keyboard(),
        disable_web_page_preview=True,
    )


@router.callback_query(F.data.startswith("menu:"))
async def menu_callback(callback: CallbackQuery) -> None:
    data = callback.data or "menu:main"
    if data == "menu:upload":
        text, keyboard = UPLOAD, submenu_keyboard()
    elif data == "menu:rules":
        text, keyboard = RULES, submenu_keyboard()
    elif data == "menu:help":
        text, keyboard = HELP, submenu_keyboard()
    else:
        text, keyboard = WELCOME, main_keyboard()

    await callback.answer()
    if isinstance(callback.message, Message):
        try:
            await callback.message.edit_text(
                text,
                reply_markup=keyboard,
                disable_web_page_preview=True,
            )
            return
        except TelegramBadRequest as exc:
            if "message is not modified" in str(exc).lower():
                return

    await callback.bot.send_message(
        callback.from_user.id,
        text,
        reply_markup=keyboard,
        disable_web_page_preview=True,
    )


@router.message(F.document)
async def handle_document(
    message: Message,
    bot: Bot,
    settings: Settings,
    limiter: RateLimiter,
) -> None:
    document = message.document
    if document is None:
        return

    name = document.file_name or "diplom.docx"
    suffix = Path(name).suffix.casefold()
    if suffix != ".docx":
        await message.reply(
            "<b>❌ Формат не подходит</b>\n\n"
            "Нужен файл <b>.docx</b>. PDF, .doc и архивы не подходят: "
            "в них не сохраняются шрифты, стили и параметры абзацев.\n\n"
            "💡 В Word: <i>Файл → Сохранить как → Документ Word (*.docx)</i>"
        )
        return

    max_bytes = settings.max_file_mb * 1024 * 1024
    if document.file_size and document.file_size > max_bytes:
        await message.reply(
            f"<b>⚠️ Файл слишком большой</b>\n\n"
            f"Лимит — <b>{settings.max_file_mb} МБ</b>. "
            "Уменьшите документ или удалите лишние изображения."
        )
        return

    user_id = message.from_user.id if message.from_user else message.chat.id
    allowed, retry_after = limiter.allow(user_id)
    if not allowed:
        await message.reply(
            "<b>⏳ Слишком много проверок подряд</b>\n\n"
            f"Попробуйте через <b>{retry_after} с</b>."
        )
        return

    status = await message.answer(
        "🔍 <b>Проверяю документ…</b>\n"
        "<i>Обычно это занимает несколько секунд.</i>"
    )
    temp_dir = Path(tempfile.mkdtemp(prefix="dyplomcheck_"))
    path = temp_dir / "upload.docx"
    try:
        async with ChatActionSender(bot=bot, chat_id=message.chat.id, action=ChatAction.TYPING):
            await bot.download(document, destination=path, timeout=120)
            if path.stat().st_size > max_bytes:
                raise AnalysisError(
                    f"Файл больше {settings.max_file_mb} МБ. Уменьшите размер документа."
                )
            _validate_archive(path)
            result = await asyncio.to_thread(
                analyze_docx,
                path,
                file_name=name,
                max_issues=settings.max_report_issues,
            )

        await _safe_delete(status)
        summary = render_summary_html(result)
        await message.answer(
            summary,
            reply_markup=after_check_keyboard(),
            disable_web_page_preview=True,
        )

        full_report = render_full_text(result).encode("utf-8-sig")
        if result.errors or result.warnings:
            caption = (
                f"📎 <b>Полный отчёт</b>\n"
                f"🔴 Ошибок: {result.errors} · 🟠 Предупреждений: {result.warnings}\n"
                "<i>Бот не исправляет файл — только показывает, что не по ГОСТ.</i>"
            )
        else:
            caption = (
                "📎 <b>Полный отчёт</b>\n"
                "✅ Нарушений по проверяемым пунктам не найдено."
            )
        await message.answer_document(
            BufferedInputFile(full_report, filename=report_filename(name)),
            caption=caption,
        )
        logger.info(
            "Проверен %s: ошибок=%s, предупреждений=%s, за %.2f с",
            name,
            result.errors,
            result.warnings,
            result.duration_sec,
        )
    except AnalysisError as exc:
        logger.warning("Ошибка разбора %s: %s", name, exc.detail or exc.user_message)
        await _safe_edit(
            status,
            "<b>❌ Не удалось проверить файл</b>\n\n"
            f"{html.escape(exc.user_message)}",
        )
    except Exception:
        logger.exception("Не удалось обработать документ %s", name)
        await _safe_edit(
            status,
            "<b>❌ Не удалось проверить файл</b>\n\n"
            "Убедитесь, что это обычный .docx без пароля и повреждений, "
            "и попробуйте ещё раз.",
        )
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


async def _safe_delete(message: Message) -> None:
    try:
        await message.delete()
    except Exception:
        logger.debug("Не удалось удалить служебное сообщение", exc_info=True)


async def _safe_edit(message: Message, text: str) -> None:
    try:
        await message.edit_text(text, disable_web_page_preview=True)
    except Exception:
        try:
            await message.answer(text, disable_web_page_preview=True)
        except Exception:
            logger.exception("Не удалось отправить сообщение об ошибке")


@router.message(F.text)
async def fallback_text(message: Message) -> None:
    await message.answer(
        "<b>📄 Нужен файл .docx</b>\n\n"
        "Пришлите диплом как <b>документ Word</b> — не текстом и не PDF.\n\n"
        "💡 Откройте меню командой /start."
    )


@router.message()
async def fallback_other(message: Message) -> None:
    await message.answer(
        "<b>📎 Пришлите документ</b>\n\n"
        "Я проверяю только файлы <b>.docx</b>. "
        "Отправьте работу как документ — и я начну проверку."
    )
