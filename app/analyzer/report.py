from __future__ import annotations

import html
from datetime import datetime

from .models import AnalysisResult, Issue
from .units import fmt_num

SUMMARY_DIVIDER = "━━━━━━━━━━━━━━━━━━━━"
CATEGORY_ICONS = {
    "Поля страницы": "📐",
    "Шрифт": "🔤",
    "Абзац и интервал": "📏",
    "Заголовки": "📑",
    "Нумерация страниц": "🔢",
    "Рисунки": "🖼",
    "Таблицы": "📊",
    "Список литературы": "📚",
    "Ссылки в тексте": "🔗",
    "Структура": "🧩",
    "Прочее": "📌",
}


def _plural(n: int, one: str, few: str, many: str) -> str:
    n = abs(n) % 100
    if 11 <= n <= 19:
        return many
    n %= 10
    if n == 1:
        return one
    if 2 <= n <= 4:
        return few
    return many


def _issues_word(n: int) -> str:
    return _plural(n, "нарушение", "нарушения", "нарушений")


def _sources_word(n: int) -> str:
    return _plural(n, "источник", "источника", "источников")


def _issue_html(issue: Issue, number: int | None = None) -> str:
    prefix = f"{number}. " if number is not None else ""
    parts = [
        f"<b>{prefix}{issue.severity.icon} {html.escape(issue.title)}</b>",
        f"📍 {html.escape(issue.location)}",
        f"<b>Найдено:</b> {html.escape(issue.found)}",
        f"<b>Нужно:</b> {html.escape(issue.expected)}",
    ]
    if issue.quote:
        parts.append(f"💬 «{html.escape(issue.quote)}»")
    return "\n".join(parts)


def render_summary_html(result: AnalysisResult, max_issues: int = 6) -> str:
    has_issues = bool(result.errors or result.warnings or result.infos)
    lines = [
        f"<b>{'📊' if has_issues else '✅'} Проверка завершена</b>",
        f"<i>Файл: <code>{html.escape(result.file_name)}</code></i>",
        f"<i>Профиль: {html.escape(result.standard)}</i>",
        SUMMARY_DIVIDER,
        "",
    ]

    if has_issues:
        lines.append("<b>📊 Итог</b>")
        lines.append(f"🔴 Ошибки: <b>{result.errors}</b>")
        lines.append(f"🟠 Предупреждения: <b>{result.warnings}</b>")
        if result.infos:
            lines.append(f"🔵 Информация: <b>{result.infos}</b>")
    else:
        lines.append("✅ <b>Нарушений по проверяемым пунктам не найдено</b>")

    lines += [
        "",
        (
            f"📄 Абзацев: {result.paragraph_count} · 📊 Таблиц: {result.table_count} · "
            f"🖼 Рисунков: {result.figure_count} · 📚 Источников: {result.source_count}"
        ),
    ]

    categories = result.by_category()
    if categories:
        lines += ["", "<b>🗂 По категориям</b>"]
        for category, count in categories:
            icon = CATEGORY_ICONS.get(category, "📌")
            lines.append(f"• {icon} {html.escape(category)} — <b>{count}</b>")

    issues = result.sorted_issues()
    top: list[Issue] = []
    per_category: dict[str, int] = {}
    # В кратком отчёте показываем не шесть однотипных шрифтовых ошибок,
    # а по 1–2 самых важных пункта из разных категорий.
    for issue in issues:
        if per_category.get(issue.category, 0) >= 2:
            continue
        top.append(issue)
        per_category[issue.category] = per_category.get(issue.category, 0) + 1
        if len(top) >= max_issues:
            break
    if len(top) < max_issues:
        for issue in issues:
            if issue in top:
                continue
            top.append(issue)
            if len(top) >= max_issues:
                break

    if top:
        lines += ["", "<b>🛠 Что исправить в первую очередь</b>"]
        for index, issue in enumerate(top, start=1):
            lines += ["", _issue_html(issue, index)]

    if len(issues) > len(top):
        lines += [
            "",
            (
                f"➕ Ещё {len(issues) - len(top)} {_issues_word(len(issues) - len(top))} — "
                "в полном отчёте ниже."
            ),
        ]

    if result.truncated_count:
        lines += [
            "",
            (
                f"⚠️ Отчёт обрезан: ещё {result.truncated_count} "
                f"{_issues_word(result.truncated_count)} не поместились в лимит."
            ),
        ]

    lines += [
        "",
        SUMMARY_DIVIDER,
        "📎 <b>Полный отчёт</b> — в файле ниже.",
        "🚫 <i>Бот ничего не исправляет — только показывает места.</i>",
    ]
    text = "\n".join(lines)
    if len(text) > 3900:
        # Telegram ограничивает сообщение 4096 символами. Обрезаем по границе
        # абзаца, чтобы не разорвать HTML-тег.
        text = text[:3800].rsplit("\n", 1)[0]
        text += "\n\n…Полный список — в приложенном отчёте."
    return text


def render_full_text(result: AnalysisResult) -> str:
    lines = [
        "ОТЧЁТ О ПРОВЕРКЕ ОФОРМЛЕНИЯ",
        "=" * 72,
        f"Файл: {result.file_name}",
        f"Профиль: {result.standard}",
        f"Дата проверки: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        "",
        f"Абзацев: {result.paragraph_count}",
        f"Таблиц: {result.table_count}",
        f"Рисунков: {result.figure_count}",
        f"Источников: {result.source_count} ({_sources_word(result.source_count)})",
        "",
        f"Ошибок: {result.errors}",
        f"Предупреждений: {result.warnings}",
        f"Информационных пунктов: {result.infos}",
        "",
    ]

    if not result.issues:
        lines.append("Нарушений по проверяемым пунктам не найдено.")
    else:
        lines.append("НАРУШЕНИЯ")
        lines.append("-" * 72)
        current_category = None
        for index, issue in enumerate(result.sorted_issues(), start=1):
            if issue.category != current_category:
                current_category = issue.category
                lines += ["", f"### {current_category}", ""]
            lines.append(f"{index}. [{issue.severity.label.upper()}] {issue.title}")
            lines.append(f"   Место: {issue.location}")
            lines.append(f"   Найдено: {issue.found}")
            lines.append(f"   Нужно: {issue.expected}")
            if issue.quote:
                lines.append(f"   Цитата: «{issue.quote}»")
            if issue.hint:
                lines.append(f"   Подсказка: {issue.hint}")
            if issue.rule_ref:
                lines.append(f"   Норма: {issue.rule_ref}")
            lines.append("")

    if result.truncated_count:
        lines += [
            "",
            (
                "ВНИМАНИЕ: отчёт обрезан. Не показано ещё "
                f"{result.truncated_count} {_issues_word(result.truncated_count)}."
            ),
        ]

    if result.skipped:
        lines += [
            "",
            "ЧТО НЕ ПРОВЕРЯЛОСЬ",
            "-" * 72,
        ]
        for item in result.skipped:
            lines.append(f"• {item.title}: {item.reason}")

    lines += [
        "",
        "=" * 72,
        "Бот проверяет оформление и ничего не исправляет в файле.",
        "Отчёт не гарантирует прохождение нормоконтроля: требования кафедры",
        "могут отличаться от выбранного профиля ГОСТ.",
        "",
    ]
    return "\n".join(lines)


def report_filename(original_name: str) -> str:
    stem = original_name.rsplit(".", 1)[0].strip() or "diplom"
    safe = "".join(ch if ch.isalnum() or ch in "-_ " else "_" for ch in stem)
    safe = "_".join(safe.split())[:80] or "diplom"
    return f"otchet_{safe}.txt"


def seconds_label(value: float) -> str:
    return f"{fmt_num(value, 1)} с"
