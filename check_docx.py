"""Проверка .docx без Telegram: удобно для демо и отладки.

Пример:
    .\\.venv\\Scripts\\python.exe check_docx.py "Моя работа.docx"
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from app.analyzer import AnalysisError, analyze_docx
from app.analyzer.report import render_full_text


def main() -> int:
    parser = argparse.ArgumentParser(description="Проверить оформление .docx по ГОСТ")
    parser.add_argument("file", type=Path, help="путь к .docx")
    parser.add_argument(
        "-o",
        "--out",
        type=Path,
        help="куда сохранить полный отчёт .txt (по умолчанию только вывод в консоль)",
    )
    args = parser.parse_args()

    if args.file.suffix.casefold() != ".docx":
        print("Нужен файл .docx", file=sys.stderr)
        return 2
    if not args.file.exists():
        print(f"Файл не найден: {args.file}", file=sys.stderr)
        return 2

    try:
        result = analyze_docx(args.file, file_name=args.file.name)
    except AnalysisError as exc:
        print(f"Ошибка: {exc.user_message}", file=sys.stderr)
        if exc.detail:
            print(f"Детали: {exc.detail}", file=sys.stderr)
        return 1

    report = render_full_text(result)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    print(report)
    if args.out:
        args.out.write_text(report, encoding="utf-8-sig")
        print(f"\nОтчёт сохранён: {args.out}")

    return 0 if result.errors == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
