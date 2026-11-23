from __future__ import annotations

import re
from dataclasses import dataclass, field
from enum import Enum


class AnalysisError(Exception):
    """Ошибка, которую можно показать пользователю без технических деталей."""

    def __init__(self, user_message: str, *, detail: str = "") -> None:
        super().__init__(user_message)
        self.user_message = user_message
        self.detail = detail


class Severity(str, Enum):
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"

    @property
    def label(self) -> str:
        return {
            Severity.ERROR: "Ошибка",
            Severity.WARNING: "Предупреждение",
            Severity.INFO: "Информация",
        }[self]

    @property
    def icon(self) -> str:
        return {
            Severity.ERROR: "🔴",
            Severity.WARNING: "🟠",
            Severity.INFO: "🔵",
        }[self]


CATEGORY_ORDER = (
    "Поля страницы",
    "Шрифт",
    "Абзац и интервал",
    "Заголовки",
    "Нумерация страниц",
    "Рисунки",
    "Таблицы",
    "Список литературы",
    "Ссылки в тексте",
    "Структура",
    "Прочее",
)


@dataclass(frozen=True)
class Issue:
    rule_id: str
    category: str
    severity: Severity
    title: str
    location: str
    found: str
    expected: str
    quote: str = ""
    hint: str = ""
    rule_ref: str = ""

    def sort_key(self) -> tuple[int, int, int, str]:
        severity_rank = {Severity.ERROR: 0, Severity.WARNING: 1, Severity.INFO: 2}
        try:
            category_rank = CATEGORY_ORDER.index(self.category)
        except ValueError:
            category_rank = len(CATEGORY_ORDER)
        match = re.search(r"(?:абзац|секция)\s+(\d+)", self.location, re.IGNORECASE)
        location_number = int(match.group(1)) if match else 0
        return severity_rank[self.severity], category_rank, location_number, self.location


@dataclass(frozen=True)
class SkippedCheck:
    title: str
    reason: str


@dataclass
class AnalysisResult:
    file_name: str
    standard: str
    issues: list[Issue] = field(default_factory=list)
    skipped: list[SkippedCheck] = field(default_factory=list)
    paragraph_count: int = 0
    table_count: int = 0
    figure_count: int = 0
    source_count: int = 0
    citation_count: int = 0
    duration_sec: float = 0.0
    truncated_count: int = 0

    @property
    def errors(self) -> int:
        return sum(1 for issue in self.issues if issue.severity is Severity.ERROR)

    @property
    def warnings(self) -> int:
        return sum(1 for issue in self.issues if issue.severity is Severity.WARNING)

    @property
    def infos(self) -> int:
        return sum(1 for issue in self.issues if issue.severity is Severity.INFO)

    def by_category(self) -> list[tuple[str, int]]:
        counts: dict[str, int] = {}
        for issue in self.issues:
            counts[issue.category] = counts.get(issue.category, 0) + 1
        return sorted(
            counts.items(),
            key=lambda item: (
                CATEGORY_ORDER.index(item[0]) if item[0] in CATEGORY_ORDER else len(CATEGORY_ORDER),
                item[0],
            ),
        )

    def sorted_issues(self) -> list[Issue]:
        return sorted(self.issues, key=Issue.sort_key)
