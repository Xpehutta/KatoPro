"""Вспомогательные функции для дат и имён листов."""

from __future__ import annotations

import calendar
import re
from datetime import date, datetime
from typing import Optional

MONTH_NAMES_RU = {
    1: "Январь",
    2: "Февраль",
    3: "Март",
    4: "Апрель",
    5: "Май",
    6: "Июнь",
    7: "Июль",
    8: "Август",
    9: "Сентябрь",
    10: "Октябрь",
    11: "Ноябрь",
    12: "Декабрь",
}

MONTH_NAME_TO_NUM = {name.lower(): num for num, name in MONTH_NAMES_RU.items()}

SHEET_NAME_RE = re.compile(
    r"^(Январь|Февраль|Март|Апрель|Май|Июнь|Июль|Август|Сентябрь|Октябрь|Ноябрь|Декабрь)"
    r"(\d{2})$",
    re.IGNORECASE,
)


def sheet_name_for(year: int, month: int) -> str:
    """Имя листа по шаблону: Март26."""
    return f"{MONTH_NAMES_RU[month]}{year % 100:02d}"


def parse_sheet_name(name: str) -> Optional[tuple[int, int]]:
    """Разбор имени листа → (year, month) или None."""
    match = SHEET_NAME_RE.match(name.strip())
    if not match:
        return None
    month = MONTH_NAME_TO_NUM[match.group(1).lower()]
    yy = int(match.group(2))
    year = 2000 + yy if yy < 70 else 1900 + yy
    return year, month


def previous_month(year: int, month: int) -> tuple[int, int]:
    if month == 1:
        return year - 1, 12
    return year, month - 1


def days_in_month(year: int, month: int) -> int:
    return calendar.monthrange(year, month)[1]


def month_dates(year: int, month: int) -> list[date]:
    return [date(year, month, day) for day in range(1, days_in_month(year, month) + 1)]


def is_weekend(d: date) -> bool:
    return d.weekday() >= 5


def to_date(value) -> Optional[date]:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None
