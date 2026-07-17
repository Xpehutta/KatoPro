"""Получение праздничных дней РФ из внешних API с резервным списком."""

from __future__ import annotations

import json
import re
import time
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

import httpx
from loguru import logger

from .config_loader import HolidaysApiConfig

# Основные федеральные праздники (без учёта переносов) — резерв при недоступности API.
FIXED_HOLIDAYS: list[tuple[int, int]] = [
    (1, 1), (1, 2), (1, 3), (1, 4), (1, 5), (1, 6), (1, 7), (1, 8),
    (2, 23),
    (3, 8),
    (5, 1), (5, 9),
    (6, 12),
    (11, 4),
]


class HolidaysProvider:
    """Кеширует праздники на время TTL (по умолчанию 24 часа)."""

    def __init__(self, api_config: HolidaysApiConfig, cache_dir: Optional[Path] = None):
        self.api_config = api_config
        self.cache_dir = cache_dir or Path("/tmp/katopro_holidays")
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self._memory: dict[int, tuple[float, set[date]]] = {}

    def get_holidays(self, year: int) -> set[date]:
        now = time.time()
        cached = self._memory.get(year)
        if cached and now - cached[0] < self.api_config.cache_ttl_hours * 3600:
            return set(cached[1])

        disk = self._load_disk_cache(year)
        if disk is not None:
            holidays, ts = disk
            if now - ts < self.api_config.cache_ttl_hours * 3600:
                self._memory[year] = (ts, holidays)
                return set(holidays)

        holidays = self._fetch(year)
        self._memory[year] = (now, holidays)
        self._save_disk_cache(year, holidays, now)
        return set(holidays)

    def get_month_holidays(self, year: int, month: int) -> list[date]:
        return sorted(d for d in self.get_holidays(year) if d.month == month)

    def _fetch(self, year: int) -> set[date]:
        timeout = self.api_config.timeout_seconds
        urls = [
            self.api_config.url.format(year=year),
            self.api_config.fallback_url.format(year=year),
            f"https://date.nager.at/api/v3/PublicHolidays/{year}/RU",
        ]
        for url in urls:
            try:
                logger.info("Fetching holidays from {}", url)
                with httpx.Client(timeout=timeout, follow_redirects=True) as client:
                    response = client.get(url)
                    response.raise_for_status()
                    holidays = self._parse_response(url, response.text, year)
                    if holidays:
                        logger.info("Loaded {} holiday dates for {} from {}", len(holidays), year, url)
                        return holidays
            except Exception as exc:
                logger.warning("Holiday API failed ({}): {}", url, exc)

        logger.warning("All holiday APIs unavailable — using built-in fallback for {}", year)
        return self._fallback(year)

    def _parse_response(self, url: str, body: str, year: int) -> set[date]:
        if "xmlcalendar.ru" in url:
            return self._parse_xmlcalendar(body, year)
        if "isdayoff.ru" in url:
            return self._parse_isdayoff(body, year)
        if "nager.at" in url:
            return self._parse_nager(body)
        if "production-calendar" in url:
            return self._parse_production_calendar(body)
        # Heuristic by content
        body_strip = body.strip()
        if body_strip.startswith("{"):
            data = json.loads(body_strip)
            if "months" in data:
                return self._parse_xmlcalendar(body, year)
            if "holidays" in data:
                return self._parse_production_calendar(body)
        if re.fullmatch(r"[0-4]+", body_strip):
            return self._parse_isdayoff(body, year)
        if body_strip.startswith("["):
            return self._parse_nager(body)
        raise ValueError(f"Unknown holiday API response format from {url}")

    @staticmethod
    def _parse_xmlcalendar(body: str, year: int) -> set[date]:
        """
        xmlcalendar.ru: в days перечислены все нерабочие дни.
        Суффикс * — сокращённый рабочий (не праздник), + — перенесённый выходной.
        Для NETWORKDAYS и подсветки берём все дни без '*'.
        """
        data = json.loads(body)
        holidays: set[date] = set()
        for month_info in data.get("months", []):
            month = int(month_info["month"])
            raw_days = str(month_info.get("days", ""))
            if not raw_days:
                continue
            for token in raw_days.split(","):
                token = token.strip()
                if not token or token.endswith("*"):
                    continue  # pre-holiday shortened workday
                day = int(re.sub(r"\D", "", token))
                holidays.add(date(year, month, day))
        return holidays

    @staticmethod
    def _parse_isdayoff(body: str, year: int) -> set[date]:
        """Строка кодов: 1 — нерабочий день."""
        codes = body.strip()
        holidays: set[date] = set()
        current = date(year, 1, 1)
        for i, code in enumerate(codes):
            d = current + timedelta(days=i)
            if d.year != year:
                break
            if code in {"1", "2"}:  # 1 day off; treat 2 (short) as non-holiday
                if code == "1":
                    holidays.add(d)
        return holidays

    @staticmethod
    def _parse_nager(body: str) -> set[date]:
        data = json.loads(body)
        return {date.fromisoformat(item["date"]) for item in data}

    @staticmethod
    def _parse_production_calendar(body: str) -> set[date]:
        data = json.loads(body)
        holidays: set[date] = set()
        for item in data.get("holidays", []):
            if isinstance(item, str):
                holidays.add(date.fromisoformat(item[:10]))
            elif isinstance(item, dict) and "date" in item:
                holidays.add(date.fromisoformat(str(item["date"])[:10]))
        return holidays

    @staticmethod
    def _fallback(year: int) -> set[date]:
        result: set[date] = set()
        for month, day in FIXED_HOLIDAYS:
            try:
                result.add(date(year, month, day))
            except ValueError:
                continue
        return result

    def _cache_path(self, year: int) -> Path:
        return self.cache_dir / f"holidays_{year}.json"

    def _load_disk_cache(self, year: int) -> Optional[tuple[set[date], float]]:
        path = self._cache_path(year)
        if not path.exists():
            return None
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            holidays = {date.fromisoformat(d) for d in data["dates"]}
            return holidays, float(data["ts"])
        except Exception as exc:
            logger.warning("Failed to read holiday cache {}: {}", path, exc)
            return None

    def _save_disk_cache(self, year: int, holidays: set[date], ts: float) -> None:
        path = self._cache_path(year)
        payload = {
            "ts": ts,
            "dates": sorted(d.isoformat() for d in holidays),
        }
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
