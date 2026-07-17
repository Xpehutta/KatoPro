"""Загрузка и валидация конфигурации YAML."""

from __future__ import annotations

import os
import re
import unicodedata
from pathlib import Path
from typing import Optional

import yaml
from loguru import logger
from pydantic import BaseModel, Field, PrivateAttr

REPORT_NAME_RE = re.compile(
    r"^МесОтч\d{6}_(.+)\.xlsx$",
    re.IGNORECASE,
)


class PointConfig(BaseModel):
    name: str
    file_path: str


class HolidaysApiConfig(BaseModel):
    url: str = "https://xmlcalendar.ru/data/ru/{year}/calendar.json"
    fallback_url: str = "https://isdayoff.ru/api/getdata?year={year}"
    timeout_seconds: int = 10
    cache_ttl_hours: int = 24


class AppConfig(BaseModel):
    points: list[PointConfig] = Field(default_factory=list)
    output_dir: str = "./generated"
    protection_password: Optional[str] = "secret"
    holidays_api: HolidaysApiConfig = Field(default_factory=HolidaysApiConfig)

    _config_path: Optional[Path] = PrivateAttr(default=None)
    _points_path: Optional[Path] = PrivateAttr(default=None)
    _base_dir: Optional[Path] = PrivateAttr(default=None)

    def point_by_name(self, name: str) -> Optional[PointConfig]:
        needle = name.strip().casefold()
        for point in self.points:
            if point.name.strip().casefold() == needle:
                return point
        return None

    @property
    def data_dir(self) -> Path:
        base = self._base_dir or Path.cwd()
        return (base / "data").resolve()

    @property
    def trash_dir(self) -> Path:
        base = self._base_dir or Path.cwd()
        return (base / "trash").resolve()

    @property
    def points_path(self) -> Path:
        if self._points_path:
            return self._points_path
        return self.data_dir / "points.yaml"


def resolve_path(path: str, base_dir: Path) -> Path:
    p = Path(path)
    if p.is_absolute():
        return p
    return (base_dir / p).resolve()


def infer_point_name_from_filename(filename: str) -> Optional[str]:
    """МесОтч202603_Смола.xlsx → Смола"""
    name = unicodedata.normalize("NFC", Path(filename).name)
    match = REPORT_NAME_RE.match(name)
    if match:
        return match.group(1).strip()
    stem = Path(name).stem.strip()
    return stem or None


def get_config_path(config_path: Optional[str] = None) -> Path:
    if config_path is None:
        config_path = os.environ.get("CONFIG_PATH", "config.yaml")
    path = Path(config_path)
    if path.is_absolute():
        return path
    candidates = [Path.cwd() / path, Path(__file__).resolve().parent.parent / path]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


def _to_storage_path(abs_path: Path, base_dir: Path) -> str:
    """Сохраняем относительный путь, если файл внутри проекта."""
    try:
        rel = abs_path.resolve().relative_to(base_dir.resolve())
        return f"./{rel.as_posix()}"
    except ValueError:
        return str(abs_path.resolve())


def load_points_file(points_path: Path, base_dir: Path) -> list[PointConfig]:
    if not points_path.exists():
        return []
    with points_path.open("r", encoding="utf-8") as fh:
        raw = yaml.safe_load(fh) or {}
    items = raw.get("points", []) if isinstance(raw, dict) else raw
    points: list[PointConfig] = []
    for item in items or []:
        point = PointConfig.model_validate(item)
        point.file_path = str(resolve_path(point.file_path, base_dir))
        points.append(point)
    return points


def save_points(config: AppConfig) -> Path:
    """Сохраняет список точек в data/points.yaml (относительные пути)."""
    points_path = config.points_path
    points_path.parent.mkdir(parents=True, exist_ok=True)
    base = config._base_dir or points_path.parent.parent
    payload = {
        "points": [
            {
                "name": p.name,
                "file_path": _to_storage_path(Path(p.file_path), base),
            }
            for p in config.points
        ]
    }
    with points_path.open("w", encoding="utf-8") as fh:
        yaml.safe_dump(payload, fh, allow_unicode=True, sort_keys=False)
    logger.info("Сохранены торговые точки: {} → {}", len(config.points), points_path)
    return points_path


def load_config(config_path: Optional[str] = None) -> AppConfig:
    """Загружает config.yaml и динамический список точек из data/points.yaml."""
    path = get_config_path(config_path)
    if not path.exists():
        raise FileNotFoundError(f"Конфиг не найден: {path}")

    with path.open("r", encoding="utf-8") as fh:
        raw = yaml.safe_load(fh) or {}

    config = AppConfig.model_validate(raw)
    base = path.parent
    config._config_path = path
    config._base_dir = base
    config.output_dir = str(resolve_path(config.output_dir, base))

    points_path = base / "data" / "points.yaml"
    config._points_path = points_path

    # Динамический список точек имеет приоритет; если его нет — берём из config.yaml и сидируем
    dynamic = load_points_file(points_path, base)
    if dynamic:
        config.points = dynamic
    else:
        for point in config.points:
            point.file_path = str(resolve_path(point.file_path, base))
        if config.points:
            save_points(config)

    return config


def scan_data_workbooks(data_dir: Path) -> list[dict]:
    """Список Excel-файлов в data/ с предполагаемым именем точки."""
    if not data_dir.exists():
        return []
    result = []
    for path in sorted(data_dir.glob("*.xlsx")):
        result.append(
            {
                "filename": path.name,
                "path": str(path.resolve()),
                "suggested_name": infer_point_name_from_filename(path.name),
            }
        )
    return result
