"""FastAPI-сервис генерации ежемесячных отчётных книг Excel."""

from __future__ import annotations

import os
import re
from contextlib import asynccontextmanager
from datetime import date
from pathlib import Path
from typing import Optional
from urllib.parse import unquote

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from loguru import logger

from .config_loader import (
    AppConfig,
    PointConfig,
    infer_point_name_from_filename,
    load_config,
    save_points,
    scan_data_workbooks,
)
from .generator import ReportGenerator
from .holidays import HolidaysProvider
from .models import (
    AddPointManualRequest,
    GenerateRequest,
    GenerateResponse,
    PointInfo,
    PointResult,
)
from .utils import MONTH_NAMES_RU

config: AppConfig | None = None
generator: ReportGenerator | None = None

BASE_DIR = Path(__file__).resolve().parent
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

SAFE_FILENAME_RE = re.compile(r"^[\w.\- А-Яа-яЁё]+$", re.UNICODE)


def _setup_logging() -> None:
    log_dir = Path(os.environ.get("LOG_DIR", "logs"))
    log_dir.mkdir(parents=True, exist_ok=True)
    logger.add(
        log_dir / "katopro_{time:YYYY-MM-DD}.log",
        rotation="1 day",
        retention="14 days",
        encoding="utf-8",
    )


def _require_config() -> AppConfig:
    if config is None:
        raise HTTPException(status_code=503, detail="Сервис ещё не инициализирован")
    return config


def _list_generated_files() -> list[dict]:
    cfg = _require_config()
    out_dir = Path(cfg.output_dir)
    if not out_dir.exists():
        return []
    files = sorted(
        out_dir.glob("*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return [{"name": f.name, "size": f.stat().st_size} for f in files[:30]]


def _default_period() -> tuple[int, int]:
    today = date.today()
    if today.month == 12:
        return today.year + 1, 1
    return today.year, today.month + 1


def _point_info(point: PointConfig) -> PointInfo:
    path = Path(point.file_path)
    return PointInfo(
        name=point.name,
        file_path=point.file_path,
        file_exists=path.exists(),
        filename=path.name if path.name else None,
    )


def _serialize_points() -> list[PointInfo]:
    cfg = _require_config()
    return [_point_info(p) for p in cfg.points]


def _safe_upload_filename(original: str) -> str:
    name = Path(unquote(original)).name.strip()
    if not name.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Нужен файл Excel (.xlsx)")
    if not SAFE_FILENAME_RE.match(name) or ".." in name:
        raise HTTPException(
            status_code=400,
            detail="Некорректное имя файла. Используйте буквы, цифры, пробел, _ . -",
        )
    return name


def _upsert_point(name: str, file_path: Path, replace: bool = False) -> PointInfo:
    cfg = _require_config()
    existing = cfg.point_by_name(name)
    if existing and not replace:
        raise HTTPException(
            status_code=409,
            detail=f"Точка «{name}» уже есть. Включите замену или выберите другое имя.",
        )

    point = PointConfig(name=name, file_path=str(file_path.resolve()))
    if existing:
        cfg.points = [point if p.name.casefold() == name.casefold() else p for p in cfg.points]
    else:
        cfg.points.append(point)

    save_points(cfg)
    if generator is not None:
        generator.config = cfg
    logger.info("Точка «{}» → {}", name, file_path)
    return _point_info(point)


@asynccontextmanager
async def lifespan(_app: FastAPI):
    global config, generator
    _setup_logging()
    config = load_config()
    cache_dir = Path(os.environ.get("HOLIDAYS_CACHE_DIR", "/tmp/katopro_holidays"))
    holidays = HolidaysProvider(config.holidays_api, cache_dir=cache_dir)
    generator = ReportGenerator(config, holidays)
    logger.info(
        "KatoPro запущен: точек={}, каталог={}",
        len(config.points),
        config.output_dir,
    )
    yield


app = FastAPI(
    title="KatoPro",
    description="Автоматизация генерации ежемесячных отчётных книг Excel для торговых точек",
    version="1.0.0",
    lifespan=lifespan,
)

app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    cfg = _require_config()
    year, month = _default_period()
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "points": _serialize_points(),
            "data_files": scan_data_workbooks(cfg.data_dir),
            "months": list(MONTH_NAMES_RU.items()),
            "default_year": year,
            "default_month": month,
            "files": _list_generated_files(),
        },
    )


@app.get("/health")
def health():
    return {"status": "ok", "message": "Сервис работает"}


@app.get("/api/points", response_model=list[PointInfo])
def api_points():
    return _serialize_points()


@app.get("/api/data-files")
def api_data_files():
    cfg = _require_config()
    return scan_data_workbooks(cfg.data_dir)


@app.post("/api/points/upload", response_model=PointInfo)
async def upload_point(
    name: Optional[str] = Form(default=None),
    replace: bool = Form(default=False),
    file: UploadFile = File(...),
):
    """Добавить точку, загрузив Excel-файл."""
    cfg = _require_config()
    original = file.filename or "report.xlsx"
    filename = _safe_upload_filename(original)
    point_name = (name or "").strip() or infer_point_name_from_filename(filename)
    if not point_name:
        raise HTTPException(
            status_code=400,
            detail="Укажите название точки или загрузите файл вида МесОтч202603_Название.xlsx",
        )

    dest = cfg.data_dir / filename
    cfg.data_dir.mkdir(parents=True, exist_ok=True)
    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="Файл пустой")
    dest.write_bytes(content)
    return _upsert_point(point_name, dest, replace=replace)


@app.post("/api/points/manual", response_model=PointInfo)
def add_point_manual(body: AddPointManualRequest):
    """Добавить точку вручную, привязав уже существующий файл из data/."""
    cfg = _require_config()
    path = cfg.data_dir / body.filename
    if not path.exists() or not path.is_file():
        raise HTTPException(
            status_code=404,
            detail=f"Файл не найден в data/: {body.filename}",
        )
    return _upsert_point(body.name, path, replace=body.replace)


@app.delete("/api/points/{name}", response_model=dict)
def delete_point(name: str, delete_file: bool = False):
    cfg = _require_config()
    point = cfg.point_by_name(name)
    if point is None:
        raise HTTPException(status_code=404, detail=f"Точка «{name}» не найдена")

    cfg.points = [p for p in cfg.points if p.name.casefold() != name.strip().casefold()]
    save_points(cfg)
    if generator is not None:
        generator.config = cfg

    removed_file = False
    if delete_file:
        path = Path(point.file_path)
        try:
            path.resolve().relative_to(cfg.data_dir.resolve())
            if path.exists() and path.is_file():
                path.unlink()
                removed_file = True
        except ValueError:
            logger.warning("Файл точки вне data/, не удаляем: {}", path)

    return {"status": "ok", "name": name, "file_deleted": removed_file}


@app.get("/api/files")
def api_files():
    return _list_generated_files()


@app.get("/download/{filename}")
def download_file(filename: str):
    cfg = _require_config()
    if not SAFE_FILENAME_RE.match(filename) or ".." in filename or "/" in filename or "\\" in filename:
        raise HTTPException(status_code=400, detail="Некорректное имя файла")
    path = Path(cfg.output_dir) / filename
    if not path.exists() or not path.is_file():
        raise HTTPException(status_code=404, detail="Файл не найден")
    return FileResponse(
        path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/generate", response_model=GenerateResponse)
def generate(request: GenerateRequest):
    if generator is None:
        raise HTTPException(status_code=503, detail="Сервис ещё не инициализирован")
    cfg = _require_config()

    logger.info(
        "Запрос генерации: год={} месяц={} точки={}",
        request.year,
        request.month,
        request.points,
    )
    try:
        raw_results = generator.generate(request.year, request.month, request.points)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc

    results = [PointResult(**item) for item in raw_results]
    if not results:
        overall = "error"
    elif all(r.status == "ok" for r in results):
        overall = "ok"
    elif any(r.status == "ok" for r in results):
        overall = "partial_error"
    else:
        overall = "error"

    return GenerateResponse(
        status=overall,
        year=request.year,
        month=request.month,
        results=results,
    )


def create_app() -> FastAPI:
    return app
