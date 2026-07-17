"""FastAPI-сервис генерации ежемесячных отчётных книг Excel."""

from __future__ import annotations

import os
import re
import shutil
import unicodedata
from contextlib import asynccontextmanager
from datetime import date, datetime
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
    ExistingOutputFile,
    GenerateConflictResponse,
    GenerateRequest,
    GenerateResponse,
    PointInfo,
    PointResult,
    UploadBatchResponse,
    UploadItemResult,
)
from .utils import MONTH_NAMES_RU, sheet_name_for

config: AppConfig | None = None
generator: ReportGenerator | None = None

BASE_DIR = Path(__file__).resolve().parent
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

SAFE_FILENAME_RE = re.compile(r"^[\w.\- А-Яа-яЁё]+$", re.UNICODE)
PERIOD_IN_NAME_RE = re.compile(r"^МесОтч(\d{6})_", re.IGNORECASE)


def _normalize_text(value: str) -> str:
    """macOS часто отдаёт NFD (й = и + знак); приводим к NFC."""
    return unicodedata.normalize("NFC", value or "").strip()


def _period_key_from_filename(filename: str) -> int:
    """МесОтч202604_Смола.xlsx → 202604; если периода нет — 0."""
    match = PERIOD_IN_NAME_RE.match(Path(filename).name)
    return int(match.group(1)) if match else 0


def _safe_upload_filename(original: str) -> str:
    name = _normalize_text(Path(unquote(original)).name)
    if not name.lower().endswith(".xlsx"):
        raise ValueError("Нужен файл Excel (.xlsx)")
    if ".." in name or "/" in name or "\\" in name:
        raise ValueError("Некорректное имя файла")
    if not SAFE_FILENAME_RE.match(name):
        raise ValueError(
            "Некорректное имя файла. Используйте буквы, цифры, пробел, _ . -"
        )
    return name


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


def _list_dir_xlsx(directory: Path) -> list[dict]:
    if not directory.exists():
        return []
    files = []
    for path in sorted(directory.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True):
        stat = path.stat()
        files.append(
            {
                "name": path.name,
                "size": stat.st_size,
                "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
                "suggested_name": infer_point_name_from_filename(path.name),
            }
        )
    return files


def _list_generated_files() -> list[dict]:
    cfg = _require_config()
    return _list_dir_xlsx(Path(cfg.output_dir))


def _list_storage() -> dict:
    cfg = _require_config()
    linked = {
        Path(p.file_path).name.casefold(): p.name
        for p in cfg.points
    }
    data_files = _list_dir_xlsx(cfg.data_dir)
    for item in data_files:
        item["linked_point"] = linked.get(item["name"].casefold())
    return {
        "data": data_files,
        "generated": _list_dir_xlsx(Path(cfg.output_dir)),
    }


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


def _upsert_point(
    name: str,
    file_path: Path,
    replace: bool = False,
    *,
    persist: bool = True,
) -> PointInfo:
    cfg = _require_config()
    existing = cfg.point_by_name(name)
    if existing and not replace:
        raise ValueError(
            f"Точка «{name}» уже есть. Включите замену или выберите другое имя."
        )

    point = PointConfig(name=name, file_path=str(file_path.resolve()))
    if existing:
        cfg.points = [
            point if p.name.casefold() == name.casefold() else p for p in cfg.points
        ]
    else:
        cfg.points.append(point)

    if persist:
        save_points(cfg)
        if generator is not None:
            generator.config = cfg
    logger.info("Точка «{}» → {}", name, file_path)
    return _point_info(point)


def _persist_points() -> None:
    cfg = _require_config()
    save_points(cfg)
    if generator is not None:
        generator.config = cfg


def _move_to_trash(source: Path, kind: str) -> Path:
    """
    Перемещает файл в trash/{kind}/ с меткой времени в имени,
    чтобы не затирать ранее удалённые файлы с тем же именем.
    """
    cfg = _require_config()
    if kind not in {"data", "generated"}:
        raise ValueError("kind должен быть data или generated")

    trash_subdir = cfg.trash_dir / kind
    trash_subdir.mkdir(parents=True, exist_ok=True)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = trash_subdir / f"{stamp}__{source.name}"
    counter = 1
    while dest.exists():
        dest = trash_subdir / f"{stamp}_{counter}__{source.name}"
        counter += 1

    shutil.move(str(source), str(dest))
    logger.info("Файл перемещён в корзину: {} → {}", source, dest)
    return dest


async def _save_upload_file(upload: UploadFile, dest: Path) -> None:
    content = await upload.read()
    if not content:
        raise ValueError("Файл пустой")
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(content)


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
    _require_config()
    year, month = _default_period()
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "points": _serialize_points(),
            "storage": _list_storage(),
            "months": list(MONTH_NAMES_RU.items()),
            "default_year": year,
            "default_month": month,
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


@app.post("/api/points/upload", response_model=UploadBatchResponse)
async def upload_points(
    replace: bool = Form(default=False),
    name: Optional[str] = Form(default=None),
    files: list[UploadFile] = File(..., description="Один или несколько Excel-файлов"),
):
    """
    Добавить одну или несколько точек загрузкой Excel.

    При нескольких файлах название берётся из имени каждого файла
    (`МесОтч202603_Берёзка.xlsx` → «Берёзка»).
    Поле `name` используется только если загружен ровно один файл.

    Файлы всегда сохраняются в data/. Точка привязывается к файлу
    с самым свежим периодом МесОтчYYYYMM (или принудительно при replace).
    """
    cfg = _require_config()
    if not files:
        raise HTTPException(status_code=400, detail="Выберите хотя бы один файл")

    explicit_name = _normalize_text(name or "") or None
    if explicit_name and len(files) > 1:
        raise HTTPException(
            status_code=400,
            detail=(
                "Название точки можно указать только при загрузке одного файла. "
                "Для пакета используйте имена вида МесОтчГГГГММ_Название.xlsx"
            ),
        )

    results: list[UploadItemResult] = []
    # point_name → лучший файл в этой загрузке
    best_in_batch: dict[str, tuple[str, Path, int]] = {}
    changed = False

    for upload in files:
        original = upload.filename or "report.xlsx"
        try:
            filename = _safe_upload_filename(original)
            point_name = explicit_name or infer_point_name_from_filename(filename)
            if not point_name:
                raise ValueError(
                    "Не удалось определить название. "
                    "Используйте файл вида МесОтч202603_Название.xlsx"
                )
            point_name = _normalize_text(point_name)
            dest = cfg.data_dir / filename
            await _save_upload_file(upload, dest)
            period = _period_key_from_filename(filename)
            key = point_name.casefold()
            prev = best_in_batch.get(key)
            if prev is None or period >= prev[2]:
                best_in_batch[key] = (point_name, dest, period)

            results.append(
                UploadItemResult(
                    filename=filename,
                    status="ok",
                    name=point_name,
                    message=f"Файл сохранён: {filename}",
                    point=None,
                )
            )
        except Exception as exc:
            logger.warning("Ошибка загрузки {}: {}", original, exc)
            results.append(
                UploadItemResult(
                    filename=_normalize_text(Path(unquote(original)).name),
                    status="error",
                    message=str(exc),
                )
            )

    # Привязка точек к лучшим файлам пакета
    for point_name, dest, period in best_in_batch.values():
        existing = cfg.point_by_name(point_name)
        existing_period = (
            _period_key_from_filename(Path(existing.file_path).name)
            if existing
            else -1
        )
        should_bind = (
            existing is None
            or replace
            or period >= existing_period
        )
        if not should_bind:
            for item in results:
                if item.status == "ok" and (item.name or "").casefold() == point_name.casefold():
                    item.message = (
                        f"Файл сохранён. Точка «{point_name}» уже привязана "
                        f"к более новому файлу «{Path(existing.file_path).name}»"
                    )
            continue

        try:
            point = _upsert_point(point_name, dest, replace=True, persist=False)
            changed = True
            for item in results:
                if item.status == "ok" and (item.name or "").casefold() == point_name.casefold():
                    item.point = point
                    if existing is None:
                        item.message = f"Точка «{point_name}» добавлена → {dest.name}"
                    elif Path(existing.file_path).name == dest.name:
                        item.message = f"Точка «{point_name}» обновлена ({dest.name})"
                    else:
                        item.message = (
                            f"Точка «{point_name}» привязана к {dest.name}"
                            if item.filename == dest.name
                            else f"Файл сохранён; точка «{point_name}» → {dest.name}"
                        )
        except Exception as exc:
            logger.warning("Ошибка привязки точки {}: {}", point_name, exc)
            for item in results:
                if item.status == "ok" and (item.name or "").casefold() == point_name.casefold():
                    item.status = "error"
                    item.message = str(exc)

    if changed:
        _persist_points()

    succeeded = sum(1 for r in results if r.status == "ok")
    failed = len(results) - succeeded
    if failed == 0:
        status = "ok"
    elif succeeded == 0:
        status = "error"
    else:
        status = "partial_error"

    return UploadBatchResponse(
        status=status,
        total=len(results),
        succeeded=succeeded,
        failed=failed,
        results=results,
    )


@app.post("/api/points/import-data", response_model=UploadBatchResponse)
def import_points_from_data(replace: bool = False):
    """Зарегистрировать все Excel из data/, ещё не добавленные в список точек."""
    cfg = _require_config()
    known_files = {Path(p.file_path).name.casefold() for p in cfg.points}
    results: list[UploadItemResult] = []
    changed = False

    for item in scan_data_workbooks(cfg.data_dir):
        filename = item["filename"]
        suggested = item.get("suggested_name")
        if filename.casefold() in known_files and not replace:
            results.append(
                UploadItemResult(
                    filename=filename,
                    status="error",
                    name=suggested,
                    message="Файл уже привязан к точке (включите замену)",
                )
            )
            continue
        if not suggested:
            results.append(
                UploadItemResult(
                    filename=filename,
                    status="error",
                    message="Не удалось определить название из имени файла",
                )
            )
            continue
        try:
            point = _upsert_point(
                suggested,
                Path(item["path"]),
                replace=replace,
                persist=False,
            )
            changed = True
            results.append(
                UploadItemResult(
                    filename=filename,
                    status="ok",
                    name=point.name,
                    message=f"Точка «{point.name}» добавлена",
                    point=point,
                )
            )
        except Exception as exc:
            results.append(
                UploadItemResult(
                    filename=filename,
                    status="error",
                    name=suggested,
                    message=str(exc),
                )
            )

    if changed:
        _persist_points()

    if not results:
        return UploadBatchResponse(
            status="ok",
            total=0,
            succeeded=0,
            failed=0,
            results=[],
        )

    succeeded = sum(1 for r in results if r.status == "ok")
    failed = len(results) - succeeded
    if failed == 0:
        status = "ok"
    elif succeeded == 0:
        status = "error"
    else:
        status = "partial_error"

    return UploadBatchResponse(
        status=status,
        total=len(results),
        succeeded=succeeded,
        failed=failed,
        results=results,
    )


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
    try:
        return _upsert_point(body.name, path, replace=body.replace)
    except ValueError as exc:
        status = 409 if "уже есть" in str(exc) else 400
        raise HTTPException(status_code=status, detail=str(exc)) from exc


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

    trash_path = None
    if delete_file:
        path = Path(point.file_path)
        try:
            path.resolve().relative_to(cfg.data_dir.resolve())
            if path.exists() and path.is_file():
                trash_path = str(_move_to_trash(path, "data"))
        except ValueError:
            logger.warning("Файл точки вне data/, не перемещаем в корзину: {}", path)

    return {
        "status": "ok",
        "name": name,
        "file_deleted": trash_path is not None,
        "trash_path": trash_path,
    }


@app.get("/api/storage")
def api_storage():
    """Список исходных (data/) и сгенерированных (generated/) Excel-файлов."""
    return _list_storage()


@app.get("/api/files")
def api_files():
    """Обратная совместимость: только generated/."""
    return _list_generated_files()


@app.delete("/api/storage/{kind}/{filename}", response_model=dict)
def delete_storage_file(kind: str, filename: str, unlink_point: bool = True):
    """
    Переместить Excel из data/ или generated/ в папку trash/.

    Для data/: по умолчанию также убирает из списка точку, привязанную к этому файлу.
    """
    cfg = _require_config()
    if kind not in {"data", "generated"}:
        raise HTTPException(status_code=400, detail="kind должен быть data или generated")
    if not SAFE_FILENAME_RE.match(filename) or ".." in filename or "/" in filename or "\\" in filename:
        raise HTTPException(status_code=400, detail="Некорректное имя файла")

    folder = cfg.data_dir if kind == "data" else Path(cfg.output_dir)
    path = (folder / filename).resolve()
    try:
        path.relative_to(folder.resolve())
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="Некорректный путь к файлу") from exc

    if not path.exists() or not path.is_file():
        raise HTTPException(status_code=404, detail="Файл не найден")

    trash_path = _move_to_trash(path, kind)

    unlinked_points: list[str] = []
    if kind == "data" and unlink_point:
        remaining = []
        for point in cfg.points:
            if Path(point.file_path).name.casefold() == filename.casefold():
                unlinked_points.append(point.name)
            else:
                remaining.append(point)
        if unlinked_points:
            cfg.points = remaining
            save_points(cfg)
            if generator is not None:
                generator.config = cfg
            logger.info("Отвязаны точки после удаления файла: {}", unlinked_points)

    return {
        "status": "ok",
        "kind": kind,
        "filename": filename,
        "trash_path": str(trash_path),
        "unlinked_points": unlinked_points,
    }


@app.get("/download/{kind}/{filename}")
def download_file(kind: str, filename: str):
    cfg = _require_config()
    if kind not in {"data", "generated"}:
        raise HTTPException(status_code=400, detail="kind должен быть data или generated")
    if not SAFE_FILENAME_RE.match(filename) or ".." in filename or "/" in filename or "\\" in filename:
        raise HTTPException(status_code=400, detail="Некорректное имя файла")
    folder = cfg.data_dir if kind == "data" else Path(cfg.output_dir)
    path = folder / filename
    if not path.exists() or not path.is_file():
        raise HTTPException(status_code=404, detail="Файл не найден")
    return FileResponse(
        path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/download/{filename}")
def download_generated_compat(filename: str):
    """Старый путь скачивания из generated/."""
    return download_file("generated", filename)


@app.post("/generate", response_model=GenerateResponse)
def generate(request: GenerateRequest):
    if generator is None:
        raise HTTPException(status_code=503, detail="Сервис ещё не инициализирован")
    cfg = _require_config()

    logger.info(
        "Запрос генерации: год={} месяц={} точки={} force={}",
        request.year,
        request.month,
        request.points,
        request.force,
    )

    try:
        targets = generator._resolve_points(request.points)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    existing: list[ExistingOutputFile] = []
    out_dir = Path(cfg.output_dir)
    for point in targets:
        filename = f"МесОтч{request.year}{request.month:02d}_{point.name}.xlsx"
        path = out_dir / filename
        if path.exists():
            existing.append(
                ExistingOutputFile(
                    name=point.name,
                    filename=filename,
                    path=str(path.resolve()),
                )
            )

    if existing and not request.force:
        period = f"{MONTH_NAMES_RU[request.month]} {request.year}"
        sheet = sheet_name_for(request.year, request.month)
        conflict = GenerateConflictResponse(
            message=(
                f"За период «{period}» уже есть сгенерированные файлы "
                f"(лист {sheet}). Подтвердите пересоздание или выберите другой период/точку."
            ),
            year=request.year,
            month=request.month,
            existing=existing,
        )
        raise HTTPException(status_code=409, detail=conflict.model_dump())

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
