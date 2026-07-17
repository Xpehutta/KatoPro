"""Pydantic-модели REST API."""

from __future__ import annotations

from typing import Optional

from pydantic import BaseModel, Field, field_validator


class GenerateRequest(BaseModel):
    year: int = Field(..., ge=2000, le=2100, description="Год отчёта")
    month: int = Field(..., ge=1, le=12, description="Месяц отчёта (1–12)")
    points: Optional[list[str]] = Field(
        default=None,
        description="Имена торговых точек. Если не указано — обрабатываются все из конфигурации.",
    )
    force: bool = Field(
        default=False,
        description="Перезаписать уже существующие файлы за этот период без дополнительного отказа API",
    )
    skip_previous_check: bool = Field(
        default=False,
        description="Не проверять наличие файла предыдущего периода в data/",
    )
    point_overrides: Optional[list["PointPeriodOverride"]] = Field(
        default=None,
        description="Индивидуальный год/месяц генерации для отдельных точек",
    )

    @field_validator("points")
    @classmethod
    def nonempty_names(cls, value: Optional[list[str]]) -> Optional[list[str]]:
        if value is not None and len(value) == 0:
            raise ValueError(
                "Список точек не должен быть пустым; не передавайте поле, чтобы обработать все точки"
            )
        return value


class PointPeriodOverride(BaseModel):
    name: str = Field(..., min_length=1)
    year: int = Field(..., ge=2000, le=2100)
    month: int = Field(..., ge=1, le=12)


class ExistingOutputFile(BaseModel):
    name: str
    filename: str
    path: str


class GenerateConflictResponse(BaseModel):
    code: str = "already_exists"
    message: str
    year: int
    month: int
    existing: list[ExistingOutputFile]


class MissingPreviousPeriod(BaseModel):
    name: str
    requested_year: int
    requested_month: int
    previous_year: int
    previous_month: int
    previous_filename: str
    latest_data_file: Optional[str] = None
    latest_year: Optional[int] = None
    latest_month: Optional[int] = None
    recommended_year: Optional[int] = None
    recommended_month: Optional[int] = None
    message: str


class GenerateMissingPreviousResponse(BaseModel):
    code: str = "missing_previous"
    message: str
    year: int
    month: int
    missing: list[MissingPreviousPeriod]


class PointResult(BaseModel):
    name: str
    status: str
    output_file: Optional[str] = None
    sheet_name: Optional[str] = None
    message: Optional[str] = None


class GenerateResponse(BaseModel):
    status: str
    year: int
    month: int
    results: list[PointResult]


class PointInfo(BaseModel):
    name: str
    file_path: str
    file_exists: bool = True
    filename: Optional[str] = None


class UploadItemResult(BaseModel):
    filename: str
    status: str  # ok | error
    name: Optional[str] = None
    message: Optional[str] = None
    point: Optional[PointInfo] = None


class UploadBatchResponse(BaseModel):
    status: str  # ok | partial_error | error
    total: int
    succeeded: int
    failed: int
    results: list[UploadItemResult]


class AddPointManualRequest(BaseModel):
    name: str = Field(..., min_length=1, max_length=80, description="Название торговой точки")
    filename: str = Field(
        ...,
        min_length=1,
        description="Имя уже загруженного Excel-файла из папки data/",
    )
    replace: bool = Field(
        default=False,
        description="Заменить точку, если имя уже существует",
    )

    @field_validator("name")
    @classmethod
    def strip_name(cls, value: str) -> str:
        value = value.strip()
        if not value:
            raise ValueError("Название точки не может быть пустым")
        return value

    @field_validator("filename")
    @classmethod
    def validate_filename(cls, value: str) -> str:
        value = value.strip()
        if not value.lower().endswith(".xlsx"):
            raise ValueError("Нужен файл Excel (.xlsx)")
        if "/" in value or "\\" in value or ".." in value:
            raise ValueError("Некорректное имя файла")
        return value
