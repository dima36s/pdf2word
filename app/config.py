"""Настройки приложения. Всё задаётся переменными окружения с разумными значениями по умолчанию."""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path


def _int(name: str, default: int) -> int:
    try:
        return int(os.environ.get(name, default))
    except ValueError:
        return default


def _default_workers() -> int:
    cpu = os.cpu_count() or 2
    return max(1, min(4, cpu - 1 if cpu > 2 else 1))


@dataclass(frozen=True)
class Settings:
    # Где хранятся загруженные и сконвертированные файлы
    data_dir: Path = field(default_factory=lambda: Path(os.environ.get("DATA_DIR", "./data")).resolve())
    # Ограничения
    max_file_mb: int = field(default_factory=lambda: _int("MAX_FILE_MB", 100))
    max_pages: int = field(default_factory=lambda: _int("MAX_PAGES", 500))
    # Сколько конвертаций выполняется одновременно (каждая — отдельный процесс)
    workers: int = field(default_factory=lambda: _int("CONVERT_WORKERS", _default_workers()))
    # Максимальное время одной конвертации, секунд
    convert_timeout: int = field(default_factory=lambda: _int("CONVERT_TIMEOUT", 900))
    # Через сколько минут файлы удаляются с сервера
    file_ttl_min: int = field(default_factory=lambda: _int("FILE_TTL_MIN", 60))
    # OCR
    ocr_dpi: int = field(default_factory=lambda: _int("OCR_DPI", 300))
    ocr_default_lang: str = field(default_factory=lambda: os.environ.get("OCR_DEFAULT_LANG", "rus+eng"))
    tesseract_cmd: str = field(default_factory=lambda: os.environ.get("TESSERACT_CMD", "tesseract"))

    @property
    def max_file_bytes(self) -> int:
        return self.max_file_mb * 1024 * 1024

    @property
    def tasks_dir(self) -> Path:
        return self.data_dir / "tasks"


settings = Settings()
