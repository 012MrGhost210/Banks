"""Копия книги и вложения в архив рядом с приложением. Только shutil."""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


@dataclass(frozen=True)
class ArchiveResult:
    folder: Path
    workbook: Path
    attachment: Path | None


def archive_sources(
    workbook: Path | str,
    attachment: Path | str | None,
    app_dir: Path | str,
    now: datetime | None = None,
) -> ArchiveResult:
    """Скопировать выбранные файлы в archive/ГГГГММДД_ЧЧММСС до отправки."""
    workbook_path = Path(workbook)
    attachment_path = Path(attachment) if attachment else None
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Файл не найден: {workbook_path}")
    if attachment_path is not None and not attachment_path.is_file():
        raise FileNotFoundError(f"Файл не найден: {attachment_path}")

    moment = now or datetime.now()
    stamp = moment.strftime("%Y%m%d_%H%M%S")
    folder = _unique_folder(Path(app_dir) / "archive", stamp)
    folder.mkdir(parents=True)
    try:
        copied_book = folder / workbook_path.name
        shutil.copy2(workbook_path, copied_book)
        copied_attachment = None
        if attachment_path is not None:
            name = attachment_path.name
            if name.casefold() == workbook_path.name.casefold():
                name = f"вложение_{name}"
            copied_attachment = folder / name
            shutil.copy2(attachment_path, copied_attachment)
    except Exception:
        shutil.rmtree(folder, ignore_errors=True)
        raise
    return ArchiveResult(
        folder=folder,
        workbook=copied_book,
        attachment=copied_attachment,
    )


def retarget_attachment(messages: list[dict], attachment: Path | None) -> list[dict]:
    """Подставить в план путь архивной копии вложения, не меняя исходный план."""
    path = str(attachment) if attachment else None
    return [{**message, "attachment": path} for message in messages]


def _unique_folder(root: Path, stamp: str) -> Path:
    folder = root / stamp
    index = 1
    while folder.exists():
        folder = root / f"{stamp}_{index}"
        index += 1
    return folder
