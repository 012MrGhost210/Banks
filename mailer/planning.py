"""Разбор колонки адресов и план писем без Outlook и Excel."""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from pathlib import Path

# Простая форма адреса: один @, точка в домене, без списка получателей в одной ячейке.
_EMAIL = re.compile(r"^[^@\s<>\",;]+@[^@\s<>\",;]+\.[^@\s<>\",;]{2,}$")


@dataclass(frozen=True)
class SkippedCell:
    row: int
    value: str
    reason: str


@dataclass(frozen=True)
class ColumnParse:
    addresses: list[str]
    skipped: list[SkippedCell]


def is_email(value: str) -> bool:
    return _EMAIL.fullmatch(value) is not None


def cell_text(value: object) -> str | None:
    """Текст ячейки. Пустые и пробельные значения — это пустая ячейка."""
    if value is None:
        return None
    if isinstance(value, str):
        text = value.strip()
    else:
        text = str(value).strip()
    return text or None


def parse_column(cells: list[object]) -> ColumnParse:
    """Разобрать колонку A одного листа.

    Пустые ячейки пропускаются. Если в A1 не адрес (заголовок вроде «email»
    или «почта»), строка пропускается. Непохожие на адрес ячейки и повторы
    внутри листа тоже пропускаются. Повторы сравниваются без учёта регистра.
    """
    addresses: list[str] = []
    seen: set[str] = set()
    skipped: list[SkippedCell] = []

    for index, raw in enumerate(cells):
        row = index + 1
        text = cell_text(raw)
        if text is None:
            continue
        if row == 1 and not is_email(text):
            # «почта» и «email» — заголовок. Ячейка с @, но не адрес, — ошибка в журнале.
            reason = "неверный адрес" if "@" in text else "заголовок"
            skipped.append(SkippedCell(row, text, reason))
            continue
        if not is_email(text):
            skipped.append(SkippedCell(row, text, "неверный адрес"))
            continue
        key = text.casefold()
        if key in seen:
            skipped.append(SkippedCell(row, text, "повтор адреса"))
            continue
        seen.add(key)
        addresses.append(text)

    return ColumnParse(addresses=addresses, skipped=skipped)


def plan_messages(
    sheets: list[dict],
    subject: str,
    body: str,
    attachment: str | None = None,
) -> list[dict]:
    """Собрать письма, которые уйдут по одному на адрес.

    Лист без отметки или без адресов не попадает в план. В каждом элементе
    ровно один получатель. Ключи: sheet, to, subject, body, attachment.
    Outlook для этого не нужен.
    """
    if isinstance(attachment, str):
        attachment_path = attachment.strip() or None
    elif attachment is None:
        attachment_path = None
    else:
        attachment_path = str(attachment)

    planned: list[dict] = []
    for sheet in sheets:
        if not sheet.get("selected"):
            continue
        name = str(sheet.get("name", ""))
        for address in sheet.get("addresses") or []:
            recipient = str(address).strip()
            if not recipient:
                continue
            planned.append(
                {
                    "sheet": name,
                    "to": recipient,
                    "subject": subject,
                    "body": body,
                    "attachment": attachment_path,
                }
            )
    return planned


def confirmation_text(count: int, subject: str) -> str:
    shown = one_line(subject)
    if len(shown) > 120:
        shown = shown[:117] + "..."
    return (
        f"Число отдельных писем: {count}.\n\n"
        "Каждое письмо уйдёт только одному адресу. "
        "Остальные получатели в это письмо не попадут.\n\n"
        f"Тема: {shown}\n\n"
        "Продолжить?"
    )


def delivery_problem(
    *,
    workbook: str,
    subject: str,
    body: str,
    attachment: str,
    loaded_path: str | None,
) -> str | None:
    """Проверка полей перед планом писем. None — можно продолжать."""
    path = workbook.strip()
    if not path:
        return "Укажите путь к файлу .xlsx."
    file_path = Path(path)
    if file_path.suffix.lower() != ".xlsx":
        return "Нужен файл Excel с расширением .xlsx."
    if not file_path.is_file():
        return f"Файл не найден: {file_path}"
    if loaded_path is None:
        return "Сначала загрузите листы из книги Excel."
    try:
        same = os.path.normcase(str(file_path.resolve())) == os.path.normcase(
            str(Path(loaded_path).resolve())
        )
    except OSError:
        same = False
    if not same:
        return "Путь к книге изменился. Загрузите листы ещё раз."
    if not subject.strip() or not body.strip():
        return "Заполните тему и текст письма. Они одинаковые для всех получателей."
    attachment_text = attachment.strip()
    if attachment_text and not Path(attachment_text).is_file():
        return "Файл вложения не найден."
    return None


def one_line(value: object, limit: int = 300) -> str:
    text = " ".join(str(value).split())
    if len(text) > limit:
        return text[: limit - 3] + "..."
    return text


def format_log(sheet: object, address: object, result: object) -> str:
    return f"{one_line(sheet)} | {one_line(address)} | {one_line(result)}"


def format_skip(sheet: str, skipped: SkippedCell) -> str:
    return format_log(
        sheet,
        skipped.value,
        f"пропущено: {skipped.reason} (строка {skipped.row})",
    )
