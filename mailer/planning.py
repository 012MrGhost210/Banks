"""Разбор колонки адресов и план писем без Outlook и Excel."""

from __future__ import annotations

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
    cc: str = "",
) -> list[dict]:
    """Собрать письма, которые уйдут по одному на адрес.

    Получатель «Кому» берётся из листа. Тема, текст и копия (Cc) — из черновика
    и одинаковы для всех. Лист без отметки или без адресов не попадает в план.
    Ключи: sheet, to, subject, body, cc, attachment. Outlook для этого не нужен.
    """
    if isinstance(attachment, str):
        attachment_path = attachment.strip() or None
    elif attachment is None:
        attachment_path = None
    else:
        attachment_path = str(attachment)
    cc_value = str(cc or "").strip()

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
                    "cc": cc_value,
                    "attachment": attachment_path,
                }
            )
    return planned


def cc_people(cc: str) -> list[str]:
    """Разобрать строку копии Outlook на отдельных людей для показа в окне."""
    people: list[str] = []
    for part in str(cc or "").replace("\n", ";").split(";"):
        person = part.strip().strip(",")
        if person:
            people.append(person)
    return people


def confirmation_text(count: int, subject: str, cc: str = "") -> str:
    shown = one_line(subject)
    if len(shown) > 120:
        shown = shown[:117] + "..."
    people = cc_people(cc)
    copy_line = ""
    if people:
        copy_line = "В копии каждого письма: " + "; ".join(people) + "\n\n"
    return (
        f"Число отдельных писем: {count}.\n\n"
        "Каждое письмо уйдёт только одному адресу в поле «Кому». "
        "Остальные адреса из таблицы в это письмо не попадут.\n\n"
        f"{copy_line}"
        f"Тема: {shown}\n\n"
        "Продолжить?"
    )


def missing_workbook_message(path: str) -> str | None:
    """Если книгу по константе открыть нельзя — текст ошибки, иначе None.

    В тексте есть тот путь, который проверяли.
    """
    text = str(path).strip().strip('"')
    if not text:
        return (
            "Путь к книге не задан.\n"
            "Откройте app.py и в константе WORKBOOK_PATH пропишите путь к файлу .xlsx."
        )
    file_path = Path(text)
    if not file_path.is_file():
        return f"Файл книги не найден.\nИскали здесь:\n{text}"
    if file_path.suffix.lower() != ".xlsx":
        return f"Нужен файл Excel с расширением .xlsx.\nИскали здесь:\n{text}"
    return None


def missing_draft_message(path: str) -> str | None:
    """Если черновик по константе открыть нельзя — текст ошибки, иначе None."""
    text = str(path).strip().strip('"')
    if not text:
        return (
            "Путь к черновику не задан.\n"
            "Откройте app.py и в константе DRAFT_PATH пропишите путь к файлу .msg."
        )
    file_path = Path(text)
    if not file_path.is_file():
        return f"Файл черновика не найден.\nИскали здесь:\n{text}"
    if file_path.suffix.lower() != ".msg":
        return (
            "Черновик должен быть файлом Outlook .msg.\n"
            f"Искали здесь:\n{text}"
        )
    return None


def delivery_problem(
    *,
    workbook: str,
    draft: str,
    subject: str,
    body: str,
    attachment: str,
    sheets_loaded: bool,
    draft_loaded: bool,
) -> str | None:
    """Проверка полей перед планом писем. None — можно продолжать."""
    missing = missing_workbook_message(workbook)
    if missing:
        return missing
    missing_draft = missing_draft_message(draft)
    if missing_draft:
        return missing_draft
    if not sheets_loaded:
        return "Книга не открыта. Проверьте путь WORKBOOK_PATH в файле app.py."
    if not draft_loaded:
        return "Черновик не открыт. Проверьте путь DRAFT_PATH в файле app.py."
    if not subject.strip() or not body.strip():
        return "Заполните тему и текст письма. Черновик только подставляет их в окно."
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
