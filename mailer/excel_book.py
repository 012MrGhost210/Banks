"""Чтение колонки A через Excel.Application (win32com)."""

from __future__ import annotations

import gc
from pathlib import Path

# Excel constant xlUp. Конец колонки ищется от нижней строки вверх.
XL_UP = -4162


class ExcelError(Exception):
    """Книгу прочитать нельзя: нет файла, pywin32 или Excel."""


def normalize_column(raw: object, last_row: int) -> list:
    """Привести ответ Excel по диапазону колонки к списку значений ячеек.

    Для одной ячейки Excel возвращает само значение. Для нескольких строк —
    последовательность кортежей вида ((значение,), ...). Иногда это плоский
    список значений. Пустой диапазон остаётся списком пустых ячеек.
    """
    if last_row <= 1:
        return [raw]
    if raw is None:
        return [None] * last_row
    values = []
    for item in raw:
        if isinstance(item, (list, tuple)):
            values.append(item[0] if item else None)
        else:
            values.append(item)
    return values


def read_workbook(path: Path | str) -> list[tuple[str, list]]:
    """Открыть .xlsx отдельным процессом Excel и вернуть (имя листа, колонка A).

    Используется DispatchEx, а не Dispatch: иначе Quit закрыл бы Excel,
    который пользователь уже открыл. Процесс, запущенный здесь, в конце закрывается.
    """
    file_path = Path(path)
    if file_path.suffix.lower() != ".xlsx":
        raise ExcelError("Нужен файл Excel с расширением .xlsx.")
    if not file_path.is_file():
        raise ExcelError(f"Файл не найден: {file_path}")

    try:
        import win32com.client
    except ImportError as exc:
        raise ExcelError(
            "Не установлен модуль pywin32.\n"
            "В папке приложения выполните:\n"
            "pip install -r requirements.txt"
        ) from exc

    excel = None
    workbook = None
    try:
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(str(file_path.resolve()), 0, True)
        except ExcelError:
            raise
        except Exception as exc:
            raise ExcelError(
                "Не удалось открыть книгу в Excel. Установите Microsoft Excel. "
                "Файл должен открываться без пароля.\n"
                f"{exc}"
            ) from exc

        sheets: list[tuple[str, list]] = []
        try:
            for sheet in workbook.Worksheets:
                sheets.append((str(sheet.Name), _column_a(sheet)))
        except ExcelError:
            raise
        except Exception as exc:
            raise ExcelError(
                "Не удалось прочитать колонку A. Проверьте, что на листах обычные ячейки.\n"
                f"{exc}"
            ) from exc
        return sheets
    finally:
        _close_excel(workbook, excel)


def _column_a(sheet) -> list:
    last_row = sheet.Cells(sheet.Rows.Count, 1).End(XL_UP).Row
    try:
        last_row = int(last_row)
    except (TypeError, ValueError) as exc:
        raise ExcelError("Не удалось определить последнюю строку колонки A.") from exc
    if last_row < 1:
        return []
    if last_row == 1:
        return [sheet.Cells(1, 1).Value]
    raw = sheet.Range(sheet.Cells(1, 1), sheet.Cells(last_row, 1)).Value
    return normalize_column(raw, last_row)


def _close_excel(workbook, excel) -> None:
    if workbook is not None:
        try:
            workbook.Close(SaveChanges=False)
        except Exception:
            pass
    if excel is not None:
        try:
            excel.Quit()
        except Exception:
            pass
    workbook = None
    excel = None
    gc.collect()
