"""Запуск окна генератора писем."""

# сюда пропиши свой путь
WORKBOOK_PATH = r"C:\путь\к\файлу\banks.xlsx"

# сюда пропиши путь к черновику
DRAFT_PATH = r"C:\путь\к\файлу\draft.msg"

from mailer.ui import run


def main() -> None:
    run(WORKBOOK_PATH, DRAFT_PATH)


if __name__ == "__main__":
    main()
