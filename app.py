"""Запуск окна генератора писем."""

# сюда пропиши свой путь
WORKBOOK_PATH = r"C:\путь\к\файлу\banks.xlsx"

from mailer.ui import run


def main() -> None:
    run(WORKBOOK_PATH)


if __name__ == "__main__":
    main()
