"""Проверки плана писем, архива и отказов без реальной отправки."""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from mailer.archive import archive_sources, retarget_attachment
from mailer.excel_book import ExcelError, normalize_column, read_workbook
from mailer.outlook_mail import OutlookError, OutlookSender
from mailer.planning import (
    confirmation_text,
    delivery_problem,
    format_log,
    format_skip,
    missing_draft_message,
    missing_workbook_message,
    parse_column,
    plan_messages,
)


class ParseColumnTest(unittest.TestCase):
    def test_header_blanks_invalid_and_duplicates(self) -> None:
        parsed = parse_column(
            ["  почта  ", None, "  ", "ivan@bank.ru", "Ivan@Bank.ru", "нет адреса", "maria@bank.ru"]
        )
        self.assertEqual(parsed.addresses, ["ivan@bank.ru", "maria@bank.ru"])
        self.assertEqual(
            [(item.row, item.value, item.reason) for item in parsed.skipped],
            [
                (1, "почта", "заголовок"),
                (5, "Ivan@Bank.ru", "повтор адреса"),
                (6, "нет адреса", "неверный адрес"),
            ],
        )

    def test_english_header_and_casefold_duplicate(self) -> None:
        parsed = parse_column(["Email", "Ivan@Bank.ru", "ivan@bank.ru"])
        self.assertEqual(parsed.addresses, ["Ivan@Bank.ru"])
        self.assertEqual(parsed.skipped[0].reason, "заголовок")
        self.assertEqual(parsed.skipped[1].reason, "повтор адреса")
        self.assertEqual(parsed.skipped[1].row, 3)

    def test_first_cell_email_is_kept(self) -> None:
        parsed = parse_column(["a@b.ru", "почта"])
        self.assertEqual(parsed.addresses, ["a@b.ru"])
        self.assertEqual(parsed.skipped[0].reason, "неверный адрес")
        self.assertEqual(parsed.skipped[0].row, 2)

    def test_blank_a1_does_not_hide_a_later_header_word(self) -> None:
        parsed = parse_column([None, "ПОЧТА", "a@b.ru"])
        self.assertEqual(parsed.addresses, ["a@b.ru"])
        self.assertEqual(parsed.skipped[0].row, 2)
        self.assertEqual(parsed.skipped[0].reason, "неверный адрес")

    def test_cell_with_several_addresses_is_not_sent(self) -> None:
        parsed = parse_column(["a@b.ru, c@d.ru", "Name <a@b.ru>", "a b@c.ru"])
        self.assertEqual(parsed.addresses, [])
        self.assertEqual([item.reason for item in parsed.skipped], ["неверный адрес"] * 3)

    def test_empty_sheet(self) -> None:
        parsed = parse_column([None, "", "   "])
        self.assertEqual(parsed.addresses, [])
        self.assertEqual(parsed.skipped, [])

    def test_skip_log_line(self) -> None:
        parsed = parse_column(["почта"])
        self.assertEqual(
            format_skip("Банк 1", parsed.skipped[0]),
            "Банк 1 | почта | пропущено: заголовок (строка 1)",
        )


class PlanMessagesTest(unittest.TestCase):
    def test_one_message_per_address_and_sheet(self) -> None:
        messages = plan_messages(
            [
                {"name": "Банк 1", "addresses": ["a@bank.ru", "b@bank.ru"], "selected": True},
                {"name": "Банк 2", "addresses": ["a@bank.ru"], "selected": True},
                {"name": "Банк 3", "addresses": ["c@bank.ru"], "selected": False},
                {"name": "Пусто", "addresses": [], "selected": True},
            ],
            subject="Тема",
            body="Текст письма",
            attachment="C:/docs/file.pdf",
        )
        self.assertEqual(
            [(item["sheet"], item["to"]) for item in messages],
            [("Банк 1", "a@bank.ru"), ("Банк 1", "b@bank.ru"), ("Банк 2", "a@bank.ru")],
        )
        for item in messages:
            self.assertEqual(set(item), {"sheet", "to", "subject", "body", "cc", "attachment"})
            self.assertEqual(item["subject"], "Тема")
            self.assertEqual(item["body"], "Текст письма")
            self.assertEqual(item["cc"], "")
            self.assertEqual(item["attachment"], "C:/docs/file.pdf")
            self.assertNotIn(",", item["to"])
            self.assertNotIn(";", item["to"])

    def test_to_comes_from_sheet_and_letter_from_draft(self) -> None:
        messages = plan_messages(
            [
                {"name": "Банк 1", "addresses": ["ivan@bank.ru", "maria@bank.ru"], "selected": True},
                {"name": "Банк 2", "addresses": ["petr@bank.ru"], "selected": True},
            ],
            subject="Тема из черновика",
            body="Текст из черновика",
            cc="boss@example.com; lawyer@example.com",
        )
        self.assertEqual(
            [item["to"] for item in messages],
            ["ivan@bank.ru", "maria@bank.ru", "petr@bank.ru"],
        )
        for item in messages:
            self.assertEqual(item["subject"], "Тема из черновика")
            self.assertEqual(item["body"], "Текст из черновика")
            self.assertEqual(item["cc"], "boss@example.com; lawyer@example.com")
            self.assertNotIn(";", item["to"])
            self.assertNotIn(",", item["to"])

    def test_blank_attachment_is_none(self) -> None:
        messages = plan_messages(
            [{"name": "Банк 1", "addresses": ["a@b.ru"], "selected": True}],
            "Тема",
            "Текст",
            "   ",
        )
        self.assertIsNone(messages[0]["attachment"])

    def test_confirmation_states_the_count(self) -> None:
        text = confirmation_text(12, "  Тема\nписьма  ")
        self.assertIn("Число отдельных писем: 12.", text)
        self.assertIn("Тема: Тема письма", text)

    def test_result_log_line(self) -> None:
        self.assertEqual(
            format_log("Банк 1", "a@b.ru", "отправлено"),
            "Банк 1 | a@b.ru | отправлено",
        )


class DeliveryProblemTest(unittest.TestCase):
    def test_missing_workbook_names_the_path(self) -> None:
        path = r"C:\путь\к\файлу\banks.xlsx"
        text = missing_workbook_message(path)
        self.assertIsNotNone(text)
        self.assertIn("Файл книги не найден.", text)
        self.assertIn("Искали здесь:", text)
        self.assertIn(path, text)
        problem = delivery_problem(
            workbook=path,
            draft=r"C:\путь\к\файлу\draft.msg",
            subject="Тема",
            body="Текст",
            attachment="",
            sheets_loaded=False,
            draft_loaded=False,
        )
        self.assertIn(path, problem)

    def test_missing_draft_names_the_path(self) -> None:
        path = r"C:\путь\к\файлу\draft.msg"
        text = missing_draft_message(path)
        self.assertIsNotNone(text)
        self.assertIn("Файл черновика не найден.", text)
        self.assertIn("Искали здесь:", text)
        self.assertIn(path, text)
        with tempfile.TemporaryDirectory() as tmp:
            book = Path(tmp) / "banks.xlsx"
            book.write_bytes(b"x")
            problem = delivery_problem(
                workbook=str(book),
                draft=path,
                subject="Тема",
                body="Текст",
                attachment="",
                sheets_loaded=True,
                draft_loaded=False,
            )
        self.assertIn(path, problem)

    def test_existing_workbook_can_be_sent(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            book = Path(tmp) / "banks.xlsx"
            book.write_bytes(b"x")
            draft = Path(tmp) / "draft.msg"
            draft.write_bytes(b"msg")
            self.assertIsNone(missing_workbook_message(str(book)))
            self.assertIsNone(missing_draft_message(str(draft)))
            problem = delivery_problem(
                workbook=str(book),
                draft=str(draft),
                subject="Тема",
                body="Текст",
                attachment="",
                sheets_loaded=True,
                draft_loaded=True,
            )
        self.assertIsNone(problem)

    def test_requires_draft_text_and_existing_attachment(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            book = Path(tmp) / "banks.xlsx"
            book.write_bytes(b"x")
            draft = Path(tmp) / "draft.msg"
            draft.write_bytes(b"msg")
            empty = delivery_problem(
                workbook=str(book),
                draft=str(draft),
                subject="  ",
                body="Текст",
                attachment="",
                sheets_loaded=True,
                draft_loaded=True,
            )
            missing = delivery_problem(
                workbook=str(book),
                draft=str(draft),
                subject="Тема",
                body="Текст",
                attachment=str(Path(tmp) / "no.pdf"),
                sheets_loaded=True,
                draft_loaded=True,
            )
            closed = delivery_problem(
                workbook=str(book),
                draft=str(draft),
                subject="Тема",
                body="Текст",
                attachment="",
                sheets_loaded=False,
                draft_loaded=True,
            )
        self.assertIn("тему и текст", empty)
        self.assertIn("вложения", missing)
        self.assertIn("не открыта", closed)


class ArchiveTest(unittest.TestCase):
    def test_copies_workbook_and_attachment_into_timestamped_folder(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            app_dir = root / "app"
            app_dir.mkdir()
            book = root / "banks.xlsx"
            book.write_bytes(b"book")
            attachment = root / "note.pdf"
            attachment.write_bytes(b"pdf")
            moment = datetime(2026, 11, 2, 17, 5, 6)

            first = archive_sources(book, attachment, app_dir, now=moment)
            second = archive_sources(book, None, app_dir, now=moment)

            self.assertEqual(first.folder, app_dir / "archive" / "20261102_170506")
            self.assertEqual(first.workbook.read_bytes(), b"book")
            self.assertEqual(first.attachment.read_bytes(), b"pdf")
            self.assertEqual(book.read_bytes(), b"book")
            self.assertEqual(second.folder.name, "20261102_170506_1")
            self.assertIsNone(second.attachment)
            self.assertTrue(second.workbook.is_file())

    def test_same_file_name_gets_a_prefix(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            book = root / "file.xlsx"
            nested = root / "nested"
            nested.mkdir()
            attachment = nested / "file.xlsx"
            book.write_bytes(b"book")
            attachment.write_bytes(b"attach")
            result = archive_sources(book, attachment, root, now=datetime(2026, 1, 1, 0, 0, 0))
            self.assertEqual(result.attachment.name, "вложение_file.xlsx")
            self.assertEqual(result.attachment.read_bytes(), b"attach")
            self.assertEqual(result.workbook.read_bytes(), b"book")

    def test_retarget_does_not_mutate_the_plan(self) -> None:
        messages = plan_messages(
            [{"name": "Банк 1", "addresses": ["a@b.ru"], "selected": True}],
            "Тема",
            "Текст",
            "note.pdf",
        )
        archived = Path("/tmp/archive/note.pdf")
        updated = retarget_attachment(messages, archived)
        self.assertEqual(messages[0]["attachment"], "note.pdf")
        self.assertEqual(updated[0]["attachment"], str(archived))
        cleared = retarget_attachment(messages, None)
        self.assertIsNone(cleared[0]["attachment"])


class ExcelValuesTest(unittest.TestCase):
    def test_normalize_column_range(self) -> None:
        raw = (("a@b.ru",), ("почта",), (None,))
        self.assertEqual(normalize_column(raw, 3), ["a@b.ru", "почта", None])
        self.assertEqual(normalize_column("only@one.ru", 1), ["only@one.ru"])

    def test_rejects_non_xlsx_before_excel(self) -> None:
        with self.assertRaises(ExcelError) as caught:
            read_workbook("book.xls")
        self.assertIn(".xlsx", str(caught.exception))

    def test_missing_file_before_excel(self) -> None:
        with self.assertRaises(ExcelError) as caught:
            read_workbook("no-such-banks-file.xlsx")
        self.assertIn("не найден", str(caught.exception).lower())

    def test_clear_error_when_excel_stack_is_unavailable(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "book.xlsx"
            path.write_bytes(b"not a real workbook")
            with self.assertRaises(ExcelError) as caught:
                read_workbook(path)
        text = str(caught.exception)
        self.assertTrue("pywin32" in text or "Excel" in text)


class FakeMail:
    def __init__(self) -> None:
        self.Subject = ""
        self.Body = ""
        self.To = ""
        self.CC = ""
        self.BCC = ""
        self.added: list[str] = []
        self.sent = False
        self.Attachments = self

    def Add(self, path: str) -> None:
        self.added.append(path)

    def Send(self) -> None:
        self.sent = True


class _FakeDraftItem:
    def __init__(self) -> None:
        self.Subject = "Тема из черновика"
        self.Body = "Текст из черновика"
        self.CC = "boss@example.com; lawyer@example.com"
        self.closed: int | None = None
        self.sent = False

    def Close(self, mode: int) -> None:
        self.closed = mode

    def Send(self) -> None:
        self.sent = True


class _FakeDraftOutlook:
    def __init__(self, item: _FakeDraftItem) -> None:
        self.item = item
        self.opened: str | None = None

    def GetNamespace(self, name: str) -> "_FakeDraftOutlook":
        if name != "MAPI":
            raise AssertionError(name)
        return self

    def OpenSharedItem(self, path: str) -> _FakeDraftItem:
        self.opened = path
        return self.item


class FakeOutlook:
    def __init__(self) -> None:
        self.created: list[FakeMail] = []

    def CreateItem(self, kind: int) -> FakeMail:
        if kind != 0:
            raise AssertionError(kind)
        mail = FakeMail()
        self.created.append(mail)
        return mail


class OutlookGuardTest(unittest.TestCase):
    def test_send_before_connect(self) -> None:
        sender = OutlookSender()
        with self.assertRaises(OutlookError):
            sender.send({"to": "a@b.ru", "subject": "s", "body": "b"})

    def test_refuses_a_combined_recipient_list(self) -> None:
        sender = OutlookSender()
        sender._app = FakeOutlook()
        with self.assertRaises(OutlookError):
            sender.send({"to": "a@b.ru; c@d.ru", "subject": "s", "body": "b"})
        self.assertEqual(sender._app.created, [])

    def test_one_mail_item_per_address_without_outlook(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            attachment = Path(tmp) / "note.txt"
            attachment.write_text("hi", encoding="utf-8")
            sender = OutlookSender()
            fake = FakeOutlook()
            sender._app = fake
            messages = plan_messages(
                [{"name": "Банк 1", "addresses": ["a@b.ru", "c@d.ru"], "selected": True}],
                subject="Тема из черновика",
                body="Текст из черновика",
                attachment=str(attachment),
                cc="boss@example.com; lawyer@example.com",
            )
            for message in messages:
                sender.send(message)

        self.assertEqual(len(fake.created), 2)
        self.assertEqual([mail.To for mail in fake.created], ["a@b.ru", "c@d.ru"])
        for mail in fake.created:
            self.assertEqual(mail.Subject, "Тема из черновика")
            self.assertEqual(mail.Body, "Текст из черновика")
            self.assertEqual(mail.CC, "boss@example.com; lawyer@example.com")
            self.assertEqual(mail.BCC, "")
            self.assertNotIn(";", mail.To)
            self.assertTrue(mail.sent)
            self.assertEqual(mail.added, [str(attachment.resolve())])

    def test_read_draft_copies_cc_and_does_not_send(self) -> None:
        item = _FakeDraftItem()
        outlook = _FakeDraftOutlook(item)
        sender = OutlookSender()
        sender._app = outlook
        draft = sender.read_draft(r"C:\путь\к\файлу\draft.msg")
        self.assertEqual(draft.subject, "Тема из черновика")
        self.assertEqual(draft.body, "Текст из черновика")
        self.assertEqual(draft.cc, "boss@example.com; lawyer@example.com")
        self.assertEqual(item.closed, 1)
        self.assertFalse(item.sent)
        self.assertTrue(outlook.opened)

    def test_connect_reports_a_clear_error_and_does_not_send(self) -> None:
        sender = OutlookSender()
        with self.assertRaises(OutlookError) as caught:
            sender.connect()
        text = str(caught.exception)
        self.assertTrue("Outlook" in text or "pywin32" in text)
        self.assertIsNone(sender._app)


class UiSmokeTest(unittest.TestCase):
    def _tk(self):
        try:
            import tkinter as tk
        except ModuleNotFoundError:
            self.skipTest("нет tkinter")
        if not os.environ.get("DISPLAY"):
            self.skipTest("нет дисплея")
        return tk

    def test_window_has_russian_controls(self) -> None:
        tk = self._tk()

        from mailer.ui import LetterApp

        root = tk.Tk()
        try:
            root.withdraw()
            app = LetterApp(
                root,
                r"C:\путь\к\файлу\banks.xlsx",
                r"C:\путь\к\файлу\draft.msg",
            )
            root.update_idletasks()
            labels: list[str] = []

            def walk(widget) -> None:
                try:
                    text = widget.cget("text")
                except tk.TclError:
                    text = ""
                if text:
                    labels.append(str(text))
                for child in widget.winfo_children():
                    walk(child)

            walk(root)
            joined = "\n".join(labels)
            for expected in (
                "Обновить",
                "Отправить",
                "Обзор…",
                "Отметить все",
                "Снять все",
                "Тема:",
                "Текст письма:",
                "Копия (Cc):",
                r"C:\путь\к\файлу\banks.xlsx",
                r"C:\путь\к\файлу\draft.msg",
            ):
                self.assertIn(expected, joined)
            self.assertNotIn("Загрузить листы", joined)
            self.assertEqual(root.title(), "Banks — генератор писем")
            app.sheet_list.set_sheets([("Банк 1", 2), ("Банк 2", 0)])
            root.update_idletasks()
            self.assertEqual(app.sheet_list.selected_names(), ["Банк 1"])
            app.sheet_list.select_none()
            self.assertEqual(app.sheet_list.selected_names(), [])
            app.sheet_list.select_all()
            self.assertEqual(app.sheet_list.selected_names(), ["Банк 1"])
        finally:
            root.destroy()

    def test_draft_text_stays_editable(self) -> None:
        tk = self._tk()
        from mailer.ui import LetterApp

        root = tk.Tk()
        try:
            root.withdraw()
            app = LetterApp(
                root,
                r"C:\путь\к\файлу\banks.xlsx",
                r"C:\путь\к\файлу\draft.msg",
            )
            app.draft_cc = "boss@example.com"
            app.draft_loaded = True
            app._show_draft("Тема из черновика", "Текст из черновика", "boss@example.com")
            self.assertEqual(app.body.cget("state"), "normal")
            self.assertEqual(app.cc_box.cget("state"), "disabled")
            app.body.insert("end", "\nПравлено в окне.")
            app.subject_var.set("Новая тема")
            messages = plan_messages(
                [{"name": "Банк 1", "addresses": ["ivan@bank.ru"], "selected": True}],
                app.current_subject(),
                app.current_body(),
                cc=app.draft_cc,
            )
            self.assertEqual(messages[0]["to"], "ivan@bank.ru")
            self.assertEqual(messages[0]["subject"], "Новая тема")
            self.assertEqual(messages[0]["body"], "Текст из черновика\nПравлено в окне.")
            self.assertEqual(messages[0]["cc"], "boss@example.com")
        finally:
            root.destroy()

    def test_missing_file_dialog_names_the_path(self) -> None:
        tk = self._tk()
        from unittest.mock import patch

        from mailer.ui import LetterApp

        path = r"C:\путь\к\файлу\banks.xlsx"
        root = tk.Tk()
        try:
            root.withdraw()
            app = LetterApp(root, path, r"C:\путь\к\файлу\draft.msg")
            with patch("mailer.ui.messagebox.showerror") as show:
                app.open_workbook()
            show.assert_called_once()
            self.assertEqual(show.call_args[0][0], "Книга Excel")
            self.assertIn("Искали здесь:", show.call_args[0][1])
            self.assertIn(path, show.call_args[0][1])
            self.assertIsNone(app.loaded_path)
            self.assertEqual(app.loaded, [])
        finally:
            root.destroy()

    def test_missing_draft_dialog_names_the_path(self) -> None:
        tk = self._tk()
        from unittest.mock import patch

        from mailer.ui import LetterApp

        path = r"C:\путь\к\файлу\draft.msg"
        root = tk.Tk()
        try:
            root.withdraw()
            app = LetterApp(root, r"C:\путь\к\файлу\banks.xlsx", path)
            with patch("mailer.ui.messagebox.showerror") as show:
                app.open_draft()
            show.assert_called_once()
            self.assertEqual(show.call_args[0][0], "Черновик")
            self.assertIn("Искали здесь:", show.call_args[0][1])
            self.assertIn(path, show.call_args[0][1])
            self.assertFalse(app.draft_loaded)
        finally:
            root.destroy()


if __name__ == "__main__":
    unittest.main()
