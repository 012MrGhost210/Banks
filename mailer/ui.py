"""Окно генератора писем."""

from __future__ import annotations

import tkinter as tk
import tkinter.font as tkfont
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from mailer.archive import archive_sources, retarget_attachment
from mailer.excel_book import ExcelError, read_workbook
from mailer.outlook_mail import OutlookError, OutlookSender
from mailer.planning import (
    ColumnParse,
    cc_people,
    confirmation_text,
    delivery_problem,
    format_log,
    format_skip,
    missing_draft_message,
    missing_workbook_message,
    one_line,
    parse_column,
    plan_messages,
)


def app_directory() -> Path:
    """Папка с app.py. Рядом с ней создаётся archive."""
    return Path(__file__).resolve().parent.parent


class SheetList(ttk.Frame):
    """Список листов: флажок, имя и число адресов."""

    def __init__(self, master) -> None:
        super().__init__(master)
        self.rows: list[tuple[str, tk.BooleanVar, int]] = []
        self.canvas = tk.Canvas(self, height=180, highlightthickness=0)
        self.scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self._window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scroll.pack(side="right", fill="y")
        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<Enter>", self._bind_wheel)
        self.canvas.bind("<Leave>", self._unbind_wheel)
        self.show_message("Открываю книгу…")

    def destroy(self) -> None:
        self._unbind_wheel(None)
        super().destroy()

    def show_message(self, text: str) -> None:
        self._clear()
        ttk.Label(self.inner, text=text).pack(anchor="w", padx=6, pady=4)

    def set_sheets(self, items: list[tuple[str, int]]) -> None:
        self._clear()
        if not items:
            self.show_message("В книге нет листов.")
            return
        for name, count in items:
            var = tk.BooleanVar(value=count > 0)
            state = "normal" if count > 0 else "disabled"
            label = f"{name}  —  адресов: {count}"
            ttk.Checkbutton(
                self.inner,
                text=label,
                variable=var,
                state=state,
            ).pack(anchor="w", fill="x", padx=4, pady=1)
            self.rows.append((name, var, count))

    def selected_names(self) -> list[str]:
        return [name for name, var, count in self.rows if count > 0 and var.get()]

    def select_all(self) -> None:
        for _name, var, count in self.rows:
            if count > 0:
                var.set(True)

    def select_none(self) -> None:
        for _name, var, _count in self.rows:
            var.set(False)

    def _clear(self) -> None:
        for child in self.inner.winfo_children():
            child.destroy()
        self.rows.clear()

    def _on_inner_configure(self, _event) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event) -> None:
        self.canvas.itemconfigure(self._window_id, width=event.width)

    def _bind_wheel(self, _event) -> None:
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _unbind_wheel(self, _event) -> None:
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _on_mousewheel(self, event) -> None:
        if getattr(event, "delta", 0):
            self.canvas.yview_scroll(int(-event.delta / 120), "units")
        elif getattr(event, "num", None) == 4:
            self.canvas.yview_scroll(-1, "units")
        elif getattr(event, "num", None) == 5:
            self.canvas.yview_scroll(1, "units")


class LetterApp:
    def __init__(self, root: tk.Tk, workbook_path: str, draft_path: str) -> None:
        self.root = root
        self.workbook_path = workbook_path
        self.draft_path = draft_path
        self.loaded: list[tuple[str, ColumnParse]] = []
        self.loaded_path: str | None = None
        self.draft_subject = ""
        self.draft_body = ""
        self.draft_cc = ""
        self.draft_loaded = False
        self._busy = False
        self._font = self._choose_font()
        self.root.title("Banks — генератор писем")
        self.root.minsize(920, 640)
        self.root.geometry("1040x760")
        self._build()

    def _choose_font(self) -> tuple[str, int]:
        families = set(tkfont.families(self.root))
        if "Segoe UI" in families:
            chosen = "Segoe UI"
        elif "DejaVu Sans" in families:
            chosen = "DejaVu Sans"
        else:
            chosen = tkfont.nametofont("TkDefaultFont").actual()["family"]
        self.root.option_add("*Font", (chosen, 10))
        style = ttk.Style(self.root)
        style.configure(".", font=(chosen, 10))
        return (chosen, 10)

    def _build(self) -> None:
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(1, weight=1)

        ttk.Label(
            outer,
            text=(
                "Тема и текст берутся из черновика. "
                "Каждый адрес из таблицы получит отдельное письмо. "
                "Люди из копии черновика будут в Cc каждого письма."
            ),
            wraplength=980,
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        left = ttk.Frame(outer)
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 6))
        right = ttk.Frame(outer)
        right.grid(row=1, column=1, sticky="nsew", padx=(6, 0))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=2)
        right.rowconfigure(3, weight=1)

        book = ttk.LabelFrame(left, text="Книга Excel", padding=8)
        book.grid(row=0, column=0, sticky="ew")
        self.book_label = ttk.Label(book, text=self.workbook_path, wraplength=440, justify="left")
        self.book_label.pack(anchor="w")
        self.load_button = ttk.Button(book, text="Обновить", command=self.open_sources)
        self.load_button.pack(anchor="w", pady=(8, 0))

        sheets = ttk.LabelFrame(left, text="Листы", padding=8)
        sheets.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        toggles = ttk.Frame(sheets)
        toggles.pack(anchor="w")
        ttk.Button(toggles, text="Отметить все", command=self._select_all).pack(side="left")
        ttk.Button(toggles, text="Снять все", command=self._select_none).pack(side="left", padx=(6, 0))
        self.sheet_list = SheetList(sheets)
        self.sheet_list.pack(fill="both", expand=True, pady=(6, 0))

        letter = ttk.LabelFrame(right, text="Черновик", padding=8)
        letter.grid(row=0, column=0, sticky="nsew")
        self.draft_label = ttk.Label(letter, text=self.draft_path, wraplength=460, justify="left")
        self.draft_label.pack(anchor="w")
        ttk.Label(letter, text="Тема:").pack(anchor="w", pady=(6, 0))
        self.subject_label = ttk.Label(letter, text="—", wraplength=460, justify="left")
        self.subject_label.pack(anchor="w", fill="x")
        ttk.Label(letter, text="Текст письма:").pack(anchor="w", pady=(6, 0))
        self.body = tk.Text(letter, height=7, wrap="word", font=self._font)
        self.body.pack(fill="both", expand=True, pady=(2, 0))
        ttk.Label(letter, text="Копия (Cc):").pack(anchor="w", pady=(6, 0))
        self.cc_box = tk.Text(letter, height=3, wrap="word", font=self._font)
        self.cc_box.pack(fill="x", pady=(2, 0))
        self._show_draft("", "", "")

        attachment = ttk.LabelFrame(right, text="Вложение (необязательно)", padding=8)
        attachment.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(attachment, text="Путь к файлу:").pack(anchor="w")
        attachment_row = ttk.Frame(attachment)
        attachment_row.pack(fill="x", pady=(2, 0))
        self.attachment_var = tk.StringVar()
        ttk.Entry(attachment_row, textvariable=self.attachment_var).pack(side="left", fill="x", expand=True)
        ttk.Button(attachment_row, text="Обзор…", command=self.browse_attachment).pack(
            side="left", padx=(6, 0)
        )

        self.send_button = ttk.Button(right, text="Отправить", command=self.on_send)
        self.send_button.grid(row=2, column=0, sticky="w", pady=(10, 0))

        log_frame = ttk.LabelFrame(right, text="Журнал", padding=8)
        log_frame.grid(row=3, column=0, sticky="nsew", pady=(8, 0))
        self.log = ScrolledText(log_frame, height=6, wrap="word", font=self._font)
        self.log.pack(fill="both", expand=True)
        self.log.bind("<Key>", self._log_key)
        self.append_log("Открываю книгу и черновик…")

    def browse_attachment(self) -> None:
        path = filedialog.askopenfilename(title="Выберите вложение")
        if path:
            self.attachment_var.set(path)

    def open_workbook(self) -> None:
        if self._busy:
            return
        file_path = Path(self.workbook_path.strip().strip('"'))
        missing = missing_workbook_message(self.workbook_path)
        if missing:
            self.loaded = []
            self.loaded_path = None
            self.sheet_list.show_message("Книга не открыта.")
            self.append_log(one_line(missing))
            messagebox.showerror("Книга Excel", missing)
            return

        self._busy = True
        self.load_button.configure(state="disabled")
        self.root.update_idletasks()
        try:
            try:
                columns = read_workbook(file_path)
                parsed_rows = [(name, parse_column(cells)) for name, cells in columns]
            except ExcelError as exc:
                self.loaded = []
                self.loaded_path = None
                self.sheet_list.show_message("Книга не открыта.")
                self.append_log(one_line(exc))
                messagebox.showerror("Excel", str(exc))
                return
            except Exception as exc:
                self.loaded = []
                self.loaded_path = None
                self.sheet_list.show_message("Книга не открыта.")
                self.append_log(one_line(exc))
                messagebox.showerror("Excel", f"Не удалось прочитать книгу.\n{exc}")
                return
        finally:
            self._busy = False
            self.load_button.configure(state="normal")

        self.loaded = parsed_rows
        self.loaded_path = str(file_path.resolve())
        self.clear_log()
        self.sheet_list.set_sheets(
            [(name, len(parsed.addresses)) for name, parsed in parsed_rows]
        )
        for name, parsed in parsed_rows:
            for skipped in parsed.skipped:
                self.append_log(format_skip(name, skipped))
        total = sum(len(parsed.addresses) for _name, parsed in parsed_rows)
        self.append_log(f"Загружено листов: {len(parsed_rows)}. Уникальных адресов: {total}.")

    def open_draft(self) -> None:
        if self._busy:
            return
        missing = missing_draft_message(self.draft_path)
        if missing:
            self._clear_draft()
            self.append_log(one_line(missing))
            messagebox.showerror("Черновик", missing)
            return

        self._busy = True
        self.load_button.configure(state="disabled")
        self.root.update_idletasks()
        sender = OutlookSender()
        try:
            try:
                sender.connect()
                draft = sender.read_draft(self.draft_path)
            except OutlookError as exc:
                self._clear_draft()
                self.append_log(one_line(exc))
                messagebox.showerror("Черновик", str(exc))
                return
        finally:
            sender.release()
            self._busy = False
            self.load_button.configure(state="normal")

        self.draft_subject = draft.subject
        self.draft_body = draft.body
        self.draft_cc = draft.cc
        self.draft_loaded = True
        self._show_draft(draft.subject, draft.body, draft.cc)
        people = cc_people(draft.cc)
        if people:
            self.append_log(f"Черновик открыт. В копии: {len(people)}.")
        else:
            self.append_log("Черновик открыт. В копии никого нет.")

    def open_sources(self) -> None:
        self.open_workbook()
        self.open_draft()

    def on_send(self) -> None:
        if self._busy:
            return
        problem = delivery_problem(
            workbook=self.workbook_path,
            draft=self.draft_path,
            subject=self.draft_subject,
            body=self.draft_body,
            attachment=self.attachment_var.get(),
            sheets_loaded=self.loaded_path is not None,
            draft_loaded=self.draft_loaded,
        )
        if problem:
            messagebox.showwarning("Проверьте данные", problem)
            return

        subject = self.draft_subject.strip()
        body = self.draft_body.strip()
        attachment_text = self.attachment_var.get().strip()
        selected = set(self.sheet_list.selected_names())
        sheets = [
            {
                "name": name,
                "addresses": parsed.addresses,
                "selected": name in selected,
            }
            for name, parsed in self.loaded
        ]
        messages = plan_messages(
            sheets,
            subject,
            body,
            attachment_text or None,
            self.draft_cc,
        )
        if not messages:
            messagebox.showinfo(
                "Нечего отправлять",
                "Среди отмеченных листов нет адресов. Письма не отправлялись.",
            )
            return
        if not messagebox.askyesno(
            "Подтверждение отправки",
            confirmation_text(len(messages), subject, self.draft_cc),
        ):
            self.append_log("Отправка отменена.")
            return

        self._busy = True
        self.send_button.configure(state="disabled", text="Отправка…")
        self.load_button.configure(state="disabled")
        sender = OutlookSender()
        try:
            try:
                sender.connect()
            except OutlookError as exc:
                self.append_log(one_line(exc))
                messagebox.showerror("Outlook", str(exc))
                return
            try:
                archived = archive_sources(
                    Path(self.loaded_path),
                    Path(attachment_text) if attachment_text else None,
                    app_directory(),
                )
            except Exception as exc:
                self.append_log(one_line(f"Архив не создан: {exc}"))
                messagebox.showerror(
                    "Архив",
                    "Не удалось скопировать файлы перед отправкой. Письма не отправлялись.\n"
                    f"{exc}",
                )
                return

            outgoing = retarget_attachment(messages, archived.attachment)
            self.append_log(f"Архив: {archived.folder}")
            sent = 0
            failed = 0
            for message in outgoing:
                try:
                    sender.send(message)
                except Exception as exc:
                    failed += 1
                    self.append_log(
                        format_log(message["sheet"], message["to"], f"ошибка: {exc}")
                    )
                else:
                    sent += 1
                    self.append_log(format_log(message["sheet"], message["to"], "отправлено"))
                self.root.update_idletasks()
            self.append_log(f"Готово. Отправлено: {sent}. Ошибок: {failed}.")
        finally:
            sender.release()
            self._busy = False
            self.send_button.configure(state="normal", text="Отправить")
            self.load_button.configure(state="normal")

    def _select_all(self) -> None:
        self.sheet_list.select_all()

    def _select_none(self) -> None:
        self.sheet_list.select_none()

    def _clear_draft(self) -> None:
        self.draft_subject = ""
        self.draft_body = ""
        self.draft_cc = ""
        self.draft_loaded = False
        self._show_draft("", "", "")

    def _show_draft(self, subject: str, body: str, cc: str) -> None:
        self.subject_label.configure(text=subject.strip() or "—")
        people = cc_people(cc)
        cc_text = "\n".join(people) if people else "В черновике никого нет в копии."
        self._set_readonly(self.body, body.strip() or "Текст появится из черновика.")
        self._set_readonly(self.cc_box, cc_text)

    def _set_readonly(self, widget: tk.Text, text: str) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.configure(state="disabled")

    def clear_log(self) -> None:
        self.log.delete("1.0", "end")

    def append_log(self, line: str) -> None:
        self.log.insert("end", line + "\n")
        self.log.see("end")

    def _log_key(self, event) -> str | None:
        if event.state & 0x4 and event.keysym.lower() in {"c", "a"}:
            return None
        if event.keysym in {
            "Left",
            "Right",
            "Up",
            "Down",
            "Home",
            "End",
            "Prior",
            "Next",
            "Tab",
            "Shift_L",
            "Shift_R",
            "Control_L",
            "Control_R",
        }:
            return None
        return "break"


def run(workbook_path: str, draft_path: str) -> None:
    root = tk.Tk()
    app = LetterApp(root, workbook_path, draft_path)
    root.after(200, app.open_sources)
    root.mainloop()
