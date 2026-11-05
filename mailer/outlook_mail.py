"""Чтение черновика .msg и отправка одного письма через Outlook.Application."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

# Outlook constant olMailItem.
OL_MAIL_ITEM = 0
# MailItem.Close: не сохранять и не отправлять открытый черновик.
OL_DISCARD = 1


@dataclass(frozen=True)
class DraftLetter:
    subject: str
    body: str
    cc: str


class OutlookError(Exception):
    """Outlook недоступен или письмо нельзя собрать как одно адресное."""


class OutlookSender:
    """Сессия Outlook. Quit не вызывается: Dispatch подключается к почте пользователя."""

    def __init__(self) -> None:
        self._app = None

    def connect(self) -> None:
        try:
            import win32com.client
        except ImportError as exc:
            raise OutlookError(
                "Не установлен модуль pywin32.\n"
                "В папке приложения выполните:\n"
                "pip install -r requirements.txt"
            ) from exc
        try:
            app = win32com.client.Dispatch("Outlook.Application")
            _ = app.Version
        except Exception as exc:
            raise OutlookError(
                "Не удалось подключиться к Outlook. Установите Microsoft Outlook, "
                "войдите в учётную запись и откройте Outlook хотя бы один раз.\n"
                f"{exc}"
            ) from exc
        self._app = app

    def read_draft(self, path: str) -> DraftLetter:
        """Прочитать тему, текст и Cc из файла .msg. Письмо не отправляется.

        Outlook открывает такой файл методом OpenSharedItem: это сообщение
        с уже заполненной копией, а не обычный текст .eml.
        """
        if self._app is None:
            raise OutlookError("Outlook не подключён.")
        full_path = str(Path(path).expanduser().resolve())
        try:
            namespace = self._app.GetNamespace("MAPI")
            item = namespace.OpenSharedItem(full_path)
        except OutlookError:
            raise
        except Exception as exc:
            raise OutlookError(
                "Не удалось открыть черновик Outlook.\n"
                f"Искали здесь:\n{path}\n{exc}"
            ) from exc
        try:
            subject = str(getattr(item, "Subject", "") or "")
            body = str(getattr(item, "Body", "") or "")
            cc = str(getattr(item, "CC", "") or "")
        finally:
            try:
                item.Close(OL_DISCARD)
            except Exception:
                pass
        return DraftLetter(subject=subject, body=body, cc=cc)

    def send(self, message: dict) -> None:
        """Создать новое MailItem на один адрес и отправить его.

        В To попадает только message['to']. Cc берётся из черновика.
        Bcc остаётся пустым. Сам файл черновика не отправляется.
        """
        if self._app is None:
            raise OutlookError("Outlook не подключён.")
        recipient = str(message.get("to", "")).strip()
        if not recipient or any(separator in recipient for separator in (",", ";")):
            raise OutlookError("В письме должен быть ровно один адрес.")

        mail = self._app.CreateItem(OL_MAIL_ITEM)
        mail.Subject = str(message.get("subject", ""))
        mail.Body = str(message.get("body", ""))
        mail.To = recipient
        mail.CC = str(message.get("cc", "") or "")
        mail.BCC = ""
        attachment = message.get("attachment")
        if attachment:
            mail.Attachments.Add(str(Path(attachment).resolve()))
        mail.Send()

    def release(self) -> None:
        self._app = None
