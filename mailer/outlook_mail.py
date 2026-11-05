"""Отправка одного письма на адрес через Outlook.Application (win32com)."""

from __future__ import annotations

from pathlib import Path

# Outlook constant olMailItem.
OL_MAIL_ITEM = 0


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

    def send(self, message: dict) -> None:
        """Создать MailItem на один адрес и отправить его.

        В To попадает только message['to']. Cc и Bcc очищаются.
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
        mail.CC = ""
        mail.BCC = ""
        attachment = message.get("attachment")
        if attachment:
            mail.Attachments.Add(str(Path(attachment).resolve()))
        mail.Send()

    def release(self) -> None:
        self._app = None
