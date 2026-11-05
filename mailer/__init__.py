"""Генератор отдельных писем Outlook по листам Excel."""

from mailer.planning import parse_column, plan_messages

__all__ = ["parse_column", "plan_messages"]
