"""Shared behaviour for widgets that manage rate and currency data."""
from __future__ import annotations

from typing import Iterable, Iterator

from PySide6.QtWidgets import QTableWidgetItem

from logic.number_format import (
    convert_rate_value,
    decimal_separator_for_lang,
)


class RateCurrencyMixin:
    """Provide helpers for converting rate values inside widgets."""

    def iter_rate_items(self) -> Iterable[QTableWidgetItem | None]:  # pragma: no cover - abstract
        raise NotImplementedError

    def _rate_decimal_separator(self) -> str:
        lang = getattr(self, "lang", "ru")
        return decimal_separator_for_lang(lang)

    def convert_rates(self, multiplier: float) -> None:
        separator = self._rate_decimal_separator()
        for item in self._iter_existing_items():
            value = item.text() if item is not None else "0"
            updated = convert_rate_value(value, multiplier, separator)
            if item is None:
                continue
            item.setText(updated)
        self._after_rates_converted()

    def _iter_existing_items(self) -> Iterator[QTableWidgetItem | None]:
        for item in self.iter_rate_items():
            yield item

    def _after_rates_converted(self) -> None:
        update = getattr(self, "update_sums", None)
        if callable(update):
            update()
