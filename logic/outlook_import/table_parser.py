"""Utility helpers for parsing HTML tables from Outlook messages."""

from __future__ import annotations

import html
from html.parser import HTMLParser
from typing import List


class _SimpleTableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.tables: List[List[List[str]]] = []
        self._current_table: List[List[str]] | None = None
        self._current_row: List[str] | None = None
        self._current_cell: List[str] | None = None
        self._table_depth = 0

    # HTMLParser overrides -------------------------------------------------
    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "table":
            self._table_depth += 1
            if self._table_depth == 1:
                self._current_table = []
        elif tag in {"tr", "th", "td"} and self._table_depth == 1:
            if tag == "tr":
                self._current_row = []
            elif tag in {"td", "th"}:
                self._current_cell = []

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag == "table":
            if self._table_depth == 1 and self._current_table is not None:
                if self._current_row is not None:
                    self._finalize_row()
                if self._current_table:
                    self.tables.append(self._current_table)
            self._current_table = None
            self._current_row = None
            self._current_cell = None
            self._table_depth = max(0, self._table_depth - 1)
        elif tag == "tr" and self._table_depth == 1:
            self._finalize_row()
        elif tag in {"td", "th"} and self._table_depth == 1:
            if self._current_cell is not None:
                text = _clean_text("".join(self._current_cell))
                if self._current_row is not None:
                    self._current_row.append(text)
            self._current_cell = None

    def handle_data(self, data):
        if self._current_cell is not None:
            self._current_cell.append(data)

    # Internal helpers -----------------------------------------------------
    def _finalize_row(self):
        if self._current_row is not None and self._current_table is not None:
            cleaned = [cell for cell in self._current_row]
            if any(cell.strip() for cell in cleaned):
                self._current_table.append(cleaned)
        self._current_row = None


def _clean_text(value: str) -> str:
    return html.unescape(value).strip()


def extract_first_table(html_content: str) -> List[List[str]]:
    """Return the first table found in ``html_content``."""

    parser = _SimpleTableParser()
    parser.feed(html_content)
    parser.close()
    return parser.tables[0] if parser.tables else []
