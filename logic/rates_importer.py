"""Excel rates importer and language pair matching utilities."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple

import langcodes
import openpyxl

RateRecord = Dict[str, float]
RatesMap = Dict[Tuple[str, str], RateRecord]


def _normalize_language(name: str) -> str:
    """Return a normalized ISO language code for *name*.

    The function is tolerant to different spellings and languages
    (e.g. "русский", "Russian (US)").  Region information in
    parenthesis is ignored.
    """
    cleaned = re.sub(r"\s*\(.*?\)", "", name).strip()
    return langcodes.find(cleaned).language


def _language_name(code: str) -> str:
    """Return English display name for a language *code*."""
    return langcodes.Language.make(code).display_name("en")


def load_rates_from_excel(path: str, currency: str, rate_type: str) -> RatesMap:
    """Load rates from *path* for the given *currency* and *rate_type*.

    Parameters
    ----------
    path:
        Path to the Excel workbook.
    currency:
        Currency code selected in the GUI (e.g. ``"USD"``).
    rate_type:
        Rate table version, such as ``"R1"`` or ``"R2"``.

    Returns
    -------
    dict
        Mapping of ``(src_lang, tgt_lang)`` pairs to rate dictionaries
        with keys ``basic``, ``complex`` and ``hour``.
    """
    sheet_name = f"{rate_type}_{currency}".upper()
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet {sheet_name} not found in {path}")

    ws = wb[sheet_name]
    rates: RatesMap = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        src, tgt, basic, complex_, hour = row
        if not src or not tgt:
            continue
        src_code = _normalize_language(str(src))
        tgt_code = _normalize_language(str(tgt))
        rates[(src_code, tgt_code)] = {
            "basic": float(basic or 0),
            "complex": float(complex_ or 0),
            "hour": float(hour or 0),
        }
    return rates


@dataclass
class PairMatch:
    """Result of matching a GUI language pair with Excel rates."""

    gui_source: str
    gui_target: str
    excel_source: str
    excel_target: str
    rates: Optional[RateRecord]


def match_pairs(gui_pairs: Iterable[Tuple[str, str]], rates: RatesMap) -> List[PairMatch]:
    """Match GUI language pairs against loaded *rates*.

    Parameters
    ----------
    gui_pairs:
        Iterable of ``(source, target)`` language names as entered in the GUI.
    rates:
        Rates map produced by :func:`load_rates_from_excel`.

    Returns
    -------
    list[PairMatch]
        A list describing how each GUI pair was matched.  If a corresponding
        rate was not found, ``rates`` attribute will be ``None``.
    """
    results: List[PairMatch] = []
    for gui_src, gui_tgt in gui_pairs:
        src_code = _normalize_language(gui_src)
        tgt_code = _normalize_language(gui_tgt)
        rate = rates.get((src_code, tgt_code))
        excel_src = _language_name(src_code)
        excel_tgt = _language_name(tgt_code)
        results.append(
            PairMatch(
                gui_source=gui_src,
                gui_target=gui_tgt,
                excel_source=excel_src,
                excel_target=excel_tgt,
                rates=rate,
            )
        )
    return results
