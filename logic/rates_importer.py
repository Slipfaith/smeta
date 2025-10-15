"""Excel rates importer and language pair matching utilities."""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Dict, Iterable, IO, List, Optional, Tuple, Union

import langcodes
import openpyxl

from .xml_parser_common import expand_language_code, language_identity

RateRecord = Dict[str, float]
RatesMap = Dict[Tuple[str, str], RateRecord]


def _normalize_language(name: str) -> str:
    """Return a normalized ISO language code for *name*.

    The helper first tries :func:`logic.xml_parser_common.language_identity`
    so that script or territory hints from the Excel sheet are preserved.
    When the value cannot be resolved that way, we fall back to the
    previous heuristics based on :func:`langcodes.standardize_tag` and
    :func:`langcodes.find`, finally returning the lower-cased input if all
    lookups fail.
    """
    language, script, territory = language_identity(name)
    if language:
        parts = [language]
        if script:
            parts.append(script.lower())
        if territory:
            parts.append(territory.lower())
        return "-".join(parts)

    cleaned = re.sub(r"\s*\(.*?\)", "", name)
    cleaned = re.split(r"[,/\-]", cleaned, maxsplit=1)[0].strip()
    try:
        tag = langcodes.standardize_tag(cleaned)
    except Exception:
        tag = ""
    if tag:
        return tag.replace("_", "-").lower()

    try:
        return langcodes.find(cleaned).language
    except LookupError:
        return cleaned.lower()


def _language_name(code: str) -> str:
    """Return English display name for a language *code*."""
    if not code:
        return ""

    normalized = code.replace("_", "-")
    try:
        tag = langcodes.standardize_tag(normalized)
    except Exception:
        tag = normalized

    pretty = expand_language_code(tag, locale="en")
    if pretty:
        return pretty

    try:
        return langcodes.Language.get(tag).display_name("en")
    except Exception:
        try:
            return langcodes.Language.make(tag).display_name("en")
        except Exception:
            return code


def load_rates_from_excel(
    path_or_stream: Union[str, os.PathLike[str], IO[bytes]],
    currency: str,
    rate_type: str,
) -> RatesMap:
    """Load rates from *path* for the given *currency* and *rate_type*.

    Parameters
    ----------
    path_or_stream:
        Path to the Excel workbook or a binary stream containing the workbook contents.
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
    stream = path_or_stream
    if hasattr(stream, "seek"):
        stream.seek(0)
    wb = openpyxl.load_workbook(stream, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet {sheet_name} not found in workbook")

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
