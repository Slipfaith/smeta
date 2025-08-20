from __future__ import annotations
from typing import Dict, List, Tuple
import xml.etree.ElementTree as ET

from .service_config import ServiceConfig

ROW_NAMES = [row["name"] for row in ServiceConfig.TRANSLATION_ROWS]


def _norm_lang(code: str) -> str:
    if not code:
        return ""
    return code.split("-")[0].upper()


def _get_value(row: ET.Element, unit: str) -> float:
    unit = unit.capitalize()
    val = row.attrib.get(unit)
    if val is None:
        child = row.find(unit)
        if child is not None and child.text:
            val = child.text
    try:
        return float(val) if val is not None else 0.0
    except (TypeError, ValueError):
        return 0.0


def _parse_language_direction(ld: ET.Element, unit: str) -> Tuple[str, Dict[str, float]]:
    src = (
        ld.attrib.get("SourceLanguageCode")
        or ld.attrib.get("SourceLanguage")
        or ld.attrib.get("Source")
        or ld.findtext("SourceLanguageCode")
        or ld.findtext("SourceLanguage")
        or ld.findtext("Source")
        or ""
    )
    tgt = (
        ld.attrib.get("TargetLanguageCode")
        or ld.attrib.get("TargetLanguage")
        or ld.attrib.get("Target")
        or ld.findtext("TargetLanguageCode")
        or ld.findtext("TargetLanguage")
        or ld.findtext("Target")
        or ""
    )
    pair_key = f"{_norm_lang(src)} → {_norm_lang(tgt)}"
    values: Dict[str, float] = {name: 0.0 for name in ROW_NAMES}
    for row in ld.findall('.//Row'):
        cat = (row.attrib.get('Category') or row.attrib.get('Type') or '').lower()
        minv = row.attrib.get('MinimumMatchValue') or row.attrib.get('Min') or row.attrib.get('Minimum')
        maxv = row.attrib.get('MaximumMatchValue') or row.attrib.get('Max') or row.attrib.get('Maximum')
        try:
            imin = int(minv) if minv is not None else 0
        except ValueError:
            imin = 0
        try:
            imax = int(maxv) if maxv is not None else 0
        except ValueError:
            imax = 0
        val = _get_value(row, unit)
        if cat in ('new', 'no match') or (cat == 'fuzzy' and imax <= 74):
            values[ROW_NAMES[0]] += val
        elif cat == 'fuzzy' and imax <= 94:
            values[ROW_NAMES[1]] += val
        elif cat == 'fuzzy' and imax <= 99:
            values[ROW_NAMES[2]] += val
        elif cat in ('crossfilerepetitions', 'cross-file repetitions', 'repetition', 'repetitions', 'exact', 'incontextexact', 'in-context exact', 'contextmatch'):
            values[ROW_NAMES[3]] += val
        # Locked и прочее игнорируем
    return pair_key, values


def parse_reports(paths: List[str], unit: str = "Words") -> Tuple[Dict[str, Dict[str, float]], List[str]]:
    """Парсит Trados XML отчёты и возвращает агрегированные объёмы по парам языков."""
    results: Dict[str, Dict[str, float]] = {}
    warnings: List[str] = []
    for path in paths:
        try:
            tree = ET.parse(path)
            root = tree.getroot()
            dirs = root.findall('.//LanguageDirection')
            if not dirs:
                dirs = [root]
            for ld in dirs:
                pair_key, vals = _parse_language_direction(ld, unit)
                if not pair_key.strip(' →'):
                    continue
                acc = results.setdefault(pair_key, {name: 0.0 for name in ROW_NAMES})
                for name in ROW_NAMES:
                    acc[name] += vals.get(name, 0.0)
        except Exception as e:  # noqa: BLE001
            warnings.append(f"{path}: {e}")
    return results, warnings

