from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple
import xml.etree.ElementTree as ET

from logger import get_logger

from .service_config import ServiceConfig
from .xml_parser_common import expand_language_code, resolve_language_display


ROW_NAMES = ServiceConfig.ROW_NAMES
SMARTCAT_NS = "urn:schemas-microsoft-com:office:spreadsheet"


logger = get_logger(__name__)


def is_smartcat_report(path: str) -> bool:
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            head = f.read(2048)
    except OSError:
        head = ""

    lower = head.lower()
    if "<workbook" in lower and "urn:schemas-microsoft-com:office:spreadsheet" in lower:
        return True

    if "statistics for project" in lower:
        return True

    filename = Path(path).name
    normalized = filename.lstrip("\ufeff").lstrip()
    normalized_lower = normalized.lower()

    if re.search(r"^\[[^\]]{1,15}\]", normalized_lower):
        return True

    if "statistics for project" in normalized_lower:
        return True

    return False


def parse_smartcat_report(
    path: str, unit: str
) -> Tuple[Dict[str, Dict[str, float]], List[str], bool, str]:
    filename = Path(path).name
    logger.info("Detected Smartcat report: %s", path)

    results: Dict[str, Dict[str, float]] = {}
    warnings: List[str] = []
    processed = False
    placeholder = Path(path).stem

    try:
        tree = ET.parse(path)
    except ET.ParseError as exc:
        msg = f"{filename}: XML Parse Error - {exc}"
        logger.error(msg)
        warnings.append(msg)
        return results, warnings, processed, placeholder
    except Exception as exc:
        msg = f"{filename}: Unexpected error - {exc}"
        logger.exception(msg)
        warnings.append(msg)
        return results, warnings, processed, placeholder

    root = tree.getroot()
    logger.debug("Root element: %s", root.tag)

    target_lang = _extract_smartcat_target_language(root)
    if not target_lang:
        target_lang = _target_language_from_filename(filename)
    if target_lang:
        placeholder = target_lang
    pair_key = target_lang or placeholder
    logger.info("Smartcat target language: '%s'", pair_key)

    statistics_rows: List[Tuple[str, float]] = []
    for worksheet in root.findall(f"{{{SMARTCAT_NS}}}Worksheet"):
        sheet_name = worksheet.get(f"{{{SMARTCAT_NS}}}Name", "") or "<unnamed>"
        logger.debug("  Inspecting worksheet: %s", sheet_name)
        rows = _worksheet_to_rows(worksheet)
        statistics_rows = _find_smartcat_statistics(rows, unit)
        if statistics_rows:
            logger.info("  → Statistics found in worksheet '%s'", sheet_name)
            break

    values = {name: 0.0 for name in ROW_NAMES}
    statistics_found = False

    if statistics_rows:
        statistics_found = True
        for label, number in statistics_rows:
            category = _categorize_smartcat_row(label)
            if not category:
                logger.debug("    Skipping row '%s'", label)
                continue
            logger.debug("    %s -> %s: %s", label, category, number)
            values[category] += number
    else:
        fallback_values, fallback_found = _parse_smartcat_task_statistics(root, unit)
        if fallback_found:
            statistics_found = True
            values = fallback_values

    if not statistics_found:
        msg = f"{filename}: Не удалось извлечь статистику Smartcat"
        logger.warning(msg)
        warnings.append(msg)
        return results, warnings, processed, pair_key

    total = sum(values.values())
    results[pair_key] = values

    if total > 0:
        processed = True
        logger.info("Smartcat report processed: %s total %s", pair_key, total)
    else:
        msg = f"{filename}: Статистика Smartcat не содержит слов"
        logger.warning(msg)
        warnings.append(msg)

    return results, warnings, processed, pair_key


def _target_language_from_filename(filename: str) -> str:
    match = re.search(r"\[([^\[\]]{1,30})\]", filename)
    if not match:
        return ""

    raw_contents = match.group(1).strip()
    if not raw_contents:
        return ""

    tokens = [token.strip() for token in re.split(r"[,;/\\|]+", raw_contents) if token.strip()]
    if not tokens:
        tokens = [token.strip() for token in re.split(r"\s+", raw_contents) if token.strip()]

    for token in tokens:
        display = resolve_language_display(token)
        if display:
            return display

    return resolve_language_display(raw_contents)


def _parse_number(text: str) -> float:
    if not text:
        return 0.0

    cleaned = text.strip().replace("\xa0", "").replace(" ", "")
    if not cleaned:
        return 0.0

    if cleaned.count(",") and cleaned.count("."):
        if cleaned.rfind(".") > cleaned.rfind(","):
            cleaned = cleaned.replace(",", "")
        else:
            cleaned = cleaned.replace(".", "")
            cleaned = cleaned.replace(",", ".")
    elif "," in cleaned and "." not in cleaned:
        cleaned = cleaned.replace(",", ".")

    try:
        return float(cleaned)
    except ValueError:
        cleaned = re.sub(r"[^0-9.+-]", "", cleaned)
        try:
            return float(cleaned)
        except ValueError:
            return 0.0


def _smartcat_unit_keywords(unit: str) -> List[str]:
    unit_lc = unit.lower()
    if "char" in unit_lc or "символ" in unit_lc:
        return ["character", "characters", "символ", "символы", "знаков"]
    return ["word", "words", "слово", "слова"]


def _worksheet_to_rows(worksheet: ET.Element) -> List[List[str]]:
    table = worksheet.find(f"{{{SMARTCAT_NS}}}Table")
    if table is None:
        return []

    rows: List[List[str]] = []
    for row in table.findall(f"{{{SMARTCAT_NS}}}Row"):
        values: List[str] = []
        for cell in row.findall(f"{{{SMARTCAT_NS}}}Cell"):
            index_attr = cell.get(f"{{{SMARTCAT_NS}}}Index")
            if index_attr:
                try:
                    index = int(index_attr) - 1
                    while len(values) < index:
                        values.append("")
                except ValueError:
                    pass

            data_elem = cell.find(f"{{{SMARTCAT_NS}}}Data")
            text = ""
            if data_elem is not None and data_elem.text:
                text = data_elem.text.strip()
            values.append(text)
        rows.append(values)
    return rows


def _smartcat_candidates_from_text(text: str) -> List[str]:
    candidates: List[str] = []
    if not text:
        return candidates

    stripped = text.strip()
    if not stripped:
        return candidates

    bracket_matches = re.findall(r"\[([A-Za-z0-9_-]{2,})\]", stripped)
    candidates.extend(bracket_matches)

    arrow_match = re.search(r"(?:→|->|➔|➡|➞|⟶|⟹)\s*([^\]]+)$", stripped)
    if arrow_match:
        candidates.append(arrow_match.group(1).strip())

    if ":" in stripped:
        label, _, value = stripped.partition(":")
        if any(
            keyword in label.lower()
            for keyword in ("target", "language", "язык", "целевой")
        ):
            candidates.append(value.strip())

    lower = stripped.lower()
    if lower.startswith("statistics for project"):
        rest = stripped[len("statistics for project") :].strip(" :-")
        if rest:
            candidates.append(rest)
    elif lower.startswith("statistics"):
        rest = stripped[len("statistics") :].strip(" :-")
        if rest and len(rest.split()) <= 3:
            candidates.append(rest)

    if len(stripped) <= 15 and re.fullmatch(r"[A-Za-z]{2,3}(?:-[A-Za-z]{2,3})?", stripped):
        candidates.append(stripped)
    elif (
        len(stripped.split()) <= 3
        and not any(
            keyword in lower
            for keyword in (
                "statistics",
                "project",
                "match",
                "segment",
                "words",
                "characters",
                "total",
            )
        )
    ):
        candidates.append(stripped)

    return [c for c in candidates if c]


def _extract_smartcat_target_language(root: ET.Element) -> str:
    codes: List[str] = []

    for elem in root.iter():
        tag = elem.tag
        if "}" in tag:
            tag = tag.split("}", 1)[1]
        tag_lc = tag.lower()
        if tag_lc != "language":
            continue

        attr_map = {key.lower(): value for key, value in elem.attrib.items() if value}
        type_value = attr_map.get("type", "").lower()
        if type_value and type_value not in {"target", "targetlanguage", "target language"}:
            continue

        for key in ("names", "name", "code"):
            if key in attr_map:
                tokens = re.split(r"[,;/\\s]+", attr_map[key])
                codes.extend(filter(None, tokens))

    for code in codes:
        display = expand_language_code(code)
        if display:
            return display

    candidates: List[str] = []
    for worksheet in root.findall(f"{{{SMARTCAT_NS}}}Worksheet"):
        name_attr = worksheet.get(f"{{{SMARTCAT_NS}}}Name", "")
        candidates.extend(_smartcat_candidates_from_text(name_attr))

        table = worksheet.find(f"{{{SMARTCAT_NS}}}Table")
        if table is None:
            continue

        for row in table.findall(f"{{{SMARTCAT_NS}}}Row"):
            for cell in row.findall(f"{{{SMARTCAT_NS}}}Cell"):
                data_elem = cell.find(f"{{{SMARTCAT_NS}}}Data")
                if data_elem is None or not data_elem.text:
                    continue
                text = data_elem.text.strip()
                if not text:
                    continue
                lower = text.lower()
                if any(
                    keyword in lower for keyword in ("target", "language", "язык", "целевой")
                ):
                    candidates.extend(_smartcat_candidates_from_text(text))
                elif "[" in text and "]" in text:
                    candidates.extend(_smartcat_candidates_from_text(text))
            if candidates:
                break
        if candidates:
            break

    seen = set()
    for candidate in candidates:
        candidate_norm = candidate.strip()
        if not candidate_norm:
            continue
        key = candidate_norm.lower()
        if key in seen:
            continue
        seen.add(key)
        display = resolve_language_display(candidate_norm)
        if display:
            return display

    return ""


def _find_smartcat_statistics(rows: List[List[str]], unit: str) -> List[Tuple[str, float]]:
    if not rows:
        return []

    keywords = _smartcat_unit_keywords(unit)

    header_idx: Optional[int] = None
    value_col: Optional[int] = None

    for idx, row in enumerate(rows):
        normalized = [cell.strip().lower() for cell in row]
        if not any(normalized):
            continue
        first_cell = normalized[0] if normalized else ""
        if header_idx is None and any(
            token in first_cell for token in ("segment", "match", "тип", "совпад")
        ):
            for col, cell in enumerate(normalized):
                if any(keyword in cell for keyword in keywords):
                    header_idx = idx
                    value_col = col
                    break
        if header_idx is not None:
            break

    if header_idx is None or value_col is None:
        return []

    stats: List[Tuple[str, float]] = []
    blank_rows = 0

    for row in rows[header_idx + 1 :]:
        normalized = [cell.strip() for cell in row]
        if not any(normalized):
            blank_rows += 1
            if blank_rows >= 2 and stats:
                break
            continue

        blank_rows = 0

        if len(row) <= value_col:
            continue

        label = normalized[0]
        if not label:
            continue

        lower_label = label.lower()
        if lower_label in {"total", "subtotal", "итого"}:
            break
        if any(
            keyword in lower_label
            for keyword in ("segment type", "match type", "workflow", "language pair")
        ):
            continue

        value = _parse_number(row[value_col])
        if value <= 0:
            continue

        stats.append((label, value))

    return stats


def _categorize_smartcat_row(label: str) -> Optional[str]:
    norm = re.sub(r"\s+", " ", label).strip()
    if not norm:
        return None

    lower = norm.lower()
    if any(keyword in lower for keyword in ("repeat", "повтор", "context", "perfect")):
        return ROW_NAMES[3]

    digits = [int(val) for val in re.findall(r"\d+", lower)]

    if any(keyword in lower for keyword in ("machine", "mt", "итого", "total")):
        if "total" in lower or "итого" in lower:
            return None
        if "machine" in lower or "mt" in lower:
            return None

    if any(value >= 100 for value in digits):
        return ROW_NAMES[3]

    if any(95 <= value <= 99 for value in digits):
        return ROW_NAMES[2]

    if any(75 <= value <= 94 for value in digits):
        return ROW_NAMES[1]

    if any(keyword in lower for keyword in ("no match", "без совп", "новые", "new")):
        return ROW_NAMES[0]

    if any(60 <= value <= 74 for value in digits):
        return ROW_NAMES[0]

    if "tm" in lower and digits:
        if any(value >= 95 for value in digits):
            return ROW_NAMES[2]
        if any(value >= 75 for value in digits):
            return ROW_NAMES[1]

    if "match" in lower and "100" in lower:
        return ROW_NAMES[3]

    return ROW_NAMES[0]


def _strip_namespace(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _iter_elements_by_name(root: ET.Element, local_name: str) -> List[ET.Element]:
    target = local_name.lower()
    return [elem for elem in root.iter() if _strip_namespace(elem.tag).lower() == target]


def _find_child_by_name(parent: ET.Element, local_name: str) -> Optional[ET.Element]:
    target = local_name.lower()
    for child in parent:
        if _strip_namespace(child.tag).lower() == target:
            return child
    return None


def _iter_children_by_name(parent: ET.Element, local_name: str) -> List[ET.Element]:
    target = local_name.lower()
    return [child for child in parent if _strip_namespace(child.tag).lower() == target]


def _stat_value_from_attributes(elem: ET.Element, unit: str) -> str:
    attr_map = {
        key.lower(): value
        for key, value in elem.attrib.items()
        if value is not None
    }

    unit_lc = unit.lower()
    candidates = [unit_lc]

    if unit_lc.endswith("s"):
        candidates.append(unit_lc[:-1])
    else:
        candidates.append(f"{unit_lc}s")

    candidates.extend(
        [
            "words",
            "word",
            "characters",
            "character",
            "chars",
            "char",
            "символ",
            "символы",
            "знаков",
            "count",
            "value",
        ]
    )

    for candidate in candidates:
        key = candidate.lower()
        if key in attr_map:
            return attr_map[key]

    if elem.text:
        return elem.text.strip()
    return ""


def _parse_smartcat_analyse_element(
    analyse: ET.Element, unit: str
) -> Tuple[Dict[str, float], bool]:
    values = {name: 0.0 for name in ROW_NAMES}
    has_data = False

    def _add_value(row_name: str, amount: float) -> None:
        nonlocal has_data
        if amount <= 0:
            return
        values[row_name] += amount
        has_data = True

    unit_attr = unit.lower()

    new_elem = _find_child_by_name(analyse, "new")
    if new_elem is not None:
        amount = _parse_number(_stat_value_from_attributes(new_elem, unit_attr))
        _add_value(ROW_NAMES[0], amount)

    for fuzzy in _iter_children_by_name(analyse, "fuzzy"):
        amount = _parse_number(_stat_value_from_attributes(fuzzy, unit_attr))
        if amount <= 0:
            continue
        max_val = int(_parse_number(fuzzy.get("max", "0")))

        if max_val <= 74:
            _add_value(ROW_NAMES[0], amount)
        elif max_val <= 94:
            _add_value(ROW_NAMES[1], amount)
        elif max_val <= 99:
            _add_value(ROW_NAMES[2], amount)
        else:
            _add_value(ROW_NAMES[3], amount)

    for tag in (
        "exact",
        "repeated",
        "crossFileRepeated",
        "inContextExact",
        "perfect",
        "locked",
    ):
        for elem in _iter_children_by_name(analyse, tag):
            amount = _parse_number(_stat_value_from_attributes(elem, unit_attr))
            _add_value(ROW_NAMES[3], amount)

    return values, has_data


def _parse_smartcat_task_statistics(
    root: ET.Element, unit: str
) -> Tuple[Dict[str, float], bool]:
    logger.debug("  Attempting Smartcat task-format parsing...")

    values = {name: 0.0 for name in ROW_NAMES}
    processed_any = False
    processed_ids: Set[int] = set()

    file_elements = _iter_elements_by_name(root, "file")
    if file_elements:
        logger.info("  Found %s file elements in Smartcat report", len(file_elements))
        for idx, file_elem in enumerate(file_elements, 1):
            file_name = file_elem.get("name") or file_elem.get("path") or f"file_{idx}"
            logger.info(
                "    Processing file %s/%s: %s", idx, len(file_elements), file_name
            )
            analyse = _find_child_by_name(file_elem, "analyse")
            if analyse is None:
                logger.warning("      No <analyse> element found")
                continue

            analyse_values, has_data = _parse_smartcat_analyse_element(analyse, unit)
            if not has_data:
                logger.debug("      Analyse element contained no data")
                continue

            processed_any = True
            processed_ids.add(id(analyse))
            for name in ROW_NAMES:
                amount = analyse_values[name]
                if amount > 0:
                    values[name] += amount
                    logger.debug(
                        "      %s: +%s (now %s)", name, amount, values[name]
                    )

        if processed_any:
            return values, True

    analyse_elements = [
        elem
        for elem in _iter_elements_by_name(root, "analyse")
        if id(elem) not in processed_ids
    ]
    if analyse_elements:
        logger.info(
            "  Found %s analyse elements for fallback parsing", len(analyse_elements)
        )
        for idx, analyse in enumerate(analyse_elements, 1):
            analyse_values, has_data = _parse_smartcat_analyse_element(analyse, unit)
            if not has_data:
                continue

            processed_any = True
            logger.info("    Analyse block #%s", idx)
            for name in ROW_NAMES:
                amount = analyse_values[name]
                if amount > 0:
                    values[name] += amount
                    logger.debug(
                        "      %s: +%s (now %s)", name, amount, values[name]
                    )

    return values, processed_any
