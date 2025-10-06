from __future__ import annotations

from collections import defaultdict
from typing import Dict, List, Optional, Set, Tuple
import xml.etree.ElementTree as ET
import re
from pathlib import Path

from .service_config import ServiceConfig
from .xml_parser_common import expand_language_code, normalize_language_name
from .sc_xml_parser import is_smartcat_report, parse_smartcat_report
from logger import get_logger


ROW_NAMES = ServiceConfig.ROW_NAMES


logger = get_logger(__name__)


def _extract_languages_from_filename(filename: str) -> Tuple[str, str]:
    """Извлекает языки из имени файла типа 'Analyze Files en-US_ru-RU(23).xml'."""
    logger.debug("Extracting languages from filename: %s", filename)

    pattern = r"([a-z]{2,3}(?:-[A-Z]{2})?)[_-]([a-z]{2,3}(?:-[A-Z]{2})?)"
    match = re.search(pattern, filename, re.IGNORECASE)

    if match:
        src = match.group(1)
        tgt = match.group(2)
        logger.debug("  Found language pattern: %s -> %s", src, tgt)

        src_expanded = expand_language_code(src)
        tgt_expanded = expand_language_code(tgt)

        logger.debug("  Expanded: %s -> %s", src_expanded, tgt_expanded)
        return src_expanded, tgt_expanded

    logger.debug("  No language pattern found in filename")
    return "", ""


def _extract_language_from_taskinfo(taskinfo: ET.Element) -> str:
    """Извлекает целевой язык из элемента taskInfo."""
    logger.debug("Extracting language from taskInfo...")

    lang_element = taskinfo.find("language")
    if lang_element is not None:
        lang_name = lang_element.get("name", "").strip()
        lcid = lang_element.get("lcid", "").strip()

        logger.debug("  Language element found: name='%s', lcid='%s'", lang_name, lcid)

        normalized = normalize_language_name(lang_name)
        if normalized:
            logger.debug("  -> Normalized language: '%s'", normalized)
            return normalized

        normalized = normalize_language_name(lcid)
        if normalized:
            logger.debug("  -> Normalized language from LCID: '%s'", normalized)
            return normalized

        if lang_name:
            logger.debug("  -> Returning raw language name: '%s'", lang_name)
            return lang_name

    logger.debug("  No language found in taskInfo")
    return ""


def _parse_analyse_element(analyse: ET.Element, unit: str = "words") -> Dict[str, float]:
    """Парсит элемент <analyse> и возвращает объемы по категориям."""
    logger.debug("  Parsing analyse element...")

    values = {name: 0.0 for name in ROW_NAMES}
    unit_attr = unit.lower()

    new_elem = analyse.find("new")
    if new_elem is not None:
        new_words = float(new_elem.get(unit_attr, 0))
        values[ROW_NAMES[0]] += new_words
        logger.debug("    New words: %s", new_words)

    fuzzy_elements = analyse.findall("fuzzy")
    for fuzzy in fuzzy_elements:
        min_val = int(fuzzy.get("min", 0))
        max_val = int(fuzzy.get("max", 100))
        words = float(fuzzy.get(unit_attr, 0))

        logger.debug("    Fuzzy %s-%s%%: %s words", min_val, max_val, words)

        if words > 0:
            if max_val <= 74:
                values[ROW_NAMES[0]] += words
            elif max_val <= 94:
                values[ROW_NAMES[1]] += words
            elif max_val <= 99:
                values[ROW_NAMES[2]] += words

    exact_elem = analyse.find("exact")
    if exact_elem is not None:
        exact_words = float(exact_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += exact_words
        logger.debug("    Exact matches: %s", exact_words)

    repeated_elem = analyse.find("repeated")
    if repeated_elem is not None:
        repeated_words = float(repeated_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += repeated_words
        logger.debug("    Repeated: %s", repeated_words)

    cross_repeated_elem = analyse.find("crossFileRepeated")
    if cross_repeated_elem is not None:
        cross_words = float(cross_repeated_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += cross_words
        logger.debug("    Cross-file repeated: %s", cross_words)

    in_context_elem = analyse.find("inContextExact")
    if in_context_elem is not None:
        in_context_words = float(in_context_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += in_context_words
        logger.debug("    In-context exact: %s", in_context_words)

    perfect_elem = analyse.find("perfect")
    if perfect_elem is not None:
        perfect_words = float(perfect_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += perfect_words
        logger.debug("    Perfect matches: %s", perfect_words)

    locked_elem = analyse.find("locked")
    if locked_elem is not None:
        locked_words = float(locked_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += locked_words
        logger.debug("    Locked: %s", locked_words)

    total_words = sum(values.values())
    logger.debug("    Total words processed: %s", total_words)

    return values


def _ensure_placeholder_entry(
    results: Dict[str, Dict[str, float]], preferred_name: str, fallback: str
) -> str:
    base = preferred_name.strip() or fallback.strip()
    if not base:
        base = "Пустая языковая пара"

    candidate = base
    index = 2
    while candidate in results:
        candidate = f"{base} ({index})"
        index += 1

    results[candidate] = {name: 0.0 for name in ROW_NAMES}
    return candidate


def _merge_pair_results(
    target: Dict[str, Dict[str, float]], additions: Dict[str, Dict[str, float]]
) -> None:
    for pair_key, values in additions.items():
        if pair_key not in target:
            target[pair_key] = {name: 0.0 for name in ROW_NAMES}
        for name in ROW_NAMES:
            target[pair_key][name] += values.get(name, 0.0)


def _parse_trados_report(
    path: str, unit: str
) -> Tuple[Dict[str, Dict[str, float]], List[str], bool, str]:
    filename = Path(path).name
    logger.info("Processing Trados report: %s", path)

    results: Dict[str, Dict[str, float]] = {}
    warnings: List[str] = []
    processed = False
    placeholder = Path(path).stem

    try:
        tree = ET.parse(path)
        root = tree.getroot()

        logger.debug("Root element: %s", root.tag)
        logger.debug("Root attributes: %s", root.attrib)

        if root.tag != "task":
            logger.warning("Expected 'task' root element, got '%s'", root.tag)

        taskinfo = root.find("taskInfo")
        if taskinfo is None:
            warning_msg = f"{filename}: No taskInfo element found"
            logger.error(warning_msg)
            warnings.append(warning_msg)
            return results, warnings, processed, placeholder

        logger.debug("TaskInfo found: %s", taskinfo.attrib)

        filename_only = Path(path).name
        src_lang, tgt_lang = _extract_languages_from_filename(filename_only)
        taskinfo_lang = _extract_language_from_taskinfo(taskinfo)

        pair_key: Optional[str] = None
        determined_source_lang = ""
        determined_target_lang = ""

        if src_lang and tgt_lang:
            determined_source_lang = src_lang
            determined_target_lang = tgt_lang
            logger.info("Language pair from filename: %s → %s", src_lang, tgt_lang)
            if (
                taskinfo_lang
                and len(tgt_lang) <= 3
                and len(taskinfo_lang) > len(tgt_lang)
            ):
                determined_target_lang = taskinfo_lang
            pair_key = f"{determined_source_lang} → {determined_target_lang}"
        elif taskinfo_lang:
            determined_source_lang = "EN"
            determined_target_lang = taskinfo_lang
            pair_key = f"EN → {taskinfo_lang}"
            logger.info("Language pair from taskInfo: %s", pair_key)

        if not pair_key:
            warning_msg = (
                f"{filename}: Could not determine language pair (src='{src_lang}', tgt='{tgt_lang}', taskinfo='{taskinfo_lang}')."
                " Please assign the correct language pair manually."
            )
            logger.error(warning_msg)
            warnings.append(warning_msg)
            return results, warnings, processed, placeholder

        placeholder = determined_target_lang or pair_key

        logger.info(
            "Final determined pair: '%s' → '%s'",
            determined_source_lang,
            determined_target_lang,
        )
        logger.debug("Pair key: '%s'", pair_key)

        file_elements = root.findall("file")
        logger.info("Found %s file elements", len(file_elements))

        if not file_elements:
            warning_msg = f"{filename}: No file elements found"
            logger.error(warning_msg)
            warnings.append(warning_msg)
            return results, warnings, processed, placeholder

        pair_values = {name: 0.0 for name in ROW_NAMES}
        pair_total_words = 0.0
        files_processed_in_pair = 0

        for j, file_elem in enumerate(file_elements):
            file_name = file_elem.get("name", f"file_{j}")
            logger.info(
                "Processing file %s/%s: %s", j + 1, len(file_elements), file_name
            )

            analyse_elem = file_elem.find("analyse")
            if analyse_elem is None:
                logger.warning("No analyse element in file %s", file_name)
                continue

            file_values = _parse_analyse_element(analyse_elem, unit)

            for name in ROW_NAMES:
                add_val = file_values[name]
                pair_values[name] += add_val
                if add_val > 0:
                    logger.debug(
                        "    %s: +%s (now %s)", name, add_val, pair_values[name]
                    )

            file_total = sum(file_values.values())
            pair_total_words += file_total
            logger.debug("    File total: %s words", file_total)
            files_processed_in_pair += 1

        results[pair_key] = pair_values

        logger.info(
            "Pair %s total: %s words from %s files",
            pair_key,
            pair_total_words,
            files_processed_in_pair,
        )

        if pair_total_words > 0:
            processed = True
            logger.info("Successfully processed report: %s", pair_key)
        else:
            warning_msg = f"{filename}: No words found in any file"
            logger.warning(warning_msg)
            warnings.append(warning_msg)

    except ET.ParseError as exc:
        error_msg = f"{filename}: XML Parse Error - {exc}"
        logger.error(error_msg)
        warnings.append(error_msg)
    except Exception as exc:
        error_msg = f"{filename}: Unexpected error - {exc}"
        logger.exception(error_msg)
        warnings.append(error_msg)

    return results, warnings, processed, placeholder


def parse_reports(
    paths: List[str], unit: str = "Words"
) -> Tuple[Dict[str, Dict[str, float]], List[str], Dict[str, List[str]]]:
    """Парсит Trados и Smartcat XML отчёты и возвращает агрегированные объёмы."""
    logger.info("Starting to parse %s XML reports...", len(paths))
    logger.info("Unit: %s", unit)

    results: Dict[str, Dict[str, float]] = {}
    warnings: List[str] = []
    sources_map: Dict[str, Set[str]] = defaultdict(set)
    unit_attr = unit.lower()
    successfully_processed = 0

    for i, path in enumerate(paths):
        logger.info(
            "Processing file %s/%s: %s", i + 1, len(paths), path
        )
        filename = Path(path).name

        if is_smartcat_report(path):
            file_results, file_warnings, processed, placeholder = parse_smartcat_report(
                path, unit_attr
            )
        else:
            file_results, file_warnings, processed, placeholder = _parse_trados_report(
                path, unit_attr
            )

        warnings.extend(file_warnings)

        if file_results:
            for pair_key in file_results:
                if pair_key not in results:
                    logger.info("Created new entry for pair: %s", pair_key)
                else:
                    logger.debug("Adding to existing pair: %s", pair_key)
                sources_map[pair_key].add(filename)
            _merge_pair_results(results, file_results)
        else:
            if not file_warnings:
                msg = f"{filename}: Отчёт не удалось обработать"
                logger.error(msg)
                warnings.append(msg)
            placeholder_name = placeholder or filename
            created_key = _ensure_placeholder_entry(results, placeholder_name, filename)
            logger.info("Added placeholder entry for '%s'", created_key)
            sources_map[created_key].add(filename)

        if processed:
            successfully_processed += 1

    logger.info("=== FINAL RESULTS ===")
    logger.info(
        "Successfully processed: %s/%s reports", successfully_processed, len(paths)
    )
    logger.info("Found %s unique language pairs:", len(results))

    sorted_pairs = sorted(results.items(), key=lambda x: x[0])

    for i, (pair_key, values) in enumerate(sorted_pairs, 1):
        total = sum(values.values())
        logger.info("  %s. %s: %s total words", i, pair_key, f"{total:,.0f}")
        for name, value in values.items():
            if value > 0:
                logger.info("     • %s: %s", name, f"{value:,.0f}")

    logger.info("Unique language pairs detected:")
    for pair_key in sorted(results.keys()):
        logger.info("  • %s", pair_key)

    if warnings:
        logger.warning("Warnings/Errors (%s):", len(warnings))
        for warning in warnings:
            logger.warning("  ❌ %s", warning)

    sources = {pair: sorted(names) for pair, names in sources_map.items()}

    return results, warnings, sources
