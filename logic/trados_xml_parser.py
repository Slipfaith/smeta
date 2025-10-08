from __future__ import annotations

from collections import defaultdict
from typing import Dict, List, Optional, Set, Tuple
import xml.etree.ElementTree as ET
import re
from pathlib import Path

from .service_config import ServiceConfig
from .xml_parser_common import expand_language_code, normalize_language_name
from .sc_xml_parser import is_smartcat_report, parse_smartcat_report


ROW_NAMES = ServiceConfig.ROW_NAMES


def _extract_languages_from_filename(filename: str) -> Tuple[str, str]:
    """Извлекает языки из имени файла типа 'Analyze Files en-US_ru-RU(23).xml'."""
    print(f"Extracting languages from filename: {filename}")

    # Allow matching extended BCP 47 subtags (e.g. ``es-419``) where the
    # territory part may include digits such as the UN M.49 codes that Trados
    # uses for "Latin America".  Previously the pattern only permitted
    # two-letter country codes which meant that reports like
    # ``Analyze Files en-US_es-419.xml`` failed to detect the target language
    # entirely.  Broadening the character class keeps backwards compatibility
    # while covering the new format.
    pattern = r"([a-z]{2,3}(?:-[A-Za-z0-9]{2,8})?)[_-]([a-z]{2,3}(?:-[A-Za-z0-9]{2,8})?)"
    match = re.search(pattern, filename, re.IGNORECASE)

    if match:
        src = match.group(1)
        tgt = match.group(2)
        print(f"  Found language pattern: {src} -> {tgt}")

        src_expanded = expand_language_code(src)
        tgt_expanded = expand_language_code(tgt)

        print(f"  Expanded: {src_expanded} -> {tgt_expanded}")
        return src_expanded, tgt_expanded

    print("  No language pattern found in filename")
    return "", ""


def _extract_language_from_taskinfo(taskinfo: ET.Element) -> str:
    """Извлекает целевой язык из элемента taskInfo."""
    print("Extracting language from taskInfo...")

    lang_element = taskinfo.find("language")
    if lang_element is not None:
        lang_name = lang_element.get("name", "").strip()
        lcid = lang_element.get("lcid", "").strip()

        print(f"  Language element found: name='{lang_name}', lcid='{lcid}'")

        normalized = normalize_language_name(lang_name)
        if normalized:
            print(f"  -> Normalized language: '{normalized}'")
            return normalized

        normalized = normalize_language_name(lcid)
        if normalized:
            print(f"  -> Normalized language from LCID: '{normalized}'")
            return normalized

        if lang_name:
            print(f"  -> Returning raw language name: '{lang_name}'")
            return lang_name

    print("  No language found in taskInfo")
    return ""


def _parse_analyse_element(analyse: ET.Element, unit: str = "words") -> Dict[str, float]:
    """Парсит элемент <analyse> и возвращает объемы по категориям."""
    print("  Parsing analyse element...")

    values = {name: 0.0 for name in ROW_NAMES}
    unit_attr = unit.lower()

    new_elem = analyse.find("new")
    if new_elem is not None:
        new_words = float(new_elem.get(unit_attr, 0))
        values[ROW_NAMES[0]] += new_words
        print(f"    New words: {new_words}")

    fuzzy_elements = analyse.findall("fuzzy")
    for fuzzy in fuzzy_elements:
        min_val = int(fuzzy.get("min", 0))
        max_val = int(fuzzy.get("max", 100))
        words = float(fuzzy.get(unit_attr, 0))

        print(f"    Fuzzy {min_val}-{max_val}%: {words} words")

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
        print(f"    Exact matches: {exact_words}")

    repeated_elem = analyse.find("repeated")
    if repeated_elem is not None:
        repeated_words = float(repeated_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += repeated_words
        print(f"    Repeated: {repeated_words}")

    cross_repeated_elem = analyse.find("crossFileRepeated")
    if cross_repeated_elem is not None:
        cross_words = float(cross_repeated_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += cross_words
        print(f"    Cross-file repeated: {cross_words}")

    in_context_elem = analyse.find("inContextExact")
    if in_context_elem is not None:
        in_context_words = float(in_context_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += in_context_words
        print(f"    In-context exact: {in_context_words}")

    perfect_elem = analyse.find("perfect")
    if perfect_elem is not None:
        perfect_words = float(perfect_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += perfect_words
        print(f"    Perfect matches: {perfect_words}")

    locked_elem = analyse.find("locked")
    if locked_elem is not None:
        locked_words = float(locked_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += locked_words
        print(f"    Locked: {locked_words}")

    total_words = sum(values.values())
    print(f"    Total words processed: {total_words}")

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
    print(f"Processing Trados report: {path}")

    results: Dict[str, Dict[str, float]] = {}
    warnings: List[str] = []
    processed = False
    placeholder = Path(path).stem

    try:
        tree = ET.parse(path)
        root = tree.getroot()

        print(f"Root element: {root.tag}")
        print(f"Root attributes: {root.attrib}")

        if root.tag != "task":
            print(f"WARNING: Expected 'task' root element, got '{root.tag}'")

        taskinfo = root.find("taskInfo")
        if taskinfo is None:
            warning_msg = f"{filename}: No taskInfo element found"
            print(f"ERROR: {warning_msg}")
            warnings.append(warning_msg)
            return results, warnings, processed, placeholder

        print(f"TaskInfo found: {taskinfo.attrib}")

        filename_only = Path(path).name
        src_lang, tgt_lang = _extract_languages_from_filename(filename_only)
        taskinfo_lang = _extract_language_from_taskinfo(taskinfo)

        pair_key: Optional[str] = None
        determined_source_lang = ""
        determined_target_lang = ""

        if src_lang and tgt_lang:
            determined_source_lang = src_lang
            determined_target_lang = tgt_lang
            print(f"Language pair from filename: {src_lang} → {tgt_lang}")
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
            print(f"Language pair from taskInfo: {pair_key}")

        if not pair_key:
            warning_msg = (
                f"{filename}: Could not determine language pair (src='{src_lang}', tgt='{tgt_lang}', taskinfo='{taskinfo_lang}')."
                " Please assign the correct language pair manually."
            )
            print(f"ERROR: {warning_msg}")
            warnings.append(warning_msg)
            return results, warnings, processed, placeholder

        placeholder = determined_target_lang or pair_key

        print(
            f"Final determined pair: '{determined_source_lang}' → '{determined_target_lang}'"
        )
        print(f"Pair key: '{pair_key}'")

        file_elements = root.findall("file")
        print(f"Found {len(file_elements)} file elements")

        if not file_elements:
            warning_msg = f"{filename}: No file elements found"
            print(f"ERROR: {warning_msg}")
            warnings.append(warning_msg)
            return results, warnings, processed, placeholder

        pair_values = {name: 0.0 for name in ROW_NAMES}
        pair_total_words = 0.0
        files_processed_in_pair = 0

        for j, file_elem in enumerate(file_elements):
            file_name = file_elem.get("name", f"file_{j}")
            print(f"\n  Processing file {j + 1}/{len(file_elements)}: {file_name}")

            analyse_elem = file_elem.find("analyse")
            if analyse_elem is None:
                print(f"    No analyse element in file {file_name}")
                continue

            file_values = _parse_analyse_element(analyse_elem, unit)

            for name in ROW_NAMES:
                add_val = file_values[name]
                pair_values[name] += add_val
                if add_val > 0:
                    print(f"    {name}: +{add_val} (now {pair_values[name]})")

            file_total = sum(file_values.values())
            pair_total_words += file_total
            print(f"    File total: {file_total} words")
            files_processed_in_pair += 1

        results[pair_key] = pair_values

        print(
            f"\nPair {pair_key} total: {pair_total_words} words from {files_processed_in_pair} files"
        )

        if pair_total_words > 0:
            processed = True
            print(f"✓ Successfully processed report: {pair_key}")
        else:
            warning_msg = f"{filename}: No words found in any file"
            print(f"WARNING: {warning_msg}")
            warnings.append(warning_msg)

    except ET.ParseError as exc:
        error_msg = f"{filename}: XML Parse Error - {exc}"
        print(f"ERROR: {error_msg}")
        warnings.append(error_msg)
    except Exception as exc:
        error_msg = f"{filename}: Unexpected error - {exc}"
        print(f"ERROR: {error_msg}")
        warnings.append(error_msg)

    return results, warnings, processed, placeholder


def parse_reports(
    paths: List[str], unit: str = "Words"
) -> Tuple[Dict[str, Dict[str, float]], List[str], Dict[str, List[str]]]:
    """Парсит Trados и Smartcat XML отчёты и возвращает агрегированные объёмы."""
    print(f"Starting to parse {len(paths)} XML reports...")
    print(f"Unit: {unit}")

    results: Dict[str, Dict[str, float]] = {}
    warnings: List[str] = []
    sources_map: Dict[str, Set[str]] = defaultdict(set)
    unit_attr = unit.lower()
    successfully_processed = 0

    for i, path in enumerate(paths):
        print(f"\n--- Processing file {i + 1}/{len(paths)}: {path} ---")
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
                    print(f"✓ Created new entry for pair: {pair_key}")
                else:
                    print(f"→ Adding to existing pair: {pair_key}")
                sources_map[pair_key].add(filename)
            _merge_pair_results(results, file_results)
        else:
            if not file_warnings:
                msg = f"{filename}: Отчёт не удалось обработать"
                print(f"ERROR: {msg}")
                warnings.append(msg)
            placeholder_name = placeholder or filename
            created_key = _ensure_placeholder_entry(results, placeholder_name, filename)
            print(f"  Added placeholder entry for '{created_key}'")
            sources_map[created_key].add(filename)

        if processed:
            successfully_processed += 1

    print("\n=== FINAL RESULTS ===")
    print(f"Successfully processed: {successfully_processed}/{len(paths)} reports")
    print(f"Found {len(results)} unique language pairs:")

    sorted_pairs = sorted(results.items(), key=lambda x: x[0])

    for i, (pair_key, values) in enumerate(sorted_pairs, 1):
        total = sum(values.values())
        print(f"  {i}. {pair_key}: {total:,.0f} total words")
        for name, value in values.items():
            if value > 0:
                print(f"     • {name}: {value:,.0f}")

    print("\nUnique language pairs detected:")
    for pair_key in sorted(results.keys()):
        print(f"  • {pair_key}")

    if warnings:
        print(f"\nWarnings/Errors ({len(warnings)}):")
        for warning in warnings:
            print(f"  ❌ {warning}")

    sources = {pair: sorted(names) for pair, names in sources_map.items()}

    return results, warnings, sources
