from __future__ import annotations
from typing import Dict, List, Tuple, Optional
import xml.etree.ElementTree as ET
import re
import csv
from pathlib import Path

import langcodes
import pycountry

from .language_codes import (
    determine_short_code,
    localise_territory_code,
    replace_territory_with_code,
)
from .service_config import ServiceConfig

# Фиксированный порядок строк статистики
ROW_NAMES = ServiceConfig.ROW_NAMES


# ====== Загрузка таблицы языков ======

LANGUAGE_CODE_MAP: Dict[str, str] = {}
LANGUAGE_NAME_MAP: Dict[str, str] = {}


SMARTCAT_NS = "urn:schemas-microsoft-com:office:spreadsheet"


def _load_languages_csv() -> None:
    """Загружает сопоставление кодов языков из languages/languages.csv."""
    csv_path = Path(__file__).resolve().parents[1] / "languages" / "languages.csv"
    try:
        with open(csv_path, encoding="utf-8-sig") as f:
            reader = csv.DictReader(f, delimiter=";")
            for row in reader:
                code = row.get("Код", "").strip().lower()
                lang_en = row.get("Язык (EN)", "").strip()
                country_en = row.get("Страна (EN)", "").strip()
                lang_ru = row.get("Язык (RU)", "").strip()
                country_ru = row.get("Страна (RU)", "").strip()

                if not lang_en:
                    continue

                country_code = determine_short_code(code, lang_en, country_en, country_ru)
                display_en = f"{lang_en} ({country_code})" if country_code else lang_en

                # Попытка использовать русские названия из CSV
                lang_ru_final = lang_ru if re.search("[А-Яа-я]", lang_ru) else ""
                country_ru_final = (
                    country_ru if re.search("[А-Яа-я]", country_ru) else ""
                )

                # Если перевода нет, пытаемся получить его через langcodes
                if not lang_ru_final:
                    try:
                        lang_code = code or langcodes.find(lang_en)
                        lang_ru_final = langcodes.Language.get(lang_code).language_name('ru')
                    except Exception:
                        lang_ru_final = lang_en

                if not country_ru_final and code and "-" in code:
                    try:
                        country_ru_final = langcodes.Language.get(code).territory_name('ru')
                    except Exception:
                        country_ru_final = country_en

                country_code_ru = localise_territory_code(country_code, "ru")
                if country_code_ru:
                    display_ru = f"{lang_ru_final} ({country_code_ru})"
                elif country_ru_final:
                    display_ru = f"{lang_ru_final} ({country_ru_final})"
                else:
                    display_ru = lang_ru_final

                display = display_ru or display_en
                if display:
                    display = display[0].upper() + display[1:]

                if code:
                    LANGUAGE_CODE_MAP[code] = display

                for key in filter(
                    None,
                    [
                        lang_en.lower(),
                        display_en.lower(),
                        f"{lang_en.lower()} ({country_en.lower()})"
                        if lang_en and country_en
                        else None,
                        lang_ru.lower(),
                        display_ru.lower(),
                        f"{lang_ru.lower()} ({country_ru.lower()})"
                        if lang_ru and country_ru
                        else None,
                        f"{lang_ru.lower()} ({country_ru_final.lower()})"
                        if lang_ru and country_ru_final
                        else None,
                    ],
                ):
                    LANGUAGE_NAME_MAP[key] = display
    except FileNotFoundError:
        # Файл со списком языков отсутствует — будем использовать только стандартные методы
        pass


_load_languages_csv()


def _lookup_language(value: str) -> str:
    """Возвращает название языка из CSV по коду или имени."""
    if not value:
        return ""
    norm = value.strip().lower()
    code_key = norm.replace("_", "-")
    return LANGUAGE_CODE_MAP.get(code_key) or LANGUAGE_NAME_MAP.get(norm, "")


def _norm_lang(code: str) -> str:
    if not code:
        return ""
    return code.split("-")[0].upper()


def _extract_languages_from_filename(filename: str) -> Tuple[str, str]:
    """Извлекает языки из имени файла типа 'Analyze Files en-US_ru-RU(23).xml'"""
    print(f"Extracting languages from filename: {filename}")

    # Ищем паттерн типа en-US_ru-RU, en_ru или трёхбуквенный код вроде bez-TZ
    pattern = r'([a-z]{2,3}(?:-[A-Z]{2})?)[_-]([a-z]{2,3}(?:-[A-Z]{2})?)'
    match = re.search(pattern, filename, re.IGNORECASE)

    if match:
        src = match.group(1)
        tgt = match.group(2)
        print(f"  Found language pattern: {src} -> {tgt}")

        # Расширяем коды до полных названий для важных языков
        src_expanded = _expand_language_code(src)
        tgt_expanded = _expand_language_code(tgt)

        print(f"  Expanded: {src_expanded} -> {tgt_expanded}")
        return src_expanded, tgt_expanded

    print("  No language pattern found in filename")
    return "", ""


def _expand_language_code(code: str) -> str:
    """Преобразует языковой код в человекочитаемое название (на русском)."""
    if not code:
        return ""
    normalized = code.replace('_', '-')

    csv_name = _lookup_language(normalized)
    if csv_name:
        print(f"    Expanded {code} -> {csv_name} (csv)")
        return csv_name

    try:
        result = langcodes.Language.get(normalized).display_name('ru')
        result = replace_territory_with_code(result, 'ru')
        print(f"    Expanded {code} -> {result}")
        return result
    except langcodes.LanguageTagError:
        simple_code = _norm_lang(code)
        print(f"    Normalized {code} -> {simple_code}")
        return simple_code


def _normalize_language_name(name: str) -> str:
    """Нормализует название или код языка и возвращает его на русском."""
    if not name:
        return ""

    name = name.strip()

    csv_name = _lookup_language(name)
    if csv_name:
        print(f"  -> Normalized using CSV: '{csv_name}'")
        return csv_name

    # Прямое использование кода, если он уже корректный
    try:
        result = langcodes.Language.get(name).display_name('ru')
        return replace_territory_with_code(result, 'ru')
    except langcodes.LanguageTagError:
        pass

    try:
        if '(' in name and ')' in name:
            lang_part, region_part = name.split('(', 1)
            lang_part = lang_part.strip()
            region_part = region_part.strip(') ').strip()

            # Особые варианты для китайского языка
            if region_part.lower() in {'simplified', 'traditional'}:
                code = 'zh-Hans' if region_part.lower() == 'simplified' else 'zh-Hant'
            else:
                lang = pycountry.languages.lookup(lang_part)
                country = pycountry.countries.lookup(region_part)
                code = f"{lang.alpha_2}-{country.alpha_2}"
        else:
            lang = pycountry.languages.lookup(name)
            code = getattr(lang, 'alpha_2', '') or getattr(lang, 'alpha_3', '')

        if code:
            result = langcodes.Language.get(code).display_name('ru')
            return replace_territory_with_code(result, 'ru')
    except LookupError:
        try:
            code = langcodes.find(name)
            result = langcodes.Language.get(code).display_name('ru')
            return replace_territory_with_code(result, 'ru')
        except Exception:
            return ""
    except langcodes.LanguageTagError:
        return ""

    return ""


def _extract_language_from_taskinfo(taskinfo: ET.Element) -> str:
    """Извлекает целевой язык из элемента taskInfo"""
    print("Extracting language from taskInfo...")

    lang_element = taskinfo.find('language')
    if lang_element is not None:
        lang_name = lang_element.get('name', '').strip()
        lcid = lang_element.get('lcid', '').strip()

        print(f"  Language element found: name='{lang_name}', lcid='{lcid}'")

        normalized = _normalize_language_name(lang_name)
        if normalized:
            print(f"  -> Normalized language: '{normalized}'")
            return normalized

        normalized = _normalize_language_name(lcid)
        if normalized:
            print(f"  -> Normalized language from LCID: '{normalized}'")
            return normalized

        if lang_name:
            print(f"  -> Returning raw language name: '{lang_name}'")
            return lang_name

    print("  No language found in taskInfo")
    return ""


def _parse_analyse_element(analyse: ET.Element, unit: str = "words") -> Dict[str, float]:
    """Парсит элемент <analyse> и возвращает объемы по категориям"""
    print("  Parsing analyse element...")

    values = {name: 0.0 for name in ROW_NAMES}
    unit_attr = unit.lower()

    # Новые слова (100% новый контент)
    new_elem = analyse.find('new')
    if new_elem is not None:
        new_words = float(new_elem.get(unit_attr, 0))
        values[ROW_NAMES[0]] += new_words  # "Перевод, новые слова (100%)"
        print(f"    New words: {new_words}")

    # Нечеткие совпадения разных диапазонов
    fuzzy_elements = analyse.findall('fuzzy')
    for fuzzy in fuzzy_elements:
        min_val = int(fuzzy.get('min', 0))
        max_val = int(fuzzy.get('max', 100))
        words = float(fuzzy.get(unit_attr, 0))

        print(f"    Fuzzy {min_val}-{max_val}%: {words} words")

        if words > 0:
            if max_val <= 74:
                values[ROW_NAMES[0]] += words  # Новые слова (считаем как новые)
            elif max_val <= 94:
                values[ROW_NAMES[1]] += words  # "Перевод, совпадения 75-94% (66%)"
            elif max_val <= 99:
                values[ROW_NAMES[2]] += words  # "Перевод, совпадения 95-99% (33%)"

    # Точные совпадения и повторы
    exact_elem = analyse.find('exact')
    if exact_elem is not None:
        exact_words = float(exact_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += exact_words  # "Перевод, повторы и 100% совпадения (30%)"
        print(f"    Exact matches: {exact_words}")

    repeated_elem = analyse.find('repeated')
    if repeated_elem is not None:
        repeated_words = float(repeated_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += repeated_words  # "Перевод, повторы и 100% совпадения (30%)"
        print(f"    Repeated: {repeated_words}")

    # Межфайловые повторы
    cross_repeated_elem = analyse.find('crossFileRepeated')
    if cross_repeated_elem is not None:
        cross_words = float(cross_repeated_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += cross_words  # "Перевод, повторы и 100% совпадения (30%)"
        print(f"    Cross-file repeated: {cross_words}")

    # In-context exact
    in_context_elem = analyse.find('inContextExact')
    if in_context_elem is not None:
        in_context_words = float(in_context_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += in_context_words  # "Перевод, повторы и 100% совпадения (30%)"
        print(f"    In-context exact: {in_context_words}")

    # Perfect matches
    perfect_elem = analyse.find('perfect')
    if perfect_elem is not None:
        perfect_words = float(perfect_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += perfect_words  # "Перевод, повторы и 100% совпадения (30%)"
        print(f"    Perfect matches: {perfect_words}")

    # Locked segments
    locked_elem = analyse.find('locked')
    if locked_elem is not None:
        locked_words = float(locked_elem.get(unit_attr, 0))
        values[ROW_NAMES[3]] += locked_words  # "Перевод, повторы и 100% совпадения (30%)"
        print(f"    Locked: {locked_words}")

    total_words = sum(values.values())
    print(f"    Total words processed: {total_words}")

    return values


def _is_smartcat_report(path: str) -> bool:
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            head = f.read(2048)
    except OSError:
        return False

    lower = head.lower()
    if "<workbook" not in lower:
        return False

    if "urn:schemas-microsoft-com:office:spreadsheet" in lower:
        return True

    return "statistics for project" in lower


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


def _parse_number(text: str) -> float:
    if not text:
        return 0.0

    cleaned = text.strip().replace("\xa0", "").replace(" ", "")
    if not cleaned:
        return 0.0

    if cleaned.count(",") and cleaned.count('.'):
        if cleaned.rfind('.') > cleaned.rfind(','):
            cleaned = cleaned.replace(",", "")
        else:
            cleaned = cleaned.replace('.', "")
            cleaned = cleaned.replace(',', '.')
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


def _resolve_language_display(value: str) -> str:
    value = value.strip()
    if not value:
        return ""

    display = _expand_language_code(value)
    if display:
        return display

    normalized = _normalize_language_name(value)
    if normalized:
        return normalized

    if re.fullmatch(r"[A-Za-z]{2,3}(?:-[A-Za-z]{2,3})?", value):
        return value.upper()

    return value


def _extract_smartcat_target_language(path: str, root: ET.Element) -> str:
    candidates = _smartcat_candidates_from_text(Path(path).stem)

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
        display = _resolve_language_display(candidate_norm)
        if display:
            return display

    return ""


def _find_smartcat_statistics(
    rows: List[List[str]], unit: str
) -> List[Tuple[str, float]]:
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
    if any(keyword in lower for keyword in ("machine", "mt", "итого", "total")):
        if "total" in lower or "итого" in lower:
            return None
        if "machine" in lower or "mt" in lower:
            return None

    digits = [int(val) for val in re.findall(r"\d+", lower)]

    if any(keyword in lower for keyword in ("repeat", "повтор", "context", "perfect")):
        return ROW_NAMES[3]

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


def _parse_smartcat_report(
    path: str, unit: str
) -> Tuple[Dict[str, Dict[str, float]], List[str], bool, str]:
    filename = Path(path).name
    print(f"Detected Smartcat report: {path}")

    results: Dict[str, Dict[str, float]] = {}
    warnings: List[str] = []
    processed = False
    placeholder = Path(path).stem

    try:
        tree = ET.parse(path)
    except ET.ParseError as exc:
        msg = f"{filename}: XML Parse Error - {exc}"
        print(f"ERROR: {msg}")
        warnings.append(msg)
        return results, warnings, processed, placeholder
    except Exception as exc:
        msg = f"{filename}: Unexpected error - {exc}"
        print(f"ERROR: {msg}")
        warnings.append(msg)
        return results, warnings, processed, placeholder

    root = tree.getroot()
    print(f"Root element: {root.tag}")

    target_lang = _extract_smartcat_target_language(path, root)
    if target_lang:
        placeholder = target_lang
    pair_key = target_lang or placeholder
    print(f"Smartcat target language: '{pair_key}'")

    statistics_rows: List[Tuple[str, float]] = []
    for worksheet in root.findall(f"{{{SMARTCAT_NS}}}Worksheet"):
        sheet_name = worksheet.get(f"{{{SMARTCAT_NS}}}Name", "") or "<unnamed>"
        print(f"  Inspecting worksheet: {sheet_name}")
        rows = _worksheet_to_rows(worksheet)
        statistics_rows = _find_smartcat_statistics(rows, unit)
        if statistics_rows:
            print(f"  → Statistics found in worksheet '{sheet_name}'")
            break

    if not statistics_rows:
        msg = f"{filename}: Не удалось извлечь статистику Smartcat"
        print(f"WARNING: {msg}")
        warnings.append(msg)
        return results, warnings, processed, pair_key

    values = {name: 0.0 for name in ROW_NAMES}
    for label, number in statistics_rows:
        category = _categorize_smartcat_row(label)
        if not category:
            print(f"    Skipping row '{label}'")
            continue
        print(f"    {label} -> {category}: {number}")
        values[category] += number

    total = sum(values.values())
    results[pair_key] = values

    if total > 0:
        processed = True
        print(f"✓ Smartcat report processed: {pair_key} total {total}")
    else:
        msg = f"{filename}: Статистика Smartcat не содержит слов"
        print(f"WARNING: {msg}")
        warnings.append(msg)

    return results, warnings, processed, pair_key


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

        if root.tag != 'task':
            print(f"WARNING: Expected 'task' root element, got '{root.tag}'")

        taskinfo = root.find('taskInfo')
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
                f"{filename}: Could not determine language pair (src='{src_lang}', tgt='{tgt_lang}', taskinfo='{taskinfo_lang}')"
            )
            print(f"ERROR: {warning_msg}")
            warnings.append(warning_msg)
            return results, warnings, processed, placeholder

        placeholder = determined_target_lang or pair_key

        print(
            f"Final determined pair: '{determined_source_lang}' → '{determined_target_lang}'"
        )
        print(f"Pair key: '{pair_key}'")

        file_elements = root.findall('file')
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
            file_name = file_elem.get('name', f'file_{j}')
            print(f"\n  Processing file {j + 1}/{len(file_elements)}: {file_name}")

            analyse_elem = file_elem.find('analyse')
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


def parse_reports(paths: List[str], unit: str = "Words") -> Tuple[Dict[str, Dict[str, float]], List[str]]:
    """Парсит Trados XML отчёты и возвращает агрегированные объёмы по парам языков."""
    print(f"Starting to parse {len(paths)} XML reports...")
    print(f"Unit: {unit}")

    results: Dict[str, Dict[str, float]] = {}
    warnings: List[str] = []
    unit_attr = unit.lower()
    successfully_processed = 0

    for i, path in enumerate(paths):
        print(f"\n--- Processing file {i + 1}/{len(paths)}: {path} ---")
        filename = Path(path).name

        if _is_smartcat_report(path):
            file_results, file_warnings, processed, placeholder = _parse_smartcat_report(
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
            _merge_pair_results(results, file_results)
        else:
            if not file_warnings:
                msg = f"{filename}: Отчёт не удалось обработать"
                print(f"ERROR: {msg}")
                warnings.append(msg)
            placeholder_name = placeholder or filename
            created_key = _ensure_placeholder_entry(results, placeholder_name, filename)
            print(f"  Added placeholder entry for '{created_key}'")

        if processed:
            successfully_processed += 1

    print(f"\n=== FINAL RESULTS ===")
    print(f"Successfully processed: {successfully_processed}/{len(paths)} reports")
    print(f"Found {len(results)} unique language pairs:")

    # Сортируем пары для лучшего отображения
    sorted_pairs = sorted(results.items(), key=lambda x: x[0])

    for i, (pair_key, values) in enumerate(sorted_pairs, 1):
        total = sum(values.values())
        print(f"  {i}. {pair_key}: {total:,.0f} total words")
        for name, value in values.items():
            if value > 0:
                print(f"     • {name}: {value:,.0f}")

    print(f"\nUnique language pairs detected:")
    for pair_key in sorted(results.keys()):
        print(f"  • {pair_key}")

    if warnings:
        print(f"\nWarnings/Errors ({len(warnings)}):")
        for warning in warnings:
            print(f"  ❌ {warning}")

    return results, warnings
