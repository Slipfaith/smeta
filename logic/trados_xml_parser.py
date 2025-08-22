from __future__ import annotations
from typing import Dict, List, Tuple
import xml.etree.ElementTree as ET
import re

from .service_config import ServiceConfig

# Фиксированный порядок строк статистики
ROW_NAMES = ServiceConfig.ROW_NAMES


def _norm_lang(code: str) -> str:
    if not code:
        return ""
    return code.split("-")[0].upper()


def _extract_languages_from_filename(filename: str) -> Tuple[str, str]:
    """Извлекает языки из имени файла типа 'Analyze Files en-US_ru-RU(23).xml'"""
    print(f"Extracting languages from filename: {filename}")

    # Ищем паттерн типа en-US_ru-RU или en_ru
    pattern = r'([a-z]{2}(?:-[A-Z]{2})?)[_-]([a-z]{2}(?:-[A-Z]{2})?)'
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
    """Расширяет языковые коды до полных названий с региональными вариантами"""
    if not code:
        return ""

    code_upper = code.upper()

    # Специальные случаи с региональными вариантами
    expansions = {
        'EN-US': 'English (US)',
        'EN-UK': 'English (UK)',
        'EN-GB': 'English (UK)',
        'EN-AU': 'English (Australia)',
        'EN-CA': 'English (Canada)',
        'PT-BR': 'Portuguese (Brazil)',
        'PT-PT': 'Portuguese (Portugal)',
        'ZH-CN': 'Chinese (Simplified)',
        'ZH-TW': 'Chinese (Traditional)',
        'ZH-HK': 'Chinese (Traditional)',
        'FR-CA': 'French (Canada)',
        'FR-FR': 'French (France)',
        'ES-ES': 'Spanish (Spain)',
        'ES-MX': 'Spanish (Mexico)',
        'ES-AR': 'Spanish (Argentina)',
        'DE-DE': 'German (Germany)',
        'DE-AT': 'German (Austria)',
        'DE-CH': 'German (Switzerland)',
        'AR-SA': 'Arabic (Saudi Arabia)',
        'AR-EG': 'Arabic (Egypt)',
        'AR-AE': 'Arabic (UAE)',
    }

    if code_upper in expansions:
        result = expansions[code_upper]
        print(f"    Expanded {code} -> {result}")
        return result

    # Для простых кодов возвращаем нормализованный код
    simple_code = _norm_lang(code)
    print(f"    Normalized {code} -> {simple_code}")
    return simple_code


def _extract_language_from_taskinfo(taskinfo: ET.Element) -> str:
    """Извлекает целевой язык из элемента taskInfo"""
    print("Extracting language from taskInfo...")

    lang_element = taskinfo.find('language')
    if lang_element is not None:
        # Пробуем разные атрибуты
        lang_name = lang_element.get('name', '')
        lcid = lang_element.get('lcid', '')

        print(f"  Language element found: name='{lang_name}', lcid='{lcid}'")

        if lang_name:
            # Приводим к нижнему регистру для проверок
            lang_lower = lang_name.lower()

            print(f"  Analyzing language: '{lang_name}' (lowercase: '{lang_lower}')")

            # Китайский - более гибкая проверка
            if any(word in lang_lower for word in ['chinese', 'китайский', 'zh']):
                print("  Detected Chinese language")
                if any(word in lang_lower for word in
                       ['traditional', 'taiwan', 'hong kong', 'hk', 'tw', 'традиционный']):
                    print("  -> Traditional variant")
                    return 'Chinese (Traditional)'
                elif any(word in lang_lower for word in ['simplified', 'china', 'prc', 'cn', 'упрощенный']):
                    print("  -> Simplified variant")
                    return 'Chinese (Simplified)'
                else:
                    print("  -> Generic Chinese")
                    return 'Chinese'

            # Португальский - более гибкая проверка
            elif any(word in lang_lower for word in ['portuguese', 'português', 'португальский', 'pt']):
                print("  Detected Portuguese language")
                if any(word in lang_lower for word in ['brazil', 'brasil', 'бразилия', 'br']):
                    print("  -> Brazil variant")
                    return 'Portuguese (Brazil)'
                elif any(word in lang_lower for word in ['portugal', 'португалия']):
                    print("  -> Portugal variant")
                    return 'Portuguese (Portugal)'
                else:
                    print("  -> Generic Portuguese")
                    return 'Portuguese'

            # Английский
            elif any(word in lang_lower for word in ['english', 'английский', 'en']):
                if any(word in lang_lower for word in ['united states', 'us', 'usa', 'америка']):
                    return 'English (US)'
                elif any(word in lang_lower for word in ['united kingdom', 'uk', 'britain', 'великобритания']):
                    return 'English (UK)'
                elif any(word in lang_lower for word in ['australia', 'австралия']):
                    return 'English (Australia)'
                elif any(word in lang_lower for word in ['canada', 'канада']):
                    return 'English (Canada)'
                else:
                    return 'EN'

            # Французский
            elif any(word in lang_lower for word in ['french', 'français', 'французский', 'fr']):
                if any(word in lang_lower for word in ['canada', 'канада']):
                    return 'French (Canada)'
                elif any(word in lang_lower for word in ['france', 'франция']):
                    return 'French (France)'
                else:
                    return 'FR'

            # Испанский
            elif any(word in lang_lower for word in ['spanish', 'español', 'испанский', 'es']):
                if any(word in lang_lower for word in ['mexico', 'méxico', 'мексика']):
                    return 'Spanish (Mexico)'
                elif any(word in lang_lower for word in ['spain', 'españa', 'испания']):
                    return 'Spanish (Spain)'
                elif any(word in lang_lower for word in ['argentina', 'аргентина']):
                    return 'Spanish (Argentina)'
                else:
                    return 'ES'

            # Немецкий
            elif any(word in lang_lower for word in ['german', 'deutsch', 'немецкий', 'de']):
                if any(word in lang_lower for word in ['germany', 'deutschland', 'германия']):
                    return 'German (Germany)'
                elif any(word in lang_lower for word in ['austria', 'österreich', 'австрия']):
                    return 'German (Austria)'
                elif any(word in lang_lower for word in ['switzerland', 'schweiz', 'швейцария']):
                    return 'German (Switzerland)'
                else:
                    return 'DE'

            # Арабский
            elif any(word in lang_lower for word in ['arabic', 'العربية', 'арабский', 'ar']):
                if any(word in lang_lower for word in ['saudi', 'السعودية', 'саудовская']):
                    return 'Arabic (Saudi Arabia)'
                elif any(word in lang_lower for word in ['egypt', 'مصر', 'египет']):
                    return 'Arabic (Egypt)'
                elif any(word in lang_lower for word in ['uae', 'emirates', 'الإمارات', 'оаэ']):
                    return 'Arabic (UAE)'
                else:
                    return 'Arabic'

            # Русский
            elif any(word in lang_lower for word in ['russian', 'русский', 'ru']):
                return 'RU'

            # Для всех остальных языков - пытаемся сохранить как есть или извлечь код
            else:
                print(f"  Language not in predefined list, processing as generic: '{lang_name}'")

                # Если название содержит информацию о регионе в скобках, сохраняем всё как есть
                if '(' in lang_name and ')' in lang_name:
                    print(f"  -> Keeping full name with region info: '{lang_name}'")
                    return lang_name.strip()

                # Если это название содержит региональные индикаторы, сохраняем
                regional_indicators = ['brazil', 'portugal', 'traditional', 'simplified', 'canada', 'france',
                                       'spain', 'mexico', 'germany', 'austria', 'switzerland', 'australia',
                                       'uk', 'us', 'taiwan', 'china', 'hong kong']

                if any(indicator in lang_lower for indicator in regional_indicators):
                    print(f"  -> Keeping name with regional info: '{lang_name}'")
                    return lang_name.strip()

                # Для коротких названий (вероятно коды) возвращаем как есть в верхнем регистре
                if len(lang_name.strip()) <= 3:
                    result = lang_name.strip().upper()
                    print(f"  -> Short code detected: '{result}'")
                    return result

                # Для длинных названий пытаемся извлечь код из первых букв
                clean_name = lang_name.split('(')[0].strip()
                if len(clean_name) >= 2:
                    result = clean_name[:2].upper()
                    print(f"  -> Extracting code from long name: '{result}'")
                    return result
                else:
                    result = clean_name.upper()
                    print(f"  -> Using cleaned name: '{result}'")
                    return result

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

        try:
            tree = ET.parse(path)
            root = tree.getroot()

            print(f"Root element: {root.tag}")
            print(f"Root attributes: {root.attrib}")

            if root.tag != 'task':
                print(f"WARNING: Expected 'task' root element, got '{root.tag}'")

            # Извлекаем информацию о задаче
            taskinfo = root.find('taskInfo')
            if taskinfo is None:
                warning_msg = f"{path}: No taskInfo element found"
                print(f"ERROR: {warning_msg}")
                warnings.append(warning_msg)
                continue

            print(f"TaskInfo found: {taskinfo.attrib}")

            # Пытаемся определить языки
            # 1. Из названия файла
            import os
            filename = os.path.basename(path)
            src_lang, tgt_lang = _extract_languages_from_filename(filename)

            # 2. Из taskInfo (обычно это целевой язык)
            taskinfo_lang = _extract_language_from_taskinfo(taskinfo)

            # Логика определения пары языков
            pair_key = None
            determined_source_lang = ""
            determined_target_lang = ""

            if src_lang and tgt_lang:
                determined_source_lang = src_lang
                determined_target_lang = tgt_lang
                pair_key = f"{src_lang} → {tgt_lang}"
                print(f"Language pair from filename: {pair_key}")
            elif taskinfo_lang:
                # Если из файла не удалось извлечь, предполагаем EN -> taskinfo_lang
                determined_source_lang = "EN"
                determined_target_lang = taskinfo_lang
                if taskinfo_lang != 'EN':
                    pair_key = f"EN → {taskinfo_lang}"
                else:
                    pair_key = f"EN → {taskinfo_lang}"
                print(f"Language pair from taskInfo: {pair_key}")

            if not pair_key:
                warning_msg = f"{path}: Could not determine language pair (src='{src_lang}', tgt='{tgt_lang}', taskinfo='{taskinfo_lang}')"
                print(f"ERROR: {warning_msg}")
                warnings.append(warning_msg)
                continue

            print(f"Final determined pair: '{determined_source_lang}' → '{determined_target_lang}'")
            print(f"Pair key: '{pair_key}'")

            # Ищем файлы для анализа
            file_elements = root.findall('file')
            print(f"Found {len(file_elements)} file elements")

            if not file_elements:
                warning_msg = f"{path}: No file elements found"
                print(f"ERROR: {warning_msg}")
                warnings.append(warning_msg)
                continue

            # Инициализируем аккумулятор для этой языковой пары
            if pair_key not in results:
                results[pair_key] = {name: 0.0 for name in ROW_NAMES}
                print(f"✓ Created new entry for pair: {pair_key}")
            else:
                print(f"→ Adding to existing pair: {pair_key}")

            pair_total_words = 0
            files_processed_in_pair = 0

            # Обрабатываем каждый файл
            for j, file_elem in enumerate(file_elements):
                file_name = file_elem.get('name', f'file_{j}')
                print(f"\n  Processing file {j + 1}/{len(file_elements)}: {file_name}")

                analyse_elem = file_elem.find('analyse')
                if analyse_elem is None:
                    print(f"    No analyse element in file {file_name}")
                    continue

                # Парсим данные анализа
                file_values = _parse_analyse_element(analyse_elem, unit_attr)

                # Добавляем к общему результату
                for name in ROW_NAMES:
                    old_val = results[pair_key][name]
                    add_val = file_values[name]
                    results[pair_key][name] += add_val
                    if add_val > 0:
                        print(f"    {name}: {old_val} + {add_val} = {results[pair_key][name]}")

                file_total = sum(file_values.values())
                pair_total_words += file_total
                files_processed_in_pair += 1
                print(f"    File total: {file_total} words")

            print(f"\nPair {pair_key} total: {pair_total_words} words from {files_processed_in_pair} files")

            if pair_total_words > 0:
                successfully_processed += 1
                print(f"✓ Successfully processed report {i + 1}: {pair_key}")
            else:
                warning_msg = f"{path}: No words found in any file"
                print(f"WARNING: {warning_msg}")
                warnings.append(warning_msg)

        except ET.ParseError as e:
            error_msg = f"{path}: XML Parse Error - {e}"
            print(f"ERROR: {error_msg}")
            warnings.append(error_msg)
        except Exception as e:
            error_msg = f"{path}: Unexpected error - {e}"
            print(f"ERROR: {error_msg}")
            warnings.append(error_msg)

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