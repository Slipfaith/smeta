# logic/excel_exporter.py
import os
from typing import Dict, Any, List, Optional, Tuple
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, Alignment, PatternFill, Border, Protection
from openpyxl.cell import Cell
from openpyxl.utils import get_column_letter

from .service_config import ServiceConfig

# Шаблон теперь ищем относительно корня проекта, чтобы код работал на любой машине
DEFAULT_TEMPLATE_PATH = os.path.join(
    os.path.dirname(__file__), "..", "templates", "шаблон.modern.xlsx"
)

START_PH = "{{translation_table}}"
END_PH = "{{subtotal_translation_table}}"

HDR = {
    "param": "{{taskname}}",
    "type": "{{type}}",
    "unit": "{{unit}}",
    "qty": "{{quantity}}",
    "rate": "{{rate}}",
    "total": "{{total_translation_table}}",
}
HDR_TITLES = {
    HDR["param"]: "Название работы",
    HDR["type"]: "Тип",
    HDR["unit"]: "Ед-ца",
    HDR["qty"]: "Кол-во",
    HDR["rate"]: "Ставка",
    HDR["total"]: "Итого",
}
SUBTOTAL_TITLE = "Промежуточная сумма (Руб):"

# блок запуска и управления проектом
PS_START_PH = "{{project_setup}}"
PS_END_PH = "{{subtotal_project_setup}}"
PS_HDR = {
    "param": "{{taskname.project_setup}}",
    "unit": "{{unit.project_setup}}",
    "qty": "{{quantity.project_setup}}",
    "rate": "{{rate.project_setup}}",
    "total": "{{total_{project_setup_table}}}",
}
PS_HDR_TITLES = {
    PS_HDR["param"]: "Названия работ",
    PS_HDR["unit"]: "час",
    PS_HDR["qty"]: "Кол-во",
    PS_HDR["rate"]: "Ставка",
    PS_HDR["total"]: "Итого",
}

# блок дополнительных услуг
ADD_START_PH = "{{add.service_table}}"
ADD_END_PH = "{{subtotal.add.service}}"
ADD_HDR = {
    "param": "{{add.service.taskname}}",
    "unit": "{{add.service.unit}}",
    "qty": "{{add.service.quantity}}",
    "rate": "{{add.service.rate}}",
    "total": "{{add.service.table}}",
}
ADD_HDR_TITLES = {
    ADD_HDR["param"]: "Параметр",
    ADD_HDR["unit"]: "Ед-ца",
    ADD_HDR["qty"]: "Кол-во",
    ADD_HDR["rate"]: "Ставка",
    ADD_HDR["total"]: "Итого",
}


class ExcelExporter:
    """Экспорт проектных данных по блоку {{translation_table}} … {{subtotal_translation_table}}."""

    def __init__(self, template_path: Optional[str] = None):
        self.template_path = template_path or DEFAULT_TEMPLATE_PATH

    # ----------------------------- ПУБЛИЧНЫЙ АПИ -----------------------------

    def export_to_excel(self, project_data: Dict[str, Any], output_path: str) -> bool:
        try:
            if not os.path.exists(self.template_path):
                raise FileNotFoundError(f"Шаблон не найден: {self.template_path}")

            wb = load_workbook(self.template_path)
            ws = wb.active

            subtot_cells: List[str] = []
            ps_cell = self._render_project_setup_table(ws, project_data)
            if ps_cell:
                subtot_cells.append(ps_cell)
            subtot_cells += self._render_translation_blocks(ws, project_data)
            add_cells = self._render_additional_services_tables(ws, project_data)
            if add_cells:
                subtot_cells += add_cells
            self._fill_text_placeholders(ws, project_data, subtot_cells)

            wb.save(output_path)
            return True
        except Exception as e:
            print(f"[ExcelExporter] Ошибка экспорта: {e}")
            return False

    # ----------------------------- ПОИСК/КОПИРОВАНИЕ -----------------------------

    @staticmethod
    def _find_first(ws: Worksheet, token: str, row_from: int = 1) -> Optional[Tuple[int, int]]:
        for r in range(row_from, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.strip() == token:
                    return r, c
        return None

    @staticmethod
    def _find_below(ws: Worksheet, start_row: int, token: str) -> Optional[Tuple[int, int]]:
        for r in range(start_row + 1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.strip() == token:
                    return r, c
        return None

    def _copy_style(self, s: Cell, d: Cell) -> None:
        d.value = s.value
        if s.font:
            d.font = Font(name=s.font.name, size=s.font.size, bold=s.font.bold,
                          italic=s.font.italic, vertAlign=s.font.vertAlign,
                          underline=s.font.underline, strike=s.font.strike, color=s.font.color)
        if s.alignment:
            d.alignment = Alignment(horizontal=s.alignment.horizontal, vertical=s.alignment.vertical,
                                    text_rotation=s.alignment.text_rotation, wrap_text=s.alignment.wrap_text,
                                    shrink_to_fit=s.alignment.shrink_to_fit, indent=s.alignment.indent)
        if s.fill:
            d.fill = PatternFill(fill_type=s.fill.fill_type, start_color=s.fill.start_color, end_color=s.fill.end_color)
        if s.border:
            d.border = Border(left=s.border.left, right=s.border.right, top=s.border.top, bottom=s.border.bottom,
                              diagonal=s.border.diagonal, diagonalUp=s.border.diagonalUp,
                              diagonalDown=s.border.diagonalDown, outline=s.border.outline,
                              vertical=s.border.vertical, horizontal=s.border.horizontal)
        if s.has_style and s.number_format:
            d.number_format = s.number_format
        if s.protection:
            d.protection = Protection(locked=s.protection.locked, hidden=s.protection.hidden)

    def _copy_block(self, dst_ws: Worksheet, src_ws: Worksheet, src_start: int, src_end: int, dst_start: int) -> None:
        """Копирует строки [src_start..src_end] со страницы src_ws на позицию dst_start в dst_ws
        со стилями и слияниями. Используется, чтобы копировать неизменённый шаблонный
        блок даже после того, как первый блок уже заполнен данными."""
        height = src_end - src_start + 1
        for i in range(height):
            sr = src_start + i
            dr = dst_start + i
            try:
                dst_ws.row_dimensions[dr].height = src_ws.row_dimensions[sr].height
            except Exception:
                pass
            for c in range(1, src_ws.max_column + 1):
                self._copy_style(src_ws.cell(sr, c), dst_ws.cell(dr, c))
        merges = []
        for m in src_ws.merged_cells.ranges:
            sr, sc, er, ec = m.min_row, m.min_col, m.max_row, m.max_col
            if src_start <= sr and er <= src_end:
                delta = dst_start - src_start
                merges.append((sr + delta, sc, er + delta, ec))
        for sr, sc, er, ec in merges:
            ref = f"{get_column_letter(sc)}{sr}:{get_column_letter(ec)}{er}"
            # Убираем перекрывающиеся с уже существующими слияния, чтобы Excel
            # не ругался на повреждённые записи при открытии файла.
            overlap = []
            for m in dst_ws.merged_cells.ranges:
                if not (m.max_row < sr or m.min_row > er or m.max_col < sc or m.min_col > ec):
                    overlap.append(str(m))
            for mref in overlap:
                try:
                    dst_ws.unmerge_cells(mref)
                except Exception:
                    pass
            try:
                dst_ws.merge_cells(ref)
            except Exception:
                pass

    # ----------------------------- МАП КОЛОНОК -----------------------------

    def _header_map(self, ws: Worksheet, headers_row: int, hdr_tokens: Dict[str, str] = HDR) -> Dict[str, int]:
        mapping: Dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(headers_row, c).value
            if isinstance(v, str):
                t = v.strip()
                for key, tok in hdr_tokens.items():
                    if t == tok:
                        mapping[key] = c
        if not mapping:
            mapping = {"param": 1, "type": 2, "unit": 3, "qty": 4, "rate": 5, "total": 6}
        return mapping

    # ----------------------------- ОСНОВНОЙ РЕНДЕР -----------------------------

    def _render_translation_blocks(self, ws: Worksheet, project_data: Dict[str, Any]) -> List[str]:
        pairs: List[Dict[str, Any]] = project_data.get("language_pairs", [])
        if not pairs:
            return []

        pairs = sorted(pairs, key=lambda p: p.get("pair_name", "").split(" - ")[1] if " - " in p.get("pair_name", "") else p.get("pair_name", ""))

        start = self._find_first(ws, START_PH)
        if not start:
            raise RuntimeError("В шаблоне не найден {{translation_table}}")
        start_row, _ = start

        end = self._find_below(ws, start_row, END_PH)
        if not end:
            raise RuntimeError("В шаблоне не найден {{subtotal_translation_table}} ниже {{translation_table}}")
        end_row, _ = end

        template_height = end_row - start_row + 1
        headers_rel = 1
        first_data_rel = 2
        subtotal_rel = template_height - 1

        template_wb = load_workbook(self.template_path)
        template_ws = template_wb.active

        subtot_cells: List[str] = []
        current_row = start_row

        for i, pair in enumerate(pairs):
            # Если не первый блок - добавляем новый
            if i > 0:
                ws.insert_rows(current_row, template_height)
                self._copy_block(ws, template_ws, start_row, end_row, current_row)

            block_top = current_row
            t_headers_row = block_top + headers_rel
            t_first_data = block_top + first_data_rel
            t_subtotal_row = block_top + subtotal_rel

            # Заголовок блока (заменяем плейсхолдер)
            for c in range(1, ws.max_column + 1):
                if ws.cell(block_top, c).value == START_PH:
                    ws.cell(block_top, c, pair.get("pair_name") or pair.get("header_title") or "")
                    break

            # Заголовки колонок (заменяем плейсхолдеры на русские названия)
            for c in range(1, ws.max_column + 1):
                v = ws.cell(t_headers_row, c).value
                if isinstance(v, str) and v.strip() in HDR_TITLES:
                    ws.cell(t_headers_row, c, HDR_TITLES[v.strip()])

            hmap = self._header_map(ws, t_headers_row)

            # Подготовка данных: фиксированные 4 строки статистики
            translation_rows = (pair.get("services") or {}).get("translation", [])
            data_map = {r.get("key"): r for r in translation_rows}

            items: List[Dict[str, Any]] = []
            for cfg in ServiceConfig.TRANSLATION_ROWS:
                key = cfg.get("key")
                src_row = data_map.get(key, {})
                items.append({
                    "parameter": cfg.get("name"),
                    "volume": float(src_row.get("volume") or 0),
                    "rate": float(src_row.get("rate") or 0),
                    "multiplier": cfg.get("multiplier"),
                    "is_base": bool(cfg.get("is_base")),
                })

            # Подгонка количества строк данных
            cur_cap = max(0, t_subtotal_row - t_first_data)
            need = len(items)

            if need > cur_cap:
                add = need - cur_cap
                ws.insert_rows(t_subtotal_row, add)
                # Копируем стиль строки данных
                if t_first_data < t_subtotal_row:
                    tpl_row = t_first_data
                else:
                    tpl_row = t_subtotal_row - 1
                for k in range(add):
                    dst = t_subtotal_row + k
                    for c in range(1, ws.max_column + 1):
                        self._copy_style(ws.cell(tpl_row, c), ws.cell(dst, c))
                t_subtotal_row += add
            elif need < cur_cap:
                delete_count = cur_cap - need
                if delete_count > 0:
                    ws.delete_rows(t_first_data + need, delete_count)
                    t_subtotal_row -= delete_count

            # Очистка всех строк данных
            for rr in range(t_first_data, t_subtotal_row):
                for c in range(1, ws.max_column + 1):
                    ws.cell(rr, c).value = None

            # Заполнение данными
            col_param = hmap.get("param", 1)
            col_type = hmap.get("type", 2)
            col_unit = hmap.get("unit", 3)
            col_qty = hmap.get("qty", 4)
            col_rate = hmap.get("rate", 5)
            col_total = hmap.get("total", 6)

            r = t_first_data
            row_numbers: List[int] = []
            for it in items:
                row_numbers.append(r)
                ws.cell(r, col_param, it["parameter"])
                ws.cell(r, col_type, "")
                ws.cell(r, col_unit, "Слово")
                ws.cell(r, col_qty, it["volume"])
                r += 1

            base_idx = next((idx for idx, it in enumerate(items) if it.get("is_base")), None)
            base_rate_cell = None
            if base_idx is not None and base_idx < len(row_numbers):
                base_rate_cell = f"{get_column_letter(col_rate)}{row_numbers[base_idx]}"

            qtyL = get_column_letter(col_qty)
            rateL = get_column_letter(col_rate)
            for idx, it in enumerate(items):
                rr = row_numbers[idx]
                if it.get("is_base"):
                    cell = ws.cell(rr, col_rate, it["rate"])
                elif base_rate_cell and it.get("multiplier") is not None:
                    cell = ws.cell(rr, col_rate, f"={base_rate_cell}*{it['multiplier']}")
                else:
                    cell = ws.cell(rr, col_rate, it["rate"])
                cell.number_format = "0.###"
                total_cell = ws.cell(rr, col_total, f"={qtyL}{rr}*{rateL}{rr}")
                total_cell.number_format = "0.00"

            # Субтотал (заменяем плейсхолдер и ставим формулу)
            for c in range(1, ws.max_column + 1):
                if ws.cell(t_subtotal_row, c).value == END_PH:
                    ws.cell(t_subtotal_row, c, SUBTOTAL_TITLE)
                    break

            totalL = get_column_letter(col_total)
            if r == t_first_data:
                subtotal_cell = ws.cell(t_subtotal_row, col_total, 0)
            else:
                subtotal_cell = ws.cell(t_subtotal_row, col_total, f"=SUM({totalL}{t_first_data}:{totalL}{r - 1})")
            subtotal_cell.number_format = "0.00"

            subtot_cells.append(f"{totalL}{t_subtotal_row}")

            # Переходим к следующему блоку
            current_row = t_subtotal_row + 1

        return subtot_cells

    def _render_project_setup_table(self, ws: Worksheet, project_data: Dict[str, Any]) -> Optional[str]:
        items: List[Dict[str, Any]] = project_data.get("project_setup", [])
        if not items:
            return None

        start = self._find_first(ws, PS_START_PH)
        if not start:
            return None
        start_row, _ = start
        end = self._find_below(ws, start_row, PS_END_PH)
        if not end:
            raise RuntimeError("В шаблоне не найден {{subtotal_project_setup}} ниже {{project_setup}}")
        end_row, _ = end

        template_height = end_row - start_row + 1
        headers_rel = 1
        first_data_rel = 2
        subtotal_rel = template_height - 1

        block_top = start_row
        t_headers_row = block_top + headers_rel
        t_first_data = block_top + first_data_rel
        t_subtotal_row = block_top + subtotal_rel

        # Заголовок блока
        for c in range(1, ws.max_column + 1):
            if ws.cell(block_top, c).value == PS_START_PH:
                ws.cell(block_top, c, "Запуск и управление проектом")
                break

        # Заголовки колонок
        for c in range(1, ws.max_column + 1):
            v = ws.cell(t_headers_row, c).value
            if isinstance(v, str) and v.strip() in PS_HDR_TITLES:
                ws.cell(t_headers_row, c, PS_HDR_TITLES[v.strip()])

        hmap = self._header_map(ws, t_headers_row, PS_HDR)

        need = len(items)
        cur_cap = max(0, t_subtotal_row - t_first_data)
        if need > cur_cap:
            add = need - cur_cap
            ws.insert_rows(t_subtotal_row, add)
            tpl_row = t_first_data if t_first_data < t_subtotal_row else t_subtotal_row - 1
            for k in range(add):
                dst = t_subtotal_row + k
                for c in range(1, ws.max_column + 1):
                    self._copy_style(ws.cell(tpl_row, c), ws.cell(dst, c))
            t_subtotal_row += add
        elif need < cur_cap:
            delete_count = cur_cap - need
            if delete_count > 0:
                ws.delete_rows(t_first_data + need, delete_count)
                t_subtotal_row -= delete_count

        # очистка строк
        for rr in range(t_first_data, t_subtotal_row):
            for c in range(1, ws.max_column + 1):
                ws.cell(rr, c).value = None

        col_param = hmap.get("param", 1)
        col_unit = hmap.get("unit", 3)
        col_qty = hmap.get("qty", 4)
        col_rate = hmap.get("rate", 5)
        col_total = hmap.get("total", 6)

        # Ensure header is merged across the table width
        first_col = min(hmap.values())
        last_col = max(hmap.values())
        ref = f"{get_column_letter(first_col)}{block_top}:{get_column_letter(last_col)}{block_top}"
        try:
            ws.unmerge_cells(ref)
        except Exception:
            pass
        ws.merge_cells(ref)
        ws.cell(block_top, first_col, "Запуск и управление проектом")

        r = t_first_data
        for it in items:
            ws.cell(r, col_param, it.get("parameter", ""))
            ws.cell(r, col_unit, "час")
            ws.cell(r, col_qty, it.get("volume", 0))
            rate_cell = ws.cell(r, col_rate, it.get("rate", 0))
            rate_cell.number_format = "0.00"
            qtyL = get_column_letter(col_qty)
            rateL = get_column_letter(col_rate)
            total_cell = ws.cell(r, col_total, f"={qtyL}{r}*{rateL}{r}")
            total_cell.number_format = "0.00"
            r += 1

        for c in range(1, ws.max_column + 1):
            if ws.cell(t_subtotal_row, c).value == PS_END_PH:
                ws.cell(t_subtotal_row, c, SUBTOTAL_TITLE)
                break

        totalL = get_column_letter(col_total)
        if r == t_first_data:
            subtotal_cell = ws.cell(t_subtotal_row, col_total, 0)
        else:
            subtotal_cell = ws.cell(t_subtotal_row, col_total, f"=SUM({totalL}{t_first_data}:{totalL}{r - 1})")
        subtotal_cell.number_format = "0.00"

        return f"{totalL}{t_subtotal_row}"

    def _render_additional_services_tables(self, ws: Worksheet, project_data: Dict[str, Any]) -> List[str]:
        blocks: List[Dict[str, Any]] = project_data.get("additional_services") or []
        if not blocks:
            return []

        # исходный шаблон — чтобы копировать новые блоки
        template_wb = load_workbook(self.template_path)
        template_ws = template_wb.active
        tpl_start = self._find_first(template_ws, ADD_START_PH)
        if not tpl_start:
            return []
        tpl_start_row, _ = tpl_start
        tpl_end = self._find_below(template_ws, tpl_start_row, ADD_END_PH)
        if not tpl_end:
            return []
        tpl_end_row, _ = tpl_end
        template_height = tpl_end_row - tpl_start_row + 1

        subtot_cells: List[str] = []
        search_row = 1

        for i, block in enumerate(blocks):
            # для каждого нового блока — вставляем шаблонный блок
            if i > 0:
                ws.insert_rows(search_row, template_height)
                self._copy_block(ws, template_ws, tpl_start_row, tpl_end_row, search_row)

            start = self._find_first(ws, ADD_START_PH, search_row)
            if not start:
                break
            start_row, _ = start
            end = self._find_below(ws, start_row, ADD_END_PH)
            if not end:
                break
            subtotal_row, _ = end

            headers_row = start_row + 1
            first_data_row = start_row + 2

            hmap = self._header_map(ws, headers_row, ADD_HDR)
            first_col = min(hmap.values())
            last_col = max(hmap.values())

            # ширины колонок берём из шаблона
            for c in range(first_col, last_col + 1):
                letter = get_column_letter(c)
                tpl_width = template_ws.column_dimensions[letter].width
                if tpl_width is not None:
                    ws.column_dimensions[letter].width = tpl_width

            header_title = block.get("header_title", "Дополнительные услуги")
            header_ref = f"{get_column_letter(first_col)}{start_row}:{get_column_letter(last_col)}{start_row}"
            try:
                ws.unmerge_cells(header_ref)
            except Exception:
                pass
            ws.merge_cells(header_ref)
            ws.cell(start_row, first_col, header_title)

            # заголовки колонок
            for c in range(first_col, last_col + 1):
                v = ws.cell(headers_row, c).value
                if isinstance(v, str) and v.strip() in ADD_HDR_TITLES:
                    ws.cell(headers_row, c, ADD_HDR_TITLES[v.strip()])

            items: List[Dict[str, Any]] = block.get("rows", [])

            need = len(items)
            cur_cap = max(0, subtotal_row - first_data_row)
            if need > cur_cap:
                add = need - cur_cap
                ws.insert_rows(subtotal_row, add)
                tpl_row = first_data_row if cur_cap > 0 else subtotal_row - 1
                merges_to_copy = [
                    m for m in ws.merged_cells.ranges
                    if m.min_row == m.max_row == tpl_row
                    and first_col <= m.min_col and m.max_col <= last_col
                ]
                for k in range(add):
                    dst = subtotal_row + k
                    for c in range(first_col, last_col + 1):
                        self._copy_style(ws.cell(tpl_row, c), ws.cell(dst, c))
                    for m in merges_to_copy:
                        sc, ec = m.min_col, m.max_col
                        ref = f"{get_column_letter(sc)}{dst}:{get_column_letter(ec)}{dst}"
                        ws.merge_cells(ref)
                subtotal_row += add
            elif need < cur_cap:
                delete_count = cur_cap - need
                for _ in range(delete_count):
                    ws.delete_rows(subtotal_row - 1)
                    subtotal_row -= 1

            # очистка блока
            for r in range(first_data_row, subtotal_row):
                for c in range(first_col, last_col + 1):
                    ws.cell(r, c).value = None

            col_param = hmap["param"]
            col_unit = hmap["unit"]
            col_qty = hmap["qty"]
            col_rate = hmap["rate"]
            col_total = hmap["total"]

            r = first_data_row
            for it in items:
                ws.cell(r, col_param, it.get("parameter", ""))
                ws.cell(r, col_unit, it.get("unit", ""))
                ws.cell(r, col_qty, it.get("volume", 0))
                rate_cell = ws.cell(r, col_rate, it.get("rate", 0))
                rate_cell.number_format = "0.00"
                qtyL = get_column_letter(col_qty)
                rateL = get_column_letter(col_rate)
                total_cell = ws.cell(r, col_total, f"={qtyL}{r}*{rateL}{r}")
                total_cell.number_format = "0.00"
                r += 1

            for c in range(first_col, last_col + 1):
                if ws.cell(subtotal_row, c).value == ADD_END_PH:
                    ws.cell(subtotal_row, c, SUBTOTAL_TITLE)
                    break

            totalL = get_column_letter(col_total)
            if r == first_data_row:
                subtotal_cell = ws.cell(subtotal_row, col_total, 0)
            else:
                subtotal_cell = ws.cell(subtotal_row, col_total, f"=SUM({totalL}{first_data_row}:{totalL}{r - 1})")
            subtotal_cell.number_format = "0.00"
            subtot_cells.append(f"{totalL}{subtotal_row}")

            search_row = subtotal_row + 1

        return subtot_cells

    # ----------------------------- ТЕКСТОВЫЕ ПЛЕЙСХОЛДЕРЫ -----------------------------

    def _fill_text_placeholders(self, ws: Worksheet, project_data: Dict[str, Any], subtot_cells: List[str]) -> None:
        # общий итог — именно как формула SUM по найденным субтоталам (чтобы Excel считал сам)
        total_formula = f"=SUM({','.join(subtot_cells)})" if subtot_cells else "0"

        strict_map = {
            "{{project_name}}": project_data.get("project_name", ""),
            "{{client}}": project_data.get("client_name", ""),
            "{{Entity}}": project_data.get("entity_name", project_data.get("company_name", "")),
            "{{Entity_address}}": project_data.get("entity_address", project_data.get("email", "")),
            "{{client_name}}": project_data.get("contact_person", ""),
            "{{PM_name}}": project_data.get("pm_name", "PM"),
            "{{PM_email}}": project_data.get("pm_email", "pm@company.com"),
            "{{target_langs}}": ", ".join([p.get("pair_name", "") for p in project_data.get("language_pairs", [])]),
        }

        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str):
                    t = v.strip()
                    if t == "{{total}}":
                        ws.cell(r, c, total_formula)
                    else:
                        new_v = v
                        for ph, val in strict_map.items():
                            if ph in new_v:
                                new_v = new_v.replace(ph, str(val))
                        if new_v != v:
                            ws.cell(r, c, new_v)
