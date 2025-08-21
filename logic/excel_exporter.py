# logic/excel_exporter.py
import os
from typing import Dict, Any, List, Optional, Tuple
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, Alignment, PatternFill, Border, Protection
from openpyxl.cell import Cell
from openpyxl.utils import get_column_letter

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

            # Подготовка данных
            items: List[Dict[str, Any]] = []
            for _svc, rows in (pair.get("services") or {}).items():
                for r in rows:
                    items.append({
                        "parameter": r.get("parameter", ""),
                        "volume": float(r.get("volume") or 0),
                        "rate": float(r.get("rate") or 0),
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
            for it in items:
                ws.cell(r, col_param, it["parameter"])
                ws.cell(r, col_type, "")
                ws.cell(r, col_unit, "Слово")
                ws.cell(r, col_qty, it["volume"])
                ws.cell(r, col_rate, it["rate"])
                qtyL = get_column_letter(col_qty)
                rateL = get_column_letter(col_rate)
                ws.cell(r, col_total, f"={qtyL}{r}*{rateL}{r}")
                r += 1

            # Субтотал (заменяем плейсхолдер и ставим формулу)
            for c in range(1, ws.max_column + 1):
                if ws.cell(t_subtotal_row, c).value == END_PH:
                    ws.cell(t_subtotal_row, c, SUBTOTAL_TITLE)
                    break

            totalL = get_column_letter(col_total)
            if r == t_first_data:
                ws.cell(t_subtotal_row, col_total, 0)
            else:
                ws.cell(t_subtotal_row, col_total, f"=SUM({totalL}{t_first_data}:{totalL}{r - 1})")

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

        r = t_first_data
        for it in items:
            ws.cell(r, col_param, it.get("parameter", ""))
            ws.cell(r, col_unit, "час")
            ws.cell(r, col_qty, it.get("volume", 0))
            ws.cell(r, col_rate, it.get("rate", 0))
            qtyL = get_column_letter(col_qty)
            rateL = get_column_letter(col_rate)
            ws.cell(r, col_total, f"={qtyL}{r}*{rateL}{r}")
            r += 1

        for c in range(1, ws.max_column + 1):
            if ws.cell(t_subtotal_row, c).value == PS_END_PH:
                ws.cell(t_subtotal_row, c, SUBTOTAL_TITLE)
                break

        totalL = get_column_letter(col_total)
        if r == t_first_data:
            ws.cell(t_subtotal_row, col_total, 0)
        else:
            ws.cell(t_subtotal_row, col_total, f"=SUM({totalL}{t_first_data}:{totalL}{r - 1})")

        return f"{totalL}{t_subtotal_row}"

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
