# logic/excel_exporter.py
import os
import logging
import re
from typing import Dict, Any, List, Optional, Tuple, Callable
from copy import deepcopy
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, Alignment, PatternFill, Border, Protection
from openpyxl.cell import Cell
from openpyxl.utils import get_column_letter
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, TwoCellAnchor
from openpyxl.drawing.image import Image as XLImage

from .service_config import ServiceConfig
from .translation_config import tr

CURRENCY_SYMBOLS = {"RUB": "₽", "EUR": "€", "USD": "$"}

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

# блок запуска и управления проектом
PS_START_PH = "{{project_setup}}"
PS_END_PH = "{{subtotal_project_setup}}"
PS_HDR = {
    "param": "{{taskname.project_setup}}",
    "unit": "{{unit.project_setup}}",
    "qty": "{{quantity.project_setup}}",
    "rate": "{{rate.project_setup}}",
    "total": "{{total_project_setup_table}}",
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

    def __init__(
        self,
        template_path: Optional[str] = None,
        currency: str = "RUB",
        log_path: str = "excel_export.md",
        lang: str = "ru",
    ):
        self.template_path = template_path or DEFAULT_TEMPLATE_PATH
        self.currency = currency
        self.currency_symbol = CURRENCY_SYMBOLS.get(currency, "")
        self.rate_fmt = self._currency_format(3)
        self.total_fmt = self._currency_format(2)
        self.lang = lang
        self.hdr_titles = {k: tr(v, lang) for k, v in HDR_TITLES.items()}
        self.ps_hdr_titles = {k: tr(v, lang) for k, v in PS_HDR_TITLES.items()}
        self.add_hdr_titles = {k: tr(v, lang) for k, v in ADD_HDR_TITLES.items()}
        self.subtotal_title = f"{tr('Промежуточная сумма', lang)} ({currency}):"
        self.logger = logging.getLogger("ExcelExporter")
        self.logger.setLevel(logging.DEBUG)
        self.logger.propagate = False
        if not self.logger.handlers:
            handler = logging.FileHandler(log_path, mode="w", encoding="utf-8")
            formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)
        self.logger.debug(
            "Initialized ExcelExporter with template %s", self.template_path
        )

    def _currency_format(self, decimals: int) -> str:
        sym = self.currency_symbol
        if self.currency == "USD":
            return f'"{sym}"#,##0.{"0"*decimals}'
        return f'#,##0.{"0"*decimals} "{sym}"'

    def _to_number(self, value: Any) -> Any:
        if isinstance(value, str):
            try:
                return float(value.replace(",", "."))
            except ValueError:
                return value
        return value

    def _apply_rate_format(self, cell: Cell) -> None:
        """Apply correct number format for rate cells.

        If the rate has no fractional part, use two decimals instead of three.
        """
        val = cell.value
        num: Optional[float] = None
        if isinstance(val, str):
            if val.startswith("="):
                m = re.match(r"=([A-Z]+\d+)\*([0-9.,]+)$", val)
                if m:
                    ref, mult = m.groups()
                    try:
                        base = cell.parent[ref].value
                        if isinstance(base, str):
                            base = float(base.replace(",", "."))
                        num = float(base) * float(mult.replace(",", "."))
                    except Exception:
                        num = None
                else:
                    num = None
            else:
                try:
                    num = float(val.replace(",", "."))
                except ValueError:
                    num = None
        elif isinstance(val, (int, float)):
            num = float(val)

        if num is not None and not (isinstance(val, str) and val.startswith("=")):
            cell.value = num
        if num is not None:
            if num.is_integer():
                cell.number_format = self.total_fmt
                return
            scaled = round(num * 1000)
            if scaled % 10 == 0:
                cell.number_format = self.total_fmt
                return
        cell.number_format = self.rate_fmt

    def _set_print_area(self, ws: Worksheet) -> None:
        last_col = get_column_letter(ws.max_column)
        ws.print_area = f"A1:{last_col}{ws.max_row}"

    def _fit_sheet_to_page(self, ws: Worksheet) -> None:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToHeight = 1
        ws.page_setup.fitToWidth = 1
        self._set_print_area(ws)

    def _fix_trailing_zeroes(self, wb: Workbook) -> None:
        """Заменяет три нуля после разделителя на два нуля во всех ячейках."""

        pattern = re.compile(r"(\d+[,.]\d*?)000\b")
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    value = cell.value
                    if isinstance(value, (int, float)):
                        try:
                            num = float(value)
                        except Exception:
                            continue
                        scaled = round(num * 1000)
                        if scaled % 10 == 0:
                            fmt = cell.number_format or ""
                            if "000" in fmt:
                                cell.number_format = fmt.replace("000", "00", 1)
                            else:
                                cell.number_format = self.total_fmt
                    elif isinstance(value, str):
                        new_val = pattern.sub(r"\1" + "00", value)
                        if new_val != value:
                            cell.value = new_val

    # ----------------------------- ПУБЛИЧНЫЙ АПИ -----------------------------

    def export_to_excel(
        self,
        project_data: Dict[str, Any],
        output_path: str,
        fit_to_page: bool = False,
        progress_callback: Optional[Callable[[int, str], None]] = None,
    ) -> bool:
        try:
            self.logger.info("Starting export to %s", output_path)
            if not os.path.exists(self.template_path):
                raise FileNotFoundError(f"Шаблон не найден: {self.template_path}")

            pairs_count = len(project_data.get("language_pairs", []))
            add_count = len(project_data.get("additional_services", []))
            total_steps = 7 + pairs_count + add_count
            progress = 0

            def step(message: str = "") -> None:
                nonlocal progress
                progress += 1
                if progress_callback:
                    percent = int(progress / total_steps * 100)
                    progress_callback(percent, message)

            if progress_callback:
                progress_callback(0, "Загрузка шаблона")

            wb = load_workbook(self.template_path)
            step("Шаблон загружен")

            quotation_ws = (
                wb["Quotation"] if "Quotation" in wb.sheetnames else wb.active
            )

            subtot_cells: List[str] = []
            current_row = 13

            last_row, ps_cell = self._render_project_setup_table(
                quotation_ws, project_data, current_row
            )
            if ps_cell:
                subtot_cells.append(ps_cell)
            step("Настройка проекта")
            current_row = last_row + 1

            last_row, tr_cells = self._render_translation_blocks(
                quotation_ws,
                project_data,
                current_row,
                progress_callback=lambda name: step(f"Перевод {name}"),
            )
            subtot_cells += tr_cells
            current_row = last_row + 1

            last_row, add_cells = self._render_additional_services_tables(
                quotation_ws,
                project_data,
                current_row,
                progress_callback=lambda name: step(f"{name}"),
            )
            subtot_cells += add_cells

            self.logger.debug("Subtotal cells collected: %s", subtot_cells)

            for name in (
                "ProjectSetup",
                "Setupfee",
                "Languages",
                "AdditionalServices",
                "Addservice",
            ):
                if name in wb.sheetnames:
                    del wb[name]

            step("Удаление шаблонных листов")

            self._fill_text_placeholders(
                quotation_ws, project_data, subtot_cells, start_row=1, wb=wb
            )
            step("Заполнение плейсхолдеров")

            if "Vat" in wb.sheetnames:
                del wb["Vat"]

            if fit_to_page:
                self._fit_sheet_to_page(quotation_ws)
                step("Подгонка страницы")
            else:
                self._set_print_area(quotation_ws)
                step("Установка области печати")

            self._fix_trailing_zeroes(wb)

            self.logger.info("Saving workbook to %s", output_path)
            wb.save(output_path)
            step("Сохранение файла")

            # openpyxl may drop images embedded in the template.  After saving the
            # workbook we re-open it via the COM Excel API and copy pictures from
            # the original template so that logos are preserved.
            self._restore_images_via_com(output_path)
            step("Восстановление изображений")

            if progress_callback:
                progress_callback(100, "Готово")

            self.logger.info("Export completed successfully")
            return True
        except Exception as e:
            self.logger.exception("Export failed")
            print(f"[ExcelExporter] Ошибка экспорта: {e}")
            return False

    # ----------------------------- ПОИСК/КОПИРОВАНИЕ -----------------------------

    def _find_first(
        self, ws: Worksheet, token: str, row_from: int = 1
    ) -> Optional[Tuple[int, int]]:
        self.logger.debug("Searching for '%s' starting from row %d", token, row_from)
        for r in range(row_from, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.strip() == token:
                    self.logger.debug("  found '%s' at row %d col %d", token, r, c)
                    return r, c
        self.logger.debug("  token '%s' not found", token)
        return None

    def _find_below(
        self, ws: Worksheet, start_row: int, token: str
    ) -> Optional[Tuple[int, int]]:
        self.logger.debug("Searching below row %d for '%s'", start_row, token)
        for r in range(start_row + 1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.strip() == token:
                    self.logger.debug("  found '%s' at row %d col %d", token, r, c)
                    return r, c
        self.logger.debug("  token '%s' not found below row %d", token, start_row)
        return None

    def _copy_style(self, s: Cell, d: Cell) -> None:
        d.value = s.value
        if s.font:
            d.font = Font(
                name=s.font.name,
                size=s.font.size,
                bold=s.font.bold,
                italic=s.font.italic,
                vertAlign=s.font.vertAlign,
                underline=s.font.underline,
                strike=s.font.strike,
                color=s.font.color,
            )
        if s.alignment:
            d.alignment = Alignment(
                horizontal=s.alignment.horizontal,
                vertical=s.alignment.vertical,
                text_rotation=s.alignment.text_rotation,
                wrap_text=s.alignment.wrap_text,
                shrink_to_fit=s.alignment.shrink_to_fit,
                indent=s.alignment.indent,
            )
        if s.fill:
            d.fill = PatternFill(
                fill_type=s.fill.fill_type,
                start_color=s.fill.start_color,
                end_color=s.fill.end_color,
            )
        if s.border:
            d.border = Border(
                left=s.border.left,
                right=s.border.right,
                top=s.border.top,
                bottom=s.border.bottom,
                diagonal=s.border.diagonal,
                diagonalUp=s.border.diagonalUp,
                diagonalDown=s.border.diagonalDown,
                outline=s.border.outline,
                vertical=s.border.vertical,
                horizontal=s.border.horizontal,
            )
        if s.has_style and s.number_format:
            d.number_format = s.number_format
        if s.protection:
            d.protection = Protection(
                locked=s.protection.locked, hidden=s.protection.hidden
            )

    def _copy_block(
        self,
        dst_ws: Worksheet,
        src_ws: Worksheet,
        src_start: int,
        src_end: int,
        dst_start: int,
    ) -> None:
        """Копирует строки [src_start..src_end] со страницы src_ws на позицию dst_start в dst_ws
        со стилями и слияниями. Используется, чтобы копировать неизменённый шаблонный
        блок даже после того, как первый блок уже заполнен данными."""
        height = src_end - src_start + 1
        self.logger.debug(
            "- Copying rows %d-%d from template to position %d-%d",
            src_start,
            src_end,
            dst_start,
            dst_start + height - 1,
        )
        for i in range(height):
            sr = src_start + i
            dr = dst_start + i
            self.logger.debug("  copying row %d -> %d", sr, dr)
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
                if not (
                    m.max_row < sr or m.min_row > er or m.max_col < sc or m.min_col > ec
                ):
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

        # Копирование изображений, закреплённых в указанном диапазоне
        delta = dst_start - src_start
        for img in getattr(src_ws, "_images", []):
            anchor = getattr(img, "anchor", None)
            if anchor is None:
                continue
            if isinstance(anchor, TwoCellAnchor):
                start_row = anchor.from_.row + 1
                end_row = anchor.to.row + 1
                if src_start <= start_row and end_row <= src_end:
                    new_anchor = deepcopy(anchor)
                    new_anchor.from_.row += delta
                    new_anchor.to.row += delta
                    new_img = XLImage(img._data())
                    new_img.anchor = new_anchor
                    dst_ws.add_image(new_img)
            elif isinstance(anchor, OneCellAnchor):
                row = anchor._from.row + 1
                if src_start <= row <= src_end:
                    new_anchor = deepcopy(anchor)
                    new_anchor._from.row += delta
                    new_img = XLImage(img._data())
                    new_img.anchor = new_anchor
                    dst_ws.add_image(new_img)

    def _restore_images_via_com(self, output_path: str) -> None:
        """Re-open the saved workbook and copy images from the template via COM.

        This step is required on Windows because openpyxl does not preserve
        embedded pictures (such as company logos) when saving the workbook.  The
        method uses ``win32com.client`` to copy every picture from the template's
        ``Quotation`` sheet to the corresponding sheet in the generated file,
        keeping their positions and sizes intact.
        """
        excel = tpl_wb = out_wb = None
        orig_decimal = orig_thousands = orig_use_sys = None
        custom_sep = None
        try:
            import win32com.client  # type: ignore

            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            lang_lc = self.lang.lower()
            if lang_lc.startswith("en"):
                custom_sep = (".", ",")
            elif lang_lc.startswith("ru"):
                custom_sep = (",", " ")

            if custom_sep is not None:
                try:
                    orig_decimal = excel.DecimalSeparator
                    orig_thousands = excel.ThousandsSeparator
                    orig_use_sys = excel.UseSystemSeparators
                    excel.DecimalSeparator, excel.ThousandsSeparator = custom_sep
                    excel.UseSystemSeparators = False
                except Exception:
                    pass

            tpl_path = os.path.abspath(self.template_path)
            out_path = os.path.abspath(output_path)

            tpl_wb = excel.Workbooks.Open(tpl_path)
            out_wb = excel.Workbooks.Open(out_path)
            tpl_ws = tpl_wb.Worksheets("Quotation")
            out_ws = out_wb.Worksheets("Quotation")

            # remove any existing shapes to avoid duplicates
            while out_ws.Shapes.Count > 0:
                out_ws.Shapes(1).Delete()

            for shape in tpl_ws.Shapes:
                # 13 corresponds to msoPicture
                if int(getattr(shape, "Type", 0)) == 13:
                    shape.Copy()
                    out_ws.Paste()
                    pasted = out_ws.Shapes(out_ws.Shapes.Count)
                    pasted.Left = shape.Left
                    pasted.Top = shape.Top
                    pasted.Width = shape.Width
                    pasted.Height = shape.Height

            out_wb.Save()
        except Exception as e:  # pragma: no cover - Windows only
            self.logger.debug("COM image restoration skipped: %s", e)
        finally:  # pragma: no cover - ensure COM objects are closed
            try:
                tpl_wb.Close(False)
            except Exception:
                pass
            try:
                out_wb.Close(True)
            except Exception:
                pass
            if excel is not None and custom_sep is not None and orig_use_sys is not None:
                try:
                    excel.DecimalSeparator = orig_decimal
                    excel.ThousandsSeparator = orig_thousands
                    excel.UseSystemSeparators = orig_use_sys
                except Exception:
                    pass
            try:
                excel.Quit()
            except Exception:
                pass

    def _shift_images(self, ws: Worksheet, idx: int, amount: int) -> None:
        """Сдвигает изображения, расположенные на листе, если были вставлены строки."""
        for image in getattr(ws, "_images", []):
            anchor = getattr(image, "anchor", None)
            if anchor is None:
                continue
            if isinstance(anchor, TwoCellAnchor):
                if anchor.from_.row >= idx - 1:
                    anchor.from_.row += amount
                    anchor.to.row += amount
            elif isinstance(anchor, OneCellAnchor):
                if anchor._from.row >= idx - 1:
                    anchor._from.row += amount

    def _insert_rows(self, ws: Worksheet, idx: int, amount: int) -> None:
        """Вставляет строки и корректирует позиции изображений."""
        ws.insert_rows(idx, amount)
        self._shift_images(ws, idx, amount)

    # ----------------------------- МАП КОЛОНОК -----------------------------

    def _header_map(
        self, ws: Worksheet, headers_row: int, hdr_tokens: Dict[str, str] = HDR
    ) -> Dict[str, int]:
        mapping: Dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(headers_row, c).value
            if isinstance(v, str):
                t = v.strip()
                for key, tok in hdr_tokens.items():
                    if t == tok:
                        mapping[key] = c
        if not mapping:
            mapping = {
                "param": 1,
                "type": 2,
                "unit": 3,
                "qty": 4,
                "rate": 5,
                "total": 6,
            }
            self.logger.debug(
                "Header tokens not found at row %d, using default mapping %s",
                headers_row,
                mapping,
            )
        else:
            self.logger.debug("Header map at row %d: %s", headers_row, mapping)
        return mapping

    # ----------------------------- ОСНОВНОЙ РЕНДЕР -----------------------------

    def _render_translation_blocks(
        self,
        ws: Worksheet,
        project_data: Dict[str, Any],
        start_row: int,
        progress_callback: Optional[Callable[[str], None]] = None,
    ) -> Tuple[int, List[str]]:
        pairs: List[Dict[str, Any]] = project_data.get("language_pairs", [])
        if not pairs:
            return start_row - 1, []
        self.logger.debug("Rendering %d translation pair(s)", len(pairs))

        pairs = sorted(
            pairs,
            key=lambda p: (
                p.get("pair_name", "").split(" - ")[1]
                if " - " in p.get("pair_name", "")
                else p.get("pair_name", "")
            ),
        )

        template_ws = (
            ws.parent["Languages"] if "Languages" in ws.parent.sheetnames else ws
        )
        tpl_start = self._find_first(template_ws, START_PH)
        if not tpl_start:
            raise RuntimeError("В шаблоне не найден {{translation_table}}")
        tpl_start_row, _ = tpl_start
        tpl_end = self._find_below(template_ws, tpl_start_row, END_PH)
        if not tpl_end:
            raise RuntimeError(
                "В шаблоне не найден {{subtotal_translation_table}} ниже {{translation_table}}"
            )
        tpl_end_row, _ = tpl_end

        template_height = tpl_end_row - tpl_start_row + 1
        headers_rel = 1
        first_data_rel = 2
        subtotal_rel = template_height - 1

        subtot_cells: List[str] = []
        current_row = start_row

        for pair in pairs:
            self.logger.debug("Rendering translation block '%s'", pair.get("pair_name"))
            self._insert_rows(ws, current_row, template_height)
            self._copy_block(ws, template_ws, tpl_start_row, tpl_end_row, current_row)

            block_top = current_row
            t_headers_row = block_top + headers_rel
            t_first_data = block_top + first_data_rel
            t_subtotal_row = block_top + subtotal_rel

            # Заголовок блока (заменяем плейсхолдер)
            for c in range(1, ws.max_column + 1):
                if ws.cell(block_top, c).value == START_PH:
                    ws.cell(
                        block_top,
                        c,
                        pair.get("pair_name") or pair.get("header_title") or "",
                    )
                    break

            hmap = self._header_map(ws, t_headers_row)
            self.logger.debug(
                "Header map for translation block '%s': %s",
                pair.get("pair_name"),
                hmap,
            )
            # Заголовки колонок (заменяем плейсхолдеры на русские названия)
            for c in range(1, ws.max_column + 1):
                v = ws.cell(t_headers_row, c).value
                if isinstance(v, str) and v.strip() in self.hdr_titles:
                    ws.cell(t_headers_row, c, self.hdr_titles[v.strip()])

            first_col = min(hmap.values())
            last_col = max(hmap.values())

            # Подготовка данных: фиксированные 4 строки статистики
            translation_rows = (pair.get("services") or {}).get("translation", [])
            data_map: Dict[str, Dict[str, Any]] = {}
            for row in translation_rows:
                key = row.get("key") or row.get("name")
                if not key:
                    continue
                data_map[key] = row
                lname = str(key).lower()
                if "new" in lname or "нов" in lname:
                    data_map.setdefault("new", row)
                elif "95" in lname:
                    data_map.setdefault("fuzzy_95_99", row)
                elif "75" in lname:
                    data_map.setdefault("fuzzy_75_94", row)
                elif "100" in lname:
                    data_map.setdefault("reps_100_30", row)
                self.logger.debug(
                    "  mapped '%s' -> %s",
                    key,
                    {k: v for k, v in row.items() if k in ("volume", "rate")},
                )

            items: List[Dict[str, Any]] = []
            for cfg in ServiceConfig.TRANSLATION_ROWS:
                key = cfg.get("key")
                src_row = data_map.get(key, {})
                items.append(
                    {
                        "parameter": tr(cfg.get("name"), self.lang),
                        "volume": self._to_number(src_row.get("volume") or 0),
                        "rate": self._to_number(src_row.get("rate") or 0),
                        "multiplier": cfg.get("multiplier"),
                        "is_base": bool(cfg.get("is_base")),
                    }
                )
            self.logger.debug(
                "Items prepared for '%s': %s", pair.get("pair_name"), items
            )

            # Подгонка количества строк данных
            cur_cap = max(0, t_subtotal_row - t_first_data)
            need = len(items)

            if need > cur_cap:
                add = need - cur_cap
                self._insert_rows(ws, t_subtotal_row, add)
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
                self.logger.debug(
                    "  inserted %d row(s) at %d", add, t_subtotal_row - add
                )
            elif need < cur_cap:
                delete_count = cur_cap - need
                for _ in range(delete_count):
                    ws.delete_rows(t_subtotal_row - 1)
                    t_subtotal_row -= 1
                if delete_count:
                    self.logger.debug(
                        "  deleted %d extra row(s) before %d",
                        delete_count,
                        t_subtotal_row,
                    )

            # Очистка всех строк данных
            for rr in range(t_first_data, t_subtotal_row):
                for c in range(first_col, last_col + 1):
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
                self.logger.debug(
                    "Writing translation row %d: parameter=%s volume=%s rate=%s",
                    r,
                    it.get("parameter"),
                    it.get("volume"),
                    it.get("rate"),
                )
                ws.cell(r, col_param, it["parameter"])
                ws.cell(r, col_type, "")
                unit_cell = ws.cell(r, col_unit, tr("Слово", self.lang))
                unit_cell.alignment = Alignment(horizontal="center", vertical="center")
                qty_cell = ws.cell(r, col_qty, self._to_number(it["volume"]))
                qty_cell.alignment = Alignment(horizontal="right", vertical="top")
                qty_cell.number_format = "General"
                self.logger.debug(
                    "  cells: %s%d='%s', %s%d='%s', %s%d=%s",
                    get_column_letter(col_param),
                    r,
                    it["parameter"],
                    get_column_letter(col_unit),
                    r,
                    tr("Слово", self.lang),
                    get_column_letter(col_qty),
                    r,
                    it["volume"],
                )
                r += 1

            base_idx = next(
                (idx for idx, it in enumerate(items) if it.get("is_base")), None
            )
            base_rate_cell = None
            if base_idx is not None and base_idx < len(row_numbers):
                base_rate_cell = f"{get_column_letter(col_rate)}{row_numbers[base_idx]}"

            qtyL = get_column_letter(col_qty)
            rateL = get_column_letter(col_rate)
            for idx, it in enumerate(items):
                rr = row_numbers[idx]
                if it.get("is_base"):
                    cell = ws.cell(rr, col_rate, self._to_number(it["rate"]))
                    self.logger.debug(
                        "  rate base %s%d=%s",
                        get_column_letter(col_rate),
                        rr,
                        it["rate"],
                    )
                elif base_rate_cell and it.get("multiplier") is not None:
                    multiplier = self._to_number(it["multiplier"])
                    cell = ws.cell(
                        rr, col_rate, f"={base_rate_cell}*{multiplier}"
                    )
                    self.logger.debug(
                        "  rate formula %s%d=%s*%s",
                        get_column_letter(col_rate),
                        rr,
                        base_rate_cell,
                        multiplier,
                    )
                else:
                    cell = ws.cell(rr, col_rate, self._to_number(it["rate"]))
                    self.logger.debug(
                        "  rate %s%d=%s", get_column_letter(col_rate), rr, it["rate"]
                    )
                self._apply_rate_format(cell)
                total_cell = ws.cell(rr, col_total, f"={qtyL}{rr}*{rateL}{rr}")
                total_cell.number_format = self.total_fmt
                self.logger.debug(
                    "  total formula %s%d=%s%d*%s%d",
                    get_column_letter(col_total),
                    rr,
                    qtyL,
                    rr,
                    rateL,
                    rr,
                )

            # Субтотал (заменяем плейсхолдер и ставим формулу)
            for c in range(1, ws.max_column + 1):
                if ws.cell(t_subtotal_row, c).value == END_PH:
                    ws.cell(t_subtotal_row, c, self.subtotal_title)
                    break

            totalL = get_column_letter(col_total)
            if r == t_first_data:
                subtotal_cell = ws.cell(t_subtotal_row, col_total, 0)
            else:
                subtotal_cell = ws.cell(
                    t_subtotal_row,
                    col_total,
                    f"=SUM({totalL}{t_first_data}:{totalL}{r - 1})",
                )
            subtotal_cell.number_format = self.total_fmt

            subtot_cells.append(f"{totalL}{t_subtotal_row}")
            self.logger.debug(
                "Subtotal for '%s' stored in %s",
                pair.get("pair_name"),
                f"{totalL}{t_subtotal_row}",
            )

            # Переходим к следующему блоку
            current_row = t_subtotal_row + 1
            if progress_callback:
                progress_callback(pair.get("pair_name", ""))
        # Если на листе остался оригинальный шаблонный блок с плейсхолдерами,
        # удаляем его, чтобы в результате не было дублей таблиц.
        extra_start = self._find_first(ws, START_PH, start_row)
        if extra_start:
            extra_end = self._find_below(ws, extra_start[0], END_PH)
            if extra_end:
                ws.delete_rows(extra_start[0], extra_end[0] - extra_start[0] + 1)

        return current_row - 1, subtot_cells

    def _render_project_setup_table(
        self, ws: Worksheet, project_data: Dict[str, Any], start_row: int
    ) -> Tuple[int, Optional[str]]:
        items: List[Dict[str, Any]] = project_data.get("project_setup", [])
        if not items:
            return start_row - 1, None

        template_ws = (
            ws.parent["Setupfee"]
            if "Setupfee" in ws.parent.sheetnames
            else (
                ws.parent["ProjectSetup"]
                if "ProjectSetup" in ws.parent.sheetnames
                else ws
            )
        )
        tpl_start = self._find_first(template_ws, PS_START_PH)
        if not tpl_start:
            return start_row - 1, None
        tpl_start_row, _ = tpl_start
        tpl_end = self._find_below(template_ws, tpl_start_row, PS_END_PH)
        if not tpl_end:
            raise RuntimeError(
                "В шаблоне не найден {{subtotal_project_setup}} ниже {{project_setup}}"
            )
        tpl_end_row, _ = tpl_end

        template_height = tpl_end_row - tpl_start_row + 1
        headers_rel = 1
        first_data_rel = 2
        subtotal_rel = template_height - 1

        self._insert_rows(ws, start_row, template_height)
        self._copy_block(ws, template_ws, tpl_start_row, tpl_end_row, start_row)
        # После вставки проверяем, не остался ли на листе исходный шаблонный
        # блок с плейсхолдерами, и при обнаружении удаляем его.
        ps_tail = self._find_first(ws, PS_START_PH, start_row + template_height)
        if ps_tail:
            ps_tail_end = self._find_below(ws, ps_tail[0], PS_END_PH)
            if ps_tail_end:
                ws.delete_rows(ps_tail[0], ps_tail_end[0] - ps_tail[0] + 1)

        block_top = start_row
        t_headers_row = block_top + headers_rel
        t_first_data = block_top + first_data_rel
        t_subtotal_row = block_top + subtotal_rel

        # Заголовок блока
        for c in range(1, ws.max_column + 1):
            if ws.cell(block_top, c).value == PS_START_PH:
                ws.cell(
                    block_top,
                    c,
                    tr("Запуск и управление проектом", self.lang),
                )
                break

        hmap = self._header_map(ws, t_headers_row, PS_HDR)
        self.logger.debug("Header map for project setup: %s", hmap)
        # Заголовки колонок
        for c in range(1, ws.max_column + 1):
            v = ws.cell(t_headers_row, c).value
            if isinstance(v, str) and v.strip() in self.ps_hdr_titles:
                ws.cell(t_headers_row, c, self.ps_hdr_titles[v.strip()])

        need = len(items)
        cur_cap = max(0, t_subtotal_row - t_first_data)
        if need > cur_cap:
            add = need - cur_cap
            self._insert_rows(ws, t_subtotal_row, add)
            tpl_row = (
                t_first_data if t_first_data < t_subtotal_row else t_subtotal_row - 1
            )
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
        ws.cell(block_top, first_col, tr("Запуск и управление проектом", self.lang))

        self.logger.debug("Rendering project setup table with %d items", len(items))
        r = t_first_data
        for it in items:
            self.logger.debug(
                "Project setup row %d: parameter=%s volume=%s rate=%s",
                r,
                it.get("parameter"),
                it.get("volume"),
                it.get("rate"),
            )
            ws.cell(r, col_param, it.get("parameter", ""))
            unit_cell = ws.cell(r, col_unit, tr("час", self.lang))
            unit_cell.alignment = Alignment(horizontal="center", vertical="center")
            qty_cell = ws.cell(r, col_qty, self._to_number(it.get("volume", 0)))
            qty_cell.alignment = Alignment(horizontal="right", vertical="top")
            qty_cell.number_format = "General"
            rate_cell = ws.cell(r, col_rate, self._to_number(it.get("rate", 0)))
            self._apply_rate_format(rate_cell)
            qtyL = get_column_letter(col_qty)
            rateL = get_column_letter(col_rate)
            total_cell = ws.cell(r, col_total, f"={qtyL}{r}*{rateL}{r}")
            total_cell.number_format = self.total_fmt
            r += 1

        for c in range(1, ws.max_column + 1):
            if ws.cell(t_subtotal_row, c).value == PS_END_PH:
                ws.cell(t_subtotal_row, c, self.subtotal_title)
                break

        totalL = get_column_letter(col_total)
        if r == t_first_data:
            subtotal_cell = ws.cell(t_subtotal_row, col_total, 0)
        else:
            subtotal_cell = ws.cell(
                t_subtotal_row,
                col_total,
                f"=SUM({totalL}{t_first_data}:{totalL}{r - 1})",
            )
        subtotal_cell.number_format = self.total_fmt
        self.logger.debug(
            "Project setup subtotal stored in %s", f"{totalL}{t_subtotal_row}"
        )

        return t_subtotal_row, f"{totalL}{t_subtotal_row}"

    def _render_additional_services_tables(
        self,
        ws: Worksheet,
        project_data: Dict[str, Any],
        start_row: int,
        progress_callback: Optional[Callable[[str], None]] = None,
    ) -> Tuple[int, List[str]]:
        blocks: List[Dict[str, Any]] = project_data.get("additional_services") or []
        if not blocks:
            return start_row - 1, []
        self.logger.debug("Rendering %d additional services block(s)", len(blocks))

        template_ws = (
            ws.parent["Addservice"]
            if "Addservice" in ws.parent.sheetnames
            else (
                ws.parent["AdditionalServices"]
                if "AdditionalServices" in ws.parent.sheetnames
                else ws
            )
        )
        tpl_start = self._find_first(template_ws, ADD_START_PH)
        if not tpl_start:
            return start_row - 1, []
        tpl_start_row, _ = tpl_start
        tpl_end = self._find_below(template_ws, tpl_start_row, ADD_END_PH)
        if not tpl_end:
            return start_row - 1, []
        tpl_end_row, _ = tpl_end
        template_height = tpl_end_row - tpl_start_row + 1

        subtot_cells: List[str] = []
        current_row = start_row

        for block in blocks:
            self._insert_rows(ws, current_row, template_height)
            self._copy_block(ws, template_ws, tpl_start_row, tpl_end_row, current_row)

            headers_row = current_row + 1
            first_data_row = current_row + 2
            subtotal_row = current_row + template_height - 1

            hmap = self._header_map(ws, headers_row, ADD_HDR)
            first_col = min(hmap.values())
            last_col = max(hmap.values())

            for c in range(first_col, last_col + 1):
                letter = get_column_letter(c)
                tpl_width = template_ws.column_dimensions[letter].width
                if tpl_width is not None:
                    ws.column_dimensions[letter].width = tpl_width

            header_title = block.get("header_title", "Дополнительные услуги")
            self.logger.debug(
                "Rendering additional services block '%s' with %d rows",
                header_title,
                len(block.get("rows", [])),
            )
            header_ref = f"{get_column_letter(first_col)}{current_row}:{get_column_letter(last_col)}{current_row}"
            try:
                ws.unmerge_cells(header_ref)
            except Exception:
                pass
            ws.merge_cells(header_ref)
            ws.cell(current_row, first_col, tr(header_title, self.lang))

            for c in range(first_col, last_col + 1):
                v = ws.cell(headers_row, c).value
                if isinstance(v, str) and v.strip() in self.add_hdr_titles:
                    ws.cell(headers_row, c, self.add_hdr_titles[v.strip()])

            items: List[Dict[str, Any]] = block.get("rows", [])

            need = len(items)
            cur_cap = max(0, subtotal_row - first_data_row)
            if need > cur_cap:
                add = need - cur_cap
                self._insert_rows(ws, subtotal_row, add)
                tpl_row = first_data_row if cur_cap > 0 else subtotal_row - 1
                merges_to_copy = [
                    m
                    for m in ws.merged_cells.ranges
                    if m.min_row == m.max_row == tpl_row
                    and first_col <= m.min_col
                    and m.max_col <= last_col
                ]
                for k in range(add):
                    dst = subtotal_row + k
                    for c in range(first_col, last_col + 1):
                        self._copy_style(ws.cell(tpl_row, c), ws.cell(dst, c))
                    for m in merges_to_copy:
                        sc, ec = m.min_col, m.max_col
                        ref = (
                            f"{get_column_letter(sc)}{dst}:{get_column_letter(ec)}{dst}"
                        )
                        ws.merge_cells(ref)
                subtotal_row += add
            elif need < cur_cap:
                delete_count = cur_cap - need
                for _ in range(delete_count):
                    ws.delete_rows(subtotal_row - 1)
                    subtotal_row -= 1

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
                self.logger.debug(
                    "Additional service row %d: parameter=%s unit=%s volume=%s rate=%s",
                    r,
                    it.get("parameter"),
                    it.get("unit"),
                    it.get("volume"),
                    it.get("rate"),
                )
                ws.cell(r, col_param, it.get("parameter", ""))
                unit_cell = ws.cell(r, col_unit, tr(it.get("unit", ""), self.lang))
                unit_cell.alignment = Alignment(horizontal="center", vertical="center")
                qty_cell = ws.cell(r, col_qty, self._to_number(it.get("volume", 0)))
                qty_cell.alignment = Alignment(horizontal="right", vertical="top")
                qty_cell.number_format = "General"
                rate_cell = ws.cell(r, col_rate, self._to_number(it.get("rate", 0)))
                self._apply_rate_format(rate_cell)
                qtyL = get_column_letter(col_qty)
                rateL = get_column_letter(col_rate)
                total_cell = ws.cell(r, col_total, f"={qtyL}{r}*{rateL}{r}")
                total_cell.number_format = self.total_fmt
                r += 1

            for c in range(first_col, last_col + 1):
                if ws.cell(subtotal_row, c).value == ADD_END_PH:
                    ws.cell(subtotal_row, c, self.subtotal_title)
                    break

            totalL = get_column_letter(col_total)
            if r == first_data_row:
                subtotal_cell = ws.cell(subtotal_row, col_total, 0)
            else:
                subtotal_cell = ws.cell(
                    subtotal_row,
                    col_total,
                    f"=SUM({totalL}{first_data_row}:{totalL}{r - 1})",
                )
            subtotal_cell.number_format = self.total_fmt
            subtot_cells.append(f"{totalL}{subtotal_row}")
            self.logger.debug(
                "Subtotal for '%s' stored in %s",
                header_title,
                f"{totalL}{subtotal_row}",
            )

            current_row = subtotal_row + 1
            if progress_callback:
                progress_callback(header_title)
        # Если на листе остался невостребованный блок шаблона с плейсхолдерами,
        # удаляем его, чтобы таблица не дублировалась в конце.
        add_tail = self._find_first(ws, ADD_START_PH, start_row)
        if add_tail:
            add_tail_end = self._find_below(ws, add_tail[0], ADD_END_PH)
            if add_tail_end:
                ws.delete_rows(add_tail[0], add_tail_end[0] - add_tail[0] + 1)

        return current_row - 1, subtot_cells

    # ----------------------------- ТЕКСТОВЫЕ ПЛЕЙСХОЛДЕРЫ -----------------------------

    def _fill_text_placeholders(
        self,
        ws: Worksheet,
        project_data: Dict[str, Any],
        subtot_cells: List[str],
        start_row: int = 1,
        wb: Optional[Workbook] = None,
    ) -> None:
        # общий итог — именно как формула SUM по найденным субтоталам (чтобы Excel считал сам)
        total_formula = f"=SUM({','.join(subtot_cells)})" if subtot_cells else "0"
        self.logger.debug("Total formula calculated: %s", total_formula)

        # формируем строку всех языков: сорс (первый), затем уникальные таргеты
        source_lang = ""
        target_langs: List[str] = []
        for p in project_data.get("language_pairs", []):
            pair_name = p.get("pair_name", "")
            if "→" in pair_name:
                src, tgt = [s.strip() for s in pair_name.split("→", 1)]
            elif "-" in pair_name:
                src, tgt = [s.strip() for s in pair_name.split("-", 1)]
            else:
                src, tgt = "", pair_name.strip()
            if not source_lang and src:
                source_lang = src
            if tgt:
                target_langs.append(tgt)
        # удаляем дубликаты в порядке появления
        uniq_targets: List[str] = []
        for lang in target_langs:
            if lang and lang not in uniq_targets:
                uniq_targets.append(lang)
        target_langs_str = ", ".join(filter(None, [source_lang] + uniq_targets))

        currency_code = project_data.get("currency", self.currency)

        strict_map = {
            "{{project_name}}": project_data.get("project_name", ""),
            "{{client}}": project_data.get("client_name", ""),
            "{{client_name}}": project_data.get("contact_person", ""),
            "{{PM_name}}": project_data.get("pm_name", ""),
            "{{PM_email}}": project_data.get("pm_email", ""),
            "{{Entity}}": project_data.get(
                "entity_name", project_data.get("company_name", "")
            ),
            "{{Entity_address}}": project_data.get(
                "entity_address", project_data.get("email", "")
            ),
            "{{target_langs}}": target_langs_str,
            "{{currency}}": self.currency,
        }
        self.logger.debug("Filling text placeholders with map: %s", strict_map)

        total_cell_ref: Optional[str] = None

        for r in range(start_row, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str):
                    t = v.strip()
                    if t == "{{total}}":
                        total_cell = ws.cell(r, c, total_formula)
                        total_cell.number_format = self.total_fmt
                        total_cell_ref = total_cell.coordinate
                        # Replace currency code placeholder in the total row
                        for c2 in range(1, ws.max_column + 1):
                            cell2 = ws.cell(r, c2)
                            val2 = cell2.value
                            if isinstance(val2, str) and "{{$}}" in val2:
                                cell2.value = val2.replace("{{$}}", currency_code)
                    else:
                        new_v = v
                        for ph, val in strict_map.items():
                            if ph in new_v:
                                new_v = new_v.replace(ph, str(val))
                        if new_v != v:
                            ws.cell(r, c, new_v)

        if wb is not None and total_cell_ref:
            self._insert_vat_section(wb, ws, total_cell_ref, project_data)

    # ----------------------------- НДС -----------------------------

    def _insert_vat_section(
        self,
        wb,
        ws: Worksheet,
        total_cell_ref: str,
        project_data: Dict[str, Any],
    ) -> None:
        """Copy VAT table from 'Vat' sheet and insert after total row."""
        vat_rate = self._to_number(project_data.get("vat_rate", 0) or 0)
        if not isinstance(vat_rate, (int, float)) or vat_rate <= 0 or "Vat" not in wb.sheetnames:
            return

        vat_ws = wb["Vat"]
        rows = vat_ws.max_row
        cols = vat_ws.max_column
        total_row = ws[total_cell_ref].row

        self._insert_rows(ws, total_row + 1, rows)

        for r in range(rows):
            for c in range(1, cols + 1):
                src = vat_ws.cell(r + 1, c)
                dst = ws.cell(total_row + 1 + r, c)
                self._copy_style(src, dst)

        repl_map = {
            "{{%vat}}": f"{vat_rate}%",
            "{{total_vat}}": f"={total_cell_ref}*{vat_rate}/100",
            "{{total.with_vat}}": f"={total_cell_ref}*{100+vat_rate}/100",
            "{{$}}": self.currency,
        }

        for r in range(total_row + 1, total_row + 1 + rows):
            for c in range(1, cols + 1):
                cell = ws.cell(r, c)
                if not isinstance(cell.value, str):
                    continue
                new_val = cell.value
                had_total = False
                for ph, val in repl_map.items():
                    if ph in new_val:
                        if ph in ("{{total_vat}}", "{{total.with_vat}}"):
                            had_total = True
                            rep = val
                            if new_val.strip() != ph and rep.startswith("="):
                                rep = rep.lstrip("=")
                        else:
                            rep = val
                        new_val = new_val.replace(ph, rep)
                if new_val != cell.value:
                    cell.value = new_val
                if had_total:
                    cell.number_format = self.total_fmt
