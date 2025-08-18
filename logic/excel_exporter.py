import os
from datetime import datetime
from typing import Dict, Any, List
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill

class ExcelExporter:
    """Класс для экспорта данных в Excel"""

    def __init__(self, template_path: str = None):
        self.template_path = template_path

    def export_to_excel(self, project_data: Dict[str, Any], output_path: str):
        try:
            if not self.template_path or not os.path.exists(self.template_path):
                raise Exception("Не указан путь к шаблону Excel или файл не существует")

            wb = load_workbook(self.template_path)
            ws = wb.active

            self.fill_template_data(ws, project_data)

            wb.save(output_path)
            return True
        except Exception as e:
            print(f"Ошибка при экспорте в Excel: {e}")
            return False

    def fill_template_data(self, ws, project_data: Dict[str, Any]):
        # Плейсхолдеры шапки
        for row in ws.iter_rows():
            for cell in row:
                if not cell.value:
                    continue
                val = str(cell.value)
                if "{{project_name}}" in val:
                    cell.value = val.replace("{{project_name}}", project_data.get('project_name', ''))
                elif "{{target_langs}}" in val:
                    pairs = [pair['pair_name'] for pair in project_data.get('language_pairs', [])]
                    cell.value = val.replace("{{target_langs}}", ", ".join(pairs))
                elif "{{client}}" in val:
                    cell.value = val.replace("{{client}}", project_data.get('client_name', ''))
                elif "{{Entity}}" in val:
                    cell.value = val.replace("{{Entity}}", project_data.get('client_name', ''))
                elif "{{Entity_address}}" in val:
                    cell.value = val.replace("{{Entity_address}}", project_data.get('email', ''))
                elif "{{client_name}}" in val:
                    cell.value = val.replace("{{client_name}}", project_data.get('contact_person', ''))
                elif "{{PM_name}}" in val:
                    cell.value = val.replace("{{PM_name}}", "Project Manager")
                elif "{{PM_email}}" in val:
                    cell.value = val.replace("{{PM_email}}", "pm@company.com")

        # Языки
        self.fill_language_pairs_data(ws, project_data)
        # Доп. услуги
        self.fill_additional_services_data(ws, project_data)

    def fill_language_pairs_data(self, ws, project_data: Dict[str, Any]):
        language_pairs = project_data.get('language_pairs', [])
        if not language_pairs:
            return

        table_sections = self.find_table_sections(ws)

        if len(language_pairs) > len(table_sections):
            self.duplicate_table_sections(ws, table_sections, len(language_pairs))
            table_sections = self.find_table_sections(ws)

        for i, pair_data in enumerate(language_pairs):
            if i < len(table_sections):
                self.fill_single_language_table(ws, pair_data, table_sections[i])

    def find_table_sections(self, ws):
        out = []
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                if self.is_table_header(cell):
                    info = self.get_table_boundaries(ws, r, c)
                    if info:
                        out.append(info)
        return out

    def is_table_header(self, cell):
        if not cell.value:
            return False
        val = str(cell.value).strip()
        indicators = ["Китайский", "Английский", "Немецкий", "Французский", "→", "Перевод", "Локализация", "Название работы"]
        has_fill = cell.fill and cell.fill.start_color and cell.fill.start_color.index != '00000000'
        is_bold = cell.font and cell.font.bold
        return any(i in val for i in indicators) or has_fill or is_bold

    def get_table_boundaries(self, ws, header_row, header_col):
        headers_row = None
        for r in range(header_row + 1, min(header_row + 5, ws.max_row + 1)):
            cell_a = ws.cell(r, 1)
            if cell_a.value and ("Название работы" in str(cell_a.value) or "Параметр" in str(cell_a.value)):
                headers_row = r
                break
        if not headers_row:
            return None

        last_data_row = headers_row
        for r in range(headers_row + 1, ws.max_row + 1):
            cell_a = ws.cell(r, 1)
            if (not cell_a.value or
                "Промежуточная сумма" in str(cell_a.value) or
                "КОНЕЧНАЯ СТОИМОСТЬ" in str(cell_a.value)):
                break
            if self.is_table_header(cell_a):
                break
            last_data_row = r

        return {
            'header_row': header_row,
            'column_headers_row': headers_row,
            'first_data_row': headers_row + 1,
            'last_data_row': last_data_row,
            'subtotal_row': last_data_row + 1
        }

    def duplicate_table_sections(self, ws, sections, needed_count):
        if not sections or len(sections) >= needed_count:
            return
        tmpl = sections[-1]
        height = tmpl['subtotal_row'] - tmpl['header_row'] + 2
        for i in range(len(sections), needed_count):
            new_start = tmpl['subtotal_row'] + 2 + (i - len(sections) + 1) * height
            for off in range(height):
                src_row = tmpl['header_row'] + off
                dst_row = new_start + off
                for col in range(1, 6):  # A..E
                    s = ws.cell(src_row, col)
                    t = ws.cell(dst_row, col)
                    t.value = s.value
                    if s.font:
                        t.font = Font(name=s.font.name, size=s.font.size, bold=s.font.bold, color=s.font.color)
                    if s.fill:
                        t.fill = PatternFill(fill_type=s.fill.fill_type, start_color=s.fill.start_color, end_color=s.fill.end_color)
                    if s.alignment:
                        t.alignment = Alignment(horizontal=s.alignment.horizontal, vertical=s.alignment.vertical)

    def fill_single_language_table(self, ws, pair_data, info):
        # Заголовок = header_title (если есть), иначе pair_name
        header_cell = ws.cell(info['header_row'], 1)
        header_cell.value = pair_data.get('header_title') or pair_data.get('pair_name')

        # Очистка данных
        for r in range(info['first_data_row'], info['last_data_row'] + 1):
            for c in range(1, 6):
                ws.cell(r, c).value = None

        # Заполнение
        row = info['first_data_row']
        subtotal = 0.0
        for service_name, service_data in pair_data.get('services', {}).items():
            for d in service_data:
                if d['volume'] > 0 or d['rate'] > 0:
                    if row <= info['last_data_row']:
                        ws.cell(row, 1).value = d['parameter']
                        ws.cell(row, 2).value = "Слово"
                        ws.cell(row, 3).value = d['volume']
                        ws.cell(row, 4).value = d['rate']
                        ws.cell(row, 5).value = d['total']
                        subtotal += d['total']
                        row += 1

        ws.cell(info['subtotal_row'], 5).value = f"{subtotal:.2f}р."

    # --- Доп. услуги + итог (без изменений) ---
    def fill_additional_services_data(self, ws, project_data: Dict[str, Any]):
        additional = project_data.get('additional_services', {})
        if not additional:
            return
        prep = self.find_preparation_section(ws)
        if prep:
            self.fill_preparation_section(ws, prep, additional)
        self.update_final_total(ws, project_data)

    def find_preparation_section(self, ws):
        for r in range(1, ws.max_row + 1):
            cell = ws.cell(r, 1)
            if cell.value and "Подготовка проекта" in str(cell.value):
                headers_row = None
                for rr in range(r + 1, min(r + 5, ws.max_row + 1)):
                    if ws.cell(rr, 1).value and "Название работы" in str(ws.cell(rr, 1).value):
                        headers_row = rr
                        break
                if headers_row:
                    return {'header_row': r, 'column_headers_row': headers_row, 'first_data_row': headers_row + 1, 'last_data_row': headers_row + 10}
        return None

    def fill_preparation_section(self, ws, section, additional):
        cur = section['first_data_row']
        for r in range(section['first_data_row'], section['last_data_row'] + 1):
            for c in range(1, 6):
                ws.cell(r, c).value = None
        for svc_name, svc_data in additional.items():
            for d in svc_data:
                if d['volume'] > 0 or d['rate'] > 0:
                    if cur <= section['last_data_row']:
                        ws.cell(cur, 1).value = f"{svc_name}: {d['parameter']}"
                        ws.cell(cur, 2).value = "Час"
                        ws.cell(cur, 3).value = d['volume']
                        ws.cell(cur, 4).value = d['rate']
                        ws.cell(cur, 5).value = d['total']
                        cur += 1

    def update_final_total(self, ws, project_data):
        total = self.calculate_total_from_data(project_data)
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                if cell.value and "КОНЕЧНАЯ СТОИМОСТЬ" in str(cell.value).upper():
                    ws.cell(r, 5).value = f"{total:.2f}р."
                    return

    def calculate_total_from_data(self, project_data: Dict[str, Any]) -> float:
        total = 0.0
        for pair in project_data.get('language_pairs', []):
            for svc in pair.get('services', {}).values():
                for row in svc:
                    total += row['total']
        for svc in project_data.get('additional_services', {}).values():
            for row in svc:
                total += row['total']
        return total
