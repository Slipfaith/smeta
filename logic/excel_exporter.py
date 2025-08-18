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
        """Экспортирует данные проекта в Excel файл на основе шаблона"""
        try:
            # Обязательно загружаем шаблон
            if not self.template_path or not os.path.exists(self.template_path):
                raise Exception("Не указан путь к шаблону Excel или файл не существует")

            # Загружаем шаблон с сохранением формул и форматирования
            wb = load_workbook(self.template_path)
            ws = wb.active

            # Заполняем основные данные проекта в соответствующие ячейки
            self.fill_template_data(ws, project_data)

            # Сохраняем файл
            wb.save(output_path)
            return True

        except Exception as e:
            print(f"Ошибка при экспорте в Excel: {e}")
            return False

    def fill_template_data(self, ws, project_data: Dict[str, Any]):
        """Заполняет данные в шаблон Excel"""

        # Основная информация проекта — плейсхолдеры
        for row in ws.iter_rows():
            for cell in row:
                if cell.value:
                    cell_value = str(cell.value)

                    if "{{project_name}}" in cell_value:
                        cell.value = cell_value.replace("{{project_name}}", project_data.get('project_name', ''))
                    elif "{{target_langs}}" in cell_value:
                        pairs = [pair['pair_name'] for pair in project_data.get('language_pairs', [])]
                        cell.value = cell_value.replace("{{target_langs}}", ", ".join(pairs))
                    elif "{{client}}" in cell_value:
                        cell.value = cell_value.replace("{{client}}", project_data.get('client_name', ''))
                    elif "{{Entity}}" in cell_value:
                        cell.value = cell_value.replace("{{Entity}}", project_data.get('client_name', ''))
                    elif "{{Entity_address}}" in cell_value:
                        cell.value = cell_value.replace("{{Entity_address}}", project_data.get('email', ''))
                    elif "{{client_name}}" in cell_value:
                        cell.value = cell_value.replace("{{client_name}}", project_data.get('contact_person', ''))
                    elif "{{PM_name}}" in cell_value:
                        cell.value = cell_value.replace("{{PM_name}}", "Project Manager")  # или из настроек
                    elif "{{PM_email}}" in cell_value:
                        cell.value = cell_value.replace("{{PM_email}}", "pm@company.com")  # или из настроек

        # Заполняем данные языковых пар
        self.fill_language_pairs_data(ws, project_data)

        # Заполняем дополнительные услуги
        self.fill_additional_services_data(ws, project_data)

    def fill_language_pairs_data(self, ws, project_data: Dict[str, Any]):
        """Заполняет данные языковых пар в существующие таблицы шаблона"""
        language_pairs = project_data.get('language_pairs', [])
        if not language_pairs:
            return

        # Находим все таблицы в шаблоне
        table_sections = self.find_table_sections(ws)

        # Если языковых пар больше, чем таблиц в шаблоне, дублируем последнюю таблицу
        if len(language_pairs) > len(table_sections):
            self.duplicate_table_sections(ws, table_sections, len(language_pairs))
            table_sections = self.find_table_sections(ws)  # Обновляем список после дублирования

        # Заполняем каждую таблицу данными соответствующей языковой пары
        for i, pair_data in enumerate(language_pairs):
            if i < len(table_sections):
                table_info = table_sections[i]
                self.fill_single_language_table(ws, pair_data, table_info)

    def find_table_sections(self, ws):
        """Находит все секции с таблицами в шаблоне"""
        table_sections = []

        # Ищем заголовки таблиц (например, "Китайский", или цветные ячейки)
        for row_num in range(1, ws.max_row + 1):
            for col_num in range(1, ws.max_column + 1):
                cell = ws.cell(row_num, col_num)

                # Проверяем, является ли ячейка заголовком таблицы
                if self.is_table_header(cell):
                    # Находим границы таблицы
                    table_info = self.get_table_boundaries(ws, row_num, col_num)
                    if table_info:
                        table_sections.append(table_info)

        return table_sections

    def is_table_header(self, cell):
        """Определяет, является ли ячейка заголовком таблицы"""
        if not cell.value:
            return False

        cell_value = str(cell.value).strip()

        indicators = [
            "Китайский", "Английский", "Немецкий", "Французский",
            "→", "Перевод", "Локализация", "Название работы"
        ]

        # Цвет/жирность — доп. признаки
        has_fill = cell.fill and cell.fill.start_color and cell.fill.start_color.index != '00000000'
        is_bold = cell.font and cell.font.bold

        return any(indicator in cell_value for indicator in indicators) or has_fill or is_bold

    def get_table_boundaries(self, ws, header_row, header_col):
        """Определяет границы таблицы начиная от заголовка"""
        # Ищем строку с заголовками колонок
        column_headers_row = None
        for row in range(header_row + 1, min(header_row + 5, ws.max_row + 1)):
            cell_a = ws.cell(row, 1)
            if cell_a.value and ("Название работы" in str(cell_a.value) or "Параметр" in str(cell_a.value)):
                column_headers_row = row
                break

        if not column_headers_row:
            return None

        # Последняя строка таблицы до пустой/итоговой
        last_data_row = column_headers_row
        for row in range(column_headers_row + 1, ws.max_row + 1):
            cell_a = ws.cell(row, 1)

            if (not cell_a.value or
                "Промежуточная сумма" in str(cell_a.value) or
                "КОНЕЧНАЯ СТОИМОСТЬ" in str(cell_a.value)):
                break

            if self.is_table_header(cell_a):
                break

            last_data_row = row

        return {
            'header_row': header_row,
            'column_headers_row': column_headers_row,
            'first_data_row': column_headers_row + 1,
            'last_data_row': last_data_row,
            'subtotal_row': last_data_row + 1
        }

    def duplicate_table_sections(self, ws, existing_sections, needed_count):
        """Дублирует таблицы если языковых пар больше чем таблиц"""
        if not existing_sections or len(existing_sections) >= needed_count:
            return

        # Берем последнюю таблицу как образец
        template_section = existing_sections[-1]

        # Высота одной таблицы
        table_height = template_section['subtotal_row'] - template_section['header_row'] + 2

        # Дублируем
        for i in range(len(existing_sections), needed_count):
            new_start_row = template_section['subtotal_row'] + 2 + (i - len(existing_sections) + 1) * table_height

            for offset in range(table_height):
                source_row = template_section['header_row'] + offset
                target_row = new_start_row + offset

                # копируем A..E (под твой текущий шаблон/код)
                for col in range(1, 6):
                    source_cell = ws.cell(source_row, col)
                    target_cell = ws.cell(target_row, col)

                    target_cell.value = source_cell.value
                    if source_cell.font:
                        target_cell.font = Font(
                            name=source_cell.font.name,
                            size=source_cell.font.size,
                            bold=source_cell.font.bold,
                            color=source_cell.font.color
                        )
                    if source_cell.fill:
                        target_cell.fill = PatternFill(
                            fill_type=source_cell.fill.fill_type,
                            start_color=source_cell.fill.start_color,
                            end_color=source_cell.fill.end_color
                        )
                    if source_cell.alignment:
                        target_cell.alignment = Alignment(
                            horizontal=source_cell.alignment.horizontal,
                            vertical=source_cell.alignment.vertical
                        )

    def fill_single_language_table(self, ws, pair_data, table_info):
        """Заполняет одну таблицу данными языковой пары"""
        # Заголовок
        header_cell = ws.cell(table_info['header_row'], 1)
        header_cell.value = pair_data['pair_name']

        # Очистка строк данных
        for row in range(table_info['first_data_row'], table_info['last_data_row'] + 1):
            for col in range(1, 6):
                ws.cell(row, col).value = None

        # Заполнение
        current_row = table_info['first_data_row']
        total_sum = 0

        for service_name, service_data in pair_data.get('services', {}).items():
            for row_data in service_data:
                if row_data['volume'] > 0 or row_data['rate'] > 0:
                    if current_row <= table_info['last_data_row']:
                        ws.cell(current_row, 1).value = row_data['parameter']  # Название работы
                        ws.cell(current_row, 2).value = "Слово"               # Тип
                        ws.cell(current_row, 3).value = row_data['volume']    # Кол-во
                        ws.cell(current_row, 4).value = row_data['rate']      # Ставка
                        ws.cell(current_row, 5).value = row_data['total']     # Итого

                        total_sum += row_data['total']
                        current_row += 1

        # Промежуточная сумма
        subtotal_cell = ws.cell(table_info['subtotal_row'], 5)
        subtotal_cell.value = f"{total_sum:.2f}р."

    def fill_additional_services_data(self, ws, project_data: Dict[str, Any]):
        """Заполняет дополнительные услуги в существующие разделы шаблона"""
        additional_services = project_data.get('additional_services', {})
        if not additional_services:
            return

        # Ищем раздел "Подготовка проекта" (или аналогичный)
        prep_section = self.find_preparation_section(ws)
        if prep_section:
            self.fill_preparation_section(ws, prep_section, additional_services)

        # Итог
        self.update_final_total(ws, project_data)

    def find_preparation_section(self, ws):
        """Находит раздел 'Подготовка проекта' или аналогичный"""
        for row_num in range(1, ws.max_row + 1):
            cell = ws.cell(row_num, 1)
            if cell.value and "Подготовка проекта" in str(cell.value):
                headers_row = None
                for r in range(row_num + 1, min(row_num + 5, ws.max_row + 1)):
                    if ws.cell(r, 1).value and "Название работы" in str(ws.cell(r, 1).value):
                        headers_row = r
                        break

                if headers_row:
                    return {
                        'header_row': row_num,
                        'column_headers_row': headers_row,
                        'first_data_row': headers_row + 1,
                        'last_data_row': headers_row + 10  # до 10 строк по умолчанию
                    }
        return None

    def fill_preparation_section(self, ws, section, additional_services):
        """Заполняет раздел подготовки проекта дополнительными услугами"""
        current_row = section['first_data_row']

        # Очистка
        for row in range(section['first_data_row'], section['last_data_row'] + 1):
            for col in range(1, 6):
                ws.cell(row, col).value = None

        # Заполнение
        for service_name, service_data in additional_services.items():
            for row_data in service_data:
                if row_data['volume'] > 0 or row_data['rate'] > 0:
                    if current_row <= section['last_data_row']:
                        ws.cell(current_row, 1).value = f"{service_name}: {row_data['parameter']}"
                        ws.cell(current_row, 2).value = "Час"  # при необходимости поменяешь
                        ws.cell(current_row, 3).value = row_data['volume']
                        ws.cell(current_row, 4).value = row_data['rate']
                        ws.cell(current_row, 5).value = row_data['total']
                        current_row += 1

    def update_final_total(self, ws, project_data):
        """Обновляет итоговую сумму в шаблоне"""
        total_sum = self.calculate_total_from_data(project_data)

        # Ищем "КОНЕЧНАЯ СТОИМОСТЬ"
        for row_num in range(1, ws.max_row + 1):
            for col_num in range(1, ws.max_column + 1):
                cell = ws.cell(row_num, col_num)
                if cell.value and "КОНЕЧНАЯ СТОИМОСТЬ" in str(cell.value).upper():
                    # результат в колонке E (как в твоем коде)
                    result_cell = ws.cell(row_num, 5)
                    result_cell.value = f"{total_sum:.2f}р."
                    return

    def calculate_total_from_data(self, project_data: Dict[str, Any]) -> float:
        """Вычисляет общую сумму из данных проекта"""
        total = 0.0

        # Языковые пары
        for pair_data in project_data.get('language_pairs', []):
            for service_data in pair_data.get('services', {}).values():
                for row_data in service_data:
                    total += row_data['total']

        # Дополнительные услуги
        for service_data in project_data.get('additional_services', {}).values():
            for row_data in service_data:
                total += row_data['total']

        return total

    # Ниже — альтернативные методы построения с нуля (оставлены как утилиты)

    def write_header(self, ws, project_data: Dict[str, Any], start_row: int):
        """Записывает заголовок документа"""
        ws[f'A{start_row}'] = "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ"
        ws[f'A{start_row}'].font = Font(bold=True, size=16)

        ws[f'A{start_row + 1}'] = f"Проект: {project_data.get('project_name', '')}"
        ws[f'A{start_row + 2}'] = f"Номер/код: {project_data.get('project_code', '')}"
        ws[f'A{start_row + 3}'] = f"Клиент: {project_data.get('client_name', '')}"
        ws[f'A{start_row + 4}'] = f"Дата: {datetime.now().strftime('%d.%m.%Y')}"

    def write_language_pair(self, ws, pair_data: Dict[str, Any], start_row: int) -> int:
        """Записывает данные языковой пары"""
        current_row = start_row

        ws[f'A{current_row}'] = pair_data['pair_name']
        ws[f'A{current_row}'].font = Font(bold=True, size=12)
        current_row += 1

        headers = ["Параметр", "Объем", "Ставка (руб)", "Сумма (руб)"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(current_row, col, header)
            cell.font = Font(bold=True)
        current_row += 1

        for service_name, service_data in pair_data.get('services', {}).items():
            ws[f'A{current_row}'] = service_name.upper()
            ws[f'A{current_row}'].font = Font(bold=True)
            current_row += 1

            for row_data in service_data:
                ws[f'A{current_row}'] = row_data['parameter']
                ws[f'B{current_row}'] = row_data['volume']
                ws[f'C{current_row}'] = row_data['rate']
                ws[f'D{current_row}'] = row_data['total']
                current_row += 1

            current_row += 1  # пустая строка после услуги

        return current_row

    def write_additional_services(self, ws, services_data: Dict[str, Any], start_row: int) -> int:
        """Записывает дополнительные услуги"""
        current_row = start_row

        ws[f'A{current_row}'] = "ДОПОЛНИТЕЛЬНЫЕ УСЛУГИ"
        ws[f'A{current_row}'].font = Font(bold=True, size=12)
        current_row += 2

        for service_name, service_data in services_data.items():
            ws[f'A{current_row}'] = service_name
            ws[f'A{current_row}'].font = Font(bold=True)
            current_row += 1

            headers = ["Параметр", "Объем", "Ставка (руб)", "Сумма (руб)"]
            for col, header in enumerate(headers, 1):
                cell = ws.cell(current_row, col, header)
                cell.font = Font(bold=True)
            current_row += 1

            for row_data in service_data:
                ws[f'A{current_row}'] = row_data['parameter']
                ws[f'B{current_row}'] = row_data['volume']
                ws[f'C{current_row}'] = row_data['rate']
                ws[f'D{current_row}'] = row_data['total']
                current_row += 1

            current_row += 1  # пустая строка

        return current_row
