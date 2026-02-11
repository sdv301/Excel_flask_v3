# reports/template_report_generator.py - ПОЛНАЯ ВЕРСИЯ БЕЗ ОГРАНИЧЕНИЙ
import os
import shutil
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime, date
import json

class TemplateReportGenerator:
    def __init__(self, db_connection, template_path: str = None):
        self.db = db_connection
        self.reports_dir = 'reports_output'
        os.makedirs(self.reports_dir, exist_ok=True)
        
        if template_path is None:
            self.template_path = 'report_templates/Сводный_отчет_шаблон.xlsx'
        else:
            self.template_path = template_path
        
        if not os.path.exists(self.template_path):
            possible_paths = [
                'report_templates/Сводный_отчет_шаблон.xlsx',
                '../report_templates/Сводный_отчет_шаблон.xlsx',
                './report_templates/Сводный_отчет_шаблон.xlsx'
            ]
            for path in possible_paths:
                if os.path.exists(path):
                    self.template_path = path
                    break
            else:
                raise FileNotFoundError(f"Шаблон не найден. Искал: {possible_paths}")

    def generate_report(self, report_date: date = None) -> str:
        try:
            if report_date is None:
                report_date = datetime.now().date()

            print(f"\n🎯 ГЕНЕРАЦИЯ ОТЧЕТА НА {report_date.strftime('%d.%m.%Y')}")

            aggregated_data = self.db.get_aggregated_data()
            if not aggregated_data:
                raise Exception("Нет данных в БД")

            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f'Сводный_отчет_{timestamp}.xlsx'
            output_path = os.path.join(self.reports_dir, filename)
            
            shutil.copy2(self.template_path, output_path)

            wb = load_workbook(output_path)
            self._update_report_info(wb, report_date, aggregated_data)
            self._fill_all_company_data(wb, aggregated_data)
            wb.save(output_path)

            if os.path.exists(output_path):
                print(f"✅ Отчет создан успешно: {output_path}")
                return output_path
            else:
                raise Exception("Файл не был создан")
        except Exception as e:
            print(f"❌ Ошибка: {e}")
            raise

    def _update_report_info(self, wb, report_date: date, aggregated_data: dict):
        date_str = report_date.strftime('%d.%m.%Y')
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in range(1, 6):
                for col in range(1, 10):
                    cell = ws.cell(row=row, column=col)
                    if cell.value and 'дата' in str(cell.value).lower():
                        ws.cell(row=row, column=col+1).value = date_str

    def _fill_all_company_data(self, wb, aggregated_data: dict):
        if '1-Структура' in wb.sheetnames:
            self._fill_structure_sheet_full(wb['1-Структура'], aggregated_data)
        if '2-Потребность' in wb.sheetnames:
            self._fill_demand_sheet_full(wb['2-Потребность'], aggregated_data)
        if '3-Остатки' in wb.sheetnames:
            self._fill_stocks_sheet_full(wb['3-Остатки'], aggregated_data)
        if '4-Поставка' in wb.sheetnames:
            self._fill_supply_sheet_full(wb['4-Поставка'], aggregated_data)
        if '5-Реализация' in wb.sheetnames:
            self._fill_sales_sheet_full(wb['5-Реализация'], aggregated_data)
        if '6-Авиатопливо' in wb.sheetnames:
            self._fill_aviation_sheet_full(wb['6-Авиатопливо'], aggregated_data)
        if '7-Комментарии' in wb.sheetnames:
            self._fill_comments_sheet_full(wb['7-Комментарии'], aggregated_data)

    def _fill_structure_sheet_full(self, ws, aggregated_data: dict):
        start_row = 13
        current_row = start_row
        for company_name, company_data in aggregated_data.items():
            for record in company_data.get('sheet1', []):
                if 'наименование компаний' in str(record.get('company_name', '')).lower(): continue
                self._set_cell_value(ws, current_row, 1, record.get('affiliation', ''))
                self._set_cell_value(ws, current_row, 2, record.get('company_name', company_name))
                self._set_cell_value(ws, current_row, 3, record.get('oil_depots_count', 0))
                self._set_cell_value(ws, current_row, 4, record.get('azs_count', 0))
                self._set_cell_value(ws, current_row, 5, record.get('working_azs_count', 0))
                current_row += 1

    def _fill_demand_sheet_full(self, ws, aggregated_data: dict):
        year_row = 7
        month_row = 13
        cur_year_row = year_row
        cur_month_row = month_row
        for company_name, company_data in aggregated_data.items():
            data = company_data.get('sheet2', {})
            if data:
                # Год
                self._set_cell_value(ws, cur_year_row, 1, company_name)
                self._set_cell_value(ws, cur_year_row, 4, data.get('gasoline_ai92', 0))
                self._set_cell_value(ws, cur_year_row, 5, data.get('gasoline_ai95', 0))
                self._set_cell_value(ws, cur_year_row, 8, data.get('diesel_total', 0))
                # Месяц
                self._set_cell_value(ws, cur_month_row, 1, company_name)
                self._set_cell_value(ws, cur_month_row, 4, data.get('monthly_gasoline_total', 0) / 2 if data.get('monthly_gasoline_total') else 0)
                self._set_cell_value(ws, cur_month_row, 5, data.get('monthly_gasoline_total', 0) / 2 if data.get('monthly_gasoline_total') else 0)
                self._set_cell_value(ws, cur_month_row, 8, data.get('monthly_diesel_total', 0))
                cur_year_row += 1
                cur_month_row += 1

    def _fill_stocks_sheet_full(self, ws, aggregated_data: dict):
        start_row = 9
        current_row = start_row
        for company_name, company_data in aggregated_data.items():
            sheet3_recs = company_data.get('sheet3_data', [])
            print(f"DEBUG: Filling Sheet 3 for {company_name}, records: {len(sheet3_recs)}")
            for loc in sheet3_recs:
                self._set_cell_value(ws, current_row, 2, company_name)
                self._set_cell_value(ws, current_row, 3, loc.get('location_name', ''))
                # Stocks (Columns 4-11: 76, 92, 95, 98, Winter, Arctic, Summer, Intermediate)
                self._set_cell_value(ws, current_row, 5, loc.get('stock_ai92', 0))
                self._set_cell_value(ws, current_row, 6, loc.get('stock_ai95', 0))
                self._set_cell_value(ws, current_row, 7, loc.get('stock_ai98_ai100', 0))
                self._set_cell_value(ws, current_row, 8, loc.get('stock_diesel_winter', 0))
                self._set_cell_value(ws, current_row, 9, loc.get('stock_diesel_arctic', 0))
                self._set_cell_value(ws, current_row, 10, loc.get('stock_diesel_summer', 0))
                # Transit (Columns 12-19)
                self._set_cell_value(ws, current_row, 13, loc.get('transit_ai92', 0))
                self._set_cell_value(ws, current_row, 14, loc.get('transit_ai95', 0))
                self._set_cell_value(ws, current_row, 15, loc.get('transit_ai98_ai100', 0))
                self._set_cell_value(ws, current_row, 16, loc.get('transit_diesel_winter', 0))
                self._set_cell_value(ws, current_row, 17, loc.get('transit_diesel_arctic', 0))
                self._set_cell_value(ws, current_row, 18, loc.get('transit_diesel_summer', 0))
                # Capacity (Columns 20-27)
                self._set_cell_value(ws, current_row, 21, loc.get('capacity_ai92', 0))
                self._set_cell_value(ws, current_row, 22, loc.get('capacity_ai95', 0))
                self._set_cell_value(ws, current_row, 23, loc.get('capacity_ai98_ai100', 0))
                self._set_cell_value(ws, current_row, 24, loc.get('capacity_diesel_winter', 0))
                self._set_cell_value(ws, current_row, 25, loc.get('capacity_diesel_arctic', 0))
                self._set_cell_value(ws, current_row, 26, loc.get('capacity_diesel_summer', 0))
                current_row += 1

    def _fill_supply_sheet_full(self, ws, aggregated_data: dict):
        start_row = 9
        current_row = start_row
        for company_name, company_data in aggregated_data.items():
            for supply in company_data.get('sheet4_data', []):
                self._set_cell_value(ws, current_row, 2, company_name)
                self._set_cell_value(ws, current_row, 3, supply.get('oil_depot_name', ''))
                self._set_cell_value(ws, current_row, 4, str(supply.get('supply_date', '')))
                self._set_cell_value(ws, current_row, 6, supply.get('supply_ai92', 0))
                self._set_cell_value(ws, current_row, 7, supply.get('supply_ai95', 0))
                self._set_cell_value(ws, current_row, 8, supply.get('supply_ai98_100', 0))
                self._set_cell_value(ws, current_row, 9, supply.get('supply_diesel_winter', 0))
                self._set_cell_value(ws, current_row, 10, supply.get('supply_diesel_arctic', 0))
                self._set_cell_value(ws, current_row, 11, supply.get('supply_diesel_summer', 0))
                current_row += 1

    def _fill_sales_sheet_full(self, ws, aggregated_data: dict):
        start_row = 9
        current_row = start_row
        for company_name, company_data in aggregated_data.items():
            for sales in company_data.get('sheet5_data', []):
                self._set_cell_value(ws, current_row, 2, company_name)
                self._set_cell_value(ws, current_row, 3, sales.get('location_name', ''))
                # Daily
                self._set_cell_value(ws, current_row, 5, sales.get('daily_ai92', 0))
                self._set_cell_value(ws, current_row, 6, sales.get('daily_ai95', 0))
                self._set_cell_value(ws, current_row, 7, sales.get('daily_ai98_100', 0))
                self._set_cell_value(ws, current_row, 8, sales.get('daily_winter', 0))
                self._set_cell_value(ws, current_row, 9, sales.get('daily_arctic', 0))
                self._set_cell_value(ws, current_row, 10, sales.get('daily_summer', 0))
                # Monthly
                self._set_cell_value(ws, current_row, 13, sales.get('monthly_ai92', 0))
                self._set_cell_value(ws, current_row, 14, sales.get('monthly_ai95', 0))
                self._set_cell_value(ws, current_row, 15, sales.get('monthly_ai98_100', 0))
                self._set_cell_value(ws, current_row, 16, sales.get('monthly_diesel_winter', 0))
                self._set_cell_value(ws, current_row, 17, sales.get('monthly_diesel_arctic', 0))
                self._set_cell_value(ws, current_row, 18, sales.get('monthly_diesel_summer', 0))
                current_row += 1

    def _fill_aviation_sheet_full(self, ws, aggregated_data: dict):
        start_row = 8
        current_row = start_row
        for company_name, company_data in aggregated_data.items():
            for item in company_data.get('sheet6_data', []):
                self._set_cell_value(ws, current_row, 1, item.get('airport_name', ''))
                self._set_cell_value(ws, current_row, 2, item.get('tzk_name', ''))
                self._set_cell_value(ws, current_row, 3, item.get('contracts_info', ''))
                self._set_cell_value(ws, current_row, 4, item.get('supply_week', 0))
                self._set_cell_value(ws, current_row, 5, item.get('supply_month_start', 0))
                self._set_cell_value(ws, current_row, 6, item.get('monthly_demand', 0))
                self._set_cell_value(ws, current_row, 7, item.get('consumption_week', 0))
                self._set_cell_value(ws, current_row, 8, item.get('consumption_month_start', 0))
                self._set_cell_value(ws, current_row, 9, item.get('end_of_day_balance', 0))
                current_row += 1

    def _fill_comments_sheet_full(self, ws, aggregated_data: dict):
        start_row = 6
        current_row = start_row
        for company_name, company_data in aggregated_data.items():
            for item in company_data.get('sheet7_data', []):
                self._set_cell_value(ws, current_row, 1, item.get('fuel_type', ''))
                self._set_cell_value(ws, current_row, 2, item.get('situation', ''))
                self._set_cell_value(ws, current_row, 3, item.get('comments', ''))
                current_row += 1

    def _set_cell_value(self, ws, row: int, col: int, value):
        try:
            if value is None: value = 0 if isinstance(value, (int, float)) else ""
            if row > 0 and col > 0:
                ws.cell(row=row, column=col).value = value
                return True
        except: return False

def generate_complete_report(db_connection, template_path=None):
    generator = TemplateReportGenerator(db_connection, template_path)
    return generator.generate_report()

if __name__ == "__main__":
    from database.queries import db
    report_path = generate_complete_report(db)
    print(f"Отчет готов: {report_path}")
