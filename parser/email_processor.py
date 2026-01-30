# parser/excel_parser.py
import pandas as pd
import numpy as np
from datetime import datetime
import re
from typing import Dict, List, Any, Optional
from dataclasses import dataclass
import warnings
warnings.filterwarnings('ignore')

@dataclass
class ExcelFileMetadata:
    filename: str
    report_date: datetime
    company_name: str
    sheets_data: Dict[str, pd.DataFrame]

class FuelReportParser:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.metadata = None
        self.xls = None
        
    def parse(self) -> ExcelFileMetadata:
        """Основной метод парсинга файла"""
        try:
            # Чтение всех листов
            self.xls = pd.ExcelFile(self.file_path, engine='openpyxl')
            sheets_data = {}
            
            for sheet_name in self.xls.sheet_names:
                try:
                    # Читаем без заголовков
                    df = pd.read_excel(self.xls, sheet_name=sheet_name, 
                                     header=None, engine='openpyxl')
                    sheets_data[sheet_name] = df
                except Exception as e:
                    print(f"Ошибка при чтении листа {sheet_name}: {e}")
                    sheets_data[sheet_name] = pd.DataFrame()
            
            # Извлечение метаданных
            metadata = self._extract_metadata(sheets_data)
            metadata.sheets_data = sheets_data
            
            self.metadata = metadata
            return metadata
            
        except Exception as e:
            raise Exception(f"Ошибка парсинга файла {self.file_path}: {e}")
    
    def _extract_metadata(self, sheets_data: Dict) -> ExcelFileMetadata:
        """Извлечение метаданных из файла"""
        # Извлекаем дату отчета (обычно в ячейке B2 на листе 1)
        sheet1 = sheets_data.get('1-Структура')
        report_date_str = None
        
        if sheet1 is not None and not sheet1.empty:
            # Поиск строки с "Информация по состоянию за:"
            for idx, row in sheet1.iterrows():
                if isinstance(row[0], str) and 'Информация по состоянию за:' in str(row[0]):
                    if len(row) > 1:
                        report_date_str = row[1]
                    break
        
        # Парсинг даты
        report_date = datetime.now()
        if report_date_str:
            try:
                if isinstance(report_date_str, datetime):
                    report_date = report_date_str
                elif isinstance(report_date_str, str):
                    # Пробуем разные форматы даты
                    for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y']:
                        try:
                            report_date = datetime.strptime(str(report_date_str).strip(), fmt)
                            break
                        except:
                            continue
            except:
                pass
        
        # Определяем компанию по имени файла
        filename = self.file_path.split('/')[-1]
        company_name = self._detect_company_name(filename)
        
        return ExcelFileMetadata(
            filename=filename,
            report_date=report_date,
            company_name=company_name,
            sheets_data={}
        )
    
    def _detect_company_name(self, filename: str) -> str:
        """Определение названия компании по имени файла"""
        filename_lower = filename.lower()
        
        # Более точное определение компании
        if 'саха' in filename_lower:
            return 'Саханефтегазсбыт'
        elif 'туймаада' in filename_lower:
            return 'Туймаада-Нефть'
        elif 'сибойл' in filename_lower:
            return 'Сибойл'
        elif 'экто' in filename_lower:
            return 'ЭКТО-Ойл'
        elif 'сибирск' in filename_lower:
            return 'Сибирское топливо'
        elif 'паритет' in filename_lower:
            return 'Паритет'
        else:
            # Пробуем извлечь из содержимого файла
            return 'Неизвестная компания'
    
    def _safe_int(self, value) -> int:
        """Безопасное преобразование в int"""
        try:
            if pd.isna(value) or value is None:
                return 0
            # Если это строка, убираем лишние символы
            if isinstance(value, str):
                value = value.replace(',', '.').strip()
                # Убираем все нецифровые символы кроме минуса и точки
                value = re.sub(r'[^\d.-]', '', value)
            return int(float(value))
        except:
            return 0
    
    def _safe_float(self, value) -> float:
        """Безопасное преобразование в float"""
        try:
            if pd.isna(value) or value is None:
                return 0.0
            # Если это строка, убираем лишние символы
            if isinstance(value, str):
                value = value.replace(',', '.').strip()
                # Убираем все нецифровые символы кроме минуса и точки
                value = re.sub(r'[^\d.-]', '', value)
            return float(value)
        except:
            return 0.0
    
    def _safe_str(self, value) -> str:
        """Безопасное преобразование в строку"""
        if pd.isna(value) or value is None:
            return ''
        return str(value).strip()
    
    def extract_all_data(self) -> Dict[str, Any]:
        """Извлечение данных со всех листов"""
        if not self.metadata:
            self.parse()
        
        return {
            'sheet1': self.extract_sheet1_data(),
            'sheet2': self.extract_sheet2_data(),
            'sheet3': self.extract_sheet3_data(),
            'sheet4': self.extract_sheet4_data(),
            'sheet5': self.extract_sheet5_data(),
            'sheet6': self.extract_sheet6_data(),
            'sheet7': self.extract_sheet7_data(),
        }
    
    def extract_sheet1_data(self) -> List[Dict]:
        """Извлечение данных из листа 1 (Структура)"""
        if not self.metadata:
            return []
        
        sheet1 = self.metadata.sheets_data.get('1-Структура')
        if sheet1 is None or sheet1.empty:
            return []
        
        data = []
        # Поиск начала таблицы №1 (Нефтебазы и АЗС)
        table_start = None
        for idx, row in sheet1.iterrows():
            cell_value = self._safe_str(row[0])
            if 'Таблица №1' in cell_value and 'Нефтебазы и автозаправочные станции' in cell_value:
                table_start = idx + 3  # Пропускаем заголовок и строку с номерами колонок
                break
        
        if table_start is None:
            return []
        
        # Извлечение строк таблицы
        for idx in range(table_start, len(sheet1)):
            row = sheet1.iloc[idx]
            
            # Проверка на конец таблицы
            cell_value = self._safe_str(row[0])
            if pd.isna(row[0]) or not cell_value or 'Таблица №2' in cell_value:
                break
            
            data.append({
                'affiliation': self._safe_str(row[0]),
                'company_name': self._safe_str(row[1]),
                'oil_depots_count': self._safe_int(row[2]),
                'azs_count': self._safe_int(row[3]),
                'working_azs_count': self._safe_int(row[4])
            })
        
        return data
    
    def extract_sheet2_data(self) -> Dict[str, Any]:
        """Извлечение данных из листа 2 (Потребность)"""
        if not self.metadata:
            return {}
        
        sheet2 = self.metadata.sheets_data.get('2-Потребность')
        if sheet2 is None or sheet2.empty:
            return {}
        
        data = {}
        
        # Поиск годовой потребности
        for idx, row in sheet2.iterrows():
            cell_value = self._safe_str(row[0])
            if 'ГОД' in cell_value:
                year_match = re.search(r'ГОД\s+(\d{4})', cell_value)
                year = int(year_match.group(1)) if year_match else datetime.now().year
                
                data.update({
                    'year': year,
                    'gasoline_total': self._safe_float(row[1]) if len(row) > 1 else 0,
                    'gasoline_ai76_80': self._safe_float(row[2]) if len(row) > 2 else 0,
                    'gasoline_ai92': self._safe_float(row[3]) if len(row) > 3 else 0,
                    'gasoline_ai95': self._safe_float(row[4]) if len(row) > 4 else 0,
                    'gasoline_ai98_100': self._safe_float(row[5]) if len(row) > 5 else 0,
                    'diesel_total': self._safe_float(row[6]) if len(row) > 6 else 0,
                    'diesel_winter': self._safe_float(row[7]) if len(row) > 7 else 0,
                    'diesel_arctic': self._safe_float(row[8]) if len(row) > 8 else 0,
                    'diesel_summer': self._safe_float(row[9]) if len(row) > 9 else 0,
                    'diesel_intermediate': self._safe_float(row[10]) if len(row) > 10 else 0,
                })
                break
        
        # Поиск месячной потребности
        for idx, row in sheet2.iterrows():
            cell_value = self._safe_str(row[0])
            if 'МЕСЯЦ' in cell_value:
                month_match = re.search(r'МЕСЯЦ\s+(\w+)', cell_value)
                month = month_match.group(1) if month_match else ''
                
                data.update({
                    'month': month,
                    'monthly_gasoline_total': self._safe_float(row[1]) if len(row) > 1 else 0,
                    'monthly_gasoline_ai76_80': self._safe_float(row[2]) if len(row) > 2 else 0,
                    'monthly_gasoline_ai92': self._safe_float(row[3]) if len(row) > 3 else 0,
                    'monthly_gasoline_ai95': self._safe_float(row[4]) if len(row) > 4 else 0,
                    'monthly_gasoline_ai98_100': self._safe_float(row[5]) if len(row) > 5 else 0,
                    'monthly_diesel_total': self._safe_float(row[6]) if len(row) > 6 else 0,
                    'monthly_diesel_winter': self._safe_float(row[7]) if len(row) > 7 else 0,
                    'monthly_diesel_arctic': self._safe_float(row[8]) if len(row) > 8 else 0,
                    'monthly_diesel_summer': self._safe_float(row[9]) if len(row) > 9 else 0,
                    'monthly_diesel_intermediate': self._safe_float(row[10]) if len(row) > 10 else 0,
                })
                break
        
        return data
    
    def extract_sheet3_data(self) -> List[Dict]:
        """Извлечение данных из листа 3 (Остатки)"""
        if not self.metadata:
            return []
        
        sheet3 = self.metadata.sheets_data.get('3-Остатки')
        if sheet3 is None or sheet3.empty:
            return []
        
        data = []
        # Поиск начала таблицы №5
        table_start = None
        for idx, row in sheet3.iterrows():
            cell_value = self._safe_str(row[0])
            if 'Таблица №5' in cell_value and 'Наличие моторного топлива' in cell_value:
                table_start = idx + 3  # Пропускаем заголовок и строку с заголовками колонок
                break
        
        if table_start is None:
            return []
        
        # Извлечение строк таблицы
        for idx in range(table_start, min(table_start + 100, len(sheet3))):  # Ограничиваем 100 строками
            row = sheet3.iloc[idx]
            
            # Проверка на конец таблицы (пустая строка или начало новой секции)
            cell_value = self._safe_str(row[0])
            if pd.isna(row[0]) or not cell_value or cell_value.startswith('='):
                # Пропускаем строки с формулами
                if cell_value.startswith('='):
                    continue
                # Проверяем, не началась ли следующая таблица
                check_idx = idx + 1
                if check_idx < len(sheet3):
                    next_cell = self._safe_str(sheet3.iloc[check_idx][0])
                    if 'Таблица №' in next_cell:
                        break
            
            data.append({
                'affiliation': self._safe_str(row[0]),
                'company_name': self._safe_str(row[1]) if len(row) > 1 else '',
                'location_type': self._safe_str(row[2]) if len(row) > 2 else '',
                'location_name': self._safe_str(row[2]) if len(row) > 2 else '',  # Дублируем для удобства
                # Имеющиеся запасы
                'stock_ai76_80': self._safe_float(row[3]) if len(row) > 3 else 0,
                'stock_ai92': self._safe_float(row[4]) if len(row) > 4 else 0,
                'stock_ai95': self._safe_float(row[5]) if len(row) > 5 else 0,
                'stock_ai98_100': self._safe_float(row[6]) if len(row) > 6 else 0,
                'stock_diesel_winter': self._safe_float(row[7]) if len(row) > 7 else 0,
                'stock_diesel_arctic': self._safe_float(row[8]) if len(row) > 8 else 0,
                'stock_diesel_summer': self._safe_float(row[9]) if len(row) > 9 else 0,
                'stock_diesel_intermediate': self._safe_float(row[10]) if len(row) > 10 else 0,
                # Товар в пути
                'transit_ai76_80': self._safe_float(row[11]) if len(row) > 11 else 0,
                'transit_ai92': self._safe_float(row[12]) if len(row) > 12 else 0,
                'transit_ai95': self._safe_float(row[13]) if len(row) > 13 else 0,
                'transit_ai98_100': self._safe_float(row[14]) if len(row) > 14 else 0,
                'transit_diesel_winter': self._safe_float(row[15]) if len(row) > 15 else 0,
                'transit_diesel_arctic': self._safe_float(row[16]) if len(row) > 16 else 0,
                'transit_diesel_summer': self._safe_float(row[17]) if len(row) > 17 else 0,
                'transit_diesel_intermediate': self._safe_float(row[18]) if len(row) > 18 else 0,
                # Емкость хранения
                'capacity_ai76_80': self._safe_float(row[19]) if len(row) > 19 else 0,
                'capacity_ai92': self._safe_float(row[20]) if len(row) > 20 else 0,
                'capacity_ai95': self._safe_float(row[21]) if len(row) > 21 else 0,
                'capacity_ai98_100': self._safe_float(row[22]) if len(row) > 22 else 0,
                'capacity_diesel_winter': self._safe_float(row[23]) if len(row) > 23 else 0,
                'capacity_diesel_arctic': self._safe_float(row[24]) if len(row) > 24 else 0,
                'capacity_diesel_summer': self._safe_float(row[25]) if len(row) > 25 else 0,
                'capacity_diesel_intermediate': self._safe_float(row[26]) if len(row) > 26 else 0,
            })
        
        return data
    
    def extract_sheet4_data(self) -> List[Dict]:
        """Извлечение данных из листа 4 (Поставка)"""
        if not self.metadata:
            return []
        
        sheet4 = self.metadata.sheets_data.get('4-Поставка')
        if sheet4 is None or sheet4.empty:
            return []
        
        data = []
        # Поиск начала таблицы №6
        table_start = None
        for idx, row in sheet4.iterrows():
            cell_value = self._safe_str(row[0])
            if 'Таблица №6' in cell_value and 'Поставка моторного топлива' in cell_value:
                table_start = idx + 3  # Пропускаем заголовок
                break
        
        if table_start is None:
            return []
        
        # Извлечение строк таблицы
        for idx in range(table_start, min(table_start + 50, len(sheet4))):
            row = sheet4.iloc[idx]
            
            # Проверка на конец таблицы
            cell_value = self._safe_str(row[0])
            if pd.isna(row[0]) or not cell_value:
                break
            
            # Парсим дату поставки
            supply_date = None
            if len(row) > 3:
                date_value = row[3]
                if isinstance(date_value, datetime):
                    supply_date = date_value
                elif isinstance(date_value, str):
                    try:
                        supply_date = datetime.strptime(str(date_value).split()[0], '%Y-%m-%d')
                    except:
                        pass
            
            data.append({
                'affiliation': self._safe_str(row[0]),
                'company_name': self._safe_str(row[1]) if len(row) > 1 else '',
                'oil_depot_name': self._safe_str(row[2]) if len(row) > 2 else '',
                'supply_date': supply_date,
                'supply_ai76_80': self._safe_float(row[4]) if len(row) > 4 else 0,
                'supply_ai92': self._safe_float(row[5]) if len(row) > 5 else 0,
                'supply_ai95': self._safe_float(row[6]) if len(row) > 6 else 0,
                'supply_ai98_100': self._safe_float(row[7]) if len(row) > 7 else 0,
                'supply_diesel_winter': self._safe_float(row[8]) if len(row) > 8 else 0,
                'supply_diesel_arctic': self._safe_float(row[9]) if len(row) > 9 else 0,
                'supply_diesel_summer': self._safe_float(row[10]) if len(row) > 10 else 0,
                'supply_diesel_intermediate': self._safe_float(row[11]) if len(row) > 11 else 0,
            })
        
        return data
    
    def extract_sheet5_data(self) -> List[Dict]:
        """Извлечение данных из листа 5 (Реализация)"""
        if not self.metadata:
            return []
        
        sheet5 = self.metadata.sheets_data.get('5-Реализация')
        if sheet5 is None or sheet5.empty:
            return []
        
        data = []
        # Поиск начала таблицы №7
        table_start = None
        for idx, row in sheet5.iterrows():
            cell_value = self._safe_str(row[0])
            if 'Таблица №7' in cell_value and 'Реализация моторного топлива' in cell_value:
                table_start = idx + 3  # Пропускаем заголовок
                break
        
        if table_start is None:
            return []
        
        # Извлечение строк таблицы
        for idx in range(table_start, min(table_start + 100, len(sheet5))):
            row = sheet5.iloc[idx]
            
            # Проверка на конец таблицы
            cell_value = self._safe_str(row[0])
            if pd.isna(row[0]) or not cell_value or cell_value.startswith('='):
                if cell_value.startswith('='):
                    continue
                break
            
            data.append({
                'affiliation': self._safe_str(row[0]),
                'company_name': self._safe_str(row[1]) if len(row) > 1 else '',
                'location_type': self._safe_str(row[2]) if len(row) > 2 else '',
                'location_name': self._safe_str(row[2]) if len(row) > 2 else '',
                # Реализация за сутки
                'daily_ai76_80': self._safe_float(row[3]) if len(row) > 3 else 0,
                'daily_ai92': self._safe_float(row[4]) if len(row) > 4 else 0,
                'daily_ai95': self._safe_float(row[5]) if len(row) > 5 else 0,
                'daily_ai98_100': self._safe_float(row[6]) if len(row) > 6 else 0,
                'daily_diesel_winter': self._safe_float(row[7]) if len(row) > 7 else 0,
                'daily_diesel_arctic': self._safe_float(row[8]) if len(row) > 8 else 0,
                'daily_diesel_summer': self._safe_float(row[9]) if len(row) > 9 else 0,
                'daily_diesel_intermediate': self._safe_float(row[10]) if len(row) > 10 else 0,
                # Реализация с начала месяца
                'monthly_ai76_80': self._safe_float(row[11]) if len(row) > 11 else 0,
                'monthly_ai92': self._safe_float(row[12]) if len(row) > 12 else 0,
                'monthly_ai95': self._safe_float(row[13]) if len(row) > 13 else 0,
                'monthly_ai98_100': self._safe_float(row[14]) if len(row) > 14 else 0,
                'monthly_diesel_winter': self._safe_float(row[15]) if len(row) > 15 else 0,
                'monthly_diesel_arctic': self._safe_float(row[16]) if len(row) > 16 else 0,
                'monthly_diesel_summer': self._safe_float(row[17]) if len(row) > 17 else 0,
                'monthly_diesel_intermediate': self._safe_float(row[18]) if len(row) > 18 else 0,
            })
        
        return data
    
    def extract_sheet6_data(self) -> List[Dict]:
        """Извлечение данных из листа 6 (Авиатопливо)"""
        if not self.metadata:
            return []
        
        sheet6 = self.metadata.sheets_data.get('6-Авиатопливо')
        if sheet6 is None or sheet6.empty:
            return []
        
        data = []
        # Поиск начала таблицы №8
        table_start = None
        for idx, row in sheet6.iterrows():
            cell_value = self._safe_str(row[0])
            if 'Таблица №8' in cell_value and 'Авиатопливо' in cell_value:
                table_start = idx + 3  # Пропускаем заголовок
                break
        
        if table_start is None:
            return []
        
        # Извлечение строк таблицы
        for idx in range(table_start, min(table_start + 50, len(sheet6))):
            row = sheet6.iloc[idx]
            
            # Проверка на конец таблицы
            cell_value = self._safe_str(row[0])
            if pd.isna(row[0]) or not cell_value:
                break
            
            data.append({
                'airport_name': self._safe_str(row[0]),
                'tzk_name': self._safe_str(row[1]) if len(row) > 1 else '',
                'contracts_info': self._safe_str(row[2]) if len(row) > 2 else '',
                'supply_week': self._safe_float(row[3]) if len(row) > 3 else 0,
                'supply_month_start': self._safe_float(row[4]) if len(row) > 4 else 0,
                'monthly_demand': self._safe_float(row[5]) if len(row) > 5 else 0,
                'consumption_week': self._safe_float(row[6]) if len(row) > 6 else 0,
                'consumption_month_start': self._safe_float(row[7]) if len(row) > 7 else 0,
                'end_of_day_balance': self._safe_float(row[8]) if len(row) > 8 else 0,
            })
        
        return data
    
    def extract_sheet7_data(self) -> List[Dict]:
        """Извлечение данных из листа 7 (Справка)"""
        if not self.metadata:
            return []
        
        sheet7 = self.metadata.sheets_data.get('7-Справка')
        if sheet7 is None or sheet7.empty:
            return []
        
        data = []
        # Поиск начала таблицы №9
        table_start = None
        for idx, row in sheet7.iterrows():
            cell_value = self._safe_str(row[0])
            if 'Таблица №9' in cell_value and 'Комментарии' in cell_value:
                table_start = idx + 3  # Пропускаем заголовок
                break
        
        if table_start is None:
            return []
        
        # Извлечение строк таблицы
        for idx in range(table_start, min(table_start + 10, len(sheet7))):
            row = sheet7.iloc[idx]
            
            # Проверка на конец таблицы
            cell_value = self._safe_str(row[0])
            if pd.isna(row[0]) or not cell_value:
                break
            
            data.append({
                'fuel_type': self._safe_str(row[0]),
                'situation': self._safe_str(row[1]) if len(row) > 1 else '',
                'comments': self._safe_str(row[2]) if len(row) > 2 else '',
            })
        
        return data