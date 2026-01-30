# parser/fuel_report_parser_fixed.py
import pandas as pd
import numpy as np
from datetime import datetime
import re
from typing import Dict, List, Any, Optional
import warnings
warnings.filterwarnings('ignore')

class FixedFuelReportParser:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.xls = None
        self.sheets_data = {}
        
    def parse_all(self) -> Dict[str, Any]:
        """Парсинг всех листов файла"""
        try:
            print(f"Парсинг файла: {self.file_path}")
            
            # Читаем Excel файл
            self.xls = pd.ExcelFile(self.file_path, engine='openpyxl')
            
            # Извлекаем метаданные
            metadata = self._extract_metadata()
            
            # Парсим все листы
            all_data = {
                'metadata': metadata,
                'sheet1': self._parse_sheet1(),
                'sheet2': self._parse_sheet2(),
                'sheet3': self._parse_sheet3(),
                'sheet4': self._parse_sheet4(),
                'sheet5': self._parse_sheet5(),
                'sheet6': self._parse_sheet6(),
                'sheet7': self._parse_sheet7(),
            }
            
            return all_data
            
        except Exception as e:
            print(f"Ошибка парсинга файла: {e}")
            raise
    
    def _extract_metadata(self) -> Dict[str, Any]:
        """Извлечение метаданных из файла"""
        try:
            # Читаем первые строки листа 1
            df = pd.read_excel(self.file_path, sheet_name='1-Структура', 
                             header=None, nrows=20, engine='openpyxl')
            
            metadata = {
                'filename': self.file_path.split('/')[-1],
                'company_name': self._detect_company_name(),
                'report_date': datetime.now(),
                'executor': '',
                'phone': '',
                'region': ''
            }
            
            # Ищем ключевые строки
            for i in range(len(df)):
                row = df.iloc[i]
                if len(row) > 0:
                    cell_str = self._safe_str(row[0])
                    
                    if 'Информация по состоянию за:' in cell_str and len(row) > 1:
                        metadata['report_date'] = self._parse_date(row[1])
                    
                    elif 'Исполнитель:' in cell_str and len(row) > 1:
                        metadata['executor'] = self._safe_str(row[1])
                    
                    elif 'Контактный телефон:' in cell_str and len(row) > 1:
                        metadata['phone'] = self._safe_str(row[1])
                    
                    elif 'Субъект Российской Федерации' in cell_str and len(row) > 1:
                        metadata['region'] = self._safe_str(row[1])
            
            return metadata
            
        except Exception as e:
            print(f"Ошибка извлечения метаданных: {e}")
            return {'company_name': 'Неизвестная компания', 'report_date': datetime.now()}
    
    def _detect_company_name(self) -> str:
        """Определение названия компании"""
        filename = self.file_path.split('/')[-1].lower()
        
        # Сначала пробуем определить по имени файла
        company_patterns = [
            ('саханефтегазсбыт', 'Саханефтегазсбыт'),
            ('туймаада', 'Туймаада-Нефть'),
            ('сибойл', 'Сибойл'),
            ('экто-ойл', 'ЭКТО-Ойл'),
            ('эктоойл', 'ЭКТО-Ойл'),
            ('сибирское', 'Сибирское топливо'),
            ('паритет', 'Паритет'),
        ]
        
        for pattern, company in company_patterns:
            if pattern in filename:
                return company
        
        # Если не нашли в имени файла, смотрим в содержимом
        try:
            df = pd.read_excel(self.file_path, sheet_name='1-Структура', 
                             header=None, nrows=50, engine='openpyxl')
            
            for i in range(len(df)):
                row = df.iloc[i]
                for cell in row:
                    cell_str = self._safe_str(cell)
                    for pattern, company in company_patterns:
                        if pattern in cell_str.lower():
                            return company
        except:
            pass
        
        return 'Неизвестная компания'
    
    def _parse_sheet1(self) -> List[Dict]:
        """Парсинг листа 1: Структура"""
        try:
            # Читаем лист, пропуская первые строки
            df = pd.read_excel(self.file_path, sheet_name='1-Структура', 
                             header=None, engine='openpyxl')
            
            data = []
            found_table = False
            
            for i in range(len(df)):
                row = df.iloc[i]
                row0_str = self._safe_str(row[0])
                
                # Ищем начало таблицы
                if 'Таблица №1' in row0_str:
                    found_table = True
                    continue
                
                if found_table and '1' in row0_str and '2' in self._safe_str(row[1]) if len(row) > 1 else '':
                    # Это заголовки столбцов, пропускаем
                    continue
                
                if found_table and row0_str and row0_str != '1' and row0_str != '2':
                    # Проверяем что это строка с данными (есть числа в колонках 2-4)
                    has_numbers = False
                    if len(row) >= 5:
                        try:
                            col2 = pd.to_numeric(row[2], errors='coerce')
                            col3 = pd.to_numeric(row[3], errors='coerce')
                            col4 = pd.to_numeric(row[4], errors='coerce')
                            has_numbers = not (pd.isna(col2) and pd.isna(col3))
                        except:
                            pass
                    
                    if has_numbers and len(row) >= 5:
                        data.append({
                            'affiliation': self._safe_str(row[0]),
                            'company_name': self._safe_str(row[1]),
                            'oil_depots_count': self._safe_int(row[2]),
                            'azs_count': self._safe_int(row[3]),
                            'working_azs_count': self._safe_int(row[4])
                        })
                    elif 'Таблица №2' in row0_str:
                        # Конец первой таблицы
                        break
            
            return data
            
        except Exception as e:
            print(f"Ошибка парсинга листа 1: {e}")
            return []
    
    def _parse_sheet2(self) -> Dict[str, Any]:
        """Парсинг листа 2: Потребность"""
        try:
            df = pd.read_excel(self.file_path, sheet_name='2-Потребность', 
                             header=None, engine='openpyxl')
            
            data = {'year': datetime.now().year, 'month': ''}
            
            # Ищем годовую потребность
            for i in range(len(df)):
                row = df.iloc[i]
                cell_str = self._safe_str(row[0])
                
                if cell_str and ('ГОД' in cell_str.upper() or 'год' in cell_str.lower()):
                    # Извлекаем год
                    year_match = re.search(r'(\d{4})', cell_str)
                    if year_match:
                        data['year'] = int(year_match.group(1))
                    
                    # Парсим данные годовой потребности
                    if len(row) >= 11:
                        data.update({
                            'gasoline_total': self._safe_float(row[1]),
                            'gasoline_ai76_80': self._safe_float(row[2]),
                            'gasoline_ai92': self._safe_float(row[3]),
                            'gasoline_ai95': self._safe_float(row[4]),
                            'gasoline_ai98_100': self._safe_float(row[5]),
                            'diesel_total': self._safe_float(row[6]),
                            'diesel_winter': self._safe_float(row[7]),
                            'diesel_arctic': self._safe_float(row[8]),
                            'diesel_summer': self._safe_float(row[9]),
                            'diesel_intermediate': self._safe_float(row[10]),
                        })
                    break
            
            # Ищем месячную потребность
            for i in range(len(df)):
                row = df.iloc[i]
                cell_str = self._safe_str(row[0])
                
                if cell_str and ('МЕСЯЦ' in cell_str.upper() or 'месяц' in cell_str.lower()):
                    # Извлекаем месяц
                    month_match = re.search(r'(\w+)(?=,|$)', cell_str)
                    if month_match:
                        data['month'] = month_match.group(1).strip()
                    
                    # Парсим данные месячной потребности
                    if len(row) >= 11:
                        data.update({
                            'monthly_gasoline_total': self._safe_float(row[1]),
                            'monthly_gasoline_ai76_80': self._safe_float(row[2]),
                            'monthly_gasoline_ai92': self._safe_float(row[3]),
                            'monthly_gasoline_ai95': self._safe_float(row[4]),
                            'monthly_gasoline_ai98_100': self._safe_float(row[5]),
                            'monthly_diesel_total': self._safe_float(row[6]),
                            'monthly_diesel_winter': self._safe_float(row[7]),
                            'monthly_diesel_arctic': self._safe_float(row[8]),
                            'monthly_diesel_summer': self._safe_float(row[9]),
                            'monthly_diesel_intermediate': self._safe_float(row[10]),
                        })
                    break
            
            return data
            
        except Exception as e:
            print(f"Ошибка парсинга листа 2: {e}")
            return {}
    
    def _parse_sheet3(self) -> List[Dict]:
        """Парсинг листа 3: Остатки"""
        try:
            # Лист 3 имеет сложную структуру, читаем как есть
            df = pd.read_excel(self.file_path, sheet_name='3-Остатки', 
                             header=None, engine='openpyxl')
            
            data = []
            found_table = False
            start_row = None
            
            # Ищем начало таблицы
            for i in range(len(df)):
                row = df.iloc[i]
                cell_str = self._safe_str(row[0])
                
                if 'Таблица №5' in cell_str:
                    found_table = True
                    start_row = i + 3  # Пропускаем заголовки
                    break
            
            if not start_row:
                return []
            
            # Парсим строки таблицы
            for i in range(start_row, min(start_row + 100, len(df))):
                row = df.iloc[i]
                
                # Проверяем что это строка с данными
                if len(row) < 3 or pd.isna(row[0]) or self._safe_str(row[0]) == '':
                    continue
                
                # Определяем тип локации
                location_type = 'АЗС'
                if len(row) > 2:
                    location_name = self._safe_str(row[2])
                    if 'НБ' in location_name or 'нефтебаза' in location_name.lower():
                        location_type = 'Нефтебаза'
                    elif 'АЗС' in location_name:
                        location_type = 'АЗС'
                
                # Создаем запись
                record = {
                    'affiliation': self._safe_str(row[0]),
                    'company_name': self._safe_str(row[1]) if len(row) > 1 else '',
                    'location_type': location_type,
                    'location_name': location_name if len(row) > 2 else '',
                }
                
                # Парсим числовые данные с учетом смещения
                # Имеющиеся запасы (колонки 3-10)
                if len(row) >= 11:
                    record.update({
                        'stock_ai76_80': self._safe_float(row[3]),
                        'stock_ai92': self._safe_float(row[4]),
                        'stock_ai95': self._safe_float(row[5]),
                        'stock_ai98_100': self._safe_float(row[6]),
                        'stock_diesel_winter': self._safe_float(row[7]),
                        'stock_diesel_arctic': self._safe_float(row[8]),
                        'stock_diesel_summer': self._safe_float(row[9]),
                        'stock_diesel_intermediate': self._safe_float(row[10]),
                    })
                
                # Товар в пути (колонки 11-18)
                if len(row) >= 19:
                    record.update({
                        'transit_ai76_80': self._safe_float(row[11]),
                        'transit_ai92': self._safe_float(row[12]),
                        'transit_ai95': self._safe_float(row[13]),
                        'transit_ai98_100': self._safe_float(row[14]),
                        'transit_diesel_winter': self._safe_float(row[15]),
                        'transit_diesel_arctic': self._safe_float(row[16]),
                        'transit_diesel_summer': self._safe_float(row[17]),
                        'transit_diesel_intermediate': self._safe_float(row[18]),
                    })
                
                # Емкость хранения (колонки 19-26)
                if len(row) >= 27:
                    record.update({
                        'capacity_ai76_80': self._safe_float(row[19]),
                        'capacity_ai92': self._safe_float(row[20]),
                        'capacity_ai95': self._safe_float(row[21]),
                        'capacity_ai98_100': self._safe_float(row[22]),
                        'capacity_diesel_winter': self._safe_float(row[23]),
                        'capacity_diesel_arctic': self._safe_float(row[24]),
                        'capacity_diesel_summer': self._safe_float(row[25]),
                        'capacity_diesel_intermediate': self._safe_float(row[26]),
                    })
                
                data.append(record)
            
            return data
            
        except Exception as e:
            print(f"Ошибка парсинга листа 3: {e}")
            return []
    
    def _parse_sheet4(self) -> List[Dict]:
        """Парсинг листа 4: Поставка"""
        try:
            df = pd.read_excel(self.file_path, sheet_name='4-Поставка', 
                             header=None, engine='openpyxl')
            
            data = []
            found_table = False
            start_row = None
            
            # Ищем начало таблицы
            for i in range(len(df)):
                row = df.iloc[i]
                cell_str = self._safe_str(row[0])
                
                if 'Таблица №6' in cell_str:
                    found_table = True
                    start_row = i + 3  # Пропускаем заголовки
                    break
            
            if not start_row:
                return []
            
            # Парсим строки таблицы
            for i in range(start_row, min(start_row + 50, len(df))):
                row = df.iloc[i]
                
                # Проверяем что это строка с данными
                if len(row) < 4 or pd.isna(row[0]) or self._safe_str(row[0]) == '':
                    continue
                
                # Парсим дату поставки
                supply_date = None
                if len(row) > 3:
                    date_val = row[3]
                    if isinstance(date_val, datetime):
                        supply_date = date_val.date()
                    elif isinstance(date_val, str):
                        try:
                            supply_date = pd.to_datetime(date_val).date()
                        except:
                            pass
                
                # Создаем запись
                record = {
                    'affiliation': self._safe_str(row[0]),
                    'company_name': self._safe_str(row[1]) if len(row) > 1 else '',
                    'oil_depot_name': self._safe_str(row[2]) if len(row) > 2 else '',
                    'supply_date': supply_date,
                }
                
                # Парсим объемы поставок (колонки 4-11)
                if len(row) >= 12:
                    record.update({
                        'supply_ai76_80': self._safe_float(row[4]),
                        'supply_ai92': self._safe_float(row[5]),
                        'supply_ai95': self._safe_float(row[6]),
                        'supply_ai98_100': self._safe_float(row[7]),
                        'supply_diesel_winter': self._safe_float(row[8]),
                        'supply_diesel_arctic': self._safe_float(row[9]),
                        'supply_diesel_summer': self._safe_float(row[10]),
                        'supply_diesel_intermediate': self._safe_float(row[11]),
                    })
                
                data.append(record)
            
            return data
            
        except Exception as e:
            print(f"Ошибка парсинга листа 4: {e}")
            return []
    
    def _parse_sheet5(self) -> List[Dict]:
        """Парсинг листа 5: Реализация"""
        try:
            df = pd.read_excel(self.file_path, sheet_name='5-Реализация', 
                             header=None, engine='openpyxl')
            
            data = []
            found_table = False
            start_row = None
            
            # Ищем начало таблицы
            for i in range(len(df)):
                row = df.iloc[i]
                cell_str = self._safe_str(row[0])
                
                if 'Таблица №7' in cell_str:
                    found_table = True
                    start_row = i + 3  # Пропускаем заголовки
                    break
            
            if not start_row:
                return []
            
            # Парсим строки таблицы
            for i in range(start_row, min(start_row + 100, len(df))):
                row = df.iloc[i]
                
                # Проверяем что это строка с данными
                if len(row) < 3 or pd.isna(row[0]) or self._safe_str(row[0]) == '':
                    continue
                
                # Определяем тип локации
                location_type = 'АЗС'
                if len(row) > 2:
                    location_name = self._safe_str(row[2])
                    if 'НБ' in location_name:
                        location_type = 'Нефтебаза'
                    elif 'АЗС' in location_name:
                        location_type = 'АЗС'
                
                # Создаем запись
                record = {
                    'affiliation': self._safe_str(row[0]),
                    'company_name': self._safe_str(row[1]) if len(row) > 1 else '',
                    'location_type': location_type,
                    'location_name': location_name,
                }
                
                # Парсим реализацию за сутки (колонки 3-10)
                if len(row) >= 11:
                    record.update({
                        'daily_ai76_80': self._safe_float(row[3]),
                        'daily_ai92': self._safe_float(row[4]),
                        'daily_ai95': self._safe_float(row[5]),
                        'daily_ai98_100': self._safe_float(row[6]),
                        'daily_diesel_winter': self._safe_float(row[7]),
                        'daily_diesel_arctic': self._safe_float(row[8]),
                        'daily_diesel_summer': self._safe_float(row[9]),
                        'daily_diesel_intermediate': self._safe_float(row[10]),
                    })
                
                # Парсим реализацию с начала месяца (колонки 11-18)
                if len(row) >= 19:
                    record.update({
                        'monthly_ai76_80': self._safe_float(row[11]),
                        'monthly_ai92': self._safe_float(row[12]),
                        'monthly_ai95': self._safe_float(row[13]),
                        'monthly_ai98_100': self._safe_float(row[14]),
                        'monthly_diesel_winter': self._safe_float(row[15]),
                        'monthly_diesel_arctic': self._safe_float(row[16]),
                        'monthly_diesel_summer': self._safe_float(row[17]),
                        'monthly_diesel_intermediate': self._safe_float(row[18]),
                    })
                
                data.append(record)
            
            return data
            
        except Exception as e:
            print(f"Ошибка парсинга листа 5: {e}")
            return []
    
    def _parse_sheet6(self) -> List[Dict]:
        """Парсинг листа 6: Авиатопливо"""
        try:
            df = pd.read_excel(self.file_path, sheet_name='6-Авиатопливо', 
                             header=None, engine='openpyxl')
            
            data = []
            found_table = False
            start_row = None
            
            # Ищем начало таблицы
            for i in range(len(df)):
                row = df.iloc[i]
                cell_str = self._safe_str(row[0])
                
                if 'Таблица №8' in cell_str:
                    found_table = True
                    start_row = i + 3  # Пропускаем заголовки
                    break
            
            if not start_row:
                return []
            
            # Парсим строки таблицы
            for i in range(start_row, min(start_row + 50, len(df))):
                row = df.iloc[i]
                
                # Проверяем что это строка с данными
                if len(row) < 2 or pd.isna(row[0]) or self._safe_str(row[0]) == '':
                    continue
                
                # Создаем запись
                record = {
                    'airport_name': self._safe_str(row[0]),
                    'tzk_name': self._safe_str(row[1]) if len(row) > 1 else '',
                    'contracts_info': self._safe_str(row[2]) if len(row) > 2 else '',
                    'supply_week': self._safe_float(row[3]) if len(row) > 3 else 0,
                    'supply_month_start': self._safe_float(row[4]) if len(row) > 4 else 0,
                    'monthly_demand': self._safe_float(row[5]) if len(row) > 5 else 0,
                    'consumption_week': self._safe_float(row[6]) if len(row) > 6 else 0,
                    'consumption_month_start': self._safe_float(row[7]) if len(row) > 7 else 0,
                    'end_of_day_balance': self._safe_float(row[8]) if len(row) > 8 else 0,
                }
                
                data.append(record)
            
            return data
            
        except Exception as e:
            print(f"Ошибка парсинга листа 6: {e}")
            return []
    
    def _parse_sheet7(self) -> List[Dict]:
        """Парсинг листа 7: Справка"""
        try:
            df = pd.read_excel(self.file_path, sheet_name='7-Справка', 
                             header=None, engine='openpyxl')
            
            data = []
            found_table = False
            start_row = None
            
            # Ищем начало таблицы
            for i in range(len(df)):
                row = df.iloc[i]
                cell_str = self._safe_str(row[0])
                
                if 'Таблица №9' in cell_str:
                    found_table = True
                    start_row = i + 3  # Пропускаем заголовки
                    break
            
            if not start_row:
                return []
            
            # Парсим строки таблицы
            for i in range(start_row, min(start_row + 10, len(df))):
                row = df.iloc[i]
                
                # Проверяем что это строка с данными
                if len(row) < 2 or pd.isna(row[0]) or self._safe_str(row[0]) == '':
                    continue
                
                # Создаем запись
                record = {
                    'fuel_type': self._safe_str(row[0]),
                    'situation': self._safe_str(row[1]) if len(row) > 1 else '',
                    'comments': self._safe_str(row[2]) if len(row) > 2 else '',
                }
                
                data.append(record)
            
            return data
            
        except Exception as e:
            print(f"Ошибка парсинга листа 7: {e}")
            return []
    
    # Вспомогательные методы
    def _safe_str(self, value) -> str:
        """Безопасное преобразование в строку"""
        if pd.isna(value) or value is None:
            return ''
        return str(value).strip()
    
    def _safe_int(self, value) -> int:
        """Безопасное преобразование в int"""
        try:
            if pd.isna(value) or value is None:
                return 0
            if isinstance(value, str):
                value = value.replace(',', '.').strip()
                # Извлекаем число из строки
                match = re.search(r'[-+]?\d*\.?\d+', value)
                if match:
                    value = match.group(0)
            return int(float(value))
        except:
            return 0
    
    def _safe_float(self, value) -> float:
        """Безопасное преобразование в float"""
        try:
            if pd.isna(value) or value is None:
                return 0.0
            if isinstance(value, str):
                value = value.replace(',', '.').strip()
                # Извлекаем число из строки
                match = re.search(r'[-+]?\d*\.?\d+', value)
                if match:
                    value = match.group(0)
            return float(value)
        except:
            return 0.0
    
    def _parse_date(self, value) -> datetime:
        """Парсинг даты"""
        try:
            if isinstance(value, datetime):
                return value
            elif isinstance(value, str):
                for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y']:
                    try:
                        return datetime.strptime(value.strip(), fmt)
                    except:
                        continue
        except:
            pass
        return datetime.now()