# parser/simple_all_parser_fixed_v2.py - ИСПРАВЛЕННАЯ ВЕРСИЯ
import pandas as pd
from datetime import datetime
import re
from typing import Dict, List, Any
import os

class SimpleAllParserV2:
    """Улучшенный упрощенный парсер для всех листов"""
    
    def __init__(self, file_path: str):
        self.file_path = file_path
    
    def parse_all(self) -> Dict[str, Any]:
        """Простой парсинг всех данных - УЛУЧШЕННАЯ ВЕРСИЯ"""
        try:
            print(f"Парсинг файла: {self.file_path}")
            
            result = {
                'metadata': self._parse_metadata_v2(),
                'sheet1': self._parse_sheet1_v2(),
                'sheet2': self._parse_sheet2_v2(),
                'sheet3': self._parse_sheet3_v2(),
                'sheet4': self._parse_sheet4_v2(),
                'sheet5': self._parse_sheet5_v2(),
                'sheet6': self._parse_sheet6_v2(),
                'sheet7': self._parse_sheet7_v2(),
            }
            
            print(f"Результаты парсинга:")
            print(f"  Лист 1: {len(result['sheet1'])} записей")
            print(f"  Лист 2: {len(result['sheet2'])} записей")
            print(f"  Лист 3: {len(result['sheet3'])} записей")
            print(f"  Лист 4: {len(result['sheet4'])} записей")
            print(f"  Лист 5: {len(result['sheet5'])} записей")
            print(f"  Лист 6: {len(result['sheet6'])} записей")
            print(f"  Лист 7: {len(result['sheet7'])} записей")
            
            return result
            
        except Exception as e:
            print(f"Ошибка парсинга: {e}")
            import traceback
            traceback.print_exc()
            raise

    def _parse_metadata_v2(self) -> Dict[str, Any]:
        """Упрощенный парсинг метаданных - только по имени файла"""
        try:
            filename = os.path.basename(self.file_path).lower()
            print(f"Определение компании по файлу: {filename}")
            
            # Определяем компанию строго по имени файла
            if 'сибойл' in filename or 'сибирь' in filename:
                company = 'Сибойл'
            elif 'саханефтегазсбыт' in filename or 'снгс' in filename or 'саха' in filename:
                company = 'Саханефтегазсбыт'
            elif 'туймаада' in filename:
                company = 'Туймаада-Нефть'
            elif 'экто' in filename:
                company = 'ЭКТО-Ойл'
            elif 'газпром' in filename:
                company = 'Газпром'
            elif 'роснефть' in filename:
                company = 'Роснефть'
            else:
                company = 'Неизвестная компания'
            
            metadata = {
                'company': company,
                'report_date': datetime.now(),
                'executor': '',
                'phone': '',
                'filename': os.path.basename(self.file_path)
            }
            
            print(f"Компания определена: {company}")
            return metadata
            
        except Exception as e:
            print(f"Ошибка парсинга метаданных: {e}")
            return {'company': 'Неизвестная компания', 'report_date': datetime.now(), 'filename': os.path.basename(self.file_path)}
        
    def _parse_sheet1_v2(self) -> List[Dict]:
        """Улучшенный парсинг листа 1"""
        try:
            print("Парсинг листа 1: Структура...")
            df = pd.read_excel(self.file_path, sheet_name='1-Структура', 
                             header=None, engine='openpyxl')
            
            data = []
            found_table = False
            
            for i in range(len(df)):
                row = df.iloc[i]
                if len(row) > 0:
                    cell_str = str(row[0])
                    
                    # Ищем начало таблицы
                    if 'Таблица №1' in cell_str:
                        found_table = True
                        print("Найдено начало таблицы №1")
                        continue
                    
                    # Пропускаем строки с формулами
                    if found_table and len(row) >= 5 and pd.notna(row[0]) and isinstance(row[0], str) and row[0].startswith('='):
                        continue
                    
                    # Извлекаем данные после начала таблицы
                    if found_table and len(row) >= 5:
                        # Проверяем, что это числовые данные или названия компаний
                        company_name = ''
                        azs_count = 0
                        
                        # Ищем название компании (обычно в колонке B)
                        if len(row) > 1 and pd.notna(row[1]) and isinstance(row[1], str):
                            company_name = str(row[1]).strip()
                        
                        # Ищем количество АЗС (колонка D)
                        if len(row) > 3:
                            azs_count = self._extract_number(row[3])
                        
                        # Если есть название компании или количество АЗС
                        if company_name or azs_count > 0:
                            record = {
                                'affiliation': str(row[0]).strip() if pd.notna(row[0]) else '',
                                'company_name': company_name,
                                'oil_depots_count': self._extract_number(row[2]) if len(row) > 2 else 0,
                                'azs_count': azs_count,
                                'working_azs_count': self._extract_number(row[4]) if len(row) > 4 else 0,
                            }
                            
                            data.append(record)
                            print(f"  Найдена запись: {company_name}, АЗС: {azs_count}")
            
            print(f"Лист 1: найдено {len(data)} записей")
            return data
            
        except Exception as e:
            print(f"Ошибка парсинга листа 1: {e}")
            return []
    
    def _parse_sheet2_v2(self) -> List[Dict]:
        """Улучшенный парсинг листа 2 - ПОТРЕБНОСТЬ"""
        try:
            print("Парсинг листа 2: Потребность...")
            df = pd.read_excel(self.file_path, sheet_name='2-Потребность', 
                            header=None, engine='openpyxl')
            
            data = []
            
            # Ищем строку с "МЕСЯЦ" - месячная потребность
            for i in range(len(df)):
                row = df.iloc[i]
                if len(row) > 0:
                    cell_str = str(row[0]) if pd.notna(row[0]) else ''
                    
                    # Ищем строку с месячной потребностью
                    if 'МЕСЯЦ' in cell_str.upper():
                        print(f"Найдена строка с месячной потребностью в строке {i+1}: {cell_str}")
                        
                        # Пропускаем заголовки таблицы (обычно следующие 1-2 строки)
                        for j in range(i+2, min(i+5, len(df))):  # Ищем в следующих 3 строках
                            data_row = df.iloc[j]
                            if len(data_row) > 10:
                                # Пропускаем строки с формулами
                                if pd.notna(data_row[0]) and isinstance(data_row[0], str) and data_row[0].startswith('='):
                                    continue
                                
                                # Извлекаем числовые данные
                                record = {
                                    'period': str(data_row[0]).strip() if pd.notna(data_row[0]) else '',
                                    'gasoline_total': self._extract_number(data_row[1]),
                                    'gasoline_ai76_80': self._extract_number(data_row[2]),
                                    'gasoline_ai92': self._extract_number(data_row[3]),
                                    'gasoline_ai95': self._extract_number(data_row[4]),
                                    'gasoline_ai98_100': self._extract_number(data_row[5]),
                                    'diesel_total': self._extract_number(data_row[6]),
                                    'diesel_winter': self._extract_number(data_row[7]),
                                    'diesel_arctic': self._extract_number(data_row[8]),
                                    'diesel_summer': self._extract_number(data_row[9]),
                                    'diesel_intermediate': self._extract_number(data_row[10]),
                                }
                                
                                # Проверяем, есть ли числовые данные
                                has_data = any([
                                    record['gasoline_total'] > 0,
                                    record['gasoline_ai92'] > 0,
                                    record['gasoline_ai95'] > 0,
                                    record['diesel_total'] > 0
                                ])
                                
                                if has_data:
                                    data.append(record)
                                    print(f"  Месячная потребность: бензин={record['gasoline_total']}, дизель={record['diesel_total']}")
                                    return data
                        break
            
            print(f"Лист 2: найдено {len(data)} записей")
            return data
            
        except Exception as e:
            print(f"Ошибка парсинга листа 2: {e}")
            import traceback
            traceback.print_exc()
            return []

    def _parse_sheet3_v2(self) -> List[Dict]:
        """Улучшенный парсинг листа 3 - ОСТАТКИ"""
        try:
            print("Парсинг листа 3: Остатки...")
            df = pd.read_excel(self.file_path, sheet_name='3-Остатки', 
                             header=None, engine='openpyxl')
            
            data = []
            found_table = False
            
            for i in range(len(df)):
                row = df.iloc[i]
                if len(row) > 0:
                    cell_str = str(row[0]) if pd.notna(row[0]) else ''
                    
                    # Ищем начало таблицы №5
                    if 'Таблица №5' in cell_str:
                        found_table = True
                        print("Найдено начало таблицы №5 (Остатки)")
                        continue
                    
                    # Пропускаем строки с формулами
                    if found_table and pd.notna(row[0]) and isinstance(row[0], str) and row[0].startswith('='):
                        continue
                    
                    # Извлекаем данные начиная с строки 9
                    if found_table and i >= 8:  # Начинаем с 9 строки (индекс 8)
                        # Пропускаем пустые строки
                        if all(pd.isna(cell) for cell in row[:3]):
                            continue
                        
                        # Пропускаем строки заголовков
                        if pd.notna(row[0]) and '1' in str(row[0]) and pd.notna(row[1]) and '2' in str(row[1]):
                            continue
                        
                        # Извлекаем данные компании
                        affiliation = str(row[0]).strip() if pd.notna(row[0]) else ''
                        company_name = str(row[1]).strip() if len(row) > 1 and pd.notna(row[1]) else ''
                        
                        # Если нет названия компании, но есть принадлежность
                        if not company_name and affiliation:
                            company_name = affiliation
                        
                        # Пропускаем строки без названия компании
                        if not company_name or len(company_name) < 3:
                            continue
                        
                        # Извлекаем числовые данные
                        record = {
                            'affiliation': affiliation,
                            'company_name': company_name,
                            'location_name': str(row[2]).strip() if len(row) > 2 and pd.notna(row[2]) else '',
                            'stock_ai76_80': self._extract_number(row[3]) if len(row) > 3 else 0,
                            'stock_ai92': self._extract_number(row[4]) if len(row) > 4 else 0,
                            'stock_ai95': self._extract_number(row[5]) if len(row) > 5 else 0,
                            'stock_ai98_100': self._extract_number(row[6]) if len(row) > 6 else 0,
                            'stock_diesel_winter': self._extract_number(row[7]) if len(row) > 7 else 0,
                            'stock_diesel_arctic': self._extract_number(row[8]) if len(row) > 8 else 0,
                            'stock_diesel_summer': self._extract_number(row[9]) if len(row) > 9 else 0,
                            'stock_diesel_intermediate': self._extract_number(row[10]) if len(row) > 10 else 0,
                        }
                        
                        # Проверяем, есть ли числовые данные
                        has_data = any([
                            record['stock_ai92'] > 0,
                            record['stock_ai95'] > 0,
                            record['stock_diesel_winter'] > 0,
                            record['stock_diesel_arctic'] > 0
                        ])
                        
                        if has_data:
                            data.append(record)
                            print(f"  Остатки: {company_name}, AI92={record['stock_ai92']}, AI95={record['stock_ai95']}")
            
            print(f"Лист 3: найдено {len(data)} записей")
            return data
            
        except Exception as e:
            print(f"Ошибка парсинга листа 3: {e}")
            import traceback
            traceback.print_exc()
            return []

    def _parse_sheet4_v2(self) -> List[Dict]:
        """Улучшенный парсинг листа 4 - ПОСТАВКИ"""
        try:
            print("Парсинг листа 4: Поставка...")
            df = pd.read_excel(self.file_path, sheet_name='4-Поставка', 
                             header=None, engine='openpyxl')
            
            data = []
            found_table = False
            
            for i in range(len(df)):
                row = df.iloc[i]
                if len(row) > 0:
                    cell_str = str(row[0]) if pd.notna(row[0]) else ''
                    
                    # Ищем начало таблицы №6
                    if 'Таблица №6' in cell_str:
                        found_table = True
                        print("Найдено начало таблицы №6 (Поставки)")
                        continue
                    
                    # Пропускаем строки с формулами
                    if found_table and pd.notna(row[0]) and isinstance(row[0], str) and row[0].startswith('='):
                        continue
                    
                    # Извлекаем данные начиная с строки 8
                    if found_table and i >= 7:  # Начинаем с 8 строки (индекс 7)
                        # Пропускаем пустые строки
                        if all(pd.isna(cell) for cell in row[:3]):
                            continue
                        
                        # Пропускаем строки заголовков
                        if pd.notna(row[0]) and '1' in str(row[0]) and pd.notna(row[1]) and '2' in str(row[1]):
                            continue
                        
                        # Извлекаем данные компании
                        affiliation = str(row[0]).strip() if pd.notna(row[0]) else ''
                        company_name = str(row[1]).strip() if len(row) > 1 and pd.notna(row[1]) else ''
                        
                        # Если нет названия компании, но есть поставщик в колонке 2
                        if not company_name and len(row) > 2 and pd.notna(row[2]):
                            company_name = str(row[2]).strip()
                        
                        # Пропускаем строки без названия компании
                        if not company_name or len(company_name) < 3:
                            continue
                        
                        # Извлекаем числовые данные
                        record = {
                            'affiliation': affiliation,
                            'company_name': company_name,
                            'oil_depot_name': str(row[2]).strip() if len(row) > 2 and pd.notna(row[2]) else '',
                            'supply_date_text': str(row[3]).strip() if len(row) > 3 and pd.notna(row[3]) else '',
                            'supply_ai76_80': self._extract_number(row[4]) if len(row) > 4 else 0,
                            'supply_ai92': self._extract_number(row[5]) if len(row) > 5 else 0,
                            'supply_ai95': self._extract_number(row[6]) if len(row) > 6 else 0,
                            'supply_ai98_100': self._extract_number(row[7]) if len(row) > 7 else 0,
                            'supply_diesel_winter': self._extract_number(row[8]) if len(row) > 8 else 0,
                            'supply_diesel_arctic': self._extract_number(row[9]) if len(row) > 9 else 0,
                            'supply_diesel_summer': self._extract_number(row[10]) if len(row) > 10 else 0,
                            'supply_diesel_intermediate': self._extract_number(row[11]) if len(row) > 11 else 0,
                        }
                        
                        # Проверяем, есть ли числовые данные
                        has_data = any([
                            record['supply_ai92'] > 0,
                            record['supply_ai95'] > 0,
                            record['supply_diesel_winter'] > 0,
                            record['supply_diesel_arctic'] > 0
                        ])
                        
                        if has_data:
                            data.append(record)
                            print(f"  Поставка: {company_name}, AI92={record['supply_ai92']}, AI95={record['supply_ai95']}")
            
            print(f"Лист 4: найдено {len(data)} записей")
            return data
            
        except Exception as e:
            print(f"Ошибка парсинга листа 4: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def _parse_sheet5_v2(self) -> List[Dict]:
        """Улучшенный парсинг листа 5 - РЕАЛИЗАЦИЯ"""
        try:
            print("Парсинг листа 5: Реализация...")
            df = pd.read_excel(self.file_path, sheet_name='5-Реализация', 
                             header=None, engine='openpyxl')
            
            data = []
            found_table = False
            
            for i in range(len(df)):
                row = df.iloc[i]
                if len(row) > 0:
                    cell_str = str(row[0]) if pd.notna(row[0]) else ''
                    
                    # Ищем начало таблицы №7
                    if 'Таблица №7' in cell_str:
                        found_table = True
                        print("Найдено начало таблицы №7 (Реализация)")
                        continue
                    
                    # Пропускаем строки с формулами
                    if found_table and pd.notna(row[0]) and isinstance(row[0], str) and row[0].startswith('='):
                        continue
                    
                    # Извлекаем данные начиная с строки 9
                    if found_table and i >= 8:  # Начинаем с 9 строки (индекс 8)
                        # Пропускаем пустые строки
                        if all(pd.isna(cell) for cell in row[:3]):
                            continue
                        
                        # Пропускаем строки заголовков
                        if pd.notna(row[0]) and '1' in str(row[0]) and pd.notna(row[1]) and '2' in str(row[1]):
                            continue
                        
                        # Извлекаем данные компании
                        affiliation = str(row[0]).strip() if pd.notna(row[0]) else ''
                        company_name = str(row[1]).strip() if len(row) > 1 and pd.notna(row[1]) else ''
                        
                        # Пропускаем строки без названия компании
                        if not company_name or len(company_name) < 3:
                            continue
                        
                        # Извлекаем числовые данные (суточная реализация)
                        record = {
                            'affiliation': affiliation,
                            'company_name': company_name,
                            'location_name': str(row[2]).strip() if len(row) > 2 and pd.notna(row[2]) else '',
                            'daily_ai76_80': self._extract_number(row[3]) if len(row) > 3 else 0,
                            'daily_ai92': self._extract_number(row[4]) if len(row) > 4 else 0,
                            'daily_ai95': self._extract_number(row[5]) if len(row) > 5 else 0,
                            'daily_ai98_100': self._extract_number(row[6]) if len(row) > 6 else 0,
                            'daily_diesel_winter': self._extract_number(row[7]) if len(row) > 7 else 0,
                            'daily_diesel_arctic': self._extract_number(row[8]) if len(row) > 8 else 0,
                            'daily_diesel_summer': self._extract_number(row[9]) if len(row) > 9 else 0,
                            'daily_diesel_intermediate': self._extract_number(row[10]) if len(row) > 10 else 0,
                        }
                        
                        # Проверяем, есть ли числовые данные
                        has_data = any([
                            record['daily_ai92'] > 0,
                            record['daily_ai95'] > 0,
                            record['daily_diesel_winter'] > 0,
                            record['daily_diesel_arctic'] > 0
                        ])
                        
                        if has_data:
                            data.append(record)
                            print(f"  Реализация: {company_name}, AI92={record['daily_ai92']}, AI95={record['daily_ai95']}")
            
            print(f"Лист 5: найдено {len(data)} записей")
            return data
            
        except Exception as e:
            print(f"Ошибка парсинга листа 5: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def _parse_sheet6_v2(self) -> List[Dict]:
        """Улучшенный парсинг листа 6 - Авиатопливо"""
        try:
            print("Парсинг листа 6: Авиатопливо...")
            df = pd.read_excel(self.file_path, sheet_name='6-Авиатопливо', 
                             header=None, engine='openpyxl')
            
            data = []
            found_table = False
            
            for i in range(len(df)):
                row = df.iloc[i]
                if len(row) > 0:
                    cell_str = str(row[0]) if pd.notna(row[0]) else ''
                    
                    # Ищем начало таблицы №8
                    if 'Таблица №8' in cell_str:
                        found_table = True
                        print("Найдено начало таблицы №8 (Авиатопливо)")
                        continue
                    
                    # Извлекаем данные
                    if found_table and len(row) >= 2:
                        # Пропускаем пустые строки
                        if all(pd.isna(cell) for cell in row[:2]):
                            continue
                        
                        record = {
                            'airport_name': str(row[0]) if pd.notna(row[0]) else '',
                            'tzk_name': str(row[1]) if len(row) > 1 and pd.notna(row[1]) else '',
                            'contracts_info': str(row[2]) if len(row) > 2 and pd.notna(row[2]) else '',
                            'supply_week': self._extract_number(row[3]) if len(row) > 3 else 0,
                            'supply_month_start': self._extract_number(row[4]) if len(row) > 4 else 0,
                            'monthly_demand': self._extract_number(row[5]) if len(row) > 5 else 0,
                            'consumption_week': self._extract_number(row[6]) if len(row) > 6 else 0,
                            'consumption_month_start': self._extract_number(row[7]) if len(row) > 7 else 0,
                            'end_of_day_balance': self._extract_number(row[8]) if len(row) > 8 else 0,
                        }
                        
                        if record['airport_name']:
                            data.append(record)
            
            print(f"Лист 6: найдено {len(data)} записей")
            return data
            
        except Exception as e:
            print(f"Ошибка парсинга листа 6: {e}")
            return []
    
    def _parse_sheet7_v2(self) -> List[Dict]:
        """Улучшенный парсинг листа 7 - Справка"""
        try:
            print("Парсинг листа 7: Справка...")
            df = pd.read_excel(self.file_path, sheet_name='7-Справка', 
                             header=None, engine='openpyxl')
            
            data = []
            found_table = False
            
            for i in range(len(df)):
                row = df.iloc[i]
                if len(row) > 0:
                    cell_str = str(row[0]) if pd.notna(row[0]) else ''
                    
                    # Ищем начало таблицы №9
                    if 'Таблица №9' in cell_str:
                        found_table = True
                        print("Найдено начало таблицы №9 (Справка)")
                        continue
                    
                    # Извлекаем данные
                    if found_table and len(row) >= 3:
                        # Пропускаем пустые строки
                        if all(pd.isna(cell) for cell in row[:3]):
                            continue
                        
                        # Пропускаем заголовки
                        if pd.notna(row[0]) and '1' in str(row[0]):
                            continue
                        
                        record = {
                            'fuel_type': str(row[0]).strip() if pd.notna(row[0]) else '',
                            'situation': str(row[1]).strip() if len(row) > 1 and pd.notna(row[1]) else '',
                            'comments': str(row[2]).strip() if len(row) > 2 and pd.notna(row[2]) else '',
                        }
                        
                        if record['fuel_type']:
                            data.append(record)
            
            print(f"Лист 7: найдено {len(data)} записей")
            return data
            
        except Exception as e:
            print(f"Ошибка парсинга листа 7: {e}")
            return []
    
    def _extract_number(self, value) -> float:
        """Извлекает число из значения, игнорируя формулы"""
        try:
            if pd.isna(value) or value is None:
                return 0.0
            
            # Если это строка, проверяем не формула ли это
            if isinstance(value, str):
                # Пропускаем формулы
                if value.startswith('='):
                    return 0.0
                
                # Извлекаем число из строки
                value = value.replace(',', '.').strip()
                match = re.search(r'[-+]?\d*\.?\d+', value)
                if match:
                    return float(match.group(0))
                else:
                    return 0.0
            
            # Если это уже число
            return float(value)
            
        except Exception as e:
            return 0.0