# parser/simple_all_parser.py - ПОЛНАЯ ИСПРАВЛЕННАЯ ВЕРСИЯ
import pandas as pd
from datetime import datetime
import re
from typing import Dict, List, Any

class SimpleAllParser:
    """Упрощенный парсер для всех листов"""
    
    def __init__(self, file_path: str):
        self.file_path = file_path
    
    def parse_all(self) -> dict:
        """Простой парсинг всех данных"""
        try:
            print(f"Парсинг файла: {self.file_path}")
            
            result = {
                'metadata': self._parse_metadata(),
                'sheet1': self._parse_sheet1_simple(),
                'sheet2': self._parse_sheet2_simple(),
                'sheet3': self._parse_sheet3_simple(),
                'sheet4': self._parse_sheet4_simple(),
                'sheet5': self._parse_sheet5_simple(),
                'sheet6': self._parse_sheet6_simple(),
                'sheet7': self._parse_sheet7_simple(),
            }
            
            return result
            
        except Exception as e:
            print(f"Ошибка парсинга: {e}")
            raise
    
    def _parse_metadata(self) -> dict:
        """Парсинг метаданных"""
        try:
            # Читаем лист 1 для метаданных
            df = pd.read_excel(self.file_path, sheet_name='1-Структура', 
                             header=None, nrows=10, engine='openpyxl')
            
            metadata = {
                'company': 'Неизвестная компания',
                'report_date': datetime.now(),
                'executor': '',
                'phone': ''
            }
            
            # Ищем данные в первых строках
            for i in range(min(10, len(df))):
                row = df.iloc[i]
                if len(row) > 1:
                    cell_value = str(row[0])
                    
                    if 'Информация по состоянию за:' in cell_value:
                        try:
                            metadata['report_date'] = pd.to_datetime(row[1])
                        except:
                            pass
                    
                    elif 'Исполнитель:' in cell_value:
                        metadata['executor'] = str(row[1]) if len(row) > 1 else ''
                    
                    elif 'Контактный телефон:' in cell_value:
                        metadata['phone'] = str(row[1]) if len(row) > 1 else ''
            
            # Определяем компанию по имени файла
            import os
            filename = os.path.basename(self.file_path).lower()
            
            # Обновленные паттерны для определения компании
            company_patterns = [
                ('саханефтегазсбыт', 'Саханефтегазсбыт'),
                ('саха', 'Саханефтегазсбыт'),  # Добавил сокращенный вариант
                ('туймаада', 'Туймаада-Нефть'),
                ('сибойл', 'Сибойл'),
                ('экто-ойл', 'ЭКТО-Ойл'),
                ('эктоойл', 'ЭКТО-Ойл'),
                ('сибирское', 'Сибирское топливо'),
                ('паритет', 'Паритет'),
            ]
            
            for pattern, company in company_patterns:
                if pattern in filename:
                    metadata['company'] = company
                    break
            
            return metadata
            
        except Exception as e:
            print(f"Ошибка парсинга метаданных: {e}")
            return {'company': 'Неизвестная компания', 'report_date': datetime.now()}
    
    def _parse_sheet1_simple(self) -> list:
        """Простой парсинг листа 1"""
        try:
            df = pd.read_excel(self.file_path, sheet_name='1-Структура', 
                             header=None, engine='openpyxl')
            
            data = []
            # Ищем таблицу с данными
            for i in range(len(df)):
                row = df.iloc[i]
                # Ищем строки где есть числовые данные в колонках 2-4
                if len(row) >= 5:
                    try:
                        # Проверяем что это данные, а не заголовки
                        oil_depots = pd.to_numeric(row[2], errors='coerce')
                        azs = pd.to_numeric(row[3], errors='coerce')
                        working_azs = pd.to_numeric(row[4], errors='coerce')
                        
                        if not pd.isna(oil_depots) or not pd.isna(azs):
                            # Проверяем что это не заголовок таблицы
                            if str(row[0]).strip() not in ['1', '2', 'Принадлежность']:
                                data.append({
                                    'affiliation': str(row[0]) if not pd.isna(row[0]) else '',
                                    'company_name': str(row[1]) if len(row) > 1 and not pd.isna(row[1]) else '',
                                    'oil_depots_count': int(oil_depots) if not pd.isna(oil_depots) else 0,
                                    'azs_count': int(azs) if not pd.isna(azs) else 0,
                                    'working_azs_count': int(working_azs) if not pd.isna(working_azs) else 0
                                })
                    except:
                        continue
            
            return data[:50]  # Ограничиваем 50 записями
            
        except Exception as e:
            print(f"Ошибка парсинга листа 1: {e}")
            return []
    
    def _parse_sheet2_simple(self) -> dict:
        """Простой парсинг листа 2"""
        try:
            df = pd.read_excel(self.file_path, sheet_name='2-Потребность', 
                             header=None, engine='openpyxl')
            
            data = {'year': datetime.now().year, 'month': ''}
            
            # Ищем числовые данные в определенных позициях
            for i in range(len(df)):
                row = df.iloc[i]
                cell_str = str(row[0]) if len(row) > 0 and not pd.isna(row[0]) else ''
                
                # Ищем годовую потребность
                if cell_str and ('ГОД' in cell_str.upper() or 'год' in cell_str.lower()):
                    # Извлекаем год
                    year_match = re.search(r'(\d{4})', cell_str)
                    if year_match:
                        data['year'] = int(year_match.group(1))
                    
                    # Парсим данные
                    if len(row) >= 11:
                        for col_idx, key in enumerate([
                            'gasoline_total', 'gasoline_ai76_80', 'gasoline_ai92', 
                            'gasoline_ai95', 'gasoline_ai98_100', 'diesel_total',
                            'diesel_winter', 'diesel_arctic', 'diesel_summer', 
                            'diesel_intermediate'
                        ]):
                            if col_idx < len(row) - 1:
                                try:
                                    data[key] = float(pd.to_numeric(row[col_idx + 1], errors='coerce') or 0)
                                except:
                                    data[key] = 0
                    break
            
            # Ищем месячную потребность
            for i in range(len(df)):
                row = df.iloc[i]
                cell_str = str(row[0]) if len(row) > 0 and not pd.isna(row[0]) else ''
                
                # Ищем месячную потребность
                if cell_str and ('МЕСЯЦ' in cell_str.upper() or 'месяц' in cell_str.lower()):
                    # Извлекаем месяц
                    month_match = re.search(r'(\w+)(?=,|$)', cell_str)
                    if month_match:
                        data['month'] = month_match.group(1).strip()
                    
                    # Парсим данные
                    if len(row) >= 11:
                        for col_idx, key in enumerate([
                            'monthly_gasoline_total', 'monthly_gasoline_ai76_80', 
                            'monthly_gasoline_ai92', 'monthly_gasoline_ai95', 
                            'monthly_gasoline_ai98_100', 'monthly_diesel_total',
                            'monthly_diesel_winter', 'monthly_diesel_arctic', 
                            'monthly_diesel_summer', 'monthly_diesel_intermediate'
                        ]):
                            if col_idx < len(row) - 1:
                                try:
                                    data[key] = float(pd.to_numeric(row[col_idx + 1], errors='coerce') or 0)
                                except:
                                    data[key] = 0
                    break
            
            return data
            
        except Exception as e:
            print(f"Ошибка парсинга листа 2: {e}")
            return {}
    
    def _parse_sheet3_simple(self) -> List[Dict]:
        """Простой парсинг листа 3"""
        try:
            df = pd.read_excel(self.file_path, sheet_name='3-Остатки', 
                             header=None, engine='openpyxl')
            
            data = []
            # Упрощенный парсинг - ищем строки с данными
            for i in range(len(df)):
                row = df.iloc[i]
                
                # Пропускаем пустые строки и заголовки
                if len(row) < 3 or pd.isna(row[0]) or str(row[0]).strip() == '':
                    continue
                
                # Проверяем что это может быть строка с данми
                try:
                    # Проверяем наличие числовых данных в колонках 3-10
                    has_numbers = False
                    for col in range(3, min(11, len(row))):
                        val = pd.to_numeric(row[col], errors='coerce')
                        if not pd.isna(val) and val != 0:
                            has_numbers = True
                            break
                    
                    if has_numbers:
                        record = {
                            'affiliation': str(row[0]) if not pd.isna(row[0]) else '',
                            'company_name': str(row[1]) if len(row) > 1 and not pd.isna(row[1]) else '',
                            'location_type': 'Нефтебаза' if 'НБ' in str(row[2]) else 'АЗС',
                            'location_name': str(row[2]) if len(row) > 2 and not pd.isna(row[2]) else '',
                        }
                        
                        # Добавляем основные числовые данные
                        if len(row) >= 11:
                            record.update({
                                'stock_ai92': float(pd.to_numeric(row[4], errors='coerce') or 0),
                                'stock_ai95': float(pd.to_numeric(row[5], errors='coerce') or 0),
                                'stock_diesel_winter': float(pd.to_numeric(row[7], errors='coerce') or 0),
                                'stock_diesel_arctic': float(pd.to_numeric(row[8], errors='coerce') or 0),
                            })
                        
                        data.append(record)
                except:
                    continue
            
            return data[:100]  # Ограничиваем 100 записями
            
        except Exception as e:
            print(f"Ошибка парсинга листа 3: {e}")
            return []
    
    def _parse_sheet4_simple(self) -> List[Dict]:
        """Простой парсинг листа 4"""
        try:
            df = pd.read_excel(self.file_path, sheet_name='4-Поставка', 
                             header=None, engine='openpyxl')
            
            data = []
            # Упрощенный парсинг
            for i in range(len(df)):
                row = df.iloc[i]
                
                # Пропускаем пустые строки
                if len(row) < 4 or pd.isna(row[0]) or str(row[0]).strip() == '':
                    continue
                
                # Проверяем что это может быть строка с данными
                try:
                    # Проверяем наличие даты в колонке 3
                    has_date = False
                    if len(row) > 3:
                        try:
                            date_val = pd.to_datetime(row[3], errors='coerce')
                            has_date = not pd.isna(date_val)
                        except:
                            pass
                    
                    # Или проверяем наличие числовых данных
                    has_numbers = False
                    for col in range(4, min(12, len(row))):
                        val = pd.to_numeric(row[col], errors='coerce')
                        if not pd.isna(val) and val != 0:
                            has_numbers = True
                            break
                    
                    if has_date or has_numbers:
                        record = {
                            'affiliation': str(row[0]) if not pd.isna(row[0]) else '',
                            'company_name': str(row[1]) if len(row) > 1 and not pd.isna(row[1]) else '',
                            'oil_depot_name': str(row[2]) if len(row) > 2 and not pd.isna(row[2]) else '',
                        }
                        
                        # Парсим дату
                        if len(row) > 3:
                            try:
                                record['supply_date'] = pd.to_datetime(row[3]).date()
                            except:
                                record['supply_date'] = None
                        
                        # Добавляем основные числовые данные
                        if len(row) >= 12:
                            record.update({
                                'supply_ai92': float(pd.to_numeric(row[5], errors='coerce') or 0),
                                'supply_ai95': float(pd.to_numeric(row[6], errors='coerce') or 0),
                                'supply_diesel_winter': float(pd.to_numeric(row[8], errors='coerce') or 0),
                                'supply_diesel_arctic': float(pd.to_numeric(row[9], errors='coerce') or 0),
                            })
                        
                        data.append(record)
                except:
                    continue
            
            return data[:50]  # Ограничиваем 50 записями
            
        except Exception as e:
            print(f"Ошибка парсинга листа 4: {e}")
            return []
    
    def _parse_sheet5_simple(self) -> List[Dict]:
        """Простой парсинг листа 5"""
        try:
            df = pd.read_excel(self.file_path, sheet_name='5-Реализация', 
                             header=None, engine='openpyxl')
            
            data = []
            # Упрощенный парсинг
            for i in range(len(df)):
                row = df.iloc[i]
                
                # Пропускаем пустые строки
                if len(row) < 3 or pd.isna(row[0]) or str(row[0]).strip() == '':
                    continue
                
                # Проверяем что это может быть строка с данными
                try:
                    # Проверяем наличие числовых данных
                    has_numbers = False
                    for col in range(3, min(19, len(row))):
                        val = pd.to_numeric(row[col], errors='coerce')
                        if not pd.isna(val) and val != 0:
                            has_numbers = True
                            break
                    
                    if has_numbers:
                        record = {
                            'affiliation': str(row[0]) if not pd.isna(row[0]) else '',
                            'company_name': str(row[1]) if len(row) > 1 and not pd.isna(row[1]) else '',
                            'location_type': 'Нефтебаза' if 'НБ' in str(row[2]) else 'АЗС',
                            'location_name': str(row[2]) if len(row) > 2 and not pd.isna(row[2]) else '',
                        }
                        
                        # Добавляем основные числовые данные (реализация с начала месяца)
                        if len(row) >= 19:
                            record.update({
                                'monthly_ai92': float(pd.to_numeric(row[12], errors='coerce') or 0),
                                'monthly_ai95': float(pd.to_numeric(row[13], errors='coerce') or 0),
                                'monthly_diesel_winter': float(pd.to_numeric(row[15], errors='coerce') or 0),
                                'monthly_diesel_arctic': float(pd.to_numeric(row[16], errors='coerce') or 0),
                            })
                        
                        data.append(record)
                except:
                    continue
            
            return data[:100]  # Ограничиваем 100 записями
            
        except Exception as e:
            print(f"Ошибка парсинга листа 5: {e}")
            return []
    
    def _parse_sheet6_simple(self) -> List[Dict]:
        """Простой парсинг листа 6"""
        try:
            df = pd.read_excel(self.file_path, sheet_name='6-Авиатопливо', 
                             header=None, engine='openpyxl')
            
            data = []
            # Упрощенный парсинг - ищем строки с названиями аэропортов
            for i in range(len(df)):
                row = df.iloc[i]
                
                # Ищем строки, начинающиеся с названий аэропортов
                if len(row) > 0 and not pd.isna(row[0]):
                    cell_str = str(row[0]).strip()
                    # Проверяем что это название аэропорта, а не заголовок
                    if cell_str and not any(x in cell_str for x in ['Таблица', 'Наименование', '1', '2']):
                        record = {
                            'airport_name': cell_str,
                            'tzk_name': str(row[1]) if len(row) > 1 and not pd.isna(row[1]) else '',
                        }
                        
                        # Пытаемся добавить числовые данные
                        for col_idx, key in enumerate(['supply_week', 'supply_month_start', 
                                                     'monthly_demand', 'consumption_week',
                                                     'consumption_month_start', 'end_of_day_balance']):
                            if col_idx + 2 < len(row):
                                try:
                                    record[key] = float(pd.to_numeric(row[col_idx + 2], errors='coerce') or 0)
                                except:
                                    record[key] = 0
                        
                        data.append(record)
            
            return data[:50]  # Ограничиваем 50 записями
            
        except Exception as e:
            print(f"Ошибка парсинга листа 6: {e}")
            return []
    
    def _parse_sheet7_simple(self) -> List[Dict]:
        """Простой парсинг листа 7"""
        try:
            df = pd.read_excel(self.file_path, sheet_name='7-Справка', 
                             header=None, engine='openpyxl')
            
            data = []
            # Ищем строки с типами топлива
            fuel_types = ['Бензин', 'Дизельное', 'Авиатопливо', 'бензин', 'дизельное', 'авиатопливо']
            
            for i in range(len(df)):
                row = df.iloc[i]
                
                if len(row) > 0 and not pd.isna(row[0]):
                    cell_str = str(row[0]).strip()
                    # Проверяем что это тип топлива
                    if any(ft in cell_str for ft in fuel_types):
                        record = {
                            'fuel_type': cell_str,
                            'situation': str(row[1]) if len(row) > 1 and not pd.isna(row[1]) else '',
                            'comments': str(row[2]) if len(row) > 2 and not pd.isna(row[2]) else '',
                        }
                        data.append(record)
            
            return data[:10]  # Ограничиваем 10 записями
            
        except Exception as e:
            print(f"Ошибка парсинга листа 7: {e}")
            return []