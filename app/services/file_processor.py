# app/services/file_processor.py
import traceback
from datetime import datetime
from database.queries import db
from parser.unified_parser import UnifiedParser

class FileProcessor:
    def __init__(self):
        self.parsers = self._get_available_parsers()
    
    def _get_available_parsers(self):
        """Получение списка доступных парсеров"""
        parsers = []
        
        # Всегда используем UnifiedParser как основной
        parsers.append({
            'name': 'UnifiedParser',
            'class': UnifiedParser,
            'priority': 1
        })
        
        # Проверяем другие парсеры
        try:
            from parser.simple_all_parser_fixed_v2 import SimpleAllParserV2
            parsers.append({
                'name': 'SimpleAllParserV2',
                'class': SimpleAllParserV2,
                'priority': 2
            })
        except ImportError:
            pass
            
        try:
            from parser.simple_all_parser import SimpleAllParser
            parsers.append({
                'name': 'SimpleAllParser',
                'class': SimpleAllParser,
                'priority': 3
            })
        except ImportError:
            pass
            
        # Сортируем по приоритету
        parsers.sort(key=lambda x: x['priority'])
        return parsers
    
    def process_file(self, filename, file_path):
        """Обработка файла с использованием доступных парсеров"""
        print(f"\n=== НАЧАЛО ОБРАБОТКИ ФАЙЛА: {filename} ===")
        print(f"Файл сохранен: {file_path}")
        
        # Пробуем все доступные парсеры по порядку
        for parser_info in self.parsers:
            try:
                print(f"Пробуем использовать парсер: {parser_info['name']}...")
                result = self._process_with_parser(
                    parser_info['class'], 
                    filename, 
                    file_path,
                    parser_info['name']
                )
                return result
            except Exception as e:
                print(f"Парсер {parser_info['name']} не сработал: {e}")
                continue
        
        # Если ни один парсер не сработал
        return {
            'error': 'Ни один из парсеров не смог обработать файл',
            'success': False
        }
    
    def _process_with_parser(self, parser_class, filename, file_path, parser_name):
        """Обработка файла конкретным парсером"""
        parser = parser_class(file_path)
        all_data = parser.parse_all()
        
        metadata = all_data['metadata']
        
        print(f"\n{parser_name} результаты:")
        print(f"  Компания: {metadata['company']}")
        print(f"  Лист 1: {len(all_data.get('sheet1', []))} записей")
        print(f"  Лист 2: {len(all_data.get('sheet2', []))} записей")
        print(f"  Лист 3: {len(all_data.get('sheet3', []))} записей")
        print(f"  Лист 4: {len(all_data.get('sheet4', []))} записей")
        print(f"  Лист 5: {len(all_data.get('sheet5', []))} записей")
        print(f"  Лист 6: {len(all_data.get('sheet6', []))} записей")
        print(f"  Лист 7: {len(all_data.get('sheet7', []))} записей")
        
        # Сохраняем в БД информацию о файле
        file_id, company_id = db.save_uploaded_file(
            filename=filename,
            file_path=file_path,
            company_name=metadata['company'],
            report_date=metadata['report_date'].date()
        )
        
        print(f"Файл сохранен в БД: ID={file_id}, Company ID={company_id}")
        
        # Сохраняем все данные
        saved_counts = self._save_all_data(all_data, file_id, company_id, metadata)
        
        # Обновляем статус файла
        db.update_file_status(file_id, 'processed')
        print(f"✓ Статус файла обновлен на 'processed'")
        
        print(f"=== ЗАВЕРШЕНО ОБРАБОТКА ФАЙЛА: {filename} ===\n")
        
        return {
            'success': True,
            'message': f'Файл успешно обработан ({parser_name})',
            'company': metadata['company'],
            'report_date': metadata['report_date'].strftime('%Y-%m-%d'),
            'data_extracted': {
                'sheet1': len(all_data.get('sheet1', [])),
                'sheet2': len(all_data.get('sheet2', [])),
                'sheet3': len(all_data.get('sheet3', [])),
                'sheet4': len(all_data.get('sheet4', [])),
                'sheet5': len(all_data.get('sheet5', [])),
                'sheet6': len(all_data.get('sheet6', [])),
                'sheet7': len(all_data.get('sheet7', [])),
            },
            'data_saved': saved_counts,
            'file_info': {
                'file_id': file_id,
                'company_id': company_id,
                'filename': filename
            }
        }
    
    def _save_all_data(self, all_data, file_id, company_id, metadata):
        """Сохранение всех данных из парсера"""
        saved_counts = {}
        
        # Сохраняем данные из каждого листа
        save_functions = {
            'sheet1': (db.save_sheet1_data, 'Sheet1 сохранен'),
            'sheet2': (db.save_sheet2_data, 'Sheet2 сохранен'),
            'sheet3': (db.save_sheet3_data, 'Sheet3 сохранен'),
            'sheet4': (db.save_sheet4_data, 'Sheet4 сохранен'),
            'sheet5': (db.save_sheet5_data, 'Sheet5 сохранен'),
            'sheet6': (db.save_sheet6_data, 'Sheet6 сохранен'),
            'sheet7': (db.save_sheet7_data, 'Sheet7 сохранен'),
        }
        
        for sheet_key, (save_func, success_msg) in save_functions.items():
            if all_data.get(sheet_key):
                try:
                    data = all_data[sheet_key]
                    if sheet_key == 'sheet2':
                        # Для sheet2 передаем dict, а не список
                        save_func(file_id, company_id, metadata['report_date'].date(), data)
                        saved_counts[sheet_key] = 1
                    else:
                        # Для остальных листов передаем список
                        save_func(file_id, company_id, metadata['report_date'].date(), data)
                        saved_counts[sheet_key] = len(data)
                    print(f"✓ {success_msg}: {saved_counts[sheet_key]} записей")
                except Exception as e:
                    print(f"✗ Ошибка сохранения {sheet_key}: {e}")
                    saved_counts[sheet_key] = 0
                    traceback.print_exc()
        
        return saved_counts