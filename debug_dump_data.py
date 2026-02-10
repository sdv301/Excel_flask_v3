#!/usr/bin/env python3
"""
Скрипт для создания дампа данных из базы данных
"""

import os
import sys
import json
from datetime import datetime

# Добавляем путь к проекту в sys.path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

from database.queries import DatabaseQueries

def create_data_dump():
    """Создает полный дамп данных из базы данных"""
    print("=== СОЗДАНИЕ ДАМПА ДАННЫХ ИЗ БАЗЫ ДАННЫХ ===")
    
    db = DatabaseQueries()
    
    # Получаем агрегированные данные
    aggregated_data = db.get_aggregated_data()
    
    # Создаем структурированный дамп
    dump = {
        "timestamp": datetime.now().isoformat(),
        "total_companies": len(aggregated_data),
        "companies": {}
    }
    
    for company_name, company_data in aggregated_data.items():
        dump["companies"][company_name] = {
            "metadata": {
                "name": company_name,
                "has_sheet1": len(company_data.get('sheet1', [])) > 0,
                "has_sheet2": bool(company_data.get('sheet2', {})),
                "has_sheet3": len(company_data.get('sheet3_data', [])) > 0,
                "has_sheet4": len(company_data.get('sheet4_data', [])) > 0,
                "has_sheet5": len(company_data.get('sheet5_data', [])) > 0,
            },
            "sheet1_sample": company_data.get('sheet1', [])[:2] if company_data.get('sheet1') else [],
            "sheet2_data": company_data.get('sheet2', {}),
            "sheet3_sample": company_data.get('sheet3_data', [])[:3] if company_data.get('sheet3_data') else [],
            "sheet4_sample": company_data.get('sheet4_data', [])[:2] if company_data.get('sheet4_data') else [],
            "sheet5_sample": company_data.get('sheet5_data', [])[:3] if company_data.get('sheet5_data') else [],
            "sheet3_totals": company_data.get('sheet3_totals', {}),
            "sheet5_totals": company_data.get('sheet5_totals', {})
        }
    
    # Сохраняем дамп в файл
    dump_filename = f"data_dump_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    dump_path = os.path.join(project_root, 'data_dumps', dump_filename)
    
    os.makedirs(os.path.dirname(dump_path), exist_ok=True)
    
    with open(dump_path, 'w', encoding='utf-8') as f:
        json.dump(dump, f, indent=2, ensure_ascii=False, default=str)
    
    print(f"✅ Дамп сохранен: {dump_path}")
    print(f"📊 Обработано компаний: {len(aggregated_data)}")
    
    # Выводим краткую статистику
    print("\n=== КРАТКАЯ СТАТИСТИКА ===")
    for company_name, data in dump["companies"].items():
        print(f"\n🏢 {company_name}:")
        print(f"   Sheet1: {len(data.get('sheet1_sample', []))} записей")
        print(f"   Sheet3: {len(data.get('sheet3_sample', []))} локаций")
        print(f"   Sheet4: {len(data.get('sheet4_sample', []))} поставок")
        print(f"   Sheet5: {len(data.get('sheet5_sample', []))} записей реализации")
        
        if data.get('sheet3_sample'):
            first_location = data['sheet3_sample'][0]
            print(f"   Пример локации: {first_location.get('location_name')}")
            print(f"     Ключи: {list(first_location.keys())}")
    
    return dump_path

def analyze_data_structure(dump_path):
    """Анализирует структуру данных для генератора отчетов"""
    with open(dump_path, 'r', encoding='utf-8') as f:
        dump = json.load(f)
    
    print("\n=== АНАЛИЗ СТРУКТУРЫ ДАННЫХ ===")
    
    key_analysis = {
        'sheet3_keys': set(),
        'sheet4_keys': set(),
        'sheet5_keys': set()
    }
    
    for company_name, company_data in dump['companies'].items():
        # Анализируем sheet3_data
        for record in company_data.get('sheet3_sample', []):
            key_analysis['sheet3_keys'].update(record.keys())
        
        # Анализируем sheet4_data
        for record in company_data.get('sheet4_sample', []):
            key_analysis['sheet4_keys'].update(record.keys())
        
        # Анализируем sheet5_data
        for record in company_data.get('sheet5_sample', []):
            key_analysis['sheet5_keys'].update(record.keys())
    
    print("📋 Sheet3 (Остатки) - ключи:", sorted(key_analysis['sheet3_keys']))
    print("📋 Sheet4 (Поставки) - ключи:", sorted(key_analysis['sheet4_keys']))
    print("📋 Sheet5 (Реализация) - ключи:", sorted(key_analysis['sheet5_keys']))
    
    # Генерируем код для генератора отчетов
    generate_template_code(key_analysis)

def generate_template_code(key_analysis):
    """Генерирует код для генератора отчетов на основе анализа данных"""
    print("\n=== ШАБЛОН КОДА ДЛЯ GENERATOR ===")
    
    # Sheet3 код
    print("\n# Sheet3 (Остатки):")
    for key in sorted(key_analysis['sheet3_keys']):
        if key not in ['location_name']:
            col_map = {
                'stock_ai92': 5, 'stock_ai95': 6, 'stock_diesel_winter': 8, 
                'stock_diesel_arctic': 9, 'stock_diesel_summer': 10,
                'transit_ai92': 13, 'transit_ai95': 14, 'transit_diesel_winter': 16,
                'transit_diesel_arctic': 17, 'capacity_ai92': 21, 'capacity_ai95': 22
            }
            col = col_map.get(key, '?')
            print(f"self._set_cell_value(ws, current_row, {col}, round(location_data.get('{key}', 0), 3))")
    
    # Sheet4 код
    print("\n# Sheet4 (Поставки):")
    for key in sorted(key_analysis['sheet4_keys']):
        if key not in ['oil_depot_name', 'supply_date', 'report_date']:
            col_map = {
                'supply_ai92': 6, 'supply_ai95': 7, 'supply_diesel_winter': 9,
                'supply_diesel_arctic': 10, 'supply_diesel_summer': 11
            }
            col = col_map.get(key, '?')
            print(f"self._set_cell_value(ws, current_row, {col}, round(supply_data.get('{key}', 0), 3))")
    
    # Sheet5 код
    print("\n# Sheet5 (Реализация):")
    for key in sorted(key_analysis['sheet5_keys']):
        if key not in ['location_name']:
            col_map = {
                'daily_ai92': 5, 'daily_ai95': 6, 'daily_winter': 8, 'daily_arctic': 9,
                'monthly_ai92': 13, 'monthly_ai95': 14, 'monthly_winter': 16, 'monthly_arctic': 17
            }
            col = col_map.get(key, '?')
            print(f"self._set_cell_value(ws, current_row, {col}, round(sales_data.get('{key}', 0), 3))")

if __name__ == "__main__":
    dump_path = create_data_dump()
    analyze_data_structure(dump_path)
