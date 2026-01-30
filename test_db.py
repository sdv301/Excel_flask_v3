# test_db.py - тест сохранения данных в БД
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from datetime import datetime
from database.connection import db_connection
from database.models import *
from database.queries import DatabaseQueries

# Тестируем сохранение данных
def test_save_data():
    print("=== ТЕСТ СОХРАНЕНИЯ ДАННЫХ В БД ===")
    
    # Создаем сессию
    session = db_connection.get_session()
    
    try:
        # 1. Проверяем, есть ли компании в БД
        companies = session.query(Company).all()
        print(f"Компаний в БД: {len(companies)}")
        for c in companies:
            print(f"  - {c.name} (ID: {c.id})")
        
        # 2. Проверяем, есть ли загруженные файлы
        files = session.query(UploadedFile).all()
        print(f"\nЗагруженных файлов в БД: {len(files)}")
        for f in files:
            print(f"  - {f.filename} (ID: {f.id}, Компания ID: {f.company_id}, Дата: {f.report_date})")
        
        # 3. Проверяем данные из sheet1
        sheet1_data = session.query(Sheet1Structure).all()
        print(f"\nДанных в Sheet1Structure: {len(sheet1_data)}")
        for i, d in enumerate(sheet1_data[:5]):  # Показываем первые 5 записей
            print(f"  [{i+1}] Компания: {d.company_name}, АЗС: {d.azs_count}, Файл ID: {d.file_id}")
        
        # 4. Проверяем данные из sheet3
        sheet3_data = session.query(Sheet3Balance).all()
        print(f"\nДанных в Sheet3Balance: {len(sheet3_data)}")
        for i, d in enumerate(sheet3_data[:5]):
            print(f"  [{i+1}] Компания: {d.company_name}, AI92: {d.stock_ai92}, Файл ID: {d.file_id}")
        
        # 5. Проверяем данные из sheet5
        sheet5_data = session.query(Sheet5Sales).all()
        print(f"\nДанных в Sheet5Sales: {len(sheet5_data)}")
        for i, d in enumerate(sheet5_data[:5]):
            print(f"  [{i+1}] Компания: {d.company_name}, AI92 (мес): {d.monthly_ai92}, Файл ID: {d.file_id}")
        
        # 6. Тестируем get_aggregated_data
        print("\n=== ТЕСТ get_aggregated_data ===")
        db_queries = DatabaseQueries()
        today = datetime.now().date()
        print(f"Запрашиваем данные на дату: {today}")
        
        aggregated = db_queries.get_aggregated_data(today)
        print(f"Найдено компаний: {len(aggregated)}")
        
        for company_name, data in aggregated.items():
            print(f"\nКомпания: {company_name}")
            print(f"  Sheet1 записей: {len(data.get('sheet1', []))}")
            print(f"  Sheet2 записей: {len(data.get('sheet2_data', []))}")
            print(f"  Sheet3 записей: {len(data.get('sheet3_data', []))}")
            print(f"  Sheet4 записей: {len(data.get('sheet4_data', []))}")
            print(f"  Sheet5 записей: {len(data.get('sheet5_data', []))}")
            
    finally:
        db_connection.close_session()

if __name__ == "__main__":
    test_save_data()