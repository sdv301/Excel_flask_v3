# test_parser.py
import os
from parser.unified_parser import UnifiedParser

def test_parser():
    """Тестируем парсер на реальном файле"""
    try:
        # Используем правильный путь для Windows
        test_file = os.path.join("uploads", "FORMA_OTCHETNOSTI_01.02.2026_.xlsx")
        
        print(f"🔍 Тестируем файл: {test_file}")
        print(f"📁 Файл существует: {os.path.exists(test_file)}")
        
        if not os.path.exists(test_file):
            print("❌ Файл не найден! Проверьте путь.")
            # Попробуем найти любой Excel файл в папке uploads
            uploads_dir = "uploads"
            if os.path.exists(uploads_dir):
                files = [f for f in os.listdir(uploads_dir) if f.endswith('.xlsx')]
                if files:
                    test_file = os.path.join(uploads_dir, files[0])
                    print(f"🔍 Найден файл: {test_file}")
                else:
                    print("❌ В папке uploads нет Excel файлов")
                    return
        
        parser = UnifiedParser(test_file)
        result = parser.parse_all()
        
        print("\n" + "="*50)
        print("📊 РЕЗУЛЬТАТЫ ПАРСИНГА")
        print("="*50)
        
        print(f"🏢 Компания: {result['metadata']['company']}")
        print(f"📅 Дата отчета: {result['metadata']['report_date']}")
        print(f"📋 Доступные листы: {result['metadata']['sheets_available']}")
        
        print(f"\n📈 ДАННЫЕ ПО ЛИСТАМ:")
        print(f"   Лист 1 (Структура): {len(result['sheet1'])} записей")
        print(f"   Лист 2 (Потребность): {len(result['sheet2'])} показателей")
        print(f"   Лист 3 (Остатки): {len(result['sheet3'])} записей")
        print(f"   Лист 4 (Поставки): {len(result['sheet4'])} записей")
        print(f"   Лист 5 (Реализация): {len(result['sheet5'])} записей")
        print(f"   Лист 6 (Авиатопливо): {len(result['sheet6'])} записей")
        print(f"   Лист 7 (Справка): {len(result['sheet7'])} записей")
        
        # Покажем примеры данных
        if result['sheet3']:
            print(f"\n📋 ПЕРВЫЕ 3 ЗАПИСИ ЛИСТА 3 (Остатки):")
            for i, record in enumerate(result['sheet3'][:3]):
                print(f"   {i+1}. Компания: {record.get('company', 'N/A')}")
                print(f"      Объект: {record.get('object_name', 'N/A')}")
                print(f"      АИ-92: {record.get('ai92', 0)}")
                print(f"      АИ-95: {record.get('ai95', 0)}")
                print(f"      АИ-98/100: {record.get('ai98_100', 0)}")
                print(f"      Дизель зимний {record.get('diesel_winter', 0)}")
                print(f"      Дизель арктик: {record.get('diesel_arctic', 0)}")
                print(f"      Дизель летнее: {record.get('diesel_summe', 0)}")
                
                print(f"      АИ-92: {record.get('capacity_ai92', 0)}")
                print(f"      АИ-95: {record.get('capacity_ai95', 0)}")
                print(f"      АИ-98/100: {record.get('capacity_ai98_100', 0)}")
                print(f"      Дизель зимний {record.get('capacity_diesel_winter', 0)}")
                print(f"      Дизель арктик: {record.get('capacity_diesel_arctic', 0)}")
                print(f"      Дизель летнее: {record.get('capacity_diesel_summe', 0)}")
                print()
                
        if result['sheet5']:
            print(f"📋 ПЕРВЫЕ 3 ЗАПИСИ ЛИСТА 5 (Реализация):")
            for i, record in enumerate(result['sheet5'][:3]):
                print(f"   {i+1}. Компания: {record.get('company', 'N/A')}")
                print(f"      Реализация месяц АИ-92: {record.get('monthly_ai92', 0)}")
                print(f"      Реализация месяц АИ-95: {record.get('monthly_ai95', 0)}")
                print(f"      Реализация месяц Дизель зимнее: {record.get('monthly_winter', 0)}")
                print(f"      Реализация месяц Дизель Артик: {record.get('monthly_arctic', 0)}")
                
                print(f"      Реализация день АИ-92: {record.get('daily_ai92', 0)}")
                print(f"      Реализация день АИ-95: {record.get('daily_ai95', 0)}")
                print(f"      Реализация день Дизель зимнее: {record.get('daily_winter', 0)}")
                print(f"      Реализация день Дизель Артик: {record.get('daily_arctic', 0)}")
                print()


            
    except Exception as e:
        print(f"❌ Ошибка тестирования: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_parser()
