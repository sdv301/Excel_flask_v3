# reprocess_files.py
import os
import sys

# Добавляем текущую директорию в путь поиска модулей
sys.path.append(os.path.abspath(os.path.dirname(__file__)))

from app.services.file_processor import FileProcessor

def reprocess():
    processor = FileProcessor()
    upload_folder = 'uploads'
    
    if not os.path.exists(upload_folder):
        print(f"Папка {upload_folder} не найдена.")
        return

    files = [f for f in os.listdir(upload_folder) if f.endswith(('.xlsx', '.xls'))]
    
    if not files:
        print("В папке uploads нет файлов для обработки.")
        return

    print(f"Найдено файлов для обработки: {len(files)}")
    
    for filename in files:
        file_path = os.path.join(upload_folder, filename)
        print(f"\nОбработка {filename}...")
        result = processor.process_file(filename, file_path)
        if result.get('success'):
            print(f"✅ Успешно: {result['message']}")
            print(f"   Данные: {result['data_extracted']}")
        else:
            print(f"❌ Ошибка: {result.get('error')}")

if __name__ == "__main__":
    reprocess()
