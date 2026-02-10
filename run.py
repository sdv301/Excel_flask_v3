# run.py
from app import create_app
import os

app = create_app()

if __name__ == '__main__':
    # print("=" * 50)
    # print("Система отчетов по топливообеспечению")
    # print("=" * 50)
    # print("Доступные эндпоинты:")
    # print("  GET  /              - Главная страница")
    # print("  POST /upload        - Загрузка файла")
    # print("  POST /generate-report - Генерация отчета")
    # print("  GET  /download-report/<filename> - Скачивание отчета")
    # print("  GET  /api/recent-files - API последних файлов")
    # print("  GET  /api/companies - API списка компаний")
    # print("=" * 50)
    
    # Создаем необходимые папки
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('reports_output', exist_ok=True)
    os.makedirs('report_templates', exist_ok=True)
    
    app.run(debug=True, host='0.0.0.0', port=5000)