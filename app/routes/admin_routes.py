from flask import Blueprint, render_template, jsonify, request, send_file
from database.queries import DatabaseQueries
from parser.unified_parser import UnifiedParser
from reports.template_report_generator import TemplateReportGenerator
from database.models import UploadedFile, Company  # Добавляем импорт моделей
from database.connection import db_connection  # Добавляем импорт соединения
import os
import traceback
from datetime import datetime

admin_bp = Blueprint('admin', __name__)

@admin_bp.route('/admin')
def admin_dashboard():
    """Главная страница админ-панели"""
    db = DatabaseQueries()
    recent_files = db.get_recent_files(limit=10)
    companies = db.get_companies()
    
    return render_template('admin.html', 
                         recent_files=recent_files, 
                         companies=companies,
                         now=datetime.now())

@admin_bp.route('/admin/test-parse')
def test_parse():
    """Тестирование парсера"""
    try:
        # Ищем последний загруженный файл для тестирования
        db = DatabaseQueries()
        recent_files = db.get_recent_files(limit=1)
        
        if not recent_files:
            return jsonify({
                'success': False,
                'error': 'Нет загруженных файлов для тестирования'
            })
        
        file_info = recent_files[0]
        file_path = f"uploads/{file_info['filename']}"
        
        if not os.path.exists(file_path):
            return jsonify({
                'success': False,
                'error': f'Файл не найден: {file_path}'
            })
        
        # Парсим файл
        parser = UnifiedParser(file_path)
        result = parser.parse_all()
        
        return jsonify({
            'success': True,
            'filename': file_info['filename'],
            'company': result.get('metadata', {}).get('company'),
            'data_extracted': {
                'sheet1': len(result.get('sheet1', [])),
                'sheet2': len(result.get('sheet2', {})),
                'sheet3': len(result.get('sheet3', [])),
                'sheet4': len(result.get('sheet4', [])),
                'sheet5': len(result.get('sheet5', [])),
                'sheet6': len(result.get('sheet6', [])),
                'sheet7': len(result.get('sheet7', []))
            }
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e),
            'details': traceback.format_exc()
        })

@admin_bp.route('/admin/check-db-data')
def check_db_data():
    """Проверка данных в базе"""
    try:
        db = DatabaseQueries()
        summary = db.get_all_data_summary()
        
        return jsonify({
            'success': True,
            'summary': summary
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })

@admin_bp.route('/admin/generate-from-existing')
def generate_from_existing():
    """Создание отчета с существующими данными из базы"""
    try:
        db = DatabaseQueries()
        generator = TemplateReportGenerator(db)
        
        # Генерируем отчет
        report_path = generator.generate_report()
        
        if report_path and os.path.exists(report_path):
            filename = os.path.basename(report_path)
            return jsonify({
                'success': True,
                'message': 'Отчет успешно создан из данных базы',
                'filename': filename,
                'download_url': f'/download-report/{filename}'
            })
        else:
            return jsonify({
                'success': False,
                'error': 'Не удалось создать отчет'
            })
            
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e),
            'details': traceback.format_exc()
        })

@admin_bp.route('/admin/debug-template')
def debug_template():
    """Отладка структуры шаблона"""
    try:
        generator = TemplateReportGenerator(None)
        generator.debug_template_structure()
        
        return jsonify({
            'success': True,
            'message': 'Структура шаблона выведена в консоль сервера'
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })

@admin_bp.route('/admin/system-status')
def system_status():
    """Статус системы"""
    try:
        # Используем прямое соединение с БД
        session = db_connection.get_session()
        total_files = session.query(UploadedFile).count()
        total_companies = session.query(Company).count()
        processed_files = session.query(UploadedFile).filter(
            UploadedFile.status == 'processed'
        ).count()
        
        # Проверяем наличие шаблона
        template_path = 'report_templates/Сводный_отчет_шаблон.xlsx'
        template_exists = os.path.exists(template_path)
        
        # Проверяем папку загрузок
        uploads_dir = 'uploads'
        uploads_exists = os.path.exists(uploads_dir)
        uploads_files = len(os.listdir(uploads_dir)) if uploads_exists else 0
        
        return jsonify({
            'success': True,
            'status': {
                'database': {
                    'files_total': total_files,
                    'files_processed': processed_files,
                    'companies': total_companies
                },
                'filesystem': {
                    'template_exists': template_exists,
                    'uploads_exists': uploads_exists,
                    'uploads_files': uploads_files
                },
                'timestamp': datetime.now().isoformat()
            }
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })
    finally:
        db_connection.close_session()