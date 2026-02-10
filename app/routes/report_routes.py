# app/routes/report_routes.py
import os
import glob
from flask import Blueprint, request, jsonify, send_file
from reports.template_report_generator import TemplateReportGenerator
from database.queries import DatabaseQueries
from datetime import datetime
import traceback

report_bp = Blueprint('report', __name__)

@report_bp.route('/generate-report', methods=['POST'])
def generate_report():
    """Генерация сводного отчета"""
    try:
        data = request.get_json()
        report_date = data.get('report_date')
        
        if report_date:
            report_date = datetime.strptime(report_date, '%Y-%m-%d').date()
        else:
            report_date = datetime.now().date()
        
        db = DatabaseQueries()
        generator = TemplateReportGenerator(db)
        
        report_path = generator.generate_report(report_date)
        
        if report_path and os.path.exists(report_path):
            filename = os.path.basename(report_path)
            
            # Возвращаем только имя файла, путь будем искать динамически
            return jsonify({
                'success': True,
                'message': 'Отчет успешно сгенерирован',
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

@report_bp.route('/download-report/<filename>')
def download_report(filename):
    """Универсальное скачивание отчета - ищем файл по всему проекту"""
    try:
        print(f"🔍 Поиск файла: {filename}")
        
        # Безопасная обработка имени файла
        if not filename or '..' in filename or '/' in filename:
            return jsonify({'success': False, 'error': 'Некорректное имя файла'}), 400
        
        # 1. Определяем корневую директорию проекта
        current_dir = os.getcwd()
        print(f"📁 Текущая директория: {current_dir}")
        
        # 2. Ищем файл во всех возможных местах
        search_locations = [
            # В текущей директории и поддиректориях
            current_dir,
            # Папка reports_output в корне
            os.path.join(current_dir, 'reports_output'),
            # Папка app/reports_output
            os.path.join(current_dir, 'app', 'reports_output'),
            # На уровень выше
            os.path.dirname(current_dir),
            # На уровень выше + reports_output
            os.path.join(os.path.dirname(current_dir), 'reports_output'),
        ]
        
        # Добавляем стандартные папки проекта
        search_locations.extend([
            'reports_output',
            '../reports_output',
            './reports_output',
            'app/reports_output',
            '../app/reports_output'
        ])
        
        found_path = None
        for location in search_locations:
            if not os.path.exists(location):
                continue
                
            # Ищем файл в этой локации
            potential_path = os.path.join(location, filename)
            if os.path.exists(potential_path):
                found_path = potential_path
                print(f"✅ Файл найден: {found_path}")
                break
                
            # Ищем рекурсивно во всех подпапках
            for root, dirs, files in os.walk(location):
                if filename in files:
                    found_path = os.path.join(root, filename)
                    print(f"✅ Файл найден рекурсивно: {found_path}")
                    break
            if found_path:
                break
        
        if not found_path:
            # Покажем все доступные отчеты для отладки
            print("📊 Доступные файлы отчетов:")
            for location in search_locations:
                if os.path.exists(location):
                    try:
                        files = os.listdir(location)
                        xlsx_files = [f for f in files if f.endswith('.xlsx')]
                        if xlsx_files:
                            print(f"   {location}: {xlsx_files}")
                    except PermissionError:
                        continue
            
            return jsonify({
                'success': False,
                'error': f'Файл {filename} не найден',
                'search_locations': search_locations,
                'current_directory': current_dir
            }), 404
        
        # Скачиваем файл
        return send_file(
            found_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
            
    except Exception as e:
        print(f"❌ Критическая ошибка при скачивании: {e}")
        return jsonify({
            'success': False,
            'error': f'Ошибка при скачивании: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500

@report_bp.route('/list-reports')
def list_reports():
    """Список всех доступных отчетов (для отладки)"""
    try:
        reports = []
        search_dirs = ['reports_output', '../reports_output', 'app/reports_output']
        
        for dir_path in search_dirs:
            if os.path.exists(dir_path):
                for file in os.listdir(dir_path):
                    if file.endswith('.xlsx') and 'сводный' in file.lower():
                        full_path = os.path.join(dir_path, file)
                        stats = os.stat(full_path)
                        reports.append({
                            'filename': file,
                            'path': full_path,
                            'size': stats.st_size,
                            'modified': datetime.fromtimestamp(stats.st_mtime).strftime('%d.%m.%Y %H:%M'),
                            'absolute_path': os.path.abspath(full_path)
                        })
        
        return jsonify({
            'success': True,
            'reports': reports,
            'current_directory': os.getcwd(),
            'absolute_current': os.path.abspath(os.getcwd())
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500
