# app/services/report_generator.py - ИСПРАВЛЕННАЯ ВЕРСИЯ
# Убраны лишние аргументы при вызове generate_report
# Теперь оба метода (summary и template) используют один и тот же генератор
# aggregated_data всегда берётся внутри генератора (самые свежие данные)

from datetime import datetime, date
from flask import jsonify, request
import os
import traceback
from database.queries import db
from reports.template_report_generator import TemplateReportGenerator  # Ваш полный генератор

class ReportGenerator:

    def __init__(self):
        self.db = db
    
    def generate_summary_report(self, request):
        """Генерация сводного отчёта (старый эндпоинт)"""
        report_date = self._get_report_date_from_request(request)
        
        print(f"\n=== ГЕНЕРАЦИЯ СВОДНОГО ОТЧЁТА ===")
        print(f"Дата отчёта: {report_date}")
        
        # Используем тот же генератор, что и для шаблона
        generator = TemplateReportGenerator(self.db)
        
        # УБРАН aggregated_data из аргументов — генератор сам его получит
        try:
            report_path = generator.generate_report(report_date)
            report_filename = os.path.basename(report_path)
            
            if request.method == 'GET':
                return self._render_report_html(report_filename, report_date)
            else:
                return self._render_report_json(report_filename, report_path)
                
        except Exception as e:
            print(f"Ошибка генерации: {e}")
            traceback.print_exc()
            return self._handle_error(request, str(e))
    
    def generate_template_report(self, request):
        """Генерация отчёта по шаблону (новый/основной эндпоинт)"""
        report_date = self._get_report_date_from_request(request)
        
        print(f"\n=== ГЕНЕРАЦИЯ ОТЧЁТА ПО ШАБЛОНУ ===")
        print(f"Дата отчёта: {report_date}")
        
        generator = TemplateReportGenerator(self.db)
        
        try:
            report_path = generator.generate_report(report_date)
            report_filename = os.path.basename(report_path)
            
            if request.method == 'GET':
                return self._render_template_report_html(report_filename, report_date)
            else:
                return self._render_template_report_json(report_filename, report_path)
                
        except Exception as e:
            print(f"Ошибка генерации шаблонного отчёта: {e}")
            traceback.print_exc()
            return self._handle_error(request, str(e))
    
    # ===================================================================
    # Вспомогательные методы (без изменений)
    # ===================================================================
    def _get_report_date_from_request(self, request) -> date:
        if request.is_json:
            data = request.get_json()
            date_str = data.get('report_date')
        else:
            date_str = request.args.get('report_date') or request.form.get('report_date')
        
        if date_str:
            try:
                return datetime.strptime(date_str, '%Y-%m-%d').date()
            except:
                pass
        return datetime.now().date()
    
    def _handle_error(self, request, message: str):
        error_details = traceback.format_exc()
        if request.method == 'GET':
            html = f"""
            <!DOCTYPE html>
            <html>
            <body>
                <h1>Ошибка генерации отчёта</h1>
                <p style="color: red;">{message}</p>
                <pre>{error_details}</pre>
                <a href="/">← На главную</a>
            </body>
            </html>
            """
            return html, 500
        else:
            return jsonify({'error': message, 'details': error_details}), 500
    
    # HTML/JSON рендеры (оставьте как были или упростите)
    def _render_report_html(self, filename, report_date):
        return f"""
        <h1>Сводный отчёт сгенерирован</h1>
        <p>Файл: {filename}</p>
        <p>Дата: {report_date.strftime('%d.%m.%Y')}</p>
        <a href="/download-report/{filename}">Скачать</a>
        """
    
    def _render_template_report_html(self, filename, report_date):
        return f"""
        <h1>Отчёт по шаблону сгенерирован</h1>
        <p>Файл: {filename}</p>
        <p>Дата: {report_date.strftime('%d.%m.%Y')}</p>
        <a class="btn btn-download" href="/download-report/{filename}">Скачать Excel</a>
        """
    
    def _render_report_json(self, filename, report_path):
        return jsonify({
            'success': True,
            'filename': filename,
            'download_url': f'/download-report/{filename}',
            'message': 'Сводный отчёт готов'
        })
    
    def _render_template_report_json(self, filename, report_path):
        return jsonify({
            'success': True,
            'filename': filename,
            'download_url': f'/download-report/{filename}',
            'message': 'Отчёт по шаблону готов'
        })