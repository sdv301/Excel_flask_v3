# app/routes/api_routes.py
from flask import Blueprint, jsonify
from database.queries import db
from database.models import UploadedFile, Company  # Импортируем модели напрямую
from database.connection import db_connection  # Импортируем соединение с БД

api_bp = Blueprint('api', __name__)

@api_bp.route('/api/recent-files')
def api_recent_files():
    """API для получения последних файлов"""
    try:
        files = db.get_recent_files(limit=20)
        return jsonify(files)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@api_bp.route('/api/companies')
def api_companies():
    """API для получения списка компаний"""
    try:
        companies = db.get_companies()
        return jsonify([{
            'id': c.id,
            'name': c.name,
            'code': c.code
        } for c in companies])
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@api_bp.route('/api/file-details/<int:file_id>')
def api_file_details(file_id):
    """API для получения деталей файла"""
    try:
        # Используем прямое соединение с БД вместо db.db
        session = db_connection.get_session()
        file = session.query(UploadedFile).filter(UploadedFile.id == file_id).first()
        
        if file:
            company = session.query(Company).filter(Company.id == file.company_id).first()
            return jsonify({
                'id': file.id,
                'filename': file.filename,
                'company_name': company.name if company else 'Неизвестно',
                'report_date': file.report_date.strftime('%d.%m.%Y') if file.report_date else 'Н/Д',
                'status': file.status,
                'upload_date': file.upload_date.strftime('%d.%m.%Y %H:%M') if file.upload_date else 'Н/Д'
            })
        else:
            return jsonify({'error': 'Файл не найден'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        db_connection.close_session()

@api_bp.route('/api/stats')
def api_stats():
    """API для получения статистики системы"""
    try:
        session = db_connection.get_session()
        
        total_files = session.query(UploadedFile).count()
        processed_files = session.query(UploadedFile).filter(UploadedFile.status == 'processed').count()
        total_companies = session.query(Company).filter(Company.is_active == True).count()
        
        # Последний загруженный файл
        last_file = session.query(UploadedFile).order_by(UploadedFile.upload_date.desc()).first()
        
        return jsonify({
            'total_files': total_files,
            'processed_files': processed_files,
            'total_companies': total_companies,
            'last_upload': last_file.upload_date.strftime('%d.%m.%Y %H:%M') if last_file else 'Нет данных'
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        db_connection.close_session()
