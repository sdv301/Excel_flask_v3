# app/routes/upload_routes.py
from flask import Blueprint, request, jsonify, current_app
from werkzeug.utils import secure_filename
import os
import traceback
from app.services.file_processor import FileProcessor

upload_bp = Blueprint('upload', __name__)

@upload_bp.route('/upload', methods=['POST'])
def upload_file():
    """Загрузка файла"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Файл не выбран'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'Файл не выбран'}), 400
        
        if not file.filename.lower().endswith('.xlsx'):
            return jsonify({'error': 'Только Excel файлы (.xlsx)'}), 400
        
        # Сохраняем файл
        filename = secure_filename(file.filename)
        file_path = os.path.join(current_app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # Обрабатываем файл
        processor = FileProcessor()
        result = processor.process_file(filename, file_path)
        
        return jsonify(result)
        
    except Exception as e:
        error_details = traceback.format_exc()
        print(f"Ошибка при загрузке файла: {error_details}")
        return jsonify({'error': str(e), 'details': error_details}), 500