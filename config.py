# config.py
import os
from dotenv import load_dotenv

# Загружаем переменные окружения
load_dotenv()

basedir = os.path.abspath(os.path.dirname(__file__))

class Config:
    # Безопасность
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'dev-secret-key-change-in-production'
    
    # База данных
    SQLALCHEMY_DATABASE_URI = os.environ.get('DATABASE_URL') or \
        f"sqlite:///{os.path.join(basedir, 'fuel_reports.db')}"
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    
    # Настройки загрузки файлов
    UPLOAD_FOLDER = os.path.join(basedir, 'uploads')
    REPORTS_FOLDER = os.path.join(basedir, 'reports_output')
    MAX_CONTENT_LENGTH = 100 * 1024 * 1024  # 100MB
    
    # Создаем директории если их нет
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(REPORTS_FOLDER, exist_ok=True)