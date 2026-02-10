# app/routes/__init__.py
from .main_routes import main_bp
from .upload_routes import upload_bp
from .report_routes import report_bp
from .api_routes import api_bp
from .admin_routes import admin_bp

# Экспорт всех blueprint'ов
__all__ = ['main_bp', 'upload_bp', 'report_bp', 'api_bp', 'admin_bp', ]