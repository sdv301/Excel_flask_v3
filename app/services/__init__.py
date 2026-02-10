# app/services/__init__.py

# Импортируем основные сервисы
from .report_generator import ReportGenerator
from .file_processor import FileProcessor

# Экспорт всех сервисов
__all__ = ['ReportGenerator', 'FileProcessor']