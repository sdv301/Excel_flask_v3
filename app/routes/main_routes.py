# app/routes/main_routes.py
from flask import Blueprint, render_template
from datetime import datetime
from database.queries import db

main_bp = Blueprint('main', __name__)

@main_bp.route('/')
def index():
    """Главная страница"""
    try:
        companies = db.get_companies()
        recent_files = db.get_recent_files(limit=10)
        return render_template('index.html', 
                             companies=companies, 
                             recent_files=recent_files,
                             now=datetime.now())
    except Exception as e:
        import traceback
        return f"Ошибка: {str(e)}<br><pre>{traceback.format_exc()}</pre>"

@main_bp.route('/admin')
def admin():
    """Страница администратора"""
    try:
        companies = db.get_companies()
        recent_files = db.get_recent_files(limit=10)
        return render_template('admin.html', 
                             companies=companies, 
                             recent_files=recent_files,
                             now=datetime.now())
    except Exception as e:
        return f"Ошибка: {str(e)}"
    
@main_bp.route('/debug-paths')
def debug_paths():
    return render_template('debug_paths.html')
