# app/__init__.py
from flask import Flask
from config import Config
from database.connection import db_connection

def create_app():
    app = Flask(__name__, 
                template_folder='../templates',
                static_folder='../static')
    app.config.from_object(Config)
    
    # Инициализация базы данных
    init_database(app)
    
    # Регистрация маршрутов
    register_blueprints(app)
    
    return app

def init_database(app):
    """Инициализация базы данных"""
    with app.app_context():
        try:
            db_connection.create_tables()
            print("Таблицы базы данных созданы успешно")
            
            # Добавляем тестовые компании если их нет
            from database.queries import db
            session = db_connection.get_session()
            from database.models import Company
            
            existing = session.query(Company).count()
            if existing == 0:
                test_companies = [
                    ("Саханефтегазсбыт", "СНГС"),
                    ("Туймаада-Нефть", "ТУЙМААДА"),
                    ("Сибойл", "СИБОЙЛ"),
                    ("ЭКТО-Ойл", "ЭКТО"),
                    ("Сибирское топливо", "СИБТОП"),
                    ("Паритет", "ПАРИТЕТ")
                ]
                
                for name, code in test_companies:
                    company = Company(name=name, code=code)
                    session.add(company)
                
                session.commit()
                print("Тестовые компании добавлены")
        except Exception as e:
            print(f"Ошибка при инициализации БД: {e}")
        finally:
            db_connection.close_session()

def register_blueprints(app):
    """Регистрация маршрутов"""
    from app.routes import main_bp, upload_bp, report_bp, api_bp, admin_bp
    app.register_blueprint(main_bp)
    app.register_blueprint(upload_bp)
    app.register_blueprint(report_bp)
    app.register_blueprint(api_bp)
    app.register_blueprint(admin_bp)