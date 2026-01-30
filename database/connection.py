# database/connection.py
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, scoped_session
from config import Config

class DatabaseConnection:
    def __init__(self):
        self.engine = create_engine(Config.SQLALCHEMY_DATABASE_URI)
        self.session_factory = sessionmaker(bind=self.engine)
        self.Session = scoped_session(self.session_factory)
    
    def get_session(self):
        return self.Session()
    
    def close_session(self):
        self.Session.remove()
    
    def create_tables(self):
        """Создание всех таблиц в базе данных"""
        from .models import Base
        Base.metadata.create_all(self.engine)
        print("Таблицы базы данных созданы успешно")
    
    def drop_tables(self):
        """Удаление всех таблиц (для тестирования)"""
        from .models import Base
        Base.metadata.drop_all(self.engine)
        print("Таблицы базы данных удалены")

# Создаем глобальный экземпляр подключения
db_connection = DatabaseConnection()