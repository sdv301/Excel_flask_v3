# database/models.py
from sqlalchemy import Column, Integer, String, Float, Date, DateTime, Boolean, Text, JSON, ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship
from datetime import datetime

Base = declarative_base()

class Company(Base):
    __tablename__ = 'companies'
    
    id = Column(Integer, primary_key=True)
    name = Column(String(255), nullable=False, unique=True)
    code = Column(String(50), unique=True)
    email_pattern = Column(String(255))
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.now)
    
    uploaded_files = relationship("UploadedFile", back_populates="company")

class UploadedFile(Base):
    __tablename__ = 'uploaded_files'
    
    id = Column(Integer, primary_key=True)
    company_id = Column(Integer, ForeignKey('companies.id'))
    filename = Column(String(500), nullable=False)
    file_path = Column(String(1000), nullable=False)
    report_date = Column(Date, nullable=False)
    upload_date = Column(DateTime, default=datetime.now)
    file_size = Column(Integer)
    status = Column(String(50), default='uploaded')
    error_message = Column(Text)
    
    company = relationship("Company", back_populates="uploaded_files")
    
    __table_args__ = (
        {'sqlite_autoincrement': True},
    )

class Sheet1Structure(Base):
    __tablename__ = 'sheet1_structure'
    
    id = Column(Integer, primary_key=True)
    file_id = Column(Integer, ForeignKey('uploaded_files.id'))
    company_id = Column(Integer, ForeignKey('companies.id'))
    report_date = Column(Date, nullable=False)
    
    affiliation = Column(String(500))
    company_name = Column(String(500))
    oil_depots_count = Column(Integer)
    azs_count = Column(Integer)
    working_azs_count = Column(Integer)
    
    created_at = Column(DateTime, default=datetime.now)

class Sheet2Demand(Base):
    __tablename__ = 'sheet2_demand'
    
    id = Column(Integer, primary_key=True)
    file_id = Column(Integer, ForeignKey('uploaded_files.id'))
    company_id = Column(Integer, ForeignKey('companies.id'))
    report_date = Column(Date, nullable=False)
    
    # Годовая потребность
    year = Column(Integer)
    gasoline_total = Column(Float)
    gasoline_ai76_80 = Column(Float)
    gasoline_ai92 = Column(Float)
    gasoline_ai95 = Column(Float)
    gasoline_ai98_100 = Column(Float)
    diesel_total = Column(Float)
    diesel_winter = Column(Float)
    diesel_arctic = Column(Float)
    diesel_summer = Column(Float)
    diesel_intermediate = Column(Float)
    
    # Месячная потребность
    month = Column(String(20))
    monthly_gasoline_total = Column(Float)
    monthly_gasoline_ai76_80 = Column(Float)
    monthly_gasoline_ai92 = Column(Float)
    monthly_gasoline_ai95 = Column(Float)
    monthly_gasoline_ai98_100 = Column(Float)
    monthly_diesel_total = Column(Float)
    monthly_diesel_winter = Column(Float)
    monthly_diesel_arctic = Column(Float)
    monthly_diesel_summer = Column(Float)
    monthly_diesel_intermediate = Column(Float)
    
    created_at = Column(DateTime, default=datetime.now)

class Sheet3Balance(Base):
    __tablename__ = 'sheet3_balance'
    
    id = Column(Integer, primary_key=True)
    file_id = Column(Integer, ForeignKey('uploaded_files.id'))
    company_id = Column(Integer, ForeignKey('companies.id'))
    report_date = Column(Date, nullable=False)
    
    affiliation = Column(String(500))
    company_name = Column(String(500))
    location_type = Column(String(100))
    location_name = Column(String(500))
    
    # Имеющиеся запасы
    stock_ai76_80 = Column(Float)
    stock_ai92 = Column(Float)
    stock_ai95 = Column(Float)
    stock_ai98_100 = Column(Float)
    stock_diesel_winter = Column(Float)
    stock_diesel_arctic = Column(Float)
    stock_diesel_summer = Column(Float)
    stock_diesel_intermediate = Column(Float)
    
    # Товар в пути
    transit_ai76_80 = Column(Float)
    transit_ai92 = Column(Float)
    transit_ai95 = Column(Float)
    transit_ai98_100 = Column(Float)
    transit_diesel_winter = Column(Float)
    transit_diesel_arctic = Column(Float)
    transit_diesel_summer = Column(Float)
    transit_diesel_intermediate = Column(Float)
    
    # Емкость хранения
    capacity_ai76_80 = Column(Float)
    capacity_ai92 = Column(Float)
    capacity_ai95 = Column(Float)
    capacity_ai98_100 = Column(Float)
    capacity_diesel_winter = Column(Float)
    capacity_diesel_arctic = Column(Float)
    capacity_diesel_summer = Column(Float)
    capacity_diesel_intermediate = Column(Float)
    
    created_at = Column(DateTime, default=datetime.now)

class Sheet4Supply(Base):
    __tablename__ = 'sheet4_supply'
    
    id = Column(Integer, primary_key=True)
    file_id = Column(Integer, ForeignKey('uploaded_files.id'))
    company_id = Column(Integer, ForeignKey('companies.id'))
    report_date = Column(Date, nullable=False)
    
    affiliation = Column(String(500))
    company_name = Column(String(500))
    oil_depot_name = Column(String(500))
    supply_date = Column(Date)
    
    supply_ai76_80 = Column(Float)
    supply_ai92 = Column(Float)
    supply_ai95 = Column(Float)
    supply_ai98_100 = Column(Float)
    supply_diesel_winter = Column(Float)
    supply_diesel_arctic = Column(Float)
    supply_diesel_summer = Column(Float)
    supply_diesel_intermediate = Column(Float)
    
    created_at = Column(DateTime, default=datetime.now)

class Sheet5Sales(Base):
    __tablename__ = 'sheet5_sales'
    
    id = Column(Integer, primary_key=True)
    file_id = Column(Integer, ForeignKey('uploaded_files.id'))
    company_id = Column(Integer, ForeignKey('companies.id'))
    report_date = Column(Date, nullable=False)
    
    affiliation = Column(String(500))
    company_name = Column(String(500))
    location_type = Column(String(100))
    location_name = Column(String(500))
    
    # Реализация за сутки
    daily_ai76_80 = Column(Float)
    daily_ai92 = Column(Float)
    daily_ai95 = Column(Float)
    daily_ai98_100 = Column(Float)
    daily_diesel_winter = Column(Float)
    daily_diesel_arctic = Column(Float)
    daily_diesel_summer = Column(Float)
    daily_diesel_intermediate = Column(Float)
    
    # Реализация с начала месяца
    monthly_ai76_80 = Column(Float)
    monthly_ai92 = Column(Float)
    monthly_ai95 = Column(Float)
    monthly_ai98_100 = Column(Float)
    monthly_diesel_winter = Column(Float)
    monthly_diesel_arctic = Column(Float)
    monthly_diesel_summer = Column(Float)
    monthly_diesel_intermediate = Column(Float)
    
    created_at = Column(DateTime, default=datetime.now)
# В models.py добавляем эти классы после Sheet5Sales:

class Sheet6Aviation(Base):
    __tablename__ = 'sheet6_aviation'
    
    id = Column(Integer, primary_key=True)
    file_id = Column(Integer, ForeignKey('uploaded_files.id'))
    company_id = Column(Integer, ForeignKey('companies.id'))
    report_date = Column(Date, nullable=False)
    
    airport_name = Column(String(500))
    tzk_name = Column(String(500))
    contracts_info = Column(Text)
    
    supply_week = Column(Float)
    supply_month_start = Column(Float)
    monthly_demand = Column(Float)
    consumption_week = Column(Float)
    consumption_month_start = Column(Float)
    end_of_day_balance = Column(Float)
    
    created_at = Column(DateTime, default=datetime.now)

class Sheet7Comments(Base):
    __tablename__ = 'sheet7_comments'
    
    id = Column(Integer, primary_key=True)
    file_id = Column(Integer, ForeignKey('uploaded_files.id'))
    company_id = Column(Integer, ForeignKey('companies.id'))
    report_date = Column(Date, nullable=False)
    
    fuel_type = Column(String(100))
    situation = Column(String(100))
    comments = Column(Text)
    
    created_at = Column(DateTime, default=datetime.now)

class DataHistory(Base):
    __tablename__ = 'data_history'
    
    id = Column(Integer, primary_key=True)
    table_name = Column(String(100), nullable=False)
    record_id = Column(Integer, nullable=False)
    operation = Column(String(10), nullable=False)  # INSERT, UPDATE, DELETE
    old_data = Column(JSON)
    new_data = Column(JSON)
    changed_by = Column(String(100))
    changed_at = Column(DateTime, default=datetime.now)

class ReportConfig(Base):
    __tablename__ = 'report_configs'
    
    id = Column(Integer, primary_key=True)
    name = Column(String(255), nullable=False)
    description = Column(Text)
    template_path = Column(String(1000))
    config = Column(JSON, nullable=False)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.now)

class GeneratedReport(Base):
    __tablename__ = 'generated_reports'
    
    id = Column(Integer, primary_key=True)
    report_config_id = Column(Integer, ForeignKey('report_configs.id'))
    report_date = Column(Date, nullable=False)
    file_path = Column(String(1000), nullable=False)
    file_size = Column(Integer)
    generated_by = Column(String(100))
    generated_at = Column(DateTime, default=datetime.now)
    status = Column(String(50), default='generated')