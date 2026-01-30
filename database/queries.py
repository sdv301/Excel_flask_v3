# database/queries.py - ПОЛНОСТЬЮ ИСПРАВЛЕННАЯ ВЕРСИЯ
from .connection import db_connection
from .models import *
from datetime import datetime, date
from typing import List, Dict, Any
import json

class DatabaseQueries:
    def __init__(self):
        self.db = db_connection
    
    def normalize_company_name(self, name: str) -> str:
        """Нормализация названия компании для поиска - УПРОЩЕННАЯ ВЕРСИЯ"""
        if not name:
            return "Неизвестная компания"
        
        clean = str(name).strip()
        print(f"Нормализация: '{name}' -> '{clean}'")
        
        # Определяем компанию по ключевым словам
        clean_lower = clean.lower()
        
        if any(x in clean_lower for x in ['сибирь ойл', 'сибир', 'сибойл', 'ооо "сибирь ойл"']):
            result = 'Сибойл'
        elif 'туймаада' in clean_lower:
            result = 'Туймаада-Нефть'
        elif any(x in clean_lower for x in ['саханефтегазсбыт', 'саха нефтегазсбыт', 'снгс', 'ао "саханефтегазсбыт"']):
            result = 'Саханефтегазсбыт'
        elif 'экто' in clean_lower:
            result = 'ЭКТО-Ойл'
        elif 'газпром' in clean_lower:
            result = 'Газпром'
        elif 'роснефть' in clean_lower:
            result = 'Роснефть'
        elif 'татнефть' in clean_lower:
            result = 'Татнефть'
        elif 'паритет' in clean_lower:
            result = 'Паритет'
        elif 'сибирское топливо' in clean_lower:
            result = 'Сибирское топливо'
        else:
            # Если не нашли, возвращаем как есть
            result = clean
        
        print(f"  Результат: '{result}'")
        return result
    
    def add_company(self, name: str, code: str = None, email_pattern: str = None) -> Company:
        """Добавление компании"""
        session = self.db.get_session()
        try:
            company = Company(
                name=name,
                code=code,
                email_pattern=email_pattern
            )
            session.add(company)
            session.commit()
            return company
        except Exception as e:
            session.rollback()
            raise e
        finally:
            self.db.close_session()
    
    def save_uploaded_file(self, filename: str, file_path: str, 
                      company_name: str, report_date: date) -> tuple:
        """Сохранение информации о загруженном файле"""
        session = self.db.get_session()
        try:
            # Нормализуем название компании
            normalized_name = self.normalize_company_name(company_name)
            print(f"Сохранение файла для компании: '{company_name}' -> '{normalized_name}'")
            
            # Ищем компанию в базе по нормализованному имени
            company = None
            for c in session.query(Company).all():
                if normalized_name.lower() in c.name.lower() or c.name.lower() in normalized_name.lower():
                    company = c
                    print(f"  Найдена существующая компания: {c.name} (ID: {c.id})")
                    break
            
            # Если не нашли, создаем новую компанию
            if not company:
                company = Company(name=normalized_name)
                session.add(company)
                session.commit()
                print(f"  Создана новая компания: {normalized_name} (ID: {company.id})")
            
            # Проверяем, нет ли уже файла на эту дату для этой компании
            existing = session.query(UploadedFile).filter(
                UploadedFile.company_id == company.id,
                UploadedFile.report_date == report_date
            ).first()
            
            if existing:
                # Обновляем существующий файл
                existing.filename = filename
                existing.file_path = file_path
                existing.upload_date = datetime.now()
                existing.status = 'processed'
                file_id = existing.id
                print(f"  Обновлен существующий файл ID: {file_id}")
            else:
                # Создаем новый файл
                uploaded_file = UploadedFile(
                    company_id=company.id,
                    filename=filename,
                    file_path=file_path,
                    report_date=report_date,
                    file_size=0,
                    status='processed'
                )
                session.add(uploaded_file)
                session.commit()
                file_id = uploaded_file.id
                print(f"  Создан новый файл ID: {file_id} для компании ID: {company.id}")
            
            return file_id, company.id
            
        except Exception as e:
            session.rollback()
            print(f"Ошибка сохранения файла: {e}")
            raise e
        finally:
            self.db.close_session()
    
    def get_companies(self) -> List[Company]:
        """Получение списка всех компаний"""
        session = self.db.get_session()
        try:
            return session.query(Company).filter(Company.is_active == True).all()
        finally:
            self.db.close_session()
    
    def get_recent_files(self, limit: int = 10) -> List[Dict]:
        """Получение последних загруженных файлов"""
        session = self.db.get_session()
        try:
            files = session.query(
                UploadedFile,
                Company.name.label('company_name')
            ).join(
                Company, UploadedFile.company_id == Company.id
            ).order_by(
                UploadedFile.upload_date.desc()
            ).limit(limit).all()
            
            return [
                {
                    'id': file.UploadedFile.id,
                    'filename': file.UploadedFile.filename,
                    'company_name': file.company_name,
                    'report_date': file.UploadedFile.report_date,
                    'upload_date': file.UploadedFile.upload_date,
                    'status': file.UploadedFile.status
                }
                for file in files
            ]
        finally:
            self.db.close_session()
    
    def save_sheet1_data(self, file_id: int, company_id: int, report_date: date, data: List[Dict]):
        """Сохранение данных из листа 1"""
        session = self.db.get_session()
        try:
            # Удаляем старые данные для этого файла
            session.query(Sheet1Structure).filter(
                Sheet1Structure.file_id == file_id
            ).delete()
            
            # Сохраняем каждую запись
            for item in data:
                sheet1 = Sheet1Structure(
                    file_id=file_id,
                    company_id=company_id,
                    report_date=report_date,
                    affiliation=item.get('affiliation'),
                    company_name=item.get('company_name'),
                    oil_depots_count=item.get('oil_depots_count'),
                    azs_count=item.get('azs_count'),
                    working_azs_count=item.get('working_azs_count')
                )
                session.add(sheet1)
            
            session.commit()
            print(f"  Sheet1 сохранено: {len(data)} записей для файла {file_id}")
        except Exception as e:
            session.rollback()
            print(f"Ошибка сохранения sheet1: {e}")
            raise e
        finally:
            self.db.close_session()
    
    def save_sheet2_data(self, file_id: int, company_id: int, report_date: date, data: Dict):
        """Сохранение данных из листа 2"""
        if not data:
            return
        
        session = self.db.get_session()
        try:
            # Удаляем старые данные
            session.query(Sheet2Demand).filter(
                Sheet2Demand.file_id == file_id
            ).delete()
            
            # Сохраняем данные
            sheet2 = Sheet2Demand(
                file_id=file_id,
                company_id=company_id,
                report_date=report_date,
                year=data.get('year'),
                gasoline_total=data.get('gasoline_total'),
                gasoline_ai76_80=data.get('gasoline_ai76_80'),
                gasoline_ai92=data.get('gasoline_ai92'),
                gasoline_ai95=data.get('gasoline_ai95'),
                gasoline_ai98_100=data.get('gasoline_ai98_100'),
                diesel_total=data.get('diesel_total'),
                diesel_winter=data.get('diesel_winter'),
                diesel_arctic=data.get('diesel_arctic'),
                diesel_summer=data.get('diesel_summer'),
                diesel_intermediate=data.get('diesel_intermediate'),
                month=data.get('month'),
                monthly_gasoline_total=data.get('monthly_gasoline_total'),
                monthly_gasoline_ai76_80=data.get('monthly_gasoline_ai76_80'),
                monthly_gasoline_ai92=data.get('monthly_gasoline_ai92'),
                monthly_gasoline_ai95=data.get('monthly_gasoline_ai95'),
                monthly_gasoline_ai98_100=data.get('monthly_gasoline_ai98_100'),
                monthly_diesel_total=data.get('monthly_diesel_total'),
                monthly_diesel_winter=data.get('monthly_diesel_winter'),
                monthly_diesel_arctic=data.get('monthly_diesel_arctic'),
                monthly_diesel_summer=data.get('monthly_diesel_summer'),
                monthly_diesel_intermediate=data.get('monthly_diesel_intermediate'),
            )
            session.add(sheet2)
            session.commit()
            print(f"  Sheet2 сохранено для файла {file_id}")
        except Exception as e:
            session.rollback()
            print(f"Ошибка сохранения sheet2: {e}")
            raise e
        finally:
            self.db.close_session()
    
    def save_sheet3_data(self, file_id: int, company_id: int, report_date: date, data: List[Dict]):
        """Сохранение данных из листа 3"""
        session = self.db.get_session()
        try:
            # Удаляем старые данные
            session.query(Sheet3Balance).filter(
                Sheet3Balance.file_id == file_id
            ).delete()
            
            # Сохраняем каждую запись
            for item in data:
                sheet3 = Sheet3Balance(
                    file_id=file_id,
                    company_id=company_id,
                    report_date=report_date,
                    affiliation=item.get('affiliation'),
                    company_name=item.get('company_name'),
                    location_type=item.get('location_type'),
                    location_name=item.get('location_name'),
                    # Имеющиеся запасы
                    stock_ai76_80=item.get('stock_ai76_80'),
                    stock_ai92=item.get('stock_ai92'),
                    stock_ai95=item.get('stock_ai95'),
                    stock_ai98_100=item.get('stock_ai98_100'),
                    stock_diesel_winter=item.get('stock_diesel_winter'),
                    stock_diesel_arctic=item.get('stock_diesel_arctic'),
                    stock_diesel_summer=item.get('stock_diesel_summer'),
                    stock_diesel_intermediate=item.get('stock_diesel_intermediate'),
                    # Товар в пути
                    transit_ai76_80=item.get('transit_ai76_80'),
                    transit_ai92=item.get('transit_ai92'),
                    transit_ai95=item.get('transit_ai95'),
                    transit_ai98_100=item.get('transit_ai98_100'),
                    transit_diesel_winter=item.get('transit_diesel_winter'),
                    transit_diesel_arctic=item.get('transit_diesel_arctic'),
                    transit_diesel_summer=item.get('transit_diesel_summer'),
                    transit_diesel_intermediate=item.get('transit_diesel_intermediate'),
                    # Емкость хранения
                    capacity_ai76_80=item.get('capacity_ai76_80'),
                    capacity_ai92=item.get('capacity_ai92'),
                    capacity_ai95=item.get('capacity_ai95'),
                    capacity_ai98_100=item.get('capacity_ai98_100'),
                    capacity_diesel_winter=item.get('capacity_diesel_winter'),
                    capacity_diesel_arctic=item.get('capacity_diesel_arctic'),
                    capacity_diesel_summer=item.get('capacity_diesel_summer'),
                    capacity_diesel_intermediate=item.get('capacity_diesel_intermediate'),
                )
                session.add(sheet3)
            
            session.commit()
            print(f"  Sheet3 сохранено: {len(data)} записей для файла {file_id}")
        except Exception as e:
            session.rollback()
            print(f"Ошибка сохранения sheet3: {e}")
            raise e
        finally:
            self.db.close_session()
    
    def save_sheet4_data(self, file_id: int, company_id: int, report_date: date, data: List[Dict]):
        """Сохранение данных из листа 4"""
        session = self.db.get_session()
        try:
            # Удаляем старые данные
            session.query(Sheet4Supply).filter(
                Sheet4Supply.file_id == file_id
            ).delete()
            
            # Сохраняем каждую запись
            for item in data:
                sheet4 = Sheet4Supply(
                    file_id=file_id,
                    company_id=company_id,
                    report_date=report_date,
                    affiliation=item.get('affiliation'),
                    company_name=item.get('company_name'),
                    oil_depot_name=item.get('oil_depot_name'),
                    supply_date=item.get('supply_date'),
                    supply_ai76_80=item.get('supply_ai76_80'),
                    supply_ai92=item.get('supply_ai92'),
                    supply_ai95=item.get('supply_ai95'),
                    supply_ai98_100=item.get('supply_ai98_100'),
                    supply_diesel_winter=item.get('supply_diesel_winter'),
                    supply_diesel_arctic=item.get('supply_diesel_arctic'),
                    supply_diesel_summer=item.get('supply_diesel_summer'),
                    supply_diesel_intermediate=item.get('supply_diesel_intermediate'),
                )
                session.add(sheet4)
            
            session.commit()
            print(f"  Sheet4 сохранено: {len(data)} записей для файла {file_id}")
        except Exception as e:
            session.rollback()
            print(f"Ошибка сохранения sheet4: {e}")
            raise e
        finally:
            self.db.close_session()
    
    def save_sheet5_data(self, file_id: int, company_id: int, report_date: date, data: List[Dict]):
        """Сохранение данных из листа 5"""
        session = self.db.get_session()
        try:
            # Удаляем старые данные
            session.query(Sheet5Sales).filter(
                Sheet5Sales.file_id == file_id
            ).delete()
            
            # Сохраняем каждую запись
            for item in data:
                sheet5 = Sheet5Sales(
                    file_id=file_id,
                    company_id=company_id,
                    report_date=report_date,
                    affiliation=item.get('affiliation'),
                    company_name=item.get('company_name'),
                    location_type=item.get('location_type'),
                    location_name=item.get('location_name'),
                    # Реализация за сутки
                    daily_ai76_80=item.get('daily_ai76_80'),
                    daily_ai92=item.get('daily_ai92'),
                    daily_ai95=item.get('daily_ai95'),
                    daily_ai98_100=item.get('daily_ai98_100'),
                    daily_diesel_winter=item.get('daily_diesel_winter'),
                    daily_diesel_arctic=item.get('daily_diesel_arctic'),
                    daily_diesel_summer=item.get('daily_diesel_summer'),
                    daily_diesel_intermediate=item.get('daily_diesel_intermediate'),
                    # Реализация с начала месяца
                    monthly_ai76_80=item.get('monthly_ai76_80'),
                    monthly_ai92=item.get('monthly_ai92'),
                    monthly_ai95=item.get('monthly_ai95'),
                    monthly_ai98_100=item.get('monthly_ai98_100'),
                    monthly_diesel_winter=item.get('monthly_diesel_winter'),
                    monthly_diesel_arctic=item.get('monthly_diesel_arctic'),
                    monthly_diesel_summer=item.get('monthly_diesel_summer'),
                    monthly_diesel_intermediate=item.get('monthly_diesel_intermediate'),
                )
                session.add(sheet5)
            
            session.commit()
            print(f"  Sheet5 сохранено: {len(data)} записей для файла {file_id}")
        except Exception as e:
            session.rollback()
            print(f"Ошибка сохранения sheet5: {e}")
            raise e
        finally:
            self.db.close_session()
    
    def save_sheet6_data(self, file_id: int, company_id: int, report_date: date, data: List[Dict]):
        """Сохранение данных из листа 6"""
        session = self.db.get_session()
        try:
            # Удаляем старые данные
            session.query(Sheet6Aviation).filter(
                Sheet6Aviation.file_id == file_id
            ).delete()
            
            # Сохраняем каждую запись
            for item in data:
                sheet6 = Sheet6Aviation(
                    file_id=file_id,
                    company_id=company_id,
                    report_date=report_date,
                    airport_name=item.get('airport_name'),
                    tzk_name=item.get('tzk_name'),
                    contracts_info=item.get('contracts_info'),
                    supply_week=item.get('supply_week'),
                    supply_month_start=item.get('supply_month_start'),
                    monthly_demand=item.get('monthly_demand'),
                    consumption_week=item.get('consumption_week'),
                    consumption_month_start=item.get('consumption_month_start'),
                    end_of_day_balance=item.get('end_of_day_balance'),
                )
                session.add(sheet6)
            
            session.commit()
            print(f"  Sheet6 сохранено: {len(data)} записей для файла {file_id}")
        except Exception as e:
            session.rollback()
            print(f"Ошибка сохранения sheet6: {e}")
            raise e
        finally:
            self.db.close_session()
    
    def save_sheet7_data(self, file_id: int, company_id: int, report_date: date, data: List[Dict]):
        """Сохранение данных из листа 7"""
        session = self.db.get_session()
        try:
            # Удаляем старые данные
            session.query(Sheet7Comments).filter(
                Sheet7Comments.file_id == file_id
            ).delete()
            
            # Сохраняем каждую запись
            for item in data:
                sheet7 = Sheet7Comments(
                    file_id=file_id,
                    company_id=company_id,
                    report_date=report_date,
                    fuel_type=item.get('fuel_type'),
                    situation=item.get('situation'),
                    comments=item.get('comments'),
                )
                session.add(sheet7)
            
            session.commit()
            print(f"  Sheet7 сохранено: {len(data)} записей для файла {file_id}")
        except Exception as e:
            session.rollback()
            print(f"Ошибка сохранения sheet7: {e}")
            raise e
        finally:
            self.db.close_session()
    
    def get_aggregated_data(self, report_date: date = None, company_id: int = None) -> Dict[str, Any]:
            """Получение агрегированных данных для сводного отчета - ИСПРАВЛЕННАЯ ВЕРСИЯ"""
            session = self.db.get_session()
            try:
                print(f"\n=== ПОИСК АГРЕГИРОВАННЫХ ДАННЫХ ===")
                print(f"Запрошена дата: {report_date}")
                print(f"Запрошена компания ID: {company_id}")
                
                result = {}
                
                # Если указана конкретная компания, ищем только ее
                if company_id:
                    companies = session.query(Company).filter(Company.id == company_id).all()
                    print(f"Ищем только компанию ID: {company_id}")
                else:
                    # Ищем все компании из БД
                    companies = session.query(Company).filter(Company.is_active == True).all()
                    print(f"Ищем все активные компании: {len(companies)} шт.")
                
                for company in companies:
                    print(f"\n--- Обработка компании: {company.name} (ID: {company.id}) ---")
                    
                    company_data = {
                        'name': company.name,
                        'sheet1': [],
                        'sheet2': {},
                        'sheet3_totals': {},
                        'sheet5_totals': {},
                        'sheet3_data': [],
                        'sheet4_data': [],
                        'sheet5_data': []
                    }
                    
                    has_data = False
                    
                    # 1. Данные из листа 1 - структура
                    # Ищем последние данные, а не по конкретной дате
                    sheet1_query = session.query(Sheet1Structure).filter(
                        Sheet1Structure.company_id == company.id
                    ).order_by(Sheet1Structure.report_date.desc())
                    
                    sheet1_items = sheet1_query.all()
                    print(f"  Sheet1 найдено записей: {len(sheet1_items)}")
                    
                    # Берем только последнюю дату для этой компании
                    if sheet1_items:
                        last_report_date = sheet1_items[0].report_date
                        sheet1_items = [item for item in sheet1_items if item.report_date == last_report_date]
                        print(f"  Используем данные на дату: {last_report_date}")
                    
                    for item in sheet1_items:
                        company_data['sheet1'].append({
                            'affiliation': item.affiliation,
                            'company_name': item.company_name,
                            'oil_depots_count': item.oil_depots_count,
                            'azs_count': item.azs_count,
                            'working_azs_count': item.working_azs_count
                        })
                        if item.azs_count or item.oil_depots_count:
                            has_data = True
                    
                    # 2. Данные из листа 2 - потребность
                    # Ищем последние данные
                    sheet2_query = session.query(Sheet2Demand).filter(
                        Sheet2Demand.company_id == company.id
                    ).order_by(Sheet2Demand.report_date.desc())
                    
                    sheet2_items = sheet2_query.all()
                    print(f"  Sheet2 найдено записей: {len(sheet2_items)}")
                    
                    # Берем данные с последней даты
                    if sheet2_items:
                        # Сгруппируем по дате и выберем последнюю
                        dates = sorted(set(item.report_date for item in sheet2_items), reverse=True)
                        if dates:
                            last_date = dates[0]
                            recent_items = [item for item in sheet2_items if item.report_date == last_date]
                            
                            if recent_items:
                                # Суммируем данные за последнюю дату
                                demand_data = {
                                    'year': last_date.year,
                                    'gasoline_total': sum(item.gasoline_total or 0 for item in recent_items),
                                    'diesel_total': sum(item.diesel_total or 0 for item in recent_items),
                                    'monthly_gasoline_total': sum(item.monthly_gasoline_total or 0 for item in recent_items),
                                    'monthly_diesel_total': sum(item.monthly_diesel_total or 0 for item in recent_items),
                                    'gasoline_ai92': sum(item.gasoline_ai92 or 0 for item in recent_items),
                                    'gasoline_ai95': sum(item.gasoline_ai95 or 0 for item in recent_items),
                                    'diesel_winter': sum(item.diesel_winter or 0 for item in recent_items),
                                    'diesel_arctic': sum(item.diesel_arctic or 0 for item in recent_items),
                                    'report_date': last_date
                                }
                                
                                company_data['sheet2'] = demand_data
                                if demand_data['gasoline_total'] or demand_data['diesel_total']:
                                    has_data = True
                    
                    # 3. Данные из листа 3 - остатки
                    # Ищем последние данные
                    sheet3_query = session.query(Sheet3Balance).filter(
                        Sheet3Balance.company_id == company.id
                    ).order_by(Sheet3Balance.report_date.desc())
                    
                    sheet3_items = sheet3_query.all()
                    print(f"  Sheet3 найдено записей: {len(sheet3_items)}")
                    
                    # Группируем по дате и берем последнюю
                    if sheet3_items:
                        dates = sorted(set(item.report_date for item in sheet3_items), reverse=True)
                        if dates:
                            last_date = dates[0]
                            recent_items = [item for item in sheet3_items if item.report_date == last_date]
                            
                            if recent_items:
                                total_stock_ai92 = sum(item.stock_ai92 or 0 for item in recent_items)
                                total_stock_ai95 = sum(item.stock_ai95 or 0 for item in recent_items)
                                total_stock_diesel_winter = sum(item.stock_diesel_winter or 0 for item in recent_items)
                                total_stock_diesel_arctic = sum(item.stock_diesel_arctic or 0 for item in recent_items)
                                
                                company_data['sheet3_totals'] = {
                                    'total_stock_ai92': total_stock_ai92,
                                    'total_stock_ai95': total_stock_ai95,
                                    'total_stock_diesel_winter': total_stock_diesel_winter,
                                    'total_stock_diesel_arctic': total_stock_diesel_arctic,
                                    'record_count': len(recent_items),
                                    'report_date': last_date
                                }
                                
                                company_data['sheet3_data'] = [
                                    {
                                        'location_name': item.location_name,
                                        'stock_ai92': item.stock_ai92,
                                        'stock_ai95': item.stock_ai95,
                                        'stock_diesel_winter': item.stock_diesel_winter,
                                        'stock_diesel_arctic': item.stock_diesel_arctic
                                    }
                                    for item in recent_items
                                ]
                                
                                if total_stock_ai92 or total_stock_ai95:
                                    has_data = True
                    
                    # 4. Данные из листа 4 - поставки
                    # Ищем последние данные
                    sheet4_query = session.query(Sheet4Supply).filter(
                        Sheet4Supply.company_id == company.id
                    ).order_by(Sheet4Supply.report_date.desc())
                    
                    sheet4_items = sheet4_query.all()
                    print(f"  Sheet4 найдено записей: {len(sheet4_items)}")
                    
                    # Группируем по дате
                    if sheet4_items:
                        dates = sorted(set(item.report_date for item in sheet4_items), reverse=True)
                        if dates:
                            last_date = dates[0]
                            recent_items = [item for item in sheet4_items if item.report_date == last_date]
                            
                            for item in recent_items:
                                company_data['sheet4_data'].append({
                                    'oil_depot_name': item.oil_depot_name,
                                    'supply_date': item.supply_date,
                                    'supply_ai92': item.supply_ai92,
                                    'supply_ai95': item.supply_ai95,
                                    'supply_diesel_winter': item.supply_diesel_winter,
                                    'supply_diesel_arctic': item.supply_diesel_arctic,
                                    'report_date': item.report_date
                                })
                                if item.supply_ai92 or item.supply_diesel_winter:
                                    has_data = True
                    
                    # 5. Данные из листа 5 - реализация
                    # Ищем последние данные
                    sheet5_query = session.query(Sheet5Sales).filter(
                        Sheet5Sales.company_id == company.id
                    ).order_by(Sheet5Sales.report_date.desc())
                    
                    sheet5_items = sheet5_query.all()
                    print(f"  Sheet5 найдено записей: {len(sheet5_items)}")
                    
                    # Группируем по дате и берем последнюю
                    if sheet5_items:
                        dates = sorted(set(item.report_date for item in sheet5_items), reverse=True)
                        if dates:
                            last_date = dates[0]
                            recent_items = [item for item in sheet5_items if item.report_date == last_date]
                            
                            if recent_items:
                                total_monthly_ai92 = sum(item.monthly_ai92 or 0 for item in recent_items)
                                total_monthly_ai95 = sum(item.monthly_ai95 or 0 for item in recent_items)
                                total_monthly_diesel_winter = sum(item.monthly_diesel_winter or 0 for item in recent_items)
                                total_monthly_diesel_arctic = sum(item.monthly_diesel_arctic or 0 for item in recent_items)
                                
                                company_data['sheet5_totals'] = {
                                    'total_monthly_ai92': total_monthly_ai92,
                                    'total_monthly_ai95': total_monthly_ai95,
                                    'total_monthly_diesel_winter': total_monthly_diesel_winter,
                                    'total_monthly_diesel_arctic': total_monthly_diesel_arctic,
                                    'record_count': len(recent_items),
                                    'report_date': last_date
                                }
                                
                                company_data['sheet5_data'] = [
                                    {
                                        'location_name': item.location_name,
                                        'monthly_ai92': item.monthly_ai92,
                                        'monthly_ai95': item.monthly_ai95,
                                        'monthly_diesel_winter': item.monthly_diesel_winter,
                                        'monthly_diesel_arctic': item.monthly_diesel_arctic
                                    }
                                    for item in recent_items
                                ]
                                
                                if total_monthly_ai92 or total_monthly_ai95:
                                    has_data = True
                    
                    # Добавляем компанию в результат если есть данные
                    if has_data:
                        result[company.name] = company_data
                        print(f"  ✓ Данные добавлены для компании: {company.name}")
                        print(f"    - Sheet1: {len(company_data['sheet1'])} записей")
                        print(f"    - Sheet3: AI92={company_data['sheet3_totals'].get('total_stock_ai92', 0):.3f}, AI95={company_data['sheet3_totals'].get('total_stock_ai95', 0):.3f}")
                        print(f"    - Sheet5: AI92={company_data['sheet5_totals'].get('total_monthly_ai92', 0):.3f}, AI95={company_data['sheet5_totals'].get('total_monthly_ai95', 0):.3f}")
                    else:
                        print(f"  ✗ Нет данных для компании: {company.name}")
                
                print(f"\n=== РЕЗУЛЬТАТ ===")
                print(f"Найдено компаний с данными: {len(result)}")
                for name in result.keys():
                    print(f"  - {name}")
                
                return result
                
            except Exception as e:
                print(f"Ошибка при получении агрегированных данных: {e}")
                import traceback
                traceback.print_exc()
                return {}
            finally:
                self.db.close_session() 
    
    def update_file_status(self, file_id: int, status: str, error_message: str = None):
        """Обновление статуса файла"""
        session = self.db.get_session()
        try:
            uploaded_file = session.query(UploadedFile).get(file_id)
            if uploaded_file:
                uploaded_file.status = status
                if error_message:
                    uploaded_file.error_message = error_message
                session.commit()
                print(f"Статус файла {file_id} обновлен на '{status}'")
                return True
            print(f"Файл {file_id} не найден")
            return False
        except Exception as e:
            session.rollback()
            print(f"Ошибка обновления статуса файла {file_id}: {e}")
            raise e
        finally:
            self.db.close_session()
    
    def get_all_data_summary(self):
        """Получить сводку всех данных в БД (для отладки)"""
        session = self.db.get_session()
        try:
            summary = {
                'companies': session.query(Company).count(),
                'uploaded_files': session.query(UploadedFile).count(),
                'sheet1': session.query(Sheet1Structure).count(),
                'sheet2': session.query(Sheet2Demand).count(),
                'sheet3': session.query(Sheet3Balance).count(),
                'sheet4': session.query(Sheet4Supply).count(),
                'sheet5': session.query(Sheet5Sales).count(),
                'sheet6': session.query(Sheet6Aviation).count(),
                'sheet7': session.query(Sheet7Comments).count(),
            }
            
            # Последние файлы
            recent_files = session.query(UploadedFile).order_by(
                UploadedFile.upload_date.desc()
            ).limit(5).all()
            
            summary['recent_files'] = [
                {
                    'id': f.id,
                    'filename': f.filename,
                    'company_id': f.company_id,
                    'report_date': f.report_date,
                    'status': f.status
                }
                for f in recent_files
            ]
            
            return summary
        finally:
            self.db.close_session()

# Создаем глобальный экземпляр
db = DatabaseQueries()