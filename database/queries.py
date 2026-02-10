# database/queries.py - ПОЛНЫЙ ИСПРАВЛЕННЫЙ ФАЙЛ
from .connection import db_connection
from .models import *
from datetime import datetime, date as dt_date
from typing import List, Dict, Any
import json
import os

class DatabaseQueries:
    def __init__(self):
        self.db = db_connection
    
    def normalize_company_name(self, name: str) -> str:
        """Улучшенная нормализация названия компании"""
        if not name:
            return "Неизвестная компания"
        
        original_name = name
        clean = str(name).strip()
        
        # Удаляем лишние символы и приводим к нижнему регистру
        clean_lower = clean.lower()
        clean_lower = clean_lower.replace('"', '').replace('ооо', '').replace('ао', '').replace('пао', '').replace('«', '').replace('»', '').strip()
        
        print(f"🔍 Нормализация: '{original_name}' -> '{clean_lower}'")
        
        # Расширенное сопоставление с приоритетом точных совпадений
        exact_mappings = {
            'саханефтегазсбыт': 'Саханефтегазсбыт',
            'снгс': 'Саханефтегазсбыт',
            'санги': 'Саханефтегазсбыт',
            'туймаада-нефть': 'Туймаада-Нефть', 
            'туймааданефть': 'Туймаада-Нефть',
            'туймаада': 'Туймаада-Нефть',
            'сибойл': 'Сибойл',
            'сибирьойл': 'Сибойл',
            'сибирь ойл': 'Сибойл',
            'экто-ойл': 'ЭКТО-Ойл',
            'эктоойл': 'ЭКТО-Ойл',
            'экто': 'ЭКТО-Ойл',
            'сибирское топливо': 'Сибирское топливо',
            'сибтопливо': 'Сибирское топливо',
            'паритет': 'Паритет',
        }
        
        # Сначала проверяем точные совпадения
        for pattern, normalized_name in exact_mappings.items():
            if pattern in clean_lower:
                print(f"  ✅ Точное совпадение: '{normalized_name}'")
                return normalized_name
        
        # Затем проверяем частичные совпадения
        partial_mappings = [
            (['саха', 'нефтегазсбыт'], 'Саханефтегазсбыт'),
            (['туймаада', 'нефть'], 'Туймаада-Нефть'),
            (['сиб', 'ойл'], 'Сибойл'),
            (['сибирск', 'топливо'], 'Сибирское топливо'),
            (['экто', 'ойл'], 'ЭКТО-Ойл'),
        ]
        
        for patterns, normalized_name in partial_mappings:
            if all(pattern in clean_lower for pattern in patterns):
                print(f"  ✅ Частичное совпадение: '{normalized_name}'")
                return normalized_name
        
        # Если не нашли, возвращаем оригинальное название (очищенное)
        result = clean
        print(f"  ⚠️  Совпадений не найдено, используем: '{result}'")
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
                       company_name: str, report_date: dt_date) -> tuple:
        """Улучшенное сохранение информации о загруженном файле"""
        session = self.db.get_session()
        try:
            # Нормализуем название компании
            normalized_name = self.normalize_company_name(company_name)
            print(f"💾 Сохранение файла: '{filename}'")
            print(f"   Исходное название компании: '{company_name}'")
            print(f"   Нормализованное название: '{normalized_name}'")
            
            # Ищем компанию в базе по нормализованному имени
            company = None
            all_companies = session.query(Company).all()
            
            # Сначала точное совпадение
            for c in all_companies:
                if normalized_name.lower() == c.name.lower():
                    company = c
                    print(f"   ✅ Найдена точное совпадение: {c.name} (ID: {c.id})")
                    break
            
            # Если точного нет, ищем частичное
            if not company:
                for c in all_companies:
                    if (normalized_name.lower() in c.name.lower() or 
                        c.name.lower() in normalized_name.lower()):
                        company = c
                        print(f"   ✅ Найдено частичное совпадение: {c.name} (ID: {c.id})")
                        break
            
            # Если не нашли, создаем новую компанию
            if not company:
                company = Company(name=normalized_name)
                session.add(company)
                session.commit()
                print(f"   🆕 Создана новая компания: {normalized_name} (ID: {company.id})")
            
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
                print(f"   📝 Обновлен существующий файл ID: {file_id}")
            else:
                # Создаем новый файл
                uploaded_file = UploadedFile(
                    company_id=company.id,
                    filename=filename,
                    file_path=file_path,
                    report_date=report_date,
                    file_size=os.path.getsize(file_path) if os.path.exists(file_path) else 0,
                    status='processed'
                )
                session.add(uploaded_file)
                session.commit()
                file_id = uploaded_file.id
                print(f"   📄 Создан новый файл ID: {file_id}")
            
            return file_id, company.id
            
        except Exception as e:
            session.rollback()
            print(f"❌ Ошибка сохранения файла: {e}")
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
            
    def save_sheet1_data(self, file_id: int, company_id: int, report_date: dt_date, data: List[Dict]):
        """Сохранение данных из листа 1"""
        session = self.db.get_session()
        try:
            session.query(Sheet1Structure).filter(Sheet1Structure.file_id == file_id).delete()
            for item in data:
                sheet1 = Sheet1Structure(
                    file_id=file_id,
                    company_id=company_id,
                    report_date=report_date,
                    affiliation=item.get('affiliation', ''),
                    company_name=item.get('company', ''),
                    oil_depots_count=item.get('oil_depots_count', 0),
                    azs_count=item.get('azs_count', 0),
                    working_azs_count=item.get('working_azs_count', 0)
                )
                session.add(sheet1)
            session.commit()
        except Exception as e:
            session.rollback()
            raise e
        finally:
            self.db.close_session()
    
    def save_sheet2_data(self, file_id: int, company_id: int, report_date: dt_date, data: Dict):
        """Сохранение данных из листа 2"""
        if not data: return
        session = self.db.get_session()
        try:
            session.query(Sheet2Demand).filter(Sheet2Demand.file_id == file_id).delete()
            sheet2 = Sheet2Demand(
                file_id=file_id,
                company_id=company_id,
                report_date=report_date,
                year=datetime.now().year,
                gasoline_total=data.get('yearly_gasoline_total', 0),
                gasoline_ai92=data.get('yearly_ai92', 0),
                gasoline_ai95=data.get('yearly_ai95', 0),
                diesel_total=data.get('yearly_diesel_total', 0),
                month=datetime.now().strftime('%B'),
                monthly_gasoline_total=data.get('monthly_gasoline_total', 0),
                monthly_gasoline_ai92=data.get('monthly_ai92', 0),
                monthly_gasoline_ai95=data.get('monthly_ai95', 0),
                monthly_diesel_total=data.get('monthly_diesel_total', 0),
            )
            session.add(sheet2)
            session.commit()
        except Exception as e:
            session.rollback()
            raise e
        finally:
            self.db.close_session()
    
    def save_sheet3_data(self, file_id: int, company_id: int, report_date: dt_date, data: List[Dict]):
        """Сохранение данных из листа 3"""
        session = self.db.get_session()
        try:
            session.query(Sheet3Balance).filter(Sheet3Balance.file_id == file_id).delete()
            for item in data:
                sheet3 = Sheet3Balance(
                    file_id=file_id,
                    company_id=company_id,
                    report_date=report_date,
                    affiliation=item.get('group', ''),
                    company_name=item.get('company', ''),
                    location_type='Объект',
                    location_name=item.get('object_name', ''),
                    stock_ai92=item.get('stock_ai92', 0),
                    stock_ai95=item.get('stock_ai95', 0),
                    stock_ai98_100=item.get('stock_ai98_100', 0),
                    stock_diesel_winter=item.get('stock_diesel_winter', 0),
                    stock_diesel_arctic=item.get('stock_diesel_arctic', 0),
                    stock_diesel_summer=item.get('stock_diesel_summer', 0),
                    transit_ai92=item.get('transit_ai92', 0),
                    transit_ai95=item.get('transit_ai95', 0),
                    transit_ai98_100=item.get('transit_ai98_100', 0),
                    transit_diesel_winter=item.get('transit_diesel_winter', 0),
                    transit_diesel_arctic=item.get('transit_diesel_arctic', 0),
                    transit_diesel_summer=item.get('transit_diesel_summer', 0),
                    capacity_ai92=item.get('capacity_ai92', 0),
                    capacity_ai95=item.get('capacity_ai95', 0),
                    capacity_ai98_100=item.get('capacity_ai98_100', 0),
                    capacity_diesel_winter=item.get('capacity_diesel_winter', 0),
                    capacity_diesel_arctic=item.get('capacity_diesel_arctic', 0),
                    capacity_diesel_summer=item.get('capacity_diesel_summer', 0),
                )
                session.add(sheet3)
            session.commit()
        except Exception as e:
            session.rollback()
            raise e
        finally:
            self.db.close_session()
    
    def save_sheet4_data(self, file_id: int, company_id: int, report_date: dt_date, data: List[Dict]):
        """Сохранение данных из листа 4"""
        session = self.db.get_session()
        try:
            session.query(Sheet4Supply).filter(Sheet4Supply.file_id == file_id).delete()
            for item in data:
                sheet4 = Sheet4Supply(
                    file_id=file_id,
                    company_id=company_id,
                    report_date=report_date,
                    company_name=item.get('company', ''),
                    oil_depot_name=item.get('oil_depot', ''),
                    supply_date=self._parse_date_string(item.get('supply_date', '')),
                    supply_ai92=item.get('supply_ai92', 0),
                    supply_ai95=item.get('supply_ai95', 0),
                    supply_ai98_100=item.get('supply_ai98_100', 0),
                    supply_diesel_winter=item.get('supply_diesel_winter', 0),
                    supply_diesel_arctic=item.get('supply_diesel_arctic', 0),
                    supply_diesel_summer=item.get('supply_diesel_summer', 0),
                )
                session.add(sheet4)
            session.commit()
        except Exception as e:
            session.rollback()
            raise e
        finally:
            self.db.close_session()
    
    def save_sheet5_data(self, file_id: int, company_id: int, report_date: dt_date, data: List[Dict]):
        """Сохранение данных из листа 5"""
        session = self.db.get_session()
        try:
            session.query(Sheet5Sales).filter(Sheet5Sales.file_id == file_id).delete()
            for item in data:
                sheet5 = Sheet5Sales(
                    file_id=file_id,
                    company_id=company_id,
                    report_date=report_date,
                    company_name=item.get('company', ''),
                    location_name=item.get('object_name', ''),
                    daily_ai92=item.get('daily_ai92', 0),
                    daily_ai95=item.get('daily_ai95', 0),
                    daily_ai98_100=item.get('daily_ai98_100', 0),
                    daily_diesel_winter=item.get('daily_winter', 0),
                    daily_diesel_arctic=item.get('daily_arctic', 0),
                    daily_diesel_summer=item.get('daily_summer', 0),
                    monthly_ai92=item.get('monthly_ai92', 0),
                    monthly_ai95=item.get('monthly_ai95', 0),
                    monthly_ai98_100=item.get('monthly_ai98_100', 0),
                    monthly_diesel_winter=item.get('monthly_winter', 0),
                    monthly_diesel_arctic=item.get('monthly_arctic', 0),
                    monthly_diesel_summer=item.get('monthly_summer', 0),
                )
                session.add(sheet5)
            session.commit()
        except Exception as e:
            session.rollback()
            raise e
        finally:
            self.db.close_session()
            
    def save_sheet6_data(self, file_id: int, company_id: int, report_date: dt_date, data: List[Dict]):
        session = self.db.get_session()
        try:
            session.query(Sheet6Aviation).filter(Sheet6Aviation.file_id == file_id).delete()
            for item in data:
                sheet6 = Sheet6Aviation(
                    file_id=file_id, company_id=company_id, report_date=report_date,
                    airport_name=item.get('airport', ''), tzk_name=item.get('tzk', ''),
                    contracts_info=item.get('contracts', ''), supply_week=item.get('supply_week', 0),
                    supply_month_start=item.get('supply_month_start', 0), monthly_demand=item.get('monthly_demand', 0),
                    consumption_week=item.get('consumption_week', 0), consumption_month_start=item.get('consumption_month_start', 0),
                    end_of_day_balance=item.get('end_of_day_balance', 0)
                )
                session.add(sheet6)
            session.commit()
        except Exception as e:
            session.rollback()
            raise e
        finally:
            self.db.close_session()

    def save_sheet7_data(self, file_id: int, company_id: int, report_date: dt_date, data: List[Dict]):
        session = self.db.get_session()
        try:
            session.query(Sheet7Comments).filter(Sheet7Comments.file_id == file_id).delete()
            for item in data:
                sheet7 = Sheet7Comments(
                    file_id=file_id, company_id=company_id, report_date=report_date,
                    fuel_type=item.get('fuel_type', ''), situation=item.get('situation', ''),
                    comments=item.get('comments', '')
                )
                session.add(sheet7)
            session.commit()
        except Exception as e:
            session.rollback()
            raise e
        finally:
            self.db.close_session()

    def _parse_date_string(self, date_str: str):
        if not date_str: return None
        date_str = str(date_str).strip()
        for fmt in ['%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y', '%Y.%m.%d', '%d.%m.%y']:
            try: return datetime.strptime(date_str, fmt).date()
            except: continue
        return None

    def process_parsed_file(self, file_path: str, parsed_data: Dict[str, Any]):
        metadata = parsed_data.get('metadata', {})
        company_name = metadata.get('company', 'Неизвестная компания')
        report_date = metadata.get('report_date', datetime.now())
        if hasattr(report_date, 'date'): report_date = report_date.date()
        file_id, company_id = self.save_uploaded_file(os.path.basename(file_path), file_path, company_name, report_date)
        if 'sheet1' in parsed_data: self.save_sheet1_data(file_id, company_id, report_date, parsed_data['sheet1'])
        if 'sheet2' in parsed_data: self.save_sheet2_data(file_id, company_id, report_date, parsed_data['sheet2'])
        if 'sheet3' in parsed_data: self.save_sheet3_data(file_id, company_id, report_date, parsed_data['sheet3'])
        if 'sheet4' in parsed_data: self.save_sheet4_data(file_id, company_id, report_date, parsed_data['sheet4'])
        if 'sheet5' in parsed_data: self.save_sheet5_data(file_id, company_id, report_date, parsed_data['sheet5'])
        if 'sheet6' in parsed_data: self.save_sheet6_data(file_id, company_id, report_date, parsed_data['sheet6'])
        if 'sheet7' in parsed_data: self.save_sheet7_data(file_id, company_id, report_date, parsed_data['sheet7'])
        return file_id

    def get_aggregated_data(self, report_date: datetime = None, company_id: int = None) -> Dict[str, Any]:
        session = self.db.get_session()
        try:
            result = {}
            companies = session.query(Company).filter(Company.id == company_id).all() if company_id else session.query(Company).filter(Company.is_active == True).all()
            for company in companies:
                company_data = {'name': company.name, 'sheet1': [], 'sheet2': {}, 'sheet3_data': [], 'sheet4_data': [], 'sheet5_data': [], 'sheet6_data': [], 'sheet7_data': []}
                has_data = False
                
                # Sheet 1
                s1 = session.query(Sheet1Structure).filter(Sheet1Structure.company_id == company.id).order_by(Sheet1Structure.report_date.desc()).all()
                if s1:
                    last_date = s1[0].report_date
                    for item in [x for x in s1 if x.report_date == last_date]:
                        company_data['sheet1'].append({'affiliation': item.affiliation, 'company_name': item.company_name, 'oil_depots_count': item.oil_depots_count, 'azs_count': item.azs_count, 'working_azs_count': item.working_azs_count})
                        has_data = True

                # Sheet 2
                s2 = session.query(Sheet2Demand).filter(Sheet2Demand.company_id == company.id).order_by(Sheet2Demand.report_date.desc()).first()
                if s2:
                    company_data['sheet2'] = {'year': s2.report_date.year, 'gasoline_total': s2.gasoline_total, 'gasoline_ai92': s2.gasoline_ai92, 'gasoline_ai95': s2.gasoline_ai95, 'diesel_total': s2.diesel_total, 'monthly_gasoline_total': s2.monthly_gasoline_total, 'monthly_diesel_total': s2.monthly_diesel_total}
                    has_data = True

                # Sheet 3
                s3 = session.query(Sheet3Balance).filter(Sheet3Balance.company_id == company.id).order_by(Sheet3Balance.report_date.desc()).all()
                if s3:
                    last_date = s3[0].report_date
                    for item in [x for x in s3 if x.report_date == last_date]:
                        company_data['sheet3_data'].append({
                            'location_name': item.location_name, 'stock_ai92': item.stock_ai92, 'stock_ai95': item.stock_ai95, 'stock_ai98_ai100': item.stock_ai98_100,
                            'stock_diesel_winter': item.stock_diesel_winter, 'stock_diesel_arctic': item.stock_diesel_arctic, 'stock_diesel_summer': item.stock_diesel_summer,
                            'transit_ai92': item.transit_ai92, 'transit_ai95': item.transit_ai95, 'transit_ai98_ai100': item.transit_ai98_100,
                            'transit_diesel_winter': item.transit_diesel_winter, 'transit_diesel_arctic': item.transit_diesel_arctic, 'transit_diesel_summer': item.transit_diesel_summer,
                            'capacity_ai92': item.capacity_ai92, 'capacity_ai95': item.capacity_ai95, 'capacity_ai98_ai100': item.capacity_ai98_100,
                            'capacity_diesel_winter': item.capacity_diesel_winter, 'capacity_diesel_arctic': item.capacity_diesel_arctic, 'capacity_diesel_summer': item.capacity_diesel_summer,
                        })
                        has_data = True

                # Sheet 4
                s4 = session.query(Sheet4Supply).filter(Sheet4Supply.company_id == company.id).order_by(Sheet4Supply.report_date.desc()).all()
                if s4:
                    last_date = s4[0].report_date
                    for item in [x for x in s4 if x.report_date == last_date]:
                        company_data['sheet4_data'].append({'oil_depot_name': item.oil_depot_name, 'supply_date': item.supply_date, 'supply_ai92': item.supply_ai92, 'supply_ai95': item.supply_ai95, 'supply_ai98_100': item.supply_ai98_100, 'supply_diesel_winter': item.supply_diesel_winter, 'supply_diesel_arctic': item.supply_diesel_arctic, 'supply_diesel_summer': item.supply_diesel_summer})
                        has_data = True

                # Sheet 5
                s5 = session.query(Sheet5Sales).filter(Sheet5Sales.company_id == company.id).order_by(Sheet5Sales.report_date.desc()).all()
                if s5:
                    last_date = s5[0].report_date
                    for item in [x for x in s5 if x.report_date == last_date]:
                        company_data['sheet5_data'].append({
                            'location_name': item.location_name, 'daily_ai92': item.daily_ai92, 'daily_ai95': item.daily_ai95, 'daily_ai98_100': item.daily_ai98_100, 'daily_winter': item.daily_diesel_winter, 'daily_arctic': item.daily_diesel_arctic, 'daily_summer': item.daily_diesel_summer,
                            'monthly_ai92': item.monthly_ai92, 'monthly_ai95': item.monthly_ai95, 'monthly_ai98_100': item.monthly_ai98_100, 'monthly_diesel_winter': item.monthly_diesel_winter, 'monthly_diesel_arctic': item.monthly_diesel_arctic, 'monthly_diesel_summer': item.monthly_diesel_summer
                        })
                        has_data = True

                # Sheet 6
                s6 = session.query(Sheet6Aviation).filter(Sheet6Aviation.company_id == company.id).order_by(Sheet6Aviation.report_date.desc()).all()
                if s6:
                    last_date = s6[0].report_date
                    for item in [x for x in s6 if x.report_date == last_date]:
                        company_data['sheet6_data'].append({'airport_name': item.airport_name, 'tzk_name': item.tzk_name, 'contracts_info': item.contracts_info, 'supply_week': item.supply_week, 'supply_month_start': item.supply_month_start, 'monthly_demand': item.monthly_demand, 'consumption_week': item.consumption_week, 'consumption_month_start': item.consumption_month_start, 'end_of_day_balance': item.end_of_day_balance})
                        has_data = True

                # Sheet 7
                s7 = session.query(Sheet7Comments).filter(Sheet7Comments.company_id == company.id).order_by(Sheet7Comments.report_date.desc()).all()
                if s7:
                    last_date = s7[0].report_date
                    for item in [x for x in s7 if x.report_date == last_date]:
                        company_data['sheet7_data'].append({'fuel_type': item.fuel_type, 'situation': item.situation, 'comments': item.comments})
                        has_data = True

                if has_data: result[company.name] = company_data
            return result
        except Exception as e:
            print(f"Error: {e}")
            return {}
        finally:
            self.db.close_session()

    def update_file_status(self, file_id: int, status: str, error_message: str = None):
        session = self.db.get_session()
        try:
            f = session.query(UploadedFile).get(file_id)
            if f:
                f.status = status
                if error_message: f.error_message = error_message
                session.commit()
                return True
            return False
        finally:
            self.db.close_session()

db = DatabaseQueries()