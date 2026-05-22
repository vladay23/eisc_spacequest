# seed_data.py
import os
import sys
from datetime import datetime, timedelta
import random

# Добавляем текущий путь в sys.path, чтобы импортировать модули
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from database import SessionLocal, Product, Component, MassPropertiesLog, FirmwareCompatibilityLog, TestRun, TestScenario, CompatibilityMatrix
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

def seed():
    db = SessionLocal()
    try:
        # Удаляем старые данные (опционально)
        # db.query(TestRun).delete()
        # db.query(FirmwareCompatibilityLog).delete()
        # db.query(MassPropertiesLog).delete()
        # db.query(Component).delete()
        # db.query(Product).delete()
        # db.commit()

        # Создаём изделия, если их нет
        if db.query(Product).count() == 0:
            p1 = Product(serial_number="SPC-001", name="Атмосферный дрон", status="design", within_tolerance=True,
                         notes="Изделие первой серии. Проект «Атмосфера-1».")
            p2 = Product(serial_number="SPC-002", name="Марсианский планер", status="production", within_tolerance=False,
                         notes="Превышена масса по ЦМ. Требует доработки конструктором.")
            p3 = Product(serial_number="SPC-003", name="Орбитальный зонд", status="testing", within_tolerance=True,
                         notes="На испытаниях. Сценарий: термовакуум.")
            db.add_all([p1, p2, p3])
            db.commit()
            db.refresh(p1); db.refresh(p2); db.refresh(p3)

            # Компоненты
            comps = [
                Component(product_id=p1.id, name="Контроллер питания", hardware_revision="2.0", controller_type="power"),
                Component(product_id=p1.id, name="Навигационный модуль", hardware_revision="1.5", controller_type="nav"),
                Component(product_id=p2.id, name="Маршевый двигатель", hardware_revision="3.1", controller_type="engine"),
                Component(product_id=p2.id, name="Контроллер питания", hardware_revision="2.0", controller_type="power"),
                Component(product_id=p3.id, name="Навигационный модуль", hardware_revision="2.0", controller_type="nav"),
            ]
            db.add_all(comps)
            db.commit()
            for c in comps:
                db.refresh(c)
        else:
            # Если изделия уже есть, просто получаем их
            p1 = db.query(Product).filter_by(serial_number="SPC-001").first()
            p2 = db.query(Product).filter_by(serial_number="SPC-002").first()
            p3 = db.query(Product).filter_by(serial_number="SPC-003").first()
            comps = db.query(Component).all()

        # Сценарии испытаний
        if db.query(TestScenario).count() == 0:
            s1 = TestScenario(name="Вибрация", limits={"vibration": {"max": 5.0}}, description="Вибрационные испытания")
            s2 = TestScenario(name="Термовакуум", limits={"temperature": {"min": -50, "max": 100}}, description="Термовакуумные испытания")
            s3 = TestScenario(name="ЭМС", limits={"ems": {"max": 3.0}}, description="Электромагнитная совместимость")
            db.add_all([s1, s2, s3])
            db.commit()
        else:
            s1, s2, s3 = db.query(TestScenario).all()

        # Матрица совместимости
        if db.query(CompatibilityMatrix).count() == 0:
            matrix = [
                CompatibilityMatrix(component_type="power", hardware_revision="2.0", firmware_version="1.0", is_compatible=True),
                CompatibilityMatrix(component_type="power", hardware_revision="2.0", firmware_version="1.1", is_compatible=True),
                CompatibilityMatrix(component_type="power", hardware_revision="2.0", firmware_version="2.0", is_compatible=False),
                CompatibilityMatrix(component_type="nav", hardware_revision="1.5", firmware_version="2.0", is_compatible=False),
                CompatibilityMatrix(component_type="nav", hardware_revision="1.5", firmware_version="1.8", is_compatible=True),
                CompatibilityMatrix(component_type="nav", hardware_revision="2.0", firmware_version="2.1", is_compatible=True),
                CompatibilityMatrix(component_type="engine", hardware_revision="3.1", firmware_version="3.0", is_compatible=True),
                CompatibilityMatrix(component_type="engine", hardware_revision="3.1", firmware_version="2.9", is_compatible=False),
            ]
            db.add_all(matrix)
            db.commit()

        # ---- СОЗДАЁМ ИНЦИДЕНТЫ ДЛЯ ПЕЧАТИ ----
        # 1. Несовместимость прошивок (FirmwareCompatibilityLog)
        if db.query(FirmwareCompatibilityLog).count() == 0:
            # Для изделия p1 (компонент навигации, ревизия 1.5, прошивка 2.0 — несовместима)
            comp_nav_p1 = db.query(Component).filter_by(product_id=p1.id, controller_type="nav").first()
            if comp_nav_p1:
                log1 = FirmwareCompatibilityLog(
                    product_id=p1.id,
                    component_id=comp_nav_p1.id,
                    compatible=False,
                    digital_passport='{"test": true}',
                    checked_at=datetime.utcnow() - timedelta(days=2)
                )
                db.add(log1)
            # Для изделия p2 (компонент двигателя, ревизия 3.1, прошивка 2.9 — несовместима)
            comp_engine_p2 = db.query(Component).filter_by(product_id=p2.id, controller_type="engine").first()
            if comp_engine_p2:
                log2 = FirmwareCompatibilityLog(
                    product_id=p2.id,
                    component_id=comp_engine_p2.id,
                    compatible=False,
                    digital_passport='{"test": true}',
                    checked_at=datetime.utcnow() - timedelta(days=1)
                )
                db.add(log2)
            db.commit()

        # 2. Отклонения по массе (MassPropertiesLog)
        if db.query(MassPropertiesLog).count() == 0:
            # Для p2 масса 33 кг (выше нормы)
            mass_log1 = MassPropertiesLog(
                product_id=p2.id,
                mass_kg=33.0,
                cg_x=0.1, cg_y=0.2, cg_z=0.15,
                within_tolerance=False,
                material="titanium",
                volume=0.0073,
                checked_at=datetime.utcnow() - timedelta(days=3)
            )
            db.add(mass_log1)
            # Для p1 масса 29.5 кг (в норме) — не инцидент, но для полноты
            mass_log2 = MassPropertiesLog(
                product_id=p1.id,
                mass_kg=29.5,
                cg_x=0.01, cg_y=0.02, cg_z=0.01,
                within_tolerance=True,
                material="aluminum",
                volume=0.0109,
                checked_at=datetime.utcnow() - timedelta(days=5)
            )
            db.add(mass_log2)
            db.commit()

        # 3. Проваленные испытания (TestRun со статусом failed)
        if db.query(TestRun).count() == 0:
            # Для изделия p2 запустим проваленный тест вибрации
            test_run = TestRun(
                product_id=p2.id,
                scenario_id=s1.id,  # Вибрация
                status="failed",
                started_at=datetime.utcnow() - timedelta(days=1, hours=2),
                finished_at=datetime.utcnow() - timedelta(days=1),
                report_path="/reports/failed_test.docx"
            )
            db.add(test_run)
            # Успешный тест для p3 (для статистики)
            test_run2 = TestRun(
                product_id=p3.id,
                scenario_id=s2.id,
                status="passed",
                started_at=datetime.utcnow() - timedelta(days=4),
                finished_at=datetime.utcnow() - timedelta(days=4, hours=-1),
                report_path="/reports/passed_test.docx"
            )
            db.add(test_run2)
            db.commit()

        print("✅ Тестовые данные успешно добавлены.")
        print(f"   - Изделий: {db.query(Product).count()}")
        print(f"   - Компонентов: {db.query(Component).count()}")
        print(f"   - Несовместимых прошивок: {db.query(FirmwareCompatibilityLog).filter(FirmwareCompatibilityLog.compatible == False).count()}")
        print(f"   - Отклонений по массе: {db.query(MassPropertiesLog).filter(MassPropertiesLog.within_tolerance == False).count()}")
        print(f"   - Проваленных испытаний: {db.query(TestRun).filter(TestRun.status == 'failed').count()}")

    except Exception as e:
        print(f"❌ Ошибка: {e}")
        db.rollback()
    finally:
        db.close()

if __name__ == "__main__":
    seed()