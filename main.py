import os, json, hashlib, random, time, shutil
from sqlalchemy import and_
from datetime import datetime, timedelta
from contextlib import asynccontextmanager
from fastapi import FastAPI, Request, Depends, Form, File, UploadFile, HTTPException, Query
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from sqlalchemy.orm import Session, joinedload
from database import SessionLocal, Product, Component, MassPropertiesLog, FirmwareCompatibilityLog
from database import CompatibilityMatrix, TestScenario, TestRun, TelemetryPoint, DocTemplate
from docx_generator import (
    generate_mass_report, generate_firmware_passport, generate_incidents_report,
    generate_test_act, generate_analytics_report, generate_product_summary
)

# ✅ Исправлено: __file__ вместо file
REPORTS_DIR = os.path.join(os.path.dirname(__file__), "reports")
TEMPLATES_STORE = os.path.join(os.path.dirname(__file__), "db_templates")
os.makedirs(REPORTS_DIR, exist_ok=True)
os.makedirs(TEMPLATES_STORE, exist_ok=True)

# ✅ Безопасный lifespan вместо @app.on_event
@asynccontextmanager
async def lifespan(app: FastAPI):
    db = SessionLocal()
    try:
        init_demo_data(db)
    finally:
        db.close()
    yield

app = FastAPI(title="ЕИС SpaceQuest", lifespan=lifespan)

templates = Jinja2Templates(directory="templates")

if os.path.exists("static"):
    app.mount("/static", StaticFiles(directory="static"), name="static")

# ─── DB ───────────────────────────────────────────────────────────────────────
def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

# ─── Демо-данные ──────────────────────────────────────────────────────────────
def init_demo_data(db: Session):
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
        
        comps = [
            Component(product_id=p1.id, name="Контроллер питания", hardware_revision="2.0", controller_type="power"),
            Component(product_id=p1.id, name="Навигационный модуль", hardware_revision="1.5", controller_type="nav"),
            Component(product_id=p2.id, name="Маршевый двигатель", hardware_revision="3.1", controller_type="engine"),
            Component(product_id=p2.id, name="Контроллер питания", hardware_revision="2.0", controller_type="power"),
            Component(product_id=p3.id, name="Навигационный модуль", hardware_revision="2.0", controller_type="nav"),
        ]
        db.add_all(comps)
        db.commit()
        
    if db.query(TestScenario).count() == 0:
        s1 = TestScenario(name="Вибрация", limits={"vibration": {"max": 5.0}},
                          description="Вибрационные испытания в диапазоне 20–2000 Гц")
        s2 = TestScenario(name="Термовакуум", limits={"temperature": {"min": -50, "max": 100}},
                          description="Термовакуумные испытания по ГОСТ 16962.2")
        s3 = TestScenario(name="ЭМС", limits={"ems": {"max": 3.0}},
                          description="Электромагнитная совместимость по ГОСТ Р 50414")
        db.add_all([s1, s2, s3])
        db.commit()
        
        # --- Демо-инциденты для отчётов ---
    if db.query(MassPropertiesLog).count() == 0:
        p2 = db.query(Product).filter_by(serial_number="SPC-002").first()
        if p2:
            mass_log = MassPropertiesLog(
                product_id=p2.id,
                mass_kg=33.0,
                cg_x=0.1, cg_y=0.2, cg_z=0.15,
                within_tolerance=False,
                material="titanium",
                volume=0.0073,
                checked_at=datetime.utcnow()
            )
            db.add(mass_log)
            p2.within_tolerance = False
            p2.mass_kg = 33.0
            db.add(p2)
            db.commit()

    if db.query(FirmwareCompatibilityLog).count() == 0:
        p1 = db.query(Product).filter_by(serial_number="SPC-001").first()
        if p1:
            comp_nav = db.query(Component).filter_by(product_id=p1.id, controller_type="nav").first()
            if comp_nav:
                fw_log = FirmwareCompatibilityLog(
                    product_id=p1.id,
                    component_id=comp_nav.id,
                    compatible=False,
                    digital_passport='{"test": true}',
                    checked_at=datetime.utcnow()
                )
                db.add(fw_log)
                comp_nav.firmware_version = "2.0"
                db.add(comp_nav)
                db.commit()

    if db.query(TestRun).filter(TestRun.status == "failed").count() == 0:
        p2 = db.query(Product).filter_by(serial_number="SPC-002").first()
        scenario_vibro = db.query(TestScenario).filter_by(name="Вибрация").first()
        if p2 and scenario_vibro:
            test_run = TestRun(
                product_id=p2.id,
                scenario_id=scenario_vibro.id,
                status="failed",
                started_at=datetime.utcnow() - timedelta(days=1),
                finished_at=datetime.utcnow() - timedelta(hours=2),
                report_path="/reports/failed_vibration.docx"
            )
            db.add(test_run)
            db.commit()    
        
    if db.query(CompatibilityMatrix).count() == 0:
        db.add_all([
            CompatibilityMatrix(component_type="power", hardware_revision="2.0", firmware_version="1.0", is_compatible=True),
            CompatibilityMatrix(component_type="power", hardware_revision="2.0", firmware_version="1.1", is_compatible=True),
            CompatibilityMatrix(component_type="power", hardware_revision="2.0", firmware_version="2.0", is_compatible=False),
            CompatibilityMatrix(component_type="nav", hardware_revision="1.5", firmware_version="2.0", is_compatible=False),
            CompatibilityMatrix(component_type="nav", hardware_revision="1.5", firmware_version="1.8", is_compatible=True),
            CompatibilityMatrix(component_type="nav", hardware_revision="2.0", firmware_version="2.1", is_compatible=True),
            CompatibilityMatrix(component_type="engine", hardware_revision="3.1", firmware_version="3.0", is_compatible=True),
            CompatibilityMatrix(component_type="engine", hardware_revision="3.1", firmware_version="2.9", is_compatible=False),
        ])
        db.commit()

# ─── ДАШБОРД ─────────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request, db: Session = Depends(get_db)):
    products = db.query(Product).all()
    stats = {
        "total": len(products),
        "design": len([p for p in products if p.status == "design"]),
        "production": len([p for p in products if p.status == "production"]),
        "testing": len([p for p in products if p.status == "testing"]),
        "accepted": len([p for p in products if p.status == "accepted"]),
        "out_of_tol": len([p for p in products if not p.within_tolerance]),
    }
    failed_tests = db.query(TestRun).filter(TestRun.status == "failed").count()
    compatibility_issues = db.query(FirmwareCompatibilityLog).filter(FirmwareCompatibilityLog.compatible == False).count()
    total_tests = db.query(TestRun).count()
    recent = products[-5:] 
    return templates.TemplateResponse(
        request=request,
        name="dashboard.html",
        context={
            "stats": stats,
            "failed_tests": failed_tests,
            "compatibility_issues": compatibility_issues,
            "total_tests": total_tests,
            "products": recent,
        }
    )

# ─── СПИСОК ИЗДЕЛИЙ ──────────────────────────────────────────────────────────
@app.get("/assets", response_class=HTMLResponse)
async def list_assets(request: Request, db: Session = Depends(get_db)):
    products = db.query(Product).all()
    return templates.TemplateResponse(  # ✅ Добавлен return
        request=request,
        name="assets.html",
        context={"products": products}
    )

@app.get("/assets/new", response_class=HTMLResponse)
async def new_asset_form(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="asset_new.html",
    )

@app.post("/assets/new")
async def create_asset(
    request: Request,
    serial_number: str = Form(...),
    name: str = Form(...),
    status: str = Form("design"),
    notes: str = Form(""),
    db: Session = Depends(get_db),
):
    existing = db.query(Product).filter(Product.serial_number == serial_number).first()
    if existing:
        return templates.TemplateResponse("asset_new.html", {
            "request": request, "error": f"Серийный номер {serial_number} уже существует."
        })
    p = Product(serial_number=serial_number, name=name, status=status, notes=notes)
    db.add(p)
    db.commit()
    return templates.TemplateResponse(
        request=request,
        name="asset_new.html",
        context={"success": f"Изделие {name} (С/Н: {serial_number}) создано.", "product_id": p.id}
    )

@app.get("/assets/{product_id}", response_class=HTMLResponse)
async def asset_detail(product_id: int, request: Request, db: Session = Depends(get_db)):
    product = db.query(Product).filter(Product.id == product_id).first()  # ✅ product_i d → product_id
    if not product:
        return templates.TemplateResponse("error.html", {"request": request, "message": "Изделие не найдено"}, status_code=404)
    components = db.query(Component).filter(Component.product_id == product_id).all()
    tests = db.query(TestRun).filter(TestRun.product_id == product_id).order_by(TestRun.started_at.desc()).all()  # ✅ T estRun → TestRun
    logs = db.query(FirmwareCompatibilityLog).filter(FirmwareCompatibilityLog.product_id == product_id).order_by(FirmwareCompatibilityLog.checked_at.desc()).all()  # ✅ des c() → desc()
    mass_logs = db.query(MassPropertiesLog).filter(MassPropertiesLog.product_id == product_id).order_by(MassPropertiesLog.checked_at.desc()).all()
    return templates.TemplateResponse(
        request=request,
        name="product_detail.html",
        context={"product": product, "components": components,
        "tests": tests, "compatibility_logs": logs, "mass_logs": mass_logs,}
    )

# ─── КОМПОНЕНТЫ ──────────────────────────────────────────────────────────────
@app.get("/assets/{product_id}/components/new", response_class=HTMLResponse)
async def new_component_form(product_id: int, request: Request, db: Session = Depends(get_db)):
    product = db.query(Product).filter(Product.id == product_id).first()
    return templates.TemplateResponse(
        request=request,
        name="component_new.html",
        context={"request": request, "product": product}
    )

@app.post("/assets/{product_id}/components/new")
async def create_component(
    product_id: int, request: Request,
    name: str = Form(...),
    hardware_revision: str = Form(...),
    controller_type: str = Form(...),
    db: Session = Depends(get_db),
):
    c = Component(product_id=product_id, name=name, hardware_revision=hardware_revision, controller_type=controller_type)
    db.add(c)
    db.commit()
    product = db.query(Product).filter(Product.id == product_id).first()
    return templates.TemplateResponse(
        request=request,
        name="component_new.html",
        context={"product": product,
        "success": f"Компонент «{name}» добавлен."}
    )

# ─── ИНЦИДЕНТЫ ───────────────────────────────────────────────────────────────
@app.get("/incidents", response_class=HTMLResponse)
async def list_incidents(request: Request, db: Session = Depends(get_db)):
    incidents = db.query(FirmwareCompatibilityLog).filter(FirmwareCompatibilityLog.compatible == False).all()  # ✅ FirmwareCompatibili tyLog → FirmwareCompatibilityLog
    mass_incidents = db.query(MassPropertiesLog).filter(MassPropertiesLog.within_tolerance == False).all()
    test_failures = db.query(TestRun).filter(TestRun.status == "failed").all()  # ✅ Tes tRun → TestRun
    return templates.TemplateResponse(
        request=request,
        name="incidents.html",
        context={"firmware_incidents": incidents,
        "mass_incidents": mass_incidents,
        "test_failures": test_failures,}
    )
    
@app.get("/incidents/print")
async def print_incidents(db: Session = Depends(get_db)):
    firmware_incidents = db.query(FirmwareCompatibilityLog).options(
        joinedload(FirmwareCompatibilityLog.product),
        joinedload(FirmwareCompatibilityLog.component)
    ).filter(FirmwareCompatibilityLog.compatible == False).all()
    mass_incidents = db.query(MassPropertiesLog).options(
        joinedload(MassPropertiesLog.product)
    ).filter(MassPropertiesLog.within_tolerance == False).all()
    test_failures = db.query(TestRun).options(
        joinedload(TestRun.product),
        joinedload(TestRun.scenario)
    ).filter(TestRun.status == "failed").all()
    
    filepath = generate_incidents_report(firmware_incidents, mass_incidents, test_failures)
    return FileResponse(filepath, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename=os.path.basename(filepath))

# ─── АНАЛИТИКА ───────────────────────────────────────────────────────────────
@app.get("/analytics", response_class=HTMLResponse)
async def analytics(
    request: Request,
    db: Session = Depends(get_db),
    date_from: str = Query(None),
    date_to: str = Query(None),
    status: str = Query(None),
    issue_type: str = Query(None)
):
    # ----- 1. Базовый запрос изделий с фильтрацией -----
    query = db.query(Product)
    
    if status:
        query = query.filter(Product.status == status)
    
    # Фильтр по дате создания (поле created_at есть в Product)
    if date_from:
        date_from_dt = datetime.strptime(date_from, '%Y-%m-%d')
        query = query.filter(Product.created_at >= date_from_dt)
    if date_to:
        date_to_dt = datetime.strptime(date_to, '%Y-%m-%d').replace(hour=23, minute=59, second=59)
        query = query.filter(Product.created_at <= date_to_dt)
    
    products = query.all()
    product_ids = [p.id for p in products]
    total_products = len(products)

    # ----- 2. Массовые отклонения (через MassPropertiesLog) -----
    in_tol = 0
    out_of_tol = 0
    for p in products:
        # последняя запись массы для изделия
        last_mass = db.query(MassPropertiesLog).filter(
            MassPropertiesLog.product_id == p.id
        ).order_by(MassPropertiesLog.checked_at.desc()).first()
        if last_mass and last_mass.mass_kg is not None:
            # Допуск по массе: 28.5 – 31.5 кг
            if 28.5 <= last_mass.mass_kg <= 31.5:
                in_tol += 1
            else:
                out_of_tol += 1
        else:
            # Нет данных по массе – считаем нарушением
            out_of_tol += 1

    # Если выбран фильтр "Отклонения по массе" – показываем только изделия вне допуска
    if issue_type == 'mass':
        total_products = out_of_tol
        in_tol = 0

    # ----- 3. Испытания (модель TestRun) -----
    tests_query = db.query(TestRun).filter(TestRun.product_id.in_(product_ids))
    if issue_type == 'test':
        tests_query = tests_query.filter(TestRun.status == 'failed')
    all_tests = tests_query.all()
    passed_tests = sum(1 for t in all_tests if t.status == 'passed')
    failed_tests = sum(1 for t in all_tests if t.status == 'failed')

    # ----- 4. Совместимость ПО (модель FirmwareCompatibilityLog) -----
    compat_query = db.query(FirmwareCompatibilityLog).filter(
        FirmwareCompatibilityLog.product_id.in_(product_ids)
    )
    if issue_type == 'firmware':
        compat_query = compat_query.filter(FirmwareCompatibilityLog.compatible == False)
    all_compat = compat_query.all()
    compatibility_ok = sum(1 for c in all_compat if c.compatible)
    compatibility_issues = sum(1 for c in all_compat if not c.compatible)

    # ---------------------- 5. Отправляем данные в шаблон ---------------------
    return templates.TemplateResponse(
        request=request,
        name="analytics.html",
        context={
            "date_from": date_from or "",
            "date_to": date_to or "",
            "status": status or "",
            "issue_type": issue_type or "",
            "total_products": total_products,
            "in_tol": in_tol,
            "out_of_tol": out_of_tol,
            "passed_tests": passed_tests,
            "failed_tests": failed_tests,
            "compatibility_ok": compatibility_ok,
            "compatibility_issues": compatibility_issues,}
    )
    
@app.get("/analytics/print")
async def analytics_print(
    request: Request,
    db: Session = Depends(get_db),
    date_from: str = Query(None),
    date_to: str = Query(None),
    status: str = Query(None),
    issue_type: str = Query(None)
):
    # Полностью копируем логику основного маршрута /analytics
    query = db.query(Product)
    if status:
        query = query.filter(Product.status == status)
    if date_from:
        date_from_dt = datetime.strptime(date_from, '%Y-%m-%d')
        query = query.filter(Product.created_at >= date_from_dt)
    if date_to:
        date_to_dt = datetime.strptime(date_to, '%Y-%m-%d').replace(hour=23, minute=59, second=59)
        query = query.filter(Product.created_at <= date_to_dt)
    
    products = query.all()
    product_ids = [p.id for p in products]
    total_products = len(products)
    
    # Масса
    in_tol = 0
    out_of_tol = 0
    for p in products:
        last_mass = db.query(MassPropertiesLog).filter(MassPropertiesLog.product_id == p.id).order_by(MassPropertiesLog.checked_at.desc()).first()
        if last_mass and last_mass.mass_kg is not None:
            if 28.5 <= last_mass.mass_kg <= 31.5:
                in_tol += 1
            else:
                out_of_tol += 1
        else:
            out_of_tol += 1
    if issue_type == 'mass':
        total_products = out_of_tol
        in_tol = 0
    
    # Испытания
    tests_query = db.query(TestRun).filter(TestRun.product_id.in_(product_ids))
    if issue_type == 'test':
        tests_query = tests_query.filter(TestRun.status == 'failed')
    all_tests = tests_query.all()
    passed_tests = sum(1 for t in all_tests if t.status == 'passed')
    failed_tests = sum(1 for t in all_tests if t.status == 'failed')
    
    # Совместимость ПО
    compat_query = db.query(FirmwareCompatibilityLog).filter(FirmwareCompatibilityLog.product_id.in_(product_ids))
    if issue_type == 'firmware':
        compat_query = compat_query.filter(FirmwareCompatibilityLog.compatible == False)
    all_compat = compat_query.all()
    compatibility_ok = sum(1 for c in all_compat if c.compatible)
    compatibility_issues = sum(1 for c in all_compat if not c.compatible)
    
    # Индекс готовности
    total_checks = (in_tol + out_of_tol) + (passed_tests + failed_tests) + (compatibility_ok + compatibility_issues)
    passed_checks = in_tol + passed_tests + compatibility_ok
    readiness = int(passed_checks / total_checks * 100) if total_checks > 0 else 0
    readiness_text = "Требуется устранение замечаний перед переходом на следующий этап ЖЦ." if readiness < 100 else "Изделие полностью готово к передаче заказчику."
    
    filters = {
        'date_from': date_from,
        'date_to': date_to,
        'status': status,
        'issue_type': issue_type
    }
    
    filepath = generate_analytics_report(
        total_products, in_tol, out_of_tol,
        passed_tests, failed_tests,
        compatibility_ok, compatibility_issues,
        readiness, readiness_text,
        filters
    )
    return FileResponse(
        filepath,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=os.path.basename(filepath)
    )

# ─── ИМПОРТ ──────────────────────────────────────────────────────────────────
@app.get("/import", response_class=HTMLResponse)
async def import_form(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="import.html"
    )

@app.post("/import")
async def do_import(request: Request, file: UploadFile = File(...), db: Session = Depends(get_db)):
    content = await file.read()
    try:
        data = json.loads(content)
        count = 0
        for item in data.get("products", []):  # ✅ ite m → item
            if not db.query(Product).filter(Product.serial_number == item["serial_number"]).first():
                prod = Product(
                    serial_number=item["serial_number"],
                    name=item["name"],
                    status=item.get("status", "design"),
                    within_tolerance=item.get("within_tolerance", True),
                )
                db.add(prod)
                count += 1
        db.commit()
        message = f"Успешно импортировано {count} изделий."
    except Exception as e:
        message = f"Ошибка импорта: {str(e)}"
    return templates.TemplateResponse(
        request=request,
        name="import.html",
        context={"message": message,
        }
    )

# ─── ПОДСИСТЕМА 1: Управление массой ─────────────────────────────────────────
@app.get("/product/{product_id}/mass_simulate", response_class=HTMLResponse)
async def mass_simulate_form(product_id: int, request: Request, db: Session = Depends(get_db)):
    product = db.query(Product).filter(Product.id == product_id).first()
    mass_logs = db.query(MassPropertiesLog).filter(MassPropertiesLog.product_id == product_id).order_by(MassPropertiesLog.checked_at.desc()).limit(5).all()
    return templates.TemplateResponse(
        request=request,
        name="mass_simulate.html",
        context={"product": product, "mass_logs": mass_logs}
    )

@app.post("/product/{product_id}/mass_simulate")
async def mass_simulate_post(product_id: int, request: Request, db: Session = Depends(get_db)):
    form = await request.form()
    material = form.get("material")
    volume = float(form.get("volume", 0))
    cg_x = float(form.get("cg_x", 0))
    cg_y = float(form.get("cg_y", 0))
    cg_z = float(form.get("cg_z", 0))
    density = {"aluminum": 2700, "titanium": 4500, "composite": 1600}.get(material, 2700)
    mass = volume * density
    ref_mass = 30.0
    within_mass = abs(mass - ref_mass) / ref_mass <= 0.05 if ref_mass > 0 else True
    within_cg = abs(cg_x) < 0.05 and abs(cg_y) < 0.05 and abs(cg_z) < 0.05
    within = within_mass and within_cg
    log = MassPropertiesLog(
        product_id=product_id, mass_kg=mass,
        cg_x=cg_x, cg_y=cg_y, cg_z=cg_z,
        within_tolerance=within, material=material, volume=volume,
    )
    db.add(log)
    product = db.query(Product).filter(Product.id == product_id).first()
    product.mass_kg = mass
    product.cg_x, product.cg_y, product.cg_z = cg_x, cg_y, cg_z  # ✅ cg_ z → cg_z
    product.within_tolerance = within
    db.commit()
    db.refresh(log)
    return templates.TemplateResponse(
        request=request,
        name="mass_result.html",
        context={"mass": mass, "cg_x": cg_x, "cg_y": cg_y, "cg_z": cg_z,
        "within": within, "product_id": product_id, "product": product,
        "log_id": log.id, "material": material,
        }
    )

# ─── Скачать отчёт по массе ───────────────────────────────────────────────────
@app.get("/product/{product_id}/mass_report/{log_id}")
async def download_mass_report(product_id: int, log_id: int, db: Session = Depends(get_db)):
    product = db.query(Product).filter(Product.id == product_id).first()
    log = db.query(MassPropertiesLog).filter(MassPropertiesLog.id == log_id).first()
    if not product or not log:
        raise HTTPException(status_code=404, detail="Запись не найдена")
    tmpl = db.query(DocTemplate).filter(
        DocTemplate.template_type == "mass_report", DocTemplate.is_active == True
    ).first()
    tpath = tmpl.file_path if tmpl and os.path.exists(tmpl.file_path) else None
    filepath = generate_mass_report(product, log, template_path=tpath)
    return FileResponse(
        filepath,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=os.path.basename(filepath),
    )

# ─── ПОДСИСТЕМА 2: Бортовое ПО ───────────────────────────────────────────────
@app.get("/product/{product_id}/firmware_simulate", response_class=HTMLResponse)
async def firmware_simulate_form(product_id: int, request: Request, db: Session = Depends(get_db)):
    product = db.query(Product).filter(Product.id == product_id).first()
    components = db.query(Component).filter(Component.product_id == product_id).all()
    matrix_versions = {}
    for c in components:  # ✅ f or → for
        versions = db.query(CompatibilityMatrix).filter(
            CompatibilityMatrix.component_type == c.controller_type,
            CompatibilityMatrix.hardware_revision == c.hardware_revision,  # ✅ hardware_re vision → hardware_revision
        ).all()
        matrix_versions[c.id] = [
            {"version": v.firmware_version, "compatible": v.is_compatible} for v in versions
        ]
    return templates.TemplateResponse(
        request=request,
        name="firmware_simulate.html",
        context={"product": product, "components": components,
        "matrix_versions": json.dumps(matrix_versions),
        }
    )

@app.post("/product/{product_id}/firmware_simulate")
async def firmware_simulate_post(product_id: int, request: Request, db: Session = Depends(get_db)):
    form = await request.form()
    component_id = int(form.get("component_id"))
    fw_version = form.get("firmware_version")
    fw_hash = form.get("hash") or hashlib.sha256(f"{component_id}{fw_version}{time.time()}".encode()).hexdigest()
    comp = db.query(Component).filter(Component.id == component_id).first()
    matrix = db.query(CompatibilityMatrix).filter(
        CompatibilityMatrix.component_type == comp.controller_type,
        CompatibilityMatrix.hardware_revision == comp.hardware_revision,
        CompatibilityMatrix.firmware_version == fw_version,
    ).first()
    compatible = matrix.is_compatible if matrix else False  # ✅ is_compati ble → is_compatible
    passport = {
        "product_id": product_id, "component": comp.name,
        "hardware_revision": comp.hardware_revision,
        "firmware_version": fw_version,
        "hash": fw_hash,
        "compatible": compatible,
        "timestamp": datetime.utcnow().isoformat(),
    }
    passport_str = json.dumps(passport, indent=2)
    signature = hashlib.sha256(passport_str.encode()).hexdigest()
    passport["signature"] = signature
    full_passport = json.dumps(passport, indent=2, ensure_ascii=False)
    log = FirmwareCompatibilityLog(
        product_id=product_id, component_id=component_id,
        compatible=compatible, digital_passport=full_passport,  # ✅ compa tible → compatible
    )
    db.add(log)
    comp.firmware_version = fw_version
    comp.firmware_hash = fw_hash
    db.commit()
    db.refresh(log)
    return templates.TemplateResponse(
        request=request,
        name="firmware_result.html",
        context={"comp": comp, "fw_version": fw_version, "fw_hash": fw_hash,
        "compatible": compatible, "passport": full_passport,
        "product_id": product_id, "log_id": log.id,
        }
    )

# ─── Скачать паспорт прошивки ────────────────────────────────────────────────
@app.get("/product/{product_id}/firmware_passport/{log_id}")
async def download_firmware_passport(product_id: int, log_id: int, db: Session = Depends(get_db)):
    product = db.query(Product).filter(Product.id == product_id).first()
    log = db.query(FirmwareCompatibilityLog).filter(FirmwareCompatibilityLog.id == log_id).first()
    if not product or not log:
        raise HTTPException(status_code=404, detail="Запись не найдена")
    component = db.query(Component).filter(Component.id == log.component_id).first()
    tmpl = db.query(DocTemplate).filter(
        DocTemplate.template_type == "firmware_passport", DocTemplate.is_active == True
    ).first()
    tpath = tmpl.file_path if tmpl and os.path.exists(tmpl.file_path) else None
    filepath = generate_firmware_passport(product, component, log, template_path=tpath)
    return FileResponse(
        filepath,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=os.path.basename(filepath),
    )

# ─── ПОДСИСТЕМА 3: Испытательная станция ─────────────────────────────────────
@app.get("/product/{product_id}/test_simulate", response_class=HTMLResponse)
async def test_simulate_form(product_id: int, request: Request, db: Session = Depends(get_db)):
    product = db.query(Product).filter(Product.id == product_id).first()
    scenarios = db.query(TestScenario).all()
    return templates.TemplateResponse(
        request=request,
        name="test_simulate.html",
        context={"request": request, "product": product, "scenarios": scenarios,
        }
    )

@app.post("/product/{product_id}/test_simulate")
async def test_simulate_post(product_id: int, request: Request, db: Session = Depends(get_db)):
    form = await request.form()
    scenario_id = int(form.get("scenario_id"))
    scenario = db.query(TestScenario).filter(TestScenario.id == scenario_id).first()
    test_run = TestRun(product_id=product_id, scenario_id=scenario_id,
                       status="running", started_at=datetime.utcnow())
    db.add(test_run)
    db.commit()
    limits = scenario.limits
    telemetry_points = []
    for _ in range(8):
        for param, limit in limits.items():
            if param == "vibration":
                value = random.uniform(0, 7)
                within = value <= limit.get("max", 5)
            elif param == "temperature":
                value = random.uniform(-65, 115)
                within = limit.get("min", -50) <= value <= limit.get("max", 100)
            elif param == "ems":
                value = random.uniform(0, 5)
                within = value <= limit.get("max", 3)
            else:
                value = 0; within = True
            tp = TelemetryPoint(test_run_id=test_run.id, parameter=param, value=value, within_limit=within)
            db.add(tp)
            telemetry_points.append((param, value, within))  # ✅ valu e → value
    db.commit()
    all_within = all(w for (_, _, w) in telemetry_points)  # ✅ (_, , w) → (_, _, w)
    test_run.status = "passed" if all_within else "failed"
    test_run.finished_at = datetime.utcnow()
    test_run.report_path = f"/reports/test_{test_run.id}.docx"  # ✅ f-string исправлен
    db.commit()
    product = db.query(Product).filter(Product.id == product_id).first()
    if all_within and product.status == "testing":
        product.status = "accepted"
        db.commit()
    return templates.TemplateResponse(
        request=request,
        name="test_result.html",
        context={"request": request, "scenario": scenario, "all_within": all_within,
        "telemetry_points": telemetry_points, "product_id": product_id,
        "test_run": test_run, "product": product,
        }
    )

# ─── Скачать акт испытаний ───────────────────────────────────────────────────
@app.get("/product/{product_id}/test_act/{run_id}")
async def download_test_act(product_id: int, run_id: int, db: Session = Depends(get_db)):
    product = db.query(Product).filter(Product.id == product_id).first()
    test_run = db.query(TestRun).filter(TestRun.id == run_id).first()
    if not product or not test_run:
        raise HTTPException(status_code=404, detail="Запись не найдена")
    scenario = db.query(TestScenario).filter(TestScenario.id == test_run.scenario_id).first()
    telemetry = db.query(TelemetryPoint).filter(TelemetryPoint.test_run_id == run_id).all()
    points = [(t.parameter, t.value, t.within_limit) for t in telemetry]
    tmpl = db.query(DocTemplate).filter(
        DocTemplate.template_type == "test_act", DocTemplate.is_active == True
    ).first()
    tpath = tmpl.file_path if tmpl and os.path.exists(tmpl.file_path) else None
    filepath = generate_test_act(product, test_run, scenario, points, template_path=tpath)
    return FileResponse(
        filepath,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=os.path.basename(filepath),
    )

# ─── Сводный отчёт ───────────────────────────────────────────────────────────
@app.get("/product/{product_id}/summary_report")
async def download_summary_report(product_id: int, db: Session = Depends(get_db)):
    product = db.query(Product).filter(Product.id == product_id).first()
    if not product:
        raise HTTPException(status_code=404, detail="Изделие не найдено")
    components = db.query(Component).filter(Component.product_id == product_id).all()
    mass_logs = db.query(MassPropertiesLog).filter(MassPropertiesLog.product_id == product_id).order_by(MassPropertiesLog.checked_at.desc()).all()
    fw_logs = db.query(FirmwareCompatibilityLog).filter(FirmwareCompatibilityLog.product_id == product_id).order_by(FirmwareCompatibilityLog.checked_at.desc()).all()
    test_runs = db.query(TestRun).filter(TestRun.product_id == product_id).order_by(TestRun.started_at.desc()).all()
    filepath = generate_product_summary(product, components, mass_logs, fw_logs, test_runs)
    return FileResponse(
        filepath,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=os.path.basename(filepath),
    )

# ─── Управление шаблонами WORD ──────────────────────────────────────────────
@app.get("/templates", response_class=HTMLResponse)
async def templates_list(request: Request, db: Session = Depends(get_db)):
    doc_templates = db.query(DocTemplate).order_by(DocTemplate.uploaded_at.desc()).all()
    return templates.TemplateResponse(
        request=request,
        name="doc_templates.html",
        context={"doc_templates": doc_templates,
        }
    )

@app.post("/templates/upload")
async def upload_template(
    request: Request,
    name: str = Form(...),
    template_type: str = Form(...),
    description: str = Form(""),
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
):
    if not file.filename.endswith(".docx"):
        return templates.TemplateResponse(
        request=request,
        name="doc_templates.html",
        context={"doc_templates": db.query(DocTemplate).all(),
                 "error": "Допустимы только файлы .docx",
            }
        )
    old = db.query(DocTemplate).filter(
        DocTemplate.template_type == template_type, DocTemplate.is_active == True
    ).all()
    for o in old:  # ✅ old : → old:
        o.is_active = False
    db.commit()
    filename = f"{template_type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    save_path = os.path.join(TEMPLATES_STORE, filename)
    content = await file.read()
    with open(save_path, "wb") as f:
        f.write(content)
    tmpl = DocTemplate(
        name=name, template_type=template_type,
        description=description, file_path=save_path, is_active=True,
    )
    db.add(tmpl)
    db.commit()
    doc_templates = db.query(DocTemplate).order_by(DocTemplate.uploaded_at.desc()).all()  # ✅ doc_t emplates → doc_templates
    return templates.TemplateResponse(
        request=request,
        name="doc_templates.html",
        context={"doc_templates": doc_templates,
                 "success": f"Шаблон «{name}» загружен и активирован.",
        }
    )

@app.post("/templates/{tmpl_id}/deactivate")
async def deactivate_template(tmpl_id: int, request: Request, db: Session = Depends(get_db)):
    tmpl = db.query(DocTemplate).filter(DocTemplate.id == tmpl_id).first()
    if tmpl:
        tmpl.is_active = False
        db.commit()
    return templates.TemplateResponse("doc_templates.html", {
        "request": request,
        "doc_templates": db.query(DocTemplate).order_by(DocTemplate.uploaded_at.desc()).all(),
        "success": "Шаблон деактивирован.",
    })

@app.get("/templates/{tmpl_id}/download")
async def download_template(tmpl_id: int, db: Session = Depends(get_db)):
    tmpl = db.query(DocTemplate).filter(DocTemplate.id == tmpl_id).first()
    if not tmpl or not os.path.exists(tmpl.file_path):
        raise HTTPException(status_code=404, detail="Шаблон не найден")
    return FileResponse(
        tmpl.file_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"{tmpl.name}.docx",
    )
    
@app.get("/product/{product_id}/digital_passport")
async def download_digital_passport(product_id: int, db: Session = Depends(get_db)):
    """Генерирует и скачивает Цифровой паспорт изделия (сводный отчёт)"""
    product = db.query(Product).filter(Product.id == product_id).first()
    if not product:
        raise HTTPException(status_code=404, detail="Изделие не найдено")
    
    components = db.query(Component).filter(Component.product_id == product_id).all()
    mass_logs = db.query(MassPropertiesLog).filter(MassPropertiesLog.product_id == product_id).order_by(MassPropertiesLog.checked_at.desc()).all()
    fw_logs = db.query(FirmwareCompatibilityLog).filter(FirmwareCompatibilityLog.product_id == product_id).order_by(FirmwareCompatibilityLog.checked_at.desc()).all()
    test_runs = db.query(TestRun).filter(TestRun.product_id == product_id).order_by(TestRun.started_at.desc()).all()
    
    filepath = generate_product_summary(product, components, mass_logs, fw_logs, test_runs)
    return FileResponse(
        filepath,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=os.path.basename(filepath)
    )
    
@app.get("/product/{product_id}/test_act_latest")
async def download_latest_test_act(product_id: int, db: Session = Depends(get_db)):
    product = db.query(Product).filter(Product.id == product_id).first()
    if not product:
        raise HTTPException(status_code=404, detail="Изделие не найдено")
    
    latest_test = db.query(TestRun).filter(
        TestRun.product_id == product_id,
        TestRun.status.in_(['passed', 'failed'])
    ).order_by(TestRun.finished_at.desc()).first()
    
    if not latest_test:
        raise HTTPException(status_code=404, detail="Нет завершённых испытаний для этого изделия")
    
    scenario = db.query(TestScenario).filter(TestScenario.id == latest_test.scenario_id).first()
    telemetry = db.query(TelemetryPoint).filter(TelemetryPoint.test_run_id == latest_test.id).all()
    points = [(t.parameter, t.value, t.within_limit) for t in telemetry]
    
    filepath = generate_test_act(product, latest_test, scenario, points)
    return FileResponse(
        filepath,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=os.path.basename(filepath)
    )

# ─── HEALTH ──────────────────────────────────────────────────────────────────
@app.get("/health")
def health():
    return {"status": "ok"}