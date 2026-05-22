from sqlalchemy import create_engine, Column, Integer, String, Float, Boolean, DateTime, ForeignKey, Text, JSON
from sqlalchemy.orm import declarative_base, sessionmaker, relationship
from datetime import datetime

DATABASE_URL = "sqlite:///./spacequest.db"
engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


class Product(Base):
    __tablename__ = "products"
    id = Column(Integer, primary_key=True, index=True)
    serial_number = Column(String, unique=True, index=True)
    name = Column(String)
    status = Column(String, default="design")  # design, production, testing, accepted
    within_tolerance = Column(Boolean, default=True)
    mass_kg = Column(Float, nullable=True)
    cg_x = Column(Float, nullable=True)
    cg_y = Column(Float, nullable=True)
    cg_z = Column(Float, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    notes = Column(Text, nullable=True)

    components = relationship("Component", back_populates="product")
    mass_logs = relationship("MassPropertiesLog", back_populates="product")
    firmware_logs = relationship("FirmwareCompatibilityLog", back_populates="product")
    test_runs = relationship("TestRun", back_populates="product")


class Component(Base):
    __tablename__ = "components"
    id = Column(Integer, primary_key=True, index=True)
    product_id = Column(Integer, ForeignKey("products.id"))
    name = Column(String)
    hardware_revision = Column(String)
    controller_type = Column(String)
    firmware_version = Column(String, nullable=True)
    firmware_hash = Column(String, nullable=True)

    product = relationship("Product", back_populates="components")
    firmware_logs = relationship("FirmwareCompatibilityLog", back_populates="component")


class MassPropertiesLog(Base):
    __tablename__ = "mass_properties_logs"
    id = Column(Integer, primary_key=True, index=True)
    product_id = Column(Integer, ForeignKey("products.id"))
    mass_kg = Column(Float)
    cg_x = Column(Float)
    cg_y = Column(Float)
    cg_z = Column(Float)
    within_tolerance = Column(Boolean)
    material = Column(String, nullable=True)
    volume = Column(Float, nullable=True)
    checked_at = Column(DateTime, default=datetime.utcnow)

    product = relationship("Product", back_populates="mass_logs")


class FirmwareCompatibilityLog(Base):
    __tablename__ = "firmware_compatibility_logs"
    id = Column(Integer, primary_key=True, index=True)
    product_id = Column(Integer, ForeignKey("products.id"))
    component_id = Column(Integer, ForeignKey("components.id"))
    compatible = Column(Boolean)
    digital_passport = Column(Text, nullable=True)
    checked_at = Column(DateTime, default=datetime.utcnow)

    product = relationship("Product", back_populates="firmware_logs")
    component = relationship("Component", back_populates="firmware_logs")


class CompatibilityMatrix(Base):
    __tablename__ = "compatibility_matrix"
    id = Column(Integer, primary_key=True, index=True)
    component_type = Column(String)
    hardware_revision = Column(String)
    firmware_version = Column(String)
    is_compatible = Column(Boolean)


class TestScenario(Base):
    __tablename__ = "test_scenarios"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String)
    limits = Column(JSON)
    description = Column(Text, nullable=True)

    test_runs = relationship("TestRun", back_populates="scenario")


class TestRun(Base):
    __tablename__ = "test_runs"
    id = Column(Integer, primary_key=True, index=True)
    product_id = Column(Integer, ForeignKey("products.id"))
    scenario_id = Column(Integer, ForeignKey("test_scenarios.id"))
    status = Column(String, default="running")  # running, passed, failed
    started_at = Column(DateTime, default=datetime.utcnow)
    finished_at = Column(DateTime, nullable=True)
    report_path = Column(String, nullable=True)

    product = relationship("Product", back_populates="test_runs")
    scenario = relationship("TestScenario", back_populates="test_runs")
    telemetry = relationship("TelemetryPoint", back_populates="test_run")


class TelemetryPoint(Base):
    __tablename__ = "telemetry_points"
    id = Column(Integer, primary_key=True, index=True)
    test_run_id = Column(Integer, ForeignKey("test_runs.id"))
    parameter = Column(String)
    value = Column(Float)
    within_limit = Column(Boolean)
    recorded_at = Column(DateTime, default=datetime.utcnow)

    test_run = relationship("TestRun", back_populates="telemetry")


class DocTemplate(Base):
    __tablename__ = "doc_templates"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String)
    template_type = Column(String)  # mass_report, firmware_passport, test_act
    description = Column(Text, nullable=True)
    file_path = Column(String)
    uploaded_at = Column(DateTime, default=datetime.utcnow)
    is_active = Column(Boolean, default=True)


Base.metadata.create_all(bind=engine)