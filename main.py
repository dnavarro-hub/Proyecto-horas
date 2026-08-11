from datetime import date, timedelta
from typing import List, Optional
from fastapi import FastAPI, HTTPException, Depends, UploadFile, File
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, validator
from sqlalchemy import create_engine, Column, Integer, String, Date, DateTime, Float, Boolean
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import io
import os
from datetime import datetime
import hashlib
import pandas as pd

# ==================== BASE DE DATOS ====================
DATABASE_URL = os.environ.get("DATABASE_URL", "postgresql://localhost/registros")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

engine = create_engine(DATABASE_URL, pool_pre_ping=True)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# ==================== UTILIDADES ====================

def hashear_pin(pin: str) -> str:
    return hashlib.sha256(pin.encode()).hexdigest()

# ==================== MODELOS ====================

class RegistroDB(Base):
    __tablename__ = "registros"
    id              = Column(Integer, primary_key=True, index=True)
    usuario         = Column(String, index=True)
    fecha           = Column(Date)
    tarea_principal = Column(String, index=True)
    subtarea        = Column(String)
    tiempo_minutos  = Column(Integer)
    proyecto        = Column(String, nullable=True)
    comentarios     = Column(String, nullable=True)
    created_at      = Column(DateTime, default=datetime.utcnow)
    updated_at      = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

class UsuarioDB(Base):
    __tablename__ = "usuarios"
    id      = Column(Integer, primary_key=True, index=True)
    codigo  = Column(String, unique=True, index=True)
    nombre  = Column(String)
    pin     = Column(String)
    rol     = Column(String, default="operario")

class PlantillaDB(Base):
    __tablename__ = "plantillas"
    id         = Column(Integer, primary_key=True, index=True)
    nombre     = Column(String)
    usuario    = Column(String)
    created_at = Column(DateTime, default=datetime.utcnow)

class PlantillaItemDB(Base):
    __tablename__ = "plantilla_items"
    id              = Column(Integer, primary_key=True, index=True)
    plantilla_id    = Column(Integer)
    tarea_principal = Column(String)
    subtarea        = Column(String)
    tiempo_minutos  = Column(Integer)
    proyecto        = Column(String, nullable=True)
    comentarios     = Column(String, nullable=True)

class VolumenDB(Base):
    __tablename__ = "volumenes"
    id              = Column(Integer, primary_key=True, index=True)
    fecha           = Column(Date, index=True)
    tarea_principal = Column(String, index=True)
    unidades        = Column(Integer)
    horas_teoricas  = Column(Float)
    creado_por      = Column(String)
    comentarios     = Column(String, nullable=True)
    created_at      = Column(DateTime, default=datetime.utcnow)
    updated_at      = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

class ObjetivoDB(Base):
    __tablename__ = "objetivos"
    id              = Column(Integer, primary_key=True, index=True)
    tarea_principal = Column(String, index=True)
    subtarea        = Column(String, nullable=True)
    uds_hora        = Column(Float)
    horas_jornada   = Column(Float, default=8.0)
    updated_at      = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

class ConfiguracionDB(Base):
    __tablename__ = "configuracion"
    id             = Column(Integer, primary_key=True, index=True)
    clave          = Column(String, unique=True, index=True)
    valor          = Column(String)
    descripcion    = Column(String, nullable=True)
    updated_at     = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

class TareaDB(Base):
    __tablename__ = "tareas"
    id              = Column(Integer, primary_key=True, index=True)
    nombre          = Column(String, unique=True, index=True)
    color           = Column(String, default="#94a3b8")
    activa          = Column(Boolean, default=True)
    created_at      = Column(DateTime, default=datetime.utcnow)

class SubtareaDB(Base):
    __tablename__ = "subtareas"
    id              = Column(Integer, primary_key=True, index=True)
    tarea_nombre    = Column(String, index=True)
    nombre          = Column(String)
    activa          = Column(Boolean, default=True)
    created_at      = Column(DateTime, default=datetime.utcnow)

# NUEVO: Modelo para productividad por Excel
class ProductividadExcelDB(Base):
    __tablename__ = "productividad_excel"
    id              = Column(Integer, primary_key=True, index=True)
    fecha           = Column(Date, index=True)
    usuario         = Column(String, index=True)
    tarea_principal = Column(String, index=True)
    subtarea        = Column(String, index=True)
    unidades        = Column(Integer)
    subido_por      = Column(String)
    created_at      = Column(DateTime, default=datetime.utcnow)

Base.metadata.create_all(bind=engine)

# ==================== INICIALIZACIÓN DE DATOS ====================

TAREAS_DEFAULT = [
    {"nombre": "Picking",   "color": "#3b82f6"},
    {"nombre": "Packing",   "color": "#8b5cf6"},
    {"nombre": "Inbound",   "color": "#f59e0b"},
    {"nombre": "Shipping",  "color": "#10b981"},
    {"nombre": "Ecommerce", "color": "#ef4444"},
]

SUBTAREAS_DEFAULT = {
    "Picking":   ["Picking Balda","Picking Palletl","Picking Percha","Revisión de Picking",
                  "Picking Kardex","Picking Recogepedidos","Picking Obsoleto","Inventario",
                  "Reposiciones","Traspasos","Recepción Kardex","Formacion","Compactar",
                  "Incidencias","Varios IT","Varios Trigo","Lanzar pedidos","Reuniones","Otros"],
    "Packing":   ["Mecado Piscina","Mercado Contenedor","Tienda Etiquetado","Wholesale",
                  "Tienda RFID","Tienda Empleado y Otros","Materiales","Runner","Formacion",
                  "incidencias","Reuniones","Otros","Admin"],
    "Inbound":   ["Muelle","Rec Pallet","Rec balda","Rec Percha","Rec Zapatos","Rec Trigo",
                  "Devoluciones","Compactar","Recepción Kardex","Materiales","Formacion",
                  "incidencias","Reuniones","Otros","Admin"],
    "Shipping":  ["COURIER preparacion carga","Courier Carga","FW preparación carga",
                  "FW carga","SERWELL carga","Devoluciones","Formacion","incidencias",
                  "Reuniones","Inventario","Otros"],
    "Ecommerce": ["Empaquetado","Runner","Store RQ Tienda","Calidad Devo","Calidad OneStock",
                  "Calidad Gestion","Calidad Otro","Formacion","Actividad","Limpieza",
                  "Otros","Reuniones"],
}

with SessionLocal() as session:
    # Supervisor por defecto
    if not session.query(UsuarioDB).filter(UsuarioDB.rol == "supervisor").first():
        session.add(UsuarioDB(
            codigo="ADMIN",
            nombre="Administrador",
            pin=hashear_pin("1234"),
            rol="supervisor"
        ))
        session.commit()

    # Objetivos por defecto
    for t in TAREAS_DEFAULT:
        if not session.query(ObjetivoDB).filter(
            ObjetivoDB.tarea_principal == t["nombre"],
            ObjetivoDB.subtarea == None
        ).first():
            session.add(ObjetivoDB(
                tarea_principal=t["nombre"],
                subtarea=None,
                uds_hora=100.0,
                horas_jornada=8.0
            ))

    # Tareas por defecto
    for t in TAREAS_DEFAULT:
        if not session.query(TareaDB).filter(TareaDB.nombre == t["nombre"]).first():
            session.add(TareaDB(nombre=t["nombre"], color=t["color"], activa=True))

    # Subtareas por defecto
    for tarea_nombre, subtareas in SUBTAREAS_DEFAULT.items():
        for sub in subtareas:
            if not session.query(SubtareaDB).filter(
                SubtareaDB.tarea_nombre == tarea_nombre,
                SubtareaDB.nombre == sub
            ).first():
                session.add(SubtareaDB(tarea_nombre=tarea_nombre, nombre=sub, activa=True))

    # Configuración por defecto
    if not session.query(ConfiguracionDB).filter(
        ConfiguracionDB.clave == "minutos_jornada"
    ).first():
        session.add(ConfiguracionDB(
            clave="minutos_jornada",
            valor="480",
            descripcion="Minutos máximos por jornada laboral"
        ))

    session.commit()
# ==================== APP ====================

app = FastAPI(title="API Registro de Tareas", version="11.0.0")

ALLOWED_ORIGINS = [
    "https://proyecto-horas.onrender.com",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE"],
    allow_headers=["Content-Type", "Authorization"],
)

# ==================== SCHEMAS ====================

class RegistroTarea(BaseModel):
    usuario: str
    fecha: date
    tarea_principal: str
    subtarea: str
    tiempo_minutos: int
    proyecto: Optional[str] = None
    comentarios: Optional[str] = None

    @validator("tiempo_minutos")
    def validar_tiempo(cls, v):
        if v <= 0:
            raise ValueError("El tiempo debe ser mayor que 0")
        if v > 720:
            raise ValueError("El tiempo no puede superar 720 minutos")
        return v

    @validator("fecha")
    def validar_fecha(cls, v):
        if v > date.today():
            raise ValueError("La fecha no puede ser futura")
        return v

class RegistroTareaRespuesta(RegistroTarea):
    id: int
    created_at: Optional[datetime] = None
    updated_at: Optional[datetime] = None
    class Config:
        from_attributes = True

class UsuarioCreate(BaseModel):
    codigo: str
    nombre: str
    pin: str
    rol: str = "operario"

    @validator("pin")
    def validar_pin(cls, v):
        if len(v) < 4:
            raise ValueError("El PIN debe tener al menos 4 caracteres")
        return v

    @validator("rol")
    def validar_rol(cls, v):
        if v not in ["operario", "supervisor"]:
            raise ValueError("Rol no válido")
        return v

class UsuarioRespuesta(BaseModel):
    id: int
    codigo: str
    nombre: str
    rol: str
    class Config:
        from_attributes = True

class LoginRequest(BaseModel):
    codigo: str
    pin: str

class PlantillaItemCreate(BaseModel):
    tarea_principal: str
    subtarea: str
    tiempo_minutos: int
    proyecto: Optional[str] = None
    comentarios: Optional[str] = None

class PlantillaCreate(BaseModel):
    nombre: str
    usuario: str
    items: List[PlantillaItemCreate]

class VolumenCreate(BaseModel):
    fecha: date
    tarea_principal: str
    unidades: int
    horas_teoricas: float
    creado_por: str
    comentarios: Optional[str] = None

    @validator("unidades")
    def validar_unidades(cls, v):
        if v <= 0:
            raise ValueError("Las unidades deben ser mayores que 0")
        return v

    @validator("horas_teoricas")
    def validar_horas(cls, v):
        if v <= 0:
            raise ValueError("Las horas teóricas deben ser mayores que 0")
        return v

class VolumenRespuesta(VolumenCreate):
    id: int
    created_at: Optional[datetime] = None
    updated_at: Optional[datetime] = None
    class Config:
        from_attributes = True

class ObjetivoCreate(BaseModel):
    tarea_principal: str
    subtarea: Optional[str] = None
    uds_hora: float
    horas_jornada: float = 8.0

    @validator("uds_hora")
    def validar_uds_hora(cls, v):
        if v <= 0:
            raise ValueError("Las unidades por hora deben ser mayores que 0")
        return v

class ObjetivoRespuesta(ObjetivoCreate):
    id: int
    updated_at: Optional[datetime] = None
    class Config:
        from_attributes = True

class TareaCreate(BaseModel):
    nombre: str
    color: str = "#94a3b8"
    activa: bool = True

class TareaRespuesta(TareaCreate):
    id: int
    class Config:
        from_attributes = True

class SubtareaCreate(BaseModel):
    tarea_nombre: str
    nombre: str
    activa: bool = True

class SubtareaRespuesta(SubtareaCreate):
    id: int
    class Config:
        from_attributes = True

class ConfiguracionUpdate(BaseModel):
    valor: str

class ConfiguracionRespuesta(BaseModel):
    clave: str
    valor: str
    descripcion: Optional[str] = None
    class Config:
        from_attributes = True

# NUEVO: Schema productividad Excel
class ProductividadExcelRespuesta(BaseModel):
    id: int
    fecha: date
    usuario: str
    tarea_principal: str
    subtarea: str
    unidades: int
    subido_por: str
    created_at: Optional[datetime] = None
    class Config:
        from_attributes = True

class ProductividadCalculada(BaseModel):
    usuario: str
    fecha: date
    tarea_principal: str
    subtarea: str
    unidades: int
    minutos_reales: int
    horas_reales: float
    uds_hora_real: Optional[float]
    uds_hora_objetivo: Optional[float]
    eficiencia_pct: Optional[float]

# ==================== DEPENDENCIAS ====================

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def get_minutos_jornada(db: Session) -> int:
    config = db.query(ConfiguracionDB).filter(
        ConfiguracionDB.clave == "minutos_jornada"
    ).first()
    return int(config.valor) if config else 480

def validar_rango_fechas(desde: date, hasta: date):
    if hasta < desde:
        raise HTTPException(
            status_code=400,
            detail="La fecha 'hasta' no puede ser anterior a 'desde'"
        )

# ==================== RUTAS GENERALES ====================

@app.get("/")
def leer_index():
    ruta = os.path.join(os.path.dirname(__file__), "index.html")
    if os.path.exists(ruta):
        return FileResponse(ruta)
    return {"error": "No se encuentra index.html"}

# ==================== LOGIN ====================

@app.post("/login/")
def login(datos: LoginRequest, db: Session = Depends(get_db)):
    pin_hasheado = hashear_pin(datos.pin)
    usuario = db.query(UsuarioDB).filter(
        UsuarioDB.codigo == datos.codigo,
        UsuarioDB.pin == pin_hasheado
    ).first()
    if not usuario:
        raise HTTPException(status_code=401, detail="Código o PIN incorrectos")
    return {
        "id": usuario.id,
        "codigo": usuario.codigo,
        "nombre": usuario.nombre,
        "rol": usuario.rol
    }

# ==================== USUARIOS ====================

@app.post("/usuarios/", response_model=UsuarioRespuesta, status_code=201)
def crear_usuario(usuario: UsuarioCreate, db: Session = Depends(get_db)):
    if db.query(UsuarioDB).filter(UsuarioDB.codigo == usuario.codigo).first():
        raise HTTPException(status_code=400, detail="Ya existe un usuario con ese código")
    datos = usuario.dict()
    datos["pin"] = hashear_pin(datos["pin"])
    nuevo = UsuarioDB(**datos)
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    return nuevo

@app.get("/usuarios/", response_model=List[UsuarioRespuesta])
def obtener_usuarios(db: Session = Depends(get_db)):
    return db.query(UsuarioDB).all()

@app.delete("/usuarios/{id}", status_code=204)
def eliminar_usuario(id: int, db: Session = Depends(get_db)):
    usuario = db.query(UsuarioDB).filter(UsuarioDB.id == id).first()
    if not usuario:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    db.delete(usuario)
    db.commit()

# ==================== TAREAS ====================

@app.get("/tareas/", response_model=List[TareaRespuesta])
def obtener_tareas(solo_activas: bool = True, db: Session = Depends(get_db)):
    query = db.query(TareaDB)
    if solo_activas:
        query = query.filter(TareaDB.activa == True)
    return query.order_by(TareaDB.nombre).all()

@app.post("/tareas/", response_model=TareaRespuesta, status_code=201)
def crear_tarea(tarea: TareaCreate, db: Session = Depends(get_db)):
    if db.query(TareaDB).filter(TareaDB.nombre == tarea.nombre).first():
        raise HTTPException(status_code=400, detail="Ya existe una tarea con ese nombre")
    nueva = TareaDB(**tarea.dict())
    db.add(nueva)
    db.commit()
    db.refresh(nueva)
    return nueva

@app.put("/tareas/{id}", response_model=TareaRespuesta)
def editar_tarea(id: int, tarea: TareaCreate, db: Session = Depends(get_db)):
    db_tarea = db.query(TareaDB).filter(TareaDB.id == id).first()
    if not db_tarea:
        raise HTTPException(status_code=404, detail="Tarea no encontrada")
    for k, v in tarea.dict().items():
        setattr(db_tarea, k, v)
    db.commit()
    db.refresh(db_tarea)
    return db_tarea

@app.delete("/tareas/{id}", status_code=204)
def eliminar_tarea(id: int, db: Session = Depends(get_db)):
    db_tarea = db.query(TareaDB).filter(TareaDB.id == id).first()
    if not db_tarea:
        raise HTTPException(status_code=404, detail="Tarea no encontrada")
    db.delete(db_tarea)
    db.commit()

# ==================== SUBTAREAS ====================

@app.get("/subtareas/", response_model=List[SubtareaRespuesta])
def obtener_subtareas(
    tarea_nombre: Optional[str] = None,
    solo_activas: bool = True,
    db: Session = Depends(get_db)
):
    query = db.query(SubtareaDB)
    if tarea_nombre:
        query = query.filter(SubtareaDB.tarea_nombre == tarea_nombre)
    if solo_activas:
        query = query.filter(SubtareaDB.activa == True)
    return query.order_by(SubtareaDB.tarea_nombre, SubtareaDB.nombre).all()

@app.post("/subtareas/", response_model=SubtareaRespuesta, status_code=201)
def crear_subtarea(subtarea: SubtareaCreate, db: Session = Depends(get_db)):
    if db.query(SubtareaDB).filter(
        SubtareaDB.tarea_nombre == subtarea.tarea_nombre,
        SubtareaDB.nombre == subtarea.nombre
    ).first():
        raise HTTPException(status_code=400, detail="Ya existe esa subtarea para esta tarea")
    nueva = SubtareaDB(**subtarea.dict())
    db.add(nueva)
    db.commit()
    db.refresh(nueva)
    return nueva

@app.put("/subtareas/{id}", response_model=SubtareaRespuesta)
def editar_subtarea(id: int, subtarea: SubtareaCreate, db: Session = Depends(get_db)):
    db_sub = db.query(SubtareaDB).filter(SubtareaDB.id == id).first()
    if not db_sub:
        raise HTTPException(status_code=404, detail="Subtarea no encontrada")
    for k, v in subtarea.dict().items():
        setattr(db_sub, k, v)
    db.commit()
    db.refresh(db_sub)
    return db_sub

@app.delete("/subtareas/{id}", status_code=204)
def eliminar_subtarea(id: int, db: Session = Depends(get_db)):
    db_sub = db.query(SubtareaDB).filter(SubtareaDB.id == id).first()
    if not db_sub:
        raise HTTPException(status_code=404, detail="Subtarea no encontrada")
    db.delete(db_sub)
    db.commit()

# ==================== CONFIGURACIÓN ====================

@app.get("/configuracion/", response_model=List[ConfiguracionRespuesta])
def obtener_configuracion(db: Session = Depends(get_db)):
    return db.query(ConfiguracionDB).all()

@app.put("/configuracion/{clave}", response_model=ConfiguracionRespuesta)
def actualizar_configuracion(
    clave: str,
    datos: ConfiguracionUpdate,
    db: Session = Depends(get_db)
):
    config = db.query(ConfiguracionDB).filter(ConfiguracionDB.clave == clave).first()
    if not config:
        raise HTTPException(status_code=404, detail="Configuración no encontrada")
    config.valor = datos.valor
    config.updated_at = datetime.utcnow()
    db.commit()
    db.refresh(config)
    return config
# ==================== REGISTROS ====================

@app.post("/registros/", response_model=RegistroTareaRespuesta, status_code=201)
def crear_registro(registro: RegistroTarea, db: Session = Depends(get_db)):
    minutos_jornada = get_minutos_jornada(db)
    total = db.query(RegistroDB).filter(
        RegistroDB.usuario == registro.usuario,
        RegistroDB.fecha == registro.fecha
    ).with_entities(RegistroDB.tiempo_minutos).all()
    total_mins = sum(r.tiempo_minutos for r in total)
    if total_mins + registro.tiempo_minutos > minutos_jornada:
        raise HTTPException(
            status_code=400,
            detail=f"No se pueden superar {minutos_jornada} minutos. "
                   f"Minutos ya registrados: {total_mins}"
        )
    nuevo = RegistroDB(**registro.dict())
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    return nuevo

@app.get("/registros/", response_model=List[RegistroTareaRespuesta])
def obtener_registros(
    skip: int = 0,
    limit: int = 200,
    db: Session = Depends(get_db)
):
    return (
        db.query(RegistroDB)
        .order_by(RegistroDB.fecha.desc())
        .offset(skip)
        .limit(limit)
        .all()
    )

@app.get("/registros/rango/", response_model=List[RegistroTareaRespuesta])
def obtener_registros_rango(
    desde: date,
    hasta: date,
    usuario: Optional[str] = None,
    tarea: Optional[str] = None,
    db: Session = Depends(get_db)
):
    validar_rango_fechas(desde, hasta)
    query = db.query(RegistroDB).filter(
        RegistroDB.fecha >= desde,
        RegistroDB.fecha <= hasta
    )
    if usuario:
        query = query.filter(RegistroDB.usuario == usuario)
    if tarea:
        query = query.filter(RegistroDB.tarea_principal == tarea)
    return query.order_by(RegistroDB.fecha.desc()).all()

@app.get("/registros/resumen-semanal/")
def resumen_semanal(
    usuario: Optional[str] = None,
    fecha_ref: Optional[date] = None,
    db: Session = Depends(get_db)
):
    if not fecha_ref:
        fecha_ref = date.today()
    inicio = fecha_ref - timedelta(days=fecha_ref.weekday())
    fin = inicio + timedelta(days=6)
    query = db.query(RegistroDB).filter(
        RegistroDB.fecha >= inicio,
        RegistroDB.fecha <= fin
    )
    if usuario:
        query = query.filter(RegistroDB.usuario == usuario)
    registros = query.all()
    resumen = {}
    for r in registros:
        dia = str(r.fecha)
        if dia not in resumen:
            resumen[dia] = {"total_minutos": 0, "tareas": {}}
        resumen[dia]["total_minutos"] += r.tiempo_minutos
        resumen[dia]["tareas"][r.tarea_principal] = (
            resumen[dia]["tareas"].get(r.tarea_principal, 0) + r.tiempo_minutos
        )
    return {
        "inicio": str(inicio),
        "fin": str(fin),
        "dias": resumen,
        "total_minutos": sum(r.tiempo_minutos for r in registros)
    }

@app.get("/registros/resumen-mensual/")
def resumen_mensual(
    usuario: Optional[str] = None,
    anyo: Optional[int] = None,
    mes: Optional[int] = None,
    db: Session = Depends(get_db)
):
    hoy = date.today()
    anyo = anyo or hoy.year
    mes = mes or hoy.month
    inicio = date(anyo, mes, 1)
    fin = (
        date(anyo + 1, 1, 1) - timedelta(days=1)
        if mes == 12
        else date(anyo, mes + 1, 1) - timedelta(days=1)
    )
    query = db.query(RegistroDB).filter(
        RegistroDB.fecha >= inicio,
        RegistroDB.fecha <= fin
    )
    if usuario:
        query = query.filter(RegistroDB.usuario == usuario)
    registros = query.all()
    por_usuario, por_tarea = {}, {}
    for r in registros:
        por_usuario[r.usuario] = por_usuario.get(r.usuario, 0) + r.tiempo_minutos
        por_tarea[r.tarea_principal] = por_tarea.get(r.tarea_principal, 0) + r.tiempo_minutos
    return {
        "anyo": anyo,
        "mes": mes,
        "total_minutos": sum(r.tiempo_minutos for r in registros),
        "total_registros": len(registros),
        "por_usuario": por_usuario,
        "por_tarea": por_tarea
    }

@app.get("/registros/{usuario}", response_model=List[RegistroTareaRespuesta])
def obtener_registros_por_usuario(
    usuario: str,
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    db: Session = Depends(get_db)
):
    if desde and hasta:
        validar_rango_fechas(desde, hasta)
    query = db.query(RegistroDB).filter(RegistroDB.usuario.ilike(usuario))
    if desde:
        query = query.filter(RegistroDB.fecha >= desde)
    if hasta:
        query = query.filter(RegistroDB.fecha <= hasta)
    resultado = query.order_by(RegistroDB.fecha.desc()).all()
    if not resultado:
        raise HTTPException(status_code=404, detail="No se encontraron registros")
    return resultado

@app.put("/registros/{id}", response_model=RegistroTareaRespuesta)
def editar_registro(id: int, registro: RegistroTarea, db: Session = Depends(get_db)):
    minutos_jornada = get_minutos_jornada(db)
    db_reg = db.query(RegistroDB).filter(RegistroDB.id == id).first()
    if not db_reg:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    total = db.query(RegistroDB).filter(
        RegistroDB.usuario == registro.usuario,
        RegistroDB.fecha == registro.fecha,
        RegistroDB.id != id
    ).with_entities(RegistroDB.tiempo_minutos).all()
    total_mins = sum(r.tiempo_minutos for r in total)
    if total_mins + registro.tiempo_minutos > minutos_jornada:
        raise HTTPException(
            status_code=400,
            detail=f"No se pueden superar {minutos_jornada} minutos. "
                   f"Minutos ya registrados: {total_mins}"
        )
    for k, v in registro.dict().items():
        setattr(db_reg, k, v)
    db_reg.updated_at = datetime.utcnow()
    db.commit()
    db.refresh(db_reg)
    return db_reg

@app.delete("/registros/{id}", status_code=204)
def eliminar_registro(id: int, db: Session = Depends(get_db)):
    db_reg = db.query(RegistroDB).filter(RegistroDB.id == id).first()
    if not db_reg:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    db.delete(db_reg)
    db.commit()

@app.post("/registros/{id}/duplicar", response_model=RegistroTareaRespuesta, status_code=201)
def duplicar_registro(
    id: int,
    nueva_fecha: Optional[date] = None,
    db: Session = Depends(get_db)
):
    minutos_jornada = get_minutos_jornada(db)
    original = db.query(RegistroDB).filter(RegistroDB.id == id).first()
    if not original:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    fecha_destino = nueva_fecha or original.fecha
    total = db.query(RegistroDB).filter(
        RegistroDB.usuario == original.usuario,
        RegistroDB.fecha == fecha_destino
    ).with_entities(RegistroDB.tiempo_minutos).all()
    total_mins = sum(r.tiempo_minutos for r in total)
    if total_mins + original.tiempo_minutos > minutos_jornada:
        raise HTTPException(
            status_code=400,
            detail=f"No se pueden superar {minutos_jornada} minutos. "
                   f"Minutos ya registrados: {total_mins}"
        )
    nuevo = RegistroDB(
        usuario=original.usuario,
        fecha=fecha_destino,
        tarea_principal=original.tarea_principal,
        subtarea=original.subtarea,
        tiempo_minutos=original.tiempo_minutos,
        proyecto=original.proyecto,
        comentarios=original.comentarios
    )
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    return nuevo

# ==================== PLANTILLAS ====================

@app.post("/plantillas/", status_code=201)
def crear_plantilla(plantilla: PlantillaCreate, db: Session = Depends(get_db)):
    nueva = PlantillaDB(nombre=plantilla.nombre, usuario=plantilla.usuario)
    db.add(nueva)
    db.commit()
    db.refresh(nueva)
    for item in plantilla.items:
        db.add(PlantillaItemDB(
            plantilla_id=nueva.id,
            tarea_principal=item.tarea_principal,
            subtarea=item.subtarea,
            tiempo_minutos=item.tiempo_minutos,
            proyecto=item.proyecto,
            comentarios=item.comentarios
        ))
    db.commit()
    return {"id": nueva.id, "nombre": nueva.nombre, "mensaje": "Plantilla creada correctamente"}

@app.get("/plantillas/")
def obtener_plantillas(usuario: Optional[str] = None, db: Session = Depends(get_db)):
    query = db.query(PlantillaDB)
    if usuario:
        query = query.filter(PlantillaDB.usuario == usuario)
    resultado = []
    for p in query.all():
        items = db.query(PlantillaItemDB).filter(PlantillaItemDB.plantilla_id == p.id).all()
        resultado.append({
            "id": p.id,
            "nombre": p.nombre,
            "usuario": p.usuario,
            "created_at": str(p.created_at),
            "items": [
                {
                    "id": i.id,
                    "tarea_principal": i.tarea_principal,
                    "subtarea": i.subtarea,
                    "tiempo_minutos": i.tiempo_minutos,
                    "proyecto": i.proyecto,
                    "comentarios": i.comentarios
                }
                for i in items
            ]
        })
    return resultado

@app.delete("/plantillas/{id}", status_code=204)
def eliminar_plantilla(id: int, db: Session = Depends(get_db)):
    plantilla = db.query(PlantillaDB).filter(PlantillaDB.id == id).first()
    if not plantilla:
        raise HTTPException(status_code=404, detail="Plantilla no encontrada")
    db.query(PlantillaItemDB).filter(PlantillaItemDB.plantilla_id == id).delete()
    db.delete(plantilla)
    db.commit()

# ==================== OBJETIVOS ====================

@app.get("/objetivos/", response_model=List[ObjetivoRespuesta])
def obtener_objetivos(db: Session = Depends(get_db)):
    return db.query(ObjetivoDB).all()

@app.put("/objetivos/{tarea}", response_model=ObjetivoRespuesta)
def actualizar_objetivo(
    tarea: str,
    objetivo: ObjetivoCreate,
    db: Session = Depends(get_db)
):
    db_obj = db.query(ObjetivoDB).filter(
        ObjetivoDB.tarea_principal == tarea,
        ObjetivoDB.subtarea == objetivo.subtarea
    ).first()
    if not db_obj:
        db_obj = ObjetivoDB(
            tarea_principal=tarea,
            subtarea=objetivo.subtarea,
            uds_hora=objetivo.uds_hora,
            horas_jornada=objetivo.horas_jornada
        )
        db.add(db_obj)
    else:
        db_obj.uds_hora = objetivo.uds_hora
        db_obj.horas_jornada = objetivo.horas_jornada
        db_obj.updated_at = datetime.utcnow()
    db.commit()
    db.refresh(db_obj)
    return db_obj

# ==================== VOLÚMENES ====================

@app.post("/volumenes/", response_model=VolumenRespuesta, status_code=201)
def crear_volumen(volumen: VolumenCreate, db: Session = Depends(get_db)):
    nuevo = VolumenDB(**volumen.dict())
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    return nuevo

@app.get("/volumenes/", response_model=List[VolumenRespuesta])
def obtener_volumenes(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    tarea: Optional[str] = None,
    db: Session = Depends(get_db)
):
    if desde and hasta:
        validar_rango_fechas(desde, hasta)
    query = db.query(VolumenDB)
    if desde:
        query = query.filter(VolumenDB.fecha >= desde)
    if hasta:
        query = query.filter(VolumenDB.fecha <= hasta)
    if tarea:
        query = query.filter(VolumenDB.tarea_principal == tarea)
    return query.order_by(VolumenDB.fecha.desc()).all()

@app.put("/volumenes/{id}", response_model=VolumenRespuesta)
def editar_volumen(id: int, volumen: VolumenCreate, db: Session = Depends(get_db)):
    db_vol = db.query(VolumenDB).filter(VolumenDB.id == id).first()
    if not db_vol:
        raise HTTPException(status_code=404, detail="Volumen no encontrado")
    for k, v in volumen.dict().items():
        setattr(db_vol, k, v)
    db_vol.updated_at = datetime.utcnow()
    db.commit()
    db.refresh(db_vol)
    return db_vol

@app.delete("/volumenes/{id}", status_code=204)
def eliminar_volumen(id: int, db: Session = Depends(get_db)):
    db_vol = db.query(VolumenDB).filter(VolumenDB.id == id).first()
    if not db_vol:
        raise HTTPException(status_code=404, detail="Volumen no encontrado")
    db.delete(db_vol)
    db.commit()

@app.get("/volumenes/metricas/")
def obtener_metricas(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    db: Session = Depends(get_db)
):
    if not desde:
        desde = date.today() - timedelta(days=30)
    if not hasta:
        hasta = date.today()
    validar_rango_fechas(desde, hasta)
    volumenes = db.query(VolumenDB).filter(
        VolumenDB.fecha >= desde,
        VolumenDB.fecha <= hasta
    ).all()
    registros = db.query(RegistroDB).filter(
        RegistroDB.fecha >= desde,
        RegistroDB.fecha <= hasta
    ).all()
    resultado = []
    for v in volumenes:
        horas_reales = sum(
            r.tiempo_minutos for r in registros
            if str(r.fecha) == str(v.fecha) and r.tarea_principal == v.tarea_principal
        ) / 60
        eficiencia = round((v.horas_teoricas / horas_reales) * 100, 1) if horas_reales > 0 else None
        uds_hora_real = round(v.unidades / horas_reales, 1) if horas_reales > 0 else None
        uds_hora_teorica = round(v.unidades / v.horas_teoricas, 1) if v.horas_teoricas > 0 else None
        resultado.append({
            "id": v.id,
            "fecha": str(v.fecha),
            "tarea_principal": v.tarea_principal,
            "unidades": v.unidades,
            "horas_teoricas": v.horas_teoricas,
            "horas_reales": round(horas_reales, 2),
            "eficiencia_pct": eficiencia,
            "uds_hora_real": uds_hora_real,
            "uds_hora_teorica": uds_hora_teorica,
            "desviacion_horas": round(horas_reales - v.horas_teoricas, 2),
            "comentarios": v.comentarios,
            "creado_por": v.creado_por
        })
    return sorted(resultado, key=lambda x: x["fecha"], reverse=True)

# ==================== PRODUCTIVIDAD EXCEL ====================

@app.post("/productividad/subir-excel/")
async def subir_excel_productividad(
    file: UploadFile = File(...),
    subido_por: str = "ADMIN",
    db: Session = Depends(get_db)
):
    """
    Sube un Excel con columnas obligatorias:
    fecha | operario | tarea | subtarea | unidades
    Calcula productividad cruzando con horas registradas.
    """
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="El fichero debe ser .xlsx o .xls")

    contenido = await file.read()
    try:
        df = pd.read_excel(io.BytesIO(contenido))
    except Exception:
        raise HTTPException(status_code=400, detail="No se pudo leer el fichero Excel")

    # Normalizar nombres de columnas
    df.columns = [c.strip().lower() for c in df.columns]

    columnas_requeridas = {"fecha", "operario", "tarea", "subtarea", "unidades"}
    if not columnas_requeridas.issubset(set(df.columns)):
        raise HTTPException(
            status_code=400,
            detail=f"El Excel debe tener las columnas: {', '.join(columnas_requeridas)}"
        )

    # Limpiar datos
    df = df.dropna(subset=["fecha", "operario", "tarea", "subtarea", "unidades"])
    df["fecha"] = pd.to_datetime(df["fecha"]).dt.date
    df["unidades"] = pd.to_numeric(df["unidades"], errors="coerce").fillna(0).astype(int)
    df = df[df["unidades"] > 0]

    insertados = 0
    errores = []

    for _, fila in df.iterrows():
        try:
            # Verificar que el usuario existe
            usuario_db = db.query(UsuarioDB).filter(
                UsuarioDB.codigo == str(fila["operario"]).strip()
            ).first()
            if not usuario_db:
                errores.append(f"Operario no encontrado: {fila['operario']}")
                continue

            nuevo = ProductividadExcelDB(
                fecha=fila["fecha"],
                usuario=str(fila["operario"]).strip(),
                tarea_principal=str(fila["tarea"]).strip(),
                subtarea=str(fila["subtarea"]).strip(),
                unidades=int(fila["unidades"]),
                subido_por=subido_por
            )
            db.add(nuevo)
            insertados += 1
        except Exception as e:
            errores.append(f"Fila con error: {str(e)}")

    db.commit()

    return {
        "insertados": insertados,
        "errores": errores,
        "total_filas": len(df),
        "mensaje": f"{insertados} registros importados correctamente"
    }


@app.get("/productividad/calcular/")
def calcular_productividad(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    usuario: Optional[str] = None,
    tarea: Optional[str] = None,
    subtarea: Optional[str] = None,
    ordenar_por: str = "uds_hora_real",
    orden: str = "desc",
    db: Session = Depends(get_db)
):
    """
    Cruza los datos del Excel subido con los registros de horas
    y calcula productividad por operario/subtarea.
    Permite ordenar por cualquier campo: uds_hora_real, eficiencia_pct, unidades, horas_reales.
    orden: 'asc' o 'desc'
    """
    if not desde:
        desde = date.today() - timedelta(days=30)
    if not hasta:
        hasta = date.today()
    validar_rango_fechas(desde, hasta)

    query_excel = db.query(ProductividadExcelDB).filter(
        ProductividadExcelDB.fecha >= desde,
        ProductividadExcelDB.fecha <= hasta
    )
    if usuario:
        query_excel = query_excel.filter(ProductividadExcelDB.usuario == usuario)
    if tarea:
        query_excel = query_excel.filter(ProductividadExcelDB.tarea_principal == tarea)
    if subtarea:
        query_excel = query_excel.filter(ProductividadExcelDB.subtarea == subtarea)

    registros_excel = query_excel.all()

    if not registros_excel:
        return []

    registros_horas = db.query(RegistroDB).filter(
        RegistroDB.fecha >= desde,
        RegistroDB.fecha <= hasta
    ).all()

    objetivos = {
        (o.tarea_principal, o.subtarea): o
        for o in db.query(ObjetivoDB).all()
    }

    resultado = []
    for ex in registros_excel:
        # Minutos reales del operario en esa subtarea y fecha
        minutos_reales = sum(
            r.tiempo_minutos for r in registros_horas
            if (r.usuario == ex.usuario
                and str(r.fecha) == str(ex.fecha)
                and r.tarea_principal == ex.tarea_principal
                and r.subtarea == ex.subtarea)
        )

        horas_reales = round(minutos_reales / 60, 2) if minutos_reales > 0 else 0
        uds_hora_real = round(ex.unidades / horas_reales, 1) if horas_reales > 0 else None

        # Buscar objetivo: primero por tarea+subtarea, luego solo por tarea
        obj = (
            objetivos.get((ex.tarea_principal, ex.subtarea))
            or objetivos.get((ex.tarea_principal, None))
        )
        uds_hora_objetivo = obj.uds_hora if obj else None
        eficiencia_pct = (
            round((uds_hora_real / uds_hora_objetivo) * 100, 1)
            if uds_hora_real and uds_hora_objetivo
            else None
        )

        # Nombre del operario
        usuario_db = db.query(UsuarioDB).filter(
            UsuarioDB.codigo == ex.usuario
        ).first()
        nombre_operario = usuario_db.nombre if usuario_db else ex.usuario

        resultado.append({
            "id": ex.id,
            "usuario": ex.usuario,
            "nombre_operario": nombre_operario,
            "fecha": str(ex.fecha),
            "tarea_principal": ex.tarea_principal,
            "subtarea": ex.subtarea,
            "unidades": ex.unidades,
            "minutos_reales": minutos_reales,
            "horas_reales": horas_reales,
            "uds_hora_real": uds_hora_real,
            "uds_hora_objetivo": uds_hora_objetivo,
            "eficiencia_pct": eficiencia_pct
        })

    # Ordenación
    campos_validos = {"uds_hora_real", "eficiencia_pct", "unidades", "horas_reales", "fecha"}
    if ordenar_por not in campos_validos:
        ordenar_por = "uds_hora_real"

    resultado = sorted(
        resultado,
        key=lambda x: (x[ordenar_por] is None, x[ordenar_por] or 0),
        reverse=(orden == "desc")
    )

    return resultado


@app.get("/productividad/resumen/")
def resumen_productividad(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    db: Session = Depends(get_db)
):
    """
    Resumen agregado por operario: total unidades, media uds/hora, eficiencia media.
    """
    if not desde:
        desde = date.today() - timedelta(days=30)
    if not hasta:
        hasta = date.today()
    validar_rango_fechas(desde, hasta)

    detalle = calcular_productividad(
        desde=desde, hasta=hasta, db=db,
        ordenar_por="uds_hora_real", orden="desc"
    )

    resumen = {}
    for d in detalle:
        u = d["usuario"]
        if u not in resumen:
            resumen[u] = {
                "usuario": u,
                "nombre_operario": d["nombre_operario"],
                "total_unidades": 0,
                "total_horas": 0,
                "eficiencias": [],
                "uds_hora_values": []
            }
        resumen[u]["total_unidades"] += d["unidades"]
        resumen[u]["total_horas"] += d["horas_reales"]
        if d["eficiencia_pct"] is not None:
            resumen[u]["eficiencias"].append(d["eficiencia_pct"])
        if d["uds_hora_real"] is not None:
            resumen[u]["uds_hora_values"].append(d["uds_hora_real"])

    resultado = []
    for u, datos in resumen.items():
        uds_hora_media = (
            round(sum(datos["uds_hora_values"]) / len(datos["uds_hora_values"]), 1)
            if datos["uds_hora_values"] else None
        )
        eficiencia_media = (
            round(sum(datos["eficiencias"]) / len(datos["eficiencias"]), 1)
            if datos["eficiencias"] else None
        )
        resultado.append({
            "usuario": datos["usuario"],
            "nombre_operario": datos["nombre_operario"],
            "total_unidades": datos["total_unidades"],
            "total_horas": round(datos["total_horas"], 2),
            "uds_hora_media": uds_hora_media,
            "eficiencia_media_pct": eficiencia_media
        })

    return sorted(
        resultado,
        key=lambda x: (x["uds_hora_media"] is None, x["uds_hora_media"] or 0),
        reverse=True
    )


@app.delete("/productividad/limpiar/", status_code=204)
def limpiar_productividad(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    db: Session = Depends(get_db)
):
    """Elimina registros de productividad Excel en un rango de fechas."""
    query = db.query(ProductividadExcelDB)
    if desde:
        query = query.filter(ProductividadExcelDB.fecha >= desde)
    if hasta:
        query = query.filter(ProductividadExcelDB.fecha <= hasta)
    query.delete()
    db.commit()
# Reemplaza la función subir_excel_productividad de la Parte 3 con esta versión:

@app.post("/productividad/subir-excel/")
async def subir_excel_productividad(
    file: UploadFile = File(...),
    subido_por: str = "ADMIN",
    db: Session = Depends(get_db)
):
    """
    Sube un Excel con columnas obligatorias:
    fecha | operario | tarea | subtarea | unidades
    Si ya existe fecha+operario+tarea+subtarea, sobreescribe las unidades.
    """
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="El fichero debe ser .xlsx o .xls")

    contenido = await file.read()
    try:
        df = pd.read_excel(io.BytesIO(contenido))
    except Exception:
        raise HTTPException(status_code=400, detail="No se pudo leer el fichero Excel")

    df.columns = [c.strip().lower() for c in df.columns]

    columnas_requeridas = {"fecha", "operario", "tarea", "subtarea", "unidades"}
    if not columnas_requeridas.issubset(set(df.columns)):
        raise HTTPException(
            status_code=400,
            detail=f"El Excel debe tener las columnas: {', '.join(columnas_requeridas)}"
        )

    df = df.dropna(subset=["fecha", "operario", "tarea", "subtarea", "unidades"])
    df["fecha"] = pd.to_datetime(df["fecha"]).dt.date
    df["unidades"] = pd.to_numeric(df["unidades"], errors="coerce").fillna(0).astype(int)
    df = df[df["unidades"] > 0]

    insertados = 0
    actualizados = 0
    errores = []

    for _, fila in df.iterrows():
        try:
            usuario_db = db.query(UsuarioDB).filter(
                UsuarioDB.codigo == str(fila["operario"]).strip()
            ).first()
            if not usuario_db:
                errores.append(f"Operario no encontrado: {fila['operario']}")
                continue

            existente = db.query(ProductividadExcelDB).filter(
                ProductividadExcelDB.fecha == fila["fecha"],
                ProductividadExcelDB.usuario == str(fila["operario"]).strip(),
                ProductividadExcelDB.tarea_principal == str(fila["tarea"]).strip(),
                ProductividadExcelDB.subtarea == str(fila["subtarea"]).strip()
            ).first()

            if existente:
                existente.unidades = int(fila["unidades"])
                existente.subido_por = subido_por
                actualizados += 1
            else:
                nuevo = ProductividadExcelDB(
                    fecha=fila["fecha"],
                    usuario=str(fila["operario"]).strip(),
                    tarea_principal=str(fila["tarea"]).strip(),
                    subtarea=str(fila["subtarea"]).strip(),
                    unidades=int(fila["unidades"]),
                    subido_por=subido_por
                )
                db.add(nuevo)
                insertados += 1

        except Exception as e:
            errores.append(f"Fila con error: {str(e)}")

    db.commit()

    return {
        "insertados": insertados,
        "actualizados": actualizados,
        "errores": errores,
        "total_filas": len(df),
        "mensaje": f"{insertados} registros nuevos, {actualizados} actualizados"
    }


# ==================== ESTADÍSTICAS ====================

@app.get("/estadisticas/productividad/")
def productividad(fecha_ref: Optional[date] = None, db: Session = Depends(get_db)):
    if not fecha_ref:
        fecha_ref = date.today()
    ayer = fecha_ref - timedelta(days=1)
    minutos_jornada = get_minutos_jornada(db)
    usuarios = db.query(UsuarioDB).filter(UsuarioDB.rol == "operario").all()
    resultado = []
    for u in usuarios:
        hoy_mins = db.query(RegistroDB).filter(
            RegistroDB.usuario == u.codigo,
            RegistroDB.fecha == fecha_ref
        ).with_entities(RegistroDB.tiempo_minutos).all()
        ayer_mins = db.query(RegistroDB).filter(
            RegistroDB.usuario == u.codigo,
            RegistroDB.fecha == ayer
        ).with_entities(RegistroDB.tiempo_minutos).all()
        total_hoy = sum(r.tiempo_minutos for r in hoy_mins)
        total_ayer = sum(r.tiempo_minutos for r in ayer_mins)
        resultado.append({
            "usuario": u.codigo,
            "nombre": u.nombre,
            "minutos_hoy": total_hoy,
            "minutos_ayer": total_ayer,
            "variacion_minutos": total_hoy - total_ayer,
            "porcentaje_jornada": round((total_hoy / minutos_jornada) * 100, 1)
        })
    return resultado

@app.get("/estadisticas/top-tareas/")
def top_tareas(
    fecha_desde: Optional[date] = None,
    fecha_hasta: Optional[date] = None,
    db: Session = Depends(get_db)
):
    if not fecha_desde:
        fecha_desde = date.today() - timedelta(days=30)
    if not fecha_hasta:
        fecha_hasta = date.today()
    validar_rango_fechas(fecha_desde, fecha_hasta)
    registros = db.query(RegistroDB).filter(
        RegistroDB.fecha >= fecha_desde,
        RegistroDB.fecha <= fecha_hasta
    ).all()
    por_tarea = {}
    for r in registros:
        clave = f"{r.tarea_principal} - {r.subtarea}"
        if clave not in por_tarea:
            por_tarea[clave] = {"count": 0, "minutos": 0}
        por_tarea[clave]["count"] += 1
        por_tarea[clave]["minutos"] += r.tiempo_minutos
    return [
        {
            "tarea": k,
            "count": v["count"],
            "minutos_totales": v["minutos"],
            "media_minutos": round(v["minutos"] / v["count"], 1)
        }
        for k, v in sorted(
            por_tarea.items(),
            key=lambda x: x[1]["minutos"],
            reverse=True
        )[:10]
    ]

@app.get("/estadisticas/media-por-tarea/")
def media_por_tarea(usuario: Optional[str] = None, db: Session = Depends(get_db)):
    query = db.query(RegistroDB)
    if usuario:
        query = query.filter(RegistroDB.usuario == usuario)
    por_tarea = {}
    for r in query.all():
        if r.tarea_principal not in por_tarea:
            por_tarea[r.tarea_principal] = {"count": 0, "minutos": 0}
        por_tarea[r.tarea_principal]["count"] += 1
        por_tarea[r.tarea_principal]["minutos"] += r.tiempo_minutos
    return [
        {
            "tarea": k,
            "count": v["count"],
            "media_minutos": round(v["minutos"] / v["count"], 1),
            "total_minutos": v["minutos"]
        }
        for k, v in por_tarea.items()
    ]

# ==================== RENDIMIENTO ====================

@app.get("/rendimiento/ranking/")
def ranking_operarios(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    db: Session = Depends(get_db)
):
    if not desde:
        desde = date.today() - timedelta(days=7)
    if not hasta:
        hasta = date.today()
    validar_rango_fechas(desde, hasta)
    minutos_jornada = get_minutos_jornada(db)
    usuarios = db.query(UsuarioDB).filter(UsuarioDB.rol == "operario").all()
    registros = db.query(RegistroDB).filter(
        RegistroDB.fecha >= desde,
        RegistroDB.fecha <= hasta
    ).all()
    dias_rango = (hasta - desde).days + 1
    dias_laborables = sum(
        1 for i in range(dias_rango)
        if (desde + timedelta(days=i)).weekday() < 5
    )
    resultado = []
    for u in usuarios:
        regs_u = [r for r in registros if r.usuario == u.codigo]
        total_mins = sum(r.tiempo_minutos for r in regs_u)
        dias_activo = len(set(str(r.fecha) for r in regs_u))
        max_mins = dias_laborables * minutos_jornada
        pct_jornada = round((total_mins / max_mins) * 100, 1) if max_mins > 0 else 0
        consistencia = round((dias_activo / dias_laborables) * 100, 1) if dias_laborables > 0 else 0
        media_diaria = round(total_mins / dias_activo, 0) if dias_activo > 0 else 0
        resultado.append({
            "usuario": u.codigo,
            "nombre": u.nombre,
            "total_minutos": total_mins,
            "total_horas": round(total_mins / 60, 1),
            "dias_activo": dias_activo,
            "dias_laborables": dias_laborables,
            "pct_jornada_periodo": pct_jornada,
            "consistencia_pct": consistencia,
            "media_minutos_dia": media_diaria
        })
    return sorted(resultado, key=lambda x: x["total_minutos"], reverse=True)

@app.get("/rendimiento/eficiencia-diaria/")
def eficiencia_diaria(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    usuario: Optional[str] = None,
    db: Session = Depends(get_db)
):
    if not desde:
        desde = date.today() - timedelta(days=7)
    if not hasta:
        hasta = date.today()
    validar_rango_fechas(desde, hasta)
    minutos_jornada = get_minutos_jornada(db)
    query = db.query(RegistroDB).filter(
        RegistroDB.fecha >= desde,
        RegistroDB.fecha <= hasta
    )
    if usuario:
        query = query.filter(RegistroDB.usuario == usuario)
    registros = query.all()
    por_dia = {}
    for r in registros:
        dia = str(r.fecha)
        if dia not in por_dia:
            por_dia[dia] = {"total_minutos": 0, "por_tarea": {}}
        por_dia[dia]["total_minutos"] += r.tiempo_minutos
        if r.tarea_principal not in por_dia[dia]["por_tarea"]:
            por_dia[dia]["por_tarea"][r.tarea_principal] = 0
        por_dia[dia]["por_tarea"][r.tarea_principal] += r.tiempo_minutos
    resultado = []
    for dia, datos in sorted(por_dia.items()):
        pct_jornada = round((datos["total_minutos"] / minutos_jornada) * 100, 1)
        resultado.append({
            "fecha": dia,
            "total_minutos": datos["total_minutos"],
            "total_horas": round(datos["total_minutos"] / 60, 2),
            "pct_jornada": pct_jornada,
            "por_tarea": datos["por_tarea"]
        })
    return resultado

@app.get("/rendimiento/tendencia-semanal/")
def tendencia_semanal(
    semanas: Optional[int] = 4,
    usuario: Optional[str] = None,
    db: Session = Depends(get_db)
):
    hasta = date.today()
    desde = hasta - timedelta(weeks=semanas)
    query = db.query(RegistroDB).filter(
        RegistroDB.fecha >= desde,
        RegistroDB.fecha <= hasta
    )
    if usuario:
        query = query.filter(RegistroDB.usuario == usuario)
    registros = query.all()
    semanas_data = {}
    for r in registros:
        iso = r.fecha.isocalendar()
        clave = f"{iso[0]}-S{iso[1]:02d}"
        if clave not in semanas_data:
            semanas_data[clave] = {"minutos": 0, "dias": set(), "usuarios": set()}
        semanas_data[clave]["minutos"] += r.tiempo_minutos
        semanas_data[clave]["dias"].add(str(r.fecha))
        semanas_data[clave]["usuarios"].add(r.usuario)
    return [
        {
            "semana": k,
            "total_minutos": v["minutos"],
            "total_horas": round(v["minutos"] / 60, 1),
            "dias_activos": len(v["dias"]),
            "usuarios_activos": len(v["usuarios"]),
            "media_horas_dia": round(v["minutos"] / 60 / len(v["dias"]), 1) if v["dias"] else 0
        }
        for k, v in sorted(semanas_data.items())
    ]

@app.get("/rendimiento/alertas/")
def alertas_rendimiento(db: Session = Depends(get_db)):
    hoy = date.today()
    hace_7 = hoy - timedelta(days=7)
    minutos_jornada = get_minutos_jornada(db)
    usuarios = db.query(UsuarioDB).filter(UsuarioDB.rol == "operario").all()
    registros_hoy = db.query(RegistroDB).filter(RegistroDB.fecha == hoy).all()
    registros_semana = db.query(RegistroDB).filter(
        RegistroDB.fecha >= hace_7,
        RegistroDB.fecha <= hoy
    ).all()
    alertas = []
    usuarios_con_actividad_hoy = set(r.usuario for r in registros_hoy)
    for u in usuarios:
        mins_hoy = sum(r.tiempo_minutos for r in registros_hoy if r.usuario == u.codigo)
        dias_activo_semana = len(set(
            str(r.fecha) for r in registros_semana if r.usuario == u.codigo
        ))
        pct_hoy = round((mins_hoy / minutos_jornada) * 100, 1)
        if u.codigo not in usuarios_con_actividad_hoy:
            alertas.append({
                "tipo": "sin_actividad",
                "nivel": "warning",
                "usuario": u.codigo,
                "nombre": u.nombre,
                "mensaje": f"{u.nombre} no tiene actividad registrada hoy"
            })
        if mins_hoy > 0 and pct_hoy < 70:
            alertas.append({
                "tipo": "jornada_incompleta",
                "nivel": "warning",
                "usuario": u.codigo,
                "nombre": u.nombre,
                "mensaje": f"{u.nombre} lleva solo {pct_hoy}% de la jornada ({mins_hoy} min)"
            })
        if 0 < dias_activo_semana < 3:
            alertas.append({
                "tipo": "baja_consistencia",
                "nivel": "danger",
                "usuario": u.codigo,
                "nombre": u.nombre,
                "mensaje": f"{u.nombre} solo ha registrado {dias_activo_semana} días esta semana"
            })
    volumenes = db.query(VolumenDB).filter(
        VolumenDB.fecha >= hace_7,
        VolumenDB.fecha <= hoy
    ).all()
    for v in volumenes:
        horas_reales = sum(
            r.tiempo_minutos for r in registros_semana
            if str(r.fecha) == str(v.fecha) and r.tarea_principal == v.tarea_principal
        ) / 60
        if horas_reales > 0:
            eficiencia = round((v.horas_teoricas / horas_reales) * 100, 1)
            if eficiencia < 70:
                alertas.append({
                    "tipo": "eficiencia_baja",
                    "nivel": "danger",
                    "usuario": None,
                    "nombre": v.tarea_principal,
                    "mensaje": f"Eficiencia baja en {v.tarea_principal} el {v.fecha}: {eficiencia}%"
                })
            elif abs(horas_reales - v.horas_teoricas) / v.horas_teoricas > 0.20:
                alertas.append({
                    "tipo": "desviacion",
                    "nivel": "warning",
                    "usuario": None,
                    "nombre": v.tarea_principal,
                    "mensaje": f"Desviación >20% en {v.tarea_principal} el {v.fecha}: "
                               f"{round(horas_reales - v.horas_teoricas, 1)}h"
                })
    return {"total": len(alertas), "alertas": alertas}

@app.get("/rendimiento/objetivo-vs-real/")
def objetivo_vs_real(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    db: Session = Depends(get_db)
):
    if not desde:
        desde = date.today() - timedelta(days=30)
    if not hasta:
        hasta = date.today()
    validar_rango_fechas(desde, hasta)
    registros = db.query(RegistroDB).filter(
        RegistroDB.fecha >= desde,
        RegistroDB.fecha <= hasta
    ).all()
    objetivos = {o.tarea_principal: o for o in db.query(ObjetivoDB).all()}
    por_tarea = {}
    for r in registros:
        if r.tarea_principal not in por_tarea:
            por_tarea[r.tarea_principal] = {"minutos": 0, "count": 0}
        por_tarea[r.tarea_principal]["minutos"] += r.tiempo_minutos
        por_tarea[r.tarea_principal]["count"] += 1
    resultado = []
    for tarea, datos in por_tarea.items():
        obj = objetivos.get(tarea)
        horas_reales = round(datos["minutos"] / 60, 2)
        horas_objetivo = round(obj.horas_jornada, 2) if obj else None
        pct = round((horas_reales / horas_objetivo) * 100, 1) if horas_objetivo else None
        resultado.append({
            "tarea": tarea,
            "horas_reales": horas_reales,
            "horas_objetivo": horas_objetivo,
            "uds_hora_objetivo": obj.uds_hora if obj else None,
            "pct_objetivo": pct,
            "registros": datos["count"]
        })
    return sorted(resultado, key=lambda x: x["horas_reales"], reverse=True)

# ==================== EXCEL EXPORTACIÓN ====================

@app.get("/exportar-excel/")
def exportar_excel(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    usuario: Optional[str] = None,
    db: Session = Depends(get_db)
):
    if desde and hasta:
        validar_rango_fechas(desde, hasta)
    query = db.query(RegistroDB)
    if desde:
        query = query.filter(RegistroDB.fecha >= desde)
    if hasta:
        query = query.filter(RegistroDB.fecha <= hasta)
    if usuario:
        query = query.filter(RegistroDB.usuario == usuario)
    registros = query.order_by(RegistroDB.fecha.desc()).all()

    wb = openpyxl.Workbook()
    HEADER_FILL = PatternFill("solid", fgColor="1e293b")
    HEADER_FONT = Font(bold=True, color="FFFFFF")
    HEADER_ALIGN = Alignment(horizontal="center")

    def estilo_cabecera(ws):
        for cell in ws[1]:
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = HEADER_ALIGN

    def autoajustar(ws):
        for col in ws.columns:
            max_length = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_length + 4, 40)

    # Hoja 0: Resumen Ejecutivo
    ws0 = wb.active
    ws0.title = "Resumen Ejecutivo"
    ws0.append(["RESUMEN EJECUTIVO", ""])
    ws0["A1"].font = Font(bold=True, size=14, color="1e293b")
    ws0.append(["Generado el", str(date.today())])
    ws0.append(["Periodo", f"{desde or 'Todo'} → {hasta or 'Todo'}"])
    ws0.append(["Usuario", usuario or "Todos"])
    ws0.append([])
    ws0.append(["MÉTRICAS GLOBALES", ""])
    ws0["A6"].font = Font(bold=True, color="FFFFFF")
    ws0["A6"].fill = PatternFill("solid", fgColor="3b82f6")
    total_mins = sum(r.tiempo_minutos for r in registros)
    usuarios_unicos = len(set(r.usuario for r in registros))
    dias_unicos = len(set(str(r.fecha) for r in registros))
    ws0.append(["Total registros", len(registros)])
    ws0.append(["Total horas", round(total_mins / 60, 2)])
    ws0.append(["Usuarios activos", usuarios_unicos])
    ws0.append(["Días con actividad", dias_unicos])
    ws0.append(["Media horas/día", round(total_mins / 60 / dias_unicos, 2) if dias_unicos else 0])
    autoajustar(ws0)

    # Hoja 1: Registros Detallados
    ws1 = wb.create_sheet("Registros Detallados")
    ws1.append([
        "ID", "Usuario", "Fecha", "Semana", "Mes",
        "Tarea Principal", "Subtarea", "Tiempo (min)",
        "Horas", "Proyecto", "Comentarios"
    ])
    estilo_cabecera(ws1)
    for r in registros:
        semana = r.fecha.isocalendar()[1]
        mes = r.fecha.strftime("%B %Y")
        ws1.append([
            r.id, r.usuario, str(r.fecha), f"Semana {semana}", mes,
            r.tarea_principal, r.subtarea, r.tiempo_minutos,
            round(r.tiempo_minutos / 60, 2),
            r.proyecto or "", r.comentarios or ""
        ])
    ws1.auto_filter.ref = ws1.dimensions
    ws1.freeze_panes = "A2"
    autoajustar(ws1)

    # Hoja 2: Resumen por Usuario
    ws2 = wb.create_sheet("Resumen por Usuario")
    ws2.append([
        "Usuario", "Total Registros", "Total Minutos",
        "Total Horas", "Días Activos", "Media Min/Día", "% Jornada Media"
    ])
    estilo_cabecera(ws2)
    por_usuario = {}
    for r in registros:
        if r.usuario not in por_usuario:
            por_usuario[r.usuario] = {"registros": 0, "minutos": 0, "dias": set()}
        por_usuario[r.usuario]["registros"] += 1
        por_usuario[r.usuario]["minutos"] += r.tiempo_minutos
        por_usuario[r.usuario]["dias"].add(str(r.fecha))
    minutos_jornada = get_minutos_jornada(db)
    for u, d in por_usuario.items():
        dias = len(d["dias"]) or 1
        media_dia = round(d["minutos"] / dias, 1)
        pct = round((media_dia / minutos_jornada) * 100, 1)
        ws2.append([
            u, d["registros"], d["minutos"],
            round(d["minutos"] / 60, 2), len(d["dias"]), media_dia, pct
        ])
    autoajustar(ws2)

    # Hoja 3: Resumen por Tarea
    ws3 = wb.create_sheet("Resumen por Tarea")
    ws3.append([
        "Tarea Principal", "Subtarea", "Total Registros",
        "Total Minutos", "Total Horas", "Media Min/Registro"
    ])
    estilo_cabecera(ws3)
    por_tarea = {}
    for r in registros:
        clave = (r.tarea_principal, r.subtarea)
        if clave not in por_tarea:
            por_tarea[clave] = {"registros": 0, "minutos": 0}
        por_tarea[clave]["registros"] += 1
        por_tarea[clave]["minutos"] += r.tiempo_minutos
    for (tp, st), d in sorted(por_tarea.items()):
        media = round(d["minutos"] / d["registros"], 1) if d["registros"] > 0 else 0
        ws3.append([tp, st, d["registros"], d["minutos"], round(d["minutos"] / 60, 2), media])
    autoajustar(ws3)

    # Hoja 4: Resumen por Día
    ws4 = wb.create_sheet("Resumen por Día")
    ws4.append([
        "Fecha", "Día Semana", "Total Registros",
        "Total Minutos", "Total Horas", "Usuarios Activos", "% Jornada Media"
    ])
    estilo_cabecera(ws4)
    dias_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    por_dia = {}
    for r in registros:
        dia = str(r.fecha)
        if dia not in por_dia:
            por_dia[dia] = {"registros": 0, "minutos": 0, "usuarios": set()}
        por_dia[dia]["registros"] += 1
        por_dia[dia]["minutos"] += r.tiempo_minutos
        por_dia[dia]["usuarios"].add(r.usuario)
    for dia, d in sorted(por_dia.items(), reverse=True):
        fecha_obj = date.fromisoformat(dia)
        nombre_dia = dias_semana[fecha_obj.weekday()]
        n_usuarios = len(d["usuarios"])
        pct = round((d["minutos"] / (n_usuarios * minutos_jornada)) * 100, 1) if n_usuarios else 0
        ws4.append([
            dia, nombre_dia, d["registros"], d["minutos"],
            round(d["minutos"] / 60, 2), n_usuarios, pct
        ])
    autoajustar(ws4)

    # Hoja 5: Métricas Volumen
    ws5 = wb.create_sheet("Métricas Volumen")
    ws5.append([
        "Fecha", "Tarea", "Unidades", "Horas Teóricas", "Horas Reales",
        "Eficiencia %", "Uds/Hora Real", "Uds/Hora Teórica",
        "Desviación (h)", "Comentarios"
    ])
    estilo_cabecera(ws5)
    volumenes = db.query(VolumenDB).order_by(VolumenDB.fecha.desc()).all()
    todos_registros = db.query(RegistroDB).all()
    for v in volumenes:
        horas_reales = sum(
            r.tiempo_minutos for r in todos_registros
            if str(r.fecha) == str(v.fecha) and r.tarea_principal == v.tarea_principal
        ) / 60
        eficiencia = round((v.horas_teoricas / horas_reales) * 100, 1) if horas_reales > 0 else 0
        uds_real = round(v.unidades / horas_reales, 1) if horas_reales > 0 else 0
        uds_teo = round(v.unidades / v.horas_teoricas, 1) if v.horas_teoricas > 0 else 0
        desviacion = round(horas_reales - v.horas_teoricas, 2)
        ws5.append([
            str(v.fecha), v.tarea_principal, v.unidades, v.horas_teoricas,
            round(horas_reales, 2), eficiencia, uds_real, uds_teo,
            desviacion, v.comentarios or ""
        ])
    autoajustar(ws5)

    # Hoja 6: Productividad desde Excel subido
    ws6 = wb.create_sheet("Productividad Operarios")
    ws6.append([
        "Operario", "Fecha", "Tarea", "Subtarea", "Unidades",
        "Horas Reales", "Uds/Hora Real", "Uds/Hora Objetivo", "Eficiencia %"
    ])
    estilo_cabecera(ws6)
    prod_data = calcular_productividad(db=db)
    for p in prod_data:
        ws6.append([
            p["nombre_operario"], p["fecha"], p["tarea_principal"],
            p["subtarea"], p["unidades"], p["horas_reales"],
            p["uds_hora_real"] or "—", p["uds_hora_objetivo"] or "—",
            p["eficiencia_pct"] or "—"
        ])
    autoajustar(ws6)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=registros.xlsx"}
    )

# ==================== ARRANQUE ====================

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port)

