from datetime import date, timedelta
from typing import List, Optional
from fastapi import FastAPI, HTTPException, Depends
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from sqlalchemy import create_engine, Column, Integer, String, Date, DateTime, Float, Boolean, Text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import io
import os
from datetime import datetime

# ==================== BASE DE DATOS ====================

DATABASE_URL = os.environ.get("DATABASE_URL")

if DATABASE_URL:
    if DATABASE_URL.startswith("postgres://"):
        DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)
    SQLALCHEMY_DATABASE_URL = DATABASE_URL
    engine = create_engine(SQLALCHEMY_DATABASE_URL)
else:
    SQLALCHEMY_DATABASE_URL = "sqlite:///./registros.db"
    engine = create_engine(
        SQLALCHEMY_DATABASE_URL,
        connect_args={"check_same_thread": False}
    )

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# ==================== MODELOS ====================

class RegistroDB(Base):
    __tablename__ = "registros"
    id               = Column(Integer, primary_key=True, index=True)
    usuario          = Column(String, index=True)
    fecha            = Column(Date)
    tarea_principal  = Column(String, index=True)
    subtarea         = Column(String)
    tiempo_minutos   = Column(Integer)
    proyecto         = Column(String, nullable=True)
    comentarios      = Column(String, nullable=True)
    created_at       = Column(DateTime, default=datetime.utcnow)
    updated_at       = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

class UsuarioDB(Base):
    __tablename__ = "usuarios"
    id             = Column(Integer, primary_key=True, index=True)
    codigo         = Column(String, unique=True, index=True)
    nombre         = Column(String)
    pin            = Column(String)
    rol            = Column(String, default="operario")

class PlantillaDB(Base):
    __tablename__ = "plantillas"
    id          = Column(Integer, primary_key=True, index=True)
    nombre      = Column(String)
    usuario     = Column(String)
    created_at  = Column(DateTime, default=datetime.utcnow)

class PlantillaItemDB(Base):
    __tablename__ = "plantilla_items"
    id               = Column(Integer, primary_key=True, index=True)
    plantilla_id     = Column(Integer)
    tarea_principal  = Column(String)
    subtarea         = Column(String)
    tiempo_minutos   = Column(Integer)
    proyecto         = Column(String, nullable=True)
    comentarios      = Column(String, nullable=True)

class VolumenDB(Base):
    __tablename__ = "volumenes"
    id               = Column(Integer, primary_key=True, index=True)
    fecha            = Column(Date, index=True)
    tarea_principal  = Column(String, index=True)
    unidades         = Column(Integer)
    horas_teoricas   = Column(Float)
    creado_por       = Column(String)
    comentarios      = Column(String, nullable=True)
    created_at       = Column(DateTime, default=datetime.utcnow)
    updated_at       = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

class ObjetivoDB(Base):
    __tablename__ = "objetivos"
    id               = Column(Integer, primary_key=True, index=True)
    tarea_principal  = Column(String, unique=True, index=True)
    unidades_hora    = Column(Float)
    activo           = Column(Boolean, default=True)
    created_at       = Column(DateTime, default=datetime.utcnow)
    updated_at       = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

class IncidenciaDB(Base):
    __tablename__ = "incidencias"
    id               = Column(Integer, primary_key=True, index=True)
    fecha            = Column(Date, index=True)
    tipo             = Column(String)
    tarea_principal  = Column(String, nullable=True)
    impacto_minutos  = Column(Integer, nullable=True)
    descripcion      = Column(Text)
    creado_por       = Column(String)
    created_at       = Column(DateTime, default=datetime.utcnow)

class TurnoUsuarioDB(Base):
    __tablename__ = "turnos_usuarios"
    id               = Column(Integer, primary_key=True, index=True)
    usuario          = Column(String, index=True)
    fecha            = Column(Date, index=True)
    turno            = Column(String)
    horas_previstas  = Column(Float, default=8.0)
    estado           = Column(String, default="trabaja")
    comentarios      = Column(String, nullable=True)
    created_at       = Column(DateTime, default=datetime.utcnow)

class AuditoriaDB(Base):
    __tablename__ = "auditoria"
    id               = Column(Integer, primary_key=True, index=True)
    usuario          = Column(String)
    accion           = Column(String)
    entidad          = Column(String)
    entidad_id       = Column(Integer, nullable=True)
    detalle          = Column(Text, nullable=True)
    created_at       = Column(DateTime, default=datetime.utcnow)

Base.metadata.create_all(bind=engine)

# Crear datos por defecto
with SessionLocal() as session:
    admin_existente = session.query(UsuarioDB).filter(UsuarioDB.rol == "supervisor").first()
    if not admin_existente:
        session.add(UsuarioDB(codigo="ADMIN", nombre="Administrador", pin="1234", rol="supervisor"))
        session.commit()

    objetivos_default = [
        ("Picking", 80.0),
        ("Packing", 60.0),
        ("Inbound", 50.0),
        ("Shipping", 70.0),
        ("Ecommerce", 65.0)
    ]
    for tarea, uds_hora in objetivos_default:
        if not session.query(ObjetivoDB).filter(ObjetivoDB.tarea_principal == tarea).first():
            session.add(ObjetivoDB(tarea_principal=tarea, unidades_hora=uds_hora))
    session.commit()

app = FastAPI(title="API de Registro de Tareas", version="9.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
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
    horas_teoricas: Optional[float] = None
    creado_por: str
    comentarios: Optional[str] = None

class VolumenRespuesta(BaseModel):
    id: int
    fecha: date
    tarea_principal: str
    unidades: int
    horas_teoricas: float
    creado_por: str
    comentarios: Optional[str] = None
    created_at: Optional[datetime] = None
    class Config:
        from_attributes = True

class ObjetivoCreate(BaseModel):
    tarea_principal: str
    unidades_hora: float
    activo: bool = True

class ObjetivoRespuesta(ObjetivoCreate):
    id: int
    created_at: Optional[datetime] = None
    updated_at: Optional[datetime] = None
    class Config:
        from_attributes = True

class IncidenciaCreate(BaseModel):
    fecha: date
    tipo: str
    tarea_principal: Optional[str] = None
    impacto_minutos: Optional[int] = None
    descripcion: str
    creado_por: str

class IncidenciaRespuesta(IncidenciaCreate):
    id: int
    created_at: Optional[datetime] = None
    class Config:
        from_attributes = True

class TurnoCreate(BaseModel):
    usuario: str
    fecha: date
    turno: str
    horas_previstas: float = 8.0
    estado: str = "trabaja"
    comentarios: Optional[str] = None

class TurnoRespuesta(TurnoCreate):
    id: int
    created_at: Optional[datetime] = None
    class Config:
        from_attributes = True

class AuditoriaRespuesta(BaseModel):
    id: int
    usuario: str
    accion: str
    entidad: str
    entidad_id: Optional[int] = None
    detalle: Optional[str] = None
    created_at: Optional[datetime] = None
    class Config:
        from_attributes = True

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def registrar_auditoria(db: Session, usuario: str, accion: str, entidad: str, entidad_id: int = None, detalle: str = None):
    log = AuditoriaDB(
        usuario=usuario,
        accion=accion,
        entidad=entidad,
        entidad_id=entidad_id,
        detalle=detalle,
        created_at=datetime.utcnow()
    )
    db.add(log)
    db.commit()

# ==================== RUTAS GENERALES ====================

@app.get("/")
def leer_index():
    ruta = os.path.join(os.path.dirname(__file__), "index.html")
    if os.path.exists(ruta):
        return FileResponse(ruta)
    return {"error": "No se encuentra el archivo index.html"}

# ==================== LOGIN ====================

@app.post("/login/")
def login(datos: LoginRequest, db: Session = Depends(get_db)):
    usuario = db.query(UsuarioDB).filter(
        UsuarioDB.codigo == datos.codigo,
        UsuarioDB.pin == datos.pin
    ).first()
    if not usuario:
        raise HTTPException(status_code=401, detail="Código o PIN incorrectos")
    registrar_auditoria(db, datos.codigo, "LOGIN", "usuarios")
    return {
        "id": usuario.id,
        "codigo": usuario.codigo,
        "nombre": usuario.nombre,
        "rol": usuario.rol
    }

# ==================== USUARIOS ====================

@app.post("/usuarios/", response_model=UsuarioRespuesta, status_code=201)
def crear_usuario(usuario: UsuarioCreate, db: Session = Depends(get_db)):
    existente = db.query(UsuarioDB).filter(UsuarioDB.codigo == usuario.codigo).first()
    if existente:
        raise HTTPException(status_code=400, detail="Ya existe un usuario con ese código")
    nuevo = UsuarioDB(**usuario.dict())
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    registrar_auditoria(db, usuario.codigo, "CREAR", "usuarios", nuevo.id, f"Nuevo usuario: {usuario.nombre}")
    return nuevo

@app.get("/usuarios/", response_model=List[UsuarioRespuesta])
def obtener_usuarios(db: Session = Depends(get_db)):
    return db.query(UsuarioDB).all()

@app.delete("/usuarios/{id}", status_code=204)
def eliminar_usuario(id: int, db: Session = Depends(get_db)):
    usuario = db.query(UsuarioDB).filter(UsuarioDB.id == id).first()
    if not usuario:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    registrar_auditoria(db, "ADMIN", "ELIMINAR", "usuarios", id, f"Usuario eliminado: {usuario.nombre}")
    db.delete(usuario)
    db.commit()

# ==================== REGISTROS ====================

@app.post("/registros/", response_model=RegistroTareaRespuesta, status_code=201)
def crear_registro(registro: RegistroTarea, db: Session = Depends(get_db)):
    total_minutos = db.query(RegistroDB).filter(
        RegistroDB.usuario == registro.usuario,
        RegistroDB.fecha == registro.fecha
    ).with_entities(RegistroDB.tiempo_minutos).all()
    total = sum(r.tiempo_minutos for r in total_minutos)
    if total + registro.tiempo_minutos > 480:
        raise HTTPException(status_code=400, detail=f"No se pueden superar 8 horas. Minutos ya registrados: {total}")
    nuevo = RegistroDB(**registro.dict())
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    registrar_auditoria(db, registro.usuario, "CREAR", "registros", nuevo.id, f"{registro.tarea_principal} - {registro.subtarea} ({registro.tiempo_minutos} min)")
    return nuevo

@app.get("/registros/", response_model=List[RegistroTareaRespuesta])
def obtener_registros(db: Session = Depends(get_db)):
    return db.query(RegistroDB).all()

@app.get("/registros/rango/", response_model=List[RegistroTareaRespuesta])
def obtener_registros_rango(
    desde: date,
    hasta: date,
    usuario: Optional[str] = None,
    db: Session = Depends(get_db)
):
    query = db.query(RegistroDB).filter(
        RegistroDB.fecha >= desde,
        RegistroDB.fecha <= hasta
    )
    if usuario:
        query = query.filter(RegistroDB.usuario == usuario)
    return query.order_by(RegistroDB.fecha.desc()).all()

@app.get("/registros/resumen-semanal/")
def resumen_semanal(usuario: Optional[str] = None, fecha_ref: Optional[date] = None, db: Session = Depends(get_db)):
    if not fecha_ref:
        fecha_ref = date.today()
    inicio_semana = fecha_ref - timedelta(days=fecha_ref.weekday())
    fin_semana = inicio_semana + timedelta(days=6)
    query = db.query(RegistroDB).filter(
        RegistroDB.fecha >= inicio_semana,
        RegistroDB.fecha <= fin_semana
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
        resumen[dia]["tareas"][r.tarea_principal] = resumen[dia]["tareas"].get(r.tarea_principal, 0) + r.tiempo_minutos
    return {
        "inicio": str(inicio_semana),
        "fin": str(fin_semana),
        "dias": resumen,
        "total_minutos": sum(r.tiempo_minutos for r in registros)
    }

@app.get("/registros/resumen-mensual/")
def resumen_mensual(usuario: Optional[str] = None, anyo: Optional[int] = None, mes: Optional[int] = None, db: Session = Depends(get_db)):
    hoy = date.today()
    if not anyo: anyo = hoy.year
    if not mes: mes = hoy.month
    inicio_mes = date(anyo, mes, 1)
    fin_mes = date(anyo + 1, 1, 1) - timedelta(days=1) if mes == 12 else date(anyo, mes + 1, 1) - timedelta(days=1)
    query = db.query(RegistroDB).filter(RegistroDB.fecha >= inicio_mes, RegistroDB.fecha <= fin_mes)
    if usuario:
        query = query.filter(RegistroDB.usuario == usuario)
    registros = query.all()
    por_usuario = {}
    por_tarea = {}
    for r in registros:
        por_usuario[r.usuario] = por_usuario.get(r.usuario, 0) + r.tiempo_minutos
        por_tarea[r.tarea_principal] = por_tarea.get(r.tarea_principal, 0) + r.tiempo_minutos
    return {
        "anyo": anyo, "mes": mes,
        "total_minutos": sum(r.tiempo_minutos for r in registros),
        "total_registros": len(registros),
        "por_usuario": por_usuario,
        "por_tarea": por_tarea
    }

@app.get("/registros/{usuario}", response_model=List[RegistroTareaRespuesta])
def obtener_registros_por_usuario(usuario: str, db: Session = Depends(get_db)):
    resultado = db.query(RegistroDB).filter(RegistroDB.usuario.ilike(usuario)).all()
    if not resultado:
        raise HTTPException(status_code=404, detail="No se encontraron registros")
    return resultado

@app.put("/registros/{id}", response_model=RegistroTareaRespuesta)
def editar_registro(id: int, registro: RegistroTarea, db: Session = Depends(get_db)):
    db_registro = db.query(RegistroDB).filter(RegistroDB.id == id).first()
    if not db_registro:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    total_minutos = db.query(RegistroDB).filter(
        RegistroDB.usuario == registro.usuario,
        RegistroDB.fecha == registro.fecha,
        RegistroDB.id != id
    ).with_entities(RegistroDB.tiempo_minutos).all()
    total = sum(r.tiempo_minutos for r in total_minutos)
    if total + registro.tiempo_minutos > 480:
        raise HTTPException(status_code=400, detail=f"No se pueden superar 8 horas. Minutos ya registrados: {total}")
    detalle_anterior = f"{db_registro.tarea_principal} - {db_registro.subtarea} ({db_registro.tiempo_minutos} min)"
    db_registro.usuario         = registro.usuario
    db_registro.fecha           = registro.fecha
    db_registro.tarea_principal = registro.tarea_principal
    db_registro.subtarea        = registro.subtarea
    db_registro.tiempo_minutos  = registro.tiempo_minutos
    db_registro.proyecto        = registro.proyecto
    db_registro.comentarios     = registro.comentarios
    db_registro.updated_at      = datetime.utcnow()
    db.commit()
    db.refresh(db_registro)
    registrar_auditoria(db, registro.usuario, "EDITAR", "registros", id, f"Antes: {detalle_anterior} | Después: {registro.tarea_principal} - {registro.subtarea} ({registro.tiempo_minutos} min)")
    return db_registro

@app.delete("/registros/{id}", status_code=204)
def eliminar_registro(id: int, db: Session = Depends(get_db)):
    db_registro = db.query(RegistroDB).filter(RegistroDB.id == id).first()
    if not db_registro:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    registrar_auditoria(db, db_registro.usuario, "ELIMINAR", "registros", id, f"{db_registro.tarea_principal} - {db_registro.subtarea} ({db_registro.tiempo_minutos} min)")
    db.delete(db_registro)
    db.commit()

@app.post("/registros/{id}/duplicar", response_model=RegistroTareaRespuesta, status_code=201)
def duplicar_registro(id: int, nueva_fecha: Optional[date] = None, db: Session = Depends(get_db)):
    original = db.query(RegistroDB).filter(RegistroDB.id == id).first()
    if not original:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    fecha_destino = nueva_fecha if nueva_fecha else original.fecha
    total_minutos = db.query(RegistroDB).filter(
        RegistroDB.usuario == original.usuario,
        RegistroDB.fecha == fecha_destino
    ).with_entities(RegistroDB.tiempo_minutos).all()
    total = sum(r.tiempo_minutos for r in total_minutos)
    if total + original.tiempo_minutos > 480:
        raise HTTPException(status_code=400, detail=f"No se pueden superar 8 horas. Minutos ya registrados: {total}")
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
    registrar_auditoria(db, original.usuario, "DUPLICAR", "registros", nuevo.id, f"Duplicado desde ID {id}")
    return nuevo
# ==================== PLANTILLAS ====================

@app.post("/plantillas/", status_code=201)
def crear_plantilla(plantilla: PlantillaCreate, db: Session = Depends(get_db)):
    nueva = PlantillaDB(
        nombre=plantilla.nombre,
        usuario=plantilla.usuario,
        created_at=datetime.utcnow()
    )
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
    registrar_auditoria(db, plantilla.usuario, "CREAR", "plantillas", nueva.id, f"Plantilla: {plantilla.nombre}")
    return {"id": nueva.id, "nombre": nueva.nombre, "usuario": nueva.usuario}

@app.get("/plantillas/")
def obtener_plantillas(usuario: Optional[str] = None, db: Session = Depends(get_db)):
    query = db.query(PlantillaDB)
    if usuario:
        query = query.filter(PlantillaDB.usuario == usuario)
    plantillas = query.all()
    resultado = []
    for p in plantillas:
        items = db.query(PlantillaItemDB).filter(PlantillaItemDB.plantilla_id == p.id).all()
        resultado.append({
            "id": p.id,
            "nombre": p.nombre,
            "usuario": p.usuario,
            "created_at": p.created_at,
            "items": [
                {
                    "id": i.id,
                    "tarea_principal": i.tarea_principal,
                    "subtarea": i.subtarea,
                    "tiempo_minutos": i.tiempo_minutos,
                    "proyecto": i.proyecto,
                    "comentarios": i.comentarios
                } for i in items
            ]
        })
    return resultado

@app.delete("/plantillas/{id}", status_code=204)
def eliminar_plantilla(id: int, db: Session = Depends(get_db)):
    plantilla = db.query(PlantillaDB).filter(PlantillaDB.id == id).first()
    if not plantilla:
        raise HTTPException(status_code=404, detail="Plantilla no encontrada")
    registrar_auditoria(db, plantilla.usuario, "ELIMINAR", "plantillas", id, f"Plantilla eliminada: {plantilla.nombre}")
    db.query(PlantillaItemDB).filter(PlantillaItemDB.plantilla_id == id).delete()
    db.delete(plantilla)
    db.commit()

# ==================== VOLUMEN & MÉTRICAS ====================

@app.post("/volumenes/", response_model=VolumenRespuesta, status_code=201)
def crear_volumen(volumen: VolumenCreate, db: Session = Depends(get_db)):
    horas_teoricas = volumen.horas_teoricas
    if horas_teoricas is None:
        objetivo = db.query(ObjetivoDB).filter(
            ObjetivoDB.tarea_principal == volumen.tarea_principal,
            ObjetivoDB.activo == True
        ).first()
        if objetivo and objetivo.unidades_hora > 0:
            horas_teoricas = round(volumen.unidades / objetivo.unidades_hora, 2)
        else:
            horas_teoricas = 0.0
    nuevo = VolumenDB(
        fecha=volumen.fecha,
        tarea_principal=volumen.tarea_principal,
        unidades=volumen.unidades,
        horas_teoricas=horas_teoricas,
        creado_por=volumen.creado_por,
        comentarios=volumen.comentarios
    )
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    registrar_auditoria(db, volumen.creado_por, "CREAR", "volumenes", nuevo.id, f"{volumen.tarea_principal}: {volumen.unidades} uds - {horas_teoricas}h teóricas")
    return nuevo

@app.get("/volumenes/", response_model=List[VolumenRespuesta])
def obtener_volumenes(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    tarea: Optional[str] = None,
    db: Session = Depends(get_db)
):
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
    horas_teoricas = volumen.horas_teoricas
    if horas_teoricas is None:
        objetivo = db.query(ObjetivoDB).filter(
            ObjetivoDB.tarea_principal == volumen.tarea_principal,
            ObjetivoDB.activo == True
        ).first()
        if objetivo and objetivo.unidades_hora > 0:
            horas_teoricas = round(volumen.unidades / objetivo.unidades_hora, 2)
        else:
            horas_teoricas = 0.0
    db_vol.fecha           = volumen.fecha
    db_vol.tarea_principal = volumen.tarea_principal
    db_vol.unidades        = volumen.unidades
    db_vol.horas_teoricas  = horas_teoricas
    db_vol.creado_por      = volumen.creado_por
    db_vol.comentarios     = volumen.comentarios
    db_vol.updated_at      = datetime.utcnow()
    db.commit()
    db.refresh(db_vol)
    registrar_auditoria(db, volumen.creado_por, "EDITAR", "volumenes", id, f"{volumen.tarea_principal}: {volumen.unidades} uds")
    return db_vol

@app.delete("/volumenes/{id}", status_code=204)
def eliminar_volumen(id: int, db: Session = Depends(get_db)):
    db_vol = db.query(VolumenDB).filter(VolumenDB.id == id).first()
    if not db_vol:
        raise HTTPException(status_code=404, detail="Volumen no encontrado")
    registrar_auditoria(db, db_vol.creado_por, "ELIMINAR", "volumenes", id, f"{db_vol.tarea_principal}: {db_vol.unidades} uds")
    db.delete(db_vol)
    db.commit()

@app.get("/volumenes/metricas/")
def metricas_volumenes(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    db: Session = Depends(get_db)
):
    query = db.query(VolumenDB)
    if desde:
        query = query.filter(VolumenDB.fecha >= desde)
    if hasta:
        query = query.filter(VolumenDB.fecha <= hasta)
    volumenes = query.all()
    registros_query = db.query(RegistroDB)
    if desde:
        registros_query = registros_query.filter(RegistroDB.fecha >= desde)
    if hasta:
        registros_query = registros_query.filter(RegistroDB.fecha <= hasta)
    registros = registros_query.all()
    por_tarea = {}
    for v in volumenes:
        t = v.tarea_principal
        if t not in por_tarea:
            por_tarea[t] = {"unidades": 0, "horas_teoricas": 0.0, "horas_reales": 0.0}
        por_tarea[t]["unidades"] += v.unidades
        por_tarea[t]["horas_teoricas"] += v.horas_teoricas
    for r in registros:
        t = r.tarea_principal
        if t in por_tarea:
            por_tarea[t]["horas_reales"] += round(r.tiempo_minutos / 60, 2)
    resultado = []
    for tarea, datos in por_tarea.items():
        eficiencia = 0.0
        if datos["horas_reales"] > 0:
            eficiencia = round((datos["horas_teoricas"] / datos["horas_reales"]) * 100, 1)
        resultado.append({
            "tarea": tarea,
            "unidades": datos["unidades"],
            "horas_teoricas": round(datos["horas_teoricas"], 2),
            "horas_reales": round(datos["horas_reales"], 2),
            "eficiencia_pct": eficiencia
        })
    return {"desde": str(desde), "hasta": str(hasta), "por_tarea": resultado}

# ==================== OBJETIVOS ====================

@app.post("/objetivos/", response_model=ObjetivoRespuesta, status_code=201)
def crear_objetivo(objetivo: ObjetivoCreate, db: Session = Depends(get_db)):
    existente = db.query(ObjetivoDB).filter(ObjetivoDB.tarea_principal == objetivo.tarea_principal).first()
    if existente:
        raise HTTPException(status_code=400, detail="Ya existe un objetivo para esa tarea")
    nuevo = ObjetivoDB(**objetivo.dict())
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    registrar_auditoria(db, "ADMIN", "CREAR", "objetivos", nuevo.id, f"{objetivo.tarea_principal}: {objetivo.unidades_hora} uds/h")
    return nuevo

@app.get("/objetivos/", response_model=List[ObjetivoRespuesta])
def obtener_objetivos(db: Session = Depends(get_db)):
    return db.query(ObjetivoDB).all()

@app.put("/objetivos/{id}", response_model=ObjetivoRespuesta)
def editar_objetivo(id: int, objetivo: ObjetivoCreate, db: Session = Depends(get_db)):
    db_obj = db.query(ObjetivoDB).filter(ObjetivoDB.id == id).first()
    if not db_obj:
        raise HTTPException(status_code=404, detail="Objetivo no encontrado")
    db_obj.tarea_principal = objetivo.tarea_principal
    db_obj.unidades_hora   = objetivo.unidades_hora
    db_obj.activo          = objetivo.activo
    db_obj.updated_at      = datetime.utcnow()
    db.commit()
    db.refresh(db_obj)
    registrar_auditoria(db, "ADMIN", "EDITAR", "objetivos", id, f"{objetivo.tarea_principal}: {objetivo.unidades_hora} uds/h")
    return db_obj

@app.delete("/objetivos/{id}", status_code=204)
def eliminar_objetivo(id: int, db: Session = Depends(get_db)):
    db_obj = db.query(ObjetivoDB).filter(ObjetivoDB.id == id).first()
    if not db_obj:
        raise HTTPException(status_code=404, detail="Objetivo no encontrado")
    registrar_auditoria(db, "ADMIN", "ELIMINAR", "objetivos", id, f"{db_obj.tarea_principal}")
    db.delete(db_obj)
    db.commit()

# ==================== INCIDENCIAS ====================

@app.post("/incidencias/", response_model=IncidenciaRespuesta, status_code=201)
def crear_incidencia(incidencia: IncidenciaCreate, db: Session = Depends(get_db)):
    nueva = IncidenciaDB(**incidencia.dict())
    db.add(nueva)
    db.commit()
    db.refresh(nueva)
    registrar_auditoria(db, incidencia.creado_por, "CREAR", "incidencias", nueva.id, f"{incidencia.tipo}: {incidencia.descripcion[:50]}")
    return nueva

@app.get("/incidencias/", response_model=List[IncidenciaRespuesta])
def obtener_incidencias(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    tipo: Optional[str] = None,
    db: Session = Depends(get_db)
):
    query = db.query(IncidenciaDB)
    if desde:
        query = query.filter(IncidenciaDB.fecha >= desde)
    if hasta:
        query = query.filter(IncidenciaDB.fecha <= hasta)
    if tipo:
        query = query.filter(IncidenciaDB.tipo == tipo)
    return query.order_by(IncidenciaDB.fecha.desc()).all()

@app.delete("/incidencias/{id}", status_code=204)
def eliminar_incidencia(id: int, db: Session = Depends(get_db)):
    db_inc = db.query(IncidenciaDB).filter(IncidenciaDB.id == id).first()
    if not db_inc:
        raise HTTPException(status_code=404, detail="Incidencia no encontrada")
    registrar_auditoria(db, db_inc.creado_por, "ELIMINAR", "incidencias", id, f"{db_inc.tipo}")
    db.delete(db_inc)
    db.commit()

# ==================== TURNOS ====================

@app.post("/turnos/", response_model=TurnoRespuesta, status_code=201)
def crear_turno(turno: TurnoCreate, db: Session = Depends(get_db)):
    existente = db.query(TurnoUsuarioDB).filter(
        TurnoUsuarioDB.usuario == turno.usuario,
        TurnoUsuarioDB.fecha == turno.fecha
    ).first()
    if existente:
        existente.turno          = turno.turno
        existente.horas_previstas = turno.horas_previstas
        existente.estado         = turno.estado
        existente.comentarios    = turno.comentarios
        db.commit()
        db.refresh(existente)
        registrar_auditoria(db, turno.usuario, "EDITAR", "turnos", existente.id, f"{turno.fecha} - {turno.turno} ({turno.estado})")
        return existente
    nuevo = TurnoUsuarioDB(**turno.dict())
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    registrar_auditoria(db, turno.usuario, "CREAR", "turnos", nuevo.id, f"{turno.fecha} - {turno.turno} ({turno.estado})")
    return nuevo

@app.get("/turnos/", response_model=List[TurnoRespuesta])
def obtener_turnos(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    usuario: Optional[str] = None,
    db: Session = Depends(get_db)
):
    query = db.query(TurnoUsuarioDB)
    if desde:
        query = query.filter(TurnoUsuarioDB.fecha >= desde)
    if hasta:
        query = query.filter(TurnoUsuarioDB.fecha <= hasta)
    if usuario:
        query = query.filter(TurnoUsuarioDB.usuario == usuario)
    return query.order_by(TurnoUsuarioDB.fecha.desc()).all()

@app.get("/turnos/resumen/")
def resumen_turnos(fecha: Optional[date] = None, db: Session = Depends(get_db)):
    if not fecha:
        fecha = date.today()
    inicio_semana = fecha - timedelta(days=fecha.weekday())
    fin_semana = inicio_semana + timedelta(days=6)
    turnos = db.query(TurnoUsuarioDB).filter(
        TurnoUsuarioDB.fecha >= inicio_semana,
        TurnoUsuarioDB.fecha <= fin_semana
    ).all()
    usuarios = db.query(UsuarioDB).all()
    resumen = {}
    for u in usuarios:
        resumen[u.codigo] = {
            "nombre": u.nombre,
            "horas_previstas": 0.0,
            "horas_registradas": 0.0,
            "dias": {}
        }
    for t in turnos:
        if t.usuario in resumen:
            resumen[t.usuario]["horas_previstas"] += t.horas_previstas
            resumen[t.usuario]["dias"][str(t.fecha)] = {
                "turno": t.turno,
                "estado": t.estado,
                "horas_previstas": t.horas_previstas
            }
    registros_semana = db.query(RegistroDB).filter(
        RegistroDB.fecha >= inicio_semana,
        RegistroDB.fecha <= fin_semana
    ).all()
    for r in registros_semana:
        if r.usuario in resumen:
            resumen[r.usuario]["horas_registradas"] += round(r.tiempo_minutos / 60, 2)
    return {
        "semana_inicio": str(inicio_semana),
        "semana_fin": str(fin_semana),
        "usuarios": resumen
    }

@app.delete("/turnos/{id}", status_code=204)
def eliminar_turno(id: int, db: Session = Depends(get_db)):
    db_turno = db.query(TurnoUsuarioDB).filter(TurnoUsuarioDB.id == id).first()
    if not db_turno:
        raise HTTPException(status_code=404, detail="Turno no encontrado")
    registrar_auditoria(db, db_turno.usuario, "ELIMINAR", "turnos", id, f"{db_turno.fecha} - {db_turno.turno}")
    db.delete(db_turno)
    db.commit()

# ==================== AUDITORÍA ====================

@app.get("/auditoria/", response_model=List[AuditoriaRespuesta])
def obtener_auditoria(
    usuario: Optional[str] = None,
    accion: Optional[str] = None,
    entidad: Optional[str] = None,
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    limite: int = 200,
    db: Session = Depends(get_db)
):
    query = db.query(AuditoriaDB)
    if usuario:
        query = query.filter(AuditoriaDB.usuario == usuario)
    if accion:
        query = query.filter(AuditoriaDB.accion == accion)
    if entidad:
        query = query.filter(AuditoriaDB.entidad == entidad)
    if desde:
        query = query.filter(AuditoriaDB.created_at >= datetime.combine(desde, datetime.min.time()))
    if hasta:
        query = query.filter(AuditoriaDB.created_at <= datetime.combine(hasta, datetime.max.time()))
    return query.order_by(AuditoriaDB.created_at.desc()).limit(limite).all()
# ==================== ESTADÍSTICAS ====================

@app.get("/estadisticas/")
def obtener_estadisticas(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    usuario: Optional[str] = None,
    db: Session = Depends(get_db)
):
    query = db.query(RegistroDB)
    if desde:
        query = query.filter(RegistroDB.fecha >= desde)
    if hasta:
        query = query.filter(RegistroDB.fecha <= hasta)
    if usuario:
        query = query.filter(RegistroDB.usuario == usuario)
    registros = query.all()

    if not registros:
        return {
            "total_registros": 0,
            "total_minutos": 0,
            "total_horas": 0.0,
            "por_tarea": {},
            "por_usuario": {},
            "por_dia": {},
            "top_tareas": [],
            "media_diaria_minutos": 0.0
        }

    por_tarea = {}
    por_usuario = {}
    por_dia = {}

    for r in registros:
        # Por tarea
        if r.tarea_principal not in por_tarea:
            por_tarea[r.tarea_principal] = {"minutos": 0, "registros": 0}
        por_tarea[r.tarea_principal]["minutos"] += r.tiempo_minutos
        por_tarea[r.tarea_principal]["registros"] += 1

        # Por usuario
        if r.usuario not in por_usuario:
            por_usuario[r.usuario] = {"minutos": 0, "registros": 0}
        por_usuario[r.usuario]["minutos"] += r.tiempo_minutos
        por_usuario[r.usuario]["registros"] += 1

        # Por día
        dia = str(r.fecha)
        if dia not in por_dia:
            por_dia[dia] = {"minutos": 0, "registros": 0}
        por_dia[dia]["minutos"] += r.tiempo_minutos
        por_dia[dia]["registros"] += 1

    total_minutos = sum(r.tiempo_minutos for r in registros)
    dias_unicos = len(por_dia)
    media_diaria = round(total_minutos / dias_unicos, 1) if dias_unicos > 0 else 0.0

    top_tareas = sorted(
        [{"tarea": t, **d} for t, d in por_tarea.items()],
        key=lambda x: x["minutos"],
        reverse=True
    )[:5]

    return {
        "total_registros": len(registros),
        "total_minutos": total_minutos,
        "total_horas": round(total_minutos / 60, 2),
        "por_tarea": por_tarea,
        "por_usuario": por_usuario,
        "por_dia": por_dia,
        "top_tareas": top_tareas,
        "media_diaria_minutos": media_diaria
    }


@app.get("/estadisticas/productividad/")
def productividad(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    db: Session = Depends(get_db)
):
    query = db.query(RegistroDB)
    if desde:
        query = query.filter(RegistroDB.fecha >= desde)
    if hasta:
        query = query.filter(RegistroDB.fecha <= hasta)
    registros = query.all()
    objetivos = {o.tarea_principal: o.unidades_hora for o in db.query(ObjetivoDB).filter(ObjetivoDB.activo == True).all()}
    volumenes_query = db.query(VolumenDB)
    if desde:
        volumenes_query = volumenes_query.filter(VolumenDB.fecha >= desde)
    if hasta:
        volumenes_query = volumenes_query.filter(VolumenDB.fecha <= hasta)
    volumenes = volumenes_query.all()

    por_usuario = {}
    for r in registros:
        if r.usuario not in por_usuario:
            por_usuario[r.usuario] = {"minutos": 0, "por_tarea": {}}
        por_usuario[r.usuario]["minutos"] += r.tiempo_minutos
        t = r.tarea_principal
        if t not in por_usuario[r.usuario]["por_tarea"]:
            por_usuario[r.usuario]["por_tarea"][t] = 0
        por_usuario[r.usuario]["por_tarea"][t] += r.tiempo_minutos

    por_tarea_global = {}
    for v in volumenes:
        t = v.tarea_principal
        if t not in por_tarea_global:
            por_tarea_global[t] = {"unidades": 0, "horas_teoricas": 0.0, "horas_reales": 0.0}
        por_tarea_global[t]["unidades"] += v.unidades
        por_tarea_global[t]["horas_teoricas"] += v.horas_teoricas

    for r in registros:
        t = r.tarea_principal
        if t in por_tarea_global:
            por_tarea_global[t]["horas_reales"] += r.tiempo_minutos / 60

    rendimiento_tareas = []
    for tarea, datos in por_tarea_global.items():
        eficiencia = 0.0
        if datos["horas_reales"] > 0:
            eficiencia = round((datos["horas_teoricas"] / datos["horas_reales"]) * 100, 1)
        objetivo_uds_hora = objetivos.get(tarea, None)
        uds_hora_real = 0.0
        if datos["horas_reales"] > 0:
            uds_hora_real = round(datos["unidades"] / datos["horas_reales"], 1)
        rendimiento_tareas.append({
            "tarea": tarea,
            "unidades": datos["unidades"],
            "horas_teoricas": round(datos["horas_teoricas"], 2),
            "horas_reales": round(datos["horas_reales"], 2),
            "eficiencia_pct": eficiencia,
            "objetivo_uds_hora": objetivo_uds_hora,
            "uds_hora_real": uds_hora_real,
            "cumple_objetivo": uds_hora_real >= objetivo_uds_hora if objetivo_uds_hora else None
        })

    ranking_usuarios = sorted(
        [{"usuario": u, "minutos": d["minutos"], "horas": round(d["minutos"] / 60, 2)} for u, d in por_usuario.items()],
        key=lambda x: x["minutos"],
        reverse=True
    )

    return {
        "desde": str(desde),
        "hasta": str(hasta),
        "rendimiento_por_tarea": rendimiento_tareas,
        "ranking_usuarios": ranking_usuarios,
        "total_usuarios_activos": len(por_usuario)
    }


@app.get("/estadisticas/rendimiento-operativo/")
def rendimiento_operativo(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    db: Session = Depends(get_db)
):
    hoy = date.today()
    if not desde:
        desde = hoy - timedelta(days=30)
    if not hasta:
        hasta = hoy

    registros = db.query(RegistroDB).filter(
        RegistroDB.fecha >= desde,
        RegistroDB.fecha <= hasta
    ).all()

    turnos = db.query(TurnoUsuarioDB).filter(
        TurnoUsuarioDB.fecha >= desde,
        TurnoUsuarioDB.fecha <= hasta
    ).all()

    incidencias = db.query(IncidenciaDB).filter(
        IncidenciaDB.fecha >= desde,
        IncidenciaDB.fecha <= hasta
    ).all()

    horas_previstas = sum(t.horas_previstas for t in turnos if t.estado == "trabaja")
    horas_registradas = sum(r.tiempo_minutos / 60 for r in registros)
    total_impacto_incidencias = sum(
        i.impacto_minutos for i in incidencias if i.impacto_minutos
    )

    cobertura_pct = 0.0
    if horas_previstas > 0:
        cobertura_pct = round((horas_registradas / horas_previstas) * 100, 1)

    usuarios_activos = list(set(r.usuario for r in registros))
    usuarios_sin_actividad = []
    todos_usuarios = db.query(UsuarioDB).filter(UsuarioDB.rol == "operario").all()
    for u in todos_usuarios:
        if u.codigo not in usuarios_activos:
            usuarios_sin_actividad.append(u.codigo)

    alertas = []
    if cobertura_pct < 70:
        alertas.append({"tipo": "warning", "mensaje": f"Cobertura de horas baja: {cobertura_pct}%"})
    if len(usuarios_sin_actividad) > 0:
        alertas.append({"tipo": "info", "mensaje": f"Usuarios sin actividad: {', '.join(usuarios_sin_actividad)}"})
    if total_impacto_incidencias > 120:
        alertas.append({"tipo": "danger", "mensaje": f"Alto impacto por incidencias: {total_impacto_incidencias} min"})

    return {
        "desde": str(desde),
        "hasta": str(hasta),
        "horas_previstas": round(horas_previstas, 2),
        "horas_registradas": round(horas_registradas, 2),
        "cobertura_pct": cobertura_pct,
        "total_incidencias": len(incidencias),
        "impacto_incidencias_min": total_impacto_incidencias,
        "usuarios_activos": len(usuarios_activos),
        "usuarios_sin_actividad": usuarios_sin_actividad,
        "alertas": alertas
    }


@app.get("/estadisticas/tendencia-semanal/")
def tendencia_semanal(
    semanas: int = 6,
    usuario: Optional[str] = None,
    db: Session = Depends(get_db)
):
    hoy = date.today()
    inicio_semana_actual = hoy - timedelta(days=hoy.weekday())
    resultado = []

    for i in range(semanas - 1, -1, -1):
        inicio = inicio_semana_actual - timedelta(weeks=i)
        fin = inicio + timedelta(days=6)
        query = db.query(RegistroDB).filter(
            RegistroDB.fecha >= inicio,
            RegistroDB.fecha <= fin
        )
        if usuario:
            query = query.filter(RegistroDB.usuario == usuario)
        registros = query.all()

        vol_query = db.query(VolumenDB).filter(
            VolumenDB.fecha >= inicio,
            VolumenDB.fecha <= fin
        )
        volumenes = vol_query.all()

        total_minutos = sum(r.tiempo_minutos for r in registros)
        total_unidades = sum(v.unidades for v in volumenes)
        horas_teoricas = sum(v.horas_teoricas for v in volumenes)
        horas_reales = total_minutos / 60
        eficiencia = round((horas_teoricas / horas_reales) * 100, 1) if horas_reales > 0 else 0.0

        resultado.append({
            "semana": f"{inicio.strftime('%d/%m')} - {fin.strftime('%d/%m')}",
            "inicio": str(inicio),
            "fin": str(fin),
            "total_minutos": total_minutos,
            "total_horas": round(horas_reales, 2),
            "total_unidades": total_unidades,
            "horas_teoricas": round(horas_teoricas, 2),
            "eficiencia_pct": eficiencia,
            "num_registros": len(registros)
        })

    return {"semanas": resultado}
# ==================== EXPORTAR EXCEL ====================

@app.get("/exportar-excel/")
def exportar_excel(
    desde: Optional[date] = None,
    hasta: Optional[date] = None,
    usuario: Optional[str] = None,
    db: Session = Depends(get_db)
):
    hoy = date.today()
    if not desde:
        desde = hoy.replace(day=1)
    if not hasta:
        hasta = hoy

    # ---- Colores ----
    COLOR_HEADER     = PatternFill("solid", fgColor="1F3864")
    COLOR_SUBTOTAL   = PatternFill("solid", fgColor="2E75B6")
    COLOR_ALERTA     = PatternFill("solid", fgColor="FF0000")
    COLOR_OK         = PatternFill("solid", fgColor="00B050")
    COLOR_WARNING    = PatternFill("solid", fgColor="FFC000")
    COLOR_FILA_PAR   = PatternFill("solid", fgColor="DCE6F1")
    COLOR_FILA_IMPAR = PatternFill("solid", fgColor="FFFFFF")

    FONT_HEADER  = Font(color="FFFFFF", bold=True, size=11)
    FONT_BOLD    = Font(bold=True)
    ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ALIGN_LEFT   = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def estilo_header(ws, fila, columnas):
        for col in range(1, columnas + 1):
            cell = ws.cell(row=fila, column=col)
            cell.fill = COLOR_HEADER
            cell.font = FONT_HEADER
            cell.alignment = ALIGN_CENTER

    def estilo_fila(ws, fila, columnas, par=True):
        fill = COLOR_FILA_PAR if par else COLOR_FILA_IMPAR
        for col in range(1, columnas + 1):
            cell = ws.cell(row=fila, column=col)
            cell.fill = fill
            cell.alignment = ALIGN_LEFT

    wb = openpyxl.Workbook()

    # ==================== HOJA 1: REGISTROS ====================
    ws1 = wb.active
    ws1.title = "Registros"
    ws1.append(["ID", "Usuario", "Fecha", "Tarea Principal", "Subtarea", "Tiempo (min)", "Horas", "Proyecto", "Comentarios", "Creado"])
    estilo_header(ws1, 1, 10)
    ws1.column_dimensions["A"].width = 6
    ws1.column_dimensions["B"].width = 14
    ws1.column_dimensions["C"].width = 12
    ws1.column_dimensions["D"].width = 18
    ws1.column_dimensions["E"].width = 22
    ws1.column_dimensions["F"].width = 12
    ws1.column_dimensions["G"].width = 8
    ws1.column_dimensions["H"].width = 16
    ws1.column_dimensions["I"].width = 25
    ws1.column_dimensions["J"].width = 18

    query = db.query(RegistroDB).filter(RegistroDB.fecha >= desde, RegistroDB.fecha <= hasta)
    if usuario:
        query = query.filter(RegistroDB.usuario.ilike(usuario))
    registros = query.order_by(RegistroDB.fecha.desc()).all()

    for idx, r in enumerate(registros, start=2):
        ws1.append([
            r.id, r.usuario, str(r.fecha), r.tarea_principal, r.subtarea,
            r.tiempo_minutos, round(r.tiempo_minutos / 60, 2),
            r.proyecto or "", r.comentarios or "",
            r.created_at.strftime("%d/%m/%Y %H:%M") if r.created_at else ""
        ])
        estilo_fila(ws1, idx, 10, par=(idx % 2 == 0))

    ws1.append([])
    total_row = ws1.max_row + 1
    ws1.cell(row=total_row, column=1, value="TOTAL")
    ws1.cell(row=total_row, column=6, value=sum(r.tiempo_minutos for r in registros))
    ws1.cell(row=total_row, column=7, value=round(sum(r.tiempo_minutos for r in registros) / 60, 2))
    for col in range(1, 11):
        ws1.cell(row=total_row, column=col).fill = COLOR_SUBTOTAL
        ws1.cell(row=total_row, column=col).font = FONT_HEADER
        ws1.cell(row=total_row, column=col).alignment = ALIGN_CENTER

    # ==================== HOJA 2: RESUMEN POR USUARIO ====================
    ws2 = wb.create_sheet("Resumen por Usuario")
    ws2.append(["Usuario", "Total Registros", "Total Minutos", "Total Horas", "Días Activos", "Media Min/Día"])
    estilo_header(ws2, 1, 6)
    ws2.column_dimensions["A"].width = 16
    ws2.column_dimensions["B"].width = 16
    ws2.column_dimensions["C"].width = 16
    ws2.column_dimensions["D"].width = 14
    ws2.column_dimensions["E"].width = 14
    ws2.column_dimensions["F"].width = 16

    por_usuario = {}
    for r in registros:
        if r.usuario not in por_usuario:
            por_usuario[r.usuario] = {"registros": 0, "minutos": 0, "dias": set()}
        por_usuario[r.usuario]["registros"] += 1
        por_usuario[r.usuario]["minutos"] += r.tiempo_minutos
        por_usuario[r.usuario]["dias"].add(str(r.fecha))

    for idx, (usr, datos) in enumerate(sorted(por_usuario.items()), start=2):
        dias = len(datos["dias"])
        media = round(datos["minutos"] / dias, 1) if dias > 0 else 0
        ws2.append([
            usr, datos["registros"], datos["minutos"],
            round(datos["minutos"] / 60, 2), dias, media
        ])
        estilo_fila(ws2, idx, 6, par=(idx % 2 == 0))

    # ==================== HOJA 3: RESUMEN POR TAREA ====================
    ws3 = wb.create_sheet("Resumen por Tarea")
    ws3.append(["Tarea Principal", "Subtareas Distintas", "Total Registros", "Total Minutos", "Total Horas", "% del Total"])
    estilo_header(ws3, 1, 6)
    ws3.column_dimensions["A"].width = 20
    ws3.column_dimensions["B"].width = 18
    ws3.column_dimensions["C"].width = 16
    ws3.column_dimensions["D"].width = 16
    ws3.column_dimensions["E"].width = 14
    ws3.column_dimensions["F"].width = 12

    por_tarea = {}
    for r in registros:
        if r.tarea_principal not in por_tarea:
            por_tarea[r.tarea_principal] = {"registros": 0, "minutos": 0, "subtareas": set()}
        por_tarea[r.tarea_principal]["registros"] += 1
        por_tarea[r.tarea_principal]["minutos"] += r.tiempo_minutos
        por_tarea[r.tarea_principal]["subtareas"].add(r.subtarea)

    total_min_global = sum(r.tiempo_minutos for r in registros)
    for idx, (tarea, datos) in enumerate(sorted(por_tarea.items(), key=lambda x: x[1]["minutos"], reverse=True), start=2):
        pct = round((datos["minutos"] / total_min_global) * 100, 1) if total_min_global > 0 else 0
        ws3.append([
            tarea, len(datos["subtareas"]), datos["registros"],
            datos["minutos"], round(datos["minutos"] / 60, 2), f"{pct}%"
        ])
        estilo_fila(ws3, idx, 6, par=(idx % 2 == 0))

    # ==================== HOJA 4: VOLUMEN & EFICIENCIA ====================
    ws4 = wb.create_sheet("Volumen y Eficiencia")
    ws4.append(["Fecha", "Tarea", "Unidades", "Horas Teóricas", "Horas Reales", "Eficiencia %", "Estado"])
    estilo_header(ws4, 1, 7)
    ws4.column_dimensions["A"].width = 12
    ws4.column_dimensions["B"].width = 18
    ws4.column_dimensions["C"].width = 12
    ws4.column_dimensions["D"].width = 16
    ws4.column_dimensions["E"].width = 14
    ws4.column_dimensions["F"].width = 14
    ws4.column_dimensions["G"].width = 14

    volumenes = db.query(VolumenDB).filter(
        VolumenDB.fecha >= desde,
        VolumenDB.fecha <= hasta
    ).order_by(VolumenDB.fecha.desc()).all()

    horas_reales_por_tarea = {}
    for r in registros:
        t = r.tarea_principal
        horas_reales_por_tarea[t] = horas_reales_por_tarea.get(t, 0) + r.tiempo_minutos / 60

    for idx, v in enumerate(volumenes, start=2):
        horas_reales = round(horas_reales_por_tarea.get(v.tarea_principal, 0), 2)
        eficiencia = round((v.horas_teoricas / horas_reales) * 100, 1) if horas_reales > 0 else 0.0
        estado = "✓ OK" if eficiencia >= 90 else ("⚠ Bajo" if eficiencia >= 70 else "✗ Crítico")
        ws4.append([
            str(v.fecha), v.tarea_principal, v.unidades,
            v.horas_teoricas, horas_reales, eficiencia, estado
        ])
        estilo_fila(ws4, idx, 7, par=(idx % 2 == 0))
        fill = COLOR_OK if eficiencia >= 90 else (COLOR_WARNING if eficiencia >= 70 else COLOR_ALERTA)
        ws4.cell(row=idx, column=7).fill = fill
        ws4.cell(row=idx, column=7).font = Font(bold=True, color="FFFFFF")
        ws4.cell(row=idx, column=7).alignment = ALIGN_CENTER

    # ==================== HOJA 5: INCIDENCIAS ====================
    ws5 = wb.create_sheet("Incidencias")
    ws5.append(["ID", "Fecha", "Tipo", "Tarea", "Impacto (min)", "Descripción", "Creado por"])
    estilo_header(ws5, 1, 7)
    ws5.column_dimensions["A"].width = 6
    ws5.column_dimensions["B"].width = 12
    ws5.column_dimensions["C"].width = 20
    ws5.column_dimensions["D"].width = 18
    ws5.column_dimensions["E"].width = 14
    ws5.column_dimensions["F"].width = 35
    ws5.column_dimensions["G"].width = 14

    incidencias = db.query(IncidenciaDB).filter(
        IncidenciaDB.fecha >= desde,
        IncidenciaDB.fecha <= hasta
    ).order_by(IncidenciaDB.fecha.desc()).all()

    for idx, i in enumerate(incidencias, start=2):
        ws5.append([
            i.id, str(i.fecha), i.tipo,
            i.tarea_principal or "", i.impacto_minutos or 0,
            i.descripcion, i.creado_por
        ])
        estilo_fila(ws5, idx, 7, par=(idx % 2 == 0))

    # ==================== HOJA 6: AUDITORÍA ====================
    ws6 = wb.create_sheet("Auditoría")
    ws6.append(["ID", "Usuario", "Acción", "Entidad", "Entidad ID", "Detalle", "Fecha/Hora"])
    estilo_header(ws6, 1, 7)
    ws6.column_dimensions["A"].width = 6
    ws6.column_dimensions["B"].width = 14
    ws6.column_dimensions["C"].width = 12
    ws6.column_dimensions["D"].width = 14
    ws6.column_dimensions["E"].width = 10
    ws6.column_dimensions["F"].width = 40
    ws6.column_dimensions["G"].width = 18

    auditoria = db.query(AuditoriaDB).filter(
        AuditoriaDB.created_at >= datetime.combine(desde, datetime.min.time()),
        AuditoriaDB.created_at <= datetime.combine(hasta, datetime.max.time())
    ).order_by(AuditoriaDB.created_at.desc()).limit(1000).all()

    for idx, a in enumerate(auditoria, start=2):
        ws6.append([
            a.id, a.usuario, a.accion, a.entidad,
            a.entidad_id or "", a.detalle or "",
            a.created_at.strftime("%d/%m/%Y %H:%M:%S") if a.created_at else ""
        ])
        estilo_fila(ws6, idx, 7, par=(idx % 2 == 0))
        accion_fill = {
            "CREAR": COLOR_OK,
            "EDITAR": COLOR_WARNING,
            "ELIMINAR": COLOR_ALERTA,
            "LOGIN": PatternFill("solid", fgColor="7030A0")
        }.get(a.accion, COLOR_FILA_PAR)
        ws6.cell(row=idx, column=3).fill = accion_fill
        ws6.cell(row=idx, column=3).font = Font(bold=True, color="FFFFFF")
        ws6.cell(row=idx, column=3).alignment = ALIGN_CENTER

    # ==================== GUARDAR Y DEVOLVER ====================
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    nombre_archivo = f"informe_{desde}_{hasta}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={nombre_archivo}"}
    )
