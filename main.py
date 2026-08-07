from datetime import date, timedelta
from typing import List, Optional
from fastapi import FastAPI, HTTPException, Depends
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from sqlalchemy import create_engine, Column, Integer, String, Date, DateTime, Float, Boolean
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import io
import os
from datetime import datetime

# ==================== BASE DE DATOS ====================
DATABASE_URL = os.environ.get("DATABASE_URL", "postgresql://localhost/registros")
# Render usa "postgres://" en lugar de "postgresql://", corregimos:
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

engine = create_engine(DATABASE_URL, pool_pre_ping=True)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()
# ==================== BASE DE DATOS ====================
DATABASE_URL = os.environ.get("DATABASE_URL", "postgresql://localhost/registros")
# Render usa "postgres://" en lugar de "postgresql://", corregimos:
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

engine = create_engine(DATABASE_URL, pool_pre_ping=True)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# --- SOLUCIÓN TEMPORAL PARA RECREAR LA TABLA OBJETIVOS ---
from sqlalchemy import text
try:
    with engine.connect() as connection:
        connection.execute(text("DROP TABLE IF EXISTS objetivos CASCADE;"))
        connection.commit()
    print("Tabla 'objetivos' eliminada correctamente para actualizar columnas.")
except Exception as e:
    print(f"Nota al limpiar tabla: {e}")
# ---------------------------------------------------------

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
    tarea_principal = Column(String, unique=True, index=True)
    uds_hora        = Column(Float)
    horas_jornada   = Column(Float, default=8.0)
    updated_at      = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

Base.metadata.create_all(bind=engine)

# Crear supervisor por defecto si no existe
with SessionLocal() as session:
    if not session.query(UsuarioDB).filter(UsuarioDB.rol == "supervisor").first():
        session.add(UsuarioDB(codigo="ADMIN", nombre="Administrador", pin="1234", rol="supervisor"))
        session.commit()
    # Objetivos por defecto
    tareas_default = ["Picking", "Packing", "Inbound", "Shipping", "Ecommerce"]
    for t in tareas_default:
        if not session.query(ObjetivoDB).filter(ObjetivoDB.tarea_principal == t).first():
            session.add(ObjetivoDB(tarea_principal=t, uds_hora=100.0, horas_jornada=8.0))
    session.commit()

# ==================== APP ====================
app = FastAPI(title="API Registro de Tareas", version="9.0.0")

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
    horas_teoricas: float
    creado_por: str
    comentarios: Optional[str] = None

class VolumenRespuesta(VolumenCreate):
    id: int
    created_at: Optional[datetime] = None
    updated_at: Optional[datetime] = None
    class Config:
        from_attributes = True

class ObjetivoCreate(BaseModel):
    tarea_principal: str
    uds_hora: float
    horas_jornada: float = 8.0

class ObjetivoRespuesta(ObjetivoCreate):
    id: int
    updated_at: Optional[datetime] = None
    class Config:
        from_attributes = True

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

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
    usuario = db.query(UsuarioDB).filter(
        UsuarioDB.codigo == datos.codigo,
        UsuarioDB.pin == datos.pin
    ).first()
    if not usuario:
        raise HTTPException(status_code=401, detail="Código o PIN incorrectos")
    return {"id": usuario.id, "codigo": usuario.codigo, "nombre": usuario.nombre, "rol": usuario.rol}

# ==================== USUARIOS ====================

@app.post("/usuarios/", response_model=UsuarioRespuesta, status_code=201)
def crear_usuario(usuario: UsuarioCreate, db: Session = Depends(get_db)):
    if db.query(UsuarioDB).filter(UsuarioDB.codigo == usuario.codigo).first():
        raise HTTPException(status_code=400, detail="Ya existe un usuario con ese código")
    nuevo = UsuarioDB(**usuario.dict())
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

# ==================== REGISTROS ====================

@app.post("/registros/", response_model=RegistroTareaRespuesta, status_code=201)
def crear_registro(registro: RegistroTarea, db: Session = Depends(get_db)):
    total = sum(r.tiempo_minutos for r in db.query(RegistroDB).filter(
        RegistroDB.usuario == registro.usuario,
        RegistroDB.fecha == registro.fecha
    ).with_entities(RegistroDB.tiempo_minutos).all())
    if total + registro.tiempo_minutos > 480:
        raise HTTPException(status_code=400, detail=f"No se pueden superar 8 horas. Minutos ya registrados: {total}")
    nuevo = RegistroDB(**registro.dict())
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    return nuevo

@app.get("/registros/", response_model=List[RegistroTareaRespuesta])
def obtener_registros(db: Session = Depends(get_db)):
    return db.query(RegistroDB).all()

@app.get("/registros/rango/", response_model=List[RegistroTareaRespuesta])
def obtener_registros_rango(desde: date, hasta: date, usuario: Optional[str] = None, db: Session = Depends(get_db)):
    query = db.query(RegistroDB).filter(RegistroDB.fecha >= desde, RegistroDB.fecha <= hasta)
    if usuario:
        query = query.filter(RegistroDB.usuario == usuario)
    return query.order_by(RegistroDB.fecha.desc()).all()

@app.get("/registros/resumen-semanal/")
def resumen_semanal(usuario: Optional[str] = None, fecha_ref: Optional[date] = None, db: Session = Depends(get_db)):
    if not fecha_ref:
        fecha_ref = date.today()
    inicio = fecha_ref - timedelta(days=fecha_ref.weekday())
    fin = inicio + timedelta(days=6)
    query = db.query(RegistroDB).filter(RegistroDB.fecha >= inicio, RegistroDB.fecha <= fin)
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
    return {"inicio": str(inicio), "fin": str(fin), "dias": resumen, "total_minutos": sum(r.tiempo_minutos for r in registros)}

@app.get("/registros/resumen-mensual/")
def resumen_mensual(usuario: Optional[str] = None, anyo: Optional[int] = None, mes: Optional[int] = None, db: Session = Depends(get_db)):
    hoy = date.today()
    anyo = anyo or hoy.year
    mes = mes or hoy.month
    inicio = date(anyo, mes, 1)
    fin = date(anyo + 1, 1, 1) - timedelta(days=1) if mes == 12 else date(anyo, mes + 1, 1) - timedelta(days=1)
    query = db.query(RegistroDB).filter(RegistroDB.fecha >= inicio, RegistroDB.fecha <= fin)
    if usuario:
        query = query.filter(RegistroDB.usuario == usuario)
    registros = query.all()
    por_usuario, por_tarea = {}, {}
    for r in registros:
        por_usuario[r.usuario] = por_usuario.get(r.usuario, 0) + r.tiempo_minutos
        por_tarea[r.tarea_principal] = por_tarea.get(r.tarea_principal, 0) + r.tiempo_minutos
    return {"anyo": anyo, "mes": mes, "total_minutos": sum(r.tiempo_minutos for r in registros), "total_registros": len(registros), "por_usuario": por_usuario, "por_tarea": por_tarea}

@app.get("/registros/{usuario}", response_model=List[RegistroTareaRespuesta])
def obtener_registros_por_usuario(usuario: str, db: Session = Depends(get_db)):
    resultado = db.query(RegistroDB).filter(RegistroDB.usuario.ilike(usuario)).all()
    if not resultado:
        raise HTTPException(status_code=404, detail="No se encontraron registros")
    return resultado

@app.put("/registros/{id}", response_model=RegistroTareaRespuesta)
def editar_registro(id: int, registro: RegistroTarea, db: Session = Depends(get_db)):
    db_reg = db.query(RegistroDB).filter(RegistroDB.id == id).first()
    if not db_reg:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    total = sum(r.tiempo_minutos for r in db.query(RegistroDB).filter(
        RegistroDB.usuario == registro.usuario,
        RegistroDB.fecha == registro.fecha,
        RegistroDB.id != id
    ).with_entities(RegistroDB.tiempo_minutos).all())
    if total + registro.tiempo_minutos > 480:
        raise HTTPException(status_code=400, detail=f"No se pueden superar 8 horas. Minutos ya registrados: {total}")
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
def duplicar_registro(id: int, nueva_fecha: Optional[date] = None, db: Session = Depends(get_db)):
    original = db.query(RegistroDB).filter(RegistroDB.id == id).first()
    if not original:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    fecha_destino = nueva_fecha or original.fecha
    total = sum(r.tiempo_minutos for r in db.query(RegistroDB).filter(
        RegistroDB.usuario == original.usuario,
        RegistroDB.fecha == fecha_destino
    ).with_entities(RegistroDB.tiempo_minutos).all())
    if total + original.tiempo_minutos > 480:
        raise HTTPException(status_code=400, detail=f"No se pueden superar 8 horas. Minutos ya registrados: {total}")
    nuevo = RegistroDB(
        usuario=original.usuario, fecha=fecha_destino,
        tarea_principal=original.tarea_principal, subtarea=original.subtarea,
        tiempo_minutos=original.tiempo_minutos, proyecto=original.proyecto,
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
            "id": p.id, "nombre": p.nombre, "usuario": p.usuario,
            "created_at": str(p.created_at),
            "items": [{"id": i.id, "tarea_principal": i.tarea_principal, "subtarea": i.subtarea,
                       "tiempo_minutos": i.tiempo_minutos, "proyecto": i.proyecto, "comentarios": i.comentarios}
                      for i in items]
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
def actualizar_objetivo(tarea: str, objetivo: ObjetivoCreate, db: Session = Depends(get_db)):
    db_obj = db.query(ObjetivoDB).filter(ObjetivoDB.tarea_principal == tarea).first()
    if not db_obj:
        db_obj = ObjetivoDB(tarea_principal=tarea, uds_hora=objetivo.uds_hora, horas_jornada=objetivo.horas_jornada)
        db.add(db_obj)
    else:
        db_obj.uds_hora = objetivo.uds_hora
        db_obj.horas_jornada = objetivo.horas_jornada
        db_obj.updated_at = datetime.utcnow()
    db.commit()
    db.refresh(db_obj)
    return db_obj

# ==================== VOLUMEN & MÉTRICAS ====================

@app.post("/volumenes/", response_model=VolumenRespuesta, status_code=201)
def crear_volumen(volumen: VolumenCreate, db: Session = Depends(get_db)):
    nuevo = VolumenDB(**volumen.dict())
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    return nuevo

@app.get("/volumenes/", response_model=List[VolumenRespuesta])
def obtener_volumenes(desde: Optional[date] = None, hasta: Optional[date] = None, tarea: Optional[str] = None, db: Session = Depends(get_db)):
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
def obtener_metricas(desde: Optional[date] = None, hasta: Optional[date] = None, db: Session = Depends(get_db)):
    if not desde:
        desde = date.today() - timedelta(days=30)
    if not hasta:
        hasta = date.today()
    volumenes = db.query(VolumenDB).filter(VolumenDB.fecha >= desde, VolumenDB.fecha <= hasta).all()
    registros = db.query(RegistroDB).filter(RegistroDB.fecha >= desde, RegistroDB.fecha <= hasta).all()
    resultado = []
    for v in volumenes:
        horas_reales = sum(r.tiempo_minutos for r in registros
            if str(r.fecha) == str(v.fecha) and r.tarea_principal == v.tarea_principal) / 60
        eficiencia = round((v.horas_teoricas / horas_reales) * 100, 1) if horas_reales > 0 else None
        uds_hora_real = round(v.unidades / horas_reales, 1) if horas_reales > 0 else None
        uds_hora_teorica = round(v.unidades / v.horas_teoricas, 1) if v.horas_teoricas > 0 else None
        resultado.append({
            "id": v.id, "fecha": str(v.fecha), "tarea_principal": v.tarea_principal,
            "unidades": v.unidades, "horas_teoricas": v.horas_teoricas,
            "horas_reales": round(horas_reales, 2), "eficiencia_pct": eficiencia,
            "uds_hora_real": uds_hora_real, "uds_hora_teorica": uds_hora_teorica,
            "desviacion_horas": round(horas_reales - v.horas_teoricas, 2),
            "comentarios": v.comentarios, "creado_por": v.creado_por
        })
    return sorted(resultado, key=lambda x: x["fecha"], reverse=True)

# ==================== ESTADÍSTICAS ====================

@app.get("/estadisticas/productividad/")
def productividad(fecha_ref: Optional[date] = None, db: Session = Depends(get_db)):
    if not fecha_ref:
        fecha_ref = date.today()
    ayer = fecha_ref - timedelta(days=1)
    usuarios = db.query(UsuarioDB).filter(UsuarioDB.rol == "operario").all()
    resultado = []
    for u in usuarios:
        hoy_mins = sum(r.tiempo_minutos for r in db.query(RegistroDB).filter(
            RegistroDB.usuario == u.codigo, RegistroDB.fecha == fecha_ref
        ).with_entities(RegistroDB.tiempo_minutos).all())
        ayer_mins = sum(r.tiempo_minutos for r in db.query(RegistroDB).filter(
            RegistroDB.usuario == u.codigo, RegistroDB.fecha == ayer
        ).with_entities(RegistroDB.tiempo_minutos).all())
        resultado.append({
            "usuario": u.codigo, "nombre": u.nombre,
            "minutos_hoy": hoy_mins, "minutos_ayer": ayer_mins,
            "variacion_minutos": hoy_mins - ayer_mins,
            "porcentaje_jornada": round((hoy_mins / 480) * 100, 1)
        })
    return resultado

@app.get("/estadisticas/top-tareas/")
def top_tareas(fecha_desde: Optional[date] = None, fecha_hasta: Optional[date] = None, db: Session = Depends(get_db)):
    if not fecha_desde:
        fecha_desde = date.today() - timedelta(days=30)
    if not fecha_hasta:
        fecha_hasta = date.today()
    registros = db.query(RegistroDB).filter(RegistroDB.fecha >= fecha_desde, RegistroDB.fecha <= fecha_hasta).all()
    por_tarea = {}
    for r in registros:
        clave = f"{r.tarea_principal} - {r.subtarea}"
        if clave not in por_tarea:
            por_tarea[clave] = {"count": 0, "minutos": 0}
        por_tarea[clave]["count"] += 1
        por_tarea[clave]["minutos"] += r.tiempo_minutos
    return [{"tarea": k, "count": v["count"], "minutos_totales": v["minutos"],
             "media_minutos": round(v["minutos"] / v["count"], 1)}
            for k, v in sorted(por_tarea.items(), key=lambda x: x[1]["minutos"], reverse=True)[:10]]

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
    return [{"tarea": k, "count": v["count"],
             "media_minutos": round(v["minutos"] / v["count"], 1),
             "total_minutos": v["minutos"]} for k, v in por_tarea.items()]

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
    usuarios = db.query(UsuarioDB).filter(UsuarioDB.rol == "operario").all()
    registros = db.query(RegistroDB).filter(RegistroDB.fecha >= desde, RegistroDB.fecha <= hasta).all()
    dias_rango = (hasta - desde).days + 1
    dias_laborables = sum(1 for i in range(dias_rango) if (desde + timedelta(days=i)).weekday() < 5)
    resultado = []
    for u in usuarios:
        regs_u = [r for r in registros if r.usuario == u.codigo]
        total_mins = sum(r.tiempo_minutos for r in regs_u)
        dias_activo = len(set(str(r.fecha) for r in regs_u))
        max_mins = dias_laborables * 480
        pct_jornada = round((total_mins / max_mins) * 100, 1) if max_mins > 0 else 0
        consistencia = round((dias_activo / dias_laborables) * 100, 1) if dias_laborables > 0 else 0
        media_diaria = round(total_mins / dias_activo, 0) if dias_activo > 0 else 0
        resultado.append({
            "usuario": u.codigo, "nombre": u.nombre,
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
    query = db.query(RegistroDB).filter(RegistroDB.fecha >= desde, RegistroDB.fecha <= hasta)
    if usuario:
        query = query.filter(RegistroDB.usuario == usuario)
    registros = query.all()
    objetivos = {o.tarea_principal: o for o in db.query(ObjetivoDB).all()}
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
        pct_jornada = round((datos["total_minutos"] / 480) * 100, 1)
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
    query = db.query(RegistroDB).filter(RegistroDB.fecha >= desde, RegistroDB.fecha <= hasta)
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
    ayer = hoy - timedelta(days=1)
    hace_7 = hoy - timedelta(days=7)
    usuarios = db.query(UsuarioDB).filter(UsuarioDB.rol == "operario").all()
    registros_hoy = db.query(RegistroDB).filter(RegistroDB.fecha == hoy).all()
    registros_semana = db.query(RegistroDB).filter(RegistroDB.fecha >= hace_7, RegistroDB.fecha <= hoy).all()
    alertas = []
    usuarios_con_actividad_hoy = set(r.usuario for r in registros_hoy)
    for u in usuarios:
        mins_hoy = sum(r.tiempo_minutos for r in registros_hoy if r.usuario == u.codigo)
        mins_semana = sum(r.tiempo_minutos for r in registros_semana if r.usuario == u.codigo)
        dias_activo_semana = len(set(str(r.fecha) for r in registros_semana if r.usuario == u.codigo))
        pct_hoy = round((mins_hoy / 480) * 100, 1)
        pct_semana = round((mins_semana / (5 * 480)) * 100, 1)
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
        if dias_activo_semana < 3 and dias_activo_semana > 0:
            alertas.append({
                "tipo": "baja_consistencia",
                "nivel": "danger",
                "usuario": u.codigo,
                "nombre": u.nombre,
                "mensaje": f"{u.nombre} solo ha registrado {dias_activo_semana} días esta semana"
            })
    volumenes = db.query(VolumenDB).filter(VolumenDB.fecha >= hace_7, VolumenDB.fecha <= hoy).all()
    for v in volumenes:
        horas_reales = sum(r.tiempo_minutos for r in registros_semana
            if str(r.fecha) == str(v.fecha) and r.tarea_principal == v.tarea_principal) / 60
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
                    "mensaje": f"Desviación >20% en {v.tarea_principal} el {v.fecha}: {round(horas_reales - v.horas_teoricas, 1)}h"
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
    registros = db.query(RegistroDB).filter(RegistroDB.fecha >= desde, RegistroDB.fecha <= hasta).all()
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
# ==================== EXCEL MEJORADO ====================

@app.get("/exportar-excel/")
def exportar_excel(
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

    # ---- HOJA 0: Resumen Ejecutivo ----
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

    # ---- HOJA 1: Registros detallados ----
    ws1 = wb.create_sheet("Registros Detallados")
    ws1.append(["ID", "Usuario", "Fecha", "Semana", "Mes", "Tarea Principal", "Subtarea",
                "Tiempo (min)", "Horas", "Proyecto", "Comentarios"])
    estilo_cabecera(ws1)
    dias_semana = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]
    for r in registros:
        semana = r.fecha.isocalendar()[1]
        mes = r.fecha.strftime("%B %Y")
        ws1.append([r.id, r.usuario, str(r.fecha), f"Semana {semana}", mes,
                    r.tarea_principal, r.subtarea, r.tiempo_minutos,
                    round(r.tiempo_minutos / 60, 2), r.proyecto or "", r.comentarios or ""])
    ws1.auto_filter.ref = ws1.dimensions
    ws1.freeze_panes = "A2"
    autoajustar(ws1)

    # ---- HOJA 2: Resumen por usuario ----
    ws2 = wb.create_sheet("Resumen por Usuario")
    ws2.append(["Usuario", "Total Registros", "Total Minutos", "Total Horas",
                "Días Activos", "Media Min/Día", "% Jornada Media"])
    estilo_cabecera(ws2)
    por_usuario = {}
    for r in registros:
        if r.usuario not in por_usuario:
            por_usuario[r.usuario] = {"registros": 0, "minutos": 0, "dias": set()}
        por_usuario[r.usuario]["registros"] += 1
        por_usuario[r.usuario]["minutos"] += r.tiempo_minutos
        por_usuario[r.usuario]["dias"].add(str(r.fecha))
    for u, d in por_usuario.items():
        dias = len(d["dias"]) or 1
        media_dia = round(d["minutos"] / dias, 1)
        pct = round((media_dia / 480) * 100, 1)
        ws2.append([u, d["registros"], d["minutos"], round(d["minutos"] / 60, 2),
                    len(d["dias"]), media_dia, pct])
    autoajustar(ws2)

    # ---- HOJA 3: Resumen por tarea ----
    ws3 = wb.create_sheet("Resumen por Tarea")
    ws3.append(["Tarea Principal", "Subtarea", "Total Registros",
                "Total Minutos", "Total Horas", "Media Min/Registro"])
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

    # ---- HOJA 4: Resumen por día ----
    ws4 = wb.create_sheet("Resumen por Día")
    ws4.append(["Fecha", "Día Semana", "Total Registros",
                "Total Minutos", "Total Horas", "Usuarios Activos", "% Jornada Media"])
    estilo_cabecera(ws4)
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
        pct = round((d["minutos"] / (n_usuarios * 480)) * 100, 1) if n_usuarios else 0
        ws4.append([dia, nombre_dia, d["registros"], d["minutos"],
                    round(d["minutos"] / 60, 2), n_usuarios, pct])
    autoajustar(ws4)

    # ---- HOJA 5: Métricas de Volumen ----
    ws5 = wb.create_sheet("Métricas Volumen")
    ws5.append(["Fecha", "Tarea", "Unidades", "Horas Teóricas", "Horas Reales",
                "Eficiencia %", "Uds/Hora Real", "Uds/Hora Teórica", "Desviación (h)", "Comentarios"])
    estilo_cabecera(ws5)
    volumenes = db.query(VolumenDB).order_by(VolumenDB.fecha.desc()).all()
    todos_registros = db.query(RegistroDB).all()
    for v in volumenes:
        horas_reales = sum(r.tiempo_minutos for r in todos_registros
            if str(r.fecha) == str(v.fecha) and r.tarea_principal == v.tarea_principal) / 60
        eficiencia = round((v.horas_teoricas / horas_reales) * 100, 1) if horas_reales > 0 else 0
        uds_real = round(v.unidades / horas_reales, 1) if horas_reales > 0 else 0
        uds_teo = round(v.unidades / v.horas_teoricas, 1) if v.horas_teoricas > 0 else 0
        desviacion = round(horas_reales - v.horas_teoricas, 2)
        ws5.append([str(v.fecha), v.tarea_principal, v.unidades, v.horas_teoricas,
                    round(horas_reales, 2), eficiencia, uds_real, uds_teo,
                    desviacion, v.comentarios or ""])
    autoajustar(ws5)

    # ---- HOJA 6: Ranking Rendimiento ----
    ws6 = wb.create_sheet("Ranking Rendimiento")
    ws6.append(["Posición", "Usuario", "Total Horas", "Días Activos",
                "% Jornada Periodo", "Consistencia %", "Media Min/Día"])
    estilo_cabecera(ws6)
    ranking = sorted(por_usuario.items(), key=lambda x: x[1]["minutos"], reverse=True)
    for pos, (u, d) in enumerate(ranking, 1):
        dias = len(d["dias"]) or 1
        media = round(d["minutos"] / dias, 1)
        pct_periodo = round((d["minutos"] / (dias * 480)) * 100, 1)
        ws6.append([pos, u, round(d["minutos"] / 60, 1), len(d["dias"]), pct_periodo,
                    round((dias / max(dias_unicos, 1)) * 100, 1), media])
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
