import os
import json
from datetime import datetime, date
from typing import Optional, List

from fastapi import FastAPI, Depends, HTTPException, status
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from sqlalchemy import (
    create_engine, Column, Integer, String, Float,
    Boolean, DateTime, Date, Text, ForeignKey, func
)
from sqlalchemy.orm import declarative_base, sessionmaker, Session, relationship
from pydantic import BaseModel

# ─── BASE DE DATOS ───────────────────────────────────────────────────────────
DATABASE_URL = os.environ.get(
    "DATABASE_URL",
    "postgresql://usuario:password@localhost:5432/warehouse"
)
# Render usa postgres://, SQLAlchemy necesita postgresql://
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# ─── MODELOS ─────────────────────────────────────────────────────────────────
class Usuario(Base):
    __tablename__ = "usuarios"
    id         = Column(Integer, primary_key=True, index=True)
    codigo     = Column(String(50), unique=True, nullable=False)
    nombre     = Column(String(100), nullable=False)
    pin        = Column(String(10), nullable=False)
    rol        = Column(String(20), default="operario")
    activo     = Column(Boolean, default=True)
    creado_en  = Column(DateTime, default=datetime.utcnow)

    registros  = relationship("Registro", back_populates="usuario")


class Plantilla(Base):
    __tablename__ = "plantillas"
    id          = Column(Integer, primary_key=True, index=True)
    nombre      = Column(String(100), nullable=False)
    descripcion = Column(Text, nullable=True)
    activa      = Column(Boolean, default=True)
    creado_en   = Column(DateTime, default=datetime.utcnow)

    registros   = relationship("Registro", back_populates="plantilla")


class Registro(Base):
    __tablename__ = "registros"
    id            = Column(Integer, primary_key=True, index=True)
    usuario_id    = Column(Integer, ForeignKey("usuarios.id"), nullable=False)
    plantilla_id  = Column(Integer, ForeignKey("plantillas.id"), nullable=True)
    tarea         = Column(String(200), nullable=False)
    cantidad      = Column(Float, default=0)
    unidad        = Column(String(50), default="unidades")
    turno         = Column(String(20), default="mañana")
    fecha         = Column(Date, default=date.today)
    notas         = Column(Text, nullable=True)
    creado_en     = Column(DateTime, default=datetime.utcnow)

    usuario       = relationship("Usuario", back_populates="registros")
    plantilla     = relationship("Plantilla", back_populates="registros")


class Volumen(Base):
    __tablename__ = "volumenes"
    id          = Column(Integer, primary_key=True, index=True)
    fecha       = Column(Date, default=date.today)
    turno       = Column(String(20), default="mañana")
    categoria   = Column(String(100), nullable=False)
    cantidad    = Column(Float, default=0)
    notas       = Column(Text, nullable=True)
    creado_en   = Column(DateTime, default=datetime.utcnow)


class Objetivo(Base):
    __tablename__ = "objetivos"
    id          = Column(Integer, primary_key=True, index=True)
    nombre      = Column(String(100), nullable=False)
    meta        = Column(Float, nullable=False)
    unidad      = Column(String(50), default="unidades")
    activo      = Column(Boolean, default=True)
    creado_en   = Column(DateTime, default=datetime.utcnow)


# ─── APP ─────────────────────────────────────────────────────────────────────
app = FastAPI(title="Warehouse Tracker")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

static_dir = os.path.join(os.path.dirname(__file__), "static")
if os.path.exists(static_dir):
    app.mount("/static", StaticFiles(directory=static_dir), name="static")


# ─── DEPENDENCIA DB ──────────────────────────────────────────────────────────
def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


# ─── INIT DB ─────────────────────────────────────────────────────────────────
def init_db():
    Base.metadata.create_all(bind=engine)
    db = SessionLocal()
    try:
        # Supervisor por defecto
        admin = db.query(Usuario).filter(Usuario.codigo == "ADMIN").first()
        if not admin:
            db.add(Usuario(
                codigo="ADMIN",
                nombre="Administrador",
                pin="1234",
                rol="supervisor",
                activo=True
            ))
        # Objetivo por defecto
        if db.query(Objetivo).count() == 0:
            db.add(Objetivo(
                nombre="Meta diaria general",
                meta=1000,
                unidad="unidades",
                activo=True
            ))
        db.commit()
    finally:
        db.close()


init_db()
# ─── SCHEMAS PYDANTIC ─────────────────────────────────────────────────────────

class LoginSchema(BaseModel):
    codigo: str
    pin: str

class UsuarioCreate(BaseModel):
    codigo: str
    nombre: str
    pin: str
    rol: str = "operario"

class UsuarioUpdate(BaseModel):
    nombre: Optional[str] = None
    pin: Optional[str] = None
    rol: Optional[str] = None
    activo: Optional[bool] = None

class RegistroCreate(BaseModel):
    usuario_id: int
    plantilla_id: Optional[int] = None
    tarea: str
    cantidad: float = 0
    unidad: str = "unidades"
    turno: str = "mañana"
    fecha: Optional[date] = None
    notas: Optional[str] = None

class RegistroUpdate(BaseModel):
    tarea: Optional[str] = None
    cantidad: Optional[float] = None
    unidad: Optional[str] = None
    turno: Optional[str] = None
    notas: Optional[str] = None

class PlantillaCreate(BaseModel):
    nombre: str
    descripcion: Optional[str] = None

class PlantillaUpdate(BaseModel):
    nombre: Optional[str] = None
    descripcion: Optional[str] = None
    activa: Optional[bool] = None

class VolumenCreate(BaseModel):
    fecha: Optional[date] = None
    turno: str = "mañana"
    categoria: str
    cantidad: float = 0
    notas: Optional[str] = None

class ObjetivoCreate(BaseModel):
    nombre: str
    meta: float
    unidad: str = "unidades"

class ObjetivoUpdate(BaseModel):
    nombre: Optional[str] = None
    meta: Optional[float] = None
    unidad: Optional[str] = None
    activo: Optional[bool] = None


# ─── AUTH ─────────────────────────────────────────────────────────────────────

@app.post("/api/login")
def login(data: LoginSchema, db: Session = Depends(get_db)):
    user = db.query(Usuario).filter(
        Usuario.codigo == data.codigo.upper(),
        Usuario.pin == data.pin,
        Usuario.activo == True
    ).first()
    if not user:
        raise HTTPException(status_code=401, detail="Credenciales incorrectas")
    return {
        "id": user.id,
        "codigo": user.codigo,
        "nombre": user.nombre,
        "rol": user.rol
    }


# ─── USUARIOS ─────────────────────────────────────────────────────────────────

@app.get("/api/usuarios")
def listar_usuarios(db: Session = Depends(get_db)):
    usuarios = db.query(Usuario).order_by(Usuario.nombre).all()
    return [
        {
            "id": u.id,
            "codigo": u.codigo,
            "nombre": u.nombre,
            "rol": u.rol,
            "activo": u.activo,
            "creado_en": u.creado_en
        }
        for u in usuarios
    ]

@app.post("/api/usuarios", status_code=201)
def crear_usuario(data: UsuarioCreate, db: Session = Depends(get_db)):
    existente = db.query(Usuario).filter(
        Usuario.codigo == data.codigo.upper()
    ).first()
    if existente:
        raise HTTPException(status_code=400, detail="El código ya existe")
    usuario = Usuario(
        codigo=data.codigo.upper(),
        nombre=data.nombre,
        pin=data.pin,
        rol=data.rol,
        activo=True
    )
    db.add(usuario)
    db.commit()
    db.refresh(usuario)
    return {"id": usuario.id, "codigo": usuario.codigo, "nombre": usuario.nombre}

@app.put("/api/usuarios/{usuario_id}")
def actualizar_usuario(
    usuario_id: int,
    data: UsuarioUpdate,
    db: Session = Depends(get_db)
):
    usuario = db.query(Usuario).filter(Usuario.id == usuario_id).first()
    if not usuario:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    if data.nombre is not None:
        usuario.nombre = data.nombre
    if data.pin is not None:
        usuario.pin = data.pin
    if data.rol is not None:
        usuario.rol = data.rol
    if data.activo is not None:
        usuario.activo = data.activo
    db.commit()
    return {"ok": True}

@app.delete("/api/usuarios/{usuario_id}")
def eliminar_usuario(usuario_id: int, db: Session = Depends(get_db)):
    usuario = db.query(Usuario).filter(Usuario.id == usuario_id).first()
    if not usuario:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    usuario.activo = False
    db.commit()
    return {"ok": True}
# ─── REGISTROS ───────────────────────────────────────────────────────────────

@app.get("/api/registros")
def listar_registros(
    usuario_id: Optional[int] = None,
    fecha: Optional[date] = None,
    turno: Optional[str] = None,
    db: Session = Depends(get_db)
):
    q = db.query(Registro)
    if usuario_id:
        q = q.filter(Registro.usuario_id == usuario_id)
    if fecha:
        q = q.filter(Registro.fecha == fecha)
    if turno:
        q = q.filter(Registro.turno == turno)
    registros = q.order_by(Registro.creado_en.desc()).all()
    return [
        {
            "id": r.id,
            "usuario_id": r.usuario_id,
            "usuario_nombre": r.usuario.nombre if r.usuario else None,
            "plantilla_id": r.plantilla_id,
            "plantilla_nombre": r.plantilla.nombre if r.plantilla else None,
            "tarea": r.tarea,
            "cantidad": r.cantidad,
            "unidad": r.unidad,
            "turno": r.turno,
            "fecha": r.fecha,
            "notas": r.notas,
            "creado_en": r.creado_en
        }
        for r in registros
    ]

@app.post("/api/registros", status_code=201)
def crear_registro(data: RegistroCreate, db: Session = Depends(get_db)):
    registro = Registro(
        usuario_id=data.usuario_id,
        plantilla_id=data.plantilla_id,
        tarea=data.tarea,
        cantidad=data.cantidad,
        unidad=data.unidad,
        turno=data.turno,
        fecha=data.fecha or date.today(),
        notas=data.notas
    )
    db.add(registro)
    db.commit()
    db.refresh(registro)
    return {"id": registro.id, "ok": True}

@app.put("/api/registros/{registro_id}")
def actualizar_registro(
    registro_id: int,
    data: RegistroUpdate,
    db: Session = Depends(get_db)
):
    registro = db.query(Registro).filter(Registro.id == registro_id).first()
    if not registro:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    if data.tarea is not None:
        registro.tarea = data.tarea
    if data.cantidad is not None:
        registro.cantidad = data.cantidad
    if data.unidad is not None:
        registro.unidad = data.unidad
    if data.turno is not None:
        registro.turno = data.turno
    if data.notas is not None:
        registro.notas = data.notas
    db.commit()
    return {"ok": True}

@app.delete("/api/registros/{registro_id}")
def eliminar_registro(registro_id: int, db: Session = Depends(get_db)):
    registro = db.query(Registro).filter(Registro.id == registro_id).first()
    if not registro:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    db.delete(registro)
    db.commit()
    return {"ok": True}


# ─── PLANTILLAS ───────────────────────────────────────────────────────────────

@app.get("/api/plantillas")
def listar_plantillas(db: Session = Depends(get_db)):
    plantillas = db.query(Plantilla).filter(Plantilla.activa == True).all()
    return [
        {
            "id": p.id,
            "nombre": p.nombre,
            "descripcion": p.descripcion,
            "activa": p.activa,
            "creado_en": p.creado_en
        }
        for p in plantillas
    ]

@app.post("/api/plantillas", status_code=201)
def crear_plantilla(data: PlantillaCreate, db: Session = Depends(get_db)):
    plantilla = Plantilla(
        nombre=data.nombre,
        descripcion=data.descripcion,
        activa=True
    )
    db.add(plantilla)
    db.commit()
    db.refresh(plantilla)
    return {"id": plantilla.id, "ok": True}

@app.put("/api/plantillas/{plantilla_id}")
def actualizar_plantilla(
    plantilla_id: int,
    data: PlantillaUpdate,
    db: Session = Depends(get_db)
):
    plantilla = db.query(Plantilla).filter(Plantilla.id == plantilla_id).first()
    if not plantilla:
        raise HTTPException(status_code=404, detail="Plantilla no encontrada")
    if data.nombre is not None:
        plantilla.nombre = data.nombre
    if data.descripcion is not None:
        plantilla.descripcion = data.descripcion
    if data.activa is not None:
        plantilla.activa = data.activa
    db.commit()
    return {"ok": True}

@app.delete("/api/plantillas/{plantilla_id}")
def eliminar_plantilla(plantilla_id: int, db: Session = Depends(get_db)):
    plantilla = db.query(Plantilla).filter(Plantilla.id == plantilla_id).first()
    if not plantilla:
        raise HTTPException(status_code=404, detail="Plantilla no encontrada")
    plantilla.activa = False
    db.commit()
    return {"ok": True}


# ─── VOLÚMENES ────────────────────────────────────────────────────────────────

@app.get("/api/volumenes")
def listar_volumenes(
    fecha: Optional[date] = None,
    turno: Optional[str] = None,
    db: Session = Depends(get_db)
):
    q = db.query(Volumen)
    if fecha:
        q = q.filter(Volumen.fecha == fecha)
    if turno:
        q = q.filter(Volumen.turno == turno)
    volumenes = q.order_by(Volumen.fecha.desc()).all()
    return [
        {
            "id": v.id,
            "fecha": v.fecha,
            "turno": v.turno,
            "categoria": v.categoria,
            "cantidad": v.cantidad,
            "notas": v.notas,
            "creado_en": v.creado_en
        }
        for v in volumenes
    ]

@app.post("/api/volumenes", status_code=201)
def crear_volumen(data: VolumenCreate, db: Session = Depends(get_db)):
    volumen = Volumen(
        fecha=data.fecha or date.today(),
        turno=data.turno,
        categoria=data.categoria,
        cantidad=data.cantidad,
        notas=data.notas
    )
    db.add(volumen)
    db.commit()
    db.refresh(volumen)
    return {"id": volumen.id, "ok": True}

@app.delete("/api/volumenes/{volumen_id}")
def eliminar_volumen(volumen_id: int, db: Session = Depends(get_db)):
    volumen = db.query(Volumen).filter(Volumen.id == volumen_id).first()
    if not volumen:
        raise HTTPException(status_code=404, detail="Volumen no encontrado")
    db.delete(volumen)
    db.commit()
    return {"ok": True}


# ─── OBJETIVOS ────────────────────────────────────────────────────────────────

@app.get("/api/objetivos")
def listar_objetivos(db: Session = Depends(get_db)):
    objetivos = db.query(Objetivo).filter(Objetivo.activo == True).all()
    return [
        {
            "id": o.id,
            "nombre": o.nombre,
            "meta": o.meta,
            "unidad": o.unidad,
            "activo": o.activo,
            "creado_en": o.creado_en
        }
        for o in objetivos
    ]

@app.post("/api/objetivos", status_code=201)
def crear_objetivo(data: ObjetivoCreate, db: Session = Depends(get_db)):
    objetivo = Objetivo(
        nombre=data.nombre,
        meta=data.meta,
        unidad=data.unidad,
        activo=True
    )
    db.add(objetivo)
    db.commit()
    db.refresh(objetivo)
    return {"id": objetivo.id, "ok": True}

@app.put("/api/objetivos/{objetivo_id}")
def actualizar_objetivo(
    objetivo_id: int,
    data: ObjetivoUpdate,
    db: Session = Depends(get_db)
):
    objetivo = db.query(Objetivo).filter(Objetivo.id == objetivo_id).first()
    if not objetivo:
        raise HTTPException(status_code=404, detail="Objetivo no encontrado")
    if data.nombre is not None:
        objetivo.nombre = data.nombre
    if data.meta is not None:
        objetivo.meta = data.meta
    if data.unidad is not None:
        objetivo.unidad = data.unidad
    if data.activo is not None:
        objetivo.activo = data.activo
    db.commit()
    return {"ok": True}

@app.delete("/api/objetivos/{objetivo_id}")
def eliminar_objetivo(objetivo_id: int, db: Session = Depends(get_db)):
    objetivo = db.query(Objetivo).filter(Objetivo.id == objetivo_id).first()
    if not objetivo:
        raise HTTPException(status_code=404, detail="Objetivo no encontrado")
    objetivo.activo = False
    db.commit()
    return {"ok": True}
# ─── ESTADÍSTICAS ─────────────────────────────────────────────────────────────

@app.get("/api/estadisticas/resumen")
def resumen_estadisticas(
    fecha: Optional[date] = None,
    db: Session = Depends(get_db)
):
    fecha_ref = fecha or date.today()

    total_registros_hoy = db.query(func.count(Registro.id)).filter(
        Registro.fecha == fecha_ref
    ).scalar() or 0

    total_cantidad_hoy = db.query(func.sum(Registro.cantidad)).filter(
        Registro.fecha == fecha_ref
    ).scalar() or 0.0

    objetivo_activo = db.query(Objetivo).filter(Objetivo.activo == True).first()
    meta = objetivo_activo.meta if objetivo_activo else 0
    progreso_pct = round((total_cantidad_hoy / meta * 100), 1) if meta > 0 else 0

    total_usuarios_activos = db.query(func.count(Usuario.id)).filter(
        Usuario.activo == True
    ).scalar() or 0

    return {
        "fecha": fecha_ref,
        "total_registros_hoy": total_registros_hoy,
        "total_cantidad_hoy": total_cantidad_hoy,
        "meta": meta,
        "progreso_pct": progreso_pct,
        "total_usuarios_activos": total_usuarios_activos
    }


@app.get("/api/estadisticas/por_turno")
def estadisticas_por_turno(
    fecha: Optional[date] = None,
    db: Session = Depends(get_db)
):
    fecha_ref = fecha or date.today()
    resultados = (
        db.query(Registro.turno, func.sum(Registro.cantidad))
        .filter(Registro.fecha == fecha_ref)
        .group_by(Registro.turno)
        .all()
    )
    return [{"turno": r[0], "cantidad": r[1] or 0} for r in resultados]


@app.get("/api/estadisticas/por_usuario")
def estadisticas_por_usuario(
    fecha: Optional[date] = None,
    db: Session = Depends(get_db)
):
    fecha_ref = fecha or date.today()
    resultados = (
        db.query(Usuario.nombre, func.sum(Registro.cantidad))
        .join(Registro, Registro.usuario_id == Usuario.id)
        .filter(Registro.fecha == fecha_ref)
        .group_by(Usuario.nombre)
        .order_by(func.sum(Registro.cantidad).desc())
        .all()
    )
    return [{"usuario": r[0], "cantidad": r[1] or 0} for r in resultados]


@app.get("/api/estadisticas/historico")
def estadisticas_historico(
    dias: int = 7,
    db: Session = Depends(get_db)
):
    from datetime import timedelta
    hoy = date.today()
    resultados = (
        db.query(Registro.fecha, func.sum(Registro.cantidad))
        .filter(Registro.fecha >= hoy - timedelta(days=dias))
        .group_by(Registro.fecha)
        .order_by(Registro.fecha)
        .all()
    )
    return [{"fecha": str(r[0]), "cantidad": r[1] or 0} for r in resultados]


# ─── RUTA RAÍZ ────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
def root():
    ruta = os.path.join(os.path.dirname(__file__), "static", "index.html")
    if os.path.exists(ruta):
        with open(ruta, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse(
        content="<h1>Warehouse Tracker activo.</h1><p>Coloca tu index.html en /static/</p>"
    )
