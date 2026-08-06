import os
import re
from datetime import datetime, date, timedelta
from typing import Optional, List
from contextlib import asynccontextmanager

from fastapi import FastAPI, Depends, HTTPException, Query
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse
from sqlalchemy import (
    create_engine, Column, Integer, String, Float, Date, DateTime,
    Text, Boolean, ForeignKey, func
)
from sqlalchemy.orm import declarative_base, sessionmaker, Session, relationship
from pydantic import BaseModel
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# ==================== BASE DE DATOS ====================
DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:///./registros.db")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False} if "sqlite" in DATABASE_URL else {}
)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# ==================== MODELOS ====================
class Usuario(Base):
    __tablename__ = "usuarios"
    id = Column(Integer, primary_key=True, index=True)
    codigo = Column(String, unique=True, index=True)
    nombre = Column(String)
    pin = Column(String)
    rol = Column(String, default="operario")
    activo = Column(Boolean, default=True)
    fecha_creacion = Column(DateTime, default=datetime.utcnow)
    registros = relationship("Registro", back_populates="usuario")
    turnos = relationship("TurnoUsuario", back_populates="usuario")

class Registro(Base):
    __tablename__ = "registros"
    id = Column(Integer, primary_key=True, index=True)
    usuario_id = Column(Integer, ForeignKey("usuarios.id"))
    fecha = Column(Date, default=date.today)
    tarea = Column(String)
    subtarea = Column(String, nullable=True)
    minutos = Column(Integer)
    notas = Column(Text, nullable=True)
    turno = Column(String, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    usuario = relationship("Usuario", back_populates="registros")

class Plantilla(Base):
    __tablename__ = "plantillas"
    id = Column(Integer, primary_key=True, index=True)
    usuario_id = Column(Integer, ForeignKey("usuarios.id"))
    nombre = Column(String)
    tarea = Column(String)
    subtarea = Column(String, nullable=True)
    minutos = Column(Integer)
    notas = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)

class Volumen(Base):
    __tablename__ = "volumenes"
    id = Column(Integer, primary_key=True, index=True)
    usuario_id = Column(Integer, ForeignKey("usuarios.id"))
    fecha = Column(Date, default=date.today)
    tarea = Column(String)
    unidades = Column(Integer, default=0)
    objetivo_unidades = Column(Integer, default=0)
    turno = Column(String, nullable=True)
    notas = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)

class Objetivo(Base):
    __tablename__ = "objetivos"
    id = Column(Integer, primary_key=True, index=True)
    tarea = Column(String)
    periodo = Column(String)
    metrica = Column(String)
    valor_objetivo = Column(Float)
    usuario_id = Column(Integer, ForeignKey("usuarios.id"), nullable=True)
    fecha_inicio = Column(Date, nullable=True)
    fecha_fin = Column(Date, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)

class Incidencia(Base):
    __tablename__ = "incidencias"
    id = Column(Integer, primary_key=True, index=True)
    usuario_id = Column(Integer, ForeignKey("usuarios.id"))
    fecha = Column(Date, default=date.today)
    tipo = Column(String)
    descripcion = Column(Text)
    impacto_minutos = Column(Integer, default=0)
    estado = Column(String, default="abierta")
    created_at = Column(DateTime, default=datetime.utcnow)

class TurnoUsuario(Base):
    __tablename__ = "turnos_usuarios"
    id = Column(Integer, primary_key=True, index=True)
    usuario_id = Column(Integer, ForeignKey("usuarios.id"))
    fecha = Column(Date, default=date.today)
    turno = Column(String)
    estado = Column(String, default="trabaja")
    notas = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    usuario = relationship("Usuario", back_populates="turnos")

class Auditoria(Base):
    __tablename__ = "auditoria"
    id = Column(Integer, primary_key=True, index=True)
    usuario_id = Column(Integer, ForeignKey("usuarios.id"), nullable=True)
    usuario_codigo = Column(String, nullable=True)
    accion = Column(String)
    entidad = Column(String, nullable=True)
    entidad_id = Column(Integer, nullable=True)
    detalle = Column(Text, nullable=True)
    timestamp = Column(DateTime, default=datetime.utcnow)

# ==================== INIT DB ====================
def init_db():
    Base.metadata.create_all(bind=engine)
    db = SessionLocal()
    try:
        admin = db.query(Usuario).filter(Usuario.codigo == "ADMIN").first()
        if not admin:
            db.add(Usuario(
                codigo="ADMIN",
                nombre="Administrador",
                pin="1234",
                rol="supervisor",
                activo=True
            ))
            db.commit()
    finally:
        db.close()

@asynccontextmanager
async def lifespan(app: FastAPI):
    init_db()
    yield

# ==================== APP ====================
app = FastAPI(lifespan=lifespan)

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def registrar_auditoria(
    db: Session,
    accion: str,
    entidad: str = None,
    entidad_id: int = None,
    detalle: str = None,
    usuario_id: int = None,
    usuario_codigo: str = None
):
    log = Auditoria(
        usuario_id=usuario_id,
        usuario_codigo=usuario_codigo,
        accion=accion,
        entidad=entidad,
        entidad_id=entidad_id,
        detalle=detalle
    )
    db.add(log)
    db.commit()
# ==================== SCHEMAS ====================
class UsuarioCreate(BaseModel):
    codigo: str
    nombre: str
    pin: str
    rol: str = "operario"

class UsuarioOut(BaseModel):
    id: int
    codigo: str
    nombre: str
    rol: str
    activo: bool
    class Config: from_attributes = True

class LoginRequest(BaseModel):
    codigo: str
    pin: str

class RegistroCreate(BaseModel):
    usuario_id: int
    fecha: str
    tarea: str
    subtarea: Optional[str] = None
    minutos: int
    notas: Optional[str] = None
    turno: Optional[str] = None

class RegistroOut(BaseModel):
    id: int
    usuario_id: int
    fecha: date
    tarea: str
    subtarea: Optional[str]
    minutos: int
    notas: Optional[str]
    turno: Optional[str]
    class Config: from_attributes = True

class PlantillaCreate(BaseModel):
    usuario_id: int
    nombre: str
    tarea: str
    subtarea: Optional[str] = None
    minutos: int
    notas: Optional[str] = None

class PlantillaOut(BaseModel):
    id: int
    usuario_id: int
    nombre: str
    tarea: str
    subtarea: Optional[str]
    minutos: int
    notas: Optional[str]
    class Config: from_attributes = True

class VolumenCreate(BaseModel):
    usuario_id: int
    fecha: str
    tarea: str
    unidades: int
    objetivo_unidades: int = 0
    turno: Optional[str] = None
    notas: Optional[str] = None

class VolumenOut(BaseModel):
    id: int
    usuario_id: int
    fecha: date
    tarea: str
    unidades: int
    objetivo_unidades: int
    turno: Optional[str]
    notas: Optional[str]
    class Config: from_attributes = True

class ObjetivoCreate(BaseModel):
    tarea: str
    periodo: str
    metrica: str
    valor_objetivo: float
    usuario_id: Optional[int] = None
    fecha_inicio: Optional[str] = None
    fecha_fin: Optional[str] = None

class ObjetivoOut(BaseModel):
    id: int
    tarea: str
    periodo: str
    metrica: str
    valor_objetivo: float
    usuario_id: Optional[int]
    fecha_inicio: Optional[date]
    fecha_fin: Optional[date]
    valor_real: Optional[float] = 0
    progreso: Optional[float] = 0
    class Config: from_attributes = True

class IncidenciaCreate(BaseModel):
    usuario_id: int
    fecha: str
    tipo: str
    descripcion: str
    impacto_minutos: int = 0
    estado: str = "abierta"

class IncidenciaOut(BaseModel):
    id: int
    usuario_id: int
    fecha: date
    tipo: str
    descripcion: str
    impacto_minutos: int
    estado: str
    class Config: from_attributes = True

class TurnoCreate(BaseModel):
    usuario_id: int
    fecha: str
    turno: str
    estado: str = "trabaja"
    notas: Optional[str] = None

class TurnoOut(BaseModel):
    id: int
    usuario_id: int
    fecha: date
    turno: str
    estado: str
    notas: Optional[str]
    usuario_codigo: Optional[str] = None
    class Config: from_attributes = True

class AuditoriaOut(BaseModel):
    id: int
    usuario_id: Optional[int]
    usuario_codigo: Optional[str]
    accion: str
    entidad: Optional[str]
    entidad_id: Optional[int]
    detalle: Optional[str]
    timestamp: datetime
    class Config: from_attributes = True

# ==================== ENDPOINTS: AUTH & USUARIOS ====================
@app.post("/login/")
def login(req: LoginRequest, db: Session = Depends(get_db)):
    user = db.query(Usuario).filter(
        Usuario.codigo == req.codigo.upper(),
        Usuario.pin == req.pin,
        Usuario.activo == True
    ).first()
    if not user:
        raise HTTPException(status_code=401, detail="Credenciales incorrectas")
    registrar_auditoria(
        db, "LOGIN", "usuarios", user.id,
        f"Login de {user.codigo}", user.id, user.codigo
    )
    return {"id": user.id, "codigo": user.codigo, "nombre": user.nombre, "rol": user.rol}

@app.get("/usuarios/", response_model=List[UsuarioOut])
def get_usuarios(db: Session = Depends(get_db)):
    return db.query(Usuario).filter(Usuario.activo == True).all()

@app.post("/usuarios/", response_model=UsuarioOut)
def crear_usuario(u: UsuarioCreate, db: Session = Depends(get_db)):
    existe = db.query(Usuario).filter(Usuario.codigo == u.codigo.upper()).first()
    if existe:
        raise HTTPException(status_code=400, detail="Código ya existe")
    nuevo = Usuario(
        codigo=u.codigo.upper(),
        nombre=u.nombre,
        pin=u.pin,
        rol=u.rol,
        activo=True
    )
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    registrar_auditoria(
        db, "CREAR", "usuarios", nuevo.id,
        f"Usuario {nuevo.codigo} creado"
    )
    return nuevo

@app.delete("/usuarios/{usuario_id}")
def eliminar_usuario(usuario_id: int, db: Session = Depends(get_db)):
    user = db.query(Usuario).filter(Usuario.id == usuario_id).first()
    if not user:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    user.activo = False
    db.commit()
    registrar_auditoria(
        db, "ELIMINAR", "usuarios", usuario_id,
        f"Usuario {user.codigo} desactivado"
    )
    return {"ok": True}

# ==================== ENDPOINTS: REGISTROS ====================
@app.get("/registros/", response_model=List[RegistroOut])
def get_registros(
    usuario_id: Optional[int] = None,
    fecha: Optional[str] = None,
    fecha_inicio: Optional[str] = None,
    fecha_fin: Optional[str] = None,
    tarea: Optional[str] = None,
    db: Session = Depends(get_db)
):
    q = db.query(Registro)
    if usuario_id:
        q = q.filter(Registro.usuario_id == usuario_id)
    if fecha:
        q = q.filter(Registro.fecha == date.fromisoformat(fecha))
    if fecha_inicio:
        q = q.filter(Registro.fecha >= date.fromisoformat(fecha_inicio))
    if fecha_fin:
        q = q.filter(Registro.fecha <= date.fromisoformat(fecha_fin))
    if tarea:
        q = q.filter(Registro.tarea == tarea)
    return q.order_by(Registro.fecha.desc(), Registro.created_at.desc()).all()

@app.post("/registros/", response_model=RegistroOut)
def crear_registro(r: RegistroCreate, db: Session = Depends(get_db)):
    nuevo = Registro(
        usuario_id=r.usuario_id,
        fecha=date.fromisoformat(r.fecha),
        tarea=r.tarea,
        subtarea=r.subtarea,
        minutos=r.minutos,
        notas=r.notas,
        turno=r.turno
    )
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    user = db.query(Usuario).filter(Usuario.id == r.usuario_id).first()
    registrar_auditoria(
        db, "CREAR", "registros", nuevo.id,
        f"{user.codigo if user else '?'} registró {r.tarea} {r.minutos}min",
        r.usuario_id, user.codigo if user else None
    )
    return nuevo

@app.put("/registros/{registro_id}", response_model=RegistroOut)
def editar_registro(registro_id: int, r: RegistroCreate, db: Session = Depends(get_db)):
    reg = db.query(Registro).filter(Registro.id == registro_id).first()
    if not reg:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    reg.fecha = date.fromisoformat(r.fecha)
    reg.tarea = r.tarea
    reg.subtarea = r.subtarea
    reg.minutos = r.minutos
    reg.notas = r.notas
    reg.turno = r.turno
    db.commit()
    db.refresh(reg)
    registrar_auditoria(
        db, "EDITAR", "registros", registro_id,
        f"Editado registro {registro_id}"
    )
    return reg

@app.delete("/registros/{registro_id}")
def eliminar_registro(registro_id: int, db: Session = Depends(get_db)):
    reg = db.query(Registro).filter(Registro.id == registro_id).first()
    if not reg:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    db.delete(reg)
    db.commit()
    registrar_auditoria(
        db, "ELIMINAR", "registros", registro_id,
        f"Eliminado registro {registro_id}"
    )
    return {"ok": True}

# ==================== ENDPOINTS: PLANTILLAS ====================
@app.get("/plantillas/", response_model=List[PlantillaOut])
def get_plantillas(usuario_id: Optional[int] = None, db: Session = Depends(get_db)):
    q = db.query(Plantilla)
    if usuario_id:
        q = q.filter(Plantilla.usuario_id == usuario_id)
    return q.order_by(Plantilla.nombre).all()

@app.post("/plantillas/", response_model=PlantillaOut)
def crear_plantilla(p: PlantillaCreate, db: Session = Depends(get_db)):
    nueva = Plantilla(**p.model_dump())
    db.add(nueva)
    db.commit()
    db.refresh(nueva)
    return nueva

@app.delete("/plantillas/{plantilla_id}")
def eliminar_plantilla(plantilla_id: int, db: Session = Depends(get_db)):
    p = db.query(Plantilla).filter(Plantilla.id == plantilla_id).first()
    if not p:
        raise HTTPException(status_code=404, detail="Plantilla no encontrada")
    db.delete(p)
    db.commit()
    return {"ok": True}
# ==================== ENDPOINTS: VOLÚMENES ====================
@app.get("/volumenes/", response_model=List[VolumenOut])
def get_volumenes(
    usuario_id: Optional[int] = None,
    fecha: Optional[str] = None,
    fecha_inicio: Optional[str] = None,
    fecha_fin: Optional[str] = None,
    db: Session = Depends(get_db)
):
    q = db.query(Volumen)
    if usuario_id:
        q = q.filter(Volumen.usuario_id == usuario_id)
    if fecha:
        q = q.filter(Volumen.fecha == date.fromisoformat(fecha))
    if fecha_inicio:
        q = q.filter(Volumen.fecha >= date.fromisoformat(fecha_inicio))
    if fecha_fin:
        q = q.filter(Volumen.fecha <= date.fromisoformat(fecha_fin))
    return q.order_by(Volumen.fecha.desc()).all()

@app.post("/volumenes/", response_model=VolumenOut)
def crear_volumen(v: VolumenCreate, db: Session = Depends(get_db)):
    nuevo = Volumen(
        usuario_id=v.usuario_id,
        fecha=date.fromisoformat(v.fecha),
        tarea=v.tarea,
        unidades=v.unidades,
        objetivo_unidades=v.objetivo_unidades,
        turno=v.turno,
        notas=v.notas
    )
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    user = db.query(Usuario).filter(Usuario.id == v.usuario_id).first()
    registrar_auditoria(
        db, "CREAR", "volumenes", nuevo.id,
        f"{user.codigo if user else '?'} registró {v.unidades} uds en {v.tarea}",
        v.usuario_id, user.codigo if user else None
    )
    return nuevo

@app.delete("/volumenes/{volumen_id}")
def eliminar_volumen(volumen_id: int, db: Session = Depends(get_db)):
    v = db.query(Volumen).filter(Volumen.id == volumen_id).first()
    if not v:
        raise HTTPException(status_code=404, detail="Volumen no encontrado")
    db.delete(v)
    db.commit()
    return {"ok": True}

# ==================== ENDPOINTS: INCIDENCIAS ====================
@app.get("/incidencias/", response_model=List[IncidenciaOut])
def get_incidencias(
    usuario_id: Optional[int] = None,
    fecha_inicio: Optional[str] = None,
    fecha_fin: Optional[str] = None,
    estado: Optional[str] = None,
    db: Session = Depends(get_db)
):
    q = db.query(Incidencia)
    if usuario_id:
        q = q.filter(Incidencia.usuario_id == usuario_id)
    if fecha_inicio:
        q = q.filter(Incidencia.fecha >= date.fromisoformat(fecha_inicio))
    if fecha_fin:
        q = q.filter(Incidencia.fecha <= date.fromisoformat(fecha_fin))
    if estado:
        q = q.filter(Incidencia.estado == estado)
    return q.order_by(Incidencia.fecha.desc()).all()

@app.post("/incidencias/", response_model=IncidenciaOut)
def crear_incidencia(inc: IncidenciaCreate, db: Session = Depends(get_db)):
    nueva = Incidencia(
        usuario_id=inc.usuario_id,
        fecha=date.fromisoformat(inc.fecha),
        tipo=inc.tipo,
        descripcion=inc.descripcion,
        impacto_minutos=inc.impacto_minutos,
        estado=inc.estado
    )
    db.add(nueva)
    db.commit()
    db.refresh(nueva)
    user = db.query(Usuario).filter(Usuario.id == inc.usuario_id).first()
    registrar_auditoria(
        db, "CREAR", "incidencias", nueva.id,
        f"{user.codigo if user else '?'} registró incidencia: {inc.tipo}",
        inc.usuario_id, user.codigo if user else None
    )
    return nueva

@app.put("/incidencias/{incidencia_id}")
def actualizar_incidencia(incidencia_id: int, estado: str, db: Session = Depends(get_db)):
    inc = db.query(Incidencia).filter(Incidencia.id == incidencia_id).first()
    if not inc:
        raise HTTPException(status_code=404, detail="Incidencia no encontrada")
    inc.estado = estado
    db.commit()
    registrar_auditoria(
        db, "EDITAR", "incidencias", incidencia_id,
        f"Incidencia {incidencia_id} → {estado}"
    )
    return {"ok": True}

@app.delete("/incidencias/{incidencia_id}")
def eliminar_incidencia(incidencia_id: int, db: Session = Depends(get_db)):
    inc = db.query(Incidencia).filter(Incidencia.id == incidencia_id).first()
    if not inc:
        raise HTTPException(status_code=404, detail="Incidencia no encontrada")
    db.delete(inc)
    db.commit()
    return {"ok": True}

# ==================== ENDPOINTS: TURNOS ====================
@app.get("/turnos/", response_model=List[TurnoOut])
def get_turnos(
    fecha: Optional[str] = None,
    fecha_inicio: Optional[str] = None,
    fecha_fin: Optional[str] = None,
    usuario_id: Optional[int] = None,
    db: Session = Depends(get_db)
):
    q = db.query(TurnoUsuario)
    if usuario_id:
        q = q.filter(TurnoUsuario.usuario_id == usuario_id)
    if fecha:
        q = q.filter(TurnoUsuario.fecha == date.fromisoformat(fecha))
    if fecha_inicio:
        q = q.filter(TurnoUsuario.fecha >= date.fromisoformat(fecha_inicio))
    if fecha_fin:
        q = q.filter(TurnoUsuario.fecha <= date.fromisoformat(fecha_fin))
    turnos = q.order_by(TurnoUsuario.fecha.desc()).all()
    result = []
    for t in turnos:
        user = db.query(Usuario).filter(Usuario.id == t.usuario_id).first()
        result.append(TurnoOut(
            id=t.id,
            usuario_id=t.usuario_id,
            fecha=t.fecha,
            turno=t.turno,
            estado=t.estado,
            notas=t.notas,
            usuario_codigo=user.codigo if user else None
        ))
    return result

@app.post("/turnos/", response_model=TurnoOut)
def crear_turno(t: TurnoCreate, db: Session = Depends(get_db)):
    nuevo = TurnoUsuario(
        usuario_id=t.usuario_id,
        fecha=date.fromisoformat(t.fecha),
        turno=t.turno,
        estado=t.estado,
        notas=t.notas
    )
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    user = db.query(Usuario).filter(Usuario.id == t.usuario_id).first()
    registrar_auditoria(
        db, "CREAR", "turnos", nuevo.id,
        f"Turno {t.turno} asignado a {user.codigo if user else '?'}",
        t.usuario_id, user.codigo if user else None
    )
    return TurnoOut(
        id=nuevo.id,
        usuario_id=nuevo.usuario_id,
        fecha=nuevo.fecha,
        turno=nuevo.turno,
        estado=nuevo.estado,
        notas=nuevo.notas,
        usuario_codigo=user.codigo if user else None
    )
# ==================== ENDPOINTS: OBJETIVOS ====================
@app.get("/objetivos/", response_model=List[ObjetivoOut])
def get_objetivos(db: Session = Depends(get_db)):
    objs = db.query(Objetivo).all()
    result = []
    for o in objs:
        result.append(ObjetivoOut(
            id=o.id,
            tarea=o.tarea,
            periodo=o.periodo,
            metrica=o.metrica,
            valor_objetivo=o.valor_objetivo,
            usuario_id=o.usuario_id,
            fecha_inicio=o.fecha_inicio,
            fecha_fin=o.fecha_fin,
            valor_real=0,
            progreso=0
        ))
    return result

@app.post("/objetivos/", response_model=ObjetivoOut)
def crear_objetivo(o: ObjetivoCreate, db: Session = Depends(get_db)):
    nuevo = Objetivo(
        tarea=o.tarea,
        periodo=o.periodo,
        metrica=o.metrica,
        valor_objetivo=o.valor_objetivo,
        usuario_id=o.usuario_id,
        fecha_inicio=date.fromisoformat(o.fecha_inicio) if o.fecha_inicio else None,
        fecha_fin=date.fromisoformat(o.fecha_fin) if o.fecha_fin else None
    )
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    registrar_auditoria(
        db, "CREAR", "objetivos", nuevo.id,
        f"Objetivo creado: {nuevo.tarea} → {nuevo.valor_objetivo} {nuevo.metrica}"
    )
    return ObjetivoOut(
        id=nuevo.id,
        tarea=nuevo.tarea,
        periodo=nuevo.periodo,
        metrica=nuevo.metrica,
        valor_objetivo=nuevo.valor_objetivo,
        usuario_id=nuevo.usuario_id,
        fecha_inicio=nuevo.fecha_inicio,
        fecha_fin=nuevo.fecha_fin,
        valor_real=0,
        progreso=0
    )

@app.delete("/objetivos/{objetivo_id}")
def eliminar_objetivo(objetivo_id: int, db: Session = Depends(get_db)):
    o = db.query(Objetivo).filter(Objetivo.id == objetivo_id).first()
    if not o:
        raise HTTPException(status_code=404, detail="Objetivo no encontrado")
    db.delete(o)
    db.commit()
    registrar_auditoria(
        db, "ELIMINAR", "objetivos", objetivo_id,
        f"Objetivo {objetivo_id} eliminado"
    )
    return {"ok": True}

# ==================== ENDPOINTS: ESTADÍSTICAS ====================
@app.get("/estadisticas/resumen/")
def estadisticas_resumen(
    usuario_id: Optional[int] = None,
    fecha_inicio: Optional[str] = None,
    fecha_fin: Optional[str] = None,
    db: Session = Depends(get_db)
):
    q = db.query(Registro)
    if usuario_id:
        q = q.filter(Registro.usuario_id == usuario_id)
    if fecha_inicio:
        q = q.filter(Registro.fecha >= date.fromisoformat(fecha_inicio))
    if fecha_fin:
        q = q.filter(Registro.fecha <= date.fromisoformat(fecha_fin))

    registros = q.all()
    total_minutos = sum(r.minutos for r in registros)

    por_tarea = {}
    for r in registros:
        por_tarea[r.tarea] = por_tarea.get(r.tarea, 0) + r.minutos

    por_usuario = {}
    for r in registros:
        user = db.query(Usuario).filter(Usuario.id == r.usuario_id).first()
        codigo = user.codigo if user else str(r.usuario_id)
        por_usuario[codigo] = por_usuario.get(codigo, 0) + r.minutos

    return {
        "total_registros": len(registros),
        "total_minutos": total_minutos,
        "total_horas": round(total_minutos / 60, 2),
        "por_tarea": por_tarea,
        "por_usuario": por_usuario
    }

@app.get("/estadisticas/productividad/")
def estadisticas_productividad(
    fecha_inicio: Optional[str] = None,
    fecha_fin: Optional[str] = None,
    db: Session = Depends(get_db)
):
    usuarios = db.query(Usuario).filter(Usuario.activo == True).all()
    resultado = []
    for u in usuarios:
        q = db.query(Registro).filter(Registro.usuario_id == u.id)
        if fecha_inicio:
            q = q.filter(Registro.fecha >= date.fromisoformat(fecha_inicio))
        if fecha_fin:
            q = q.filter(Registro.fecha <= date.fromisoformat(fecha_fin))
        regs = q.all()
        total_min = sum(r.minutos for r in regs)
        resultado.append({
            "usuario_id": u.id,
            "codigo": u.codigo,
            "nombre": u.nombre,
            "total_registros": len(regs),
            "total_minutos": total_min,
            "total_horas": round(total_min / 60, 2)
        })
    return resultado

# ==================== ENDPOINTS: AUDITORÍA ====================
@app.get("/auditoria/", response_model=List[AuditoriaOut])
def get_auditoria(
    usuario_id: Optional[int] = None,
    fecha_inicio: Optional[str] = None,
    fecha_fin: Optional[str] = None,
    db: Session = Depends(get_db)
):
    q = db.query(Auditoria)
    if usuario_id:
        q = q.filter(Auditoria.usuario_id == usuario_id)
    if fecha_inicio:
        q = q.filter(Auditoria.timestamp >= datetime.fromisoformat(fecha_inicio))
    if fecha_fin:
        q = q.filter(
            Auditoria.timestamp <= datetime.fromisoformat(fecha_fin + " 23:59:59")
        )
    return q.order_by(Auditoria.timestamp.desc()).limit(200).all()

# ==================== EXPORTACIÓN EXCEL ====================
@app.get("/exportar/excel/")
def exportar_excel(
    fecha_inicio: Optional[str] = None,
    fecha_fin: Optional[str] = None,
    db: Session = Depends(get_db)
):
    wb = openpyxl.Workbook()

    header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    center = Alignment(horizontal="center", vertical="center")
    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    def estilizar_hoja(ws, headers):
        ws.append(headers)
        for col_num, _ in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = thin
        ws.row_dimensions[1].height = 20

    def auto_ancho(ws):
        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value), default=10)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 40)

    # Hoja 1: Registros
    ws1 = wb.active
    ws1.title = "Registros"
    estilizar_hoja(ws1, ["ID", "Usuario", "Fecha", "Tarea", "Subtarea", "Minutos", "Horas", "Turno", "Notas"])
    q_reg = db.query(Registro)
    if fecha_inicio:
        q_reg = q_reg.filter(Registro.fecha >= date.fromisoformat(fecha_inicio))
    if fecha_fin:
        q_reg = q_reg.filter(Registro.fecha <= date.fromisoformat(fecha_fin))
    for r in q_reg.order_by(Registro.fecha.desc()).all():
        user = db.query(Usuario).filter(Usuario.id == r.usuario_id).first()
        ws1.append([
            r.id, user.codigo if user else "N/A", str(r.fecha),
            r.tarea, r.subtarea or "", r.minutos,
            round(r.minutos / 60, 2), r.turno or "", r.notas or ""
        ])
    auto_ancho(ws1)

    # Hoja 2: Volumen
    ws2 = wb.create_sheet(title="Volumen")
    estilizar_hoja(ws2, ["ID", "Usuario", "Fecha", "Tarea", "Unidades", "Objetivo", "Turno", "Notas"])
    q_vol = db.query(Volumen)
    if fecha_inicio:
        q_vol = q_vol.filter(Volumen.fecha >= date.fromisoformat(fecha_inicio))
    if fecha_fin:
        q_vol = q_vol.filter(Volumen.fecha <= date.fromisoformat(fecha_fin))
    for v in q_vol.order_by(Volumen.fecha.desc()).all():
        user = db.query(Usuario).filter(Usuario.id == v.usuario_id).first()
        ws2.append([
            v.id, user.codigo if user else "N/A", str(v.fecha),
            v.tarea, v.unidades, v.objetivo_unidades,
            v.turno or "", v.notas or ""
        ])
    auto_ancho(ws2)

    # Hoja 3: Incidencias
    ws3 = wb.create_sheet(title="Incidencias")
    estilizar_hoja(ws3, ["ID", "Usuario", "Fecha", "Tipo", "Descripción", "Impacto (min)", "Estado"])
    q_inc = db.query(Incidencia)
    if fecha_inicio:
        q_inc = q_inc.filter(Incidencia.fecha >= date.fromisoformat(fecha_inicio))
    if fecha_fin:
        q_inc = q_inc.filter(Incidencia.fecha <= date.fromisoformat(fecha_fin))
    for i in q_inc.order_by(Incidencia.fecha.desc()).all():
        user = db.query(Usuario).filter(Usuario.id == i.usuario_id).first()
        ws3.append([
            i.id, user.codigo if user else "N/A", str(i.fecha),
            i.tipo, i.descripcion, i.impacto_minutos, i.estado
        ])
    auto_ancho(ws3)

    # Hoja 4: Turnos
    ws4 = wb.create_sheet(title="Turnos")
    estilizar_hoja(ws4, ["ID", "Usuario", "Fecha", "Turno", "Estado", "Notas"])
    q_tur = db.query(TurnoUsuario)
    if fecha_inicio:
        q_tur = q_tur.filter(TurnoUsuario.fecha >= date.fromisoformat(fecha_inicio))
    if fecha_fin:
        q_tur = q_tur.filter(TurnoUsuario.fecha <= date.fromisoformat(fecha_fin))
    for t in q_tur.order_by(TurnoUsuario.fecha.desc()).all():
        user = db.query(Usuario).filter(Usuario.id == t.usuario_id).first()
        ws4.append([
            t.id, user.codigo if user else "N/A", str(t.fecha),
            t.turno, t.estado, t.notas or ""
        ])
    auto_ancho(ws4)

    # Hoja 5: Auditoría
    ws5 = wb.create_sheet(title="Auditoria")
    estilizar_hoja(ws5, ["ID", "Usuario", "Acción", "Entidad", "Entidad ID", "Detalle", "Timestamp"])
    for a in db.query(Auditoria).order_by(Auditoria.timestamp.desc()).limit(500).all():
        ws5.append([
            a.id, a.usuario_codigo or "sistema",
            a.accion, a.entidad or "", a.entidad_id or "",
            a.detalle or "", str(a.timestamp)
        ])
    auto_ancho(ws5)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=reporte_logistica.xlsx"}
    )

# ==================== FRONTEND ====================
@app.get("/")
def serve_index():
    if os.path.exists("index.html"):
        return FileResponse("index.html")
    return {"mensaje": "Backend activo. Sube index.html para el frontend."}

# ==================== INICIO ====================
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
