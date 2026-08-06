from fastapi import FastAPI, Depends, HTTPException, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from sqlalchemy import create_engine, Column, Integer, String, Date, Float, Boolean, Text, text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
from pydantic import BaseModel
from typing import Optional, List
from datetime import date, datetime
from contextlib import asynccontextmanager
import openpyxl
import os
import io
from fastapi.responses import StreamingResponse

# ── Base de datos ──────────────────────────────────────────────────────────────
DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:///./registros.db")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# ── Modelos ────────────────────────────────────────────────────────────────────
class Usuario(Base):
    __tablename__ = "usuarios"
    id             = Column(Integer, primary_key=True, index=True)
    codigo         = Column(String, unique=True, index=True)
    nombre         = Column(String)
    pin            = Column(String)
    rol            = Column(String, default="operario")
    activo         = Column(Boolean, default=True)
    fecha_creacion = Column(Date, default=date.today)

class Registro(Base):
    __tablename__ = "registros"
    id           = Column(Integer, primary_key=True, index=True)
    usuario_id   = Column(Integer)
    codigo       = Column(String)
    nombre       = Column(String)
    tarea        = Column(String)
    subtarea     = Column(String, nullable=True)
    cantidad     = Column(Integer)
    fecha        = Column(Date)
    turno        = Column(String, nullable=True)
    notas        = Column(String, nullable=True)
    duplicado_de = Column(Integer, nullable=True)

class Plantilla(Base):
    __tablename__ = "plantillas"
    id          = Column(Integer, primary_key=True, index=True)
    nombre      = Column(String)
    tarea       = Column(String)
    subtarea    = Column(String, nullable=True)
    descripcion = Column(String, nullable=True)

class Volumen(Base):
    __tablename__ = "volumenes"
    id      = Column(Integer, primary_key=True, index=True)
    fecha   = Column(Date)
    tipo    = Column(String)
    valor   = Column(Float)
    notas   = Column(String, nullable=True)

class Objetivo(Base):
    __tablename__ = "objetivos"
    id      = Column(Integer, primary_key=True, index=True)
    tarea   = Column(String)
    meta    = Column(Integer)
    periodo = Column(String, default="diario")

class Incidencia(Base):
    __tablename__ = "incidencias"
    id          = Column(Integer, primary_key=True, index=True)
    fecha       = Column(Date)
    tipo        = Column(String)
    descripcion = Column(Text)
    usuario_id  = Column(Integer, nullable=True)
    resuelto    = Column(Boolean, default=False)

class TurnoUsuario(Base):
    __tablename__ = "turnos_usuarios"
    id         = Column(Integer, primary_key=True, index=True)
    usuario_id = Column(Integer)
    fecha      = Column(Date)
    turno      = Column(String)
    hora_ini   = Column(String, nullable=True)
    hora_fin   = Column(String, nullable=True)

class Auditoria(Base):
    __tablename__ = "auditoria"
    id         = Column(Integer, primary_key=True, index=True)
    fecha      = Column(Date)
    hora       = Column(String)
    usuario_id = Column(Integer, nullable=True)
    accion     = Column(String)
    detalle    = Column(Text, nullable=True)

# ── Migración segura ───────────────────────────────────────────────────────────
def migrar_columnas():
    es_postgres = not DATABASE_URL.startswith("sqlite")
    if not es_postgres:
        return
    columnas = [
        "ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS activo BOOLEAN DEFAULT TRUE",
        "ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS fecha_creacion DATE",
        "ALTER TABLE registros ADD COLUMN IF NOT EXISTS subtarea VARCHAR",
        "ALTER TABLE registros ADD COLUMN IF NOT EXISTS turno VARCHAR",
        "ALTER TABLE registros ADD COLUMN IF NOT EXISTS notas VARCHAR",
        "ALTER TABLE registros ADD COLUMN IF NOT EXISTS duplicado_de INTEGER",
        "ALTER TABLE volumenes ADD COLUMN IF NOT EXISTS notas VARCHAR",
        "ALTER TABLE incidencias ADD COLUMN IF NOT EXISTS resuelto BOOLEAN DEFAULT FALSE",
    ]
    with engine.connect() as conn:
        for sql in columnas:
            try:
                conn.execute(text(sql))
            except Exception:
                pass
        conn.commit()

# ── Init DB ────────────────────────────────────────────────────────────────────
def init_db():
    Base.metadata.create_all(bind=engine)
    migrar_columnas()
    db = SessionLocal()
    try:
        admin = db.query(Usuario).filter(Usuario.codigo == "ADMIN").first()
        if not admin:
            db.add(Usuario(
                codigo="ADMIN",
                nombre="Administrador",
                pin="1234",
                rol="supervisor",
                activo=True,
                fecha_creacion=date.today()
            ))
        else:
            admin.pin = "1234"
            admin.rol = "supervisor"
            admin.activo = True
        obj = db.query(Objetivo).first()
        if not obj:
            db.add(Objetivo(tarea="General", meta=100, periodo="diario"))
        db.commit()
    except Exception as e:
        db.rollback()
        print(f"Error en init_db: {e}")
    finally:
        db.close()

# ── Lifespan ───────────────────────────────────────────────────────────────────
@asynccontextmanager
async def lifespan(app: FastAPI):
    init_db()
    yield

app = FastAPI(lifespan=lifespan)

# ── Dependencia DB ─────────────────────────────────────────────────────────────
def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
# ── Schemas Pydantic ───────────────────────────────────────────────────────────
class UsuarioSchema(BaseModel):
    codigo: str
    nombre: str
    pin: str
    rol: str = "operario"
    activo: bool = True

class LoginSchema(BaseModel):
    codigo: str
    pin: str

class RegistroSchema(BaseModel):
    usuario_id: int
    codigo: str
    nombre: str
    tarea: str
    subtarea: Optional[str] = None
    cantidad: int
    fecha: str
    turno: Optional[str] = None
    notas: Optional[str] = None

class PlantillaSchema(BaseModel):
    nombre: str
    tarea: str
    subtarea: Optional[str] = None
    descripcion: Optional[str] = None

class VolumenSchema(BaseModel):
    fecha: str
    tipo: str
    valor: float
    notas: Optional[str] = None

class ObjetivoSchema(BaseModel):
    tarea: str
    meta: int
    periodo: str = "diario"

class IncidenciaSchema(BaseModel):
    fecha: str
    tipo: str
    descripcion: str
    usuario_id: Optional[int] = None
    resuelto: bool = False

class TurnoSchema(BaseModel):
    usuario_id: int
    fecha: str
    turno: str
    hora_ini: Optional[str] = None
    hora_fin: Optional[str] = None

class AuditoriaSchema(BaseModel):
    usuario_id: Optional[int] = None
    accion: str
    detalle: Optional[str] = None

# ── Endpoints: Auth & Usuarios ─────────────────────────────────────────────────
@app.post("/login")
def login(data: LoginSchema, db: Session = Depends(get_db)):
    user = db.query(Usuario).filter(
        Usuario.codigo == data.codigo,
        Usuario.pin == data.pin,
        Usuario.activo == True
    ).first()
    if not user:
        raise HTTPException(status_code=401, detail="Código o PIN incorrecto")
    return {
        "id": user.id,
        "codigo": user.codigo,
        "nombre": user.nombre,
        "rol": user.rol
    }

@app.get("/usuarios")
def get_usuarios(db: Session = Depends(get_db)):
    usuarios = db.query(Usuario).all()
    return [
        {
            "id": u.id,
            "codigo": u.codigo,
            "nombre": u.nombre,
            "rol": u.rol,
            "activo": u.activo,
            "fecha_creacion": str(u.fecha_creacion) if u.fecha_creacion else None
        }
        for u in usuarios
    ]

@app.post("/usuarios")
def crear_usuario(data: UsuarioSchema, db: Session = Depends(get_db)):
    existente = db.query(Usuario).filter(Usuario.codigo == data.codigo).first()
    if existente:
        raise HTTPException(status_code=400, detail="Código ya existe")
    u = Usuario(
        codigo=data.codigo,
        nombre=data.nombre,
        pin=data.pin,
        rol=data.rol,
        activo=data.activo,
        fecha_creacion=date.today()
    )
    db.add(u)
    db.commit()
    db.refresh(u)
    return {"id": u.id, "codigo": u.codigo, "nombre": u.nombre, "rol": u.rol}

@app.put("/usuarios/{usuario_id}")
def actualizar_usuario(usuario_id: int, data: UsuarioSchema, db: Session = Depends(get_db)):
    u = db.query(Usuario).filter(Usuario.id == usuario_id).first()
    if not u:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    u.codigo  = data.codigo
    u.nombre  = data.nombre
    u.pin     = data.pin
    u.rol     = data.rol
    u.activo  = data.activo
    db.commit()
    return {"ok": True}

@app.delete("/usuarios/{usuario_id}")
def eliminar_usuario(usuario_id: int, db: Session = Depends(get_db)):
    u = db.query(Usuario).filter(Usuario.id == usuario_id).first()
    if not u:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    db.delete(u)
    db.commit()
    return {"ok": True}

# ── Endpoints: Registros ───────────────────────────────────────────────────────
@app.get("/registros")
def get_registros(
    fecha: Optional[str] = None,
    usuario_id: Optional[int] = None,
    db: Session = Depends(get_db)
):
    q = db.query(Registro)
    if fecha:
        q = q.filter(Registro.fecha == date.fromisoformat(fecha))
    if usuario_id:
        q = q.filter(Registro.usuario_id == usuario_id)
    registros = q.order_by(Registro.id.desc()).all()
    return [
        {
            "id": r.id,
            "usuario_id": r.usuario_id,
            "codigo": r.codigo,
            "nombre": r.nombre,
            "tarea": r.tarea,
            "subtarea": r.subtarea,
            "cantidad": r.cantidad,
            "fecha": str(r.fecha),
            "turno": r.turno,
            "notas": r.notas,
            "duplicado_de": r.duplicado_de
        }
        for r in registros
    ]

@app.post("/registros")
def crear_registro(data: RegistroSchema, db: Session = Depends(get_db)):
    r = Registro(
        usuario_id=data.usuario_id,
        codigo=data.codigo,
        nombre=data.nombre,
        tarea=data.tarea,
        subtarea=data.subtarea,
        cantidad=data.cantidad,
        fecha=date.fromisoformat(data.fecha),
        turno=data.turno,
        notas=data.notas
    )
    db.add(r)
    db.commit()
    db.refresh(r)
    return {"id": r.id, "ok": True}

@app.put("/registros/{registro_id}")
def actualizar_registro(registro_id: int, data: RegistroSchema, db: Session = Depends(get_db)):
    r = db.query(Registro).filter(Registro.id == registro_id).first()
    if not r:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    r.tarea    = data.tarea
    r.subtarea = data.subtarea
    r.cantidad = data.cantidad
    r.turno    = data.turno
    r.notas    = data.notas
    db.commit()
    return {"ok": True}

@app.delete("/registros/{registro_id}")
def eliminar_registro(registro_id: int, db: Session = Depends(get_db)):
    r = db.query(Registro).filter(Registro.id == registro_id).first()
    if not r:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    db.delete(r)
    db.commit()
    return {"ok": True}

@app.post("/registros/{registro_id}/duplicar")
def duplicar_registro(registro_id: int, db: Session = Depends(get_db)):
    r = db.query(Registro).filter(Registro.id == registro_id).first()
    if not r:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    nuevo = Registro(
        usuario_id=r.usuario_id,
        codigo=r.codigo,
        nombre=r.nombre,
        tarea=r.tarea,
        subtarea=r.subtarea,
        cantidad=r.cantidad,
        fecha=date.today(),
        turno=r.turno,
        notas=r.notas,
        duplicado_de=r.id
    )
    db.add(nuevo)
    db.commit()
    db.refresh(nuevo)
    return {"id": nuevo.id, "ok": True}
# ── Endpoints: Plantillas ──────────────────────────────────────────────────────
@app.get("/plantillas")
def get_plantillas(db: Session = Depends(get_db)):
    plantillas = db.query(Plantilla).all()
    return [
        {
            "id": p.id,
            "nombre": p.nombre,
            "tarea": p.tarea,
            "subtarea": p.subtarea,
            "descripcion": p.descripcion
        }
        for p in plantillas
    ]

@app.post("/plantillas")
def crear_plantilla(data: PlantillaSchema, db: Session = Depends(get_db)):
    p = Plantilla(
        nombre=data.nombre,
        tarea=data.tarea,
        subtarea=data.subtarea,
        descripcion=data.descripcion
    )
    db.add(p)
    db.commit()
    db.refresh(p)
    return {"id": p.id, "ok": True}

@app.delete("/plantillas/{plantilla_id}")
def eliminar_plantilla(plantilla_id: int, db: Session = Depends(get_db)):
    p = db.query(Plantilla).filter(Plantilla.id == plantilla_id).first()
    if not p:
        raise HTTPException(status_code=404, detail="Plantilla no encontrada")
    db.delete(p)
    db.commit()
    return {"ok": True}

# ── Endpoints: Volúmenes ───────────────────────────────────────────────────────
@app.get("/volumenes")
def get_volumenes(fecha: Optional[str] = None, db: Session = Depends(get_db)):
    q = db.query(Volumen)
    if fecha:
        q = q.filter(Volumen.fecha == date.fromisoformat(fecha))
    volumenes = q.order_by(Volumen.id.desc()).all()
    return [
        {
            "id": v.id,
            "fecha": str(v.fecha),
            "tipo": v.tipo,
            "valor": v.valor,
            "notas": v.notas
        }
        for v in volumenes
    ]

@app.post("/volumenes")
def crear_volumen(data: VolumenSchema, db: Session = Depends(get_db)):
    v = Volumen(
        fecha=date.fromisoformat(data.fecha),
        tipo=data.tipo,
        valor=data.valor,
        notas=data.notas
    )
    db.add(v)
    db.commit()
    db.refresh(v)
    return {"id": v.id, "ok": True}

@app.delete("/volumenes/{volumen_id}")
def eliminar_volumen(volumen_id: int, db: Session = Depends(get_db)):
    v = db.query(Volumen).filter(Volumen.id == volumen_id).first()
    if not v:
        raise HTTPException(status_code=404, detail="Volumen no encontrado")
    db.delete(v)
    db.commit()
    return {"ok": True}

# ── Endpoints: Objetivos ───────────────────────────────────────────────────────
@app.get("/objetivos")
def get_objetivos(db: Session = Depends(get_db)):
    objetivos = db.query(Objetivo).all()
    return [
        {
            "id": o.id,
            "tarea": o.tarea,
            "meta": o.meta,
            "periodo": o.periodo
        }
        for o in objetivos
    ]

@app.post("/objetivos")
def crear_objetivo(data: ObjetivoSchema, db: Session = Depends(get_db)):
    o = Objetivo(tarea=data.tarea, meta=data.meta, periodo=data.periodo)
    db.add(o)
    db.commit()
    db.refresh(o)
    return {"id": o.id, "ok": True}

@app.put("/objetivos/{objetivo_id}")
def actualizar_objetivo(objetivo_id: int, data: ObjetivoSchema, db: Session = Depends(get_db)):
    o = db.query(Objetivo).filter(Objetivo.id == objetivo_id).first()
    if not o:
        raise HTTPException(status_code=404, detail="Objetivo no encontrado")
    o.tarea   = data.tarea
    o.meta    = data.meta
    o.periodo = data.periodo
    db.commit()
    return {"ok": True}

@app.delete("/objetivos/{objetivo_id}")
def eliminar_objetivo(objetivo_id: int, db: Session = Depends(get_db)):
    o = db.query(Objetivo).filter(Objetivo.id == objetivo_id).first()
    if not o:
        raise HTTPException(status_code=404, detail="Objetivo no encontrado")
    db.delete(o)
    db.commit()
    return {"ok": True}

# ── Endpoints: Incidencias ─────────────────────────────────────────────────────
@app.get("/incidencias")
def get_incidencias(db: Session = Depends(get_db)):
    incidencias = db.query(Incidencia).order_by(Incidencia.id.desc()).all()
    return [
        {
            "id": i.id,
            "fecha": str(i.fecha),
            "tipo": i.tipo,
            "descripcion": i.descripcion,
            "usuario_id": i.usuario_id,
            "resuelto": i.resuelto
        }
        for i in incidencias
    ]

@app.post("/incidencias")
def crear_incidencia(data: IncidenciaSchema, db: Session = Depends(get_db)):
    i = Incidencia(
        fecha=date.fromisoformat(data.fecha),
        tipo=data.tipo,
        descripcion=data.descripcion,
        usuario_id=data.usuario_id,
        resuelto=data.resuelto
    )
    db.add(i)
    db.commit()
    db.refresh(i)
    return {"id": i.id, "ok": True}

@app.put("/incidencias/{incidencia_id}/resolver")
def resolver_incidencia(incidencia_id: int, db: Session = Depends(get_db)):
    i = db.query(Incidencia).filter(Incidencia.id == incidencia_id).first()
    if not i:
        raise HTTPException(status_code=404, detail="Incidencia no encontrada")
    i.resuelto = True
    db.commit()
    return {"ok": True}

@app.delete("/incidencias/{incidencia_id}")
def eliminar_incidencia(incidencia_id: int, db: Session = Depends(get_db)):
    i = db.query(Incidencia).filter(Incidencia.id == incidencia_id).first()
    if not i:
        raise HTTPException(status_code=404, detail="Incidencia no encontrada")
    db.delete(i)
    db.commit()
    return {"ok": True}

# ── Endpoints: Turnos ──────────────────────────────────────────────────────────
@app.get("/turnos")
def get_turnos(fecha: Optional[str] = None, db: Session = Depends(get_db)):
    q = db.query(TurnoUsuario)
    if fecha:
        q = q.filter(TurnoUsuario.fecha == date.fromisoformat(fecha))
    turnos = q.order_by(TurnoUsuario.id.desc()).all()
    return [
        {
            "id": t.id,
            "usuario_id": t.usuario_id,
            "fecha": str(t.fecha),
            "turno": t.turno,
            "hora_ini": t.hora_ini,
            "hora_fin": t.hora_fin
        }
        for t in turnos
    ]

@app.post("/turnos")
def crear_turno(data: TurnoSchema, db: Session = Depends(get_db)):
    t = TurnoUsuario(
        usuario_id=data.usuario_id,
        fecha=date.fromisoformat(data.fecha),
        turno=data.turno,
        hora_ini=data.hora_ini,
        hora_fin=data.hora_fin
    )
    db.add(t)
    db.commit()
    db.refresh(t)
    return {"id": t.id, "ok": True}

@app.delete("/turnos/{turno_id}")
def eliminar_turno(turno_id: int, db: Session = Depends(get_db)):
    t = db.query(TurnoUsuario).filter(TurnoUsuario.id == turno_id).first()
    if not t:
        raise HTTPException(status_code=404, detail="Turno no encontrado")
    db.delete(t)
    db.commit()
    return {"ok": True}

# ── Endpoints: Auditoría ───────────────────────────────────────────────────────
@app.get("/auditoria")
def get_auditoria(db: Session = Depends(get_db)):
    logs = db.query(Auditoria).order_by(Auditoria.id.desc()).limit(200).all()
    return [
        {
            "id": a.id,
            "fecha": str(a.fecha),
            "hora": a.hora,
            "usuario_id": a.usuario_id,
            "accion": a.accion,
            "detalle": a.detalle
        }
        for a in logs
    ]

@app.post("/auditoria")
def registrar_auditoria(data: AuditoriaSchema, db: Session = Depends(get_db)):
    ahora = datetime.now()
    a = Auditoria(
        fecha=ahora.date(),
        hora=ahora.strftime("%H:%M:%S"),
        usuario_id=data.usuario_id,
        accion=data.accion,
        detalle=data.detalle
    )
    db.add(a)
    db.commit()
    return {"ok": True}
# ── Endpoints: Estadísticas ────────────────────────────────────────────────────
@app.get("/estadisticas")
def get_estadisticas(
    fecha: Optional[str] = None,
    usuario_id: Optional[int] = None,
    db: Session = Depends(get_db)
):
    q = db.query(Registro)
    if fecha:
        q = q.filter(Registro.fecha == date.fromisoformat(fecha))
    if usuario_id:
        q = q.filter(Registro.usuario_id == usuario_id)
    registros = q.all()

    total_tareas    = len(registros)
    total_cantidad  = sum(r.cantidad for r in registros)
    por_tarea       = {}
    por_usuario     = {}
    por_turno       = {}

    for r in registros:
        por_tarea[r.tarea]            = por_tarea.get(r.tarea, 0) + r.cantidad
        por_usuario[r.codigo]         = por_usuario.get(r.codigo, 0) + r.cantidad
        turno_key                     = r.turno or "Sin turno"
        por_turno[turno_key]          = por_turno.get(turno_key, 0) + r.cantidad

    objetivo = db.query(Objetivo).first()
    meta     = objetivo.meta if objetivo else 100
    progreso = round((total_cantidad / meta) * 100, 1) if meta > 0 else 0

    return {
        "total_tareas":   total_tareas,
        "total_cantidad": total_cantidad,
        "por_tarea":      por_tarea,
        "por_usuario":    por_usuario,
        "por_turno":      por_turno,
        "meta":           meta,
        "progreso":       progreso
    }

@app.get("/estadisticas/rango")
def get_estadisticas_rango(
    fecha_ini: str,
    fecha_fin: str,
    db: Session = Depends(get_db)
):
    registros = db.query(Registro).filter(
        Registro.fecha >= date.fromisoformat(fecha_ini),
        Registro.fecha <= date.fromisoformat(fecha_fin)
    ).all()

    por_dia     = {}
    por_tarea   = {}
    por_usuario = {}

    for r in registros:
        dia_key                    = str(r.fecha)
        por_dia[dia_key]           = por_dia.get(dia_key, 0) + r.cantidad
        por_tarea[r.tarea]         = por_tarea.get(r.tarea, 0) + r.cantidad
        por_usuario[r.codigo]      = por_usuario.get(r.codigo, 0) + r.cantidad

    return {
        "total_cantidad": sum(r.cantidad for r in registros),
        "por_dia":        por_dia,
        "por_tarea":      por_tarea,
        "por_usuario":    por_usuario
    }

# ── Exportar Excel ─────────────────────────────────────────────────────────────
@app.get("/exportar/excel")
def exportar_excel(
    fecha_ini: Optional[str] = None,
    fecha_fin: Optional[str] = None,
    db: Session = Depends(get_db)
):
    q = db.query(Registro)
    if fecha_ini:
        q = q.filter(Registro.fecha >= date.fromisoformat(fecha_ini))
    if fecha_fin:
        q = q.filter(Registro.fecha <= date.fromisoformat(fecha_fin))
    registros = q.order_by(Registro.fecha).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Registros"

    encabezados = [
        "ID", "Código", "Nombre", "Tarea", "Subtarea",
        "Cantidad", "Fecha", "Turno", "Notas"
    ]
    ws.append(encabezados)

    for r in registros:
        ws.append([
            r.id,
            r.codigo,
            r.nombre,
            r.tarea,
            r.subtarea or "",
            r.cantidad,
            str(r.fecha),
            r.turno or "",
            r.notas or ""
        ])

    # Hoja de resumen por tarea
    ws2 = wb.create_sheet("Resumen por Tarea")
    ws2.append(["Tarea", "Total Cantidad"])
    por_tarea = {}
    for r in registros:
        por_tarea[r.tarea] = por_tarea.get(r.tarea, 0) + r.cantidad
    for tarea, total in sorted(por_tarea.items()):
        ws2.append([tarea, total])

    # Hoja de resumen por usuario
    ws3 = wb.create_sheet("Resumen por Usuario")
    ws3.append(["Código", "Nombre", "Total Cantidad"])
    por_usuario = {}
    for r in registros:
        if r.codigo not in por_usuario:
            por_usuario[r.codigo] = {"nombre": r.nombre, "total": 0}
        por_usuario[r.codigo]["total"] += r.cantidad
    for codigo, datos in sorted(por_usuario.items()):
        ws3.append([codigo, datos["nombre"], datos["total"]])

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)

    nombre_archivo = f"registros_{date.today()}.xlsx"
    return StreamingResponse(
        stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={nombre_archivo}"}
    )

# ── Exportar Excel Incidencias ─────────────────────────────────────────────────
@app.get("/exportar/incidencias")
def exportar_incidencias(db: Session = Depends(get_db)):
    incidencias = db.query(Incidencia).order_by(Incidencia.fecha).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Incidencias"
    ws.append(["ID", "Fecha", "Tipo", "Descripción", "Usuario ID", "Resuelto"])

    for i in incidencias:
        ws.append([
            i.id,
            str(i.fecha),
            i.tipo,
            i.descripcion,
            i.usuario_id or "",
            "Sí" if i.resuelto else "No"
        ])

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)

    return StreamingResponse(
        stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=incidencias_{date.today()}.xlsx"}
    )

# ── Endpoint raíz ──────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
def root():
    ruta = os.path.join(os.path.dirname(__file__), "static", "index.html")
    if os.path.exists(ruta):
        with open(ruta, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse(content="<h1>App funcionando. Coloca tu index.html en /static/</h1>")

# ── Archivos estáticos ─────────────────────────────────────────────────────────
static_path = os.path.join(os.path.dirname(__file__), "static")
if os.path.exists(static_path):
    app.mount("/static", StaticFiles(directory=static_path), name="static")
