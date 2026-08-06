from datetime import date, timedelta
from typing import List, Optional
from fastapi import FastAPI, HTTPException, Depends
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from sqlalchemy import create_engine, Column, Integer, String, Date, DateTime, func
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import io
import os
from datetime import datetime

SQLALCHEMY_DATABASE_URL = "sqlite:///./registros.db"
engine = create_engine(SQLALCHEMY_DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

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

Base.metadata.create_all(bind=engine)

# Crear usuario administrador por defecto si no existe ninguno
with SessionLocal() as session:
    admin_existente = session.query(UsuarioDB).filter(UsuarioDB.rol == "supervisor").first()
    if not admin_existente:
        nuevo_admin = UsuarioDB(
            codigo="ADMIN",
            nombre="Administrador",
            pin="1234",
            rol="supervisor"
        )
        session.add(nuevo_admin)
        session.commit()

app = FastAPI(
    title="API de Registro de Tareas",
    version="7.0.0",
)

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

class PlantillaItemRespuesta(PlantillaItemCreate):
    id: int
    plantilla_id: int
    class Config:
        from_attributes = True

class PlantillaRespuesta(BaseModel):
    id: int
    nombre: str
    usuario: str
    items: List[PlantillaItemRespuesta] = []
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
        tarea = r.tarea_principal
        resumen[dia]["tareas"][tarea] = resumen[dia]["tareas"].get(tarea, 0) + r.tiempo_minutos
    return {
        "inicio": str(inicio_semana),
        "fin": str(fin_semana),
        "dias": resumen,
        "total_minutos": sum(r.tiempo_minutos for r in registros)
    }

@app.get("/registros/resumen-mensual/")
def resumen_mensual(usuario: Optional[str] = None, anyo: Optional[int] = None, mes: Optional[int] = None, db: Session = Depends(get_db)):
    hoy = date.today()
    if not anyo:
        anyo = hoy.year
    if not mes:
        mes = hoy.month
    inicio_mes = date(anyo, mes, 1)
    if mes == 12:
        fin_mes = date(anyo + 1, 1, 1) - timedelta(days=1)
    else:
        fin_mes = date(anyo, mes + 1, 1) - timedelta(days=1)
    query = db.query(RegistroDB).filter(
        RegistroDB.fecha >= inicio_mes,
        RegistroDB.fecha <= fin_mes
    )
    if usuario:
        query = query.filter(RegistroDB.usuario == usuario)
    registros = query.all()
    por_usuario = {}
    por_tarea = {}
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
    return db_registro

@app.delete("/registros/{id}", status_code=204)
def eliminar_registro(id: int, db: Session = Depends(get_db)):
    db_registro = db.query(RegistroDB).filter(RegistroDB.id == id).first()
    if not db_registro:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
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
    return nuevo
# ==================== PLANTILLAS ====================

@app.post("/plantillas/", status_code=201)
def crear_plantilla(plantilla: PlantillaCreate, db: Session = Depends(get_db)):
    nueva = PlantillaDB(
        nombre=plantilla.nombre,
        usuario=plantilla.usuario
    )
    db.add(nueva)
    db.commit()
    db.refresh(nueva)
    for item in plantilla.items:
        nuevo_item = PlantillaItemDB(
            plantilla_id=nueva.id,
            tarea_principal=item.tarea_principal,
            subtarea=item.subtarea,
            tiempo_minutos=item.tiempo_minutos,
            proyecto=item.proyecto,
            comentarios=item.comentarios
        )
        db.add(nuevo_item)
    db.commit()
    return {"id": nueva.id, "nombre": nueva.nombre, "mensaje": "Plantilla creada correctamente"}

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
            "created_at": str(p.created_at),
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
    db.query(PlantillaItemDB).filter(PlantillaItemDB.plantilla_id == id).delete()
    db.delete(plantilla)
    db.commit()

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

    colores_tarea = {
        "Picking":   "3b82f6",
        "Packing":   "8b5cf6",
        "Inbound":   "f59e0b",
        "Shipping":  "10b981",
        "Ecommerce": "ef4444",
    }

    wb = openpyxl.Workbook()

    # ---- HOJA 1: Registros detallados ----
    ws1 = wb.active
    ws1.title = "Registros Detallados"
    cabecera = ["ID", "Usuario", "Fecha", "Semana", "Mes", "Tarea Principal", "Subtarea", "Tiempo (min)", "Horas", "Proyecto", "Comentarios"]
    ws1.append(cabecera)
    for cell in ws1[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1e293b")
        cell.alignment = Alignment(horizontal="center")
    for r in registros:
        semana = r.fecha.isocalendar()[1]
        mes = r.fecha.strftime("%B %Y")
        color = colores_tarea.get(r.tarea_principal, "94a3b8")
        fila = [r.id, r.usuario, str(r.fecha), f"Semana {semana}", mes, r.tarea_principal, r.subtarea, r.tiempo_minutos, round(r.tiempo_minutos/60, 2), r.proyecto or "", r.comentarios or ""]
        ws1.append(fila)
        for cell in ws1[ws1.max_row]:
            cell.fill = PatternFill("solid", fgColor=color + "22" if len(color) == 6 else "f1f5f9")
    ws1.auto_filter.ref = ws1.dimensions
    for col in ws1.columns:
        max_length = max(len(str(cell.value or "")) for cell in col)
        ws1.column_dimensions[col[0].column_letter].width = min(max_length + 4, 40)
    ws1.freeze_panes = "A2"

    # ---- HOJA 2: Resumen por usuario ----
    ws2 = wb.create_sheet("Resumen por Usuario")
    ws2.append(["Usuario", "Total Registros", "Total Minutos", "Total Horas", "Media Minutos/Día"])
    for cell in ws2[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1e293b")
        cell.alignment = Alignment(horizontal="center")
    por_usuario = {}
    for r in registros:
        if r.usuario not in por_usuario:
            por_usuario[r.usuario] = {"registros": 0, "minutos": 0, "dias": set()}
        por_usuario[r.usuario]["registros"] += 1
        por_usuario[r.usuario]["minutos"] += r.tiempo_minutos
        por_usuario[r.usuario]["dias"].add(str(r.fecha))
    for u, datos in por_usuario.items():
        dias = len(datos["dias"]) or 1
        ws2.append([u, datos["registros"], datos["minutos"], round(datos["minutos"]/60, 2), round(datos["minutos"]/dias, 1)])
    for col in ws2.columns:
        max_length = max(len(str(cell.value or "")) for cell in col)
        ws2.column_dimensions[col[0].column_letter].width = min(max_length + 4, 40)

    # ---- HOJA 3: Resumen por tarea ----
    ws3 = wb.create_sheet("Resumen por Tarea")
    ws3.append(["Tarea Principal", "Subtarea", "Total Registros", "Total Minutos", "Total Horas", "Media Min/Registro"])
    for cell in ws3[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1e293b")
        cell.alignment = Alignment(horizontal="center")
    por_tarea = {}
    for r in registros:
        clave = (r.tarea_principal, r.subtarea)
        if clave not in por_tarea:
            por_tarea[clave] = {"registros": 0, "minutos": 0}
        por_tarea[clave]["registros"] += 1
        por_tarea[clave]["minutos"] += r.tiempo_minutos
    for (tp, st), datos in sorted(por_tarea.items()):
        media = round(datos["minutos"] / datos["registros"], 1) if datos["registros"] > 0 else 0
        ws3.append([tp, st, datos["registros"], datos["minutos"], round(datos["minutos"]/60, 2), media])
    for col in ws3.columns:
        max_length = max(len(str(cell.value or "")) for cell in col)
        ws3.column_dimensions[col[0].column_letter].width = min(max_length + 4, 40)

    # ---- HOJA 4: Resumen por día ----
    ws4 = wb.create_sheet("Resumen por Día")
    ws4.append(["Fecha", "Día Semana", "Total Registros", "Total Minutos", "Total Horas", "Usuarios Activos"])
    for cell in ws4[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1e293b")
        cell.alignment = Alignment(horizontal="center")
    por_dia = {}
    dias_semana = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]
    for r in registros:
        dia = str(r.fecha)
        if dia not in por_dia:
            por_dia[dia] = {"registros": 0, "minutos": 0, "usuarios": set()}
        por_dia[dia]["registros"] += 1
        por_dia[dia]["minutos"] += r.tiempo_minutos
        por_dia[dia]["usuarios"].add(r.usuario)
    for dia, datos in sorted(por_dia.items(), reverse=True):
        fecha_obj = date.fromisoformat(dia)
        nombre_dia = dias_semana[fecha_obj.weekday()]
        ws4.append([dia, nombre_dia, datos["registros"], datos["minutos"], round(datos["minutos"]/60, 2), len(datos["usuarios"])])
    for col in ws4.columns:
        max_length = max(len(str(cell.value or "")) for cell in col)
        ws4.column_dimensions[col[0].column_letter].width = min(max_length + 4, 40)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=registros.xlsx"}
    )

# ==================== ESTADÍSTICAS ====================

@app.get("/estadisticas/productividad/")
def productividad(fecha_ref: Optional[date] = None, db: Session = Depends(get_db)):
    if not fecha_ref:
        fecha_ref = date.today()
    ayer = fecha_ref - timedelta(days=1)
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
        variacion = total_hoy - total_ayer
        resultado.append({
            "usuario": u.codigo,
            "nombre": u.nombre,
            "minutos_hoy": total_hoy,
            "minutos_ayer": total_ayer,
            "variacion_minutos": variacion,
            "porcentaje_jornada": round((total_hoy / 480) * 100, 1)
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
    ordenado = sorted(por_tarea.items(), key=lambda x: x[1]["minutos"], reverse=True)
    return [
        {
            "tarea": k,
            "count": v["count"],
            "minutos_totales": v["minutos"],
            "media_minutos": round(v["minutos"] / v["count"], 1)
        }
        for k, v in ordenado[:10]
    ]

@app.get("/estadisticas/media-por-tarea/")
def media_por_tarea(usuario: Optional[str] = None, db: Session = Depends(get_db)):
    query = db.query(RegistroDB)
    if usuario:
        query = query.filter(RegistroDB.usuario == usuario)
    registros = query.all()
    por_tarea = {}
    for r in registros:
        clave = r.tarea_principal
        if clave not in por_tarea:
            por_tarea[clave] = {"count": 0, "minutos": 0}
        por_tarea[clave]["count"] += 1
        por_tarea[clave]["minutos"] += r.tiempo_minutos
    return [
        {
            "tarea": k,
            "count": v["count"],
            "media_minutos": round(v["minutos"] / v["count"], 1),
            "total_minutos": v["minutos"]
        }
        for k, v in por_tarea.items()
    ]

# ==================== ARRANQUE ====================

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port)
