from datetime import date
from typing import List, Optional
from fastapi import FastAPI, HTTPException, Depends
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from sqlalchemy import create_engine, Column, Integer, String, Date
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
import openpyxl
import io
import os

SQLALCHEMY_DATABASE_URL = "sqlite:///./registros.db"
engine = create_engine(SQLALCHEMY_DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

class RegistroDB(Base):
    __tablename__ = "registros"
    id             = Column(Integer, primary_key=True, index=True)
    usuario        = Column(String, index=True)
    fecha          = Column(Date)
    tarea          = Column(String)
    tiempo_minutos = Column(Integer)
    proyecto       = Column(String, nullable=True)
    comentarios    = Column(String, nullable=True)

class UsuarioDB(Base):
    __tablename__ = "usuarios"
    id             = Column(Integer, primary_key=True, index=True)
    codigo         = Column(String, unique=True, index=True)
    nombre         = Column(String)
    pin            = Column(String)
    rol            = Column(String, default="operario")  # operario o supervisor

Base.metadata.create_all(bind=engine)

app = FastAPI(
    title="API de Registro de Tareas",
    version="6.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class RegistroTarea(BaseModel):
    usuario: str
    fecha: date
    tarea: str
    tiempo_minutos: int
    proyecto: Optional[str] = None
    comentarios: Optional[str] = None

class RegistroTareaRespuesta(RegistroTarea):
    id: int
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

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

# Servir el index.html
@app.get("/")
def leer_index():
    ruta = os.path.join(os.path.dirname(__file__), "index.html")
    if os.path.exists(ruta):
        return FileResponse(ruta)
    return {"error": "No se encuentra el archivo index.html"}

# Login
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

# Crear usuario (solo desde el panel de admin)
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

# Obtener todos los usuarios
@app.get("/usuarios/", response_model=List[UsuarioRespuesta])
def obtener_usuarios(db: Session = Depends(get_db)):
    return db.query(UsuarioDB).all()

# Eliminar usuario
@app.delete("/usuarios/{id}", status_code=204)
def eliminar_usuario(id: int, db: Session = Depends(get_db)):
    usuario = db.query(UsuarioDB).filter(UsuarioDB.id == id).first()
    if not usuario:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    db.delete(usuario)
    db.commit()

# Crear registro
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

# Obtener todos los registros
@app.get("/registros/", response_model=List[RegistroTareaRespuesta])
def obtener_registros(db: Session = Depends(get_db)):
    return db.query(RegistroDB).all()

# Obtener registros por usuario
@app.get("/registros/{usuario}", response_model=List[RegistroTareaRespuesta])
def obtener_registros_por_usuario(usuario: str, db: Session = Depends(get_db)):
    resultado = db.query(RegistroDB).filter(RegistroDB.usuario.ilike(usuario)).all()
    if not resultado:
        raise HTTPException(status_code=404, detail="No se encontraron registros")
    return resultado

# Editar registro
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
    db_registro.usuario        = registro.usuario
    db_registro.fecha          = registro.fecha
    db_registro.tarea          = registro.tarea
    db_registro.tiempo_minutos = registro.tiempo_minutos
    db_registro.proyecto       = registro.proyecto
    db_registro.comentarios    = registro.comentarios
    db.commit()
    db.refresh(db_registro)
    return db_registro

# Eliminar registro
@app.delete("/registros/{id}", status_code=204)
def eliminar_registro(id: int, db: Session = Depends(get_db)):
    db_registro = db.query(RegistroDB).filter(RegistroDB.id == id).first()
    if not db_registro:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    db.delete(db_registro)
    db.commit()

# Exportar Excel
@app.get("/exportar-excel/")
def exportar_excel(db: Session = Depends(get_db)):
    registros = db.query(RegistroDB).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Registros"
    ws.append(["ID", "Usuario", "Fecha", "Tarea", "Tiempo (min)", "Horas", "Proyecto", "Comentarios"])
    for r in registros:
        ws.append([r.id, r.usuario, str(r.fecha), r.tarea, r.tiempo_minutos, round(r.tiempo_minutos/60,2), r.proyecto or "", r.comentarios or ""])
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=registros.xlsx"}
    )

# Bloque para arrancar localmente y compatible con Render
if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port)
