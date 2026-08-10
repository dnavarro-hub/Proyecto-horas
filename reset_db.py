import os
from sqlalchemy import create_engine
from sqlalchemy.engine.url import make_url

DATABASE_URL = os.environ.get(
    "DATABASE_URL", "postgresql://localhost/registros"
)
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

url = make_url(DATABASE_URL)
admin_url = url.set(database="postgres")
admin_engine = create_engine(admin_url, isolation_level="AUTOCOMMIT")

with admin_engine.connect() as conn:
    db_name = url.database
    print(f"Borrando y recreando la base de datos '{db_name}'...")
    conn.execute(
        f"SELECT pg_terminate_backend(pid) FROM pg_stat_activity WHERE datname = '{db_name}' AND pid <> pg_backend_pid()"
    )
    conn.execute(f"DROP DATABASE IF EXISTS {db_name}")
    conn.execute(f"CREATE DATABASE {db_name}")

print("Base de datos reiniciada con éxito.")
