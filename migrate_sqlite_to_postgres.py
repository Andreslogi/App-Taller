"""Migra la base SQLite del piloto a PostgreSQL sin perder IDs ni relaciones.

Uso (PowerShell):
  $env:DATABASE_URL="postgresql://usuario:clave@host/base"
  python migrate_sqlite_to_postgres.py --sqlite "C:\\ruta\\inventario.db"

Por seguridad, el destino debe estar vacío. Use --replace solo si desea borrar el
contenido de las tablas de destino antes de migrar.
"""
from __future__ import annotations

import argparse
import os
import sqlite3
from datetime import datetime
from pathlib import Path

from sqlalchemy import DateTime, create_engine, func, select, text
from sqlalchemy.exc import SQLAlchemyError
from dotenv import load_dotenv

load_dotenv()

from database import db

TABLE_ORDER = [
    "users",
    "settings",
    "products",
    "customers",
    "salespeople",
    "services",
    "invoices",
    "invoice_items",
    "invoice_service_items",
    "movements",
]


def normalize_url(url: str) -> str:
    url = url.strip()
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]
    if url.startswith("postgresql://"):
        url = "postgresql+psycopg://" + url[len("postgresql://"):]
    return url


def convert_row(table, row: dict):
    converted = dict(row)
    for col in table.columns:
        if isinstance(col.type, DateTime) and converted.get(col.name) not in (None, ""):
            value = converted[col.name]
            if isinstance(value, str):
                try:
                    converted[col.name] = datetime.fromisoformat(value.replace("Z", "+00:00"))
                except ValueError:
                    converted[col.name] = datetime.strptime(value[:19], "%Y-%m-%d %H:%M:%S")
    return converted


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--sqlite", required=True, help="Ruta del inventario.db actual")
    parser.add_argument("--replace", action="store_true", help="Vacía las tablas PostgreSQL antes de copiar")
    args = parser.parse_args()

    sqlite_path = Path(args.sqlite).expanduser().resolve()
    if not sqlite_path.exists():
        raise SystemExit(f"No existe: {sqlite_path}")

    target_url = normalize_url(os.getenv("DATABASE_URL", ""))
    if not target_url.startswith("postgresql+"):
        raise SystemExit("Define DATABASE_URL con la conexión PostgreSQL de Render antes de ejecutar.")

    source = sqlite3.connect(sqlite_path)
    source.row_factory = sqlite3.Row
    engine = create_engine(target_url, pool_pre_ping=True)

    # Crea esquema destino usando los mismos modelos de la aplicación.
    db.metadata.create_all(engine)

    try:
        with engine.begin() as conn:
            if args.replace:
                # Orden inverso por claves foráneas.
                for name in reversed(TABLE_ORDER):
                    if name in db.metadata.tables:
                        conn.execute(text(f'TRUNCATE TABLE "{name}" RESTART IDENTITY CASCADE'))
            else:
                non_empty = []
                for name in TABLE_ORDER:
                    table = db.metadata.tables[name]
                    count = conn.execute(select(func.count()).select_from(table)).scalar_one()
                    if count:
                        non_empty.append(f"{name}={count}")
                if non_empty:
                    raise SystemExit(
                        "La base PostgreSQL no está vacía. No se migró nada. "
                        "Contenido encontrado: " + ", ".join(non_empty) +
                        ". Si realmente desea reemplazarla, ejecute con --replace."
                    )

            for name in TABLE_ORDER:
                src_exists = source.execute(
                    "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)
                ).fetchone()
                if not src_exists:
                    print(f"[OMITIDA] {name}: no existe en SQLite")
                    continue

                rows = [dict(r) for r in source.execute(f'SELECT * FROM "{name}"').fetchall()]
                if not rows:
                    print(f"[VACÍA] {name}")
                    continue

                table = db.metadata.tables[name]
                valid_cols = {c.name for c in table.columns}
                cleaned = []
                for row in rows:
                    row = {k: v for k, v in row.items() if k in valid_cols}
                    cleaned.append(convert_row(table, row))
                conn.execute(table.insert(), cleaned)
                print(f"[OK] {name}: {len(cleaned)} registros")

            # Reajusta secuencias después de conservar IDs explícitos.
            for name in TABLE_ORDER:
                table = db.metadata.tables[name]
                if "id" not in table.columns:
                    continue
                conn.execute(text(
                    f"SELECT setval(pg_get_serial_sequence('{name}', 'id'), "
                    f"COALESCE((SELECT MAX(id) FROM {name}), 1), "
                    f"(SELECT COUNT(*) > 0 FROM {name}))"
                ))

        print("\nMigración terminada correctamente.")
    except SQLAlchemyError as exc:
        raise SystemExit(f"Error de PostgreSQL: {exc}") from exc
    finally:
        source.close()
        engine.dispose()


if __name__ == "__main__":
    main()
