from __future__ import annotations

import os
import re
import sys
from contextlib import contextmanager
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable

from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text, func
from sqlalchemy.exc import IntegrityError


def writable_base() -> Path:
    if getattr(sys, "frozen", False):
        base = Path(os.getenv("LOCALAPPDATA", Path.home())) / "InventarioFacturacion"
    else:
        base = Path(__file__).resolve().parent
    base.mkdir(parents=True, exist_ok=True)
    return base


BASE_DIR = writable_base()
LOCAL_DB_PATH = BASE_DIR / "data" / "inventario.db"
LOCAL_DB_PATH.parent.mkdir(parents=True, exist_ok=True)


def database_url() -> str:
    url = os.getenv("DATABASE_URL", "").strip()
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]
    if url.startswith("postgresql://"):
        url = "postgresql+psycopg://" + url[len("postgresql://"):]
    if url:
        return url
    return f"sqlite:///{LOCAL_DB_PATH.as_posix()}"


db = SQLAlchemy()


class User(db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(100), nullable=False, unique=True)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(180), nullable=False)
    role = db.Column(db.String(30), nullable=False)
    active = db.Column(db.Integer, nullable=False, default=1, server_default=text("1"))
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, server_default=func.now())


class Setting(db.Model):
    __tablename__ = "settings"
    id = db.Column(db.Integer, primary_key=True)
    business_name = db.Column(db.String(200), nullable=False, default="Mi Negocio")
    nit = db.Column(db.String(80), default="")
    address = db.Column(db.String(250), default="")
    phone = db.Column(db.String(80), default="")
    email = db.Column(db.String(150), default="")
    invoice_prefix = db.Column(db.String(20), nullable=False, default="REM")
    next_invoice = db.Column(db.Integer, nullable=False, default=1)
    tax_rate = db.Column(db.Numeric(8, 4), nullable=False, default=0)


class Product(db.Model):
    __tablename__ = "products"
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(100), nullable=False, unique=True)
    description = db.Column(db.String(300), nullable=False)
    category = db.Column(db.String(150), default="")
    unit = db.Column(db.String(60), default="Unidad")
    stock = db.Column(db.Numeric(16, 4), nullable=False, default=0)
    min_stock = db.Column(db.Numeric(16, 4), nullable=False, default=1)
    cost = db.Column(db.Numeric(16, 2), nullable=False, default=0)
    sale_price = db.Column(db.Numeric(16, 2), nullable=False, default=0)
    notes = db.Column(db.Text, default="")
    active = db.Column(db.Integer, nullable=False, default=1, server_default=text("1"))
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, server_default=func.now())
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow, server_default=func.now())


class Customer(db.Model):
    __tablename__ = "customers"
    id = db.Column(db.Integer, primary_key=True)
    document = db.Column(db.String(100), default="")
    name = db.Column(db.String(250), nullable=False)
    phone = db.Column(db.String(100), default="")
    email = db.Column(db.String(150), default="")
    address = db.Column(db.String(300), default="")
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, server_default=func.now())


class Salesperson(db.Model):
    __tablename__ = "salespeople"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(180), nullable=False, unique=True)
    active = db.Column(db.Integer, nullable=False, default=1, server_default=text("1"))
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, server_default=func.now())
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow, server_default=func.now())


class Invoice(db.Model):
    __tablename__ = "invoices"
    id = db.Column(db.Integer, primary_key=True)
    number = db.Column(db.String(60), nullable=False, unique=True)
    customer_id = db.Column(db.Integer, db.ForeignKey("customers.id"))
    customer_name = db.Column(db.String(250), nullable=False)
    customer_document = db.Column(db.String(100), default="")
    customer_phone = db.Column(db.String(100), default="")
    customer_address = db.Column(db.String(300), default="")
    vehicle_plate = db.Column(db.String(30), default="")
    salesperson_id = db.Column(db.Integer, db.ForeignKey("salespeople.id"))
    salesperson_name = db.Column(db.String(180), default="")
    issued_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    subtotal = db.Column(db.Numeric(16, 2), nullable=False, default=0)
    tax = db.Column(db.Numeric(16, 2), nullable=False, default=0)
    total = db.Column(db.Numeric(16, 2), nullable=False, default=0)
    payment_method = db.Column(db.String(80), default="Efectivo")
    notes = db.Column(db.Text, default="")
    status = db.Column(db.String(40), nullable=False, default="Emitida", server_default=text("'Emitida'"))
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"))


class InvoiceItem(db.Model):
    __tablename__ = "invoice_items"
    id = db.Column(db.Integer, primary_key=True)
    invoice_id = db.Column(db.Integer, db.ForeignKey("invoices.id", ondelete="CASCADE"), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey("products.id"), nullable=False)
    code = db.Column(db.String(100), nullable=False)
    description = db.Column(db.String(300), nullable=False)
    quantity = db.Column(db.Numeric(16, 4), nullable=False)
    unit_price = db.Column(db.Numeric(16, 2), nullable=False)
    line_total = db.Column(db.Numeric(16, 2), nullable=False)


class Service(db.Model):
    __tablename__ = "services"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(250), nullable=False, unique=True)
    default_price = db.Column(db.Numeric(16, 2), nullable=False, default=0)
    worker_percentage = db.Column(db.Numeric(8, 4), nullable=False, default=70)
    business_percentage = db.Column(db.Numeric(8, 4), nullable=False, default=30)
    active = db.Column(db.Integer, nullable=False, default=1, server_default=text("1"))
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, server_default=func.now())
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow, server_default=func.now())


class InvoiceServiceItem(db.Model):
    __tablename__ = "invoice_service_items"
    id = db.Column(db.Integer, primary_key=True)
    invoice_id = db.Column(db.Integer, db.ForeignKey("invoices.id", ondelete="CASCADE"), nullable=False)
    service_id = db.Column(db.Integer, db.ForeignKey("services.id"))
    service_name = db.Column(db.String(250), nullable=False)
    worker_name = db.Column(db.String(180), nullable=False, default="")
    quantity = db.Column(db.Numeric(16, 4), nullable=False, default=1)
    unit_price = db.Column(db.Numeric(16, 2), nullable=False)
    line_total = db.Column(db.Numeric(16, 2), nullable=False)
    worker_percentage = db.Column(db.Numeric(8, 4), nullable=False, default=70)
    business_percentage = db.Column(db.Numeric(8, 4), nullable=False, default=30)
    worker_amount = db.Column(db.Numeric(16, 2), nullable=False)
    business_amount = db.Column(db.Numeric(16, 2), nullable=False)


class AccountReceivable(db.Model):
    __tablename__ = "accounts_receivable"
    id = db.Column(db.Integer, primary_key=True)
    invoice_id = db.Column(db.Integer, db.ForeignKey("invoices.id", ondelete="CASCADE"), nullable=False, unique=True)
    original_amount = db.Column(db.Numeric(16, 2), nullable=False, default=0)
    due_date = db.Column(db.Date, nullable=True)
    status = db.Column(db.String(30), nullable=False, default="PENDIENTE", server_default=text("'PENDIENTE'"))
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, server_default=func.now())
    closed_at = db.Column(db.DateTime, nullable=True)


class AccountReceivablePayment(db.Model):
    __tablename__ = "accounts_receivable_payments"
    id = db.Column(db.Integer, primary_key=True)
    receivable_id = db.Column(db.Integer, db.ForeignKey("accounts_receivable.id", ondelete="CASCADE"), nullable=False)
    amount = db.Column(db.Numeric(16, 2), nullable=False)
    payment_date = db.Column(db.Date, nullable=False)
    payment_method = db.Column(db.String(80), nullable=False, default="Efectivo")
    notes = db.Column(db.Text, default="")
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, server_default=func.now())

class AccountPayable(db.Model):
    __tablename__ = "accounts_payable"

    id = db.Column(db.Integer, primary_key=True)

    supplier = db.Column(
        db.String(250),
        nullable=False
    )

    invoice_number = db.Column(
        db.String(120),
        nullable=False
    )

    purchase_date = db.Column(
        db.Date,
        nullable=False
    )

    due_date = db.Column(
        db.Date,
        nullable=False
    )

    original_amount = db.Column(
        db.Numeric(16, 2),
        nullable=False,
        default=0
    )

    description = db.Column(
        db.String(500),
        default=""
    )

    notes = db.Column(
        db.Text,
        default=""
    )

    status = db.Column(
        db.String(30),
        nullable=False,
        default="PENDIENTE",
        server_default=text("'PENDIENTE'")
    )

    created_at = db.Column(
        db.DateTime,
        nullable=False,
        default=datetime.utcnow,
        server_default=func.now()
    )

    updated_at = db.Column(
        db.DateTime,
        nullable=False,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        server_default=func.now()
    )

    closed_at = db.Column(
        db.DateTime,
        nullable=True
    )

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id")
    )


class AccountPayablePayment(db.Model):
    __tablename__ = "accounts_payable_payments"

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    payable_id = db.Column(
        db.Integer,
        db.ForeignKey(
            "accounts_payable.id",
            ondelete="CASCADE"
        ),
        nullable=False
    )

    amount = db.Column(
        db.Numeric(16, 2),
        nullable=False
    )

    payment_date = db.Column(
        db.Date,
        nullable=False
    )

    payment_method = db.Column(
        db.String(80),
        nullable=False,
        default="Transferencia"
    )

    notes = db.Column(
        db.Text,
        default=""
    )

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id")
    )

    created_at = db.Column(
        db.DateTime,
        nullable=False,
        default=datetime.utcnow,
        server_default=func.now()
    )


class Movement(db.Model):
    __tablename__ = "movements"
    id = db.Column(db.Integer, primary_key=True)
    product_id = db.Column(db.Integer, db.ForeignKey("products.id"), nullable=False)
    movement_type = db.Column(db.String(60), nullable=False)
    quantity = db.Column(db.Numeric(16, 4), nullable=False)
    reference = db.Column(db.String(180), default="")
    notes = db.Column(db.Text, default="")
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, server_default=func.now())
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"))


def _normalize_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.isoformat(sep=" ", timespec="seconds")
    return value


def _row_to_dict(row: Any) -> dict[str, Any] | None:
    if row is None:
        return None
    mapping = getattr(row, "_mapping", row)
    return {k: _normalize_value(v) for k, v in dict(mapping).items()}


def _qmark_to_named(sql: str, params: Iterable[Any] | None):
    params = tuple(params or ())
    idx = 0
    out = []
    in_single = False
    in_double = False
    for ch in sql:
        if ch == "'" and not in_double:
            in_single = not in_single
            out.append(ch)
        elif ch == '"' and not in_single:
            in_double = not in_double
            out.append(ch)
        elif ch == "?" and not in_single and not in_double:
            out.append(f":p{idx}")
            idx += 1
        else:
            out.append(ch)
    if idx != len(params):
        raise ValueError(f"Cantidad de parámetros no coincide: SQL espera {idx}, recibió {len(params)}")
    return "".join(out), {f"p{i}": v for i, v in enumerate(params)}


class ResultAdapter:
    def __init__(self, result, lastrowid=None):
        self._result = result
        self.lastrowid = lastrowid
        self.rowcount = getattr(result, "rowcount", -1)

    def fetchone(self):
        try:
            return _row_to_dict(self._result.fetchone())
        except Exception:
            return None

    def fetchall(self):
        try:
            return [_row_to_dict(r) for r in self._result.fetchall()]
        except Exception:
            return []

    def __iter__(self):
        try:
            for row in self._result:
                yield _row_to_dict(row)
        except Exception:
            return


class CompatConnection:
    """Capa mínima para conservar las consultas actuales mientras el motor es SQLAlchemy/PostgreSQL."""
    def __init__(self, connection):
        self.connection = connection
        self.dialect = connection.dialect.name

    def execute(self, sql: str, params: Iterable[Any] | None = None):
        sql = sql.strip()
        named_sql, bind = _qmark_to_named(sql, params)
        lower = named_sql.lower().lstrip()
        insert_match = re.match(r"insert\s+into\s+([a-zA-Z_][a-zA-Z0-9_]*)", lower)
        lastrowid = None
        if insert_match and insert_match.group(1) in {
            "users", "products", "customers", "salespeople", "invoices", "invoice_items",
            "services", "invoice_service_items", "movements",
            "accounts_receivable", "accounts_receivable_payments",
            "accounts_payable", "accounts_payable_payments"
        } and " returning " not in lower and self.dialect == "postgresql":
            named_sql += " RETURNING id"
            result = self.connection.execute(text(named_sql), bind)
            row = result.fetchone()
            if row is not None:
                lastrowid = row[0]
            return ResultAdapter(result, lastrowid=lastrowid)
        result = self.connection.execute(text(named_sql), bind)
        if self.dialect == "sqlite":
            lastrowid = getattr(result, "lastrowid", None)
        return ResultAdapter(result, lastrowid=lastrowid)


@contextmanager
def db_conn():
    connection = db.engine.connect()
    trans = connection.begin()
    try:
        yield CompatConnection(connection)
        trans.commit()
    except Exception:
        trans.rollback()
        raise
    finally:
        connection.close()


__all__ = [
    "db", "db_conn", "IntegrityError", "BASE_DIR", "LOCAL_DB_PATH", "database_url",
    "User", "Setting", "Product", "Customer", "Salesperson", "Invoice", "InvoiceItem",
    "Service", "InvoiceServiceItem", "Movement", "AccountReceivable",
    "AccountReceivablePayment","AccountPayable","AccountPayablePayment",
]
