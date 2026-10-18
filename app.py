from __future__ import annotations

import os
import sys
import threading
import time
import webbrowser
import csv
from io import StringIO
from functools import wraps
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from flask import Flask, flash, g, redirect, render_template, request, send_file, session, url_for, Response
from flask_migrate import Migrate
from dotenv import load_dotenv
from database import db, db_conn, IntegrityError, BASE_DIR as DATA_BASE_DIR, LOCAL_DB_PATH, database_url
from openpyxl import load_workbook
from werkzeug.security import check_password_hash, generate_password_hash
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from expenses_module import expenses_bp
from profitability_module import profitability_bp
from extra_income_module import extra_income_bp
load_dotenv()

APP_NAME = "Control de Inventario y Remisiones"

# Mecánicos autorizados para asignar servicios en las remisiones.
ALLOWED_MECHANICS = ("Leo", "Daniel", "James")


def resource_path(relative: str) -> Path:
    """Ruta compatible con ejecución normal y PyInstaller."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


BASE_DIR = DATA_BASE_DIR
DB_PATH = LOCAL_DB_PATH
INVOICE_DIR = BASE_DIR / "invoices"
BACKUP_DIR = BASE_DIR / "backups"
INITIAL_XLSX = resource_path("data/Inventario_Inicial.xlsx")
SERVICES_XLSX = resource_path("data/Servicios_Iniciales.xlsx")

for folder in (INVOICE_DIR, BACKUP_DIR):
    folder.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "cambia-esta-clave-en-produccion")
app.config["SQLALCHEMY_DATABASE_URI"] = database_url()
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {"pool_pre_ping": True}
db.init_app(app)
migrate = Migrate(app, db)
app.register_blueprint(expenses_bp)
app.register_blueprint(profitability_bp)
app.register_blueprint(extra_income_bp)

def money(value: Any) -> str:
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        number = 0
    return "$ {:,.0f}".format(number).replace(",", ".")


app.jinja_env.filters["money"] = money


def parse_decimal(value: str | None, default: Decimal = Decimal("0")) -> Decimal:
    if value is None:
        return default
    cleaned = str(value).strip().replace("$", "").replace(" ", "")
    if not cleaned:
        return default
    # Soporta 1.234.567,89 y 1234567.89
    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    elif "," in cleaned:
        cleaned = cleaned.replace(",", ".")
    try:
        return Decimal(cleaned)
    except InvalidOperation as exc:
        raise ValueError(f"Valor numérico inválido: {value}") from exc


def login_required(view):
    @wraps(view)
    def wrapped_view(**kwargs):
        if g.user is None:
            flash("Debes iniciar sesión.", "warning")
            return redirect(url_for("login", next=request.path))
        return view(**kwargs)
    return wrapped_view


def role_required(*roles):
    def decorator(view):
        @wraps(view)
        def wrapped_view(**kwargs):
            if g.user is None:
                flash("Debes iniciar sesión.", "warning")
                return redirect(url_for("login"))
            if g.user["role"] not in roles:
                flash("No tienes permisos para realizar esta acción.", "danger")
                return redirect(url_for("dashboard"))
            return view(**kwargs)
        return wrapped_view
    return decorator

def arturo_required(view):
    """Permite editar/eliminar remisiones únicamente a arturo.lopez administrador."""
    @wraps(view)
    def wrapped_view(**kwargs):
        if g.user is None:
            flash("Debes iniciar sesión.", "warning")
            return redirect(url_for("login"))

        username = str(g.user["username"] or "").strip().lower()
        role = str(g.user["role"] or "").strip().lower()

        if username != "arturo.lopez" or role != "administrador":
            flash(
                "Solo el usuario arturo.lopez puede editar o eliminar remisiones.",
                "danger",
            )
            return redirect(url_for("invoices"))

        return view(**kwargs)

    return wrapped_view


@app.before_request
def load_logged_in_user():
    user_id = session.get("user_id")
    if user_id is None:
        g.user = None
    else:
        with db_conn() as conn:
            g.user = conn.execute(
                "SELECT id, username, full_name, role, active FROM users WHERE id=?",
                (user_id,),
            ).fetchone()
        if g.user is None or not g.user["active"]:
            session.clear()
            g.user = None


def init_db() -> None:
    """Crea el esquema y carga datos predeterminados de forma aditiva."""
    with app.app_context():
        db.create_all()
        ensure_schema_updates()
        with db_conn() as conn:
            settings = conn.execute("SELECT * FROM settings WHERE id=1").fetchone()
            if not settings:
                conn.execute("""INSERT INTO settings(
                    id,business_name,nit,address,phone,email,invoice_prefix,next_invoice,tax_rate
                ) VALUES(1,?,?,?,?,?,?,?,?)""",
                ("Mi Negocio", "", "", "", "", "REM", 1, 0))
            else:
                conn.execute("UPDATE settings SET invoice_prefix='REM' WHERE id=1 AND invoice_prefix='FV'")

            default_users = [
                ("administrador", "admin123*", "Administrador", "administrador"),
                ("vendedor", "vendedor123*", "Vendedor genérico", "vendedor"),
                ("arturo.lopez", "Arturo1963*", "Arturo López", "administrador"),
            ]
            for username, password, full_name, role in default_users:
                if not conn.execute("SELECT id FROM users WHERE username=?", (username,)).fetchone():
                    conn.execute(
                        "INSERT INTO users(username,password_hash,full_name,role,active) VALUES(?,?,?,?,1)",
                        (username, generate_password_hash(password), full_name, role),
                    )

            for seller_name in ("James", "Leonardo"):
                if not conn.execute("SELECT id FROM salespeople WHERE name=?", (seller_name,)).fetchone():
                    conn.execute("INSERT INTO salespeople(name,active) VALUES(?,1)", (seller_name,))

        seed_services_from_excel()
        seed_from_excel_if_empty()


def ensure_schema_updates() -> None:
    """Agrega columnas nuevas a tablas existentes (create_all no altera tablas ya creadas)."""
    from sqlalchemy import inspect

    inspector = inspect(db.engine)
    payment_columns = {c["name"] for c in inspector.get_columns("accounts_receivable_payments")}
    if "payment_kind" not in payment_columns:
        with db_conn() as conn:
            conn.execute(
                "ALTER TABLE accounts_receivable_payments "
                "ADD COLUMN payment_kind VARCHAR(20) NOT NULL DEFAULT 'ABONO'"
            )


def seed_services_from_excel() -> None:
    if not SERVICES_XLSX.exists():
        return
    wb = load_workbook(SERVICES_XLSX, data_only=True)
    added = 0
    with db_conn() as conn:
        for ws in wb.worksheets:
            # Encabezados útiles empiezan en la fila 3 en el archivo entregado.
            for row in ws.iter_rows(min_row=3, values_only=True):
                category = str(row[0] or "").strip()
                model = str(row[1] or "").strip()
                price = row[2] if len(row) > 2 else None
                if not category or not model or not isinstance(price, (int, float)):
                    continue
                name = f"{category.strip()} - {model.strip()}"
                if not conn.execute("SELECT id FROM services WHERE LOWER(name)=LOWER(?)", (name,)).fetchone():
                    conn.execute(
                        """INSERT INTO services(name,default_price,worker_percentage,business_percentage,active)
                           VALUES(?,?,70,30,1)""",
                        (name, float(price)),
                    )
                    added += 1
    if added:
        print(f"Servicios iniciales agregados: {added}")


def seed_from_excel_if_empty() -> None:
    with db_conn() as conn:
        row = conn.execute("SELECT COUNT(*) AS total FROM products").fetchone()
        if (row and int(row["total"]) > 0) or not INITIAL_XLSX.exists():
            return

        wb = load_workbook(INITIAL_XLSX, data_only=False)
        ws = wb["Inventario"]
        imported = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            code, description, category, unit, stock_initial = row[0], row[1], row[2], row[3], row[4]
            total_initial, min_stock, notes = row[8], row[11], row[13]
            if not code or not description:
                continue
            stock = float(stock_initial or 0)
            total = float(total_initial or 0)
            cost = total / stock if stock else 0
            code = str(code)
            if conn.execute("SELECT id FROM products WHERE code=?", (code,)).fetchone():
                continue
            conn.execute(
                """INSERT INTO products
                (code, description, category, unit, stock, min_stock, cost, sale_price, notes, active)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1)""",
                (code, str(description), str(category or ""), str(unit or "Unidad"), stock,
                 float(min_stock or 1), cost, cost, str(notes or "")),
            )
            imported += 1
        print(f"Productos importados: {imported}")

@app.route("/productos/exportar-csv")
def export_products_csv():

    if g.user is None:

        flash(
            "Debes iniciar sesión.",
            "warning"
        )

        return redirect(
            url_for("login")
        )

    with db_conn() as conn:

        products = conn.execute(
            """
            SELECT
                code,
                description,
                category,
                unit,
                stock,
                sale_price,
                active

            FROM products

            ORDER BY
                description ASC
            """
        ).fetchall()

    output = StringIO()

    # BOM para que Excel abra correctamente
    # tildes, ñ y caracteres especiales.

    output.write("\ufeff")

    writer = csv.writer(
        output,
        delimiter=";",
        lineterminator="\n"
    )

    writer.writerow(
        [
            "Código",
            "Descripción",
            "Categoría",
            "Unidad",
            "Stock",
            "Precio de venta",
            "Activo",
        ]
    )

    for product in products:

        writer.writerow(
            [
                product["code"],
                product["description"],
                product["category"],
                product["unit"],
                product["stock"],
                product["sale_price"],
                (
                    "Sí"
                    if product["active"]
                    else
                    "No"
                ),
            ]
        )

    csv_content = output.getvalue()

    output.close()

    return Response(
        csv_content,
        mimetype=(
            "text/csv; "
            "charset=utf-8"
        ),
        headers={
            "Content-Disposition":
                "attachment; "
                "filename=lista_precios_productos.csv"
        }
    )

def get_settings(conn=None, lock: bool = False):
    sql = "SELECT * FROM settings WHERE id=1"
    if lock and db.engine.dialect.name == "postgresql":
        sql += " FOR UPDATE"
    if conn is not None:
        return conn.execute(sql).fetchone()
    with db_conn() as local:
        return local.execute(sql).fetchone()




@app.route("/healthz")
def healthz():
    return {"status": "ok", "database": db.engine.dialect.name}, 200


@app.route("/login", methods=["GET", "POST"])
def login():
    if g.user is not None:
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        username = request.form.get("username", "").strip().lower()
        password = request.form.get("password", "")
        with db_conn() as conn:
            user = conn.execute(
                "SELECT * FROM users WHERE username=? AND active=1", (username,)
            ).fetchone()
        if user is None or not check_password_hash(user["password_hash"], password):
            flash("Usuario o contraseña incorrectos.", "danger")
        else:
            session.clear()
            session["user_id"] = user["id"]
            return redirect(request.args.get("next") or url_for("dashboard"))
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/usuarios")
@role_required("administrador")
def users():
    with db_conn() as conn:
        rows = conn.execute(
            "SELECT id,username,full_name,role,active,created_at FROM users ORDER BY full_name"
        ).fetchall()
    return render_template("users.html", users=rows)


@app.route("/usuarios/nuevo", methods=["GET", "POST"])
@role_required("administrador")
def user_new():
    if request.method == "POST":
        username = request.form.get("username", "").strip().lower()
        full_name = request.form.get("full_name", "").strip()
        password = request.form.get("password", "")
        role = request.form.get("role", "vendedor")
        try:
            if not username or not full_name:
                raise ValueError("El usuario y el nombre son obligatorios.")
            if len(password) < 8:
                raise ValueError("La contraseña debe tener al menos 8 caracteres.")
            if role not in ("administrador", "vendedor"):
                raise ValueError("Rol inválido.")
            with db_conn() as conn:
                conn.execute(
                    "INSERT INTO users(username,password_hash,full_name,role) VALUES(?,?,?,?)",
                    (username, generate_password_hash(password), full_name, role),
                )
            flash("Usuario creado correctamente.", "success")
            return redirect(url_for("users"))
        except IntegrityError:
            flash("Ese nombre de usuario ya existe.", "danger")
        except ValueError as exc:
            flash(str(exc), "danger")
    return render_template("user_form.html", user=None)


@app.route("/usuarios/<int:user_id>/editar", methods=["GET", "POST"])
@role_required("administrador")
def user_edit(user_id: int):
    with db_conn() as conn:
        user = conn.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
    if not user:
        flash("Usuario no encontrado.", "danger")
        return redirect(url_for("users"))

    if request.method == "POST":
        full_name = request.form.get("full_name", "").strip()
        role = request.form.get("role", "vendedor")
        active = 1 if request.form.get("active") == "1" else 0
        password = request.form.get("password", "")
        if role not in ("administrador", "vendedor"):
            flash("Rol inválido.", "danger")
            return render_template("user_form.html", user=user)
        if user_id == g.user["id"] and not active:
            flash("No puedes desactivar tu propio usuario.", "danger")
            return redirect(url_for("users"))
        with db_conn() as conn:
            if password:
                if len(password) < 8:
                    flash("La contraseña debe tener al menos 8 caracteres.", "danger")
                    return render_template("user_form.html", user=user)
                conn.execute(
                    "UPDATE users SET full_name=?,role=?,active=?,password_hash=? WHERE id=?",
                    (full_name, role, active, generate_password_hash(password), user_id),
                )
            else:
                conn.execute(
                    "UPDATE users SET full_name=?,role=?,active=? WHERE id=?",
                    (full_name, role, active, user_id),
                )
        flash("Usuario actualizado.", "success")
        return redirect(url_for("users"))
    return render_template("user_form.html", user=user)


@app.route("/")
@login_required
def dashboard():
    with db_conn() as conn:
        stats = conn.execute(
            """
            SELECT COUNT(*) product_count,
                   COALESCE(SUM(stock),0) total_units,
                   COALESCE(SUM(stock * cost),0) inventory_cost,
                   COALESCE(SUM(stock * sale_price),0) inventory_sale,
                   SUM(CASE WHEN stock <= 0 THEN 1 ELSE 0 END) out_count,
                   SUM(CASE WHEN stock > 0 AND stock <= min_stock THEN 1 ELSE 0 END) low_count
            FROM products WHERE active=1
            """
        ).fetchone()
        recent_invoices = conn.execute("SELECT * FROM invoices ORDER BY id DESC LIMIT 8").fetchall()
        low_products = conn.execute(
            "SELECT * FROM products WHERE active=1 AND stock <= min_stock ORDER BY stock ASC LIMIT 10"
        ).fetchall()
    return render_template("dashboard.html", stats=stats, recent_invoices=recent_invoices, low_products=low_products)


@app.route("/productos")
@login_required
def products():
    q = request.args.get("q", "").strip()
    with db_conn() as conn:
        if q:
            rows = conn.execute(
                """SELECT * FROM products WHERE active=1 AND
                (code LIKE ? OR description LIKE ? OR category LIKE ?)
                ORDER BY description""",
                (f"%{q}%", f"%{q}%", f"%{q}%"),
            ).fetchall()
        else:
            rows = conn.execute("SELECT * FROM products WHERE active=1 ORDER BY description").fetchall()
    return render_template("products.html", products=rows, q=q)


@app.route("/productos/nuevo", methods=["GET", "POST"])
@role_required("administrador")
def product_new():
    if request.method == "POST":
        try:
            values = (
                request.form["code"].strip(), request.form["description"].strip(),
                request.form.get("category", "").strip(), request.form.get("unit", "Unidad").strip(),
                float(parse_decimal(request.form.get("stock"))), float(parse_decimal(request.form.get("min_stock"), Decimal("1"))),
                float(parse_decimal(request.form.get("cost"))), float(parse_decimal(request.form.get("sale_price"))),
                request.form.get("notes", "").strip(),
            )
            with db_conn() as conn:
                conn.execute(
                    """INSERT INTO products(code, description, category, unit, stock, min_stock, cost, sale_price, notes)
                    VALUES(?,?,?,?,?,?,?,?,?)""", values
                )
            flash("Producto creado correctamente.", "success")
            return redirect(url_for("products"))
        except IntegrityError:
            flash("El código del producto ya existe.", "danger")
        except (ValueError, KeyError) as exc:
            flash(str(exc), "danger")
    return render_template("product_form.html", product=None)


@app.route("/productos/<int:product_id>/editar", methods=["GET", "POST"])
@role_required("administrador")
def product_edit(product_id: int):
    with db_conn() as conn:
        product = conn.execute("SELECT * FROM products WHERE id=?", (product_id,)).fetchone()
    if not product:
        flash("Producto no encontrado.", "danger")
        return redirect(url_for("products"))

    if request.method == "POST":
        try:
            values = (
                request.form["code"].strip(), request.form["description"].strip(),
                request.form.get("category", "").strip(), request.form.get("unit", "Unidad").strip(),
                float(parse_decimal(request.form.get("min_stock"), Decimal("1"))),
                float(parse_decimal(request.form.get("cost"))), float(parse_decimal(request.form.get("sale_price"))),
                request.form.get("notes", "").strip(), product_id,
            )
            with db_conn() as conn:
                conn.execute(
                    """UPDATE products SET code=?, description=?, category=?, unit=?, min_stock=?, cost=?,
                    sale_price=?, notes=?, updated_at=CURRENT_TIMESTAMP WHERE id=?""", values
                )
            flash("Producto actualizado.", "success")
            return redirect(url_for("products"))
        except IntegrityError:
            flash("Ese código ya está asignado a otro producto.", "danger")
        except ValueError as exc:
            flash(str(exc), "danger")
    return render_template("product_form.html", product=product)

@app.route(
    "/productos/<int:product_id>/eliminar",
    methods=["POST"]
)
@role_required("administrador")
def product_delete(product_id: int):

    try:

        with db_conn() as conn:

            product = conn.execute(
                """
                SELECT *
                FROM products
                WHERE id=?
                """,
                (product_id,)
            ).fetchone()

            if not product:
                raise ValueError(
                    "Producto no encontrado."
                )


            invoice_usage = conn.execute(
                """
                SELECT COUNT(*) AS total
                FROM invoice_items
                WHERE product_id=?
                """,
                (product_id,)
            ).fetchone()


            movement_usage = conn.execute(
                """
                SELECT COUNT(*) AS total
                FROM movements
                WHERE product_id=?
                """,
                (product_id,)
            ).fetchone()


            has_history = (
                int(invoice_usage["total"] or 0) > 0
                or
                int(movement_usage["total"] or 0) > 0
            )


            # ============================================
            # PRODUCTO SIN HISTORIAL
            # Se puede borrar completamente
            # ============================================

            if not has_history:

                conn.execute(
                    """
                    DELETE FROM products
                    WHERE id=?
                    """,
                    (product_id,)
                )

                flash(
                    f"Producto '{product['description']}' "
                    f"eliminado completamente.",
                    "success"
                )


            # ============================================
            # PRODUCTO CON HISTORIAL
            # No borramos para no dañar remisiones
            # ============================================

            else:

                conn.execute(
                    """
                    UPDATE products

                    SET
                        active=0,
                        updated_at=CURRENT_TIMESTAMP

                    WHERE id=?
                    """,
                    (product_id,)
                )

                flash(
                    f"Producto '{product['description']}' "
                    f"retirado del inventario. "
                    f"Se conservó su historial.",
                    "success"
                )


        return redirect(
            url_for("products")
        )


    except ValueError as exc:

        flash(
            str(exc),
            "danger"
        )

        return redirect(
            url_for("products")
        )    


@app.route("/movimientos", methods=["GET", "POST"])
@role_required("administrador")
def movements():
    if request.method == "POST":
        try:
            product_id = int(request.form["product_id"])
            movement_type = request.form["movement_type"].strip()
            quantity = parse_decimal(request.form["quantity"])

            if quantity <= Decimal("0"):
                raise ValueError("La cantidad debe ser mayor que cero.")

            if movement_type not in ("Entrada", "Salida"):
                raise ValueError("Tipo de movimiento inválido.")

            signed = quantity if movement_type == "Entrada" else -quantity

            with db_conn() as conn:
                product_sql = "SELECT * FROM products WHERE id=?"
                if db.engine.dialect.name == "postgresql":
                    product_sql += " FOR UPDATE"

                product = conn.execute(product_sql, (product_id,)).fetchone()
                if not product:
                    raise ValueError("Producto no encontrado.")

                current_stock = Decimal(str(product["stock"] or 0))
                new_stock = current_stock + signed

                if new_stock < Decimal("0"):
                    raise ValueError("La salida supera el inventario disponible.")

                conn.execute(
                    """UPDATE products
                       SET stock=?, updated_at=CURRENT_TIMESTAMP
                       WHERE id=?""",
                    (float(new_stock), product_id),
                )
                conn.execute(
                    """INSERT INTO movements(
                           product_id,movement_type,quantity,reference,notes,created_at,user_id
                       ) VALUES(?,?,?,?,?,?,?)""",
                    (
                        product_id,
                        movement_type,
                        float(quantity),
                        request.form.get("reference", "").strip(),
                        request.form.get("notes", "").strip(),
                        datetime.now().isoformat(timespec="seconds"),
                        g.user["id"],
                    ),
                )

            flash(
                f"Movimiento registrado correctamente. Nuevo inventario: {new_stock:g}",
                "success",
            )
            return redirect(url_for("movements"))

        except (ValueError, KeyError, TypeError) as exc:
            flash(str(exc), "danger")

    with db_conn() as conn:
        product_rows = conn.execute(
            """SELECT id,code,description,stock
               FROM products
               WHERE active=1
               ORDER BY description"""
        ).fetchall()
        rows = conn.execute(
            """SELECT m.*, p.code, p.description, u.full_name AS user_name
               FROM movements m
               JOIN products p ON p.id=m.product_id
               LEFT JOIN users u ON u.id=m.user_id
               ORDER BY m.id DESC
               LIMIT 100"""
        ).fetchall()

    return render_template(
        "movements.html",
        products=product_rows,
        movements=rows,
    )



@app.route("/vendedores")
@login_required
def salespeople():
    with db_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM salespeople ORDER BY active DESC, name"
        ).fetchall()
    return render_template("salespeople.html", salespeople=rows)


@app.route("/vendedores/nuevo", methods=["GET", "POST"])
@role_required("administrador")
def salesperson_new():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        try:
            if not name:
                raise ValueError("El nombre del vendedor es obligatorio.")
            with db_conn() as conn:
                conn.execute("INSERT INTO salespeople(name,active) VALUES(?,1)", (name,))
            flash("Vendedor agregado correctamente.", "success")
            return redirect(url_for("salespeople"))
        except IntegrityError:
            flash("Ya existe un vendedor con ese nombre.", "danger")
        except ValueError as exc:
            flash(str(exc), "danger")
    return render_template("salesperson_form.html", salesperson=None)


@app.route("/vendedores/<int:salesperson_id>/editar", methods=["GET", "POST"])
@role_required("administrador")
def salesperson_edit(salesperson_id: int):
    with db_conn() as conn:
        seller = conn.execute("SELECT * FROM salespeople WHERE id=?", (salesperson_id,)).fetchone()
    if not seller:
        flash("Vendedor no encontrado.", "danger")
        return redirect(url_for("salespeople"))

    if request.method == "POST":
        name = request.form.get("name", "").strip()
        active = 1 if request.form.get("active") == "1" else 0
        try:
            if not name:
                raise ValueError("El nombre del vendedor es obligatorio.")
            with db_conn() as conn:
                conn.execute(
                    "UPDATE salespeople SET name=?,active=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",
                    (name, active, salesperson_id),
                )
            flash("Vendedor actualizado.", "success")
            return redirect(url_for("salespeople"))
        except IntegrityError:
            flash("Ya existe otro vendedor con ese nombre.", "danger")
        except ValueError as exc:
            flash(str(exc), "danger")
    return render_template("salesperson_form.html", salesperson=seller)


@app.route("/servicios", methods=["GET", "POST"])
@login_required
def services():

    # =====================================================
    # CREAR SERVICIO
    # =====================================================

    if request.method == "POST":

        if g.user["role"] != "administrador":

            flash(
                "Solo el administrador puede crear servicios.",
                "danger"
            )

            return redirect(
                url_for("services")
            )

        try:

            name = request.form.get(
                "name",
                ""
            ).strip()

            price = float(
                parse_decimal(
                    request.form.get(
                        "default_price"
                    )
                )
            )

            if not name:

                raise ValueError(
                    "El nombre del servicio es obligatorio."
                )

            if price < 0:

                raise ValueError(
                    "El valor del servicio no puede ser negativo."
                )

            with db_conn() as conn:

                archived = conn.execute(
                    "SELECT id FROM services WHERE LOWER(name)=LOWER(?) AND active=0",
                    (name,)
                ).fetchone()

                if archived:

                    conn.execute(
                        """UPDATE services SET active=1,default_price=?,updated_at=CURRENT_TIMESTAMP
                           WHERE id=?""",
                        (price, archived["id"])
                    )

                    flash(
                        "Ese servicio estaba eliminado y se restauró con el nuevo valor.",
                        "success"
                    )

                    return redirect(
                        url_for("services")
                    )

                conn.execute(
                    """
                    INSERT INTO services(
                        name,
                        default_price,
                        worker_percentage,
                        business_percentage
                    )
                    VALUES(?,?,70,30)
                    """,
                    (
                        name,
                        price
                    )
                )

            flash(
                "Servicio creado correctamente.",
                "success"
            )

            return redirect(
                url_for("services")
            )

        except IntegrityError:

            flash(
                "Ya existe un servicio con ese nombre.",
                "danger"
            )

        except ValueError as exc:

            flash(
                str(exc),
                "danger"
            )


    # =====================================================
    # FILTROS
    # =====================================================

    date_from = request.args.get(
        "date_from",
        ""
    ).strip()

    date_to = request.args.get(
        "date_to",
        ""
    ).strip()

    worker_filter = request.args.get(
        "worker",
        ""
    ).strip()


    filters = []
    params = []


    if date_from:

        filters.append(
            "DATE(i.issued_at) >= ?"
        )

        params.append(
            date_from
        )


    if date_to:

        filters.append(
            "DATE(i.issued_at) <= ?"
        )

        params.append(
            date_to
        )


    if worker_filter:

        filters.append(
            "LOWER(TRIM(isi.worker_name)) = LOWER(?)"
        )

        params.append(
            worker_filter
        )


    where_clause = ""

    if filters:

        where_clause = (
            " WHERE "
            + " AND ".join(filters)
        )


    # =====================================================
    # CONSULTAS
    # =====================================================

    with db_conn() as conn:

        # ================================================
        # CATÁLOGO DE SERVICIOS
        # ================================================

        catalog = conn.execute(
            """
            SELECT *
            FROM services
            WHERE active=1
            ORDER BY name
            """
        ).fetchall()

        archived_services = conn.execute(
            """
            SELECT *
            FROM services
            WHERE active=0
            ORDER BY name
            """
        ).fetchall()


        # ================================================
        # HISTORIAL FILTRADO
        # ================================================

        history = conn.execute(
            f"""
            SELECT
                isi.*,
                i.number,
                i.issued_at,
                i.customer_name,
                i.vehicle_plate

            FROM invoice_service_items isi

            JOIN invoices i
                ON i.id = isi.invoice_id

            {where_clause}

            ORDER BY
                i.issued_at DESC,
                isi.id DESC

            LIMIT 1000
            """,
            tuple(params)
        ).fetchall()


        # ================================================
        # RESUMEN GENERAL FILTRADO
        # ================================================

        summary = conn.execute(
            f"""
            SELECT

                COUNT(*) AS service_lines,

                COALESCE(
                    SUM(isi.line_total),
                    0
                ) AS service_total,

                COALESCE(
                    SUM(isi.worker_amount),
                    0
                ) AS worker_total,

                COALESCE(
                    SUM(isi.business_amount),
                    0
                ) AS business_total

            FROM invoice_service_items isi

            JOIN invoices i
                ON i.id = isi.invoice_id

            {where_clause}
            """,
            tuple(params)
        ).fetchone()


        # ================================================
        # DETALLE POR TRABAJADOR Y SERVICIO
        # ================================================

        worker_filters = list(filters)
        worker_params = list(params)

        worker_filters.append(
            "TRIM(isi.worker_name) <> ''"
        )

        worker_where = (
            " WHERE "
            + " AND ".join(
                worker_filters
            )
        )


        workers = conn.execute(
            f"""
            SELECT

                isi.worker_name,

                isi.service_name,

                COUNT(*) AS service_lines,

                COALESCE(
                    SUM(isi.quantity),
                    0
                ) AS quantity_total,

                COALESCE(
                    SUM(isi.line_total),
                    0
                ) AS service_total,

                COALESCE(
                    SUM(isi.worker_amount),
                    0
                ) AS worker_total,

                COALESCE(
                    SUM(isi.business_amount),
                    0
                ) AS business_total

            FROM invoice_service_items isi

            JOIN invoices i
                ON i.id = isi.invoice_id

            {worker_where}

            GROUP BY
                isi.worker_name,
                isi.service_name

            ORDER BY
                isi.worker_name ASC,
                isi.service_name ASC
            """,
            tuple(worker_params)
        ).fetchall()


        # ================================================
        # LISTADO DE TRABAJADORES PARA FILTRO
        # ================================================

        worker_options = conn.execute(
            """
            SELECT DISTINCT
                worker_name

            FROM invoice_service_items

            WHERE
                TRIM(worker_name) <> ''

            ORDER BY
                worker_name
            """
        ).fetchall()


        # ================================================
        # RESUMEN POR DÍA
        # ================================================

        daily = conn.execute(
            f"""
            SELECT

                DATE(i.issued_at) AS service_date,

                COUNT(*) AS service_lines,

                COALESCE(
                    SUM(isi.line_total),
                    0
                ) AS service_total,

                COALESCE(
                    SUM(isi.worker_amount),
                    0
                ) AS worker_total,

                COALESCE(
                    SUM(isi.business_amount),
                    0
                ) AS business_total

            FROM invoice_service_items isi

            JOIN invoices i
                ON i.id = isi.invoice_id

            {where_clause}

            GROUP BY
                DATE(i.issued_at)

            ORDER BY
                DATE(i.issued_at) DESC
            """,
            tuple(params)
        ).fetchall()


    return render_template(
        "services.html",

        services=catalog,

        archived_services=archived_services,

        history=history,

        summary=summary,

        workers=workers,

        worker_options=worker_options,

        daily=daily,

        date_from=date_from,

        date_to=date_to,

        worker_filter=worker_filter
    )

@app.route("/servicios/<int:service_id>/editar", methods=["GET", "POST"])
@role_required("administrador")
def service_edit(service_id: int):
    with db_conn() as conn:
        service = conn.execute("SELECT * FROM services WHERE id=?", (service_id,)).fetchone()
    if not service:
        flash("Servicio no encontrado.", "danger")
        return redirect(url_for("services"))

    if request.method == "POST":
        try:
            name = request.form.get("name", "").strip()
            price = float(parse_decimal(request.form.get("default_price")))
            active = 1 if request.form.get("active") == "1" else 0
            if not name:
                raise ValueError("El nombre del servicio es obligatorio.")
            if price < 0:
                raise ValueError("El valor del servicio no puede ser negativo.")
            with db_conn() as conn:
                conn.execute(
                    """UPDATE services SET name=?,default_price=?,active=?,updated_at=CURRENT_TIMESTAMP
                       WHERE id=?""",
                    (name, price, active, service_id),
                )
            flash("Servicio actualizado.", "success")
            return redirect(url_for("services"))
        except IntegrityError:
            flash("Ya existe otro servicio con ese nombre.", "danger")
        except ValueError as exc:
            flash(str(exc), "danger")
    return render_template("service_form.html", service=service)


@app.route("/servicios/<int:service_id>/eliminar", methods=["POST"])
@role_required("administrador")
def service_delete(service_id: int):
    # El servicio se archiva (active=0) en lugar de borrarse:
    # - las remisiones y liquidaciones guardan su propia copia de nombre,
    #   valores y porcentajes, así que el historial no cambia;
    # - al editar una remisión antigua se vuelve a leer el servicio por id;
    # - los servicios del Excel inicial se recrearían al reiniciar si se borraran.
    with db_conn() as conn:
        service = conn.execute("SELECT * FROM services WHERE id=?", (service_id,)).fetchone()
        if not service:
            flash("Servicio no encontrado.", "danger")
            return redirect(url_for("services"))
        usage = conn.execute(
            "SELECT COUNT(*) AS total FROM invoice_service_items WHERE service_id=?",
            (service_id,),
        ).fetchone()
        conn.execute(
            "UPDATE services SET active=0,updated_at=CURRENT_TIMESTAMP WHERE id=?",
            (service_id,),
        )
    used = int(usage["total"] or 0)
    message = f"Servicio \"{service['name']}\" eliminado del catálogo."
    if used:
        message += f" Aparece en {used} línea(s) de remisiones, que se conservan sin cambios."
    flash(message, "success")
    return redirect(url_for("services"))


@app.route("/servicios/<int:service_id>/restaurar", methods=["POST"])
@role_required("administrador")
def service_restore(service_id: int):
    with db_conn() as conn:
        conn.execute(
            "UPDATE services SET active=1,updated_at=CURRENT_TIMESTAMP WHERE id=?",
            (service_id,),
        )
    flash("Servicio restaurado.", "success")
    return redirect(url_for("services"))


@app.route("/facturas")
@login_required
def invoices():

    with db_conn() as conn:

        rows = conn.execute(
            """
            SELECT
                i.*,
                u.full_name AS user_name,

                CASE

                    WHEN i.payment_method <> 'Crédito'
                    THEN 'PAGADA'

                    WHEN ar.id IS NULL
                    THEN 'PENDIENTE'

                    WHEN
                        COALESCE(
                            (
                                SELECT SUM(arp.amount)
                                FROM accounts_receivable_payments arp
                                WHERE arp.receivable_id = ar.id
                            ),
                            0
                        ) >= ar.original_amount
                    THEN 'PAGADA'

                    WHEN
                        COALESCE(
                            (
                                SELECT SUM(arp.amount)
                                FROM accounts_receivable_payments arp
                                WHERE arp.receivable_id = ar.id
                            ),
                            0
                        ) > 0
                    THEN 'PAGO PARCIAL'

                    ELSE 'PENDIENTE'

                END AS payment_status,

                COALESCE(
                    (
                        SELECT SUM(arp.amount)
                        FROM accounts_receivable_payments arp
                        WHERE arp.receivable_id = ar.id
                    ),
                    0
                ) AS paid_amount,

            CASE
                WHEN i.payment_method <> 'Crédito'
                THEN 0

                WHEN (
                    ar.original_amount
                    -
                    COALESCE(
                        (
                            SELECT SUM(arp.amount)
                            FROM accounts_receivable_payments arp
                            WHERE arp.receivable_id = ar.id
                        ),
                        0
                    )
                ) > 0

                THEN (
                    ar.original_amount
                    -
                    COALESCE(
                        (
                            SELECT SUM(arp.amount)
                            FROM accounts_receivable_payments arp
                            WHERE arp.receivable_id = ar.id
                        ),
                        0
                    )
                )

                ELSE 0
            END AS pending_amount
            FROM invoices i

            LEFT JOIN users u
                ON u.id = i.user_id

            LEFT JOIN accounts_receivable ar
                ON ar.invoice_id = i.id

            ORDER BY
                i.id DESC
            """
        ).fetchall()

    return render_template(
        "invoices.html",
        invoices=rows
    )

@app.route("/facturas/nueva", methods=["GET", "POST"])
@login_required
def invoice_new():

    form_data = request.form if request.method == "POST" else {}
    if request.method == "POST":
        try:
            customer_name = request.form.get("customer_name", "Consumidor final").strip() or "Consumidor final"
            customer_document = request.form.get("customer_document", "").strip()
            customer_phone = request.form.get("customer_phone", "").strip()
            customer_address = request.form.get("customer_address", "").strip()
            vehicle_plate = request.form.get("vehicle_plate", "").strip().upper()
            salesperson_id_raw = request.form.get("salesperson_id", "").strip()
            if not salesperson_id_raw:
                raise ValueError("Selecciona quién está realizando la venta.")
            salesperson_id = int(salesperson_id_raw)
            payment_method = request.form.get("payment_method", "Efectivo").strip()
            notes = request.form.get("notes", "").strip()

            credit_due_date_raw = request.form.get("credit_due_date", "").strip()
            credit_initial_payment = parse_decimal(
                request.form.get("credit_initial_payment", "0")
            )
            credit_initial_payment_method = request.form.get(
                "credit_initial_payment_method", "Efectivo"
            ).strip() or "Efectivo"

            if credit_initial_payment < 0:
                raise ValueError("El abono inicial no puede ser negativo.")

            credit_due_date = None
            if payment_method == "Crédito" and credit_due_date_raw:
                try:
                    credit_due_date = datetime.strptime(
                        credit_due_date_raw, "%Y-%m-%d"
                    ).date()
                except ValueError as exc:
                    raise ValueError("La fecha acordada de pago no es válida.") from exc

            product_ids = request.form.getlist("product_id[]")
            quantities = request.form.getlist("quantity[]")
            prices = request.form.getlist("price[]")

            service_ids = request.form.getlist("service_id[]")
            service_workers = request.form.getlist("service_worker[]")
            service_quantities = request.form.getlist("service_quantity[]")
            service_prices = request.form.getlist("service_price[]")

            # Filtrar filas vacías que puedan quedar en el formulario.
            product_rows_raw = [x for x in zip(product_ids, quantities, prices) if str(x[0]).strip()]
            service_rows_raw = [x for x in zip(service_ids, service_workers, service_quantities, service_prices) if str(x[0]).strip()]
            if not product_rows_raw and not service_rows_raw:
                raise ValueError("Agrega al menos un producto o un servicio a la remisión.")

            with db_conn() as conn:
                settings = get_settings(conn, lock=True)
                salesperson = conn.execute(
                    "SELECT id,name FROM salespeople WHERE id=? AND active=1",
                    (salesperson_id,),
                ).fetchone()
                if not salesperson:
                    raise ValueError("El vendedor seleccionado no está disponible.")
                number = f"{settings['invoice_prefix']}-{settings['next_invoice']:06d}"
                product_lines = []
                service_lines = []
                subtotal = Decimal("0")

                for product_id_raw, qty_raw, price_raw in product_rows_raw:
                    product_id = int(product_id_raw)
                    qty = parse_decimal(qty_raw)
                    price = parse_decimal(price_raw)
                    if qty <= 0 or price < 0:
                        raise ValueError("Revisa cantidades y precios de los productos.")
                    product_sql = "SELECT * FROM products WHERE id=? AND active=1"
                    if db.engine.dialect.name == "postgresql":
                        product_sql += " FOR UPDATE"
                    product = conn.execute(product_sql, (product_id,)).fetchone()
                    if not product:
                        raise ValueError("Uno de los productos ya no está disponible.")
                    if Decimal(str(product["stock"])) < qty:
                        raise ValueError(f"Stock insuficiente para {product['description']}. Disponible: {product['stock']}")
                    line_total = qty * price
                    subtotal += line_total
                    product_lines.append((product, qty, price, line_total))

                for service_id_raw, worker_name, qty_raw, price_raw in service_rows_raw:
                    service_id = int(service_id_raw)
                    worker_name = (worker_name or "").strip()
                    qty = parse_decimal(qty_raw)
                    price = parse_decimal(price_raw)
                    if not worker_name:
                        raise ValueError("Indica el trabajador responsable de cada servicio.")
                    if worker_name not in ALLOWED_MECHANICS:
                        raise ValueError(
                            f"Mecánico no válido: {worker_name}. "
                            f"Selecciona uno de: {', '.join(ALLOWED_MECHANICS)}."
                        )
                    if qty <= 0 or price < 0:
                        raise ValueError("Revisa cantidades y valores de los servicios.")
                    service = conn.execute("SELECT * FROM services WHERE id=? AND active=1", (service_id,)).fetchone()
                    if not service:
                        raise ValueError("Uno de los servicios ya no está disponible.")
                    line_total = qty * price
                    worker_pct = Decimal(str(service["worker_percentage"] or 70))
                    business_pct = Decimal(str(service["business_percentage"] or 30))
                    worker_amount = (line_total * worker_pct / Decimal("100")).quantize(Decimal("0.01"))
                    business_amount = line_total - worker_amount
                    subtotal += line_total
                    service_lines.append((service, worker_name, qty, price, line_total, worker_pct, business_pct, worker_amount, business_amount))

                tax_rate = Decimal(str(settings["tax_rate"] or 0))
                tax = (subtotal * tax_rate / Decimal("100")).quantize(Decimal("0.01"))
                total = subtotal + tax
                issued_date = request.form.get(
                    "issued_date",
                    ""
                ).strip()

                if issued_date:
                    issued_at = f"{issued_date}T12:00:00"
                else:
                    issued_at = datetime.now().isoformat(
                        timespec="seconds"
                    )
                    
                customer_id = None
                if customer_name.lower() != "consumidor final" or customer_document:
                    existing = None
                    if customer_document:
                        existing = conn.execute("SELECT id FROM customers WHERE document=? LIMIT 1", (customer_document,)).fetchone()
                    if existing:
                        customer_id = existing["id"]
                        conn.execute("UPDATE customers SET name=?,phone=?,email=?,address=? WHERE id=?",
                                     (customer_name, customer_phone, request.form.get("customer_email", ""), customer_address, customer_id))
                    else:
                        cur = conn.execute("INSERT INTO customers(document,name,phone,email,address) VALUES(?,?,?,?,?)",
                                           (customer_document, customer_name, customer_phone, request.form.get("customer_email", ""), customer_address))
                        customer_id = cur.lastrowid

                cur = conn.execute(
                    """INSERT INTO invoices(number,customer_id,customer_name,customer_document,customer_phone,
                    customer_address,vehicle_plate,salesperson_id,salesperson_name,issued_at,subtotal,tax,total,payment_method,notes,user_id)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (number, customer_id, customer_name, customer_document, customer_phone, customer_address, vehicle_plate,
                     salesperson["id"], salesperson["name"], issued_at, float(subtotal), float(tax), float(total),
                     payment_method, notes, g.user["id"]),
                )
                invoice_id = cur.lastrowid

                for product, qty, price, line_total in product_lines:
                    conn.execute(
                        """INSERT INTO invoice_items(invoice_id,product_id,code,description,quantity,unit_price,line_total)
                        VALUES(?,?,?,?,?,?,?)""",
                        (invoice_id, product["id"], product["code"], product["description"], float(qty), float(price), float(line_total)),
                    )
                    conn.execute("UPDATE products SET stock=stock-?, updated_at=CURRENT_TIMESTAMP WHERE id=?", (float(qty), product["id"]))
                    conn.execute(
                        "INSERT INTO movements(product_id,movement_type,quantity,reference,notes,created_at,user_id) VALUES(?,?,?,?,?,?,?)",
                        (product["id"], "Venta", float(qty), number, "Venta registrada en remisión", issued_at, g.user["id"]),
                    )

                for service, worker_name, qty, price, line_total, worker_pct, business_pct, worker_amount, business_amount in service_lines:
                    conn.execute(
                        """INSERT INTO invoice_service_items(
                            invoice_id,service_id,service_name,worker_name,quantity,unit_price,line_total,
                            worker_percentage,business_percentage,worker_amount,business_amount)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                        (invoice_id, service["id"], service["name"], worker_name, float(qty), float(price), float(line_total),
                         float(worker_pct), float(business_pct), float(worker_amount), float(business_amount)),
                    )

                # Si la remisión es a crédito, crear automáticamente la cuenta por cobrar.
                if payment_method == "Crédito":
                    if credit_initial_payment > total:
                        raise ValueError(
                            "El abono inicial no puede ser mayor al total de la remisión."
                        )

                    initial_status = "PAGO PARCIAL" if credit_initial_payment > 0 else "PENDIENTE"
                    receivable_cur = conn.execute(
                        """INSERT INTO accounts_receivable(
                            invoice_id,original_amount,due_date,status,created_at
                        ) VALUES(?,?,?,?,?)""",
                        (
                            invoice_id,
                            float(total),
                            credit_due_date,
                            initial_status,
                            datetime.now().isoformat(timespec="seconds"),
                        ),
                    )
                    receivable_id = receivable_cur.lastrowid

                    if credit_initial_payment > 0:
                        conn.execute(
                            """INSERT INTO accounts_receivable_payments(
                                receivable_id,amount,payment_date,payment_method,notes,user_id
                            ) VALUES(?,?,?,?,?,?)""",
                            (
                                receivable_id,
                                float(credit_initial_payment),
                                issued_at[:10],
                                credit_initial_payment_method,
                                "Abono inicial registrado al crear la remisión",
                                g.user["id"],
                            ),
                        )

                conn.execute("UPDATE settings SET next_invoice=next_invoice+1 WHERE id=1")

            generate_invoice_pdf(invoice_id)
            flash(f"Remisión {number} creada correctamente. El inventario de productos fue actualizado.", "success")
            return redirect(url_for("invoice_view", invoice_id=invoice_id))
        except (ValueError, KeyError) as exc:
            flash(str(exc), "danger")

    with db_conn() as conn:
        product_rows = conn.execute(
            "SELECT id,code,description,stock,sale_price,unit FROM products WHERE active=1 AND stock>0 ORDER BY description"
        ).fetchall()
        service_rows = conn.execute(
            "SELECT * FROM services WHERE active=1 ORDER BY name"
        ).fetchall()
        salesperson_rows = conn.execute(
            "SELECT id,name FROM salespeople WHERE active=1 ORDER BY name"
        ).fetchall()
        settings = get_settings(conn)
    return render_template(
        "invoice_form.html",
        products=product_rows,
        services=service_rows,
        salespeople=salesperson_rows,
        settings=settings,
        form_data=form_data
    )

@app.route("/finanzas")
@role_required("administrador")
def finances():
    """Cierre semanal separando productos, servicios, cartera y caja real."""

    today = datetime.now().date()
    week_start = today.fromordinal(today.toordinal() - today.weekday())
    yesterday = today.fromordinal(today.toordinal() - 1)
    previous_week_start = week_start.fromordinal(week_start.toordinal() - 7)
    previous_week_end = week_start.fromordinal(week_start.toordinal() - 1)
    month_start = today.replace(day=1)

    # Por defecto, Finanzas abre en la semana actual.
    date_from = request.args.get("date_from", "").strip() or week_start.isoformat()
    date_to = request.args.get("date_to", "").strip() or today.isoformat()

    for value, label in ((date_from, "Desde"), (date_to, "Hasta")):
        try:
            datetime.strptime(value, "%Y-%m-%d")
        except ValueError:
            flash(f"La fecha '{label}' no es válida.", "danger")
            return redirect(url_for("finances"))

    if date_from > date_to:
        flash("La fecha Desde no puede ser posterior a la fecha Hasta.", "danger")
        return redirect(url_for("finances"))

    def D(value):
        return Decimal(str(value or 0))

    def allocate_amount(amount, product_base, service_base):
        """Distribuye un cobro/saldo entre productos y servicios proporcionalmente."""
        amount = D(amount)
        product_base = D(product_base)
        service_base = D(service_base)
        base = product_base + service_base

        if amount <= 0 or base <= 0:
            return Decimal("0"), Decimal("0")

        product_part = amount * product_base / base
        service_part = amount - product_part
        return product_part, service_part

    invoice_parts_sql = """
        SELECT
            i.id,
            i.number,
            i.issued_at,
            i.customer_name,
            i.payment_method,
            i.subtotal,
            i.tax,
            i.total,
            COALESCE((
                SELECT SUM(ii.line_total)
                FROM invoice_items ii
                WHERE ii.invoice_id = i.id
            ), 0) AS product_total,
            COALESCE((
                SELECT SUM(isi.line_total)
                FROM invoice_service_items isi
                WHERE isi.invoice_id = i.id
            ), 0) AS service_total
        FROM invoices i
    """

    with db_conn() as conn:
        # ======================================================
        # 1) REMISIONES EMITIDAS EN EL PERIODO
        # ======================================================
        issued_rows = conn.execute(
            invoice_parts_sql + """
            WHERE DATE(i.issued_at) >= ?
              AND DATE(i.issued_at) <= ?
            ORDER BY i.issued_at DESC, i.id DESC
            """,
            (date_from, date_to),
        ).fetchall()

        # ======================================================
        # 2) VENTAS NO CRÉDITO COBRADAS EN EL PERIODO
        # Se reconocen como cobradas en la fecha de la remisión.
        # ======================================================
        direct_rows = conn.execute(
            invoice_parts_sql + """
            WHERE i.payment_method <> 'Crédito'
              AND DATE(i.issued_at) >= ?
              AND DATE(i.issued_at) <= ?
            ORDER BY i.issued_at DESC, i.id DESC
            """,
            (date_from, date_to),
        ).fetchall()

        # ======================================================
        # 3) PAGOS DE CARTERA RECIBIDOS EN EL PERIODO
        # El pago se distribuye proporcionalmente entre productos
        # y servicios de la remisión.
        # ======================================================
        credit_payment_rows = conn.execute(
            """
            SELECT
                p.id AS payment_id,
                p.payment_date,
                p.amount,
                p.payment_method,
                p.payment_kind,
                COALESCE((
                    SELECT SUM(sp.amount)
                    FROM accounts_receivable_payments sp
                    WHERE sp.receivable_id = ar.id
                      AND sp.payment_kind = 'SERVICIOS'
                ), 0) AS services_paid_amount,
                i.id AS invoice_id,
                i.number,
                i.customer_name,
                i.subtotal,
                i.tax,
                i.total,
                COALESCE((
                    SELECT SUM(ii.line_total)
                    FROM invoice_items ii
                    WHERE ii.invoice_id = i.id
                ), 0) AS product_total,
                COALESCE((
                    SELECT SUM(isi.line_total)
                    FROM invoice_service_items isi
                    WHERE isi.invoice_id = i.id
                ), 0) AS service_total
            FROM accounts_receivable_payments p
            JOIN accounts_receivable ar ON ar.id = p.receivable_id
            JOIN invoices i ON i.id = ar.invoice_id
            WHERE DATE(p.payment_date) >= ?
              AND DATE(p.payment_date) <= ?
            ORDER BY p.payment_date DESC, p.id DESC
            """,
            (date_from, date_to),
        ).fetchall()

        # ======================================================
        # 4) CARTERA DE REMISIONES EMITIDAS EN EL PERIODO
        # ======================================================
        period_receivable_rows = conn.execute(
            """
            SELECT
                ar.id AS receivable_id,
                ar.original_amount,
                i.id AS invoice_id,
                i.number,
                i.issued_at,
                i.customer_name,
                COALESCE((
                    SELECT SUM(p.amount)
                    FROM accounts_receivable_payments p
                    WHERE p.receivable_id = ar.id
                ), 0) AS paid_amount,
                COALESCE((
                    SELECT SUM(sp.amount)
                    FROM accounts_receivable_payments sp
                    WHERE sp.receivable_id = ar.id
                      AND sp.payment_kind = 'SERVICIOS'
                ), 0) AS services_paid_amount,
                COALESCE((
                    SELECT SUM(ii.line_total)
                    FROM invoice_items ii
                    WHERE ii.invoice_id = i.id
                ), 0) AS product_total,
                COALESCE((
                    SELECT SUM(isi.line_total)
                    FROM invoice_service_items isi
                    WHERE isi.invoice_id = i.id
                ), 0) AS service_total
            FROM accounts_receivable ar
            JOIN invoices i ON i.id = ar.invoice_id
            WHERE DATE(i.issued_at) >= ?
              AND DATE(i.issued_at) <= ?
            ORDER BY i.issued_at DESC, i.id DESC
            """,
            (date_from, date_to),
        ).fetchall()

        # ======================================================
        # 5) CARTERA TOTAL ACTUAL
        # ======================================================
        all_receivable_rows = conn.execute(
            """
            SELECT
                ar.id AS receivable_id,
                ar.original_amount,
                i.id AS invoice_id,
                COALESCE((
                    SELECT SUM(p.amount)
                    FROM accounts_receivable_payments p
                    WHERE p.receivable_id = ar.id
                ), 0) AS paid_amount,
                COALESCE((
                    SELECT SUM(sp.amount)
                    FROM accounts_receivable_payments sp
                    WHERE sp.receivable_id = ar.id
                      AND sp.payment_kind = 'SERVICIOS'
                ), 0) AS services_paid_amount,
                COALESCE((
                    SELECT SUM(ii.line_total)
                    FROM invoice_items ii
                    WHERE ii.invoice_id = i.id
                ), 0) AS product_total,
                COALESCE((
                    SELECT SUM(isi.line_total)
                    FROM invoice_service_items isi
                    WHERE isi.invoice_id = i.id
                ), 0) AS service_total
            FROM accounts_receivable ar
            JOIN invoices i ON i.id = ar.invoice_id
            """
        ).fetchall()

        # ======================================================
        # 6) SERVICIOS Y 70/30 DEL PERIODO
        # Se causan según la fecha de la remisión, sin importar si
        # el cliente ya pagó o no.
        # ======================================================
        service_summary = conn.execute(
            """
            SELECT
                COALESCE(SUM(isi.line_total), 0) AS service_total,
                COALESCE(SUM(isi.worker_amount), 0) AS worker_total,
                COALESCE(SUM(isi.business_amount), 0) AS business_total,
                COUNT(isi.id) AS service_lines
            FROM invoice_service_items isi
            JOIN invoices i ON i.id = isi.invoice_id
            WHERE DATE(i.issued_at) >= ?
              AND DATE(i.issued_at) <= ?
            """,
            (date_from, date_to),
        ).fetchone()

        # ======================================================
        # 7) INGRESOS EXTRA
        # ======================================================
        extra_income_row = conn.execute(
            """
            SELECT COALESCE(SUM(amount), 0) AS total
            FROM extra_income
            WHERE income_date >= ?
              AND income_date <= ?
            """,
            (date_from, date_to),
        ).fetchone()

        extra_income_detail = conn.execute(
            """
            SELECT income_date, category, concept, amount, payment_method
            FROM extra_income
            WHERE income_date >= ?
              AND income_date <= ?
            ORDER BY income_date DESC, id DESC
            """,
            (date_from, date_to),
        ).fetchall()

        # ======================================================
        # 8) GASTOS OPERATIVOS
        # ======================================================
        expenses_row = conn.execute(
            """
            SELECT COALESCE(SUM(amount), 0) AS total
            FROM expenses
            WHERE expense_date >= ?
              AND expense_date <= ?
            """,
            (date_from, date_to),
        ).fetchone()

        # ======================================================
        # 9) PAGOS REALES A PROVEEDORES
        # No son gasto operativo/rentabilidad, pero sí salida de caja.
        # ======================================================
        supplier_payments_row = conn.execute(
            """
            SELECT COALESCE(SUM(amount), 0) AS total
            FROM accounts_payable_payments
            WHERE payment_date >= ?
              AND payment_date <= ?
            """,
            (date_from, date_to),
        ).fetchone()

        # ======================================================
        # 10) LIQUIDACIÓN POR MECÁNICO
        # ======================================================
        worker_detail = conn.execute(
            """
            SELECT
                isi.worker_name,
                COALESCE(SUM(isi.line_total), 0) AS service_total,
                COALESCE(SUM(isi.worker_amount), 0) AS worker_total,
                COALESCE(SUM(isi.business_amount), 0) AS business_total,
                COUNT(*) AS service_lines
            FROM invoice_service_items isi
            JOIN invoices i ON i.id = isi.invoice_id
            WHERE DATE(i.issued_at) >= ?
              AND DATE(i.issued_at) <= ?
              AND TRIM(COALESCE(isi.worker_name, '')) <> ''
            GROUP BY isi.worker_name
            ORDER BY
                CASE isi.worker_name
                    WHEN 'Leo' THEN 1
                    WHEN 'Daniel' THEN 2
                    WHEN 'James' THEN 3
                    ELSE 4
                END,
                isi.worker_name ASC
            """,
            (date_from, date_to),
        ).fetchall()

        # ======================================================
        # 11) GASTOS Y PAGOS A PROVEEDORES - DETALLE
        # ======================================================
        expense_detail = conn.execute(
            """
            SELECT expense_date, category, description, amount, payment_method
            FROM expenses
            WHERE expense_date >= ?
              AND expense_date <= ?
            ORDER BY expense_date DESC, id DESC
            """,
            (date_from, date_to),
        ).fetchall()

        supplier_payment_detail = conn.execute(
            """
            SELECT
                p.payment_date,
                p.amount,
                p.payment_method,
                ap.supplier,
                ap.invoice_number
            FROM accounts_payable_payments p
            JOIN accounts_payable ap ON ap.id = p.payable_id
            WHERE p.payment_date >= ?
              AND p.payment_date <= ?
            ORDER BY p.payment_date DESC, p.id DESC
            """,
            (date_from, date_to),
        ).fetchall()

    # ==========================================================
    # VENTAS EMITIDAS: PRODUCTOS Y SERVICIOS SEPARADOS
    # ==========================================================
    product_sales = sum((D(r["product_total"]) for r in issued_rows), Decimal("0"))
    service_sales = sum((D(r["service_total"]) for r in issued_rows), Decimal("0"))
    tax_issued = sum((D(r["tax"]) for r in issued_rows), Decimal("0"))
    invoice_total_issued = sum((D(r["total"]) for r in issued_rows), Decimal("0"))
    invoice_count = len(issued_rows)

    # ==========================================================
    # COBROS DIRECTOS: SEPARADOS ENTRE PRODUCTOS Y SERVICIOS
    # ==========================================================
    direct_product_collected = Decimal("0")
    direct_service_collected = Decimal("0")
    collected_detail = []

    for row in direct_rows:
        product_part, service_part = allocate_amount(
            row["total"],
            row["product_total"],
            row["service_total"],
        )
        direct_product_collected += product_part
        direct_service_collected += service_part
        collected_detail.append({
            "collection_date": str(row["issued_at"])[:10],
            "source_type": "Venta directa",
            "number": row["number"],
            "customer_name": row["customer_name"],
            "payment_method": row["payment_method"],
            "product_amount": product_part,
            "service_amount": service_part,
            "amount": D(row["total"]),
        })

    # ==========================================================
    # COBROS DE CARTERA: CADA ABONO SE DISTRIBUYE PROPORCIONALMENTE
    # ==========================================================
    credit_product_collected = Decimal("0")
    credit_service_collected = Decimal("0")

    def credit_bases(row):
        """Bases para repartir abonos/saldo: los servicios ya pagados aparte no cuentan."""
        service_base = max(Decimal("0"), D(row["service_total"]) - D(row["services_paid_amount"]))
        return row["product_total"], service_base

    for row in credit_payment_rows:
        if row["payment_kind"] == SERVICES_PAYMENT_KIND:
            # Pago de mano de obra marcado desde la remisión: es 100% servicios.
            product_part, service_part = Decimal("0"), D(row["amount"])
        else:
            product_part, service_part = allocate_amount(row["amount"], *credit_bases(row))
        credit_product_collected += product_part
        credit_service_collected += service_part
        collected_detail.append({
            "collection_date": str(row["payment_date"])[:10],
            "source_type": (
                "Pago de servicios" if row["payment_kind"] == SERVICES_PAYMENT_KIND else "Cobro de cartera"
            ),
            "number": row["number"],
            "customer_name": row["customer_name"],
            "payment_method": row["payment_method"],
            "product_amount": product_part,
            "service_amount": service_part,
            "amount": D(row["amount"]),
        })

    collected_detail.sort(
        key=lambda x: (x["collection_date"], x["number"] or ""),
        reverse=True,
    )

    product_collected = direct_product_collected + credit_product_collected
    service_collected = direct_service_collected + credit_service_collected
    collected_total = product_collected + service_collected
    direct_collected = direct_product_collected + direct_service_collected
    credit_payments = credit_product_collected + credit_service_collected

    # ==========================================================
    # CARTERA DEL PERIODO, SEPARADA ENTRE PRODUCTOS Y SERVICIOS
    # ==========================================================
    pending_product_period = Decimal("0")
    pending_service_period = Decimal("0")
    period_receivables = []

    for row in period_receivable_rows:
        original = D(row["original_amount"])
        paid = D(row["paid_amount"])
        balance = max(Decimal("0"), original - paid)
        product_pending, service_pending = allocate_amount(balance, *credit_bases(row))
        pending_product_period += product_pending
        pending_service_period += service_pending

        period_receivables.append({
            "invoice_id": row["invoice_id"],
            "number": row["number"],
            "issued_at": row["issued_at"],
            "customer_name": row["customer_name"],
            "original_amount": original,
            "paid_amount": paid,
            "balance": balance,
            "product_pending": product_pending,
            "service_pending": service_pending,
        })

    pending_period = pending_product_period + pending_service_period

    # ==========================================================
    # CARTERA TOTAL, SEPARADA ENTRE PRODUCTOS Y SERVICIOS
    # ==========================================================
    total_product_receivables = Decimal("0")
    total_service_receivables = Decimal("0")

    for row in all_receivable_rows:
        original = D(row["original_amount"])
        paid = D(row["paid_amount"])
        balance = max(Decimal("0"), original - paid)
        product_pending, service_pending = allocate_amount(balance, *credit_bases(row))
        total_product_receivables += product_pending
        total_service_receivables += service_pending

    total_receivables = total_product_receivables + total_service_receivables

    # ==========================================================
    # SERVICIOS / CAJA / SALIDAS
    # ==========================================================
    service_worker = D(service_summary["worker_total"])
    service_business = D(service_summary["business_total"])
    extra_income_total = D(extra_income_row["total"])
    expenses_total = D(expenses_row["total"])
    supplier_payments_total = D(supplier_payments_row["total"])

    # Liquidación individual del 70% por mecánico.
    # Se construye con los valores ya guardados en invoice_service_items,
    # por lo que respeta exactamente el porcentaje del servicio.
    mechanic_totals = {
        name: {
            "service_total": Decimal("0"),
            "worker_total": Decimal("0"),
            "business_total": Decimal("0"),
            "service_lines": 0,
        }
        for name in ALLOWED_MECHANICS
    }

    for row in worker_detail:
        worker_name = str(row["worker_name"] or "").strip()
        if worker_name not in mechanic_totals:
            mechanic_totals[worker_name] = {
                "service_total": Decimal("0"),
                "worker_total": Decimal("0"),
                "business_total": Decimal("0"),
                "service_lines": 0,
            }

        mechanic_totals[worker_name] = {
            "service_total": D(row["service_total"]),
            "worker_total": D(row["worker_total"]),
            "business_total": D(row["business_total"]),
            "service_lines": int(row["service_lines"] or 0),
        }

    leo_worker_total = mechanic_totals["Leo"]["worker_total"]
    daniel_worker_total = mechanic_totals["Daniel"]["worker_total"]
    james_worker_total = mechanic_totals["James"]["worker_total"]

    total_inflows = product_collected + service_collected + extra_income_total
    total_outflows = service_worker + expenses_total + supplier_payments_total
    net_cash_flow = total_inflows - total_outflows

    # ==========================================================
    # RESUMEN DIARIO: PRODUCTOS Y SERVICIOS SEPARADOS
    # ==========================================================
    daily_map = {}

    def ensure_day(raw_date):
        key = str(raw_date)[:10]
        if key not in daily_map:
            daily_map[key] = {
                "activity_date": key,
                "product_collected": Decimal("0"),
                "service_collected": Decimal("0"),
                "extra_income": Decimal("0"),
                "worker_total": Decimal("0"),
                "expenses": Decimal("0"),
                "supplier_payments": Decimal("0"),
                "net_cash_flow": Decimal("0"),
            }
        return daily_map[key]

    for row in collected_detail:
        day = ensure_day(row["collection_date"])
        day["product_collected"] += D(row["product_amount"])
        day["service_collected"] += D(row["service_amount"])

    for row in extra_income_detail:
        ensure_day(row["income_date"])["extra_income"] += D(row["amount"])

    for row in worker_detail:
        # El detalle por mecánico está agregado por persona, no por día.
        # Para el resumen diario hacemos una consulta basada en remisiones abajo.
        pass

    # 70% por fecha de servicio/remisión.
    with db_conn() as conn:
        daily_workers = conn.execute(
            """
            SELECT DATE(i.issued_at) AS activity_date,
                   COALESCE(SUM(isi.worker_amount), 0) AS worker_total
            FROM invoice_service_items isi
            JOIN invoices i ON i.id = isi.invoice_id
            WHERE DATE(i.issued_at) >= ?
              AND DATE(i.issued_at) <= ?
            GROUP BY DATE(i.issued_at)
            """,
            (date_from, date_to),
        ).fetchall()

    for row in daily_workers:
        ensure_day(row["activity_date"])["worker_total"] += D(row["worker_total"])

    for row in expense_detail:
        ensure_day(row["expense_date"])["expenses"] += D(row["amount"])

    for row in supplier_payment_detail:
        ensure_day(row["payment_date"])["supplier_payments"] += D(row["amount"])

    daily = []
    for key in sorted(daily_map.keys(), reverse=True):
        row = daily_map[key]
        row["collected"] = row["product_collected"] + row["service_collected"]
        row["net_cash_flow"] = (
            row["product_collected"]
            + row["service_collected"]
            + row["extra_income"]
            - row["worker_total"]
            - row["expenses"]
            - row["supplier_payments"]
        )
        daily.append(row)

    return render_template(
        "finances.html",
        date_from=date_from,
        date_to=date_to,

        # Ventas emitidas
        product_sales=product_sales,
        service_sales=service_sales,
        tax_issued=tax_issued,
        invoice_total_issued=invoice_total_issued,
        invoice_count=invoice_count,

        # Cobros separados
        direct_product_collected=direct_product_collected,
        direct_service_collected=direct_service_collected,
        credit_product_collected=credit_product_collected,
        credit_service_collected=credit_service_collected,
        product_collected=product_collected,
        service_collected=service_collected,
        direct_collected=direct_collected,
        credit_payments=credit_payments,
        collected_total=collected_total,

        # Cartera separada
        pending_product_period=pending_product_period,
        pending_service_period=pending_service_period,
        pending_period=pending_period,
        total_product_receivables=total_product_receivables,
        total_service_receivables=total_service_receivables,
        total_receivables=total_receivables,

        # Servicios 70/30
        service_worker=service_worker,
        service_business=service_business,

        # Liquidación individual por mecánico
        mechanic_totals=mechanic_totals,
        leo_worker_total=leo_worker_total,
        daniel_worker_total=daniel_worker_total,
        james_worker_total=james_worker_total,

        # Caja
        extra_income_total=extra_income_total,
        expenses_total=expenses_total,
        supplier_payments_total=supplier_payments_total,
        total_inflows=total_inflows,
        total_outflows=total_outflows,
        net_cash_flow=net_cash_flow,

        # Detalles
        collected_detail=collected_detail,
        period_receivables=period_receivables,
        worker_detail=worker_detail,
        extra_income_detail=extra_income_detail,
        expense_detail=expense_detail,
        supplier_payment_detail=supplier_payment_detail,
        daily=daily,

        # Accesos rápidos
        today_iso=today.isoformat(),
        yesterday_iso=yesterday.isoformat(),
        week_start_iso=week_start.isoformat(),
        previous_week_start_iso=previous_week_start.isoformat(),
        previous_week_end_iso=previous_week_end.isoformat(),
        month_start_iso=month_start.isoformat(),
    )


def _receivable_display_status(row: dict[str, Any], today=None) -> str:
    """Calcula el estado visible de una cuenta por cobrar."""
    today = today or datetime.now().date()
    original = Decimal(str(row.get("original_amount") or 0))
    paid = Decimal(str(row.get("paid_amount") or 0))
    balance = original - paid

    if balance <= 0:
        return "PAGADA"

    due_raw = row.get("due_date")
    if due_raw:
        if hasattr(due_raw, "date") and not isinstance(due_raw, str):
            try:
                due = due_raw.date()
            except Exception:
                due = due_raw
        elif isinstance(due_raw, str):
            try:
                due = datetime.strptime(due_raw[:10], "%Y-%m-%d").date()
            except ValueError:
                due = None
        else:
            due = due_raw
        if due and due < today:
            return "VENCIDA"

    if paid > 0:
        return "PAGO PARCIAL"
    return "PENDIENTE"


@app.route("/cuentas-por-cobrar")
@login_required
def accounts_receivable():
    status_filter = request.args.get("status", "ABIERTAS").strip().upper()
    q = request.args.get("q", "").strip()

    with db_conn() as conn:
        rows = conn.execute(
            """
            SELECT
                ar.id, ar.invoice_id, ar.original_amount, ar.due_date,
                ar.status, ar.created_at, ar.closed_at,
                i.number, i.issued_at, i.customer_name, i.customer_document,
                i.customer_phone, i.vehicle_plate,
                COALESCE((
                    SELECT SUM(p.amount)
                    FROM accounts_receivable_payments p
                    WHERE p.receivable_id = ar.id
                ), 0) AS paid_amount,
                COALESCE((
                    SELECT SUM(sp.amount)
                    FROM accounts_receivable_payments sp
                    WHERE sp.receivable_id = ar.id
                      AND sp.payment_kind = 'SERVICIOS'
                ), 0) AS services_paid_amount
            FROM accounts_receivable ar
            JOIN invoices i ON i.id = ar.invoice_id
            ORDER BY i.issued_at ASC, ar.id ASC
            """
        ).fetchall()

    accounts = []
    total_pending = Decimal("0")
    total_overdue = Decimal("0")
    total_without_date = Decimal("0")

    q_lower = q.lower()
    for row in rows:
        account = dict(row)
        original = Decimal(str(account.get("original_amount") or 0))
        paid = Decimal(str(account.get("paid_amount") or 0))
        balance = max(Decimal("0"), original - paid)
        display_status = _receivable_display_status(account)

        account["paid_amount"] = paid
        account["balance"] = balance
        account["display_status"] = display_status

        if balance > 0:
            total_pending += balance
            if display_status == "VENCIDA":
                total_overdue += balance
            if not account.get("due_date"):
                total_without_date += balance

        if q_lower:
            haystack = " ".join(
                str(account.get(key) or "")
                for key in ("number", "customer_name", "customer_document", "customer_phone", "vehicle_plate")
            ).lower()
            if q_lower not in haystack:
                continue

        if status_filter == "ABIERTAS" and display_status == "PAGADA":
            continue
        if status_filter not in ("", "TODAS", "ABIERTAS") and display_status != status_filter:
            continue

        accounts.append(account)

    return render_template(
        "accounts_receivable.html",
        accounts=accounts,
        status_filter=status_filter,
        q=q,
        total_pending=total_pending,
        total_overdue=total_overdue,
        total_without_date=total_without_date,
    )


@app.route("/cuentas-por-cobrar/<int:account_id>/pago", methods=["POST"])
@login_required
def accounts_receivable_payment(account_id: int):
    try:
        amount = parse_decimal(request.form.get("amount", "0"))
        payment_method = request.form.get("payment_method", "Efectivo").strip() or "Efectivo"
        payment_date_raw = request.form.get("payment_date", "").strip()
        notes = request.form.get("notes", "").strip()

        if amount <= 0:
            raise ValueError("El valor del pago debe ser mayor a cero.")

        if payment_date_raw:
            try:
                payment_date = datetime.strptime(payment_date_raw, "%Y-%m-%d").date()
            except ValueError as exc:
                raise ValueError("La fecha del pago no es válida.") from exc
        else:
            payment_date = datetime.now().date()

        with db_conn() as conn:
            account = conn.execute(
                """
                SELECT ar.*,
                       COALESCE((
                           SELECT SUM(p.amount)
                           FROM accounts_receivable_payments p
                           WHERE p.receivable_id = ar.id
                       ), 0) AS paid_amount
                FROM accounts_receivable ar
                WHERE ar.id=?
                """,
                (account_id,),
            ).fetchone()

            if not account:
                raise ValueError("La cuenta por cobrar no existe.")

            original = Decimal(str(account["original_amount"] or 0))
            already_paid = Decimal(str(account["paid_amount"] or 0))
            balance = original - already_paid

            if balance <= 0:
                raise ValueError("Esta cuenta ya se encuentra pagada.")
            if amount > balance:
                raise ValueError(f"El pago supera el saldo pendiente de {money(balance)}.")

            conn.execute(
                """INSERT INTO accounts_receivable_payments(
                    receivable_id,amount,payment_date,payment_method,notes,user_id
                ) VALUES(?,?,?,?,?,?)""",
                (account_id, float(amount), payment_date, payment_method, notes, g.user["id"]),
            )

            new_balance = balance - amount
            new_status = "PAGADA" if new_balance <= 0 else "PAGO PARCIAL"
            closed_at = datetime.now().isoformat(timespec="seconds") if new_balance <= 0 else None
            conn.execute(
                "UPDATE accounts_receivable SET status=?,closed_at=? WHERE id=?",
                (new_status, closed_at, account_id),
            )

        flash("Pago registrado correctamente.", "success")
    except ValueError as exc:
        flash(str(exc), "danger")

    return redirect(url_for("accounts_receivable"))


@app.route("/cuentas-por-cobrar/<int:account_id>/pagar-total", methods=["POST"])
@login_required
def accounts_receivable_pay_full(account_id: int):
    try:
        payment_method = request.form.get("payment_method", "Efectivo").strip() or "Efectivo"
        payment_date_raw = request.form.get("payment_date", "").strip()
        notes = request.form.get("notes", "").strip() or "Pago total de la cuenta"

        if payment_date_raw:
            try:
                payment_date = datetime.strptime(payment_date_raw, "%Y-%m-%d").date()
            except ValueError as exc:
                raise ValueError("La fecha del pago no es válida.") from exc
        else:
            payment_date = datetime.now().date()

        with db_conn() as conn:
            account = conn.execute(
                """
                SELECT ar.*,
                       COALESCE((
                           SELECT SUM(p.amount)
                           FROM accounts_receivable_payments p
                           WHERE p.receivable_id = ar.id
                       ), 0) AS paid_amount
                FROM accounts_receivable ar
                WHERE ar.id=?
                """,
                (account_id,),
            ).fetchone()

            if not account:
                raise ValueError("La cuenta por cobrar no existe.")

            original = Decimal(str(account["original_amount"] or 0))
            paid = Decimal(str(account["paid_amount"] or 0))
            balance = original - paid

            if balance > 0:
                conn.execute(
                    """INSERT INTO accounts_receivable_payments(
                        receivable_id,amount,payment_date,payment_method,notes,user_id
                    ) VALUES(?,?,?,?,?,?)""",
                    (account_id, float(balance), payment_date, payment_method, notes, g.user["id"]),
                )

            conn.execute(
                "UPDATE accounts_receivable SET status='PAGADA',closed_at=? WHERE id=?",
                (datetime.now().isoformat(timespec="seconds"), account_id),
            )

        flash("Cuenta marcada como pagada correctamente.", "success")
    except ValueError as exc:
        flash(str(exc), "danger")

    return redirect(url_for("accounts_receivable"))


@app.route("/cuentas-por-cobrar/<int:account_id>/historial")
@login_required
def accounts_receivable_history(account_id: int):
    with db_conn() as conn:
        account = conn.execute(
            """
            SELECT ar.*, i.number, i.issued_at, i.customer_name, i.customer_phone, i.vehicle_plate
            FROM accounts_receivable ar
            JOIN invoices i ON i.id=ar.invoice_id
            WHERE ar.id=?
            """,
            (account_id,),
        ).fetchone()
        payments = conn.execute(
            """
            SELECT p.*, u.full_name AS user_name
            FROM accounts_receivable_payments p
            LEFT JOIN users u ON u.id=p.user_id
            WHERE p.receivable_id=?
            ORDER BY p.payment_date DESC, p.id DESC
            """,
            (account_id,),
        ).fetchall()

    if not account:
        flash("Cuenta por cobrar no encontrada.", "danger")
        return redirect(url_for("accounts_receivable"))

    total_paid = sum((Decimal(str(p["amount"] or 0)) for p in payments), Decimal("0"))
    balance = max(Decimal("0"), Decimal(str(account["original_amount"] or 0)) - total_paid)
    return render_template(
        "accounts_receivable_history.html",
        account=account,
        payments=payments,
        total_paid=total_paid,
        balance=balance,
    )

# ============================================================
# CUENTAS POR PAGAR / FACTURAS PENDIENTES
# ============================================================


def _parse_payable_date(value):
    """Convierte distintos formatos devueltos por SQLite/PostgreSQL a date."""

    if not value:
        return None

    if hasattr(value, "date") and not isinstance(value, str):
        try:
            return value.date()
        except Exception:
            pass

    if isinstance(value, str):
        try:
            return datetime.strptime(
                value[:10],
                "%Y-%m-%d"
            ).date()
        except ValueError:
            return None

    return value


def _payable_display_status(row, today=None):
    """
    Estado visual tipo semáforo.

    PAGADA             -> pago completo
    VENCIDA            -> fecha límite ya pasó
    PROXIMA A VENCER   -> faltan entre 0 y 7 días
    PENDIENTE          -> faltan más de 7 días
    """

    today = today or datetime.now().date()

    original_amount = Decimal(
        str(
            row.get(
                "original_amount"
            )
            or
            0
        )
    )

    paid_amount = Decimal(
        str(
            row.get(
                "paid_amount"
            )
            or
            0
        )
    )

    balance = (
        original_amount
        -
        paid_amount
    )

    if balance <= Decimal("0"):
        return "PAGADA"

    due_date = _parse_payable_date(
        row.get(
            "due_date"
        )
    )

    if not due_date:
        return "PENDIENTE"

    days_remaining = (
        due_date
        -
        today
    ).days

    if days_remaining < 0:
        return "VENCIDA"

    if days_remaining <= 7:
        return "PROXIMA A VENCER"

    return "PENDIENTE"


def _payable_days_remaining(row, today=None):

    today = today or datetime.now().date()

    due_date = _parse_payable_date(
        row.get(
            "due_date"
        )
    )

    if not due_date:
        return None

    return (
        due_date
        -
        today
    ).days


# ============================================================
# LISTADO
# ============================================================

@app.route("/facturas-pendientes")
@role_required("administrador")
def accounts_payable():

    status_filter = (
        request.args.get(
            "status",
            "ABIERTAS"
        )
        .strip()
        .upper()
    )

    q = (
        request.args.get(
            "q",
            ""
        )
        .strip()
    )

    with db_conn() as conn:

        rows = conn.execute(
            """
            SELECT
                ap.*,

                COALESCE(
                    (
                        SELECT SUM(p.amount)
                        FROM accounts_payable_payments p
                        WHERE p.payable_id = ap.id
                    ),
                    0
                ) AS paid_amount

            FROM accounts_payable ap

            ORDER BY
                ap.due_date ASC,
                ap.id DESC
            """
        ).fetchall()


    accounts = []

    total_open = Decimal("0")
    total_pending = Decimal("0")
    total_upcoming = Decimal("0")
    total_overdue = Decimal("0")

    count_pending = 0
    count_upcoming = 0
    count_overdue = 0

    q_lower = q.lower()


    for row in rows:

        account = dict(row)

        original_amount = Decimal(
            str(
                account.get(
                    "original_amount"
                )
                or
                0
            )
        )

        paid_amount = Decimal(
            str(
                account.get(
                    "paid_amount"
                )
                or
                0
            )
        )

        balance = max(
            Decimal("0"),
            original_amount
            -
            paid_amount
        )

        display_status = (
            _payable_display_status(
                account
            )
        )

        days_remaining = (
            _payable_days_remaining(
                account
            )
        )


        account["paid_amount"] = (
            paid_amount
        )

        account["balance"] = (
            balance
        )

        account["display_status"] = (
            display_status
        )

        account["days_remaining"] = (
            days_remaining
        )


        if balance > 0:

            total_open += balance

            if display_status == "PENDIENTE":

                total_pending += balance
                count_pending += 1

            elif display_status == "PROXIMA A VENCER":

                total_upcoming += balance
                count_upcoming += 1

            elif display_status == "VENCIDA":

                total_overdue += balance
                count_overdue += 1


        # Buscador
        if q_lower:

            searchable = " ".join(
                str(
                    account.get(key)
                    or
                    ""
                )
                for key in (
                    "supplier",
                    "invoice_number",
                    "description",
                    "notes",
                )
            ).lower()

            if q_lower not in searchable:
                continue


        # Filtros
        if (
            status_filter == "ABIERTAS"
            and
            display_status == "PAGADA"
        ):
            continue


        if (
            status_filter
            not in (
                "",
                "TODAS",
                "ABIERTAS"
            )
            and
            display_status
            !=
            status_filter
        ):
            continue


        accounts.append(
            account
        )


    return render_template(
        "accounts_payable.html",

        accounts=accounts,

        status_filter=status_filter,

        q=q,

        total_open=total_open,

        total_pending=total_pending,

        total_upcoming=total_upcoming,

        total_overdue=total_overdue,

        count_pending=count_pending,

        count_upcoming=count_upcoming,

        count_overdue=count_overdue,
    )


# ============================================================
# CREAR FACTURA
# ============================================================

@app.route(
    "/facturas-pendientes/nueva",
    methods=["GET", "POST"]
)
@role_required("administrador")
def account_payable_new():

    if request.method == "POST":

        try:

            supplier = (
                request.form.get(
                    "supplier",
                    ""
                )
                .strip()
            )

            invoice_number = (
                request.form.get(
                    "invoice_number",
                    ""
                )
                .strip()
            )

            purchase_date_raw = (
                request.form.get(
                    "purchase_date",
                    ""
                )
                .strip()
            )

            due_date_raw = (
                request.form.get(
                    "due_date",
                    ""
                )
                .strip()
            )

            amount = parse_decimal(
                request.form.get(
                    "original_amount",
                    "0"
                )
            )

            description = (
                request.form.get(
                    "description",
                    ""
                )
                .strip()
            )

            notes = (
                request.form.get(
                    "notes",
                    ""
                )
                .strip()
            )


            if not supplier:

                raise ValueError(
                    "El proveedor es obligatorio."
                )


            if not invoice_number:

                raise ValueError(
                    "El número de factura es obligatorio."
                )


            if not purchase_date_raw:

                raise ValueError(
                    "La fecha de compra es obligatoria."
                )


            if not due_date_raw:

                raise ValueError(
                    "La fecha límite de pago es obligatoria."
                )


            try:

                purchase_date = datetime.strptime(
                    purchase_date_raw,
                    "%Y-%m-%d"
                ).date()

            except ValueError as exc:

                raise ValueError(
                    "La fecha de compra no es válida."
                ) from exc


            try:

                due_date = datetime.strptime(
                    due_date_raw,
                    "%Y-%m-%d"
                ).date()

            except ValueError as exc:

                raise ValueError(
                    "La fecha límite de pago no es válida."
                ) from exc


            if due_date < purchase_date:

                raise ValueError(
                    "La fecha límite de pago no puede ser anterior a la fecha de compra."
                )


            if amount <= Decimal("0"):

                raise ValueError(
                    "El valor de la factura debe ser mayor a cero."
                )


            with db_conn() as conn:

                duplicate = conn.execute(
                    """
                    SELECT id
                    FROM accounts_payable
                    WHERE LOWER(supplier)=LOWER(?)
                    AND LOWER(invoice_number)=LOWER(?)
                    """,
                    (
                        supplier,
                        invoice_number
                    )
                ).fetchone()


                if duplicate:

                    raise ValueError(
                        "Ya existe una factura con ese número para este proveedor."
                    )


                conn.execute(
                    """
                    INSERT INTO accounts_payable(
                        supplier,
                        invoice_number,
                        purchase_date,
                        due_date,
                        original_amount,
                        description,
                        notes,
                        status,
                        user_id,
                        created_at,
                        updated_at
                    )
                    VALUES(
                        ?,?,?,?,?,?,?,?,
                        ?,?,?
                    )
                    """,
                    (
                        supplier,
                        invoice_number,
                        purchase_date,
                        due_date,
                        float(amount),
                        description,
                        notes,
                        "PENDIENTE",
                        g.user["id"],
                        datetime.now().isoformat(
                            timespec="seconds"
                        ),
                        datetime.now().isoformat(
                            timespec="seconds"
                        ),
                    )
                )


            flash(
                "Factura pendiente registrada correctamente.",
                "success"
            )

            return redirect(
                url_for(
                    "accounts_payable"
                )
            )


        except ValueError as exc:

            flash(
                str(exc),
                "danger"
            )


    return render_template(
        "account_payable_form.html",
        account=None
    )


# ============================================================
# EDITAR FACTURA
# ============================================================

@app.route(
    "/facturas-pendientes/<int:account_id>/editar",
    methods=["GET", "POST"]
)
@role_required("administrador")
def account_payable_edit(account_id):

    with db_conn() as conn:

        account = conn.execute(
            """
            SELECT ap.*,

                   COALESCE(
                       (
                           SELECT SUM(p.amount)
                           FROM accounts_payable_payments p
                           WHERE p.payable_id=ap.id
                       ),
                       0
                   ) AS paid_amount

            FROM accounts_payable ap

            WHERE ap.id=?
            """,
            (
                account_id,
            )
        ).fetchone()


    if not account:

        flash(
            "Factura pendiente no encontrada.",
            "danger"
        )

        return redirect(
            url_for(
                "accounts_payable"
            )
        )


    if request.method == "POST":

        try:

            supplier = (
                request.form.get(
                    "supplier",
                    ""
                )
                .strip()
            )

            invoice_number = (
                request.form.get(
                    "invoice_number",
                    ""
                )
                .strip()
            )

            purchase_date_raw = (
                request.form.get(
                    "purchase_date",
                    ""
                )
                .strip()
            )

            due_date_raw = (
                request.form.get(
                    "due_date",
                    ""
                )
                .strip()
            )

            amount = parse_decimal(
                request.form.get(
                    "original_amount",
                    "0"
                )
            )

            description = (
                request.form.get(
                    "description",
                    ""
                )
                .strip()
            )

            notes = (
                request.form.get(
                    "notes",
                    ""
                )
                .strip()
            )


            if not supplier:
                raise ValueError(
                    "El proveedor es obligatorio."
                )


            if not invoice_number:
                raise ValueError(
                    "El número de factura es obligatorio."
                )


            try:

                purchase_date = datetime.strptime(
                    purchase_date_raw,
                    "%Y-%m-%d"
                ).date()

                due_date = datetime.strptime(
                    due_date_raw,
                    "%Y-%m-%d"
                ).date()

            except ValueError as exc:

                raise ValueError(
                    "Las fechas ingresadas no son válidas."
                ) from exc


            if due_date < purchase_date:

                raise ValueError(
                    "La fecha límite no puede ser anterior a la fecha de compra."
                )


            if amount <= 0:

                raise ValueError(
                    "El valor de la factura debe ser mayor a cero."
                )


            paid_amount = Decimal(
                str(
                    account.get(
                        "paid_amount"
                    )
                    or
                    0
                )
            )


            if amount < paid_amount:

                raise ValueError(
                    "El valor de la factura no puede ser menor al valor que ya fue pagado."
                )


            with db_conn() as conn:

                duplicate = conn.execute(
                    """
                    SELECT id
                    FROM accounts_payable

                    WHERE LOWER(supplier)=LOWER(?)
                    AND LOWER(invoice_number)=LOWER(?)
                    AND id<>?
                    """,
                    (
                        supplier,
                        invoice_number,
                        account_id
                    )
                ).fetchone()


                if duplicate:

                    raise ValueError(
                        "Ya existe otra factura con ese número para este proveedor."
                    )


                balance = (
                    amount
                    -
                    paid_amount
                )


                new_status = (
                    "PAGADA"
                    if balance <= 0
                    else
                    (
                        "PAGO PARCIAL"
                        if paid_amount > 0
                        else
                        "PENDIENTE"
                    )
                )


                closed_at = (
                    datetime.now().isoformat(
                        timespec="seconds"
                    )
                    if new_status == "PAGADA"
                    else
                    None
                )


                conn.execute(
                    """
                    UPDATE accounts_payable

                    SET
                        supplier=?,
                        invoice_number=?,
                        purchase_date=?,
                        due_date=?,
                        original_amount=?,
                        description=?,
                        notes=?,
                        status=?,
                        closed_at=?,
                        updated_at=CURRENT_TIMESTAMP

                    WHERE id=?
                    """,
                    (
                        supplier,
                        invoice_number,
                        purchase_date,
                        due_date,
                        float(amount),
                        description,
                        notes,
                        new_status,
                        closed_at,
                        account_id
                    )
                )


            flash(
                "Factura actualizada correctamente.",
                "success"
            )

            return redirect(
                url_for(
                    "accounts_payable"
                )
            )


        except ValueError as exc:

            flash(
                str(exc),
                "danger"
            )


    return render_template(
        "account_payable_form.html",
        account=account
    )


# ============================================================
# PAGO PARCIAL
# ============================================================

@app.route(
    "/facturas-pendientes/<int:account_id>/pago",
    methods=["POST"]
)
@role_required("administrador")
def account_payable_payment(account_id):

    try:

        amount = parse_decimal(
            request.form.get(
                "amount",
                "0"
            )
        )

        payment_method = (
            request.form.get(
                "payment_method",
                "Transferencia"
            )
            .strip()
            or
            "Transferencia"
        )

        payment_date_raw = (
            request.form.get(
                "payment_date",
                ""
            )
            .strip()
        )

        notes = (
            request.form.get(
                "notes",
                ""
            )
            .strip()
        )


        if amount <= 0:

            raise ValueError(
                "El valor del pago debe ser mayor a cero."
            )


        if payment_date_raw:

            try:

                payment_date = datetime.strptime(
                    payment_date_raw,
                    "%Y-%m-%d"
                ).date()

            except ValueError as exc:

                raise ValueError(
                    "La fecha del pago no es válida."
                ) from exc

        else:

            payment_date = (
                datetime.now().date()
            )


        with db_conn() as conn:

            account = conn.execute(
                """
                SELECT
                    ap.*,

                    COALESCE(
                        (
                            SELECT SUM(p.amount)
                            FROM accounts_payable_payments p
                            WHERE p.payable_id=ap.id
                        ),
                        0
                    ) AS paid_amount

                FROM accounts_payable ap

                WHERE ap.id=?
                """,
                (
                    account_id,
                )
            ).fetchone()


            if not account:

                raise ValueError(
                    "La factura pendiente no existe."
                )


            original = Decimal(
                str(
                    account["original_amount"]
                    or
                    0
                )
            )

            already_paid = Decimal(
                str(
                    account["paid_amount"]
                    or
                    0
                )
            )

            balance = (
                original
                -
                already_paid
            )


            if balance <= 0:

                raise ValueError(
                    "Esta factura ya se encuentra pagada."
                )


            if amount > balance:

                raise ValueError(
                    f"El pago supera el saldo pendiente de {money(balance)}."
                )


            conn.execute(
                """
                INSERT INTO accounts_payable_payments(
                    payable_id,
                    amount,
                    payment_date,
                    payment_method,
                    notes,
                    user_id
                )
                VALUES(?,?,?,?,?,?)
                """,
                (
                    account_id,
                    float(amount),
                    payment_date,
                    payment_method,
                    notes,
                    g.user["id"]
                )
            )


            new_balance = (
                balance
                -
                amount
            )


            new_status = (
                "PAGADA"
                if new_balance <= 0
                else
                "PAGO PARCIAL"
            )


            closed_at = (
                datetime.now().isoformat(
                    timespec="seconds"
                )
                if new_balance <= 0
                else
                None
            )


            conn.execute(
                """
                UPDATE accounts_payable
                SET
                    status=?,
                    closed_at=?,
                    updated_at=CURRENT_TIMESTAMP
                WHERE id=?
                """,
                (
                    new_status,
                    closed_at,
                    account_id
                )
            )


        flash(
            "Pago registrado correctamente.",
            "success"
        )


    except ValueError as exc:

        flash(
            str(exc),
            "danger"
        )


    return redirect(
        url_for(
            "accounts_payable"
        )
    )


# ============================================================
# PAGAR TODO
# ============================================================

@app.route(
    "/facturas-pendientes/<int:account_id>/pagar-total",
    methods=["POST"]
)
@role_required("administrador")
def account_payable_pay_full(account_id):

    try:

        payment_method = (
            request.form.get(
                "payment_method",
                "Transferencia"
            )
            .strip()
            or
            "Transferencia"
        )

        payment_date_raw = (
            request.form.get(
                "payment_date",
                ""
            )
            .strip()
        )

        notes = (
            request.form.get(
                "notes",
                ""
            )
            .strip()
            or
            "Pago total de factura"
        )


        if payment_date_raw:

            payment_date = datetime.strptime(
                payment_date_raw,
                "%Y-%m-%d"
            ).date()

        else:

            payment_date = (
                datetime.now().date()
            )


        with db_conn() as conn:

            account = conn.execute(
                """
                SELECT
                    ap.*,

                    COALESCE(
                        (
                            SELECT SUM(p.amount)
                            FROM accounts_payable_payments p
                            WHERE p.payable_id=ap.id
                        ),
                        0
                    ) AS paid_amount

                FROM accounts_payable ap

                WHERE ap.id=?
                """,
                (
                    account_id,
                )
            ).fetchone()


            if not account:

                raise ValueError(
                    "Factura pendiente no encontrada."
                )


            original = Decimal(
                str(
                    account["original_amount"]
                    or
                    0
                )
            )

            already_paid = Decimal(
                str(
                    account["paid_amount"]
                    or
                    0
                )
            )

            balance = (
                original
                -
                already_paid
            )


            if balance > 0:

                conn.execute(
                    """
                    INSERT INTO accounts_payable_payments(
                        payable_id,
                        amount,
                        payment_date,
                        payment_method,
                        notes,
                        user_id
                    )
                    VALUES(?,?,?,?,?,?)
                    """,
                    (
                        account_id,
                        float(balance),
                        payment_date,
                        payment_method,
                        notes,
                        g.user["id"]
                    )
                )


            conn.execute(
                """
                UPDATE accounts_payable

                SET
                    status='PAGADA',
                    closed_at=?,
                    updated_at=CURRENT_TIMESTAMP

                WHERE id=?
                """,
                (
                    datetime.now().isoformat(
                        timespec="seconds"
                    ),
                    account_id
                )
            )


        flash(
            "Factura marcada como pagada.",
            "success"
        )


    except ValueError as exc:

        flash(
            str(exc),
            "danger"
        )


    return redirect(
        url_for(
            "accounts_payable"
        )
    )


# ============================================================
# HISTORIAL
# ============================================================

@app.route(
    "/facturas-pendientes/<int:account_id>/historial"
)
@role_required("administrador")
def account_payable_history(account_id):

    with db_conn() as conn:

        account = conn.execute(
            """
            SELECT ap.*,

                   COALESCE(
                       (
                           SELECT SUM(p.amount)
                           FROM accounts_payable_payments p
                           WHERE p.payable_id=ap.id
                       ),
                       0
                   ) AS paid_amount

            FROM accounts_payable ap

            WHERE ap.id=?
            """,
            (
                account_id,
            )
        ).fetchone()


        payments = conn.execute(
            """
            SELECT
                p.*,
                u.full_name AS user_name

            FROM accounts_payable_payments p

            LEFT JOIN users u
                ON u.id=p.user_id

            WHERE p.payable_id=?

            ORDER BY
                p.payment_date DESC,
                p.id DESC
            """,
            (
                account_id,
            )
        ).fetchall()


    if not account:

        flash(
            "Factura pendiente no encontrada.",
            "danger"
        )

        return redirect(
            url_for(
                "accounts_payable"
            )
        )


    account = dict(
        account
    )


    original = Decimal(
        str(
            account.get(
                "original_amount"
            )
            or
            0
        )
    )


    paid = Decimal(
        str(
            account.get(
                "paid_amount"
            )
            or
            0
        )
    )


    balance = max(
        Decimal("0"),
        original
        -
        paid
    )


    account["balance"] = (
        balance
    )

    account["display_status"] = (
        _payable_display_status(
            account
        )
    )


    return render_template(
        "account_payable_history.html",
        account=account,
        payments=payments,
        total_paid=paid,
        balance=balance,
    )


SERVICES_PAYMENT_KIND = "SERVICIOS"


def _line_sum(rows) -> Decimal:
    return sum((Decimal(str(r["line_total"] or 0)) for r in rows), Decimal("0"))


def get_invoice_payment_summary(conn, invoice, items, service_items) -> dict[str, Any]:
    """Subtotales por tipo y estado de pago (abonos, servicios pagados y saldo) de una remisión."""
    product_subtotal = _line_sum(items)
    service_subtotal = _line_sum(service_items)
    total = Decimal(str(invoice["total"] or 0))
    is_credit = invoice["payment_method"] == "Crédito"

    receivable = None
    services_payment = None
    paid_total = Decimal("0")
    if is_credit:
        receivable = conn.execute(
            "SELECT * FROM accounts_receivable WHERE invoice_id=?",
            (invoice["id"],),
        ).fetchone()
    if receivable:
        payments = conn.execute(
            "SELECT * FROM accounts_receivable_payments WHERE receivable_id=? ORDER BY id",
            (receivable["id"],),
        ).fetchall()
        paid_total = sum((Decimal(str(p["amount"] or 0)) for p in payments), Decimal("0"))
        services_payment = next(
            (p for p in payments if (p.get("payment_kind") or "") == SERVICES_PAYMENT_KIND),
            None,
        )
        original = Decimal(str(receivable["original_amount"] or 0))
    else:
        original = total if is_credit else Decimal("0")

    services_paid_amount = Decimal(str(services_payment["amount"])) if services_payment else Decimal("0")
    return {
        "product_subtotal": product_subtotal,
        "service_subtotal": service_subtotal,
        "is_credit": is_credit,
        "receivable": receivable,
        "services_payment": services_payment,
        "services_paid": services_payment is not None,
        "services_paid_amount": services_paid_amount,
        "other_paid": paid_total - services_paid_amount,
        "paid_total": paid_total,
        "balance": max(Decimal("0"), original - paid_total) if is_credit else Decimal("0"),
    }


def _refresh_receivable_status(conn, receivable) -> None:
    """Recalcula estado/cierre de una cuenta por cobrar según sus pagos actuales."""
    paid_row = conn.execute(
        "SELECT COALESCE(SUM(amount),0) AS paid FROM accounts_receivable_payments WHERE receivable_id=?",
        (receivable["id"],),
    ).fetchone()
    paid = Decimal(str((paid_row or {}).get("paid") or 0))
    original = Decimal(str(receivable["original_amount"] or 0))
    if paid >= original and original > 0:
        status, closed_at = "PAGADA", datetime.now().isoformat(timespec="seconds")
    elif paid > 0:
        status, closed_at = "PAGO PARCIAL", None
    else:
        status, closed_at = "PENDIENTE", None
    conn.execute(
        "UPDATE accounts_receivable SET status=?,closed_at=? WHERE id=?",
        (status, closed_at, receivable["id"]),
    )


def can_mark_services_paid(user) -> bool:
    """Marcan servicios pagados quienes generan remisiones y cobran: arturo.lopez y el rol vendedor."""
    if user is None:
        return False
    username = str(user["username"] or "").strip().lower()
    role = str(user["role"] or "").strip().lower()
    return (username == "arturo.lopez" and role == "administrador") or role == "vendedor"


@app.route("/facturas/<int:invoice_id>/servicios-pagados", methods=["POST"])
@login_required
def invoice_services_paid(invoice_id: int):
    """Marca o desmarca que el cliente ya pagó los servicios (mano de obra) de una remisión a crédito.

    Marcar registra un abono de tipo SERVICIOS por el subtotal de servicios, así el saldo,
    cuentas por cobrar y finanzas lo reflejan sin lógica aparte. Desmarcar elimina ese abono.
    """
    if not can_mark_services_paid(g.user):
        flash("Solo arturo.lopez o un vendedor pueden marcar los servicios como pagados.", "danger")
        return redirect(url_for("invoice_view", invoice_id=invoice_id))

    mark = request.form.get("services_paid") == "1"
    try:
        with db_conn() as conn:
            invoice = conn.execute("SELECT * FROM invoices WHERE id=?", (invoice_id,)).fetchone()
            if not invoice:
                raise ValueError("Remisión no encontrada.")
            items = conn.execute("SELECT line_total FROM invoice_items WHERE invoice_id=?", (invoice_id,)).fetchall()
            service_items = conn.execute(
                "SELECT line_total FROM invoice_service_items WHERE invoice_id=?", (invoice_id,)
            ).fetchall()
            summary = get_invoice_payment_summary(conn, invoice, items, service_items)
            receivable = summary["receivable"]

            if not summary["is_credit"] or not receivable:
                raise ValueError("Solo las remisiones a crédito tienen saldo pendiente para descontar servicios.")

            if mark:
                if summary["services_paid"]:
                    raise ValueError("Los servicios de esta remisión ya están marcados como pagados.")
                service_subtotal = summary["service_subtotal"]
                if service_subtotal <= 0:
                    raise ValueError("Esta remisión no tiene servicios.")
                if service_subtotal > summary["balance"]:
                    raise ValueError(
                        f"El saldo pendiente ({money(summary['balance'])}) es menor que los servicios "
                        f"({money(service_subtotal)}). Registra el pago como abono en cuentas por cobrar."
                    )
                payment_method = request.form.get("payment_method", "Efectivo").strip() or "Efectivo"
                conn.execute(
                    """INSERT INTO accounts_receivable_payments(
                        receivable_id,amount,payment_date,payment_method,notes,payment_kind,user_id
                    ) VALUES(?,?,?,?,?,?,?)""",
                    (
                        receivable["id"],
                        float(service_subtotal),
                        datetime.now().date(),
                        payment_method,
                        "Pago de servicios (mano de obra)",
                        SERVICES_PAYMENT_KIND,
                        g.user["id"],
                    ),
                )
                message = f"Servicios marcados como pagados. Se descontaron {money(service_subtotal)} del saldo."
            else:
                if not summary["services_paid"]:
                    raise ValueError("Los servicios de esta remisión no estaban marcados como pagados.")
                conn.execute(
                    "DELETE FROM accounts_receivable_payments WHERE id=?",
                    (summary["services_payment"]["id"],),
                )
                message = "Se quitó la marca de servicios pagados. El saldo volvió a incluir los servicios."

            _refresh_receivable_status(conn, receivable)

        generate_invoice_pdf(invoice_id)
        flash(message, "success")
    except ValueError as exc:
        flash(str(exc), "danger")

    return redirect(url_for("invoice_view", invoice_id=invoice_id))


@app.route("/facturas/<int:invoice_id>")
@login_required
def invoice_view(invoice_id: int):
    with db_conn() as conn:
        invoice = conn.execute(
            """
            SELECT
                i.*,
                u.full_name AS user_name
            FROM invoices i
            LEFT JOIN users u ON u.id = i.user_id
            WHERE i.id = ?
            """,
            (invoice_id,),
        ).fetchone()

        items = conn.execute(
            """
            SELECT *
            FROM invoice_items
            WHERE invoice_id = ?
            ORDER BY id
            """,
            (invoice_id,),
        ).fetchall()

        service_items = conn.execute(
            """
            SELECT *
            FROM invoice_service_items
            WHERE invoice_id = ?
            ORDER BY id
            """,
            (invoice_id,),
        ).fetchall()

        settings = get_settings(conn)
        summary = get_invoice_payment_summary(conn, invoice, items, service_items) if invoice else None

    if not invoice:
        flash("Remisión no encontrada.", "danger")
        return redirect(url_for("invoices"))

    current_username = str(g.user["username"] or "").strip().lower()
    current_role = str(g.user["role"] or "").strip().lower()
    can_edit_invoice = (
        current_username == "arturo.lopez"
        and current_role == "administrador"
    )

    return render_template(
        "invoice_view.html",
        invoice=invoice,
        items=items,
        service_items=service_items,
        settings=settings,
        summary=summary,
        can_edit_invoice=can_edit_invoice,
        can_mark_services_paid=can_mark_services_paid(g.user),
    )


@app.route(
    "/facturas/<int:invoice_id>/eliminar",
    methods=["POST"]
)
@arturo_required
def invoice_delete(invoice_id: int):

    try:

        with db_conn() as conn:

            invoice = conn.execute(
                """
                SELECT *
                FROM invoices
                WHERE id=?
                """,
                (invoice_id,),
            ).fetchone()

            if not invoice:

                flash(
                    "La remisión no existe.",
                    "warning"
                )

                return redirect(
                    url_for("invoices")
                )


            # =============================================
            # PRODUCTOS DE LA REMISIÓN
            # =============================================

            items = conn.execute(
                """
                SELECT *
                FROM invoice_items
                WHERE invoice_id=?
                """,
                (invoice_id,),
            ).fetchall()


            # =============================================
            # DEVOLVER PRODUCTOS AL INVENTARIO
            # =============================================

            for item in items:

                if item["product_id"]:

                    conn.execute(
                        """
                        UPDATE products

                        SET
                            stock = stock + ?,
                            updated_at = CURRENT_TIMESTAMP

                        WHERE id = ?
                        """,
                        (
                            float(
                                item["quantity"]
                                or
                                0
                            ),
                            item["product_id"],
                        ),
                    )


            # =============================================
            # ELIMINAR MOVIMIENTOS GENERADOS POR REMISIÓN
            # =============================================

            conn.execute(
                """
                DELETE FROM movements

                WHERE reference = ?
                  AND movement_type = 'Venta'
                """,
                (
                    invoice["number"],
                ),
            )


            # =============================================
            # CUENTAS POR COBRAR
            # =============================================

            receivable = conn.execute(
                """
                SELECT id
                FROM accounts_receivable
                WHERE invoice_id=?
                """,
                (
                    invoice_id,
                ),
            ).fetchone()


            if receivable:

                conn.execute(
                    """
                    DELETE FROM accounts_receivable_payments
                    WHERE receivable_id=?
                    """,
                    (
                        receivable["id"],
                    ),
                )

                conn.execute(
                    """
                    DELETE FROM accounts_receivable
                    WHERE id=?
                    """,
                    (
                        receivable["id"],
                    ),
                )


            # =============================================
            # DETALLES
            # =============================================

            conn.execute(
                """
                DELETE FROM invoice_service_items
                WHERE invoice_id=?
                """,
                (
                    invoice_id,
                ),
            )


            conn.execute(
                """
                DELETE FROM invoice_items
                WHERE invoice_id=?
                """,
                (
                    invoice_id,
                ),
            )


            # =============================================
            # REMISIÓN
            # =============================================

            conn.execute(
                """
                DELETE FROM invoices
                WHERE id=?
                """,
                (
                    invoice_id,
                ),
            )


        # =============================================
        # BORRAR PDF GUARDADO SI EXISTE
        # =============================================

        try:

            pdf_path = (
                INVOICE_DIR
                /
                f"Remision_{invoice['number'].replace('-', '_')}.pdf"
            )

            if pdf_path.exists():

                pdf_path.unlink()

        except Exception:

            pass


        flash(
            f"Remisión {invoice['number']} eliminada. "
            "El inventario fue restaurado.",
            "success"
        )


    except Exception as exc:

        print(
            "ERROR ELIMINANDO REMISIÓN:",
            repr(exc)
        )

        flash(
            "No fue posible eliminar la remisión.",
            "danger"
        )


    return redirect(
        url_for("invoices")
    )


@app.route("/facturas/<int:invoice_id>/editar", methods=["GET", "POST"])
@arturo_required
def invoice_edit(invoice_id: int):
    """Edita una remisión y ajusta inventario/cartera de forma transaccional."""

    with db_conn() as conn:
        invoice = conn.execute(
            "SELECT * FROM invoices WHERE id=?",
            (invoice_id,),
        ).fetchone()

    if not invoice:
        flash("Remisión no encontrada.", "danger")
        return redirect(url_for("invoices"))

    if request.method == "POST":
        try:
            customer_name = (
                request.form.get("customer_name", "Consumidor final").strip()
                or "Consumidor final"
            )
            customer_document = request.form.get("customer_document", "").strip()
            customer_phone = request.form.get("customer_phone", "").strip()
            customer_address = request.form.get("customer_address", "").strip()
            customer_email = request.form.get("customer_email", "").strip()
            vehicle_plate = request.form.get("vehicle_plate", "").strip().upper()
            salesperson_id_raw = request.form.get("salesperson_id", "").strip()
            payment_method = request.form.get("payment_method", "Efectivo").strip()
            notes = request.form.get("notes", "").strip()
            issued_date = request.form.get("issued_date", "").strip()
            credit_due_date = request.form.get("credit_due_date", "").strip()

            if not salesperson_id_raw:
                raise ValueError("Selecciona quién realizó la venta.")
            salesperson_id = int(salesperson_id_raw)

            if payment_method not in ("Efectivo", "Transferencia", "Tarjeta", "Crédito"):
                raise ValueError("Forma de pago inválida.")

            if issued_date:
                try:
                    datetime.strptime(issued_date, "%Y-%m-%d")
                except ValueError as exc:
                    raise ValueError("La fecha de la remisión no es válida.") from exc
                issued_at = f"{issued_date}T12:00:00"
            else:
                issued_at = invoice["issued_at"]

            if credit_due_date:
                try:
                    datetime.strptime(credit_due_date, "%Y-%m-%d")
                except ValueError as exc:
                    raise ValueError("La fecha acordada de pago no es válida.") from exc

            product_rows_raw = [
                x
                for x in zip(
                    request.form.getlist("product_id[]"),
                    request.form.getlist("quantity[]"),
                    request.form.getlist("price[]"),
                )
                if str(x[0]).strip()
            ]

            service_rows_raw = [
                x
                for x in zip(
                    request.form.getlist("service_id[]"),
                    request.form.getlist("service_worker[]"),
                    request.form.getlist("service_quantity[]"),
                    request.form.getlist("service_price[]"),
                )
                if str(x[0]).strip()
            ]

            if not product_rows_raw and not service_rows_raw:
                raise ValueError(
                    "La remisión debe contener al menos un producto o un servicio."
                )

            with db_conn() as conn:
                invoice_sql = "SELECT * FROM invoices WHERE id=?"
                if db.engine.dialect.name == "postgresql":
                    invoice_sql += " FOR UPDATE"
                current_invoice = conn.execute(invoice_sql, (invoice_id,)).fetchone()
                if not current_invoice:
                    raise ValueError("La remisión ya no existe.")

                salesperson = conn.execute(
                    "SELECT id,name FROM salespeople WHERE id=? AND active=1",
                    (salesperson_id,),
                ).fetchone()
                if not salesperson:
                    raise ValueError("El vendedor seleccionado no está disponible.")

                # Cantidades anteriores por producto.
                old_items = conn.execute(
                    "SELECT * FROM invoice_items WHERE invoice_id=?",
                    (invoice_id,),
                ).fetchall()
                old_quantities: dict[int, Decimal] = {}
                for item in old_items:
                    pid = int(item["product_id"])
                    old_quantities[pid] = old_quantities.get(pid, Decimal("0")) + Decimal(
                        str(item["quantity"] or 0)
                    )

                requested_product_ids = {int(row[0]) for row in product_rows_raw}
                all_product_ids = set(old_quantities) | requested_product_ids

                # Bloquear productos implicados en PostgreSQL.
                product_map: dict[int, dict[str, Any]] = {}
                for product_id in sorted(all_product_ids):
                    product_sql = "SELECT * FROM products WHERE id=?"
                    if db.engine.dialect.name == "postgresql":
                        product_sql += " FOR UPDATE"
                    product = conn.execute(product_sql, (product_id,)).fetchone()
                    if not product:
                        raise ValueError("Uno de los productos ya no existe.")
                    product_map[product_id] = product

                subtotal = Decimal("0")
                product_lines = []
                new_quantities: dict[int, Decimal] = {}

                for product_id_raw, qty_raw, price_raw in product_rows_raw:
                    product_id = int(product_id_raw)
                    qty = parse_decimal(qty_raw)
                    price = parse_decimal(price_raw)

                    if qty <= Decimal("0"):
                        raise ValueError("La cantidad de cada producto debe ser mayor que cero.")
                    if price < Decimal("0"):
                        raise ValueError("El precio de un producto no puede ser negativo.")

                    product = product_map[product_id]
                    line_total = qty * price
                    subtotal += line_total
                    product_lines.append((product, qty, price, line_total))
                    new_quantities[product_id] = new_quantities.get(
                        product_id, Decimal("0")
                    ) + qty

                # Validar y aplicar únicamente la diferencia contra la remisión anterior.
                inventory_adjustments = []
                for product_id in sorted(all_product_ids):
                    product = product_map[product_id]
                    current_stock = Decimal(str(product["stock"] or 0))
                    old_qty = old_quantities.get(product_id, Decimal("0"))
                    new_qty = new_quantities.get(product_id, Decimal("0"))
                    adjustment = old_qty - new_qty
                    new_stock = current_stock + adjustment

                    if new_stock < Decimal("0"):
                        available_for_edit = current_stock + old_qty
                        raise ValueError(
                            f"Stock insuficiente para {product['description']}. "
                            f"Disponible considerando la remisión actual: {available_for_edit:g}."
                        )

                    if adjustment != Decimal("0"):
                        inventory_adjustments.append(
                            (product_id, product, adjustment, new_stock)
                        )

                movement_time = datetime.now().isoformat(timespec="seconds")
                for product_id, product, adjustment, new_stock in inventory_adjustments:
                    conn.execute(
                        """UPDATE products
                           SET stock=?, updated_at=CURRENT_TIMESTAMP
                           WHERE id=?""",
                        (float(new_stock), product_id),
                    )
                    movement_type = (
                        "Ajuste remisión entrada"
                        if adjustment > 0
                        else "Ajuste remisión salida"
                    )
                    conn.execute(
                        """INSERT INTO movements(
                               product_id,movement_type,quantity,reference,notes,created_at,user_id
                           ) VALUES(?,?,?,?,?,?,?)""",
                        (
                            product_id,
                            movement_type,
                            float(abs(adjustment)),
                            current_invoice["number"],
                            "Edición de remisión realizada por Arturo López",
                            movement_time,
                            g.user["id"],
                        ),
                    )

                conn.execute(
                    "DELETE FROM invoice_items WHERE invoice_id=?",
                    (invoice_id,),
                )
                for product, qty, price, line_total in product_lines:
                    conn.execute(
                        """INSERT INTO invoice_items(
                               invoice_id,product_id,code,description,quantity,unit_price,line_total
                           ) VALUES(?,?,?,?,?,?,?)""",
                        (
                            invoice_id,
                            product["id"],
                            product["code"],
                            product["description"],
                            float(qty),
                            float(price),
                            float(line_total),
                        ),
                    )

                # Reconstruir servicios.
                service_lines = []
                for service_id_raw, worker_name, qty_raw, price_raw in service_rows_raw:
                    service_id = int(service_id_raw)
                    worker_name = (worker_name or "").strip()
                    qty = parse_decimal(qty_raw)
                    price = parse_decimal(price_raw)

                    if not worker_name:
                        raise ValueError(
                            "Indica el trabajador responsable de cada servicio."
                        )
                    if worker_name not in ALLOWED_MECHANICS:
                        raise ValueError(
                            f"Mecánico no válido: {worker_name}. "
                            f"Selecciona uno de: {', '.join(ALLOWED_MECHANICS)}."
                        )
                    if qty <= Decimal("0") or price < Decimal("0"):
                        raise ValueError(
                            "Revisa las cantidades y precios de los servicios."
                        )

                    service = conn.execute(
                        "SELECT * FROM services WHERE id=?",
                        (service_id,),
                    ).fetchone()
                    if not service:
                        raise ValueError("Uno de los servicios ya no existe.")

                    line_total = qty * price
                    worker_pct = Decimal(str(service["worker_percentage"] or 70))
                    business_pct = Decimal(str(service["business_percentage"] or 30))
                    worker_amount = (
                        line_total * worker_pct / Decimal("100")
                    ).quantize(Decimal("0.01"))
                    business_amount = line_total - worker_amount
                    subtotal += line_total

                    service_lines.append(
                        (
                            service,
                            worker_name,
                            qty,
                            price,
                            line_total,
                            worker_pct,
                            business_pct,
                            worker_amount,
                            business_amount,
                        )
                    )

                conn.execute(
                    "DELETE FROM invoice_service_items WHERE invoice_id=?",
                    (invoice_id,),
                )
                for (
                    service,
                    worker_name,
                    qty,
                    price,
                    line_total,
                    worker_pct,
                    business_pct,
                    worker_amount,
                    business_amount,
                ) in service_lines:
                    conn.execute(
                        """INSERT INTO invoice_service_items(
                               invoice_id,service_id,service_name,worker_name,quantity,unit_price,line_total,
                               worker_percentage,business_percentage,worker_amount,business_amount
                           ) VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            invoice_id,
                            service["id"],
                            service["name"],
                            worker_name,
                            float(qty),
                            float(price),
                            float(line_total),
                            float(worker_pct),
                            float(business_pct),
                            float(worker_amount),
                            float(business_amount),
                        ),
                    )

                settings = get_settings(conn)
                tax_rate = Decimal(str(settings["tax_rate"] or 0))
                tax = (subtotal * tax_rate / Decimal("100")).quantize(Decimal("0.01"))
                total = subtotal + tax

                # Cliente: conservar/actualizar o volver a consumidor final.
                customer_id = None
                if customer_name.lower() != "consumidor final" or customer_document:
                    existing = None
                    if customer_document:
                        existing = conn.execute(
                            "SELECT id FROM customers WHERE document=? LIMIT 1",
                            (customer_document,),
                        ).fetchone()

                    if existing:
                        customer_id = existing["id"]
                        conn.execute(
                            """UPDATE customers
                               SET name=?,phone=?,email=?,address=?
                               WHERE id=?""",
                            (
                                customer_name,
                                customer_phone,
                                customer_email,
                                customer_address,
                                customer_id,
                            ),
                        )
                    else:
                        cur_customer = conn.execute(
                            """INSERT INTO customers(document,name,phone,email,address)
                               VALUES(?,?,?,?,?)""",
                            (
                                customer_document,
                                customer_name,
                                customer_phone,
                                customer_email,
                                customer_address,
                            ),
                        )
                        customer_id = cur_customer.lastrowid

                # Cartera: conservar pagos ya realizados y recalcular saldo/estado.
                receivable = conn.execute(
                    "SELECT * FROM accounts_receivable WHERE invoice_id=?",
                    (invoice_id,),
                ).fetchone()
                paid_amount = Decimal("0")
                if receivable:
                    # Si los servicios estaban marcados como pagados, el pago sigue al nuevo subtotal de servicios.
                    new_service_subtotal = sum((line[4] for line in service_lines), Decimal("0"))
                    services_payment = conn.execute(
                        "SELECT id FROM accounts_receivable_payments WHERE receivable_id=? AND payment_kind=?",
                        (receivable["id"], SERVICES_PAYMENT_KIND),
                    ).fetchone()
                    if services_payment and new_service_subtotal <= 0:
                        conn.execute(
                            "DELETE FROM accounts_receivable_payments WHERE id=?",
                            (services_payment["id"],),
                        )
                    elif services_payment:
                        conn.execute(
                            "UPDATE accounts_receivable_payments SET amount=? WHERE id=?",
                            (float(new_service_subtotal), services_payment["id"]),
                        )
                    paid_row = conn.execute(
                        """SELECT COALESCE(SUM(amount),0) AS paid
                           FROM accounts_receivable_payments
                           WHERE receivable_id=?""",
                        (receivable["id"],),
                    ).fetchone()
                    paid_amount = Decimal(str((paid_row or {}).get("paid") or 0))

                if payment_method == "Crédito":
                    if paid_amount > total:
                        raise ValueError(
                            "No se puede reducir el total de la remisión por debajo "
                            "de los pagos que ya fueron registrados."
                        )

                    if paid_amount >= total and total > Decimal("0"):
                        credit_status = "PAGADA"
                    elif paid_amount > Decimal("0"):
                        credit_status = "PAGO PARCIAL"
                    else:
                        credit_status = "PENDIENTE"

                    if credit_due_date and paid_amount < total:
                        due_date = datetime.strptime(
                            credit_due_date, "%Y-%m-%d"
                        ).date()
                        if due_date < datetime.now().date():
                            credit_status = "VENCIDA"

                    closed_at = (
                        datetime.now().isoformat(timespec="seconds")
                        if credit_status == "PAGADA"
                        else None
                    )

                    if receivable:
                        conn.execute(
                            """UPDATE accounts_receivable
                               SET original_amount=?,due_date=?,status=?,closed_at=?
                               WHERE id=?""",
                            (
                                float(total),
                                credit_due_date or None,
                                credit_status,
                                closed_at,
                                receivable["id"],
                            ),
                        )
                    else:
                        conn.execute(
                            """INSERT INTO accounts_receivable(
                                   invoice_id,original_amount,due_date,status,created_at,closed_at
                               ) VALUES(?,?,?,?,?,?)""",
                            (
                                invoice_id,
                                float(total),
                                credit_due_date or None,
                                credit_status,
                                datetime.now().isoformat(timespec="seconds"),
                                closed_at,
                            ),
                        )
                elif receivable:
                    if paid_amount > Decimal("0"):
                        raise ValueError(
                            "Esta remisión a crédito ya tiene pagos registrados. "
                            "No puede cambiarse a otra forma de pago desde la edición."
                        )
                    conn.execute(
                        "DELETE FROM accounts_receivable_payments WHERE receivable_id=?",
                        (receivable["id"],),
                    )
                    conn.execute(
                        "DELETE FROM accounts_receivable WHERE id=?",
                        (receivable["id"],),
                    )

                conn.execute(
                    """UPDATE invoices
                       SET customer_id=?,customer_name=?,customer_document=?,customer_phone=?,
                           customer_address=?,vehicle_plate=?,salesperson_id=?,salesperson_name=?,
                           issued_at=?,subtotal=?,tax=?,total=?,payment_method=?,notes=?
                       WHERE id=?""",
                    (
                        customer_id,
                        customer_name,
                        customer_document,
                        customer_phone,
                        customer_address,
                        vehicle_plate,
                        salesperson["id"],
                        salesperson["name"],
                        issued_at,
                        float(subtotal),
                        float(tax),
                        float(total),
                        payment_method,
                        notes,
                        invoice_id,
                    ),
                )

            generate_invoice_pdf(invoice_id)
            flash(
                f"Remisión {invoice['number']} actualizada correctamente.",
                "success",
            )
            return redirect(url_for("invoice_view", invoice_id=invoice_id))

        except (ValueError, KeyError, TypeError) as exc:
            flash(str(exc), "danger")

    with db_conn() as conn:
        invoice = conn.execute(
            "SELECT * FROM invoices WHERE id=?",
            (invoice_id,),
        ).fetchone()
        items = conn.execute(
            "SELECT * FROM invoice_items WHERE invoice_id=? ORDER BY id",
            (invoice_id,),
        ).fetchall()
        service_items = conn.execute(
            "SELECT * FROM invoice_service_items WHERE invoice_id=? ORDER BY id",
            (invoice_id,),
        ).fetchall()
        products = conn.execute(
            """SELECT id,code,description,stock,sale_price,unit
               FROM products
               WHERE active=1
               ORDER BY description"""
        ).fetchall()
        services = conn.execute(
            "SELECT * FROM services WHERE active=1 ORDER BY name"
        ).fetchall()
        salespeople = conn.execute(
            "SELECT id,name FROM salespeople WHERE active=1 ORDER BY name"
        ).fetchall()
        receivable = conn.execute(
            "SELECT * FROM accounts_receivable WHERE invoice_id=?",
            (invoice_id,),
        ).fetchone()

    return render_template(
        "invoice_edit.html",
        invoice=invoice,
        items=items,
        service_items=service_items,
        products=products,
        services=services,
        salespeople=salespeople,
        receivable=receivable,
    )


@app.route("/facturas/<int:invoice_id>/pdf")
@login_required
def invoice_pdf(invoice_id: int):
    pdf_path = generate_invoice_pdf(invoice_id)
    return send_file(pdf_path, as_attachment=True, download_name=pdf_path.name)

@app.route("/facturas/<int:invoice_id>/pdf/ver")
@login_required
def invoice_pdf_view(invoice_id: int):

    pdf_path = generate_invoice_pdf(
        invoice_id
    )

    return send_file(
        pdf_path,
        mimetype="application/pdf",
        as_attachment=False,
        download_name=pdf_path.name
    )


def generate_invoice_pdf(invoice_id: int) -> Path:
    """Genera una remisión compacta en PDF y devuelve la ruta del archivo creado."""
    with db_conn() as conn:
        invoice = conn.execute(
            """
            SELECT
                i.*,
                u.full_name AS user_name
            FROM invoices i
            LEFT JOIN users u ON u.id = i.user_id
            WHERE i.id = ?
            """,
            (invoice_id,),
        ).fetchone()

        items = conn.execute(
            """
            SELECT *
            FROM invoice_items
            WHERE invoice_id = ?
            ORDER BY id
            """,
            (invoice_id,),
        ).fetchall()

        service_items = conn.execute(
            """
            SELECT *
            FROM invoice_service_items
            WHERE invoice_id = ?
            ORDER BY id
            """,
            (invoice_id,),
        ).fetchall()

        settings = get_settings(conn)
        payment_summary = get_invoice_payment_summary(conn, invoice, items, service_items) if invoice else None

    if not invoice:
        raise ValueError("Remisión no encontrada")

    path = INVOICE_DIR / f"Remision_{invoice['number'].replace('-', '_')}.pdf"

    # Hoja carta con márgenes reducidos para que remisiones normales quepan en una página.
    doc = SimpleDocTemplate(
        str(path),
        pagesize=letter,
        rightMargin=7 * mm,
        leftMargin=7 * mm,
        topMargin=7 * mm,
        bottomMargin=7 * mm,
    )

    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="InvoiceSmall",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=7,
            leading=8.2,
            spaceAfter=0,
            spaceBefore=0,
        )
    )
    styles.add(
        ParagraphStyle(
            name="InvoiceTiny",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=6.3,
            leading=7.2,
            spaceAfter=0,
            spaceBefore=0,
        )
    )
    styles.add(
        ParagraphStyle(
            name="InvoiceFooter",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=6.2,
            leading=7,
            alignment=1,
            textColor=colors.HexColor("#555555"),
            spaceAfter=0,
            spaceBefore=0,
        )
    )

    small = styles["InvoiceSmall"]
    tiny = styles["InvoiceTiny"]
    footer_style = styles["InvoiceFooter"]
    story = []

    # Encabezado
    business_name = settings["business_name"] or ""
    nit = settings["nit"] or ""
    address = settings["address"] or ""
    phone = settings["phone"] or ""
    email = settings["email"] or ""

    issued_at = str(invoice["issued_at"] or "")
    issued_date = issued_at[:10] if issued_at else "-"
    status = invoice["status"] or "-"

    header_left = Paragraph(
        f"<b>{business_name}</b><br/>{nit}<br/>{address}<br/>{phone} - {email}",
        small,
    )
    header_right = Paragraph(
        "<b>REMISIÓN</b>"
        f"<br/><font size='10'><b>{invoice['number']}</b></font>"
        f"<br/>Fecha: {issued_date}"
        f"<br/>Estado: {status}",
        small,
    )

    header = Table(
        [[header_left, header_right]],
        colWidths=[125 * mm, 66 * mm],
    )
    header.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#1F4E78")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#D9EAF7")),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    story.append(header)
    story.append(Spacer(1, 2.5 * mm))

    # Datos del cliente
    customer_name = invoice["customer_name"] or "-"
    customer_document = invoice["customer_document"] or "-"
    customer_phone = invoice["customer_phone"] or "-"
    customer_address = invoice["customer_address"] or "-"
    vehicle_plate = invoice["vehicle_plate"] or "-"
    payment_method = invoice["payment_method"] or "-"
    salesperson_name = invoice["salesperson_name"] or "-"
    user_name = invoice["user_name"] or "-"
    notes = invoice["notes"] or "-"

    customer = Table(
        [
            [
                "Cliente",
                Paragraph(str(customer_name), tiny),
                "Documento",
                Paragraph(str(customer_document), tiny),
            ],
            [
                "Teléfono",
                Paragraph(str(customer_phone), tiny),
                "Dirección",
                Paragraph(str(customer_address), tiny),
            ],
            [
                "Placa",
                Paragraph(str(vehicle_plate), tiny),
                "Forma de pago",
                Paragraph(str(payment_method), tiny),
            ],
            [
                "Vendedor",
                Paragraph(str(salesperson_name), tiny),
                "Usuario",
                Paragraph(str(user_name), tiny),
            ],
            [
                "Observaciones",
                Paragraph(str(notes), tiny),
                "",
                "",
            ],
        ],
        colWidths=[23 * mm, 72 * mm, 27 * mm, 69 * mm],
    )
    customer.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#999999")),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#EEF3F8")),
                ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#EEF3F8")),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 6.7),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    story.append(customer)
    story.append(Spacer(1, 2.5 * mm))

    # Detalle de productos y servicios, cada grupo con su subtotal
    data = [["Tipo", "Descripción", "Cant.", "Valor unitario", "Total"]]
    subtotal_rows = []

    for item in items:
        data.append(
            [
                "Producto",
                Paragraph(str(item["description"] or ""), tiny),
                f"{float(item['quantity'] or 0):g}",
                money(item["unit_price"]),
                money(item["line_total"]),
            ]
        )
    if items:
        subtotal_rows.append(len(data))
        data.append(["", "", "", "Subtotal productos", money(payment_summary["product_subtotal"])])

    for item in service_items:
        worker_name = str(item["worker_name"] or "").strip()
        service_text = str(item["service_name"] or "")
        if worker_name:
            service_text += f" - {worker_name}"
        data.append(
            [
                "Servicio",
                Paragraph(service_text, tiny),
                f"{float(item['quantity'] or 0):g}",
                money(item["unit_price"]),
                money(item["line_total"]),
            ]
        )
    if service_items:
        subtotal_rows.append(len(data))
        data.append(["", "", "", "Subtotal servicios", money(payment_summary["service_subtotal"])])

    detail_end = len(data) - 1
    data.extend(
        [
            ["", "", "", "Subtotal", money(invoice["subtotal"])],
            ["", "", "", "Impuesto", money(invoice["tax"])],
            ["", "", "", "TOTAL", money(invoice["total"])],
        ]
    )
    total_row = len(data) - 1

    # Remisiones a crédito: lo ya pagado y el saldo que queda debiendo.
    if payment_summary["is_credit"]:
        if payment_summary["services_paid"]:
            data.append(["", "", "", "Servicios pagados", f"- {money(payment_summary['services_paid_amount'])}"])
        if payment_summary["other_paid"] > 0:
            data.append(["", "", "", "Abonos", f"- {money(payment_summary['other_paid'])}"])
        data.append(["", "", "", "SALDO PENDIENTE", money(payment_summary["balance"])])

    table_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 6.7),
        ("GRID", (0, 0), (-1, detail_end), 0.3, colors.HexColor("#999999")),
        ("LINEABOVE", (3, detail_end + 1), (-1, -1), 0.4, colors.HexColor("#777777")),
        ("FONTNAME", (3, total_row), (-1, total_row), "Helvetica-Bold"),
        ("BACKGROUND", (3, total_row), (-1, total_row), colors.HexColor("#D9EAF7")),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2.5),
        ("TOPPADDING", (0, 0), (-1, -1), 1.8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.8),
    ]
    for row in subtotal_rows:
        table_style.extend(
            [
                ("SPAN", (0, row), (2, row)),
                ("FONTNAME", (3, row), (-1, row), "Helvetica-Bold"),
                ("BACKGROUND", (0, row), (-1, row), colors.HexColor("#EEF3F8")),
            ]
        )
    if payment_summary["is_credit"]:
        table_style.extend(
            [
                ("FONTNAME", (3, -1), (-1, -1), "Helvetica-Bold"),
                ("BACKGROUND", (3, -1), (-1, -1), colors.HexColor("#FCE4D6")),
            ]
        )

    table = Table(
        data,
        colWidths=[22 * mm, 91 * mm, 14 * mm, 31 * mm, 33 * mm],
        repeatRows=1,
    )
    table.setStyle(TableStyle(table_style))
    story.append(table)
    story.append(Spacer(1, 3 * mm))
    story.append(
        Paragraph(
            "Documento de remisión. Gracias por su compra.",
            footer_style,
        )
    )

    # Construir y retornar: ambas líneas son obligatorias.
    doc.build(story)
    return path


@app.route("/configuracion", methods=["GET", "POST"])
@role_required("administrador")
def settings_view():
    if request.method == "POST":
        try:
            tax_rate = float(parse_decimal(request.form.get("tax_rate")))
            with db_conn() as conn:
                conn.execute(
                    """UPDATE settings SET business_name=?,nit=?,address=?,phone=?,email=?,invoice_prefix=?,tax_rate=? WHERE id=1""",
                    (request.form["business_name"].strip(), request.form.get("nit", "").strip(),
                     request.form.get("address", "").strip(), request.form.get("phone", "").strip(),
                     request.form.get("email", "").strip(), request.form.get("invoice_prefix", "REM").strip().upper(), tax_rate),
                )
            flash("Configuración guardada.", "success")
            return redirect(url_for("settings_view"))
        except ValueError as exc:
            flash(str(exc), "danger")
    settings = get_settings()
    return render_template("settings.html", settings=settings, db_path=DB_PATH, db_backend=db.engine.dialect.name)


@app.route("/respaldo")
@role_required("administrador")
def backup():
    """Exporta un respaldo JSON portable tanto para SQLite como PostgreSQL."""
    import json
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = BACKUP_DIR / f"inventario_respaldo_{stamp}.json"
    table_names = [
        "users", "settings", "products", "customers", "salespeople", "services",
        "invoices", "invoice_items", "invoice_service_items", "movements",
        "accounts_receivable", "accounts_receivable_payments",
        "accounts_payable", "accounts_payable_payments",
        "expenses", "extra_income"
    ]
    payload = {"created_at": datetime.now().isoformat(), "database": db.engine.dialect.name, "tables": {}}
    with db_conn() as conn:
        for table_name in table_names:
            payload["tables"][table_name] = conn.execute(f"SELECT * FROM {table_name}").fetchall()
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    return send_file(path, as_attachment=True, download_name=path.name)


def open_browser():
    time.sleep(1.2)
    webbrowser.open("http://127.0.0.1:5000")


# En Render/Gunicorn el módulo se importa, por eso inicializamos también en import.
# La operación es idempotente: no borra productos, usuarios, vendedores ni servicios existentes.
init_db()

if __name__ == "__main__":
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host="127.0.0.1", port=int(os.getenv("PORT", "5000")), debug=False, use_reloader=False)
