from __future__ import annotations

import os
import sys
import threading
import time
import webbrowser
from functools import wraps
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from flask import Flask, flash, g, redirect, render_template, request, send_file, session, url_for
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

load_dotenv()

APP_NAME = "Control de Inventario y Remisiones"


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


@app.route("/movimientos", methods=["GET", "POST"])
@role_required("administrador")
def movements():
    if request.method == "POST":
        try:
            product_id = int(request.form["product_id"])
            movement_type = request.form["movement_type"]
            quantity = float(parse_decimal(request.form["quantity"]))
            if quantity <= 0:
                raise ValueError("La cantidad debe ser mayor que cero.")
            signed = quantity if movement_type == "Entrada" else -quantity
            with db_conn() as conn:
                product = conn.execute("SELECT * FROM products WHERE id=?", (product_id,)).fetchone()
                if not product:
                    raise ValueError("Producto no encontrado.")
                if product["stock"] + signed < 0:
                    raise ValueError("La salida supera el inventario disponible.")
                conn.execute("UPDATE products SET stock=stock+?, updated_at=CURRENT_TIMESTAMP WHERE id=?", (signed, product_id))
                conn.execute(
                    "INSERT INTO movements(product_id,movement_type,quantity,reference,notes,created_at,user_id) VALUES(?,?,?,?,?,?,?)",
                    (product_id, movement_type, quantity, request.form.get("reference", ""), request.form.get("notes", ""), datetime.now().isoformat(timespec="seconds"), g.user["id"]),
                )
            flash("Movimiento registrado.", "success")
            return redirect(url_for("movements"))
        except (ValueError, KeyError) as exc:
            flash(str(exc), "danger")

    with db_conn() as conn:
        product_rows = conn.execute("SELECT id, code, description, stock FROM products WHERE active=1 ORDER BY description").fetchall()
        rows = conn.execute(
            """SELECT m.*, p.code, p.description, u.full_name AS user_name FROM movements m
            JOIN products p ON p.id=m.product_id
            LEFT JOIN users u ON u.id=m.user_id ORDER BY m.id DESC LIMIT 100"""
        ).fetchall()
    return render_template("movements.html", products=product_rows, movements=rows)



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
    if request.method == "POST":
        if g.user["role"] != "administrador":
            flash("Solo el administrador puede crear servicios.", "danger")
            return redirect(url_for("services"))
        try:
            name = request.form.get("name", "").strip()
            price = float(parse_decimal(request.form.get("default_price")))
            if not name:
                raise ValueError("El nombre del servicio es obligatorio.")
            if price < 0:
                raise ValueError("El valor del servicio no puede ser negativo.")
            with db_conn() as conn:
                conn.execute(
                    """INSERT INTO services(name,default_price,worker_percentage,business_percentage)
                    VALUES(?,?,70,30)""",
                    (name, price),
                )
            flash("Servicio creado correctamente.", "success")
            return redirect(url_for("services"))
        except IntegrityError:
            flash("Ya existe un servicio con ese nombre.", "danger")
        except ValueError as exc:
            flash(str(exc), "danger")

    with db_conn() as conn:
        catalog = conn.execute(
            "SELECT * FROM services WHERE active=1 ORDER BY name"
        ).fetchall()
        history = conn.execute(
            """SELECT isi.*, i.number, i.issued_at, i.customer_name, i.vehicle_plate
               FROM invoice_service_items isi
               JOIN invoices i ON i.id=isi.invoice_id
               ORDER BY isi.id DESC LIMIT 200"""
        ).fetchall()
        summary = conn.execute(
            """SELECT
                   COUNT(*) AS service_lines,
                   COALESCE(SUM(line_total),0) AS service_total,
                   COALESCE(SUM(worker_amount),0) AS worker_total,
                   COALESCE(SUM(business_amount),0) AS business_total
               FROM invoice_service_items"""
        ).fetchone()
        workers = conn.execute(
            """SELECT worker_name,
                   COUNT(*) AS service_lines,
                   COALESCE(SUM(line_total),0) AS service_total,
                   COALESCE(SUM(worker_amount),0) AS worker_total,
                   COALESCE(SUM(business_amount),0) AS business_total
               FROM invoice_service_items
               WHERE TRIM(worker_name)<>''
               GROUP BY worker_name
               ORDER BY worker_total DESC"""
        ).fetchall()
    return render_template("services.html", services=catalog, history=history, summary=summary, workers=workers)


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


@app.route("/facturas")
@login_required
def invoices():
    with db_conn() as conn:
        rows = conn.execute("""SELECT i.*, u.full_name AS user_name FROM invoices i
            LEFT JOIN users u ON u.id=i.user_id ORDER BY i.id DESC""").fetchall()
    return render_template("invoices.html", invoices=rows)


@app.route("/facturas/nueva", methods=["GET", "POST"])
@login_required
def invoice_new():
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
            payment_method = request.form.get("payment_method", "Efectivo")
            notes = request.form.get("notes", "").strip()

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
                issued_at = datetime.now().isoformat(timespec="seconds")

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
        "invoice_form.html", products=product_rows, services=service_rows,
        salespeople=salesperson_rows, settings=settings
    )


@app.route("/facturas/<int:invoice_id>")
@login_required
def invoice_view(invoice_id: int):
    with db_conn() as conn:
        invoice = conn.execute("""SELECT i.*, u.full_name AS user_name FROM invoices i
            LEFT JOIN users u ON u.id=i.user_id WHERE i.id=?""", (invoice_id,)).fetchone()
        items = conn.execute("SELECT * FROM invoice_items WHERE invoice_id=?", (invoice_id,)).fetchall()
        service_items = conn.execute("SELECT * FROM invoice_service_items WHERE invoice_id=?", (invoice_id,)).fetchall()
        settings = get_settings(conn)
    if not invoice:
        flash("Remisión no encontrada.", "danger")
        return redirect(url_for("invoices"))
    return render_template("invoice_view.html", invoice=invoice, items=items, service_items=service_items, settings=settings)


@app.route("/facturas/<int:invoice_id>/pdf")
@login_required
def invoice_pdf(invoice_id: int):
    pdf_path = generate_invoice_pdf(invoice_id)
    return send_file(pdf_path, as_attachment=True, download_name=pdf_path.name)


def generate_invoice_pdf(invoice_id: int) -> Path:
    with db_conn() as conn:
        invoice = conn.execute("""SELECT i.*, u.full_name AS user_name FROM invoices i
            LEFT JOIN users u ON u.id=i.user_id WHERE i.id=?""", (invoice_id,)).fetchone()
        items = conn.execute("SELECT * FROM invoice_items WHERE invoice_id=?", (invoice_id,)).fetchall()
        service_items = conn.execute("SELECT * FROM invoice_service_items WHERE invoice_id=?", (invoice_id,)).fetchall()
        settings = get_settings(conn)
    if not invoice:
        raise ValueError("Remisión no encontrada")

    path = INVOICE_DIR / f"Remision_{invoice['number'].replace('-', '_')}.pdf"
    doc = SimpleDocTemplate(str(path), pagesize=letter, rightMargin=16*mm, leftMargin=16*mm, topMargin=15*mm, bottomMargin=15*mm)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Small", parent=styles["Normal"], fontSize=8.5, leading=11))
    story = []

    header = Table([
        [Paragraph(f"<b>{settings['business_name']}</b><br/>{settings['nit']}<br/>{settings['address']}<br/>{settings['phone']} - {settings['email']}", styles["Small"]),
         Paragraph(f"<b>REMISIÓN</b><br/><font size=13>{invoice['number']}</font><br/>Fecha: {invoice['issued_at'][:10]}<br/>Estado: {invoice['status']}", styles["Small"])]
    ], colWidths=[110*mm, 70*mm])
    header.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 0.8, colors.HexColor("#1F4E78")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("BACKGROUND", (1,0), (1,0), colors.HexColor("#D9EAF7")),
        ("LEFTPADDING", (0,0), (-1,-1), 8), ("RIGHTPADDING", (0,0), (-1,-1), 8),
        ("TOPPADDING", (0,0), (-1,-1), 8), ("BOTTOMPADDING", (0,0), (-1,-1), 8),
    ]))
    story.extend([header, Spacer(1, 6*mm)])

    customer = Table([
        ["Cliente", invoice["customer_name"], "Documento", invoice["customer_document"] or "-"],
        ["Teléfono", invoice["customer_phone"] or "-", "Dirección", invoice["customer_address"] or "-"],
        ["Placa", invoice["vehicle_plate"] or "-", "Forma de pago", invoice["payment_method"]],
        ["Vendedor", invoice["salesperson_name"] or "-", "Usuario", invoice["user_name"] or "-"],
        ["Observaciones", invoice["notes"] or "-", "", ""],
    ], colWidths=[25*mm, 65*mm, 28*mm, 62*mm])
    customer.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.4, colors.grey),
        ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#EEF3F8")),
        ("BACKGROUND", (2,0), (2,-1), colors.HexColor("#EEF3F8")),
        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"), ("FONTNAME", (2,0), (2,-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 8), ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("PADDING", (0,0), (-1,-1), 5),
    ]))
    story.extend([customer, Spacer(1, 6*mm)])

    data = [["Tipo", "Descripción", "Cant.", "Valor unitario", "Total"]]
    for item in items:
        data.append([
            "Producto", Paragraph(item["description"], styles["Small"]), f"{item['quantity']:g}",
            money(item["unit_price"]), money(item["line_total"])
        ])
    for item in service_items:
        data.append([
            "Servicio", Paragraph(item["service_name"], styles["Small"]), f"{item['quantity']:g}",
            money(item["unit_price"]), money(item["line_total"])
        ])
    data.extend([
        ["", "", "", "Subtotal", money(invoice["subtotal"])],
        ["", "", "", "Impuesto", money(invoice["tax"])],
        ["", "", "", "TOTAL", money(invoice["total"])],
    ])
    table = Table(data, colWidths=[25*mm, 78*mm, 16*mm, 30*mm, 31*mm], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (2,1), (-1,-1), "RIGHT"),
        ("GRID", (0,0), (-1,-4), 0.35, colors.grey),
        ("LINEABOVE", (3,-3), (-1,-1), 0.5, colors.grey),
        ("FONTNAME", (3,-1), (-1,-1), "Helvetica-Bold"),
        ("BACKGROUND", (3,-1), (-1,-1), colors.HexColor("#D9EAF7")),
        ("FONTSIZE", (0,0), (-1,-1), 8.5),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,-1), 5), ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ]))
    story.extend([table, Spacer(1, 9*mm), Paragraph("Documento de remisión. Gracias por su compra.", styles["Normal"])])
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
        "invoices", "invoice_items", "invoice_service_items", "movements"
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
