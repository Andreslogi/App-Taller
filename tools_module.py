"""Inventario de herramientas del taller.

Las herramientas (llaves, taladros, pistolas de impacto, etc.) NO se venden:
no aparecen en remisiones ni afectan el inventario de productos. El objetivo
es saber cuántas hay, dónde están y quién las tiene, y detectar pérdidas
mediante conteos físicos periódicos.
"""

from __future__ import annotations

import csv
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation
from functools import wraps
from io import StringIO

from flask import (
    Blueprint,
    Response,
    flash,
    g,
    redirect,
    render_template,
    request,
    url_for,
)
from sqlalchemy import func, text

from database import IntegrityError, db, db_conn


tools_bp = Blueprint(
    "tools",
    __name__,
    url_prefix="/herramientas",
)


TOOL_STATUSES = (
    "Disponible",
    "En uso",
    "En reparación",
)

# Tipos de movimiento y su efecto sobre la cantidad.
#   +1 suma, -1 resta.
MOVEMENT_TYPES = {
    "Entrada": 1,       # compra o llegada de herramienta nueva
    "Pérdida": -1,      # se perdió o la robaron
    "Daño / baja": -1,  # se dañó y se retira
}

# Tipos que se cuentan como pérdida en los resúmenes.
LOSS_TYPES = ("Pérdida", "Faltante en conteo")


# ============================================================
# MODELOS
# ============================================================

class Tool(db.Model):
    __tablename__ = "tools"

    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(100), nullable=False, unique=True)
    name = db.Column(db.String(300), nullable=False)
    category = db.Column(db.String(150), default="")
    brand = db.Column(db.String(150), default="")
    serial = db.Column(db.String(150), default="")
    quantity = db.Column(db.Integer, nullable=False, default=1)
    unit_cost = db.Column(db.Numeric(16, 2), nullable=False, default=0)
    location = db.Column(db.String(200), default="")
    responsible = db.Column(db.String(180), default="")
    status = db.Column(db.String(40), nullable=False, default="Disponible", server_default=text("'Disponible'"))
    notes = db.Column(db.Text, default="")
    active = db.Column(db.Integer, nullable=False, default=1, server_default=text("1"))
    last_counted_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, server_default=func.now())
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow, server_default=func.now())


class ToolMovement(db.Model):
    __tablename__ = "tool_movements"

    id = db.Column(db.Integer, primary_key=True)
    tool_id = db.Column(db.Integer, db.ForeignKey("tools.id", ondelete="CASCADE"), nullable=False)
    movement_type = db.Column(db.String(60), nullable=False)
    # Cantidad con signo: positiva suma, negativa resta.
    quantity = db.Column(db.Integer, nullable=False)
    previous_quantity = db.Column(db.Integer, nullable=False, default=0)
    new_quantity = db.Column(db.Integer, nullable=False, default=0)
    notes = db.Column(db.Text, default="")
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, server_default=func.now())


# ============================================================
# SEGURIDAD
# ============================================================

def login_required(view):
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if g.user is None:
            flash("Debes iniciar sesión.", "warning")
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapped_view


def admin_required(view):
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if g.user is None:
            flash("Debes iniciar sesión.", "warning")
            return redirect(url_for("login"))
        if g.user["role"] != "administrador":
            flash("No tienes permisos para realizar esta acción.", "danger")
            return redirect(url_for("tools.index"))
        return view(*args, **kwargs)
    return wrapped_view


# ============================================================
# UTILIDADES
# ============================================================

def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _parse_int(value, field: str, minimum: int = 0) -> int:
    raw = str(value if value is not None else "").strip()
    if not raw:
        raise ValueError(f"{field} es obligatorio.")
    try:
        number = int(Decimal(raw.replace(",", ".")))
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f"{field} debe ser un número entero.") from exc
    if number < minimum:
        raise ValueError(f"{field} no puede ser menor que {minimum}.")
    return number


def _parse_money(value) -> float:
    cleaned = str(value or "").strip().replace("$", "").replace(" ", "")
    if not cleaned:
        return 0.0
    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+", cleaned):
        # 350.000 -> 350000 (punto como separador de miles)
        cleaned = cleaned.replace(".", "")
    elif "," in cleaned:
        cleaned = cleaned.replace(",", ".")
    try:
        amount = Decimal(cleaned)
    except InvalidOperation as exc:
        raise ValueError(f"Valor numérico inválido: {value}") from exc
    if amount < 0:
        raise ValueError("El costo no puede ser negativo.")
    return float(amount)


def _form_values() -> dict:
    code = request.form.get("code", "").strip()
    name = request.form.get("name", "").strip()
    if not code:
        raise ValueError("El código es obligatorio.")
    if not name:
        raise ValueError("El nombre de la herramienta es obligatorio.")
    status = request.form.get("status", "Disponible").strip()
    if status not in TOOL_STATUSES:
        raise ValueError("Estado inválido.")
    return {
        "code": code,
        "name": name,
        "category": request.form.get("category", "").strip(),
        "brand": request.form.get("brand", "").strip(),
        "serial": request.form.get("serial", "").strip(),
        "unit_cost": _parse_money(request.form.get("unit_cost")),
        "location": request.form.get("location", "").strip(),
        "responsible": request.form.get("responsible", "").strip(),
        "status": status,
        "notes": request.form.get("notes", "").strip(),
    }


def _record_movement(conn, tool_id: int, movement_type: str, delta: int,
                     previous: int, notes: str = "") -> None:
    conn.execute(
        """INSERT INTO tool_movements(
               tool_id, movement_type, quantity, previous_quantity,
               new_quantity, notes, user_id, created_at
           ) VALUES(?,?,?,?,?,?,?,?)""",
        (tool_id, movement_type, delta, previous, previous + delta,
         notes, g.user["id"], _now()),
    )


def _get_tool(conn, tool_id: int, for_update: bool = False):
    sql = "SELECT * FROM tools WHERE id=?"
    if for_update and db.engine.dialect.name == "postgresql":
        sql += " FOR UPDATE"
    return conn.execute(sql, (tool_id,)).fetchone()


# ============================================================
# LISTADO
# ============================================================

@tools_bp.route("/")
@login_required
def index():
    q = request.args.get("q", "").strip()
    status = request.args.get("status", "").strip()

    filters = ["active=1"]
    params: list = []
    if q:
        filters.append(
            "(code LIKE ? OR name LIKE ? OR category LIKE ? OR brand LIKE ? "
            "OR serial LIKE ? OR location LIKE ? OR responsible LIKE ?)"
        )
        params.extend([f"%{q}%"] * 7)
    if status in TOOL_STATUSES:
        filters.append("status=?")
        params.append(status)

    where = " AND ".join(filters)

    with db_conn() as conn:
        rows = conn.execute(
            f"SELECT * FROM tools WHERE {where} ORDER BY category, name",
            tuple(params),
        ).fetchall()

        summary = conn.execute(
            """SELECT
                   COUNT(*) AS references_count,
                   COALESCE(SUM(quantity), 0) AS units,
                   COALESCE(SUM(quantity * unit_cost), 0) AS total_value
               FROM tools WHERE active=1"""
        ).fetchone()

        loss_placeholders = ",".join("?" for _ in LOSS_TYPES)
        losses = conn.execute(
            f"""SELECT
                    COALESCE(SUM(-m.quantity), 0) AS units,
                    COALESCE(SUM(-m.quantity * t.unit_cost), 0) AS value
                FROM tool_movements m
                JOIN tools t ON t.id = m.tool_id
                WHERE m.movement_type IN ({loss_placeholders})""",
            LOSS_TYPES,
        ).fetchone()

        last_count = conn.execute(
            "SELECT MAX(last_counted_at) AS last FROM tools WHERE active=1"
        ).fetchone()

    return render_template(
        "tools.html",
        tools=rows,
        q=q,
        status=status,
        statuses=TOOL_STATUSES,
        summary=summary,
        losses=losses,
        last_count=(last_count or {}).get("last"),
    )


# ============================================================
# CREAR / EDITAR
# ============================================================

@tools_bp.route("/nueva", methods=["GET", "POST"])
@admin_required
def new():
    if request.method == "POST":
        try:
            values = _form_values()
            quantity = _parse_int(request.form.get("quantity"), "La cantidad", minimum=0)
            with db_conn() as conn:
                result = conn.execute(
                    """INSERT INTO tools(
                           code, name, category, brand, serial, quantity, unit_cost,
                           location, responsible, status, notes, active,
                           created_at, updated_at
                       ) VALUES(?,?,?,?,?,?,?,?,?,?,?,1,?,?)""",
                    (values["code"], values["name"], values["category"], values["brand"],
                     values["serial"], quantity, values["unit_cost"], values["location"],
                     values["responsible"], values["status"], values["notes"],
                     _now(), _now()),
                )
                tool_id = result.lastrowid
                if tool_id is None:
                    tool_id = conn.execute(
                        "SELECT id FROM tools WHERE code=?", (values["code"],)
                    ).fetchone()["id"]
                if quantity:
                    _record_movement(conn, tool_id, "Registro inicial", quantity, 0)
            flash("Herramienta registrada correctamente.", "success")
            return redirect(url_for("tools.index"))
        except IntegrityError:
            flash("Ya existe una herramienta con ese código.", "danger")
        except ValueError as exc:
            flash(str(exc), "danger")

    return render_template("tool_form.html", tool=None, statuses=TOOL_STATUSES)


@tools_bp.route("/<int:tool_id>/editar", methods=["GET", "POST"])
@admin_required
def edit(tool_id: int):
    with db_conn() as conn:
        tool = _get_tool(conn, tool_id)
    if not tool or not tool["active"]:
        flash("Herramienta no encontrada.", "danger")
        return redirect(url_for("tools.index"))

    if request.method == "POST":
        try:
            values = _form_values()
            with db_conn() as conn:
                conn.execute(
                    """UPDATE tools SET
                           code=?, name=?, category=?, brand=?, serial=?, unit_cost=?,
                           location=?, responsible=?, status=?, notes=?, updated_at=?
                       WHERE id=?""",
                    (values["code"], values["name"], values["category"], values["brand"],
                     values["serial"], values["unit_cost"], values["location"],
                     values["responsible"], values["status"], values["notes"],
                     _now(), tool_id),
                )
            flash("Herramienta actualizada.", "success")
            return redirect(url_for("tools.detail", tool_id=tool_id))
        except IntegrityError:
            flash("Ese código ya está asignado a otra herramienta.", "danger")
        except ValueError as exc:
            flash(str(exc), "danger")

    return render_template("tool_form.html", tool=tool, statuses=TOOL_STATUSES)


# ============================================================
# DETALLE + MOVIMIENTOS
# ============================================================

@tools_bp.route("/<int:tool_id>")
@login_required
def detail(tool_id: int):
    with db_conn() as conn:
        tool = _get_tool(conn, tool_id)
        if not tool:
            flash("Herramienta no encontrada.", "danger")
            return redirect(url_for("tools.index"))
        movements = conn.execute(
            """SELECT m.*, u.full_name AS user_name
               FROM tool_movements m
               LEFT JOIN users u ON u.id = m.user_id
               WHERE m.tool_id=?
               ORDER BY m.created_at DESC, m.id DESC""",
            (tool_id,),
        ).fetchall()

    return render_template(
        "tool_detail.html",
        tool=tool,
        movements=movements,
        movement_types=list(MOVEMENT_TYPES),
        loss_types=LOSS_TYPES,
    )


@tools_bp.route("/<int:tool_id>/movimiento", methods=["POST"])
@admin_required
def movement(tool_id: int):
    try:
        movement_type = request.form.get("movement_type", "").strip()
        if movement_type not in MOVEMENT_TYPES:
            raise ValueError("Tipo de movimiento inválido.")
        quantity = _parse_int(request.form.get("quantity"), "La cantidad", minimum=1)
        notes = request.form.get("notes", "").strip()
        if MOVEMENT_TYPES[movement_type] < 0 and not notes:
            raise ValueError("Describe qué pasó (por ejemplo, quién la tenía o dónde se vio por última vez).")

        delta = quantity * MOVEMENT_TYPES[movement_type]

        with db_conn() as conn:
            tool = _get_tool(conn, tool_id, for_update=True)
            if not tool or not tool["active"]:
                raise ValueError("Herramienta no encontrada.")
            previous = int(tool["quantity"] or 0)
            if previous + delta < 0:
                raise ValueError(f"Solo hay {previous} unidad(es) registradas.")
            conn.execute(
                "UPDATE tools SET quantity=?, updated_at=? WHERE id=?",
                (previous + delta, _now(), tool_id),
            )
            _record_movement(conn, tool_id, movement_type, delta, previous, notes)

        flash(f"{movement_type} registrada. Cantidad actual: {previous + delta}.", "success")
    except ValueError as exc:
        flash(str(exc), "danger")

    return redirect(url_for("tools.detail", tool_id=tool_id))


# ============================================================
# CONTEO FÍSICO (DETECCIÓN DE PÉRDIDAS)
# ============================================================

@tools_bp.route("/conteo", methods=["GET", "POST"])
@admin_required
def count():
    if request.method == "POST":
        general_notes = request.form.get("notes", "").strip()
        missing = extra = checked = 0
        missing_names: list[str] = []
        try:
            with db_conn() as conn:
                tools = conn.execute(
                    "SELECT id FROM tools WHERE active=1 ORDER BY id"
                ).fetchall()
                for row in tools:
                    raw = request.form.get(f"counted_{row['id']}", "").strip()
                    if raw == "":
                        continue  # no se contó esta herramienta
                    counted = _parse_int(raw, "La cantidad contada", minimum=0)
                    tool = _get_tool(conn, row["id"], for_update=True)
                    previous = int(tool["quantity"] or 0)
                    delta = counted - previous
                    checked += 1
                    if delta < 0:
                        movement_type = "Faltante en conteo"
                        missing += -delta
                        missing_names.append(f"{tool['name']} ({-delta})")
                    elif delta > 0:
                        movement_type = "Sobrante en conteo"
                        extra += delta
                    else:
                        movement_type = "Conteo sin novedad"
                    _record_movement(conn, row["id"], movement_type, delta, previous, general_notes)
                    conn.execute(
                        "UPDATE tools SET quantity=?, last_counted_at=?, updated_at=? WHERE id=?",
                        (counted, _now(), _now(), row["id"]),
                    )
        except ValueError as exc:
            flash(str(exc), "danger")
            return redirect(url_for("tools.count"))

        if not checked:
            flash("No se ingresó ninguna cantidad contada.", "warning")
            return redirect(url_for("tools.count"))

        if missing:
            flash(
                f"Conteo guardado ({checked} herramientas revisadas). "
                f"FALTAN {missing} unidad(es): " + ", ".join(missing_names),
                "danger",
            )
        else:
            flash(
                f"Conteo guardado ({checked} herramientas revisadas). No falta nada."
                + (f" Aparecieron {extra} unidad(es) de más." if extra else ""),
                "success",
            )
        return redirect(url_for("tools.index"))

    with db_conn() as conn:
        tools = conn.execute(
            "SELECT * FROM tools WHERE active=1 ORDER BY location, category, name"
        ).fetchall()
    return render_template("tool_count.html", tools=tools)


# ============================================================
# HISTORIAL DE PÉRDIDAS
# ============================================================

@tools_bp.route("/perdidas")
@login_required
def losses():
    placeholders = ",".join("?" for _ in LOSS_TYPES)
    with db_conn() as conn:
        rows = conn.execute(
            f"""SELECT m.*, t.code, t.name, t.unit_cost, t.responsible,
                       u.full_name AS user_name
                FROM tool_movements m
                JOIN tools t ON t.id = m.tool_id
                LEFT JOIN users u ON u.id = m.user_id
                WHERE m.movement_type IN ({placeholders})
                ORDER BY m.created_at DESC, m.id DESC""",
            LOSS_TYPES,
        ).fetchall()
    total_units = sum(-int(r["quantity"] or 0) for r in rows)
    total_value = sum(-int(r["quantity"] or 0) * float(r["unit_cost"] or 0) for r in rows)
    return render_template(
        "tool_losses.html",
        rows=rows,
        total_units=total_units,
        total_value=total_value,
    )


# ============================================================
# ELIMINAR
# ============================================================

@tools_bp.route("/<int:tool_id>/eliminar", methods=["POST"])
@admin_required
def delete(tool_id: int):
    with db_conn() as conn:
        tool = _get_tool(conn, tool_id)
        if not tool:
            flash("Herramienta no encontrada.", "danger")
            return redirect(url_for("tools.index"))
        history = conn.execute(
            """SELECT COUNT(*) AS total FROM tool_movements
               WHERE tool_id=? AND movement_type <> 'Registro inicial'""",
            (tool_id,),
        ).fetchone()
        if int(history["total"] or 0) == 0:
            conn.execute("DELETE FROM tool_movements WHERE tool_id=?", (tool_id,))
            conn.execute("DELETE FROM tools WHERE id=?", (tool_id,))
            flash(f"Herramienta '{tool['name']}' eliminada.", "success")
        else:
            # Con historial se oculta para no perder el registro de pérdidas.
            conn.execute(
                "UPDATE tools SET active=0, updated_at=? WHERE id=?",
                (_now(), tool_id),
            )
            flash(
                f"Herramienta '{tool['name']}' retirada del inventario. Se conservó su historial.",
                "success",
            )
    return redirect(url_for("tools.index"))


# ============================================================
# EXPORTAR
# ============================================================

@tools_bp.route("/exportar-csv")
@login_required
def export_csv():
    with db_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM tools WHERE active=1 ORDER BY category, name"
        ).fetchall()

    output = StringIO()
    output.write("﻿")
    writer = csv.writer(output, delimiter=";", lineterminator="\n")
    writer.writerow([
        "Código", "Herramienta", "Categoría", "Marca", "Serial", "Cantidad",
        "Costo unitario", "Ubicación", "Responsable", "Estado", "Último conteo",
        "Observaciones",
    ])
    for t in rows:
        writer.writerow([
            t["code"], t["name"], t["category"], t["brand"], t["serial"], t["quantity"],
            t["unit_cost"], t["location"], t["responsible"], t["status"],
            t["last_counted_at"] or "", t["notes"],
        ])

    stamp = datetime.now().strftime("%Y%m%d")
    return Response(
        output.getvalue(),
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename=herramientas_{stamp}.csv"},
    )
