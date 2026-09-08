from __future__ import annotations

from datetime import datetime, date, timedelta
from decimal import Decimal, InvalidOperation
from functools import wraps

from flask import (
    Blueprint,
    flash,
    g,
    redirect,
    render_template,
    request,
    url_for,
)
from sqlalchemy import func

from database import db, db_conn


# ============================================================
# BLUEPRINT
# ============================================================

expenses_bp = Blueprint(
    "expenses",
    __name__,
    url_prefix="/gastos"
)


# ============================================================
# MODELO
# ============================================================

class Expense(db.Model):
    __tablename__ = "expenses"

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    expense_date = db.Column(
        db.Date,
        nullable=False
    )

    category = db.Column(
        db.String(120),
        nullable=False
    )

    description = db.Column(
        db.String(300),
        nullable=False
    )

    amount = db.Column(
        db.Numeric(16, 2),
        nullable=False,
        default=0
    )

    payment_method = db.Column(
        db.String(80),
        nullable=False,
        default="Transferencia"
    )

    expense_type = db.Column(
        db.String(30),
        nullable=False,
        default="Variable"
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

    updated_at = db.Column(
        db.DateTime,
        nullable=False,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        server_default=func.now()
    )


# ============================================================
# CATÁLOGOS
# ============================================================

EXPENSE_CATEGORIES = [
    "Arriendo",
    "Salarios",
    "Seguridad social",
    "Servicios públicos",
    "Transporte",
    "Publicidad",
    "Contabilidad",
    "Mantenimiento",
    "Papelería",
    "Software",
    "Impuestos",
    "Aseo",
    "Alimentación",
    "Comisiones",
    "Otros",
]


PAYMENT_METHODS = [
    "Transferencia",
    "Efectivo",
    "Tarjeta",
    "Otro",
]


EXPENSE_TYPES = [
    "Fijo",
    "Variable",
]


# ============================================================
# SEGURIDAD
# ============================================================

def admin_required(view):

    @wraps(view)
    def wrapped(*args, **kwargs):

        if g.user is None:

            flash(
                "Debes iniciar sesión.",
                "warning"
            )

            return redirect(
                url_for("login")
            )

        if (
            str(
                g.user["role"]
                or
                ""
            ).strip().lower()
            !=
            "administrador"
        ):

            flash(
                "Solo el administrador puede acceder al módulo de gastos.",
                "danger"
            )

            return redirect(
                url_for("dashboard")
            )

        return view(
            *args,
            **kwargs
        )

    return wrapped


# ============================================================
# UTILIDADES
# ============================================================

def parse_amount(value):

    value = str(
        value
        or
        ""
    ).strip()

    value = (
        value
        .replace("$", "")
        .replace(" ", "")
    )

    # Si viene 1.500.000
    if (
        "." in value
        and
        "," not in value
    ):
        parts = value.split(".")

        if all(
            len(part) == 3
            for part in parts[1:]
        ):
            value = "".join(parts)

    # Formato colombiano 1.500.000,50
    elif (
        "." in value
        and
        "," in value
    ):
        value = (
            value
            .replace(".", "")
            .replace(",", ".")
        )

    else:
        value = value.replace(
            ",",
            "."
        )

    try:

        return Decimal(
            value
            or
            "0"
        )

    except InvalidOperation as exc:

        raise ValueError(
            "El valor del gasto no es válido."
        ) from exc


def parse_date(value):

    try:

        return datetime.strptime(
            value,
            "%Y-%m-%d"
        ).date()

    except Exception as exc:

        raise ValueError(
            "La fecha ingresada no es válida."
        ) from exc


def first_day_month(value: date):

    return value.replace(
        day=1
    )


def last_day_month(value: date):

    if value.month == 12:

        next_month = date(
            value.year + 1,
            1,
            1
        )

    else:

        next_month = date(
            value.year,
            value.month + 1,
            1
        )

    return (
        next_month
        -
        timedelta(days=1)
    )


# ============================================================
# LISTADO / DASHBOARD DE GASTOS
# ============================================================

@expenses_bp.route("/")
@admin_required
def list_expenses():

    today = date.today()

    date_from = (
        request.args.get(
            "date_from",
            ""
        )
        .strip()
    )

    date_to = (
        request.args.get(
            "date_to",
            ""
        )
        .strip()
    )

    category = (
        request.args.get(
            "category",
            ""
        )
        .strip()
    )

    expense_type = (
        request.args.get(
            "expense_type",
            ""
        )
        .strip()
    )

    q = (
        request.args.get(
            "q",
            ""
        )
        .strip()
    )


    # ========================================================
    # PERÍODO POR DEFECTO: MES ACTUAL
    # ========================================================

    if not date_from:

        date_from = (
            first_day_month(
                today
            )
            .isoformat()
        )

    if not date_to:

        date_to = (
            last_day_month(
                today
            )
            .isoformat()
        )


    filters = [
        "expense_date >= ?",
        "expense_date <= ?"
    ]

    params = [
        date_from,
        date_to
    ]


    if category:

        filters.append(
            "category = ?"
        )

        params.append(
            category
        )


    if expense_type:

        filters.append(
            "expense_type = ?"
        )

        params.append(
            expense_type
        )


    if q:

        filters.append(
            """
            (
                LOWER(description) LIKE LOWER(?)
                OR
                LOWER(category) LIKE LOWER(?)
                OR
                LOWER(COALESCE(notes,'')) LIKE LOWER(?)
            )
            """
        )

        search = (
            "%"
            +
            q
            +
            "%"
        )

        params.extend(
            [
                search,
                search,
                search
            ]
        )


    where_clause = (
        " WHERE "
        +
        " AND ".join(
            filters
        )
    )


    with db_conn() as conn:

        expenses = conn.execute(
            f"""
            SELECT
                e.*,
                u.full_name AS user_name

            FROM expenses e

            LEFT JOIN users u
                ON u.id = e.user_id

            {where_clause}

            ORDER BY
                expense_date DESC,
                e.id DESC
            """,
            params
        ).fetchall()


        totals = conn.execute(
            f"""
            SELECT

                COALESCE(
                    SUM(amount),
                    0
                )
                AS total_expenses,

                COALESCE(
                    SUM(
                        CASE
                            WHEN expense_type='Fijo'
                            THEN amount
                            ELSE 0
                        END
                    ),
                    0
                )
                AS fixed_expenses,

                COALESCE(
                    SUM(
                        CASE
                            WHEN expense_type='Variable'
                            THEN amount
                            ELSE 0
                        END
                    ),
                    0
                )
                AS variable_expenses

            FROM expenses

            {where_clause}
            """,
            params
        ).fetchone()


        categories_summary = conn.execute(
            f"""
            SELECT

                category,

                COUNT(*) AS quantity,

                COALESCE(
                    SUM(amount),
                    0
                )
                AS total

            FROM expenses

            {where_clause}

            GROUP BY category

            ORDER BY total DESC
            """,
            params
        ).fetchall()


    total_expenses = Decimal(
        str(
            totals["total_expenses"]
            or
            0
        )
    )

    fixed_expenses = Decimal(
        str(
            totals["fixed_expenses"]
            or
            0
        )
    )

    variable_expenses = Decimal(
        str(
            totals["variable_expenses"]
            or
            0
        )
    )


    return render_template(
        "expenses.html",

        expenses=expenses,

        categories=EXPENSE_CATEGORIES,

        expense_types=EXPENSE_TYPES,

        categories_summary=categories_summary,

        date_from=date_from,

        date_to=date_to,

        category_filter=category,

        expense_type_filter=expense_type,

        q=q,

        total_expenses=total_expenses,

        fixed_expenses=fixed_expenses,

        variable_expenses=variable_expenses,
    )


# ============================================================
# NUEVO GASTO
# ============================================================

@expenses_bp.route(
    "/nuevo",
    methods=["GET", "POST"]
)
@admin_required
def new_expense():

    if request.method == "POST":

        try:

            expense_date = parse_date(
                request.form.get(
                    "expense_date",
                    ""
                )
            )

            category = (
                request.form.get(
                    "category",
                    ""
                )
                .strip()
            )

            description = (
                request.form.get(
                    "description",
                    ""
                )
                .strip()
            )

            amount = parse_amount(
                request.form.get(
                    "amount",
                    ""
                )
            )

            payment_method = (
                request.form.get(
                    "payment_method",
                    "Transferencia"
                )
                .strip()
            )

            expense_type = (
                request.form.get(
                    "expense_type",
                    "Variable"
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


            if (
                category
                not in
                EXPENSE_CATEGORIES
            ):

                raise ValueError(
                    "Selecciona una categoría válida."
                )


            if not description:

                raise ValueError(
                    "La descripción es obligatoria."
                )


            if amount <= 0:

                raise ValueError(
                    "El valor del gasto debe ser mayor a cero."
                )


            if (
                payment_method
                not in
                PAYMENT_METHODS
            ):

                raise ValueError(
                    "La forma de pago no es válida."
                )


            if (
                expense_type
                not in
                EXPENSE_TYPES
            ):

                raise ValueError(
                    "El tipo de gasto no es válido."
                )


            with db_conn() as conn:

                conn.execute(
                    """
                    INSERT INTO expenses(
                        expense_date,
                        category,
                        description,
                        amount,
                        payment_method,
                        expense_type,
                        notes,
                        user_id,
                        created_at,
                        updated_at
                    )
                    VALUES(
                        ?,?,?,?,?,?,?,?,
                        CURRENT_TIMESTAMP,
                        CURRENT_TIMESTAMP
                    )
                    """,
                    (
                        expense_date,
                        category,
                        description,
                        float(amount),
                        payment_method,
                        expense_type,
                        notes,
                        g.user["id"]
                    )
                )


            flash(
                "Gasto registrado correctamente.",
                "success"
            )

            return redirect(
                url_for(
                    "expenses.list_expenses"
                )
            )


        except ValueError as exc:

            flash(
                str(exc),
                "danger"
            )


    return render_template(
        "expense_form.html",

        expense=None,

        categories=EXPENSE_CATEGORIES,

        payment_methods=PAYMENT_METHODS,

        expense_types=EXPENSE_TYPES,

        today=date.today().isoformat(),
    )


# ============================================================
# EDITAR GASTO
# ============================================================

@expenses_bp.route(
    "/<int:expense_id>/editar",
    methods=["GET", "POST"]
)
@admin_required
def edit_expense(expense_id):

    with db_conn() as conn:

        expense = conn.execute(
            """
            SELECT *
            FROM expenses
            WHERE id=?
            """,
            (
                expense_id,
            )
        ).fetchone()


    if not expense:

        flash(
            "El gasto no existe.",
            "danger"
        )

        return redirect(
            url_for(
                "expenses.list_expenses"
            )
        )


    if request.method == "POST":

        try:

            expense_date = parse_date(
                request.form.get(
                    "expense_date",
                    ""
                )
            )

            category = (
                request.form.get(
                    "category",
                    ""
                )
                .strip()
            )

            description = (
                request.form.get(
                    "description",
                    ""
                )
                .strip()
            )

            amount = parse_amount(
                request.form.get(
                    "amount",
                    ""
                )
            )

            payment_method = (
                request.form.get(
                    "payment_method",
                    ""
                )
                .strip()
            )

            expense_type = (
                request.form.get(
                    "expense_type",
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


            if (
                category
                not in
                EXPENSE_CATEGORIES
            ):

                raise ValueError(
                    "Selecciona una categoría válida."
                )


            if not description:

                raise ValueError(
                    "La descripción es obligatoria."
                )


            if amount <= 0:

                raise ValueError(
                    "El valor debe ser mayor a cero."
                )


            if (
                payment_method
                not in
                PAYMENT_METHODS
            ):

                raise ValueError(
                    "La forma de pago no es válida."
                )


            if (
                expense_type
                not in
                EXPENSE_TYPES
            ):

                raise ValueError(
                    "El tipo de gasto no es válido."
                )


            with db_conn() as conn:

                conn.execute(
                    """
                    UPDATE expenses

                    SET
                        expense_date=?,
                        category=?,
                        description=?,
                        amount=?,
                        payment_method=?,
                        expense_type=?,
                        notes=?,
                        updated_at=CURRENT_TIMESTAMP

                    WHERE id=?
                    """,
                    (
                        expense_date,
                        category,
                        description,
                        float(amount),
                        payment_method,
                        expense_type,
                        notes,
                        expense_id
                    )
                )


            flash(
                "Gasto actualizado correctamente.",
                "success"
            )

            return redirect(
                url_for(
                    "expenses.list_expenses"
                )
            )


        except ValueError as exc:

            flash(
                str(exc),
                "danger"
            )


    return render_template(
        "expense_form.html",

        expense=expense,

        categories=EXPENSE_CATEGORIES,

        payment_methods=PAYMENT_METHODS,

        expense_types=EXPENSE_TYPES,

        today=date.today().isoformat(),
    )


# ============================================================
# ELIMINAR
# ============================================================

@expenses_bp.route(
    "/<int:expense_id>/eliminar",
    methods=["POST"]
)
@admin_required
def delete_expense(expense_id):

    with db_conn() as conn:

        expense = conn.execute(
            """
            SELECT id, description
            FROM expenses
            WHERE id=?
            """,
            (
                expense_id,
            )
        ).fetchone()


        if not expense:

            flash(
                "El gasto no existe.",
                "danger"
            )

            return redirect(
                url_for(
                    "expenses.list_expenses"
                )
            )


        conn.execute(
            """
            DELETE FROM expenses
            WHERE id=?
            """,
            (
                expense_id,
            )
        )


    flash(
        "Gasto eliminado correctamente.",
        "success"
    )

    return redirect(
        url_for(
            "expenses.list_expenses"
        )
    )