from __future__ import annotations

from datetime import date, datetime, timedelta
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

from database import db, db_conn


# ============================================================
# BLUEPRINT
# ============================================================

expenses_bp = Blueprint(
    "expenses",
    __name__,
    url_prefix="/gastos",
)


# ============================================================
# MODELO
# ============================================================

class Expense(db.Model):

    __tablename__ = "expenses"

    id = db.Column(
        db.Integer,
        primary_key=True,
    )

    expense_date = db.Column(
        db.Date,
        nullable=False,
        index=True,
    )

    category = db.Column(
        db.String(100),
        nullable=False,
        index=True,
    )

    description = db.Column(
        db.String(250),
        nullable=False,
    )

    amount = db.Column(
        db.Numeric(14, 2),
        nullable=False,
        default=0,
    )

    payment_method = db.Column(
        db.String(80),
        nullable=True,
    )

    expense_type = db.Column(
        db.String(30),
        nullable=False,
        default="Variable",
        index=True,
    )

    notes = db.Column(
        db.Text,
        nullable=True,
    )

    user_id = db.Column(
        db.Integer,
        nullable=True,
    )

    created_at = db.Column(
        db.DateTime,
        nullable=False,
        default=datetime.utcnow,
    )

    updated_at = db.Column(
        db.DateTime,
        nullable=False,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
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


EXPENSE_TYPES = [
    "Fijo",
    "Variable",
]


PAYMENT_METHODS = [

    "Efectivo",

    "Transferencia",

    "Tarjeta",

    "Débito automático",

    "Otro",
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
                "warning",
            )

            return redirect(
                url_for("login")
            )


        role = str(
            g.user["role"]
            or
            ""
        ).strip().lower()


        # Permitimos ambos nombres por compatibilidad
        # con las distintas versiones del proyecto.

        if role not in (
            "admin",
            "administrador",
        ):

            flash(
                "Solo el administrador puede acceder al módulo de gastos.",
                "danger",
            )

            return redirect(
                url_for("dashboard")
            )


        return view(
            *args,
            **kwargs,
        )

    return wrapped


# ============================================================
# UTILIDADES
# ============================================================

def parse_date(value):

    if not value:

        return None

    try:

        return datetime.strptime(
            str(value)[:10],
            "%Y-%m-%d",
        ).date()

    except ValueError:

        return None


def parse_amount(value):

    if value is None:

        return None


    raw = str(value).strip()


    if not raw:

        return None


    raw = (
        raw
        .replace("$", "")
        .replace(" ", "")
    )


    # ========================================================
    # SOPORTE PARA FORMATOS:
    #
    # 1500000
    # 1500000.50
    # 1.500.000
    # 1.500.000,50
    # ========================================================

    if "," in raw and "." in raw:

        if raw.rfind(",") > raw.rfind("."):

            raw = (
                raw
                .replace(".", "")
                .replace(",", ".")
            )

        else:

            raw = (
                raw
                .replace(",", "")
            )


    elif "," in raw:

        parts = raw.split(",")

        if (
            len(parts) == 2
            and
            len(parts[1]) <= 2
        ):

            raw = raw.replace(
                ",",
                ".",
            )

        else:

            raw = raw.replace(
                ",",
                "",
            )


    elif raw.count(".") > 1:

        raw = raw.replace(
            ".",
            "",
        )


    try:

        amount = Decimal(raw)

    except (
        InvalidOperation,
        ValueError,
    ):

        return None


    if amount < 0:

        return None


    return amount


def first_day_month(value):

    return value.replace(
        day=1
    )


def last_day_month(value):

    if value.month == 12:

        next_month = date(
            value.year + 1,
            1,
            1,
        )

    else:

        next_month = date(
            value.year,
            value.month + 1,
            1,
        )


    return (
        next_month
        -
        timedelta(days=1)
    )


def current_user_id():

    try:

        return g.user["id"]

    except Exception:

        return None


# ============================================================
# COMMIT COMPATIBLE
# ============================================================

def commit_connection(conn):

    """
    Confirma los INSERT / UPDATE / DELETE.

    Esto es importante especialmente en PostgreSQL
    de Render.

    La función soporta tanto una conexión directa
    como el CompatConnection utilizado por database.py.
    """

    # ========================================================
    # Caso 1:
    # el wrapper tiene commit()
    # ========================================================

    if hasattr(
        conn,
        "commit",
    ):

        conn.commit()

        return


    # ========================================================
    # Caso 2:
    # CompatConnection contiene
    # una conexión SQLAlchemy en .connection
    # ========================================================

    inner_connection = getattr(
        conn,
        "connection",
        None,
    )


    if (
        inner_connection is not None
        and
        hasattr(
            inner_connection,
            "commit",
        )
    ):

        inner_connection.commit()

        return


    raise RuntimeError(
        "No fue posible confirmar la transacción en la base de datos."
    )


# ============================================================
# FILTROS
# ============================================================

def get_filters():

    today = date.today()


    default_from = first_day_month(
        today
    )

    default_to = last_day_month(
        today
    )


    date_from_raw = (
        request.args.get(
            "date_from",
            ""
        )
        .strip()
    )

    date_to_raw = (
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


    date_from = (
        parse_date(
            date_from_raw
        )
        or
        default_from
    )


    date_to = (
        parse_date(
            date_to_raw
        )
        or
        default_to
    )


    if date_to < date_from:

        date_from, date_to = (
            date_to,
            date_from,
        )


    return {

        "date_from": date_from,

        "date_to": date_to,

        "category": category,

        "expense_type": expense_type,

        "q": q,
    }


# ============================================================
# LISTADO DE GASTOS
# ============================================================

@expenses_bp.route("/")
@admin_required
def list_expenses():

    filters = get_filters()


    date_from = filters[
        "date_from"
    ]

    date_to = filters[
        "date_to"
    ]

    category = filters[
        "category"
    ]

    expense_type = filters[
        "expense_type"
    ]

    q = filters[
        "q"
    ]


    # ========================================================
    # WHERE DINÁMICO
    # ========================================================

    where = [

        "expense_date >= ?",

        "expense_date <= ?",
    ]


    params = [

        date_from.isoformat(),

        date_to.isoformat(),
    ]


    if category:

        where.append(
            "category = ?"
        )

        params.append(
            category
        )


    if expense_type:

        where.append(
            "expense_type = ?"
        )

        params.append(
            expense_type
        )


    if q:

        where.append(
            """
            (
                LOWER(description) LIKE ?
                OR
                LOWER(category) LIKE ?
                OR
                LOWER(COALESCE(notes, '')) LIKE ?
            )
            """
        )


        search_value = (
            "%"
            +
            q.lower()
            +
            "%"
        )


        params.extend(
            [
                search_value,
                search_value,
                search_value,
            ]
        )


    where_sql = (
        " AND ".join(
            where
        )
    )


    with db_conn() as conn:


        # ====================================================
        # DETALLE
        # ====================================================

        expenses = conn.execute(
            f"""
            SELECT

                id,

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

            FROM expenses

            WHERE
                {where_sql}

            ORDER BY
                expense_date DESC,
                id DESC
            """,
            tuple(params),
        ).fetchall()


        # ====================================================
        # KPIs
        # ====================================================

        totals = conn.execute(
            f"""
            SELECT

                COALESCE(
                    SUM(amount),
                    0
                ) AS total,

                COALESCE(
                    SUM(
                        CASE

                            WHEN expense_type = 'Fijo'
                            THEN amount

                            ELSE 0

                        END
                    ),
                    0
                ) AS fixed_total,

                COALESCE(
                    SUM(
                        CASE

                            WHEN expense_type = 'Variable'
                            THEN amount

                            ELSE 0

                        END
                    ),
                    0
                ) AS variable_total

            FROM expenses

            WHERE
                {where_sql}
            """,
            tuple(params),
        ).fetchone()


        # ====================================================
        # DISTRIBUCIÓN POR CATEGORÍA
        # ====================================================

        categories_summary = conn.execute(
            f"""
            SELECT

                category,

                COUNT(*) AS quantity,

                COALESCE(
                    SUM(amount),
                    0
                ) AS total

            FROM expenses

            WHERE
                {where_sql}

            GROUP BY
                category

            ORDER BY
                total DESC
            """,
            tuple(params),
        ).fetchall()


    # ========================================================
    # TOTALES
    # ========================================================

    if totals:

        total_expenses = Decimal(
            str(
                totals["total"]
                or
                0
            )
        )

        fixed_expenses = Decimal(
            str(
                totals["fixed_total"]
                or
                0
            )
        )

        variable_expenses = Decimal(
            str(
                totals["variable_total"]
                or
                0
            )
        )

    else:

        total_expenses = Decimal("0")

        fixed_expenses = Decimal("0")

        variable_expenses = Decimal("0")


    return render_template(
        "expenses.html",

        expenses=expenses,

        categories_summary=(
            categories_summary
        ),

        total_expenses=(
            total_expenses
        ),

        fixed_expenses=(
            fixed_expenses
        ),

        variable_expenses=(
            variable_expenses
        ),

        categories=(
            EXPENSE_CATEGORIES
        ),

        expense_types=(
            EXPENSE_TYPES
        ),

        date_from=(
            date_from.isoformat()
        ),

        date_to=(
            date_to.isoformat()
        ),

        selected_category=(
            category
        ),

        selected_type=(
            expense_type
        ),

        q=q,
    )


# ============================================================
# NUEVO GASTO
# ============================================================

@expenses_bp.route(
    "/nuevo",
    methods=[
        "GET",
        "POST",
    ],
)
@admin_required
def new_expense():

    if request.method == "POST":


        expense_date = parse_date(
            request.form.get(
                "expense_date"
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
                "amount"
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


        # ====================================================
        # VALIDACIONES
        # ====================================================

        errors = []


        if not expense_date:

            errors.append(
                "Debes seleccionar la fecha del gasto."
            )


        if not category:

            errors.append(
                "Debes seleccionar una categoría."
            )


        elif category not in EXPENSE_CATEGORIES:

            errors.append(
                "La categoría seleccionada no es válida."
            )


        if not description:

            errors.append(
                "Debes ingresar una descripción."
            )


        if (
            amount is None
            or
            amount <= 0
        ):

            errors.append(
                "El valor del gasto debe ser mayor que cero."
            )


        if not expense_type:

            errors.append(
                "Debes seleccionar si el gasto es fijo o variable."
            )


        elif expense_type not in EXPENSE_TYPES:

            errors.append(
                "El tipo de gasto seleccionado no es válido."
            )


        if (
            payment_method
            and
            payment_method not in PAYMENT_METHODS
        ):

            errors.append(
                "El método de pago seleccionado no es válido."
            )


        # ====================================================
        # MOSTRAR ERRORES
        # ====================================================

        if errors:

            for error in errors:

                flash(
                    error,
                    "danger",
                )


            return render_template(
                "expense_form.html",

                expense=None,

                categories=(
                    EXPENSE_CATEGORIES
                ),

                expense_types=(
                    EXPENSE_TYPES
                ),

                payment_methods=(
                    PAYMENT_METHODS
                ),

                today=date.today().isoformat(),
            )


        # ====================================================
        # GUARDAR
        # ====================================================

        now = datetime.now()


        try:

            with db_conn() as conn:

                conn.execute(
                    """
                    INSERT INTO expenses (
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
                    VALUES (
                        ?,
                        ?,
                        ?,
                        ?,
                        ?,
                        ?,
                        ?,
                        ?,
                        ?,
                        ?
                    )
                    """,
                    (
                        expense_date.isoformat(),

                        category,

                        description,

                        amount,

                        payment_method
                        or
                        None,

                        expense_type,

                        notes
                        or
                        None,

                        current_user_id(),

                        now,

                        now,
                    ),
                )


                # ============================================
                # MUY IMPORTANTE PARA POSTGRESQL / RENDER
                # ============================================

                commit_connection(
                    conn
                )


        except Exception as exc:

            print(
                "ERROR REGISTRANDO GASTO:",
                repr(exc),
            )


            flash(
                "No fue posible registrar el gasto. Revisa el registro del servidor.",
                "danger",
            )


            return render_template(
                "expense_form.html",

                expense=None,

                categories=(
                    EXPENSE_CATEGORIES
                ),

                expense_types=(
                    EXPENSE_TYPES
                ),

                payment_methods=(
                    PAYMENT_METHODS
                ),

                today=date.today().isoformat(),
            )


        flash(
            "Gasto registrado correctamente.",
            "success",
        )


        return redirect(
            url_for(
                "expenses.list_expenses"
            )
        )


    # ========================================================
    # GET
    # ========================================================

    return render_template(
        "expense_form.html",

        expense=None,

        categories=(
            EXPENSE_CATEGORIES
        ),

        expense_types=(
            EXPENSE_TYPES
        ),

        payment_methods=(
            PAYMENT_METHODS
        ),

        today=date.today().isoformat(),
    )


# ============================================================
# EDITAR GASTO
# ============================================================

@expenses_bp.route(
    "/<int:expense_id>/editar",
    methods=[
        "GET",
        "POST",
    ],
)
@admin_required
def edit_expense(
    expense_id,
):

    # ========================================================
    # BUSCAR REGISTRO
    # ========================================================

    with db_conn() as conn:

        expense = conn.execute(
            """
            SELECT

                id,

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

            FROM expenses

            WHERE
                id = ?
            """,
            (
                expense_id,
            ),
        ).fetchone()


    if not expense:

        flash(
            "El gasto no existe.",
            "danger",
        )

        return redirect(
            url_for(
                "expenses.list_expenses"
            )
        )


    # ========================================================
    # ACTUALIZAR
    # ========================================================

    if request.method == "POST":


        expense_date = parse_date(
            request.form.get(
                "expense_date"
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
                "amount"
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


        # ====================================================
        # VALIDACIONES
        # ====================================================

        errors = []


        if not expense_date:

            errors.append(
                "Debes seleccionar la fecha del gasto."
            )


        if not category:

            errors.append(
                "Debes seleccionar una categoría."
            )


        elif category not in EXPENSE_CATEGORIES:

            errors.append(
                "La categoría seleccionada no es válida."
            )


        if not description:

            errors.append(
                "Debes ingresar una descripción."
            )


        if (
            amount is None
            or
            amount <= 0
        ):

            errors.append(
                "El valor debe ser mayor que cero."
            )


        if expense_type not in EXPENSE_TYPES:

            errors.append(
                "Debes seleccionar un tipo de gasto válido."
            )


        if (
            payment_method
            and
            payment_method not in PAYMENT_METHODS
        ):

            errors.append(
                "El método de pago seleccionado no es válido."
            )


        if errors:

            for error in errors:

                flash(
                    error,
                    "danger",
                )


            with db_conn() as conn:

                expense = conn.execute(
                    """
                    SELECT *
                    FROM expenses
                    WHERE id = ?
                    """,
                    (
                        expense_id,
                    ),
                ).fetchone()


            return render_template(
                "expense_form.html",

                expense=expense,

                categories=(
                    EXPENSE_CATEGORIES
                ),

                expense_types=(
                    EXPENSE_TYPES
                ),

                payment_methods=(
                    PAYMENT_METHODS
                ),

                today=date.today().isoformat(),
            )


        # ====================================================
        # UPDATE
        # ====================================================

        try:

            with db_conn() as conn:

                conn.execute(
                    """
                    UPDATE expenses

                    SET

                        expense_date = ?,

                        category = ?,

                        description = ?,

                        amount = ?,

                        payment_method = ?,

                        expense_type = ?,

                        notes = ?,

                        updated_at = ?

                    WHERE
                        id = ?
                    """,
                    (
                        expense_date.isoformat(),

                        category,

                        description,

                        amount,

                        payment_method
                        or
                        None,

                        expense_type,

                        notes
                        or
                        None,

                        datetime.now(),

                        expense_id,
                    ),
                )


                # ============================================
                # CONFIRMAR CAMBIO EN POSTGRESQL
                # ============================================

                commit_connection(
                    conn
                )


        except Exception as exc:

            print(
                "ERROR EDITANDO GASTO:",
                repr(exc),
            )


            flash(
                "No fue posible actualizar el gasto.",
                "danger",
            )


            return redirect(
                url_for(
                    "expenses.edit_expense",
                    expense_id=expense_id,
                )
            )


        flash(
            "Gasto actualizado correctamente.",
            "success",
        )


        return redirect(
            url_for(
                "expenses.list_expenses"
            )
        )


    # ========================================================
    # GET
    # ========================================================

    return render_template(
        "expense_form.html",

        expense=expense,

        categories=(
            EXPENSE_CATEGORIES
        ),

        expense_types=(
            EXPENSE_TYPES
        ),

        payment_methods=(
            PAYMENT_METHODS
        ),

        today=date.today().isoformat(),
    )


# ============================================================
# ELIMINAR GASTO
# ============================================================

@expenses_bp.route(
    "/<int:expense_id>/eliminar",
    methods=[
        "POST",
    ],
)
@admin_required
def delete_expense(
    expense_id,
):

    try:

        with db_conn() as conn:


            expense = conn.execute(
                """
                SELECT
                    id,
                    description

                FROM expenses

                WHERE
                    id = ?
                """,
                (
                    expense_id,
                ),
            ).fetchone()


            if not expense:

                flash(
                    "El gasto no existe.",
                    "warning",
                )

                return redirect(
                    url_for(
                        "expenses.list_expenses"
                    )
                )


            conn.execute(
                """
                DELETE FROM expenses

                WHERE
                    id = ?
                """,
                (
                    expense_id,
                ),
            )


            # ================================================
            # CONFIRMAR ELIMINACIÓN EN POSTGRESQL
            # ================================================

            commit_connection(
                conn
            )


    except Exception as exc:

        print(
            "ERROR ELIMINANDO GASTO:",
            repr(exc),
        )


        flash(
            "No fue posible eliminar el gasto.",
            "danger",
        )


        return redirect(
            url_for(
                "expenses.list_expenses"
            )
        )


    flash(
        "Gasto eliminado correctamente.",
        "success",
    )


    return redirect(
        url_for(
            "expenses.list_expenses"
        )
    )