from __future__ import annotations

from functools import wraps
from datetime import datetime
from decimal import Decimal

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


extra_income_bp = Blueprint(
    "extra_income",
    __name__,
    url_prefix="/ingresos-extra",
)


# ============================================================
# MODELO
# ============================================================

class ExtraIncome(db.Model):

    __tablename__ = "extra_income"

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    income_date = db.Column(
        db.Date,
        nullable=False
    )

    category = db.Column(
        db.String(100),
        nullable=False
    )

    concept = db.Column(
        db.String(200),
        nullable=False
    )

    amount = db.Column(
        db.Numeric(14, 2),
        nullable=False
    )

    payment_method = db.Column(
        db.String(50),
        nullable=True
    )

    notes = db.Column(
        db.Text,
        nullable=True
    )

    user_id = db.Column(
        db.Integer,
        nullable=True
    )

    created_at = db.Column(
        db.DateTime,
        nullable=False,
        default=datetime.utcnow
    )


# ============================================================
# SEGURIDAD
# ============================================================

def admin_required(view):

    @wraps(view)
    def wrapped_view(*args, **kwargs):

        if g.user is None:

            flash(
                "Debes iniciar sesión.",
                "warning"
            )

            return redirect(
                url_for("login")
            )

        if g.user["role"] != "administrador":

            flash(
                "No tienes permisos para acceder a este módulo.",
                "danger"
            )

            return redirect(
                url_for("dashboard")
            )

        return view(*args, **kwargs)

    return wrapped_view


# ============================================================
# LISTAR / CREAR
# ============================================================

@extra_income_bp.route(
    "/",
    methods=["GET", "POST"]
)
@admin_required
def index():

    if request.method == "POST":

        try:

            income_date_raw = (
                request.form.get(
                    "income_date",
                    ""
                )
                .strip()
            )

            category = (
                request.form.get(
                    "category",
                    ""
                )
                .strip()
            )

            concept = (
                request.form.get(
                    "concept",
                    ""
                )
                .strip()
            )

            amount_raw = (
                request.form.get(
                    "amount",
                    "0"
                )
                .strip()
                .replace("$", "")
                .replace(".", "")
                .replace(",", ".")
            )

            payment_method = (
                request.form.get(
                    "payment_method",
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

            if not income_date_raw:

                raise ValueError(
                    "La fecha es obligatoria."
                )

            income_date = datetime.strptime(
                income_date_raw,
                "%Y-%m-%d"
            ).date()

            if not category:

                raise ValueError(
                    "La categoría es obligatoria."
                )

            if not concept:

                raise ValueError(
                    "El concepto es obligatorio."
                )

            amount = Decimal(
                amount_raw
            )

            if amount <= 0:

                raise ValueError(
                    "El valor debe ser mayor a cero."
                )


            with db_conn() as conn:

                conn.execute(
                    """
                    INSERT INTO extra_income(
                        income_date,
                        category,
                        concept,
                        amount,
                        payment_method,
                        notes,
                        user_id,
                        created_at
                    )
                    VALUES(?,?,?,?,?,?,?,?)
                    """,
                    (
                        income_date,
                        category,
                        concept,
                        float(amount),
                        payment_method,
                        notes,
                        g.user["id"],
                        datetime.now().isoformat(
                            timespec="seconds"
                        ),
                    )
                )


            flash(
                "Ingreso extra registrado correctamente.",
                "success"
            )

            return redirect(
                url_for(
                    "extra_income.index"
                )
            )


        except (
            ValueError,
            TypeError
        ) as exc:

            flash(
                str(exc),
                "danger"
            )


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


    filters = []
    params = []


    if date_from:

        filters.append(
            "income_date >= ?"
        )

        params.append(
            date_from
        )


    if date_to:

        filters.append(
            "income_date <= ?"
        )

        params.append(
            date_to
        )


    where_clause = ""

    if filters:

        where_clause = (
            " WHERE "
            +
            " AND ".join(
                filters
            )
        )


    with db_conn() as conn:

        rows = conn.execute(
            f"""
            SELECT *
            FROM extra_income

            {where_clause}

            ORDER BY
                income_date DESC,
                id DESC
            """,
            tuple(params)
        ).fetchall()


        summary = conn.execute(
            f"""
            SELECT
                COALESCE(
                    SUM(amount),
                    0
                ) AS total

            FROM extra_income

            {where_clause}
            """,
            tuple(params)
        ).fetchone()


    return render_template(
        "extra_income.html",
        rows=rows,
        total=summary["total"],
        date_from=date_from,
        date_to=date_to,
        today=datetime.now().date().isoformat(),
    )


# ============================================================
# ELIMINAR
# ============================================================

@extra_income_bp.route(
    "/<int:income_id>/eliminar",
    methods=["POST"]
)
@admin_required
def delete(income_id):

    with db_conn() as conn:

        conn.execute(
            """
            DELETE FROM extra_income
            WHERE id=?
            """,
            (
                income_id,
            )
        )


    flash(
        "Ingreso extra eliminado.",
        "success"
    )

    return redirect(
        url_for(
            "extra_income.index"
        )
    )