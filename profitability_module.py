from __future__ import annotations

from datetime import date, datetime, timedelta
from decimal import Decimal
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

from database import db_conn


# ============================================================
# BLUEPRINT
# ============================================================

profitability_bp = Blueprint(
    "profitability",
    __name__,
    url_prefix="/rentabilidad"
)


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

        role = str(
            g.user["role"]
            or
            ""
        ).strip().lower()

        if role != "administrador":

            flash(
                "Solo el administrador puede consultar la rentabilidad.",
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

def to_decimal(value):

    try:

        return Decimal(
            str(
                value
                or
                0
            )
        )

    except Exception:

        return Decimal("0")


def parse_date(value):

    return datetime.strptime(
        str(value)[:10],
        "%Y-%m-%d"
    ).date()


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
# PERÍODO
# ============================================================

def get_period():

    today = date.today()

    preset = (
        request.args.get(
            "preset",
            ""
        )
        .strip()
        .lower()
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


    # ========================================================
    # HOY
    # ========================================================

    if preset == "today":

        start = today
        end = today


    # ========================================================
    # ESTA SEMANA
    # Lunes a domingo
    # ========================================================

    elif preset == "week":

        start = (
            today
            -
            timedelta(
                days=today.weekday()
            )
        )

        end = (
            start
            +
            timedelta(
                days=6
            )
        )


    # ========================================================
    # MES ACTUAL
    # ========================================================

    elif preset == "month":

        start = first_day_month(
            today
        )

        end = last_day_month(
            today
        )


    # ========================================================
    # PERSONALIZADO
    # ========================================================

    elif date_from_raw and date_to_raw:

        try:

            start = parse_date(
                date_from_raw
            )

            end = parse_date(
                date_to_raw
            )

            preset = "custom"

        except ValueError:

            preset = "week"

            start = (
                today
                -
                timedelta(
                    days=today.weekday()
                )
            )

            end = (
                start
                +
                timedelta(
                    days=6
                )
            )


    # ========================================================
    # POR DEFECTO:
    # ESTA SEMANA
    # ========================================================

    else:

        preset = "week"

        start = (
            today
            -
            timedelta(
                days=today.weekday()
            )
        )

        end = (
            start
            +
            timedelta(
                days=6
            )
        )


    if end < start:

        start, end = end, start


    return (
        start,
        end,
        preset
    )


# ============================================================
# FUNCIÓN PRINCIPAL
# ============================================================

@profitability_bp.route("/")
@admin_required
def profitability():

    (
        date_from,
        date_to,
        preset
    ) = get_period()


    start_value = (
        date_from.isoformat()
    )

    end_value = (
        date_to.isoformat()
    )


    # ========================================================
    # ACUMULADORES
    # ========================================================

    product_income = Decimal("0")

    services_total = Decimal("0")

    service_business_income = Decimal("0")

    expenses_total = Decimal("0")

    fixed_expenses = Decimal("0")

    variable_expenses = Decimal("0")


    cash_invoices = 0

    paid_credit_invoices = 0

    excluded_credit_invoices = 0


    category_expenses = []


    with db_conn() as conn:


        # ====================================================
        # 1. REMISIONES NO CRÉDITO
        #
        # Efectivo, transferencia, tarjeta, etc.
        #
        # Se consideran cobradas en la fecha de la remisión.
        # ====================================================

        cash_rows = conn.execute(
            """
            SELECT
                i.id,
                i.issued_at,

                COALESCE(
                    (
                        SELECT SUM(ii.line_total)

                        FROM invoice_items ii

                        WHERE
                            ii.invoice_id = i.id
                    ),
                    0
                ) AS product_total,

                COALESCE(
                    (
                        SELECT SUM(si.line_total)

                        FROM invoice_service_items si

                        WHERE
                            si.invoice_id = i.id
                    ),
                    0
                ) AS service_total

            FROM invoices i

            WHERE

                COALESCE(
                    i.payment_method,
                    ''
                ) <> 'Crédito'

                AND DATE(i.issued_at) >= ?

                AND DATE(i.issued_at) <= ?
            """,
            (
                start_value,
                end_value
            )
        ).fetchall()


        for row in cash_rows:

            product_total = to_decimal(
                row["product_total"]
            )

            service_total = to_decimal(
                row["service_total"]
            )


            product_income += (
                product_total
            )


            services_total += (
                service_total
            )


            # 30 % del valor del servicio
            # corresponde al negocio.

            service_business_income += (
                service_total
                *
                Decimal("0.30")
            )


            cash_invoices += 1


        # ====================================================
        # 2. REMISIONES A CRÉDITO
        #
        # Solo cuentan cuando están totalmente pagadas.
        #
        # Se asignan al período del último pago,
        # porque ahí fue cuando realmente quedó cobrada.
        # ====================================================

        credit_rows = conn.execute(
            """
            SELECT
                i.id,
                i.issued_at,

                ar.id AS receivable_id,

                ar.original_amount,

                COALESCE(
                    (
                        SELECT SUM(arp.amount)

                        FROM accounts_receivable_payments arp

                        WHERE
                            arp.receivable_id = ar.id
                    ),
                    0
                ) AS paid_amount,

                (
                    SELECT MAX(arp.payment_date)

                    FROM accounts_receivable_payments arp

                    WHERE
                        arp.receivable_id = ar.id
                ) AS last_payment_date,

                COALESCE(
                    (
                        SELECT SUM(ii.line_total)

                        FROM invoice_items ii

                        WHERE
                            ii.invoice_id = i.id
                    ),
                    0
                ) AS product_total,

                COALESCE(
                    (
                        SELECT SUM(si.line_total)

                        FROM invoice_service_items si

                        WHERE
                            si.invoice_id = i.id
                    ),
                    0
                ) AS service_total

            FROM invoices i

            LEFT JOIN accounts_receivable ar
                ON ar.invoice_id = i.id

            WHERE
                i.payment_method = 'Crédito'
            """
        ).fetchall()


        for row in credit_rows:

            original_amount = to_decimal(
                row["original_amount"]
            )

            paid_amount = to_decimal(
                row["paid_amount"]
            )


            # =================================================
            # No tiene cuenta por cobrar válida
            # =================================================

            if original_amount <= 0:

                excluded_credit_invoices += 1

                continue


            # =================================================
            # Todavía tiene saldo
            # =================================================

            if paid_amount < original_amount:

                excluded_credit_invoices += 1

                continue


            # =================================================
            # Ya está totalmente pagada
            # =================================================

            last_payment_date = (
                row["last_payment_date"]
            )


            if not last_payment_date:

                excluded_credit_invoices += 1

                continue


            try:

                payment_date = parse_date(
                    last_payment_date
                )

            except ValueError:

                excluded_credit_invoices += 1

                continue


            # Solo entra si terminó de cobrarse
            # dentro del período seleccionado.

            if not (
                date_from
                <=
                payment_date
                <=
                date_to
            ):

                continue


            product_total = to_decimal(
                row["product_total"]
            )

            service_total = to_decimal(
                row["service_total"]
            )


            product_income += (
                product_total
            )


            services_total += (
                service_total
            )


            service_business_income += (
                service_total
                *
                Decimal("0.30")
            )


            paid_credit_invoices += 1


        # ====================================================
        # 3. GASTOS OPERATIVOS
        #
        # IMPORTANTE:
        #
        # Solo se consulta la tabla expenses.
        #
        # Las facturas de proveedores / inventario
        # NO se incluyen aquí.
        # ====================================================

        expense_totals = conn.execute(
            """
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
                expense_date >= ?

                AND expense_date <= ?
            """,
            (
                start_value,
                end_value
            )
        ).fetchone()


        if expense_totals:

            expenses_total = to_decimal(
                expense_totals["total"]
            )

            fixed_expenses = to_decimal(
                expense_totals[
                    "fixed_total"
                ]
            )

            variable_expenses = to_decimal(
                expense_totals[
                    "variable_total"
                ]
            )


        # ====================================================
        # 4. GASTOS POR CATEGORÍA
        # ====================================================

        category_expenses = conn.execute(
            """
            SELECT

                category,

                COUNT(*) AS quantity,

                COALESCE(
                    SUM(amount),
                    0
                ) AS total

            FROM expenses

            WHERE
                expense_date >= ?

                AND expense_date <= ?

            GROUP BY category

            ORDER BY total DESC
            """,
            (
                start_value,
                end_value
            )
        ).fetchall()


    # ========================================================
    # INGRESO TOTAL ATRIBUIBLE AL NEGOCIO
    #
    # Productos:
    #   entra el 100 %
    #
    # Servicios:
    #   entra solo el 30 %
    # ========================================================

    total_business_income = (
        product_income
        +
        service_business_income
    )


    # ========================================================
    # RESULTADO OPERATIVO
    # ========================================================

    operating_result = (
        total_business_income
        -
        expenses_total
    )


    # ========================================================
    # MARGEN
    #
    # Resultado / Ingresos x 100
    # ========================================================

    if total_business_income > 0:

        operating_margin = (
            operating_result
            /
            total_business_income
            *
            Decimal("100")
        )

    else:

        operating_margin = (
            Decimal("0")
        )


    # ========================================================
    # SEMÁFORO DE SOSTENIBILIDAD
    #
    # ROJO:
    #   hay pérdida
    #
    # NARANJA:
    #   margen positivo pero menor a 10 %
    #
    # VERDE:
    #   margen >= 10 %
    # ========================================================

    if operating_result < 0:

        health_status = (
            "PERDIDA"
        )

        health_class = (
            "red"
        )

        health_label = (
            "Pérdida operativa"
        )


    elif (
        total_business_income > 0
        and
        operating_margin < Decimal("10")
    ):

        health_status = (
            "LOW"
        )

        health_class = (
            "orange"
        )

        health_label = (
            "Margen bajo"
        )


    elif total_business_income <= 0:

        health_status = (
            "NO_DATA"
        )

        health_class = (
            "orange"
        )

        health_label = (
            "Sin ingresos registrados"
        )


    else:

        health_status = (
            "HEALTHY"
        )

        health_class = (
            "green"
        )

        health_label = (
            "Resultado positivo"
        )


    # ========================================================
    # DISTRIBUCIÓN DE INGRESOS
    # ========================================================

    if total_business_income > 0:

        product_percentage = (
            product_income
            /
            total_business_income
            *
            Decimal("100")
        )

        service_percentage = (
            service_business_income
            /
            total_business_income
            *
            Decimal("100")
        )

    else:

        product_percentage = (
            Decimal("0")
        )

        service_percentage = (
            Decimal("0")
        )


    # ========================================================
    # DISTRIBUCIÓN GASTOS
    # ========================================================

    if expenses_total > 0:

        fixed_percentage = (
            fixed_expenses
            /
            expenses_total
            *
            Decimal("100")
        )

        variable_percentage = (
            variable_expenses
            /
            expenses_total
            *
            Decimal("100")
        )

    else:

        fixed_percentage = (
            Decimal("0")
        )

        variable_percentage = (
            Decimal("0")
        )


    # ========================================================
    # RETORNO
    # ========================================================

    return render_template(
        "profitability.html",

        # Período
        date_from=start_value,
        date_to=end_value,
        preset=preset,

        # Ingresos
        product_income=(
            product_income
        ),

        services_total=(
            services_total
        ),

        service_business_income=(
            service_business_income
        ),

        total_business_income=(
            total_business_income
        ),

        # Gastos
        expenses_total=(
            expenses_total
        ),

        fixed_expenses=(
            fixed_expenses
        ),

        variable_expenses=(
            variable_expenses
        ),

        # Resultado
        operating_result=(
            operating_result
        ),

        operating_margin=(
            operating_margin
        ),

        # Distribuciones
        product_percentage=(
            product_percentage
        ),

        service_percentage=(
            service_percentage
        ),

        fixed_percentage=(
            fixed_percentage
        ),

        variable_percentage=(
            variable_percentage
        ),

        # Semáforo
        health_status=(
            health_status
        ),

        health_class=(
            health_class
        ),

        health_label=(
            health_label
        ),

        # Control cartera
        cash_invoices=(
            cash_invoices
        ),

        paid_credit_invoices=(
            paid_credit_invoices
        ),

        excluded_credit_invoices=(
            excluded_credit_invoices
        ),

        # Categorías
        category_expenses=(
            category_expenses
        ),
    )