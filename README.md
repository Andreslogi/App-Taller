# Inventario, Remisiones y Servicios — versión Web/PostgreSQL

Esta versión conserva el funcionamiento del piloto local y agrega soporte para PostgreSQL y despliegue en Render.

## Qué conserva

- Inventario inicial existente.
- Productos nuevos desde la app (solo administradores).
- Entradas y salidas de inventario (solo administradores).
- Remisiones con placa.
- Usuarios y roles.
- Usuarios iniciales:
  - `administrador` / `admin123*`
  - `vendedor` / `vendedor123*`
  - `arturo.lopez` / `Arturo1963*`
- Vendedores iniciales James y Leonardo.
- Creación de vendedores desde la app.
- Servicios con reparto 70 % trabajador / 30 % negocio.
- Servicios cargados desde `data/Servicios_Iniciales.xlsx` de forma aditiva.
- PDF de remisiones.

> Cambia las contraseñas iniciales y configura una `SECRET_KEY` segura antes de producción.

---

## 1. Desarrollo local con SQLite

Si `DATABASE_URL` no existe, la aplicación usa SQLite automáticamente.

```bat
py -3.11 -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt --timeout 120 --retries 10
python run.py
```

La app abrirá en `http://127.0.0.1:5000`.

---

## 2. PostgreSQL

La aplicación usa PostgreSQL cuando existe la variable de entorno `DATABASE_URL`.

Ejemplo de variable (no subir contraseñas reales a GitHub):

```text
DATABASE_URL=postgresql://usuario:password@host:5432/inventario
```

La aplicación convierte internamente esta URL al driver `psycopg`.

---

## 3. Migrar el SQLite actual a PostgreSQL

Antes de migrar, haz una copia de seguridad de `inventario.db`.

En PowerShell:

```powershell
$env:DATABASE_URL="postgresql://USUARIO:CLAVE@HOST/BASE"
python migrate_sqlite_to_postgres.py --sqlite "C:\ruta\inventario.db"
```

El script migra, en orden:

- users
- settings
- products
- customers
- salespeople
- services
- invoices
- invoice_items
- invoice_service_items
- movements

Por seguridad, si PostgreSQL ya contiene información el script se detiene. Solo usa `--replace` si realmente quieres vaciar PostgreSQL y reemplazarlo por el contenido de SQLite.

```powershell
python migrate_sqlite_to_postgres.py --sqlite "C:\ruta\inventario.db" --replace
```

---

## 4. Subir a GitHub

Crea un repositorio vacío en GitHub y, desde esta carpeta:

```bat
git init
git add .
git commit -m "Version web con PostgreSQL"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/TU_REPOSITORIO.git
git push -u origin main
```

`.gitignore` excluye `.env`, la base SQLite, PDFs, entornos virtuales y builds.

---

## 5. Desplegar en Render

El proyecto incluye `render.yaml`.

### Opción sencilla — Blueprint

1. Sube el proyecto a GitHub.
2. En Render crea un **Blueprint** y selecciona el repositorio.
3. Render leerá `render.yaml`.
4. Se crearán el servicio web y PostgreSQL.
5. `DATABASE_URL` se enlaza automáticamente a PostgreSQL.
6. `SECRET_KEY` se genera automáticamente.

Comando de build:

```text
pip install -r requirements.txt
```

Comando de inicio:

```text
gunicorn app:app --workers 2 --threads 4 --timeout 120
```

Al arrancar por primera vez, si la base está vacía, se crean las tablas, usuarios predeterminados, vendedores iniciales, inventario inicial y servicios iniciales.

### Si vas a migrar la base local real

Es preferible crear PostgreSQL vacío, migrar `inventario.db` usando `migrate_sqlite_to_postgres.py` y después desplegar la app. El inicializador es aditivo y no borra registros existentes.

---

## 6. Concurrencia de inventario

Cuando la aplicación trabaja con PostgreSQL, al crear una remisión:

- bloquea temporalmente el consecutivo de remisión;
- bloquea los productos vendidos durante la transacción;
- valida el stock antes de descontarlo;
- confirma remisión, movimientos y stock en una sola transacción.

Esto reduce el riesgo de que dos vendedores vendan simultáneamente la última unidad.

---

## 7. Variables de entorno

No guardes secretos en GitHub. Variables usadas:

```text
SECRET_KEY=...
DATABASE_URL=...
PORT=...
```

`PORT` normalmente lo define el proveedor automáticamente.

---

## 8. Futuras modificaciones de estructura

El proyecto incluye Flask-Migrate/Alembic. Para cambios futuros en modelos puedes usar:

```bat
set FLASK_APP=app.py
flask db init
flask db migrate -m "descripcion del cambio"
flask db upgrade
```

`flask db init` se ejecuta una sola vez por repositorio. En producción conserva la carpeta `migrations/` en GitHub.

---

## 9. Ejecutable local

La versión continúa siendo compatible con el piloto local. Para construir:

```bat
build_exe.bat
```

El `.exe` usa SQLite local si no se configura `DATABASE_URL`.
