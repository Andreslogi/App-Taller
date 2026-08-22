@echo off
setlocal
cd /d "%~dp0"

echo ===============================================
echo  GENERANDO EJECUTABLE - INVENTARIO Y REMISIONES
echo ===============================================

if not exist ".venv\Scripts\python.exe" (
    echo Creando entorno virtual Python 3.11...
    py -3.11 -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt --timeout 120 --retries 10

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist InventarioFacturacion.spec del /q InventarioFacturacion.spec

pyinstaller --noconfirm --clean --onefile --name InventarioFacturacion ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  --add-data "data\Inventario_Inicial.xlsx;data" ^
  --add-data "data\Servicios_Iniciales.xlsx;data" ^
  --collect-all reportlab ^
  --collect-all psycopg ^
  run.py

if exist "dist\InventarioFacturacion.exe" (
    echo.
    echo EJECUTABLE CREADO CORRECTAMENTE:
    echo %CD%\dist\InventarioFacturacion.exe
    echo.
    start "" "%CD%\dist"
) else (
    echo No fue posible crear el ejecutable.
)
pause
