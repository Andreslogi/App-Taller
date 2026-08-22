import os
import threading
import time
import webbrowser
from waitress import serve
from app import app


def open_browser():
    time.sleep(1.2)
    webbrowser.open("http://127.0.0.1:5000")


if __name__ == "__main__":
    # El esquema y los datos iniciales se inicializan al importar app.py.
    threading.Thread(target=open_browser, daemon=True).start()
    serve(app, host="127.0.0.1", port=int(os.getenv("PORT", "5000")), threads=6)
