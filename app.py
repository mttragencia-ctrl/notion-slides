"""App web: convierte un calendario de Notion en una presentación .pptx.

Rutas:
    GET  /          -> formulario (pide login si no hay sesión)
    POST /login     -> valida la contraseña compartida
    POST /logout    -> cierra sesión
    POST /generate  -> lee Notion y devuelve el .pptx para descargar

Variables de entorno necesarias:
    NOTION_TOKEN   token de la integración interna de Notion (secreto)
    APP_PASSWORD   contraseña compartida para entrar a la app
    SECRET_KEY     (opcional) clave para firmar la cookie de sesión
"""

import os
import re
import unicodedata
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, send_file, flash,
)

from lib.parse_url import extract_id
from lib.notion_fetch import fetch_items
from lib.slides import build_deck

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "cambia-esta-clave-en-produccion")

NOTION_TOKEN = os.environ.get("NOTION_TOKEN", "")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("auth"):
            return redirect(url_for("index"))
        return view(*args, **kwargs)
    return wrapped


def _safe_filename(title):
    """Convierte el título de la base en un nombre de archivo seguro."""
    norm = unicodedata.normalize("NFKD", title).encode("ascii", "ignore").decode()
    norm = re.sub(r"[^A-Za-z0-9]+", "_", norm).strip("_")
    return (norm or "Planificacion") + ".pptx"


@app.route("/")
def index():
    return render_template("index.html", authed=bool(session.get("auth")))


@app.route("/login", methods=["POST"])
def login():
    pwd = request.form.get("password", "")
    if APP_PASSWORD and pwd == APP_PASSWORD:
        session["auth"] = True
    else:
        flash("Contraseña incorrecta.", "error")
    return redirect(url_for("index"))


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("index"))


@app.route("/generate", methods=["POST"])
@login_required
def generate():
    notion_url = request.form.get("notion_url", "").strip()

    if not NOTION_TOKEN:
        flash("Falta configurar NOTION_TOKEN en el servidor.", "error")
        return redirect(url_for("index"))

    try:
        database_id = extract_id(notion_url)
    except ValueError as e:
        flash(str(e), "error")
        return redirect(url_for("index"))

    try:
        deck_title, items = fetch_items(NOTION_TOKEN, database_id)
    except Exception as e:
        msg = str(e)
        if "Could not find database" in msg or "object_not_found" in msg:
            flash(
                "No se pudo acceder a esa base. Verificá que la integración de "
                "Notion esté compartida con la base (… → Conexiones).",
                "error",
            )
        elif "unauthorized" in msg.lower() or "API token is invalid" in msg:
            flash("El token de Notion es inválido. Revisá NOTION_TOKEN.", "error")
        else:
            flash(f"Error leyendo Notion: {msg}", "error")
        return redirect(url_for("index"))

    if not items:
        flash("La base no tiene filas para convertir.", "error")
        return redirect(url_for("index"))

    buf = build_deck(items, deck_title)
    return send_file(
        buf,
        as_attachment=True,
        download_name=_safe_filename(deck_title),
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
