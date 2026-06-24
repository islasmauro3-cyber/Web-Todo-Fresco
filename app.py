import os
import uuid
from datetime import datetime, date
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, send_from_directory, abort
)
from werkzeug.utils import secure_filename
from PIL import Image

from database import get_db, init_db

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "cambia-esta-clave-secreta-123")

UPLOAD_FOLDER = os.path.join(app.root_path, "static", "uploads")
ALLOWED_EXTENSIONS = {"jpg", "jpeg", "png", "webp", "gif"}
MAX_IMAGE_SIZE = (800, 800)

# Contraseña de acceso (se puede cambiar por variable de entorno)
APP_PASSWORD = os.environ.get("APP_PASSWORD", "admin123")


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def save_image(file):
    """Guarda y redimensiona la imagen subida. Devuelve el nombre del archivo."""
    ext = file.filename.rsplit(".", 1)[1].lower()
    filename = f"{uuid.uuid4().hex}.{ext}"
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    img = Image.open(file)
    img.thumbnail(MAX_IMAGE_SIZE, Image.LANCZOS)
    img.save(filepath, optimize=True, quality=85)
    return filename


def delete_image(filename):
    if filename:
        path = os.path.join(UPLOAD_FOLDER, filename)
        if os.path.exists(path):
            os.remove(path)


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("autenticado"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


def oferta_activa(producto):
    """Determina si un producto está en oferta activa hoy."""
    if not producto["en_oferta"]:
        return False
    hoy = date.today().isoformat()
    desde = producto["oferta_desde"] or ""
    hasta = producto["oferta_hasta"] or ""
    if desde and hoy < desde:
        return False
    if hasta and hoy > hasta:
        return False
    return True


# ── Autenticación ─────────────────────────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("autenticado"):
        return redirect(url_for("dashboard"))
    error = None
    if request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session["autenticado"] = True
            session.permanent = True
            return redirect(url_for("dashboard"))
        error = "Contraseña incorrecta"
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ── Dashboard ─────────────────────────────────────────────────────────────────

@app.route("/")
@login_required
def dashboard():
    db = get_db()
    total_productos = db.execute("SELECT COUNT(*) FROM productos WHERE activo=1").fetchone()[0]
    en_oferta = db.execute(
        "SELECT COUNT(*) FROM productos WHERE activo=1 AND en_oferta=1"
    ).fetchone()[0]
    total_clientes = db.execute("SELECT COUNT(*) FROM clientes WHERE activo=1").fetchone()[0]
    db.close()
    return render_template(
        "dashboard.html",
        total_productos=total_productos,
        en_oferta=en_oferta,
        total_clientes=total_clientes,
    )


# ── Productos ─────────────────────────────────────────────────────────────────

@app.route("/productos")
@login_required
def productos():
    db = get_db()
    rows = db.execute(
        "SELECT * FROM productos ORDER BY activo DESC, nombre ASC"
    ).fetchall()
    db.close()
    items = [dict(r) for r in rows]
    for p in items:
        p["oferta_hoy"] = oferta_activa(p)
    return render_template("productos.html", productos=items)


@app.route("/productos/nuevo", methods=["GET", "POST"])
@login_required
def producto_nuevo():
    if request.method == "POST":
        return _guardar_producto(None)
    return render_template("producto_form.html", producto=None)


@app.route("/productos/<int:pid>/editar", methods=["GET", "POST"])
@login_required
def producto_editar(pid):
    db = get_db()
    row = db.execute("SELECT * FROM productos WHERE id=?", (pid,)).fetchone()
    db.close()
    if not row:
        abort(404)
    if request.method == "POST":
        return _guardar_producto(dict(row))
    return render_template("producto_form.html", producto=dict(row))


def _guardar_producto(existente):
    pid = existente["id"] if existente else None
    nombre = request.form.get("nombre", "").strip()
    descripcion = request.form.get("descripcion", "").strip()
    precio_raw = request.form.get("precio", "").strip().replace(",", ".")
    en_oferta = 1 if request.form.get("en_oferta") else 0
    oferta_desde = request.form.get("oferta_desde") or None
    oferta_hasta = request.form.get("oferta_hasta") or None
    activo = 1 if request.form.get("activo") else 0

    if not nombre:
        flash("El nombre es obligatorio.", "danger")
        return render_template("producto_form.html", producto=existente)

    try:
        precio = float(precio_raw)
    except ValueError:
        flash("El precio debe ser un número.", "danger")
        return render_template("producto_form.html", producto=existente)

    foto_actual = existente["foto"] if existente else None
    foto = foto_actual

    file = request.files.get("foto")
    if file and file.filename:
        if not allowed_file(file.filename):
            flash("Formato de imagen no permitido. Usá JPG, PNG o WEBP.", "danger")
            return render_template("producto_form.html", producto=existente)
        delete_image(foto_actual)
        foto = save_image(file)

    db = get_db()
    if pid:
        db.execute(
            """UPDATE productos SET nombre=?, descripcion=?, precio=?, foto=?,
               en_oferta=?, oferta_desde=?, oferta_hasta=?, activo=?
               WHERE id=?""",
            (nombre, descripcion, precio, foto, en_oferta, oferta_desde, oferta_hasta, activo, pid),
        )
        flash("Producto actualizado.", "success")
    else:
        db.execute(
            """INSERT INTO productos (nombre, descripcion, precio, foto, en_oferta,
               oferta_desde, oferta_hasta, activo)
               VALUES (?,?,?,?,?,?,?,?)""",
            (nombre, descripcion, precio, foto, en_oferta, oferta_desde, oferta_hasta, activo),
        )
        flash("Producto creado.", "success")
    db.commit()
    db.close()
    return redirect(url_for("productos"))


@app.route("/productos/<int:pid>/borrar", methods=["POST"])
@login_required
def producto_borrar(pid):
    db = get_db()
    row = db.execute("SELECT foto FROM productos WHERE id=?", (pid,)).fetchone()
    if row:
        delete_image(row["foto"])
        db.execute("DELETE FROM productos WHERE id=?", (pid,))
        db.commit()
        flash("Producto eliminado.", "success")
    db.close()
    return redirect(url_for("productos"))


@app.route("/productos/<int:pid>/toggle-oferta", methods=["POST"])
@login_required
def toggle_oferta(pid):
    db = get_db()
    row = db.execute("SELECT en_oferta FROM productos WHERE id=?", (pid,)).fetchone()
    if row:
        nuevo = 0 if row["en_oferta"] else 1
        db.execute("UPDATE productos SET en_oferta=? WHERE id=?", (nuevo, pid))
        db.commit()
    db.close()
    return redirect(url_for("productos"))


# ── Clientes ──────────────────────────────────────────────────────────────────

@app.route("/clientes")
@login_required
def clientes():
    db = get_db()
    rows = db.execute(
        "SELECT * FROM clientes ORDER BY activo DESC, nombre ASC"
    ).fetchall()
    db.close()
    return render_template("clientes.html", clientes=[dict(r) for r in rows])


@app.route("/clientes/nuevo", methods=["GET", "POST"])
@login_required
def cliente_nuevo():
    if request.method == "POST":
        return _guardar_cliente(None)
    return render_template("cliente_form.html", cliente=None)


@app.route("/clientes/<int:cid>/editar", methods=["GET", "POST"])
@login_required
def cliente_editar(cid):
    db = get_db()
    row = db.execute("SELECT * FROM clientes WHERE id=?", (cid,)).fetchone()
    db.close()
    if not row:
        abort(404)
    if request.method == "POST":
        return _guardar_cliente(dict(row))
    return render_template("cliente_form.html", cliente=dict(row))


def _guardar_cliente(existente):
    cid = existente["id"] if existente else None
    nombre = request.form.get("nombre", "").strip()
    celular = request.form.get("celular", "").strip()
    notas = request.form.get("notas", "").strip()
    activo = 1 if request.form.get("activo") else 0

    if not nombre or not celular:
        flash("Nombre y celular son obligatorios.", "danger")
        return render_template("cliente_form.html", cliente=existente)

    db = get_db()
    if cid:
        db.execute(
            "UPDATE clientes SET nombre=?, celular=?, notas=?, activo=? WHERE id=?",
            (nombre, celular, notas, activo, cid),
        )
        flash("Cliente actualizado.", "success")
    else:
        db.execute(
            "INSERT INTO clientes (nombre, celular, notas, activo) VALUES (?,?,?,?)",
            (nombre, celular, notas, activo),
        )
        flash("Cliente agregado.", "success")
    db.commit()
    db.close()
    return redirect(url_for("clientes"))


@app.route("/clientes/<int:cid>/borrar", methods=["POST"])
@login_required
def cliente_borrar(cid):
    db = get_db()
    db.execute("DELETE FROM clientes WHERE id=?", (cid,))
    db.commit()
    db.close()
    flash("Cliente eliminado.", "success")
    return redirect(url_for("clientes"))


@app.route("/clientes/<int:cid>/toggle", methods=["POST"])
@login_required
def cliente_toggle(cid):
    db = get_db()
    row = db.execute("SELECT activo FROM clientes WHERE id=?", (cid,)).fetchone()
    if row:
        db.execute("UPDATE clientes SET activo=? WHERE id=?", (1 - row["activo"], cid))
        db.commit()
    db.close()
    return redirect(url_for("clientes"))


# ── Vista previa WhatsApp ──────────────────────────────────────────────────────

@app.route("/whatsapp")
@login_required
def whatsapp_preview():
    db = get_db()
    productos_oferta = db.execute(
        "SELECT * FROM productos WHERE activo=1 AND en_oferta=1 ORDER BY nombre"
    ).fetchall()
    clientes_activos = db.execute(
        "SELECT * FROM clientes WHERE activo=1 ORDER BY nombre"
    ).fetchall()
    db.close()

    hoy = date.today().isoformat()
    ofertas_hoy = []
    for p in productos_oferta:
        p = dict(p)
        if oferta_activa(p):
            ofertas_hoy.append(p)

    mensaje = _generar_mensaje(ofertas_hoy)
    return render_template(
        "whatsapp_preview.html",
        ofertas=ofertas_hoy,
        clientes=clientes_activos,
        mensaje=mensaje,
    )


def _generar_mensaje(ofertas):
    if not ofertas:
        return "No hay productos en oferta hoy."
    lineas = ["🛍️ *¡Ofertas del día!* 🛍️", ""]
    for p in ofertas:
        precio_fmt = f"${p['precio']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        lineas.append(f"✅ *{p['nombre']}*")
        if p["descripcion"]:
            lineas.append(f"   {p['descripcion']}")
        lineas.append(f"   💲 {precio_fmt}")
        lineas.append("")
    lineas.append("¡Contactame para hacer tu pedido! 😊")
    return "\n".join(lineas)


# ── Archivos estáticos subidos ─────────────────────────────────────────────────

@app.route("/uploads/<filename>")
def uploaded_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename)


if __name__ == "__main__":
    init_db()
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    app.run(debug=True, host="0.0.0.0", port=5000)
