import os
import csv
import io
import uuid
from datetime import datetime, date
from functools import wraps

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, send_from_directory, abort, Response
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

UNIDADES = [
    "por kg",
    "por unidad",
    "por 100g",
    "por 1/4 kg",
    "por media horma",
    "por horma entera",
]

_UNIDAD_MAP = {
    # kg
    "kg": "por kg", "kilo": "por kg", "kilos": "por kg", "kilogramo": "por kg",
    "por kg": "por kg",
    # 100g
    "100g": "por 100g", "100 g": "por 100g", "por 100g": "por 100g",
    # 1/4 kg
    "1/4": "por 1/4 kg", "cuarto": "por 1/4 kg", "1/4 kg": "por 1/4 kg",
    "por 1/4 kg": "por 1/4 kg", "cuarto kg": "por 1/4 kg",
    # media horma
    "media horma": "por media horma", "1/2 horma": "por media horma",
    "por media horma": "por media horma", "media": "por media horma",
    # horma entera
    "horma": "por horma entera", "horma entera": "por horma entera",
    "por horma entera": "por horma entera",
    # unidad (default)
    "unidad": "por unidad", "u": "por unidad", "ud": "por unidad",
    "por unidad": "por unidad",
}


@app.template_filter("unidad_corta")
def unidad_corta(u):
    """'por kg' → 'kg', para mostrar '$3500 / kg'."""
    return (u or "unidad").replace("por ", "", 1)


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
    return render_template("producto_form.html", producto=None, unidades=UNIDADES)


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
    return render_template("producto_form.html", producto=dict(row), unidades=UNIDADES)


def _guardar_producto(existente):
    pid = existente["id"] if existente else None
    nombre = request.form.get("nombre", "").strip()
    descripcion = request.form.get("descripcion", "").strip()
    precio_raw = request.form.get("precio", "").strip().replace(",", ".")
    unidad = request.form.get("unidad", "por unidad")
    if unidad not in UNIDADES:
        unidad = "por unidad"
    en_oferta = 1 if request.form.get("en_oferta") else 0
    oferta_desde = request.form.get("oferta_desde") or None
    oferta_hasta = request.form.get("oferta_hasta") or None
    activo = 1 if request.form.get("activo") else 0

    if not nombre:
        flash("El nombre es obligatorio.", "danger")
        return render_template("producto_form.html", producto=existente, unidades=UNIDADES)

    try:
        precio = float(precio_raw)
    except ValueError:
        flash("El precio debe ser un número.", "danger")
        return render_template("producto_form.html", producto=existente, unidades=UNIDADES)

    foto_actual = existente["foto"] if existente else None
    foto = foto_actual

    file = request.files.get("foto")
    if file and file.filename:
        if not allowed_file(file.filename):
            flash("Formato de imagen no permitido. Usá JPG, PNG o WEBP.", "danger")
            return render_template("producto_form.html", producto=existente, unidades=UNIDADES)
        delete_image(foto_actual)
        foto = save_image(file)

    db = get_db()
    if pid:
        db.execute(
            """UPDATE productos SET nombre=?, descripcion=?, precio=?, foto=?,
               unidad=?, en_oferta=?, oferta_desde=?, oferta_hasta=?, activo=?
               WHERE id=?""",
            (nombre, descripcion, precio, foto, unidad,
             en_oferta, oferta_desde, oferta_hasta, activo, pid),
        )
        flash("Producto actualizado.", "success")
    else:
        db.execute(
            """INSERT INTO productos (nombre, descripcion, precio, foto, unidad,
               en_oferta, oferta_desde, oferta_hasta, activo)
               VALUES (?,?,?,?,?,?,?,?,?)""",
            (nombre, descripcion, precio, foto, unidad,
             en_oferta, oferta_desde, oferta_hasta, activo),
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
        u = (p.get("unidad") or "por unidad").replace("por ", "", 1)
        lineas.append(f"✅ *{p['nombre']}*")
        if p["descripcion"]:
            lineas.append(f"   {p['descripcion']}")
        lineas.append(f"   💲 {precio_fmt} / {u}")
        lineas.append("")
    lineas.append("¡Contactame para hacer tu pedido! 😊")
    return "\n".join(lineas)


# ── Importación masiva ────────────────────────────────────────────────────────

VERDADERO = {"si", "sí", "yes", "1", "true", "verdadero", "v", "s", "y"}


def _parsear_en_oferta(val):
    return 1 if str(val).strip().lower() in VERDADERO else 0


def _parsear_unidad(val):
    return _UNIDAD_MAP.get(str(val or "").strip().lower(), "por unidad")


def _parsear_precio(val):
    try:
        return float(str(val).strip().replace(",", ".").replace("$", "").replace(" ", ""))
    except (ValueError, TypeError):
        return None


def _filas_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    return rows


def _filas_csv(file_bytes):
    text = file_bytes.decode("utf-8-sig")
    reader = csv.reader(io.StringIO(text))
    return list(reader)


def _procesar_filas(rows):
    """Devuelve (productos_validos, errores) a partir de las filas crudas (incluyendo cabecera)."""
    if not rows:
        return [], ["El archivo está vacío."]

    # Normalizar cabecera
    header = [str(c).strip().lower() if c is not None else "" for c in rows[0]]

    col_map = {}
    for i, h in enumerate(header):
        if "nombre" in h:
            col_map["nombre"] = i
        elif "descrip" in h:
            col_map["descripcion"] = i
        elif "precio" in h:
            col_map["precio"] = i
        elif "oferta" in h:
            col_map["en_oferta"] = i
        elif "unidad" in h:
            col_map["unidad"] = i

    faltantes = [c for c in ("nombre", "precio") if c not in col_map]
    if faltantes:
        return [], [f"Columnas obligatorias no encontradas: {', '.join(faltantes)}. "
                    f"Verificá que la primera fila sea el encabezado."]

    productos, errores = [], []
    for n, row in enumerate(rows[1:], start=2):
        if all(c is None or str(c).strip() == "" for c in row):
            continue  # fila vacía

        nombre = str(row[col_map["nombre"]] or "").strip()
        descripcion = str(row[col_map.get("descripcion", -1)] or "").strip() if "descripcion" in col_map else ""
        precio_raw = row[col_map["precio"]]
        en_oferta_raw = row[col_map.get("en_oferta", -1)] if "en_oferta" in col_map else ""
        unidad_raw = row[col_map.get("unidad", -1)] if "unidad" in col_map else ""

        if not nombre:
            errores.append(f"Fila {n}: nombre vacío — se omitió.")
            continue

        precio = _parsear_precio(precio_raw)
        if precio is None:
            errores.append(f"Fila {n} ({nombre!r}): precio inválido ({precio_raw!r}) — se omitió.")
            continue

        productos.append({
            "nombre": nombre,
            "descripcion": descripcion,
            "precio": precio,
            "en_oferta": _parsear_en_oferta(en_oferta_raw),
            "unidad": _parsear_unidad(unidad_raw),
        })

    return productos, errores


@app.route("/productos/importar", methods=["GET", "POST"])
@login_required
def productos_importar():
    if request.method == "GET":
        return render_template("importar_productos.html")

    file = request.files.get("archivo")
    if not file or not file.filename:
        flash("Seleccioná un archivo.", "danger")
        return render_template("importar_productos.html")

    ext = file.filename.rsplit(".", 1)[-1].lower()
    if ext not in ("xlsx", "xls", "csv"):
        flash("Formato no soportado. Usá .xlsx o .csv.", "danger")
        return render_template("importar_productos.html")

    file_bytes = file.read()
    try:
        rows = _filas_excel(file_bytes) if ext in ("xlsx", "xls") else _filas_csv(file_bytes)
    except Exception as e:
        flash(f"No se pudo leer el archivo: {e}", "danger")
        return render_template("importar_productos.html")

    productos, errores = _procesar_filas(rows)

    if not productos and not errores:
        flash("El archivo no contiene filas con datos.", "warning")
        return render_template("importar_productos.html")

    if not productos:
        for err in errores:
            flash(err, "danger")
        return render_template("importar_productos.html")

    # Insertar en la base de datos
    db = get_db()
    ids_nuevos = []
    for p in productos:
        cur = db.execute(
            """INSERT INTO productos (nombre, descripcion, precio, unidad, en_oferta, activo)
               VALUES (?,?,?,?,?,1)""",
            (p["nombre"], p["descripcion"], p["precio"], p["unidad"], p["en_oferta"]),
        )
        ids_nuevos.append(cur.lastrowid)
    db.commit()

    # Traer los productos recién insertados para mostrarlos
    placeholders = ",".join("?" * len(ids_nuevos))
    nuevos = [dict(r) for r in db.execute(
        f"SELECT * FROM productos WHERE id IN ({placeholders}) ORDER BY nombre",
        ids_nuevos,
    ).fetchall()]
    db.close()

    for err in errores:
        flash(err, "warning")

    flash(f"Se importaron {len(productos)} productos correctamente.", "success")
    return render_template("importar_resultado.html", productos=nuevos, errores=errores)


@app.route("/productos/plantilla")
@login_required
def productos_plantilla():
    """Genera y descarga el archivo Excel de plantilla."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Productos"

    # Estilos
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="0D6EFD")
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ["nombre", "descripcion", "precio", "en_oferta", "unidad"]
    labels  = ["Nombre *", "Descripción", "Precio *", "En oferta (sí/no)", "Unidad de venta"]
    widths  = [30, 40, 14, 20, 22]

    for col, (label, width) in enumerate(zip(labels, widths), start=1):
        cell = ws.cell(row=1, column=col, value=label)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
        ws.column_dimensions[cell.column_letter].width = width

    ws.row_dimensions[1].height = 22

    # Filas de ejemplo
    ejemplos = [
        ("Empanadas de carne", "Rellenas con cebolla y huevo", 2500, "no",  "por unidad"),
        ("Queso cremoso",      "Cremoso suave",                3800, "sí",  "por kg"),
        ("Salame",             "Estacionado, picante",         4200, "no",  "por 100g"),
        ("Medialunas",         "Con manteca, crocantes",       1200, "sí",  "por unidad"),
        ("Queso en horma",     "Horma entera aprox. 3 kg",    11000, "no",  "por horma entera"),
        ("Queso 1/4",          "Un cuarto de horma",           2800, "sí",  "por 1/4 kg"),
    ]
    nota_fill = PatternFill("solid", fgColor="FFF9E6")
    for row_num, fila in enumerate(ejemplos, start=2):
        for col, val in enumerate(fila, start=1):
            cell = ws.cell(row=row_num, column=col, value=val)
            cell.fill = nota_fill
            cell.border = border

    # Hoja de instrucciones
    ws2 = wb.create_sheet("Instrucciones")
    unidades_str = " / ".join(UNIDADES)
    instrucciones = [
        ("INSTRUCCIONES PARA COMPLETAR LA PLANTILLA", True),
        ("", False),
        ("Columna NOMBRE (obligatorio):", True),
        ("  Escribí el nombre del producto. No puede estar vacío.", False),
        ("", False),
        ("Columna DESCRIPCIÓN:", True),
        ("  Descripción corta (opcional). Puede quedar en blanco.", False),
        ("", False),
        ("Columna PRECIO (obligatorio):", True),
        ("  Número con punto o coma decimal. Ej: 1500 o 1500,50", False),
        ("", False),
        ("Columna EN OFERTA:", True),
        ('  Escribí "sí" o "no". También acepta: si / yes / 1 / no / 0.', False),
        ("", False),
        ("Columna UNIDAD DE VENTA:", True),
        (f"  Opciones válidas: {unidades_str}", False),
        ("  También se aceptan abreviaturas: kg, 100g, 1/4, horma, media horma, etc.", False),
        ("  Si dejás la celda vacía o escribís algo distinto, se usa 'por unidad'.", False),
        ("", False),
        ("IMPORTANTE:", True),
        ("  - La primera fila DEBE ser el encabezado (no la borres).", False),
        ("  - Las fotos las agregás desde el panel, una por una, después de importar.", False),
        ("  - Podés importar el archivo como .xlsx o guardarlo como .csv.", False),
    ]
    ws2.column_dimensions["A"].width = 70
    for i, (texto, bold) in enumerate(instrucciones, start=1):
        cell = ws2.cell(row=i, column=1, value=texto)
        cell.font = Font(bold=bold, size=11 if bold else 10)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return Response(
        output.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=plantilla_productos.xlsx"},
    )


# ── Archivos estáticos subidos ─────────────────────────────────────────────────

@app.route("/uploads/<filename>")
def uploaded_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename)


if __name__ == "__main__":
    init_db()
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    app.run(debug=True, host="0.0.0.0", port=5000)
