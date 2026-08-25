import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "ventas.db")


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS productos (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre      TEXT    NOT NULL,
            descripcion TEXT    NOT NULL DEFAULT '',
            precio      REAL    NOT NULL DEFAULT 0,
            foto        TEXT,
            en_oferta   INTEGER NOT NULL DEFAULT 0,
            oferta_desde TEXT,
            oferta_hasta TEXT,
            activo      INTEGER NOT NULL DEFAULT 1,
            created_at  TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS clientes (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre      TEXT    NOT NULL,
            celular     TEXT    NOT NULL,
            fecha_alta  TEXT    NOT NULL DEFAULT (date('now','localtime')),
            activo      INTEGER NOT NULL DEFAULT 1,
            notas       TEXT    NOT NULL DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS column_mappings (
            fingerprint TEXT    PRIMARY KEY,
            mapping     TEXT    NOT NULL,
            updated_at  TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS variantes (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            producto_id INTEGER NOT NULL REFERENCES productos(id) ON DELETE CASCADE,
            unidad      TEXT    NOT NULL DEFAULT 'por unidad',
            precio      REAL    NOT NULL,
            en_oferta   INTEGER NOT NULL DEFAULT 0,
            orden       INTEGER NOT NULL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS envios (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_id        TEXT,
            cliente_id      INTEGER,
            cliente_nombre  TEXT    NOT NULL,
            celular         TEXT    NOT NULL,
            mensaje         TEXT    NOT NULL,
            twilio_sid      TEXT,
            estado          TEXT    NOT NULL DEFAULT 'pendiente',
            error_msg       TEXT,
            es_prueba       INTEGER NOT NULL DEFAULT 0,
            created_at      TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
        );
    """)

    # Migración: agregar columna unidad si no existe (bases de datos existentes)
    cols = {row[1] for row in conn.execute("PRAGMA table_info(productos)").fetchall()}
    if "unidad" not in cols:
        conn.execute("ALTER TABLE productos ADD COLUMN unidad TEXT NOT NULL DEFAULT 'por unidad'")
        conn.commit()

    # Migración: crear variante inicial para productos existentes sin variantes
    prods_sin_variantes = conn.execute("""
        SELECT id, precio, unidad, en_oferta FROM productos
        WHERE NOT EXISTS (SELECT 1 FROM variantes WHERE variantes.producto_id = productos.id)
    """).fetchall()
    if prods_sin_variantes:
        for p in prods_sin_variantes:
            conn.execute(
                "INSERT INTO variantes (producto_id, precio, unidad, en_oferta, orden) VALUES (?,?,?,?,0)",
                (p["id"], p["precio"] or 0, p["unidad"] or "por unidad", p["en_oferta"] or 0),
            )
        conn.commit()

    conn.close()
