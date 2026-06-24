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
            precio      REAL    NOT NULL,
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
    """)
    conn.commit()
    conn.close()
