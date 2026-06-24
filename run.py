"""Punto de entrada para producción (gunicorn run:app)."""
import os
from database import init_db
from app import app, UPLOAD_FOLDER

init_db()
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
