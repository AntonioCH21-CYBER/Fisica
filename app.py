"""
Física IA – Chatbot para aprender física
Backend: Python (Flask) + SQLite + Google Gemini (gratis) + gTTS (voz)
Interfaz: HTML / CSS / JavaScript  (carpeta web/)

- Cuentas de usuario (crear cuenta / iniciar sesión)
- Apartados y conversaciones guardadas por cuenta
- Texto, imágenes de la galería, preguntas por voz y respuestas con audio
- BASE DE CONOCIMIENTO: los PDF / PowerPoint de la carpeta conocimiento/ son la fuente
  principal del chatbot para responder (búsqueda de fragmentos relevantes en cada pregunta)
- Además, cada conversación puede tener sus propios documentos

En tu PC:   la API key va en .env  ->  python app.py   (datos en SQLite, carpeta datos/)
En Vercel:  variables de entorno GEMINI_API_KEY y DATABASE_URL (Neon Postgres)
"""
import importlib
import importlib.util
import os
import site
import subprocess
import sys

try:  # evita errores de tildes/emojis en la consola de Windows
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


# ───────────── Instala automáticamente lo que falte ─────────────
def instalar_dependencias():
    if os.getenv("VERCEL"):  # en Vercel las librerías se instalan desde requirements.txt
        return
    paquetes = {"flask": "flask", "google.genai": "google-genai", "gtts": "gTTS",
                "dotenv": "python-dotenv", "pypdf": "pypdf", "pptx": "python-pptx"}
    if os.getenv("DATABASE_URL"):
        paquetes["psycopg"] = "psycopg[binary]"
    faltan = []
    for modulo, paquete in paquetes.items():
        try:
            if importlib.util.find_spec(modulo) is None:
                faltan.append(paquete)
        except ModuleNotFoundError:
            faltan.append(paquete)
    if faltan:
        print("Instalando librerias necesarias (solo la primera vez):", ", ".join(faltan))
        subprocess.check_call([sys.executable, "-m", "pip", "install", *faltan])
        importlib.invalidate_caches()
        try:
            if site.getusersitepackages() not in sys.path:
                sys.path.append(site.getusersitepackages())
        except Exception:
            pass


instalar_dependencias()

import hashlib
import io
import json
import math
import os
import re
import secrets
import sqlite3
import threading
import unicodedata
import uuid
import webbrowser
from collections import Counter
from datetime import datetime, timedelta
from functools import wraps
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask, Response, g, jsonify, request, send_from_directory, session
from google import genai
from google.genai import types
from gtts import gTTS
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pypdf import PdfReader
from werkzeug.security import check_password_hash, generate_password_hash

# ───────────────────────── Configuración ─────────────────────────
RAIZ = Path(__file__).resolve().parent
EN_VERCEL = bool(os.getenv("VERCEL"))
load_dotenv(RAIZ / ".env", override=not EN_VERCEL)

DATABASE_URL = (os.getenv("DATABASE_URL") or os.getenv("POSTGRES_URL") or "").strip()
USA_POSTGRES = DATABASE_URL.startswith(("postgres://", "postgresql://"))
if USA_POSTGRES:
    import psycopg
    from psycopg.rows import dict_row

WEB = RAIZ / "web"
CONOCIMIENTO = RAIZ / "conocimiento"          # ← aquí van los PDF/PPTX que usará el chatbot
# En Vercel solo se puede escribir en /tmp (y se borra); por eso allí se usa Postgres.
DATOS = Path("/tmp/fisica_ia") if EN_VERCEL else RAIZ / "datos"
DB = DATOS / "fisica_ia.db"
DATOS.mkdir(parents=True, exist_ok=True)
if not EN_VERCEL:
    CONOCIMIENTO.mkdir(exist_ok=True)
MAX_SUBIDA_MB = 4.4 if EN_VERCEL else 60      # Vercel acepta como máximo 4.5 MB por envío

PUERTO = 5050
APARTADOS_BASE = ["General", "Mecánica", "Electricidad y magnetismo",
                  "Ondas y óptica", "Termodinámica", "Laboratorio"]
MODELOS_PREFERIDOS = ["gemini-3.6-flash", "gemini-3.5-flash", "gemini-flash-latest",
                      "gemini-3.5-flash-lite", "gemini-3.1-flash-lite", "gemini-2.5-flash"]
MIME = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".webp": "image/webp", ".gif": "image/gif"}
MIME_TODOS = {**MIME, ".mp3": "audio/mpeg", ".wav": "audio/wav", ".pdf": "application/pdf",
              ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
              ".txt": "text/plain; charset=utf-8", ".md": "text/markdown; charset=utf-8"}
EXT_DOCS = {".pdf", ".pptx", ".txt", ".md"}
MAX_MENSAJES_CONTEXTO = 30      # cuántos mensajes previos recuerda la IA
MAX_MENSAJES_CON_IMAGEN = 6     # solo se reenvían imágenes de los últimos N mensajes
MAX_CARACTERES_DOCS = 300_000   # texto máximo de documentos de la conversación que se envía a la IA
TAM_FRAGMENTO = 1400            # tamaño de cada fragmento de la base de conocimiento (caracteres)
TOP_FRAGMENTOS = 7              # cuántos fragmentos relevantes se envían por pregunta
MAX_CONTEXTO_BASE = 12_000      # caracteres máximos de la base por pregunta

SYSTEM_PROMPT = """Eres «Física IA», un tutor de física para estudiantes universitarios de ingeniería \
(curso de física general con laboratorio, Arduino y robótica). Responde siempre en español, \
de forma clara, cercana y precisa.

Cómo respondes:
- Puedes responder preguntas de cualquier tema, pero siempre que sea posible conecta la respuesta \
con la física y aprovecha para enseñar.
- En problemas de física sigue este orden: 1) Datos e incógnita, 2) Ley o concepto que se aplica y por qué, \
3) Fórmula, 4) Sustitución con unidades, 5) Resultado en unidades SI con cifras significativas razonables, \
6) Interpretación física (¿el resultado tiene sentido?).
- Escribe las ecuaciones en LaTeX entre $...$ (en línea) o $$...$$ (en bloque).
- Si recibes una imagen (ejercicio, gráfica, montaje experimental, circuito, foto de un fenómeno), \
descríbela brevemente y analízala desde la física.
- En experimentos ayuda con: planteamiento del problema, objetivos, fundamento teórico, diseño, variables \
(independiente, dependiente, controladas), registro de datos, incertidumbre, gráficas, linealización, \
ajuste por mínimos cuadrados, interpretación de la pendiente, conclusiones y recomendaciones.
- Si el estudiante comete un error conceptual, corrígelo con amabilidad y explica por qué.
- Tras explicaciones largas, propone una pregunta corta para comprobar que entendió.
- Sé conciso: si la pregunta es simple, responde simple. No inventes datos; si falta información, \
dilo o asume valores razonables indicándolo."""

PROMPT_DOCUMENTOS = """

DOCUMENTOS DE REFERENCIA:
El estudiante subió documentos (PDF o PowerPoint) a esta conversación. Úsalos como tu fuente principal:
- Redacta conforme al contenido, el enfoque y la terminología de esos documentos.
- Cuando uses información de ellos, cita la fuente así: (Nombre del documento, pág. X) o (diapositiva X).
- Si te piden resumir, explicar o redactar algo, hazlo basándote en los documentos.
- Si la respuesta no está en los documentos, dilo claramente y luego complementa con conocimiento \
general de física, indicando que eso no viene de los documentos."""

PROMPT_BASE = """

BASE DE CONOCIMIENTO DEL CURSO:
Abajo tienes fragmentos de los materiales oficiales del curso (PDF y PowerPoint) que son relevantes \
para la pregunta actual. Son tu FUENTE PRINCIPAL:
- Basa tu respuesta en estos fragmentos y redacta conforme a su contenido, enfoque, notación y terminología.
- Cita la fuente al usar información de ellos, así: (Nombre del archivo, pág. X) o (Nombre del archivo, diap. X).
- Si los fragmentos no contienen la respuesta, dilo brevemente («Esto no aparece en los materiales del curso») \
y responde con tu conocimiento general de física.
- No menciones la palabra «fragmento»; habla de «los materiales del curso».

"""

app = Flask(__name__, static_folder=None)
app.config.update(MAX_CONTENT_LENGTH=int(MAX_SUBIDA_MB * 1024 * 1024),
                  SESSION_COOKIE_SAMESITE="Lax",
                  PERMANENT_SESSION_LIFETIME=timedelta(days=30))


def clave_secreta():
    if os.getenv("SECRET_KEY"):
        return os.getenv("SECRET_KEY")
    if EN_VERCEL:  # debe ser igual en todas las instancias de Vercel
        base = (os.getenv("GEMINI_API_KEY") or "") + DATABASE_URL
        return hashlib.sha256(("fisica-ia|" + base).encode()).hexdigest()
    ruta = DATOS / "secret.key"
    if not ruta.exists():
        ruta.write_text(secrets.token_hex(32), "utf-8")
    return ruta.read_text("utf-8").strip()


app.secret_key = clave_secreta()


# ───────────────────────── Base de datos (SQLite local / Postgres en la nube) ─────────────────────────
ESQUEMA_SQLITE = """
CREATE TABLE IF NOT EXISTS usuarios(
  id INTEGER PRIMARY KEY, usuario TEXT UNIQUE NOT NULL COLLATE NOCASE,
  nombre TEXT NOT NULL, hash TEXT NOT NULL, creado TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS apartados(
  id INTEGER PRIMARY KEY, usuario_id INTEGER NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
  nombre TEXT NOT NULL, orden INTEGER DEFAULT 0, UNIQUE(usuario_id, nombre));
CREATE TABLE IF NOT EXISTS chats(
  id TEXT PRIMARY KEY, usuario_id INTEGER NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
  apartado_id INTEGER NOT NULL REFERENCES apartados(id), titulo TEXT NOT NULL,
  creado TEXT NOT NULL, actualizado TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS mensajes(
  id INTEGER PRIMARY KEY, chat_id TEXT NOT NULL REFERENCES chats(id) ON DELETE CASCADE,
  rol TEXT NOT NULL, texto TEXT, imagenes TEXT DEFAULT '[]', audio TEXT,
  via_voz INTEGER DEFAULT 0, fecha TEXT);
CREATE TABLE IF NOT EXISTS documentos(
  id INTEGER PRIMARY KEY, chat_id TEXT NOT NULL REFERENCES chats(id) ON DELETE CASCADE,
  nombre TEXT NOT NULL, tipo TEXT NOT NULL, archivo TEXT, texto TEXT, unidades INTEGER, fecha TEXT);
CREATE TABLE IF NOT EXISTS archivos(
  id TEXT PRIMARY KEY, usuario_id INTEGER NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
  mime TEXT, datos BLOB, creado TEXT);
CREATE TABLE IF NOT EXISTS indice(ruta TEXT PRIMARY KEY, hash TEXT, datos TEXT);
"""

ESQUEMA_PG = """
CREATE TABLE IF NOT EXISTS usuarios(
  id SERIAL PRIMARY KEY, usuario TEXT UNIQUE NOT NULL,
  nombre TEXT NOT NULL, hash TEXT NOT NULL, creado TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS apartados(
  id SERIAL PRIMARY KEY, usuario_id INTEGER NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
  nombre TEXT NOT NULL, orden INTEGER DEFAULT 0, UNIQUE(usuario_id, nombre));
CREATE TABLE IF NOT EXISTS chats(
  id TEXT PRIMARY KEY, usuario_id INTEGER NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
  apartado_id INTEGER NOT NULL REFERENCES apartados(id) ON DELETE CASCADE, titulo TEXT NOT NULL,
  creado TEXT NOT NULL, actualizado TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS mensajes(
  id SERIAL PRIMARY KEY, chat_id TEXT NOT NULL REFERENCES chats(id) ON DELETE CASCADE,
  rol TEXT NOT NULL, texto TEXT, imagenes TEXT DEFAULT '[]', audio TEXT,
  via_voz INTEGER DEFAULT 0, fuentes TEXT DEFAULT '[]', fecha TEXT);
CREATE TABLE IF NOT EXISTS documentos(
  id SERIAL PRIMARY KEY, chat_id TEXT NOT NULL REFERENCES chats(id) ON DELETE CASCADE,
  nombre TEXT NOT NULL, tipo TEXT NOT NULL, archivo TEXT, texto TEXT, unidades INTEGER, fecha TEXT);
CREATE TABLE IF NOT EXISTS archivos(
  id TEXT PRIMARY KEY, usuario_id INTEGER NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
  mime TEXT, datos BYTEA, creado TEXT);
CREATE TABLE IF NOT EXISTS indice(ruta TEXT PRIMARY KEY, hash TEXT, datos TEXT)
"""

_pg = None
_pg_lock = threading.Lock()


def _correr(sql, args=(), traer=False, con_id=False):
    """Ejecuta SQL (escrito con «?») en SQLite o en Postgres."""
    global _pg
    if USA_POSTGRES:
        sql = sql.replace("?", "%s") + (" RETURNING id" if con_id else "")
        with _pg_lock:
            for intento in (1, 2):
                try:
                    if _pg is None or _pg.closed:
                        _pg = psycopg.connect(DATABASE_URL, autocommit=True, row_factory=dict_row,
                                              prepare_threshold=None, connect_timeout=15)
                    cur = _pg.execute(sql, args)
                    if con_id:
                        return cur.fetchone()["id"]
                    return [dict(f) for f in cur.fetchall()] if traer else None
                except psycopg.OperationalError:
                    _pg = None  # conexión cerrada por inactividad: se reconecta una vez
                    if intento == 2:
                        raise
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    try:
        cur = con.execute(sql, args)
        filas = [dict(f) for f in cur.fetchall()] if traer else None
        con.commit()
        return cur.lastrowid if con_id else filas
    finally:
        con.close()


def consulta(sql, args=(), uno=False):
    filas = _correr(sql, args, traer=True)
    return (filas[0] if filas else None) if uno else filas


def ejecutar(sql, args=(), con_id=False):
    return _correr(sql, args, con_id=con_id)


def iniciar_bd():
    if USA_POSTGRES:
        for sentencia in filter(str.strip, ESQUEMA_PG.split(";")):
            try:
                ejecutar(sentencia)
            except Exception as e:  # p. ej. dos instancias creando la misma tabla a la vez
                print("Aviso al crear tablas:", str(e)[:120])
        return
    con = sqlite3.connect(DB)
    try:
        con.executescript(ESQUEMA_SQLITE)
        if "fuentes" not in [c[1] for c in con.execute("PRAGMA table_info(mensajes)")]:
            con.execute("ALTER TABLE mensajes ADD COLUMN fuentes TEXT DEFAULT '[]'")
        con.commit()
    finally:
        con.close()


iniciar_bd()


# ───────────────────────── Utilidades ─────────────────────────
def ahora():
    return datetime.now().isoformat(timespec="seconds")


def error(msg, codigo=400, **extra):
    return jsonify(error=msg, **extra), codigo


# Imágenes, audios y documentos se guardan DENTRO de la base de datos (tabla «archivos»),
# así funcionan igual en tu PC y en Vercel (donde no se pueden guardar archivos).
def guardar_media(uid, datos: bytes, ext: str) -> str:
    aid = uuid.uuid4().hex[:16]
    ejecutar("INSERT INTO archivos(id, usuario_id, mime, datos, creado) VALUES (?,?,?,?,?)",
             (aid, uid, MIME_TODOS.get(ext, "application/octet-stream"), datos, ahora()))
    return f"media/{aid}{ext}"


def id_media(rel):
    m = re.fullmatch(r"media/([0-9a-f]{16})\.\w+", rel or "")
    return m.group(1) if m else None


def leer_media(rel):
    aid = id_media(rel)
    if aid:
        fila = consulta("SELECT datos FROM archivos WHERE id=?", (aid,), uno=True)
        return bytes(fila["datos"]) if fila else None
    ruta = DATOS / (rel or "")  # archivos de la versión anterior (carpeta datos/media)
    return ruta.read_bytes() if rel and ruta.is_file() else None


def borrar_media(rel):
    aid = id_media(rel)
    if aid:
        ejecutar("DELETE FROM archivos WHERE id=?", (aid,))
    elif rel and (DATOS / rel).is_file():
        (DATOS / rel).unlink()


def enviar_bytes(datos, mime):
    """Respuesta con soporte de «Range» (necesario para reproducir audio en algunos navegadores)."""
    total = len(datos)
    m = re.match(r"bytes=(\d*)-(\d*)", request.headers.get("Range", ""))
    if m and (m.group(1) or m.group(2)):
        if m.group(1):
            ini, fin = int(m.group(1)), int(m.group(2)) if m.group(2) else total - 1
        else:
            ini, fin = max(0, total - int(m.group(2))), total - 1
        fin = min(fin, total - 1)
        if ini > fin:
            return Response(status=416, headers={"Content-Range": f"bytes */{total}"})
        r = Response(datos[ini:fin + 1], 206, mimetype=mime)
        r.headers["Content-Range"] = f"bytes {ini}-{fin}/{total}"
    else:
        r = Response(datos, mimetype=mime)
    r.headers["Accept-Ranges"] = "bytes"
    r.headers["Cache-Control"] = "private, max-age=86400"
    return r


def titulo_de(texto):
    linea = texto.strip().splitlines()[0] if texto.strip() else "Conversación"
    return linea[:45] + ("…" if len(linea) > 45 else "")


# ───────────────────────── Cuentas ─────────────────────────
def login_requerido(f):
    @wraps(f)
    def envoltura(*a, **k):
        uid = session.get("uid")
        if not uid or not consulta("SELECT id FROM usuarios WHERE id=?", (uid,), uno=True):
            session.clear()
            return error("Inicia sesión para continuar.", 401, login=True)
        g.uid = uid
        return f(*a, **k)
    return envoltura


def info_usuario(uid):
    return consulta("SELECT id, usuario, nombre FROM usuarios WHERE id=?", (uid,), uno=True)


def iniciar_sesion(uid):
    session.clear()
    session.permanent = True
    session["uid"] = uid


@app.post("/api/registro")
def registro():
    d = request.get_json(silent=True) or {}
    nombre = (d.get("nombre") or "").strip()[:60]
    usuario = (d.get("usuario") or "").strip().lower()
    clave_ = d.get("clave") or ""
    if not re.fullmatch(r"[a-z0-9_.]{3,30}", usuario):
        return error("El usuario debe tener de 3 a 30 caracteres: letras, números, punto o guion bajo.")
    if len(clave_) < 6:
        return error("La contraseña debe tener al menos 6 caracteres.")
    if consulta("SELECT id FROM usuarios WHERE usuario=?", (usuario,), uno=True):
        return error("Ese usuario ya existe. Elige otro o inicia sesión.")
    uid = ejecutar("INSERT INTO usuarios(usuario, nombre, hash, creado) VALUES (?,?,?,?)",
                   (usuario, nombre or usuario, generate_password_hash(clave_), ahora()), con_id=True)
    for i, a in enumerate(APARTADOS_BASE):
        ejecutar("INSERT INTO apartados(usuario_id, nombre, orden) VALUES (?,?,?)", (uid, a, i))
    iniciar_sesion(uid)
    return jsonify(usuario=info_usuario(uid))


@app.post("/api/login")
def login():
    d = request.get_json(silent=True) or {}
    u = consulta("SELECT * FROM usuarios WHERE usuario=?", ((d.get("usuario") or "").strip().lower(),), uno=True)
    if not u or not check_password_hash(u["hash"], d.get("clave") or ""):
        return error("Usuario o contraseña incorrectos.", 401)
    iniciar_sesion(u["id"])
    return jsonify(usuario=info_usuario(u["id"]))


@app.post("/api/logout")
def logout():
    session.clear()
    return jsonify(ok=True)


# ───────────────────────── Apartados y chats ─────────────────────────
def nombres_apartados(uid):
    return [a["nombre"] for a in consulta(
        "SELECT nombre FROM apartados WHERE usuario_id=? ORDER BY orden, id", (uid,))]


def id_apartado(uid, nombre):
    a = consulta("SELECT id FROM apartados WHERE usuario_id=? AND nombre=?", (uid, nombre), uno=True)
    if a:
        return a["id"]
    a = consulta("SELECT id FROM apartados WHERE usuario_id=? ORDER BY orden, id", (uid,), uno=True)
    if a:
        return a["id"]
    return ejecutar("INSERT INTO apartados(usuario_id, nombre) VALUES (?, 'General')", (uid,), con_id=True)


def crear_chat(uid, apartado, titulo="Nueva conversación"):
    cid = uuid.uuid4().hex[:12]
    ejecutar("INSERT INTO chats(id, usuario_id, apartado_id, titulo, creado, actualizado) VALUES (?,?,?,?,?,?)",
             (cid, uid, id_apartado(uid, apartado), titulo, ahora(), ahora()))
    return cid


def es_mio(uid, cid):
    return bool(cid) and consulta("SELECT 1 FROM chats WHERE id=? AND usuario_id=?", (cid, uid), uno=True)


def tocar(cid):
    ejecutar("UPDATE chats SET actualizado=? WHERE id=?", (ahora(), cid))


def mensajes_de(cid):
    msgs = consulta("""SELECT rol, texto, imagenes, audio, via_voz, fuentes, fecha
                       FROM mensajes WHERE chat_id=? ORDER BY id""", (cid,))
    for m in msgs:
        m["imagenes"] = json.loads(m["imagenes"] or "[]")
        m["fuentes"] = json.loads(m["fuentes"] or "[]")
        m["via_voz"] = bool(m["via_voz"])
    return msgs


def chat_completo(uid, cid):
    c = consulta("""SELECT c.id, c.titulo, c.creado, c.actualizado, a.nombre AS apartado
                    FROM chats c JOIN apartados a ON a.id = c.apartado_id
                    WHERE c.id=? AND c.usuario_id=?""", (cid, uid), uno=True)
    if not c:
        return None
    c["mensajes"] = mensajes_de(cid)
    c["documentos"] = consulta("""SELECT id, nombre, tipo, unidades, length(texto) AS caracteres
                                  FROM documentos WHERE chat_id=? ORDER BY id""", (cid,))
    return c


def borrar_chat(cid):
    for m in consulta("SELECT imagenes, audio FROM mensajes WHERE chat_id=?", (cid,)):
        for rel in json.loads(m["imagenes"] or "[]") + [m["audio"]]:
            borrar_media(rel)
    for d in consulta("SELECT archivo FROM documentos WHERE chat_id=?", (cid,)):
        borrar_media(d["archivo"])
    ejecutar("DELETE FROM chats WHERE id=?", (cid,))


# ───────────────────────── Documentos (PDF / PPTX) ─────────────────────────
def textos_de_forma(forma):
    textos = []
    if forma.shape_type == MSO_SHAPE_TYPE.GROUP:
        for sub in forma.shapes:
            textos += textos_de_forma(sub)
    if getattr(forma, "has_text_frame", False) and forma.text_frame.text.strip():
        textos.append(forma.text_frame.text.strip())
    if getattr(forma, "has_table", False):
        for fila in forma.table.rows:
            textos.append(" | ".join(c.text.strip() for c in fila.cells))
    return textos


def extraer_unidades(nombre, datos):
    """Devuelve ([(«pág. N» | «diap. N», texto), ...], número de páginas/diapositivas)."""
    ext = Path(nombre).suffix.lower()
    if ext == ".pdf":
        lector = PdfReader(io.BytesIO(datos))
        if lector.is_encrypted:
            lector.decrypt("")
        unidades = []
        for i, pagina in enumerate(lector.pages, 1):
            try:
                t = (pagina.extract_text() or "").strip()
            except Exception:
                t = ""
            if t:
                unidades.append((f"pág. {i}", t))
        return unidades, len(lector.pages)
    if ext == ".pptx":
        prs = Presentation(io.BytesIO(datos))
        unidades = []
        for i, diapo in enumerate(prs.slides, 1):
            textos = []
            for forma in diapo.shapes:
                textos += textos_de_forma(forma)
            if diapo.has_notes_slide and diapo.notes_slide.notes_text_frame.text.strip():
                textos.append("Notas: " + diapo.notes_slide.notes_text_frame.text.strip())
            if textos:
                unidades.append((f"diap. {i}", "\n".join(textos)))
        return unidades, len(prs.slides)
    return [("texto", datos.decode("utf-8", errors="replace"))], 1


def extraer_texto(nombre, datos):
    """Devuelve (texto con marcas de página/diapositiva, número de páginas/diapositivas)."""
    unidades, total = extraer_unidades(nombre, datos)
    return "\n\n".join(f"[{u}]\n{t}" for u, t in unidades), total


@app.post("/api/documentos")
@login_requerido
def subir_documentos():
    uid = g.uid
    archivos = request.files.getlist("archivos")
    if not archivos:
        return error("Elige un archivo PDF o PowerPoint.")
    cid = request.form.get("chat_id")
    if not es_mio(uid, cid):
        cid = None
    avisos, agregados = [], 0
    for a in archivos:
        nombre = Path(a.filename or "documento").name
        ext = Path(nombre).suffix.lower()
        if ext == ".ppt":
            avisos.append(f"«{nombre}»: los .ppt antiguos no se pueden leer. En PowerPoint usa "
                          "Archivo → Guardar como → .pptx o PDF.")
            continue
        if ext not in EXT_DOCS:
            avisos.append(f"«{nombre}»: formato no admitido (usa PDF o .pptx).")
            continue
        datos = a.read()
        try:
            texto, unidades = extraer_texto(nombre, datos)
        except Exception as e:
            avisos.append(f"«{nombre}»: no se pudo leer ({str(e)[:80]}).")
            continue
        if not texto.strip() and ext != ".pdf":
            avisos.append(f"«{nombre}»: no tiene texto que se pueda leer.")
            continue
        if not cid:
            cid = crear_chat(uid, request.form.get("apartado", ""), titulo=f"📄 {titulo_de(nombre)}")
        ejecutar("""INSERT INTO documentos(chat_id, nombre, tipo, archivo, texto, unidades, fecha)
                    VALUES (?,?,?,?,?,?,?)""",
                 (cid, nombre, ext, guardar_media(uid, datos, ext), texto, unidades, ahora()))
        agregados += 1
    if not agregados:
        return error(" ".join(avisos) or "No se pudo agregar el documento.")
    tocar(cid)
    return jsonify(chat=chat_completo(uid, cid), avisos=avisos)


@app.delete("/api/documentos/<int:did>")
@login_requerido
def quitar_documento(did):
    d = consulta("""SELECT d.id, d.chat_id, d.archivo FROM documentos d JOIN chats c ON c.id = d.chat_id
                    WHERE d.id=? AND c.usuario_id=?""", (did, g.uid), uno=True)
    if not d:
        return error("No existe ese documento.", 404)
    borrar_media(d["archivo"])
    ejecutar("DELETE FROM documentos WHERE id=?", (did,))
    return jsonify(chat=chat_completo(g.uid, d["chat_id"]))


# ───────────────────────── IA (Gemini) ─────────────────────────
def clave():
    k = (os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or "").strip().strip('"').strip("'")
    return "" if (not k or k.startswith("pega_")) else k


_cache_modelos = {}


def modelos_disponibles():
    """Modelos Gemini de texto disponibles para la clave (los recomendados primero)."""
    k = clave()
    if not k:
        return MODELOS_PREFERIDOS
    if k in _cache_modelos:
        return _cache_modelos[k]
    nombres = []
    try:
        for m in genai.Client(api_key=k).models.list():
            n = m.name.replace("models/", "")
            acciones = (getattr(m, "supported_actions", None)
                        or getattr(m, "supported_generation_methods", None) or [])
            if ((not acciones or "generateContent" in acciones) and n.startswith("gemini")
                    and not re.search(r"tts|image|embed|live|audio|robotics|computer", n)):
                nombres.append(n)
    except Exception:
        pass
    if not nombres:
        return MODELOS_PREFERIDOS

    def orden(n):
        pref = MODELOS_PREFERIDOS.index(n) if n in MODELOS_PREFERIDOS else 99
        return (pref, "flash" not in n, "preview" in n or "exp" in n, n)

    _cache_modelos[k] = sorted(nombres, key=orden)
    return _cache_modelos[k]


def es_modelo_invalido(e):
    t = str(e).lower()
    return ("404" in t or "not found" in t or "not supported" in t or "no longer available" in t
            or "deprecated" in t or "is not available" in t)


def generar(client, modelo, contents, config=None):
    """Llama a Gemini; si el modelo no existe, prueba automáticamente con otros."""
    candidatos = [modelo] + [m for m in modelos_disponibles() + MODELOS_PREFERIDOS if m != modelo]
    vistos, ultimo = set(), None
    for m in candidatos:
        if m in vistos:
            continue
        vistos.add(m)
        try:
            return client.models.generate_content(model=m, contents=contents, config=config), m
        except Exception as e:
            if not es_modelo_invalido(e):
                raise
            ultimo = e
    raise ultimo


def contexto_documentos(cid):
    """Arma el bloque de documentos de referencia que se envía antes del historial."""
    docs = consulta("SELECT nombre, tipo, archivo, texto, unidades FROM documentos WHERE chat_id=? ORDER BY id", (cid,))
    if not docs:
        return []
    parts = [types.Part.from_text(text="Estos son mis documentos de referencia para esta conversación:")]
    restante = MAX_CARACTERES_DOCS
    for d in docs:
        texto = d["texto"] or ""
        escaneado = len(texto.strip()) < 40 * max(d["unidades"] or 1, 1) and d["tipo"] == ".pdf"
        pdf = leer_media(d["archivo"]) if escaneado else None
        if pdf and len(pdf) < 15 * 1024 * 1024:
            # PDF escaneado (sin texto): se envía el archivo para que la IA lo lea directamente
            parts.append(types.Part.from_text(text=f"=== DOCUMENTO: {d['nombre']} (PDF escaneado) ==="))
            parts.append(types.Part.from_bytes(data=pdf, mime_type="application/pdf"))
            continue
        if restante <= 0:
            parts.append(types.Part.from_text(text=f"[El documento «{d['nombre']}» no se incluyó por espacio]"))
            continue
        recorte = texto[:restante]
        restante -= len(recorte)
        aviso = "\n[... documento recortado por su longitud ...]" if len(recorte) < len(texto) else ""
        parts.append(types.Part.from_text(
            text=f"=== INICIO DEL DOCUMENTO: {d['nombre']} ===\n{recorte}{aviso}\n"
                 f"=== FIN DEL DOCUMENTO: {d['nombre']} ==="))
    return [types.Content(role="user", parts=parts),
            types.Content(role="model", parts=[types.Part.from_text(
                text="Entendido. Usaré estos documentos como fuente principal y citaré el documento "
                     "y la página o diapositiva correspondiente.")])]


def a_contents(mensajes):
    """Convierte el historial guardado al formato de Gemini."""
    mensajes = mensajes[-MAX_MENSAJES_CONTEXTO:]
    while mensajes and mensajes[0]["rol"] != "user":
        mensajes = mensajes[1:]
    contents = []
    for i, m in enumerate(mensajes):
        parts = []
        reciente = i >= len(mensajes) - MAX_MENSAJES_CON_IMAGEN
        for rel in m.get("imagenes", []):
            datos = leer_media(rel) if reciente else None
            if datos:
                parts.append(types.Part.from_bytes(
                    data=datos, mime_type=MIME.get(Path(rel).suffix.lower(), "image/jpeg")))
            elif not reciente:
                parts.append(types.Part.from_text(text="[El estudiante envió una imagen antes]"))
        if m.get("texto"):
            parts.append(types.Part.from_text(text=m["texto"]))
        if not parts:
            continue
        if contents and contents[-1].role == m["rol"]:
            contents[-1].parts.extend(parts)
        else:
            contents.append(types.Content(role=m["rol"], parts=parts))
    return contents


def transcribir(client, modelo, audio_bytes):
    r, _ = generar(client, modelo, [
        types.Part.from_bytes(data=audio_bytes, mime_type="audio/wav"),
        "Transcribe literalmente este audio (probablemente en español). "
        "Devuelve solo la transcripción, sin comentarios."])
    return (r.text or "").strip()


def error_amigable(e):
    s = str(e)
    if "429" in s or "RESOURCE_EXHAUSTED" in s:
        return ("⏳ Se alcanzó el límite gratuito de la API. Espera un minuto o elige otro modelo. "
                "Si subiste documentos muy grandes, prueba con uno más pequeño.")
    if "API key" in s or "API_KEY" in s or "401" in s or "403" in s:
        return "🔑 La API key no es válida. Revisa el archivo .env y reinicia con python app.py."
    if es_modelo_invalido(e):
        return "🤖 Ese modelo no está disponible. Elige otro en Ajustes."
    if "getaddrinfo" in s or "ConnectError" in s or "Connection" in s:
        return "🌐 No hay conexión a internet."
    return f"❌ Error: {s[:300]}"


# ───────────────────────── Voz (texto → audio) ─────────────────────────
def latex_a_voz(s):
    s = re.sub(r"\\[dt]?frac\{([^{}]*)\}\{([^{}]*)\}", r" \1 sobre \2 ", s)
    s = re.sub(r"\\sqrt\{([^{}]*)\}", r" raíz de \1 ", s)
    s = re.sub(r"\\(?:text|mathrm|mathbf|vec|operatorname)\{([^{}]*)\}", r"\1", s)
    s = re.sub(r"\^\{?2\}?(?![\d])", " al cuadrado", s)
    s = re.sub(r"\^\{?3\}?(?![\d])", " al cubo", s)
    s = re.sub(r"\^\{?(-?[\w.]+)\}?", r" elevado a \1", s)
    s = re.sub(r"_\{?(\w+)\}?", r" \1", s)
    for k, v in {r"\cdot": " por ", r"\times": " por ", r"\approx": " aproximadamente ",
                 r"\Delta": " delta ", r"\pi": " pi ", r"\theta": " theta ", r"\omega": " omega ",
                 r"\alpha": " alfa ", r"\lambda": " lambda ", r"\mu": " mu ", r"\rho": " ro ",
                 r"\sigma": " sigma ", r"\pm": " más menos ", r"\Rightarrow": " entonces ",
                 r"\rightarrow": " hacia ", r"\,": " ", r"\;": " ", "^\\circ": " grados"}.items():
        s = s.replace(k, v)
    s = s.replace("=", " igual a ").replace("+", " más ").replace("/", " entre ")
    s = re.sub(r"(?<=\s)-(?=\s|\d)", " menos ", s)
    s = re.sub(r"\\([a-zA-Z]+)", r"\1", s)
    return s.replace("{", " ").replace("}", " ").replace("\\", " ")


def texto_para_voz(t):
    t = re.sub(r"```.*?```", " Te dejé el código en pantalla. ", t, flags=re.S)
    t = re.sub(r"\$\$(.+?)\$\$", lambda m: " " + latex_a_voz(m.group(1)) + ". ", t, flags=re.S)
    t = re.sub(r"\$(.+?)\$", lambda m: latex_a_voz(m.group(1)), t)
    t = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", t)
    t = re.sub(r"[*#_`>|~]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()[:3000]


def generar_audio(uid, texto):
    try:
        limpio = texto_para_voz(texto)
        if not limpio:
            return None
        buf = io.BytesIO()
        gTTS(limpio, lang="es", tld="com.mx").write_to_fp(buf)
        return guardar_media(uid, buf.getvalue(), ".mp3")
    except Exception as e:
        print("No se pudo generar el audio:", e)
        return None


# ───────────────────────── Base de conocimiento (carpeta conocimiento/) ─────────────────────────
STOPWORDS = set("""de la que el en y a los del se las por un para con no una su al lo como mas pero sus le ya o
este si porque esta entre cuando muy sin sobre tambien me hasta hay donde quien desde todo nos durante
todos uno les ni contra otros ese eso ante ellos e esto mi antes algunos unos yo otro otras otra tanto esa
estos mucho quienes nada muchos cual poco ella estar estas algunas algo nosotros es son fue ser ha han hace
cuales cuanto explica explicame dime puedes favor ejemplo quiero saber podrias ayudame entonces asi tiene
tienen cada debe pueden puede segun hacer the and of to in is""".split())


def normalizar(t):
    t = unicodedata.normalize("NFD", t.lower())
    return "".join(c for c in t if unicodedata.category(c) != "Mn")


def tokens(t):
    """Palabras clave normalizadas (sin tildes, sin palabras vacías, raíz de 7 letras)."""
    return [w[:7] for w in re.findall(r"[a-z0-9]+", normalizar(t)) if len(w) > 1 and w not in STOPWORDS]


def trocear(texto, tam=TAM_FRAGMENTO, solape=200):
    texto = re.sub(r"[ \t]+", " ", texto).strip()
    if len(texto) <= tam:
        return [texto] if texto else []
    trozos, i = [], 0
    while i < len(texto):
        fin = min(i + tam, len(texto))
        if fin < len(texto):
            corte = max(texto.rfind("\n", i + tam // 2, fin), texto.rfind(". ", i + tam // 2, fin))
            if corte > i:
                fin = corte + 1
        trozos.append(texto[i:fin].strip())
        if fin >= len(texto):
            break
        i = max(fin - solape, i + 1)
    return [t for t in trozos if t]


def leer_pdf_con_ia(ruta, paginas):
    """Para PDF escaneados (imágenes sin texto): la IA transcribe el texto una sola vez."""
    k = clave()
    if not k or ruta.stat().st_size > 19 * 1024 * 1024:
        return []
    r, _ = generar(genai.Client(api_key=k), MODELOS_PREFERIDOS[0], [
        types.Part.from_bytes(data=ruta.read_bytes(), mime_type="application/pdf"),
        "Transcribe todo el texto de este documento (incluye fórmulas en LaTeX y describe brevemente "
        "gráficos y tablas). Antes del contenido de cada página escribe exactamente [Página N]. "
        "Devuelve solo la transcripción."])
    texto = r.text or ""
    partes = re.split(r"\[P[áa]gina\s+(\d+)\]", texto)
    if len(partes) < 3:
        return [("pág. 1-%d" % paginas, texto.strip())] if texto.strip() else []
    return [(f"pág. {partes[i]}", partes[i + 1].strip()) for i in range(1, len(partes) - 1, 2)
            if partes[i + 1].strip()]


def indexar_archivo(ruta, firma):
    nombre = ruta.name
    info = {"firma": firma, "nombre": nombre, "paginas": 0, "fragmentos": [], "aviso": None}
    if ruta.suffix.lower() == ".ppt":
        info["aviso"] = "Formato .ppt antiguo: ábrelo en PowerPoint y guárdalo como .pptx o PDF."
        return info
    try:
        unidades, info["paginas"] = extraer_unidades(nombre, ruta.read_bytes())
        caracteres = sum(len(t) for _, t in unidades)
        if ruta.suffix.lower() == ".pdf" and caracteres < 40 * max(info["paginas"], 1):
            print(f"  · «{nombre}» parece escaneado: leyéndolo con la IA (solo esta vez)...")
            try:
                unidades = leer_pdf_con_ia(ruta, info["paginas"]) or unidades
            except Exception as e:
                info["aviso"] = "PDF escaneado: no se pudo leer con la IA (" + error_amigable(e)[:80] + ")."
        for u, t in unidades:
            for trozo in trocear(t):
                info["fragmentos"].append({"u": u, "t": trozo})
        if not info["fragmentos"] and not info["aviso"]:
            info["aviso"] = "No se encontró texto (¿PDF escaneado? Revisa que el .env tenga tu API key)."
    except Exception as e:
        info["aviso"] = f"No se pudo leer: {str(e)[:100]}"
    return info


class BaseConocimiento:
    EXTENSIONES = {".pdf", ".pptx", ".ppt", ".txt", ".md"}

    def __init__(self):
        self.lock = threading.Lock()
        self.firma = None
        self.archivos = {}
        self.frags, self.df, self.avgdl = [], Counter(), 1.0

    def rutas(self):
        if not CONOCIMIENTO.is_dir():
            return []
        return sorted(p for p in CONOCIMIENTO.rglob("*")
                      if p.is_file() and p.suffix.lower() in self.EXTENSIONES
                      and not p.name.startswith(("_", "~$", ".")))

    def actualizar(self, forzar=False):
        with self.lock:
            rutas = self.rutas()
            firma = tuple((str(p.relative_to(CONOCIMIENTO)), p.stat().st_size, int(p.stat().st_mtime))
                          for p in rutas)
            if firma == self.firma and not forzar:
                return
            # El índice se guarda en la base de datos: así no se vuelve a leer cada archivo
            # (ni a gastar IA en PDF escaneados) en cada arranque de Vercel.
            cache = {f["ruta"]: f for f in consulta("SELECT ruta, hash, datos FROM indice")}
            nuevos = {}
            for p in rutas:
                rel = p.relative_to(CONOCIMIENTO).as_posix()
                h = hashlib.sha1(p.read_bytes()).hexdigest()
                previo, info = cache.get(rel), None
                if previo and previo["hash"] == h:
                    info = json.loads(previo["datos"])
                    if forzar and not info.get("fragmentos"):
                        info = None  # reintentar los que fallaron
                if info is None:
                    info = indexar_archivo(p, h)
                    ejecutar("""INSERT INTO indice(ruta, hash, datos) VALUES (?,?,?)
                                ON CONFLICT(ruta) DO UPDATE SET hash=excluded.hash, datos=excluded.datos""",
                             (rel, h, json.dumps(info, ensure_ascii=False)))
                nuevos[rel] = info
            for rel in set(cache) - set(nuevos):
                ejecutar("DELETE FROM indice WHERE ruta=?", (rel,))
            self.archivos, self.firma = nuevos, firma
            self.construir()

    def construir(self):
        frags, df = [], Counter()
        for rel, info in self.archivos.items():
            extra = tokens(Path(info["nombre"]).stem)
            for fr in info["fragmentos"]:
                tf = Counter(tokens(fr["t"]) + extra)
                frags.append({"archivo": info["nombre"], "u": fr["u"], "t": fr["t"],
                              "tf": tf, "len": sum(tf.values()) or 1})
                df.update(tf.keys())
        self.avgdl = (sum(f["len"] for f in frags) / len(frags)) if frags else 1.0
        self.frags, self.df = frags, df

    def buscar(self, consulta, k=TOP_FRAGMENTOS):
        """Búsqueda BM25: devuelve los fragmentos más relevantes para la consulta."""
        q = Counter(tokens(consulta))
        if not q or not self.frags:
            return []
        n, k1, b = len(self.frags), 1.5, 0.75
        puntajes = []
        for i, f in enumerate(self.frags):
            s = 0.0
            for termino, peso in q.items():
                tf = f["tf"].get(termino)
                if tf:
                    idf = math.log(1 + (n - self.df[termino] + 0.5) / (self.df[termino] + 0.5))
                    s += peso * idf * tf * (k1 + 1) / (tf + k1 * (1 - b + b * f["len"] / self.avgdl))
            if s > 0:
                puntajes.append((s, i))
        if not puntajes:
            return []
        puntajes.sort(reverse=True)
        mejor = puntajes[0][0]
        return [self.frags[i] for s, i in puntajes[:k] if s >= 0.4 * mejor]

    def resumen(self):
        return [{"nombre": rel, "paginas": i["paginas"], "fragmentos": len(i["fragmentos"]), "aviso": i["aviso"]}
                for rel, i in self.archivos.items()]


BASE = BaseConocimiento()


def contexto_base(texto_actual, cid):
    """Busca en la base los fragmentos relevantes. Devuelve (texto para la IA, lista de fuentes)."""
    BASE.actualizar()
    previos = [m["texto"] for m in mensajes_de(cid) if m["rol"] == "user"][-3:-1]
    frags = BASE.buscar(" ".join([texto_actual, texto_actual] + previos))
    bloques, fuentes, usados = [], [], 0
    for i, f in enumerate(frags, 1):
        if usados + len(f["t"]) > MAX_CONTEXTO_BASE and bloques:
            break
        usados += len(f["t"])
        bloques.append(f"[{i}] ({f['archivo']}, {f['u']})\n{f['t']}")
        fuente = f"{f['archivo']} ({f['u']})"
        if fuente not in fuentes:
            fuentes.append(fuente)
    return ("\n\n".join(bloques), fuentes) if bloques else ("", [])


@app.get("/api/conocimiento")
@login_requerido
def ver_conocimiento():
    BASE.actualizar()
    return jsonify(archivos=BASE.resumen(), carpeta=str(CONOCIMIENTO))


@app.post("/api/conocimiento/actualizar")
@login_requerido
def recargar_conocimiento():
    BASE.actualizar(forzar=True)
    return jsonify(archivos=BASE.resumen(), carpeta=str(CONOCIMIENTO))


# ───────────────────────── Rutas (API) ─────────────────────────
@app.get("/")
def inicio():
    return send_from_directory(WEB, "index.html")


@app.get("/favicon.ico")
def favicon():
    return Response(status=204)


@app.errorhandler(413)
def demasiado_grande(_e):
    return error(f"El archivo es muy grande (máximo {MAX_SUBIDA_MB:g} MB por envío). Para materiales "
                 "grandes del curso, ponlos en la carpeta «conocimiento».", 413)


@app.get("/media/<path:ruta>")
@login_requerido
def media(ruta):
    aid = id_media("media/" + ruta)
    if aid:
        fila = consulta("SELECT mime, datos FROM archivos WHERE id=? AND usuario_id=?", (aid, g.uid), uno=True)
        if not fila:
            return error("No encontrado.", 404)
        return enviar_bytes(bytes(fila["datos"]), fila["mime"])
    if not ruta.startswith(f"u{g.uid}/"):  # archivos de la versión anterior
        return error("No autorizado.", 403)
    return send_from_directory(DATOS / "media", ruta)


@app.get("/api/estado")
def estado():
    uid = session.get("uid")
    return jsonify(clave=bool(clave()), modelos=modelos_disponibles(),
                   usuario=info_usuario(uid) if uid else None,
                   base=sum(1 for a in BASE.resumen() if a["fragmentos"]),
                   modo="postgres" if USA_POSTGRES else ("temporal" if EN_VERCEL else "local"),
                   limite_mb=MAX_SUBIDA_MB)


@app.get("/api/apartados")
@login_requerido
def listar_apartados():
    return jsonify(apartados=nombres_apartados(g.uid))


@app.post("/api/apartados")
@login_requerido
def nuevo_apartado():
    nombre = ((request.get_json(silent=True) or {}).get("nombre") or "").strip()[:40]
    if not nombre:
        return error("Escribe un nombre para el apartado.")
    if nombre not in nombres_apartados(g.uid):
        ejecutar("INSERT INTO apartados(usuario_id, nombre, orden) VALUES (?,?,?)", (g.uid, nombre, 999))
    return jsonify(apartados=nombres_apartados(g.uid))


@app.delete("/api/apartados/<path:nombre>")
@login_requerido
def borrar_apartado(nombre):
    aps = nombres_apartados(g.uid)
    if nombre not in aps or len(aps) <= 1:
        return error("No se puede eliminar ese apartado.")
    origen = id_apartado(g.uid, nombre)
    destino = id_apartado(g.uid, next(a for a in aps if a != nombre))
    ejecutar("UPDATE chats SET apartado_id=? WHERE apartado_id=?", (destino, origen))  # se mueven, no se pierden
    ejecutar("DELETE FROM apartados WHERE id=?", (origen,))
    return jsonify(apartados=nombres_apartados(g.uid))


@app.get("/api/chats")
@login_requerido
def listar_chats():
    chats = consulta("""SELECT c.id, c.titulo, c.actualizado FROM chats c JOIN apartados a ON a.id = c.apartado_id
                        WHERE c.usuario_id=? AND a.nombre=? ORDER BY c.actualizado DESC""",
                     (g.uid, request.args.get("apartado", "")))
    return jsonify(chats=chats)


@app.get("/api/chats/<cid>")
@login_requerido
def ver_chat(cid):
    chat = chat_completo(g.uid, cid)
    return jsonify(chat=chat) if chat else error("No existe esa conversación.", 404)


@app.patch("/api/chats/<cid>")
@login_requerido
def editar_chat(cid):
    if not es_mio(g.uid, cid):
        return error("No existe esa conversación.", 404)
    d = request.get_json(silent=True) or {}
    if (d.get("titulo") or "").strip():
        ejecutar("UPDATE chats SET titulo=? WHERE id=?", (d["titulo"].strip()[:80], cid))
    if d.get("apartado") in nombres_apartados(g.uid):
        ejecutar("UPDATE chats SET apartado_id=? WHERE id=?", (id_apartado(g.uid, d["apartado"]), cid))
    tocar(cid)
    return jsonify(chat=chat_completo(g.uid, cid))


@app.delete("/api/chats/<cid>")
@login_requerido
def eliminar_chat(cid):
    if es_mio(g.uid, cid):
        borrar_chat(cid)
    return jsonify(ok=True)


@app.post("/api/mensaje")
@login_requerido
def mensaje():
    uid = g.uid
    k = clave()
    if not k:
        return error("🔑 Falta tu API key. Pégala en el archivo .env y reinicia con python app.py.", 503)
    f = request.form
    modelo = f.get("modelo") or MODELOS_PREFERIDOS[0]
    cid = f.get("chat_id") if es_mio(uid, f.get("chat_id")) else crear_chat(uid, f.get("apartado", ""))
    client = genai.Client(api_key=k)
    texto = (f.get("texto") or "").strip()
    imagenes, audio_rel, via_voz = [], None, False

    for img in request.files.getlist("imagenes"):
        ext = Path(img.filename or "").suffix.lower()
        imagenes.append(guardar_media(uid, img.read(), ext if ext in MIME else ".jpg"))

    try:
        audio = request.files.get("audio")
        if audio:
            datos = audio.read()
            audio_rel, via_voz = guardar_media(uid, datos, ".wav"), True
            texto = (texto + "\n\n" + transcribir(client, modelo, datos)).strip()
        if not texto:
            if not imagenes:
                return error("No te escuché bien. Graba de nuevo o escribe tu pregunta.",
                             chat=chat_completo(uid, cid))
            texto = "Analiza esta imagen desde el punto de vista de la física."

        ejecutar("""INSERT INTO mensajes(chat_id, rol, texto, imagenes, audio, via_voz, fecha)
                    VALUES (?, 'user', ?, ?, ?, ?, ?)""",
                 (cid, texto, json.dumps(imagenes), audio_rel, int(via_voz), ahora()))
        ejecutar("UPDATE chats SET titulo=? WHERE id=? AND titulo='Nueva conversación'", (titulo_de(texto), cid))
        tocar(cid)

        docs = contexto_documentos(cid)
        texto_base, fuentes = contexto_base(texto, cid) if f.get("base", "1") == "1" else ("", [])
        instruccion = SYSTEM_PROMPT + (PROMPT_DOCUMENTOS if docs else "")
        if texto_base:
            instruccion += PROMPT_BASE + texto_base
        config = types.GenerateContentConfig(system_instruction=instruccion)
        r, modelo = generar(client, modelo, docs + a_contents(mensajes_de(cid)), config)
        respuesta = (r.text or "").strip() or "No pude generar una respuesta. ¿Puedes reformular la pregunta?"
    except Exception as e:
        return error(error_amigable(e), 502, chat=chat_completo(uid, cid))

    voz = generar_audio(uid, respuesta) if f.get("voz") == "1" else None
    ejecutar("""INSERT INTO mensajes(chat_id, rol, texto, imagenes, audio, via_voz, fuentes, fecha)
                VALUES (?, 'model', ?, '[]', ?, 0, ?, ?)""", (cid, respuesta, voz, json.dumps(fuentes), ahora()))
    tocar(cid)
    return jsonify(chat=chat_completo(uid, cid), modelo=modelo)


# ───────────────────────── Arranque ─────────────────────────
if __name__ == "__main__":
    url = f"http://127.0.0.1:{PUERTO}"
    print("\n" + "=" * 56)
    print(f"  FISICA IA corriendo en:  {url}")
    if not clave():
        print("  ⚠ Falta tu API key: pegala en el archivo .env y reinicia.")
    print("  Deja esta ventana abierta. Para apagar: Ctrl + C")
    print("=" * 56 + "\n")
    print("Leyendo la base de conocimiento (carpeta conocimiento/)...")
    BASE.actualizar()
    for a in BASE.resumen():
        print(f"  {'✔' if a['fragmentos'] else '✖'} {a['nombre']}: "
              + (a["aviso"] or f"{a['paginas']} págs/diaps, {a['fragmentos']} fragmentos"))
    if not BASE.resumen():
        print("  (vacía) Pon tus PDF o PowerPoint en la carpeta 'conocimiento' y reinicia.")
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=PUERTO, debug=False)
