# -*- coding: utf-8 -*-
"""
PRIZE - Portal Web/App de Boletas
Integración PRO: portal visual + API de boletas + estructura Render/local.

Listo para GitHub y Render:
- Un solo servidor Flask (no requiere levantar API y portal por separado).
- SQLite local o PostgreSQL en Render si defines DATABASE_URL.
- Login trabajador: DNI + correo.
- Login administrador: usuario + clave.
- Carga masiva de trabajadores por Excel.
- Carga individual/masiva de PDFs por DNI.
- API compatible: /api/health, /api/login, /api/boleta/<dni>, /api/pdf/<dni>.

Credenciales admin demo iniciales:
- admin / admin123
- adm1 / adm1
"""

import os
import re
import sqlite3
from io import BytesIO
from pathlib import Path
from datetime import datetime
from functools import wraps
from zoneinfo import ZoneInfo

try:
    import pandas as pd
except Exception:
    pd = None

try:
    import psycopg2
    import psycopg2.extras
except Exception:
    psycopg2 = None

from flask import (
    Flask, request, redirect, url_for, session, send_file,
    render_template_string, flash, jsonify, abort
)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PERSIST_DIR = os.getenv("PERSIST_DIR", "/data" if os.path.isdir("/data") else BASE_DIR)
UPLOAD_DIR = os.path.join(PERSIST_DIR, "uploads")
PDF_DIR = os.path.join(UPLOAD_DIR, "pdf_boletas")
REPORT_DIR = os.path.join(PERSIST_DIR, "reportes")
DB_PATH = os.path.join(PERSIST_DIR, "boletas_prize.db")
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
USE_POSTGRES = bool(DATABASE_URL)
APP_TZ = ZoneInfo(os.getenv("APP_TIMEZONE", "America/Lima"))
EMPRESA_NOMBRE = os.getenv("EMPRESA_NOMBRE", "PRIZE SUPERFRUITS")

for folder in [PERSIST_DIR, UPLOAD_DIR, PDF_DIR, REPORT_DIR]:
    os.makedirs(folder, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "prize-boletas-web-app-2026")
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024

# =========================================================
# UTILIDADES DB
# =========================================================
def now_app():
    return datetime.now(APP_TZ)


def now_txt():
    return now_app().strftime("%Y-%m-%d %H:%M:%S")


def today_iso():
    return now_app().date().isoformat()


def normalizar_dni(dni):
    solo = "".join(c for c in str(dni or "") if c.isdigit())
    return solo.zfill(8)[-8:] if solo else ""


def clean_text(v):
    return str(v or "").strip()


def get_conn():
    if USE_POSTGRES:
        if psycopg2 is None:
            raise RuntimeError("DATABASE_URL definido, pero psycopg2 no está instalado.")
        return psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def sql_params(sql):
    return sql.replace("?", "%s") if USE_POSTGRES else sql


def q_all(sql, params=()):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(sql_params(sql), params)
        rows = cur.fetchall()
        cur.close()
        return [dict(r) for r in rows]


def q_one(sql, params=()):
    rows = q_all(sql, params)
    return rows[0] if rows else None


def q_exec(sql, params=()):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(sql_params(sql), params)
        conn.commit()
        cur.close()


def q_exec_return_id(sql, params=()):
    with get_conn() as conn:
        cur = conn.cursor()
        if USE_POSTGRES:
            cur.execute(sql_params(sql + " RETURNING id"), params)
            rid = cur.fetchone()["id"]
        else:
            cur.execute(sql, params)
            rid = cur.lastrowid
        conn.commit()
        cur.close()
        return rid


def init_db():
    with get_conn() as conn:
        cur = conn.cursor()
        if USE_POSTGRES:
            cur.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'admin',
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT
            )""")
            cur.execute("""
            CREATE TABLE IF NOT EXISTS trabajadores (
                id SERIAL PRIMARY KEY,
                dni TEXT UNIQUE NOT NULL,
                nombre TEXT NOT NULL,
                correo TEXT,
                cargo TEXT,
                area TEXT,
                empresa TEXT,
                planilla TEXT,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT,
                updated_at TEXT
            )""")
            cur.execute("""
            CREATE TABLE IF NOT EXISTS boletas (
                id SERIAL PRIMARY KEY,
                dni TEXT NOT NULL,
                tipo TEXT NOT NULL DEFAULT 'Utilidad',
                periodo TEXT,
                archivo_nombre TEXT,
                ruta_pdf TEXT,
                fecha_subida TEXT,
                uploaded_by TEXT
            )""")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_boletas_dni ON boletas(dni)")
        else:
            cur.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'admin',
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT
            )""")
            cur.execute("""
            CREATE TABLE IF NOT EXISTS trabajadores (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                dni TEXT UNIQUE NOT NULL,
                nombre TEXT NOT NULL,
                correo TEXT,
                cargo TEXT,
                area TEXT,
                empresa TEXT,
                planilla TEXT,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT,
                updated_at TEXT
            )""")
            cur.execute("""
            CREATE TABLE IF NOT EXISTS boletas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                dni TEXT NOT NULL,
                tipo TEXT NOT NULL DEFAULT 'Utilidad',
                periodo TEXT,
                archivo_nombre TEXT,
                ruta_pdf TEXT,
                fecha_subida TEXT,
                uploaded_by TEXT
            )""")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_boletas_dni ON boletas(dni)")
        conn.commit()
        cur.close()

    for user, pw in [("admin", "admin123"), ("adm1", "adm1")]:
        if not q_one("SELECT id FROM usuarios WHERE username=?", (user,)):
            q_exec(
                "INSERT INTO usuarios(username,password_hash,role,active,created_at) VALUES(?,?,?,?,?)",
                (user, generate_password_hash(pw), "admin", 1, now_txt()),
            )

    demos = [
        ("74324033", "AZABACHE LUJAN, OMAR EDUARDO", "omar@demo.com", "ANALISTA", "RRHH", EMPRESA_NOMBRE, "GENERAL"),
        ("45148597", "CONCEPCION ZAVALETA, VICTOR", "victor@demo.com", "OPERARIO", "PRODUCCION", EMPRESA_NOMBRE, "GENERAL"),
    ]
    for dni, nom, correo, cargo, area, emp, planilla in demos:
        if not q_one("SELECT id FROM trabajadores WHERE dni=?", (dni,)):
            q_exec("""
                INSERT INTO trabajadores(dni,nombre,correo,cargo,area,empresa,planilla,active,created_at,updated_at)
                VALUES(?,?,?,?,?,?,?,?,?,?)
            """, (dni, nom, correo, cargo, area, emp, planilla, 1, now_txt(), now_txt()))


init_db()

# =========================================================
# SEGURIDAD
# =========================================================
def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if session.get("admin_user"):
            return fn(*args, **kwargs)
        flash("Ingresa como administrador.", "error")
        return redirect(url_for("admin_login"))
    return wrapper


def worker_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if session.get("worker_dni"):
            return fn(*args, **kwargs)
        return redirect(url_for("login"))
    return wrapper

# =========================================================
# HELPERS BOLETAS
# =========================================================
def latest_boleta(dni, tipo=None):
    dni = normalizar_dni(dni)
    if tipo:
        return q_one("SELECT * FROM boletas WHERE dni=? AND tipo=? ORDER BY fecha_subida DESC,id DESC", (dni, tipo))
    return q_one("SELECT * FROM boletas WHERE dni=? ORDER BY fecha_subida DESC,id DESC", (dni,))


def trabajador_data(dni):
    return q_one("SELECT * FROM trabajadores WHERE dni=? AND active=1", (normalizar_dni(dni),))


def preparar_resultado(dni):
    t = trabajador_data(dni)
    if not t:
        return None
    b = latest_boleta(dni)
    return {
        "dni": t.get("dni", ""),
        "nombre": t.get("nombre", ""),
        "correo": t.get("correo", ""),
        "cargo": t.get("cargo", ""),
        "area": t.get("area", ""),
        "empresa": t.get("empresa", EMPRESA_NOMBRE) or EMPRESA_NOMBRE,
        "planilla": t.get("planilla", ""),
        "tiene_pdf": bool(b and b.get("ruta_pdf") and os.path.exists(b.get("ruta_pdf"))),
        "url_pdf": b.get("ruta_pdf") if b else "",
        "archivo_nombre": b.get("archivo_nombre") if b else "",
        "tipo": b.get("tipo") if b else "Sin documento",
        "periodo": b.get("periodo") if b else "",
        "fecha_subida": b.get("fecha_subida") if b else "Sin fecha registrada",
        "estado_pdf": "Disponible" if b else "No disponible",
        "ultima_actualizacion": now_app().strftime("%d/%m/%Y %I:%M %p"),
    }


def guardar_pdf(file_storage, tipo="Utilidad", periodo=""):
    filename_raw = secure_filename(file_storage.filename or "")
    if not filename_raw.lower().endswith(".pdf"):
        raise ValueError("Solo se permite PDF.")
    stem = Path(filename_raw).stem
    dni = normalizar_dni(stem)
    if len(dni) != 8:
        m = re.search(r"(\d{8})", filename_raw)
        dni = m.group(1) if m else ""
    if not dni:
        raise ValueError(f"No se detectó DNI en el nombre: {filename_raw}. Usa 12345678.pdf")
    if not trabajador_data(dni):
        raise ValueError(f"DNI {dni} no existe en trabajadores. Carga primero la base de trabajadores.")
    safe_periodo = re.sub(r"[^A-Za-z0-9_\-]", "_", periodo or today_iso())
    tipo_safe = re.sub(r"[^A-Za-z0-9_\-]", "_", tipo or "Utilidad")
    folder = os.path.join(PDF_DIR, tipo_safe, safe_periodo)
    os.makedirs(folder, exist_ok=True)
    final_name = f"{dni}.pdf"
    path = os.path.join(folder, final_name)
    file_storage.save(path)
    q_exec("""
        INSERT INTO boletas(dni,tipo,periodo,archivo_nombre,ruta_pdf,fecha_subida,uploaded_by)
        VALUES(?,?,?,?,?,?,?)
    """, (dni, tipo, periodo, final_name, path, now_txt(), session.get("admin_user", "admin")))
    return dni


def importar_trabajadores_excel(file_storage):
    if pd is None:
        raise RuntimeError("Pandas no está instalado. Revisa requirements.txt.")
    df = pd.read_excel(file_storage)
    df.columns = [str(c).strip().upper() for c in df.columns]
    col_dni = next((c for c in ["DNI", "DOCUMENTO", "NRO_DOCUMENTO", "NUMERO_DOCUMENTO"] if c in df.columns), None)
    col_nombre = next((c for c in ["NOMBRE", "TRABAJADOR", "APELLIDOS Y NOMBRES", "APELLIDOS_NOMBRES"] if c in df.columns), None)
    if not col_dni or not col_nombre:
        raise ValueError("El Excel debe tener columnas DNI y NOMBRE/TRABAJADOR.")
    inserted = updated = skipped = 0
    for _, row in df.iterrows():
        dni = normalizar_dni(row.get(col_dni))
        nombre = clean_text(row.get(col_nombre))
        if not dni or not nombre or dni == "00000000":
            skipped += 1
            continue
        correo = clean_text(row.get("CORREO") or row.get("EMAIL") or "")
        cargo = clean_text(row.get("CARGO") or "")
        area = clean_text(row.get("AREA") or row.get("ÁREA") or "")
        empresa = clean_text(row.get("EMPRESA") or EMPRESA_NOMBRE)
        planilla = clean_text(row.get("PLANILLA") or "")
        exists = q_one("SELECT id FROM trabajadores WHERE dni=?", (dni,))
        if exists:
            q_exec("""
                UPDATE trabajadores SET nombre=?, correo=?, cargo=?, area=?, empresa=?, planilla=?, active=1, updated_at=? WHERE dni=?
            """, (nombre, correo, cargo, area, empresa, planilla, now_txt(), dni))
            updated += 1
        else:
            q_exec("""
                INSERT INTO trabajadores(dni,nombre,correo,cargo,area,empresa,planilla,active,created_at,updated_at)
                VALUES(?,?,?,?,?,?,?,?,?,?)
            """, (dni, nombre, correo, cargo, area, empresa, planilla, 1, now_txt(), now_txt()))
            inserted += 1
    return inserted, updated, skipped

# =========================================================
# UI BASE
# =========================================================
BASE_HTML = r'''
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>{{ title or 'PRIZE Boletas' }}</title>
<style>
:root{--bg:#eef4fb;--nav:#142239;--nav2:#1e314f;--card:#fff;--line:#d8e4f0;--text:#112033;--muted:#66758a;--blue:#1697f6;--blue2:#0f70c5;--yellow:#f6c516;--green:#16a34a;--red:#dc2626;--shadow:0 18px 44px rgba(15,23,42,.12)}
*{box-sizing:border-box} body{margin:0;font-family:Arial,Helvetica,sans-serif;color:var(--text);background:radial-gradient(circle at 12% 10%,rgba(22,151,246,.12),transparent 20%),radial-gradient(circle at 90% 12%,rgba(246,197,22,.14),transparent 18%),linear-gradient(180deg,#eef4fb,#f8fafc)}
a{text-decoration:none;color:inherit}.btn,.btn-blue,.btn-red{border:0;border-radius:14px;padding:11px 14px;font-weight:800;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:7px}.btn{background:#fff;border:1px solid var(--line);color:#17233a}.btn-blue{background:linear-gradient(180deg,var(--blue),var(--blue2));color:#fff;box-shadow:0 12px 22px rgba(15,112,197,.22)}.btn-red{background:#fee2e2;color:#991b1b;border:1px solid #fecaca}.badge{display:inline-flex;border-radius:999px;padding:6px 10px;font-size:12px;font-weight:900}.badge.ok{background:#dcfce7;color:#166534}.badge.no{background:#fee2e2;color:#991b1b}.badge.warn{background:#fef3c7;color:#92400e}
.login-page{min-height:100vh;display:grid;grid-template-columns:330px 1fr}.login-left{background:linear-gradient(180deg,var(--nav),#1c2d49);color:#eaf2fb;padding:30px 24px;display:flex;flex-direction:column;justify-content:space-between}.login-left h1{font-size:30px;line-height:1.1;margin:18px 0 10px}.login-left p{color:#c8d7e6;line-height:1.55}.pill{display:inline-flex;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.08);border-radius:999px;padding:8px 12px;font-size:12px;font-weight:800}.login-main{display:flex;align-items:center;justify-content:center;padding:26px}.login-card{width:100%;max-width:470px;background:rgba(255,255,255,.95);border:1px solid rgba(255,255,255,.8);border-radius:30px;padding:30px;box-shadow:var(--shadow)}.brand{text-align:center;margin-bottom:12px}.brand .prize{font-size:42px;font-weight:900;color:#14304e;letter-spacing:-1px}.brand .super{font-size:13px;font-weight:900;color:#f0b90b;letter-spacing:3px}.login-title{text-align:center;margin:8px 0 5px;font-size:26px}.login-sub{text-align:center;color:var(--muted);margin:0 0 22px}.field{margin-bottom:13px}.field label{font-size:12px;font-weight:900;color:#42526a;display:block;margin:0 0 6px 4px}.field input,.field select{width:100%;border:1px solid var(--line);border-radius:18px;background:#fff;padding:14px 15px;font-size:15px;outline:none;box-shadow:0 8px 20px rgba(15,23,42,.04)}.alert{border-radius:16px;padding:12px 14px;margin-bottom:14px;line-height:1.45;font-size:14px}.alert.error{background:#fff1f2;border:1px solid #fecdd3;color:#be123c}.alert.ok{background:#effcf6;border:1px solid #bbf7d0;color:#166534}
.app-shell{min-height:100vh;display:grid;grid-template-columns:280px 1fr}.sidebar{background:linear-gradient(180deg,var(--nav),#1b2f4d);color:#eaf2fb;padding:22px 16px;position:sticky;top:0;height:100vh}.side-brand{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.08);border-radius:22px;padding:16px;margin-bottom:18px}.side-brand .prize{font-size:32px;font-weight:900}.side-brand .super{color:var(--yellow);font-weight:900;letter-spacing:2px;font-size:11px}.side-nav{display:grid;gap:8px}.side-nav a{padding:12px 14px;border-radius:16px;color:#dce8f4;font-weight:800}.side-nav a:hover,.side-nav a.active{background:rgba(255,255,255,.12);color:#fff}.content{padding:24px;min-width:0}.topbar{display:flex;justify-content:space-between;gap:14px;align-items:center;margin-bottom:18px}.topbar h1{margin:0;font-size:28px}.topbar p{margin:5px 0 0;color:var(--muted)}.grid{display:grid;grid-template-columns:repeat(12,1fr);gap:16px}.card{background:rgba(255,255,255,.96);border:1px solid rgba(216,228,240,.9);border-radius:24px;box-shadow:0 12px 28px rgba(15,23,42,.08);padding:18px;min-width:0}.span-3{grid-column:span 3}.span-4{grid-column:span 4}.span-5{grid-column:span 5}.span-6{grid-column:span 6}.span-7{grid-column:span 7}.span-8{grid-column:span 8}.span-12{grid-column:span 12}.metric{display:flex;align-items:center;justify-content:space-between}.metric h3{margin:0;color:#64748b;font-size:13px}.metric strong{font-size:30px}.metric .ico{width:46px;height:46px;border-radius:16px;background:#eff6ff;display:grid;place-items:center;font-size:22px}.table-wrap{overflow:auto;border-radius:16px;border:1px solid var(--line)}table{width:100%;border-collapse:collapse;background:#fff}th,td{padding:11px 12px;border-bottom:1px solid #edf2f7;text-align:left;font-size:13px;vertical-align:middle}th{background:#f8fafc;color:#475569;font-size:12px;text-transform:uppercase}.info{display:grid;grid-template-columns:150px 1fr;gap:8px 12px}.info .label{color:#64748b;font-weight:900}.info .value{font-weight:700;word-break:break-word}.tiles{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.tile{background:#f8fbfe;border:1px solid var(--line);border-radius:18px;padding:15px}.tile h4{margin:0 0 7px}.tile p{margin:0;color:#64748b;line-height:1.45;font-size:13px}.actions{display:flex;gap:10px;flex-wrap:wrap}.form-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;align-items:end}.section-title{font-size:18px;margin:0 0 12px}.mobile-tabs{display:none;position:sticky;bottom:0;background:rgba(20,34,57,.96);padding:8px;z-index:10;gap:6px;overflow-x:auto}.mobile-tabs a{white-space:nowrap;color:#eaf2fb;background:rgba(255,255,255,.08);border-radius:999px;padding:10px 12px;font-size:12px;font-weight:900}
@media(max-width:980px){.login-page{grid-template-columns:1fr}.login-left{display:none}.app-shell{display:block}.sidebar{display:none}.content{padding:16px 14px 82px}.grid{grid-template-columns:1fr}.span-3,.span-4,.span-5,.span-6,.span-7,.span-8,.span-12{grid-column:span 1}.topbar{align-items:flex-start;flex-direction:column}.tiles{grid-template-columns:1fr}.form-grid{grid-template-columns:1fr}.info{grid-template-columns:1fr}.mobile-tabs{display:flex}.card{border-radius:20px}.login-main{padding:16px}.login-card{border-radius:24px;padding:24px 18px}}
</style>
</head>
<body>
{% with messages = get_flashed_messages(with_categories=true) %}{% if messages %}<div style="position:fixed;right:16px;top:16px;z-index:99;max-width:390px">{% for cat,msg in messages %}<div class="alert {{ 'ok' if cat=='ok' else 'error' }}">{{ msg }}</div>{% endfor %}</div>{% endif %}{% endwith %}
{{ content|safe }}
</body>
</html>
'''


def render_page(content, title="PRIZE Boletas"):
    return render_template_string(BASE_HTML, content=content, title=title)


def layout(content, active="dashboard", title="Panel"):
    nav = [
        ("dashboard", "🏠 Dashboard", url_for("admin_dashboard") if session.get("admin_user") else url_for("panel")),
        ("trabajadores", "👥 Trabajadores", url_for("trabajadores")),
        ("boletas", "📄 Boletas", url_for("boletas_admin")),
        ("api", "🔌 API", url_for("api_status")),
        ("logout", "🚪 Salir", url_for("logout")),
    ]
    if not session.get("admin_user"):
        nav = [("panel", "📄 Mis boletas", url_for("panel")), ("logout", "🚪 Salir", url_for("logout"))]
    links = "".join([f'<a class="{ "active" if k==active else "" }" href="{u}">{label}</a>' for k,label,u in nav])
    mobile = "".join([f'<a href="{u}">{label}</a>' for k,label,u in nav])
    shell = f'''
    <div class="app-shell">
      <aside class="sidebar">
        <div class="side-brand"><div class="prize">Prize</div><div class="super">SUPERFRUITS</div><p style="color:#c8d7e6;margin:8px 0 0;font-size:13px">Portal web/app de boletas</p></div>
        <nav class="side-nav">{links}</nav>
      </aside>
      <main class="content">{content}</main>
      <nav class="mobile-tabs">{mobile}</nav>
    </div>'''
    return render_page(shell, title)

# =========================================================
# LOGIN TRABAJADOR Y PANEL
# =========================================================
@app.route("/", methods=["GET", "POST"])
def login():
    if session.get("worker_dni"):
        return redirect(url_for("panel"))
    if session.get("admin_user"):
        return redirect(url_for("admin_dashboard"))
    if request.method == "POST":
        dni = normalizar_dni(request.form.get("dni"))
        correo = clean_text(request.form.get("correo")).lower()
        t = trabajador_data(dni)
        if not t:
            flash("DNI no encontrado. Verifica que esté cargado en la base.", "error")
        elif clean_text(t.get("correo")).lower() != correo:
            flash("El correo no coincide con el registrado.", "error")
        else:
            session.clear()
            session["worker_dni"] = dni
            return redirect(url_for("panel"))
    content = '''
    <div class="login-page">
      <aside class="login-left"><div><span class="pill">Portal corporativo · estilo app</span><h1>Acceso moderno para boletas y documentos</h1><p>Integrado en un solo sistema web responsive, listo para celular, GitHub y Render.</p></div><p>Admin demo: <b>admin / admin123</b><br>Trabajador demo: <b>74324033 / omar@demo.com</b></p></aside>
      <main class="login-main"><form class="login-card" method="post"><div class="brand"><div class="prize">Prize</div><div class="super">SUPERFRUITS</div></div><h2 class="login-title">Portal de Boletas</h2><p class="login-sub">Ingresa con DNI y correo registrado.</p><div class="field"><label>DNI</label><input name="dni" inputmode="numeric" maxlength="8" placeholder="Ej. 74324033" required></div><div class="field"><label>Correo</label><input name="correo" type="email" placeholder="correo@empresa.com" required></div><button class="btn-blue" style="width:100%">Ingresar</button><div class="actions" style="margin-top:14px;justify-content:center"><a class="btn" href="/admin/login">Soy administrador</a></div></form></main>
    </div>'''
    return render_page(content, "Login trabajador")


@app.route("/panel")
@worker_required
def panel():
    r = preparar_resultado(session["worker_dni"])
    if not r:
        session.clear()
        return redirect(url_for("login"))
    boletas = q_all("SELECT * FROM boletas WHERE dni=? ORDER BY fecha_subida DESC,id DESC", (r["dni"],))
    rows = "".join([f"<tr><td>{b['tipo']}</td><td>{b.get('periodo') or ''}</td><td>{b['fecha_subida']}</td><td><a class='btn-blue' href='/ver_pdf/{b['id']}' target='_blank'>Ver PDF</a></td></tr>" for b in boletas]) or "<tr><td colspan='4'>Aún no hay documentos cargados.</td></tr>"
    content = f'''
    <div class="topbar"><div><h1>Hola, {r['nombre']}</h1><p>DNI {r['dni']} · {r['empresa']} · última actualización {r['ultima_actualizacion']}</p></div><div class="actions"><a class="btn" href="/panel">Actualizar</a><a class="btn-red" href="/logout">Salir</a></div></div>
    <div class="grid">
      <section class="card span-4"><div class="metric"><div><h3>Estado PDF</h3><strong>{'OK' if r['tiene_pdf'] else 'NO'}</strong></div><div class="ico">📄</div></div><p>{'<span class="badge ok">Disponible</span>' if r['tiene_pdf'] else '<span class="badge no">No disponible</span>'}</p></section>
      <section class="card span-4"><div class="metric"><div><h3>Documentos</h3><strong>{len(boletas)}</strong></div><div class="ico">🗂️</div></div></section>
      <section class="card span-4"><div class="metric"><div><h3>Portal</h3><strong>APP</strong></div><div class="ico">📱</div></div></section>
      <section class="card span-6"><h3 class="section-title">Datos del colaborador</h3><div class="info"><div class="label">DNI</div><div class="value">{r['dni']}</div><div class="label">Nombre</div><div class="value">{r['nombre']}</div><div class="label">Correo</div><div class="value">{r['correo']}</div><div class="label">Cargo</div><div class="value">{r['cargo']}</div><div class="label">Área</div><div class="value">{r['area']}</div><div class="label">Empresa</div><div class="value">{r['empresa']}</div></div></section>
      <section class="card span-6"><h3 class="section-title">Accesos rápidos</h3><div class="tiles"><div class="tile"><h4>Boleta</h4><p>Documento principal del trabajador.</p>{'<a class="btn-blue" href="/pdf/'+r['dni']+'" target="_blank">Abrir PDF</a>' if r['tiene_pdf'] else '<span class="badge no">Sin PDF</span>'}</div><div class="tile"><h4>Responsive</h4><p>Diseño listo para PC y celular.</p></div></div></section>
      <section class="card span-12"><h3 class="section-title">Historial de documentos</h3><div class="table-wrap"><table><thead><tr><th>Tipo</th><th>Periodo</th><th>Subida</th><th>Acción</th></tr></thead><tbody>{rows}</tbody></table></div></section>
    </div>'''
    return layout(content, "panel", "Mi panel")

# =========================================================
# ADMIN
# =========================================================
@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        username = clean_text(request.form.get("username"))
        password = request.form.get("password", "")
        u = q_one("SELECT * FROM usuarios WHERE username=? AND active=1", (username,))
        if u and check_password_hash(u["password_hash"], password):
            session.clear()
            session["admin_user"] = username
            session["role"] = u.get("role", "admin")
            return redirect(url_for("admin_dashboard"))
        flash("Usuario o clave incorrecta.", "error")
    content = '''
    <div class="login-page"><aside class="login-left"><div><span class="pill">Administrador</span><h1>Gestión de boletas PRIZE</h1><p>Carga trabajadores, PDFs y revisa la API desde un solo panel.</p></div></aside><main class="login-main"><form class="login-card" method="post"><div class="brand"><div class="prize">Prize</div><div class="super">ADMIN</div></div><h2 class="login-title">Acceso administrador</h2><div class="field"><label>Usuario</label><input name="username" required></div><div class="field"><label>Clave</label><input name="password" type="password" required></div><button class="btn-blue" style="width:100%">Ingresar</button><div class="actions" style="margin-top:14px;justify-content:center"><a class="btn" href="/">Portal trabajador</a></div></form></main></div>'''
    return render_page(content, "Login admin")


@app.route("/admin")
@admin_required
def admin_dashboard():
    total_t = q_one("SELECT COUNT(*) c FROM trabajadores")["c"]
    total_b = q_one("SELECT COUNT(*) c FROM boletas")["c"]
    con_pdf = q_one("SELECT COUNT(DISTINCT dni) c FROM boletas")["c"]
    ultimos = q_all("""
        SELECT b.id,b.dni,b.tipo,b.periodo,b.fecha_subida,t.nombre,t.empresa
        FROM boletas b LEFT JOIN trabajadores t ON t.dni=b.dni
        ORDER BY b.fecha_subida DESC,b.id DESC LIMIT 10
    """)
    rows = "".join([f"<tr><td>{r['dni']}</td><td>{r.get('nombre') or ''}</td><td>{r['tipo']}</td><td>{r.get('periodo') or ''}</td><td>{r['fecha_subida']}</td><td><a class='btn' href='/ver_pdf/{r['id']}' target='_blank'>Ver</a></td></tr>" for r in ultimos]) or "<tr><td colspan='6'>Sin boletas cargadas.</td></tr>"
    content = f'''
    <div class="topbar"><div><h1>Dashboard de Boletas</h1><p>Panel tipo web/app integrado para GitHub y Render.</p></div><div class="actions"><a class="btn" href="/admin">Actualizar</a><a class="btn-red" href="/logout">Salir</a></div></div>
    <div class="grid">
      <section class="card span-3"><div class="metric"><div><h3>Trabajadores</h3><strong>{total_t}</strong></div><div class="ico">👥</div></div></section>
      <section class="card span-3"><div class="metric"><div><h3>Boletas</h3><strong>{total_b}</strong></div><div class="ico">📄</div></div></section>
      <section class="card span-3"><div class="metric"><div><h3>Con PDF</h3><strong>{con_pdf}</strong></div><div class="ico">✅</div></div></section>
      <section class="card span-3"><div class="metric"><div><h3>Modo</h3><strong>{'PG' if USE_POSTGRES else 'SQL'}</strong></div><div class="ico">☁️</div></div></section>
      <section class="card span-12"><h3 class="section-title">Últimas boletas cargadas</h3><div class="table-wrap"><table><thead><tr><th>DNI</th><th>Trabajador</th><th>Tipo</th><th>Periodo</th><th>Fecha</th><th>PDF</th></tr></thead><tbody>{rows}</tbody></table></div></section>
    </div>'''
    return layout(content, "dashboard", "Admin")


@app.route("/admin/trabajadores", methods=["GET", "POST"])
@admin_required
def trabajadores():
    if request.method == "POST":
        file = request.files.get("excel")
        if not file or not file.filename:
            flash("Selecciona un Excel.", "error")
        else:
            try:
                ins, upd, skip = importar_trabajadores_excel(file)
                flash(f"Carga terminada: insertados {ins}, actualizados {upd}, omitidos {skip}.", "ok")
            except Exception as e:
                flash(str(e), "error")
        return redirect(url_for("trabajadores"))
    buscar = clean_text(request.args.get("buscar"))
    params = []
    where = "WHERE 1=1"
    if buscar:
        where += " AND (dni LIKE ? OR nombre LIKE ? OR correo LIKE ? OR empresa LIKE ?)"
        b = f"%{buscar}%"; params = [b,b,b,b]
    rows_data = q_all(f"SELECT * FROM trabajadores {where} ORDER BY updated_at DESC,id DESC LIMIT 300", tuple(params))
    rows = "".join([f"<tr><td>{r['dni']}</td><td>{r['nombre']}</td><td>{r.get('correo') or ''}</td><td>{r.get('cargo') or ''}</td><td>{r.get('area') or ''}</td><td>{r.get('empresa') or ''}</td></tr>" for r in rows_data]) or "<tr><td colspan='6'>Sin registros.</td></tr>"
    content = f'''
    <div class="topbar"><div><h1>Trabajadores</h1><p>Carga Excel con columnas DNI, NOMBRE/TRABAJADOR, CORREO, CARGO, AREA, EMPRESA.</p></div></div>
    <div class="grid">
      <section class="card span-12"><form method="post" enctype="multipart/form-data" class="form-grid"><div class="field"><label>Excel trabajadores</label><input type="file" name="excel" accept=".xlsx,.xls" required></div><button class="btn-blue">Importar / reemplazar</button><a class="btn" href="/admin/trabajadores">Actualizar</a></form></section>
      <section class="card span-12"><form method="get" class="form-grid"><div class="field"><label>Buscar</label><input name="buscar" value="{buscar}" placeholder="DNI, nombre, correo, empresa"></div><button class="btn-blue">Filtrar</button></form></section>
      <section class="card span-12"><div class="table-wrap"><table><thead><tr><th>DNI</th><th>Nombre</th><th>Correo</th><th>Cargo</th><th>Área</th><th>Empresa</th></tr></thead><tbody>{rows}</tbody></table></div></section>
    </div>'''
    return layout(content, "trabajadores", "Trabajadores")


@app.route("/admin/boletas", methods=["GET", "POST"])
@admin_required
def boletas_admin():
    if request.method == "POST":
        tipo = clean_text(request.form.get("tipo")) or "Utilidad"
        periodo = clean_text(request.form.get("periodo")) or today_iso()
        files = request.files.getlist("pdfs")
        ok = errores = 0
        msgs = []
        for f in files:
            if not f or not f.filename:
                continue
            try:
                guardar_pdf(f, tipo, periodo)
                ok += 1
            except Exception as e:
                errores += 1
                msgs.append(str(e))
        if ok:
            flash(f"PDFs cargados correctamente: {ok}.", "ok")
        if errores:
            flash("Errores: " + " | ".join(msgs[:5]), "error")
        return redirect(url_for("boletas_admin"))
    buscar = clean_text(request.args.get("buscar"))
    params = []
    where = "WHERE 1=1"
    if buscar:
        where += " AND (b.dni LIKE ? OR t.nombre LIKE ? OR b.tipo LIKE ? OR b.periodo LIKE ?)"
        b = f"%{buscar}%"; params = [b,b,b,b]
    rows_data = q_all(f"""
        SELECT b.*,t.nombre,t.correo,t.empresa FROM boletas b LEFT JOIN trabajadores t ON t.dni=b.dni
        {where} ORDER BY b.fecha_subida DESC,b.id DESC LIMIT 300
    """, tuple(params))
    rows = "".join([f"<tr><td>{r['dni']}</td><td>{r.get('nombre') or ''}</td><td>{r['tipo']}</td><td>{r.get('periodo') or ''}</td><td>{r['fecha_subida']}</td><td><a class='btn-blue' href='/ver_pdf/{r['id']}' target='_blank'>Ver</a></td></tr>" for r in rows_data]) or "<tr><td colspan='6'>Sin PDFs cargados.</td></tr>"
    content = f'''
    <div class="topbar"><div><h1>Boletas / Documentos</h1><p>Carga PDFs nombrados con DNI, ejemplo: 74324033.pdf. También detecta DNI dentro del nombre.</p></div></div>
    <div class="grid">
      <section class="card span-12"><form method="post" enctype="multipart/form-data" class="form-grid"><div class="field"><label>Tipo</label><select name="tipo"><option>Utilidad</option><option>Normal</option><option>CTS</option><option>Gratificación</option><option>Vacaciones</option><option>Liquidación</option><option>Contrato</option><option>Otros</option></select></div><div class="field"><label>Periodo</label><input name="periodo" value="{today_iso()}" placeholder="01_2026 / 2026 / Semana 01"></div><div class="field"><label>PDFs</label><input type="file" name="pdfs" accept=".pdf" multiple required></div><button class="btn-blue">Subir PDFs</button></form></section>
      <section class="card span-12"><form method="get" class="form-grid"><div class="field"><label>Buscar</label><input name="buscar" value="{buscar}" placeholder="DNI, trabajador, tipo, periodo"></div><button class="btn-blue">Filtrar</button><a class="btn" href="/admin/boletas">Actualizar</a></form></section>
      <section class="card span-12"><div class="table-wrap"><table><thead><tr><th>DNI</th><th>Trabajador</th><th>Tipo</th><th>Periodo</th><th>Subida</th><th>PDF</th></tr></thead><tbody>{rows}</tbody></table></div></section>
    </div>'''
    return layout(content, "boletas", "Boletas")


@app.route("/admin/api")
@admin_required
def api_status():
    health = {"ok": True, "db": "PostgreSQL" if USE_POSTGRES else "SQLite", "database": "DATABASE_URL" if USE_POSTGRES else DB_PATH, "pdf_dir": PDF_DIR}
    content = f'''
    <div class="topbar"><div><h1>Estado API</h1><p>Endpoints integrados en el mismo app.py.</p></div></div>
    <div class="grid"><section class="card span-12"><pre style="white-space:pre-wrap;background:#0f172a;color:#e2e8f0;border-radius:18px;padding:16px">{health}</pre><div class="tiles"><div class="tile"><h4>/api/health</h4><p>Estado del sistema.</p></div><div class="tile"><h4>/api/boleta/&lt;dni&gt;</h4><p>Consulta de trabajador y último PDF.</p></div><div class="tile"><h4>/api/login</h4><p>Validación DNI + correo por POST JSON.</p></div><div class="tile"><h4>/api/pdf/&lt;dni&gt;</h4><p>Abre PDF por DNI.</p></div></div></section></div>'''
    return layout(content, "api", "API")

# =========================================================
# PDF Y API
# =========================================================
@app.route("/ver_pdf/<int:boleta_id>")
def ver_pdf(boleta_id):
    b = q_one("SELECT * FROM boletas WHERE id=?", (boleta_id,))
    if not b or not b.get("ruta_pdf") or not os.path.exists(b["ruta_pdf"]):
        abort(404)
    if session.get("admin_user") or session.get("worker_dni") == b["dni"]:
        return send_file(b["ruta_pdf"], mimetype="application/pdf", as_attachment=False, download_name=b.get("archivo_nombre") or os.path.basename(b["ruta_pdf"]))
    abort(403)


@app.route("/pdf/<dni>")
@worker_required
def pdf_worker(dni):
    if session.get("worker_dni") != normalizar_dni(dni):
        abort(403)
    b = latest_boleta(dni)
    if not b:
        abort(404)
    return redirect(url_for("ver_pdf", boleta_id=b["id"]))


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.after_request
def add_headers(resp):
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.errorhandler(404)
def not_found(e):
    return render_page("<div class='login-main'><div class='login-card'><h2>No encontrado</h2><p>La ruta o documento no existe.</p><a class='btn-blue' href='/'>Volver</a></div></div>", "404"), 404


@app.errorhandler(500)
def server_error(e):
    app.logger.exception("Error interno: %s", e)
    return render_page("<div class='login-main'><div class='login-card'><h2>Error interno controlado</h2><p>Revisa los logs de Render o vuelve a ingresar.</p><a class='btn-blue' href='/logout'>Limpiar sesión</a></div></div>", "Error"), 500


@app.route("/api/health")
def api_health():
    try:
        q_one("SELECT COUNT(*) c FROM trabajadores")
        return jsonify({"ok": True, "mensaje": "API y BD operativas", "db": "postgres" if USE_POSTGRES else "sqlite"})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/boleta/<dni>")
def api_boleta(dni):
    r = preparar_resultado(dni)
    if not r:
        return jsonify({"ok": False, "mensaje": "DNI no encontrado"}), 404
    out = dict(r)
    out["ok"] = True
    out["pdf_url"] = url_for("api_pdf", dni=r["dni"], _external=True) if r["tiene_pdf"] else ""
    return jsonify(out)


@app.route("/api/login", methods=["POST"])
def api_login():
    data = request.get_json(silent=True) or {}
    dni = normalizar_dni(data.get("dni"))
    correo = clean_text(data.get("correo")).lower()
    t = trabajador_data(dni)
    if not t:
        return jsonify({"ok": False, "mensaje": "DNI no encontrado"}), 404
    if clean_text(t.get("correo")).lower() != correo:
        return jsonify({"ok": False, "mensaje": "Correo no coincide"}), 401
    r = preparar_resultado(dni)
    r["ok"] = True
    return jsonify(r)


@app.route("/api/pdf/<dni>")
def api_pdf(dni):
    b = latest_boleta(dni)
    if not b or not b.get("ruta_pdf") or not os.path.exists(b["ruta_pdf"]):
        return jsonify({"ok": False, "mensaje": "PDF no disponible"}), 404
    return send_file(b["ruta_pdf"], mimetype="application/pdf", as_attachment=False, download_name=b.get("archivo_nombre") or os.path.basename(b["ruta_pdf"]))


if __name__ == "__main__":
    port = int(os.getenv("PORT", "8000"))
    app.run(host="0.0.0.0", port=port, debug=False)
