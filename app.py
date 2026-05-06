# -*- coding: utf-8 -*-
"""
Portal de Documentos PRIZE - Versión Ultra Mejorada
Listo para Render / GitHub / uso local.

Usuarios demo:
- Administrador: admin / admin123
- Trabajador: DNI 74324033 / correo omar@demo.com

Variables Render opcionales:
- SECRET_KEY
- PERSIST_DIR=/data
- APP_TIMEZONE=America/Lima
"""

import os
import re
import sqlite3
from datetime import datetime
from pathlib import Path
from functools import wraps
from zoneinfo import ZoneInfo

from flask import Flask, request, redirect, url_for, session, send_file, render_template_string, flash, jsonify, abort
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
PERSIST_DIR = Path(os.getenv("PERSIST_DIR", "/data" if Path("/data").is_dir() else str(BASE_DIR)))
STATIC_DIR = BASE_DIR / "static"
UPLOAD_DIR = PERSIST_DIR / "uploads"
DB_PATH = PERSIST_DIR / "boletas_prize.db"
APP_TZ = ZoneInfo(os.getenv("APP_TIMEZONE", "America/Lima"))

for d in (PERSIST_DIR, STATIC_DIR, UPLOAD_DIR):
    d.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "prize_documentos_ultra_2026")
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

# =============================
# CONFIGURACIÓN FUNCIONAL
# =============================
TIPOS_PAGO = [
    ("Utilidad", "Boletas utilidades", "📄"),
    ("Vacaciones", "Boletas vacaciones", "📄"),
    ("Normal", "Boletas normal", "📄"),
    ("Constancia Gratificación", "Constancia gratificación", "📄"),
    ("CTS", "Boletas CTS", "📄"),
    ("Liquidación", "Boletas liquidación", "📄"),
    ("Gratificación", "Boletas gratificación", "📄"),
]
TIPOS_EMPRESA = [
    ("Contrato de Trabajo", "Contrato de Trabajo", "📑"),
    ("Reglamento Interno", "Reglamento Interno", "📘"),
    ("Reglamento de SST", "Reglamento de SST", "🦺"),
    ("Código de Conducta", "Código de Conducta", "⚖️"),
    ("Políticas", "Políticas", "📌"),
    ("Comunicados", "Comunicados", "📣"),
    ("Formatos", "Formatos", "🧾"),
]
TIPOS_PERSONALES = [
    ("Otros", "Otros documentos", "🗂️"),
    ("Contrato Personal", "Contrato de Trabajo", "📑"),
]
ALL_TIPOS = {k: (label, icon, "pago") for k, label, icon in TIPOS_PAGO}
ALL_TIPOS.update({k: (label, icon, "empresa") for k, label, icon in TIPOS_EMPRESA})
ALL_TIPOS.update({k: (label, icon, "personal") for k, label, icon in TIPOS_PERSONALES})
EXT_ALLOWED = {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".doc", ".docx", ".xls", ".xlsx"}


def now_txt():
    return datetime.now(APP_TZ).strftime("%d/%m/%Y %I:%M %p")


def now_file():
    return datetime.now(APP_TZ).strftime("%Y%m%d_%H%M%S")


def clean(v):
    return str(v or "").strip()


def normalizar_dni(v):
    d = re.sub(r"\D", "", str(v or ""))
    return d[-8:].zfill(8) if d else ""


def safe_periodo(p):
    return re.sub(r"[^A-Za-z0-9_\- ]", "", clean(p))[:50] or datetime.now(APP_TZ).strftime("%Y-%m")


def logo_url():
    # Reconoce logo en la carpeta raíz o static: logo_prize.png, logo.png, prize.png, etc.
    nombres = ["logo_prize.png", "logo.png", "prize.png", "LOGO.png", "Logo.png", "logo_prize.jpg", "logo_prize.jpeg"]
    for folder in (BASE_DIR, STATIC_DIR):
        for name in nombres:
            p = folder / name
            if p.exists():
                if folder == STATIC_DIR:
                    return url_for("static", filename=name)
                return url_for("logo_file", filename=name)
    return url_for("logo_svg")


def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with db() as con:
        con.execute("""
        CREATE TABLE IF NOT EXISTS usuarios_admin(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT UNIQUE,
            clave_hash TEXT,
            nombre TEXT,
            rol TEXT DEFAULT 'admin',
            activo INTEGER DEFAULT 1
        )""")
        con.execute("""
        CREATE TABLE IF NOT EXISTS trabajadores(
            dni TEXT PRIMARY KEY,
            nombre TEXT,
            correo TEXT,
            cargo TEXT,
            area TEXT,
            empresa TEXT,
            activo INTEGER DEFAULT 1,
            fecha_registro TEXT
        )""")
        con.execute("""
        CREATE TABLE IF NOT EXISTS documentos(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            dni TEXT,
            categoria TEXT,
            tipo TEXT,
            periodo TEXT,
            detalle TEXT,
            observacion TEXT,
            archivo_nombre TEXT,
            ruta_archivo TEXT,
            extension TEXT,
            fecha_subida TEXT,
            uploaded_by TEXT
        )""")
        # Datos demo seguros
        if not con.execute("SELECT 1 FROM usuarios_admin WHERE usuario='admin'").fetchone():
            con.execute("INSERT INTO usuarios_admin(usuario,clave_hash,nombre,rol) VALUES(?,?,?,?)",
                        ("admin", generate_password_hash("admin123"), "Administrador PRIZE", "admin"))
        if not con.execute("SELECT 1 FROM trabajadores WHERE dni='74324033'").fetchone():
            con.execute("INSERT INTO trabajadores(dni,nombre,correo,cargo,area,empresa,activo,fecha_registro) VALUES(?,?,?,?,?,?,?,?)",
                        ("74324033", "AZABACHE LUJAN, OMAR EDUARDO", "omar@demo.com", "Analista", "RR.HH.", "PRIZE SUPERFRUITS", 1, now_txt()))
        con.commit()


init_db()

# =============================
# SEGURIDAD / DECORADORES
# =============================
def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("admin_id"):
            return redirect(url_for("admin_login"))
        return fn(*args, **kwargs)
    return wrapper


def worker_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("dni"):
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper

# =============================
# DB HELPERS
# =============================
def get_trabajador(dni):
    dni = normalizar_dni(dni)
    with db() as con:
        return con.execute("SELECT * FROM trabajadores WHERE dni=?", (dni,)).fetchone()


def listar_documentos(dni=None, tipo=None, categoria=None, periodo=None, buscar=None, limit=300):
    where, params = [], []
    if dni:
        where.append("(dni=? OR categoria='empresa')")
        params.append(normalizar_dni(dni))
    if tipo:
        where.append("tipo=?")
        params.append(tipo)
    if categoria:
        where.append("categoria=?")
        params.append(categoria)
    if periodo:
        where.append("periodo=?")
        params.append(periodo)
    if buscar:
        b = f"%{buscar}%"
        where.append("(dni LIKE ? OR tipo LIKE ? OR periodo LIKE ? OR detalle LIKE ? OR observacion LIKE ? OR archivo_nombre LIKE ?)")
        params += [b, b, b, b, b, b]
    sql = "SELECT * FROM documentos"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY id DESC LIMIT ?"
    params.append(limit)
    with db() as con:
        return con.execute(sql, params).fetchall()


def periodos_disponibles(dni=None, tipo=None, categoria=None):
    where, params = ["periodo IS NOT NULL", "periodo<>''"], []
    if dni:
        where.append("(dni=? OR categoria='empresa')"); params.append(normalizar_dni(dni))
    if tipo:
        where.append("tipo=?"); params.append(tipo)
    if categoria:
        where.append("categoria=?"); params.append(categoria)
    sql = "SELECT DISTINCT periodo FROM documentos WHERE " + " AND ".join(where) + " ORDER BY periodo DESC LIMIT 80"
    with db() as con:
        return [r[0] for r in con.execute(sql, params).fetchall() if r[0]]


def guardar_documento(file_storage, dni, tipo, periodo, detalle="", observacion="", uploaded_by="sistema"):
    if not file_storage or not file_storage.filename:
        return None
    original = secure_filename(file_storage.filename)
    ext = Path(original).suffix.lower()
    if ext not in EXT_ALLOWED:
        raise ValueError(f"Extensión no permitida: {ext}")
    tipo_info = ALL_TIPOS.get(tipo, (tipo, "📄", "personal"))
    categoria = tipo_info[2]
    dni = normalizar_dni(dni) if categoria != "empresa" else ""
    periodo = safe_periodo(periodo)
    folder = UPLOAD_DIR / categoria / re.sub(r"[^A-Za-z0-9_\-]", "_", tipo) / periodo
    if dni:
        folder = folder / dni
    folder.mkdir(parents=True, exist_ok=True)
    tipo_file = re.sub(r"[^A-Za-z0-9_\-]+", "_", tipo)
    prefijo_dni = f"{dni}_" if dni else ""
    final = f"{prefijo_dni}{tipo_file}_{periodo}_{now_file()}_{original}"
    path = folder / final
    file_storage.save(path)
    with db() as con:
        con.execute("""
        INSERT INTO documentos(dni,categoria,tipo,periodo,detalle,observacion,archivo_nombre,ruta_archivo,extension,fecha_subida,uploaded_by)
        VALUES(?,?,?,?,?,?,?,?,?,?,?)
        """, (dni, categoria, tipo, periodo, clean(detalle), clean(observacion), original, str(path), ext, now_txt(), uploaded_by))
        con.commit()
    return str(path)

# =============================
# ESTILOS Y LAYOUT
# =============================
BASE = r'''
<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{{ title }}</title>
<style>
:root{
 --ink:#071426;--nav:#0b1423;--nav2:#12233a;--nav3:#294663;--nav4:#365675;
 --line:#d8e8f3;--txt:#06162c;--mut:#60748b;--green:#18a64a;--green2:#29cc66;
 --blue:#247cad;--sky:#8fd2ed;--aqua:#bdeee7;--mint:#dff8e9;--orange:#f28b1a;
 --card:#ffffff;--soft:#f7fbfd;--shadow:0 22px 55px rgba(15,23,42,.13)
}
*{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;font-family:Inter,Segoe UI,Arial,sans-serif;color:var(--txt);background:radial-gradient(circle at -8% 82%,rgba(46,135,192,.26),transparent 30%),radial-gradient(circle at 108% 76%,rgba(28,179,86,.25),transparent 34%),linear-gradient(135deg,#fbfdff 0%,#eef9f5 100%);font-weight:650}a{text-decoration:none;color:inherit}.hidden{display:none!important}
.btn,.btn-blue,.btn-green,.btn-red{border:1px solid var(--line);border-radius:16px;padding:12px 18px;background:#fff;color:#071426;font-weight:950;cursor:pointer;display:inline-flex;align-items:center;gap:8px;box-shadow:0 8px 18px rgba(15,23,42,.05);transition:.16s}.btn:hover,.btn-blue:hover,.btn-green:hover{transform:translateY(-1px);box-shadow:0 14px 28px rgba(15,23,42,.10)}.btn-blue{background:linear-gradient(135deg,#eff9ff,#fff);border-color:#b9dff1}.btn-green{background:linear-gradient(135deg,var(--green),var(--green2));border:0;color:white}.btn-red{background:#fff1f2;border-color:#fecdd3;color:#be123c}.pill{display:inline-flex;align-items:center;gap:6px;border-radius:999px;padding:8px 12px;background:#f1f8fb;border:1px solid #dbeaf3;font-size:13px;color:#39506b}.flash{padding:13px 16px;border-radius:16px;margin:10px 0;background:#ecfdf5;border:1px solid #bbf7d0;color:#166534}.flash.err{background:#fff1f2;border-color:#fecdd3;color:#be123c}
/* LOGIN - colores PRIZE oscuros con verde/azul */
.login-body{min-height:100vh;display:grid;place-items:center;padding:20px;background:radial-gradient(circle at 8% 85%,rgba(43,132,190,.34),transparent 23%),radial-gradient(circle at 90% 82%,rgba(42,204,102,.30),transparent 24%),linear-gradient(125deg,#f9fbfd 0%,#f4fbf6 100%)}.login-card{width:min(92vw,540px);background:linear-gradient(180deg,#101a2a,#0a1321 78%);border:1px solid rgba(255,255,255,.10);border-radius:30px;padding:34px 38px 0;box-shadow:0 42px 95px rgba(15,23,42,.33);overflow:hidden;position:relative}.login-card:before{content:"";position:absolute;left:-58px;bottom:-86px;width:340px;height:180px;background:linear-gradient(135deg,#166da7,#1f8ec6);border-radius:50% 50% 0 0;transform:rotate(4deg)}.login-card:after{content:"";position:absolute;right:-55px;bottom:-78px;width:330px;height:180px;background:linear-gradient(135deg,#169545,#23c85e);border-radius:50% 50% 0 0;transform:rotate(-5deg)}.login-inner{position:relative;z-index:2}.login-logo{text-align:center}.login-logo img{max-width:245px;max-height:115px;object-fit:contain;background:#fff;border-radius:4px;padding:4px;filter:drop-shadow(0 10px 18px rgba(0,0,0,.28))}.login-title{text-align:center;color:#b7c9dc;margin:18px 0 28px}.login-title h1{margin:0 0 7px;color:#dff1ff;font-size:25px;letter-spacing:.2px}.field{display:grid;gap:7px}.field label{font-size:13px;color:#29425f;font-weight:950}.login-card .field label{color:#83b6dc}.input,select,textarea{width:100%;border:1px solid #d7e4ef;border-radius:14px;padding:13px 14px;background:#fff;color:#071426;font:inherit;outline:none}.input:focus,select:focus,textarea:focus{border-color:#7fc9e7;box-shadow:0 0 0 4px rgba(127,201,231,.18)}.login-input{display:flex;align-items:center;gap:10px;background:#fff;border-radius:15px;padding:0 13px;margin-bottom:18px;border:1px solid #e6eef6}.login-input input{border:0;padding:16px 8px;width:100%;font:inherit;outline:none}.login-card .btn-green{width:100%;justify-content:center;font-size:18px;margin:6px 0 68px}.login-links{text-align:center;margin-top:-48px;padding-bottom:22px;position:relative;z-index:3}.login-links a{color:#bfdbfe;font-size:13px}
/* APP */
.app{display:grid;grid-template-columns:320px 1fr;min-height:100vh}.side{background:linear-gradient(180deg,#091422,#14243a 68%,#102033);color:#e6eef7;position:sticky;top:0;height:100vh;overflow:auto;transition:.25s;width:320px;z-index:5;box-shadow:12px 0 35px rgba(15,23,42,.12)}.side.collapsed{width:84px}.side::-webkit-scrollbar{width:9px}.side::-webkit-scrollbar-thumb{background:#8aa0b5;border-radius:20px}.side-top{height:48px;display:flex;align-items:center;justify-content:space-between;padding:0 12px;background:#07111f;border-bottom:1px solid rgba(255,255,255,.07);position:sticky;top:0;z-index:3}.toggle{cursor:pointer;background:transparent;border:0;color:white;font-size:20px}.brand{padding:24px 16px;text-align:center}.brand img{max-width:158px;max-height:84px;background:#fff;border-radius:4px;object-fit:contain;padding:2px;box-shadow:0 10px 22px rgba(0,0,0,.25)}.brand p{color:#b8c8d8;font-size:14px;margin-top:18px}.side.collapsed .brand p,.side.collapsed .label,.side.collapsed .chev,.side.collapsed .subtxt{display:none}.side.collapsed .brand{padding:18px 8px}.side.collapsed .brand img{max-width:54px;max-height:54px;border-radius:12px}.menu-group{margin:10px 8px;border-radius:0;overflow:hidden}.menu-title{width:100%;border:0;display:flex;align-items:center;gap:12px;background:linear-gradient(90deg,#314c69,#345a78);color:#fff;padding:15px 13px;font-size:16px;font-weight:1000;cursor:pointer;text-align:left}.menu-title:hover{background:linear-gradient(90deg,#3b5d7d,#397099)}.menu-title .chev{margin-left:auto;transition:.18s}.menu-group.closed .chev{transform:rotate(-90deg)}.submenu{background:#14243b;padding:10px 0;max-height:650px;transition:max-height .28s ease,padding .18s ease}.menu-group.closed .submenu{max-height:0;padding:0;overflow:hidden}.menu-item{display:flex;align-items:center;gap:13px;padding:12px 20px 12px 46px;color:#e4edf8;font-weight:950;font-size:14px;border-left:4px solid transparent;transition:.13s}.menu-item:hover{background:#20334b;border-left-color:#91dcff}.menu-item.active{background:linear-gradient(90deg,#28415e,#26394f);border-left-color:#64d9ff;color:#fff;box-shadow:inset 0 0 0 1px rgba(255,255,255,.03)}.side.collapsed .menu-title{justify-content:center}.side.collapsed .menu-item{padding:14px;justify-content:center}.side.collapsed .submenu{display:none}.main{min-width:0;padding:28px 34px 50px;overflow:auto}.topbar{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:20px}.topbar h1{margin:0;font-size:34px;letter-spacing:-1px}.subtitle{color:#516982;font-size:18px;margin-top:5px}.grid{display:grid;grid-template-columns:repeat(12,1fr);gap:18px}.card{background:rgba(255,255,255,.95);border:1px solid #d6e8f2;border-radius:28px;box-shadow:var(--shadow);padding:22px}.mini{grid-column:span 4;display:flex;align-items:center;justify-content:space-between}.mini b{font-size:28px}.ico{width:58px;height:58px;border-radius:18px;display:grid;place-items:center;background:linear-gradient(135deg,#eef9ff,#f0fff5);font-size:25px}.span-12{grid-column:span 12}.span-8{grid-column:span 8}.span-4{grid-column:span 4}.doc-grid{display:grid;grid-template-columns:repeat(4,minmax(220px,1fr));gap:14px}.doc-card{background:linear-gradient(135deg,#f9fdff,#ffffff);border:1px solid #d9eaf3;border-radius:20px;padding:18px;min-height:144px;transition:.16s;position:relative;overflow:hidden}.doc-card:before{content:"";position:absolute;right:-35px;top:-35px;width:85px;height:85px;background:rgba(143,210,237,.22);border-radius:50%}.doc-card:hover{transform:translateY(-3px);box-shadow:0 16px 30px rgba(15,23,42,.10);border-color:#9ed9ef}.doc-card h3{margin:0 0 10px;font-size:18px}.doc-card p{margin:0 0 12px;color:#526b84;font-weight:500;line-height:1.45}.table-wrap{overflow:auto;border:1px solid #dbe8f3;border-radius:16px}table{width:100%;border-collapse:collapse;background:#fff}th,td{text-align:left;padding:13px 14px;border-bottom:1px solid #edf2f6;vertical-align:top}th{background:#f6f9fc;color:#314963;font-size:13px;text-transform:uppercase}tr:hover td{background:#fbfdff}.form-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;align-items:end}.detail-box{background:linear-gradient(135deg,#f8fcff,#ffffff);border:1px solid #dceaf5;border-radius:18px;padding:15px}.detail-box small{display:block;color:#60748b;margin-bottom:4px}.period-row{display:flex;gap:12px;align-items:end;flex-wrap:wrap}.mobile-head{display:none}
@media(max-width:1000px){.app{grid-template-columns:1fr}.side{position:fixed;left:-330px}.side.open{left:0}.side.collapsed{left:-90px}.mobile-head{display:flex;position:sticky;top:0;z-index:4;background:#091321;color:white;padding:11px 14px;align-items:center;justify-content:space-between}.main{padding:18px}.doc-grid{grid-template-columns:1fr}.mini,.span-8,.span-4{grid-column:span 12}.form-grid{grid-template-columns:1fr}.topbar h1{font-size:25px}.subtitle{font-size:14px}}
@media(max-width:1350px){.doc-grid{grid-template-columns:repeat(2,1fr)}}
</style>
<script>
function side(){return document.querySelector('.side')}
function saveSideScroll(){const s=side(); if(s){localStorage.setItem('sideScroll',s.scrollTop||0)}}
function restoreSideScroll(){const s=side(); if(s){s.scrollTop=parseInt(localStorage.getItem('sideScroll')||'0')}}
function toggleSide(){const s=side(); if(!s)return; if(window.innerWidth<1000){s.classList.toggle('open')}else{s.classList.toggle('collapsed'); localStorage.setItem('sideCollapsed',s.classList.contains('collapsed')?'1':'0')}}
function toggleGroup(id){const g=document.getElementById(id); if(!g)return; g.classList.toggle('closed'); localStorage.setItem('group_'+id,g.classList.contains('closed')?'1':'0')}
function initSide(){const s=side(); if(!s)return; if(localStorage.getItem('sideCollapsed')==='1' && window.innerWidth>=1000){s.classList.add('collapsed')} document.querySelectorAll('.menu-group[data-group]').forEach(g=>{const id=g.id; const saved=localStorage.getItem('group_'+id); if(saved==='1' && !g.classList.contains('force-open')) g.classList.add('closed')}); setTimeout(restoreSideScroll,60); document.querySelectorAll('.menu-item').forEach(a=>a.addEventListener('click',saveSideScroll));}
function filterCards(){const q=(document.getElementById('cardSearch')?.value||'').toLowerCase();document.querySelectorAll('.doc-card').forEach(c=>{c.style.display=c.innerText.toLowerCase().includes(q)?'block':'none'})}
window.addEventListener('DOMContentLoaded',initSide)
window.addEventListener('beforeunload',saveSideScroll)
</script></head><body>{{ body|safe }}</body></html>
'''


def render_page(content, title="Portal de Documentos PRIZE", active="Inicio"):
    body = f'''
    <div class="mobile-head"><button class="toggle" onclick="toggleSide()">☰</button><b>PRIZE Documentos</b><a href="/logout">Salir</a></div>
    <div class="app"><aside class="side"><div class="side-top"><button class="toggle" onclick="toggleSide()">←</button><b class="label">Nisira DMHT</b><button class="toggle" onclick="toggleSide()">☷</button></div>
      <div class="brand"><img src="{logo_url()}" alt="PRIZE"><p>Portal de documentos</p></div>{sidebar(active)}</aside><main class="main">{flashes()}{content}</main></div>'''
    return render_template_string(BASE, body=body, title=title)


def flashes():
    out = ""
    for cat, msg in list(getattr(request, 'flashes', []) or []):
        out += f"<div class='flash {'err' if cat=='error' else ''}'>{msg}</div>"
    # Flask get_flashed_messages unavailable without import? import below lazily
    from flask import get_flashed_messages
    out = "".join([f"<div class='flash {'err' if c=='error' else ''}'>{m}</div>" for c, m in get_flashed_messages(with_categories=True)])
    return out


def item(tipo, label, icon, active):
    cls = "menu-item active" if active == tipo else "menu-item"
    return f"<a class='{cls}' href='{url_for('panel_tipo', tipo=tipo)}'><span>{icon}</span><span class='label'>{label}</span></a>"


def sidebar(active):
    pago = ''.join(item(k,l,i,active) for k,l,i in TIPOS_PAGO)
    emp = ''.join(item(k,l,i,active) for k,l,i in TIPOS_EMPRESA)
    per = ''.join(item(k,l,i,active) for k,l,i in TIPOS_PERSONALES)
    def gclass(keys):
        return 'menu-group force-open' if active in keys else 'menu-group'
    admin = ""
    if session.get('admin_id'):
        admin = """
        <div id='grp_admin' data-group='admin' class='menu-group'>
          <button type='button' class='menu-title' onclick="toggleGroup('grp_admin')"><span>⚙️</span><span class='label'>Administrador</span><span class='chev'>∨</span></button>
          <div class='submenu'>
            <a class='menu-item' onclick='saveSideScroll()' href='/admin'><span>📊</span><span class='label'>Dashboard</span></a>
            <a class='menu-item' onclick='saveSideScroll()' href='/admin/trabajadores'><span>👥</span><span class='label'>Trabajadores</span></a>
            <a class='menu-item' onclick='saveSideScroll()' href='/admin/documentos'><span>⬆️</span><span class='label'>Subir documentos</span></a>
          </div>
        </div>"""
    pago_cls = gclass([k for k,_,_ in TIPOS_PAGO])
    emp_cls = gclass([k for k,_,_ in TIPOS_EMPRESA])
    per_cls = gclass([k for k,_,_ in TIPOS_PERSONALES])
    return f"""
    <nav>
      <div id='grp_pago' data-group='pago' class='{pago_cls}'>
        <button type='button' class='menu-title' onclick="toggleGroup('grp_pago')"><span>▣</span><span class='label'>Documentos de pago</span><span class='chev'>∨</span></button>
        <div class='submenu'>{pago}</div>
      </div>
      <div id='grp_empresa' data-group='empresa' class='{emp_cls}'>
        <button type='button' class='menu-title' onclick="toggleGroup('grp_empresa')"><span>▦</span><span class='label'>Documentos de la empresa</span><span class='chev'>∨</span></button>
        <div class='submenu'>{emp}</div>
      </div>
      <div id='grp_personal' data-group='personal' class='{per_cls}'>
        <button type='button' class='menu-title' onclick="toggleGroup('grp_personal')"><span>▤</span><span class='label'>Documentos personales</span><span class='chev'>∨</span></button>
        <div class='submenu'>{per}</div>
      </div>
      {admin}
      <div id='grp_cuenta' data-group='cuenta' class='menu-group'>
        <button type='button' class='menu-title' onclick="toggleGroup('grp_cuenta')"><span>👤</span><span class='label'>Mi cuenta</span><span class='chev'>∨</span></button>
        <div class='submenu'><a class='menu-item' onclick='saveSideScroll()' href='/panel'><span>🏠</span><span class='label'>Inicio</span></a><a class='menu-item' href='/logout'><span>🚪</span><span class='label'>Salir</span></a></div>
      </div>
    </nav>"""

def login_template(admin=False, error=""):
    action = url_for('admin_login') if admin else url_for('login')
    title = "Administrador PRIZE" if admin else "Portal de documentos PRIZE"
    sub = "Gestión interna de documentos" if admin else "Acceso al sistema"
    fields = """
      <div class='field'><label>Usuario</label><div class='login-input'>👤<input name='usuario' placeholder='Ingrese su usuario' required></div></div>
      <div class='field'><label>Clave</label><div class='login-input'>🔒<input name='clave' type='password' placeholder='Ingrese su clave' required></div></div>
    """ if admin else """
      <div class='field'><label>DNI</label><div class='login-input'>🪪<input name='dni' maxlength='8' placeholder='Ingrese su DNI' required></div></div>
      <div class='field'><label>Correo</label><div class='login-input'>✉️<input name='correo' type='email' placeholder='Ingrese su correo' required></div></div>
    """
    body = f"""
    <div class='login-body'><form class='login-card' method='post' action='{action}'><div class='login-inner'>
      <div class='login-logo'><img src='{logo_url()}'></div><div class='login-title'><h1>{title}</h1><b>{sub}</b></div>
      {f"<div class='flash err'>{error}</div>" if error else ""}{fields}<button class='btn-green'>Ingresar</button>
    </div><div class='login-links'>{'<a href="/">Ir al portal trabajador</a>' if admin else '<a href="/admin/login">Acceso administrador</a>'}</div></form></div>"""
    return render_template_string(BASE, body=body, title=title)

# =============================
# ROUTES ESTÁTICAS / LOGO
# =============================
@app.route('/_logo/<path:filename>')
def logo_file(filename):
    p = BASE_DIR / filename
    if p.exists():
        return send_file(p)
    abort(404)

@app.route('/logo_svg')
def logo_svg():
    svg = """<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 360 140'><rect width='360' height='140' fill='white'/><text x='45' y='75' font-family='Segoe UI,Arial' font-size='70' font-style='italic' fill='#2b668d'>Prize</text><circle cx='250' cy='63' r='26' fill='#ef8b16'/><circle cx='259' cy='62' r='17' fill='#f8c02a'/><text x='250' y='73' font-family='Arial' font-size='35' font-weight='900' fill='#244d77'>e</text><path d='M255 35c22-28 42-30 54-34-5 20-21 35-48 42z' fill='#35a34a'/><text x='112' y='117' font-family='Arial' font-size='28' font-weight='800' fill='#4cae55'>SUPERFRUITS</text></svg>"""
    return app.response_class(svg, mimetype='image/svg+xml')

# =============================
# LOGIN USUARIO
# =============================
@app.route('/', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        dni = normalizar_dni(request.form.get('dni'))
        correo = clean(request.form.get('correo')).lower()
        t = get_trabajador(dni)
        if not t or clean(t['correo']).lower() != correo:
            return login_template(False, "DNI o correo no coincide. Verifique sus datos.")
        session.clear(); session['dni'] = dni; session['nombre'] = t['nombre']
        return redirect(url_for('panel'))
    return login_template(False)

@app.route('/logout')
def logout():
    session.clear(); return redirect(url_for('login'))

@app.route('/panel')
@worker_required
def panel():
    dni = session['dni']; t = get_trabajador(dni)
    docs = listar_documentos(dni=dni, limit=999)
    ultimo = docs[0]['tipo'] if docs else 'Sin documento'
    cards = ''.join(doc_card(k,l,i) for k,l,i in (TIPOS_PAGO+TIPOS_EMPRESA+TIPOS_PERSONALES))
    content = f"""
    <div class='topbar'><div><h1>Todos mis documentos</h1><div class='subtitle'>{t['nombre']} · DNI {t['dni']} · {t['empresa']}</div></div><a class='btn' href='/panel'>Ver todo</a></div>
    <section class='grid'><div class='card mini'><div><span>Documentos</span><br><b>{len(docs)}</b></div><div class='ico'>🗂️</div></div><div class='card mini'><div><span>Último tipo</span><br><b>{ultimo}</b></div><div class='ico'>📄</div></div><div class='card mini'><div><span>Estado</span><br><b>Activo</b></div><div class='ico'>✅</div></div>
    <div class='card span-12'><div style='display:flex;justify-content:space-between;gap:12px;align-items:center;flex-wrap:wrap'><h2>Accesos por pestaña</h2><input id='cardSearch' onkeyup='filterCards()' class='input' style='max-width:310px' placeholder='Buscar pestaña...'></div><div class='doc-grid'>{cards}</div></div>
    <div class='card span-12'><h2>Últimos documentos</h2>{tabla_docs(docs)}</div></section>"""
    return render_page(content, active='Inicio')


def doc_card(k,l,i):
    return f"<div class='doc-card'><h3>{i} {l}</h3><p>Consulta, filtra por periodo y revisa el detalle del documento.</p><a class='btn-blue' href='{url_for('panel_tipo', tipo=k)}'>Abrir</a></div>"

@app.route('/documentos/<tipo>')
@worker_required
def panel_tipo(tipo):
    if tipo not in ALL_TIPOS: abort(404)
    dni = session['dni']; label, icon, categoria = ALL_TIPOS[tipo]
    periodo = clean(request.args.get('periodo'))
    pers = periodos_disponibles(dni=dni, tipo=tipo)
    docs = listar_documentos(dni=dni, tipo=tipo, periodo=periodo or None, limit=999)
    opts = "<option value=''>Todos los periodos</option>" + ''.join([f"<option {'selected' if p==periodo else ''}>{p}</option>" for p in pers])
    detalle = detalle_tipo(tipo, docs)
    upload_extra = ""
    if tipo == 'Otros':
        upload_extra = f"""
        <div class='card span-12'><h2>Adjuntar nuevo documento personal</h2><form method='post' action='/subir_personal' enctype='multipart/form-data' class='form-grid'>
        <input type='hidden' name='tipo' value='Otros'><div class='field'><label>Periodo</label><input name='periodo' value='{datetime.now(APP_TZ).strftime('%Y-%m')}'></div><div class='field'><label>Detalle</label><input name='detalle' placeholder='Ej: Certificado, solicitud, evidencia'></div><div class='field'><label>Archivo</label><input type='file' name='archivo' accept='.pdf,.png,.jpg,.jpeg,.webp,.doc,.docx,.xls,.xlsx' required></div><div class='field'><label>Observación</label><textarea name='observacion' rows='2' placeholder='Comentario u observación'></textarea></div><button class='btn-green'>Subir documento</button></form></div>"""
    content = f"""
    <div class='topbar'><div><h1>{icon} {label}</h1><div class='subtitle'>Solo se muestran documentos del tipo seleccionado.</div></div><a class='btn' href='/panel'>Volver</a></div>
    <section class='grid'><div class='card mini'><div><span>Total</span><br><b>{len(docs)}</b></div><div class='ico'>{icon}</div></div><div class='card mini'><div><span>Periodo</span><br><b>{periodo or 'Todos'}</b></div><div class='ico'>📅</div></div><div class='card mini'><div><span>Filtro</span><br><b>{tipo}</b></div><div class='ico'>🔎</div></div>
    <div class='card span-12'><form method='get' class='period-row'><div class='field'><label>Elegir periodo</label><select name='periodo'>{opts}</select></div><button class='btn-blue'>Aplicar filtro</button><a class='btn' href='{url_for('panel_tipo', tipo=tipo)}'>Limpiar</a></form></div>
    <div class='card span-12'><h2>Detalle de {label}</h2>{detalle}</div>{upload_extra}<div class='card span-12'><h2>Listado</h2>{tabla_docs(docs)}</div></section>"""
    return render_page(content, active=tipo)

@app.route('/subir_personal', methods=['POST'])
@worker_required
def subir_personal():
    try:
        guardar_documento(request.files.get('archivo'), session['dni'], clean(request.form.get('tipo')) or 'Otros', request.form.get('periodo'), request.form.get('detalle'), request.form.get('observacion'), session.get('dni'))
        flash('Documento personal subido correctamente.', 'ok')
    except Exception as e:
        flash(f'No se pudo subir: {e}', 'error')
    return redirect(url_for('panel_tipo', tipo='Otros'))


def detalle_tipo(tipo, docs):
    ult = docs[0] if docs else None
    label, icon, cat = ALL_TIPOS.get(tipo, (tipo,'📄',''))
    texto = {
        'Utilidad':'Boletas de participación de utilidades por periodo anual.',
        'Vacaciones':'Documentos relacionados a pago o liquidación de vacaciones.',
        'Normal':'Boletas de pago normal mensual, quincenal o semanal.',
        'Constancia Gratificación':'Constancias asociadas a gratificación.',
        'CTS':'Boletas o constancias de Compensación por Tiempo de Servicios.',
        'Liquidación':'Documentos de liquidación de beneficios sociales.',
        'Gratificación':'Boletas de gratificación ordinaria o extraordinaria.',
        'Otros':'Espacio para documentos personales adjuntos por el usuario o administrador.',
    }.get(tipo, 'Documento disponible para consulta y descarga.')
    return f"""
    <div class='grid'><div class='detail-box span-4'><small>Tipo</small><b>{icon} {label}</b></div><div class='detail-box span-4'><small>Último periodo</small><b>{ult['periodo'] if ult else 'Sin periodo'}</b></div><div class='detail-box span-4'><small>Última carga</small><b>{ult['fecha_subida'] if ult else 'Sin carga'}</b></div><div class='detail-box span-12'><small>Descripción</small>{texto}</div></div>"""


def tabla_docs(rows):
    if not rows:
        return "<div class='table-wrap'><table><tr><th>Tipo</th><th>Periodo</th><th>Detalle</th><th>Observación</th><th>Fecha</th><th>Archivo</th></tr><tr><td colspan='6'>No hay documentos en esta pestaña.</td></tr></table></div>"
    body = ''.join([f"<tr><td>{r['tipo']}</td><td>{r['periodo'] or ''}</td><td>{r['detalle'] or '-'}</td><td>{r['observacion'] or '-'}</td><td>{r['fecha_subida']}</td><td><a class='btn-blue' target='_blank' href='{url_for('ver_doc', doc_id=r['id'])}'>Ver/Descargar</a></td></tr>" for r in rows])
    return f"<div class='table-wrap'><table><thead><tr><th>Tipo</th><th>Periodo</th><th>Detalle</th><th>Observación</th><th>Fecha</th><th>Archivo</th></tr></thead><tbody>{body}</tbody></table></div>"

@app.route('/ver/<int:doc_id>')
def ver_doc(doc_id):
    with db() as con:
        r = con.execute("SELECT * FROM documentos WHERE id=?", (doc_id,)).fetchone()
    if not r: abort(404)
    if not session.get('admin_id'):
        dni = session.get('dni')
        if not dni or (r['categoria'] != 'empresa' and r['dni'] != dni): abort(403)
    path = Path(r['ruta_archivo'])
    if not path.exists(): abort(404)
    return send_file(path, as_attachment=False, download_name=r['archivo_nombre'])

# =============================
# ADMIN
# =============================
@app.route('/admin/login', methods=['GET','POST'])
def admin_login():
    if request.method == 'POST':
        u, c = clean(request.form.get('usuario')), clean(request.form.get('clave'))
        with db() as con:
            user = con.execute("SELECT * FROM usuarios_admin WHERE usuario=? AND activo=1", (u,)).fetchone()
        if not user or not check_password_hash(user['clave_hash'], c):
            return login_template(True, 'Usuario o clave incorrecta.')
        session.clear(); session['admin_id']=user['id']; session['admin_user']=user['usuario']; session['admin_nombre']=user['nombre']
        return redirect(url_for('admin'))
    return login_template(True)

@app.route('/admin')
@admin_required
def admin():
    with db() as con:
        trabajadores = con.execute("SELECT COUNT(*) FROM trabajadores").fetchone()[0]
        docs = con.execute("SELECT COUNT(*) FROM documentos").fetchone()[0]
        emp = con.execute("SELECT COUNT(*) FROM documentos WHERE categoria='empresa'").fetchone()[0]
        ult = con.execute("SELECT * FROM documentos ORDER BY id DESC LIMIT 12").fetchall()
    content = f"""
    <div class='topbar'><div><h1>Panel Administrador</h1><div class='subtitle'>Gestión total de trabajadores, boletas y documentos.</div></div><a class='btn-green' href='/admin/documentos'>Subir documentos</a></div>
    <section class='grid'><div class='card mini'><div><span>Trabajadores</span><br><b>{trabajadores}</b></div><div class='ico'>👥</div></div><div class='card mini'><div><span>Documentos</span><br><b>{docs}</b></div><div class='ico'>🗂️</div></div><div class='card mini'><div><span>Empresa</span><br><b>{emp}</b></div><div class='ico'>🏢</div></div>
    <div class='card span-12'><h2>Últimas cargas</h2>{tabla_docs(ult)}</div></section>"""
    return render_page(content, active='Admin')

@app.route('/admin/trabajadores', methods=['GET','POST'])
@admin_required
def admin_trabajadores():
    if request.method == 'POST':
        if 'excel' in request.files and request.files['excel'].filename:
            f = request.files['excel']; path = UPLOAD_DIR / f"base_{now_file()}_{secure_filename(f.filename)}"; f.save(path)
            wb = load_workbook(path, data_only=True); ws = wb.active
            headers = [clean(c.value).upper() for c in ws[1]]
            def idx(name):
                return headers.index(name) if name in headers else -1
            n=0
            with db() as con:
                for row in ws.iter_rows(min_row=2, values_only=True):
                    dni = normalizar_dni(row[idx('DNI')] if idx('DNI')>=0 else '')
                    if not dni: continue
                    nombre = clean(row[idx('NOMBRE')] if idx('NOMBRE')>=0 else '')
                    correo = clean(row[idx('CORREO')] if idx('CORREO')>=0 else '').lower()
                    cargo = clean(row[idx('CARGO')] if idx('CARGO')>=0 else '')
                    area = clean(row[idx('AREA')] if idx('AREA')>=0 else '')
                    empresa = clean(row[idx('EMPRESA')] if idx('EMPRESA')>=0 else 'PRIZE SUPERFRUITS')
                    con.execute("INSERT OR REPLACE INTO trabajadores(dni,nombre,correo,cargo,area,empresa,activo,fecha_registro) VALUES(?,?,?,?,?,?,1,?)", (dni,nombre,correo,cargo,area,empresa,now_txt()))
                    n+=1
                con.commit()
            flash(f'Base cargada correctamente: {n} trabajadores.', 'ok')
        else:
            dni=normalizar_dni(request.form.get('dni'))
            with db() as con:
                con.execute("INSERT OR REPLACE INTO trabajadores(dni,nombre,correo,cargo,area,empresa,activo,fecha_registro) VALUES(?,?,?,?,?,?,1,?)", (dni,clean(request.form.get('nombre')),clean(request.form.get('correo')).lower(),clean(request.form.get('cargo')),clean(request.form.get('area')),clean(request.form.get('empresa')) or 'PRIZE SUPERFRUITS',now_txt()))
                con.commit()
            flash('Trabajador guardado.', 'ok')
        return redirect(url_for('admin_trabajadores'))
    with db() as con:
        rows = con.execute("SELECT * FROM trabajadores ORDER BY nombre LIMIT 300").fetchall()
    table = ''.join([f"<tr><td>{r['dni']}</td><td>{r['nombre']}</td><td>{r['correo']}</td><td>{r['cargo'] or ''}</td><td>{r['empresa'] or ''}</td></tr>" for r in rows])
    content = f"""
    <div class='topbar'><div><h1>Trabajadores</h1><div class='subtitle'>Carga manual o masiva por Excel.</div></div></div><section class='grid'>
    <div class='card span-12'><h2>Nuevo trabajador</h2><form method='post' class='form-grid'><div class='field'><label>DNI</label><input name='dni' required></div><div class='field'><label>Nombre</label><input name='nombre' required></div><div class='field'><label>Correo</label><input name='correo' type='email' required></div><div class='field'><label>Cargo</label><input name='cargo'></div><div class='field'><label>Área</label><input name='area'></div><div class='field'><label>Empresa</label><input name='empresa' value='PRIZE SUPERFRUITS'></div><button class='btn-green'>Guardar</button></form></div>
    <div class='card span-12'><h2>Carga Excel</h2><form method='post' enctype='multipart/form-data' class='form-grid'><div class='field'><label>Excel con DNI, NOMBRE, CORREO, CARGO, AREA, EMPRESA</label><input type='file' name='excel' accept='.xlsx' required></div><button class='btn-blue'>Importar Excel</button></form></div>
    <div class='card span-12'><h2>Listado</h2><div class='table-wrap'><table><tr><th>DNI</th><th>Nombre</th><th>Correo</th><th>Cargo</th><th>Empresa</th></tr>{table}</table></div></div></section>"""
    return render_page(content, active='Trabajadores')

@app.route('/admin/documentos', methods=['GET','POST'])
@admin_required
def admin_documentos():
    if request.method == 'POST':
        tipo = clean(request.form.get('tipo'))
        dni = normalizar_dni(request.form.get('dni'))
        periodo = request.form.get('periodo')
        detalle = request.form.get('detalle')
        obs = request.form.get('observacion')
        files = request.files.getlist('archivos')
        ok=0
        try:
            for f in files:
                if f and f.filename:
                    guardar_documento(f, dni, tipo, periodo, detalle, obs, session.get('admin_user','admin')); ok += 1
            flash(f'Carga completada: {ok} archivo(s).', 'ok')
        except Exception as e:
            flash(f'Error en carga: {e}', 'error')
        return redirect(url_for('admin_documentos', tipo=tipo))
    tipo = clean(request.args.get('tipo')) or 'Utilidad'
    buscar = clean(request.args.get('buscar'))
    periodo = clean(request.args.get('periodo'))
    tipo_options = ''.join([f"<option value='{k}' {'selected' if k==tipo else ''}>{l}</option>" for k,l,i in TIPOS_PAGO+TIPOS_EMPRESA+TIPOS_PERSONALES])
    pers = periodos_disponibles(tipo=tipo)
    periodo_options = "<option value=''>Todos</option>" + ''.join([f"<option {'selected' if p==periodo else ''}>{p}</option>" for p in pers])
    rows = listar_documentos(tipo=tipo if tipo else None, periodo=periodo or None, buscar=buscar, limit=500)
    content = f"""
    <div class='topbar'><div><h1>Subir y gestionar documentos</h1><div class='subtitle'>Administrador: pago, empresa y documentos personales.</div></div></div><section class='grid'>
    <div class='card span-12'><h2>Carga de documentos</h2><form method='post' enctype='multipart/form-data' class='form-grid'><div class='field'><label>Tipo</label><select name='tipo'>{tipo_options}</select></div><div class='field'><label>DNI trabajador</label><input name='dni' placeholder='Vacío si es documento de empresa'></div><div class='field'><label>Periodo</label><input name='periodo' value='{datetime.now(APP_TZ).strftime('%Y-%m')}' list='periodos'></div><div class='field'><label>Detalle</label><input name='detalle' placeholder='Ej: Boleta semanal / Política actualizada'></div><div class='field'><label>Archivos</label><input type='file' name='archivos' accept='.pdf,.png,.jpg,.jpeg,.webp,.doc,.docx,.xls,.xlsx' multiple required></div><div class='field'><label>Observación</label><textarea name='observacion' rows='2'></textarea></div><button class='btn-green'>Subir</button></form></div>
    <div class='card span-12'><h2>Filtros</h2><form method='get' class='form-grid'><div class='field'><label>Tipo</label><select name='tipo'>{tipo_options}</select></div><div class='field'><label>Periodo</label><select name='periodo'>{periodo_options}</select></div><div class='field'><label>Buscar</label><input name='buscar' value='{buscar}' placeholder='DNI, detalle, observación'></div><button class='btn-blue'>Filtrar</button><a class='btn' href='/admin/documentos'>Limpiar</a></form></div>
    <div class='card span-12'><h2>Listado</h2>{tabla_docs(rows)}</div></section>"""
    return render_page(content, active=tipo)

# API compatibles
@app.route('/api/health')
def api_health(): return jsonify({'ok': True, 'mensaje': 'Portal PRIZE activo'})
@app.route('/api/boleta/<dni>')
def api_boleta(dni):
    docs = listar_documentos(dni=dni, categoria='pago', limit=20)
    t = get_trabajador(dni)
    return jsonify({'ok': bool(t), 'trabajador': dict(t) if t else None, 'documentos': [dict(x) for x in docs]})

if __name__ == '__main__':
    port = int(os.getenv('PORT', '5000'))
    app.run(host='0.0.0.0', port=port, debug=False)
