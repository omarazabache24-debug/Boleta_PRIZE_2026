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
from copy import copy
from functools import wraps
from zoneinfo import ZoneInfo

from flask import Flask, request, redirect, url_for, session, send_file, render_template_string, flash, jsonify, abort
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

BASE_DIR = Path(__file__).resolve().parent
PERSIST_DIR = Path(os.getenv("PERSIST_DIR", "/data" if Path("/data").is_dir() else str(BASE_DIR)))
STATIC_DIR = BASE_DIR / "static"
UPLOAD_DIR = PERSIST_DIR / "uploads"
EXCEL_LOCAL_DIR = PERSIST_DIR / "REGISTROS_EXCEL_LOCAL"
DB_PATH = PERSIST_DIR / "boletas_prize.db"
APP_TZ = ZoneInfo(os.getenv("APP_TIMEZONE", "America/Lima"))

for d in (PERSIST_DIR, STATIC_DIR, UPLOAD_DIR, EXCEL_LOCAL_DIR):
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
# Carpeta documental LOCAL y dinámica:
# Por defecto se crea al costado de app.py, en la misma carpeta donde tienes tus archivos.
# Ejemplo Windows: si app.py está en D:\MiSistema\, se crea D:\MiSistema\DOCUMENTOS_PRIZE_AUTO\
# En Render también se crea dentro del proyecto, pero para carga real masiva se recomienda uso local.
DOCUMENTOS_BASE_DIR = Path(os.getenv("DOCUMENTOS_BASE_DIR", str(BASE_DIR / "DOCUMENTOS_PRIZE_AUTO")))
DOCUMENTOS_BASE_DIR.mkdir(parents=True, exist_ok=True)

def slug_folder(v):
    v = clean(v).upper() if 'clean' in globals() else str(v or '').strip().upper()
    v = (v.replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U').replace('Ñ','N'))
    return re.sub(r"[^A-Z0-9]+", " ", v).strip() or "GENERAL"

def asegurar_carpetas_documentales(tipo=None):
    """Crea automáticamente la estructura física al entrar/click en pestañas."""
    grupos = []
    if tipo and tipo in ALL_TIPOS:
        label, icon, cat = ALL_TIPOS[tipo]
        grupos = [(tipo, label, cat)]
    else:
        grupos = [(k,l,'pago') for k,l,i in TIPOS_PAGO] + [(k,l,'empresa') for k,l,i in TIPOS_EMPRESA] + [(k,l,'personal') for k,l,i in TIPOS_PERSONALES]
    for k, label, cat in grupos:
        base = DOCUMENTOS_BASE_DIR / ({'pago':'DOCUMENTOS DE PAGO','empresa':'DOCUMENTOS DE LA EMPRESA','personal':'DOCUMENTOS PERSONALES'}.get(cat,'DOCUMENTOS')) / slug_folder(label)
        base.mkdir(parents=True, exist_ok=True)
        if k == 'Normal':
            (base / 'MENSUAL').mkdir(parents=True, exist_ok=True)
            (base / 'SEMANAL').mkdir(parents=True, exist_ok=True)
        if cat == 'pago':
            for y in range(datetime.now(APP_TZ).year-1, datetime.now(APP_TZ).year+2):
                (base / str(y)).mkdir(parents=True, exist_ok=True)
    return True


def now_txt():
    return datetime.now(APP_TZ).strftime("%d/%m/%Y %I:%M %p")


def now_file():
    return datetime.now(APP_TZ).strftime("%Y%m%d_%H%M%S")


def clean(v):
    return str(v or "").strip()



def fecha_sin_hora(v):
    """Muestra fechas sin 00:00:00, aceptando Excel datetime, ISO y texto dd/mm/aaaa."""
    if v is None:
        return ''
    if hasattr(v, 'strftime'):
        return v.strftime('%d/%m/%Y')
    txt = clean(v)
    if not txt:
        return ''
    txt = re.sub(r'\s+00:00:00$', '', txt)
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y %H:%M:%S', '%d/%m/%Y'):
        try:
            return datetime.strptime(txt, fmt).strftime('%d/%m/%Y')
        except Exception:
            pass
    return txt.split()[0] if '00:00:00' in txt else txt


def excel_cell_fecha(v):
    return fecha_sin_hora(v)


def exportar_tabla_excel(nombre_archivo, tabla, columnas):
    """Respaldo local en Excel para que la información sobreviva reinicios y se pueda auditar."""
    try:
        EXCEL_LOCAL_DIR.mkdir(parents=True, exist_ok=True)
        path = EXCEL_LOCAL_DIR / nombre_archivo
        wb = Workbook(); ws = wb.active; ws.title = tabla[:31]
        ws.append([titulo for titulo, campo in columnas])
        with db() as con:
            rows = con.execute(f"SELECT * FROM {tabla}").fetchall()
        for r in rows:
            ws.append([fecha_sin_hora(r[campo]) if 'fecha' in campo.lower() else (r[campo] if campo in r.keys() else '') for titulo, campo in columnas])
        for i, _ in enumerate(columnas, 1):
            ws.column_dimensions[chr(64+i) if i <= 26 else 'A'].width = 24
        for cell in ws[1]:
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor='1F2937')
            cell.alignment = Alignment(horizontal='center')
        ws.freeze_panes = 'A2'
        wb.save(path)
        return path
    except Exception as e:
        print('No se pudo exportar Excel local', tabla, e)
        return None


def respaldar_exceles_locales():
    exportar_tabla_excel('01_TRABAJADORES_LOCAL.xlsx', 'trabajadores', [
        ('EMPRESA','empresa'),('DNI','dni'),('TRABAJADOR','nombre'),('CARGO','cargo'),('AREA','area'),('JEFE INMEDIATO','jefe_dni'),('JEFE NOMBRE','jefe_nombre'),('PLANILLA','planilla'),('CORREO','correo'),('FECHA NACIMIENTO','fecha_nacimiento'),('FECHA INGRESO','fecha_ingreso'),('USUARIO','usuario_portal'),('CLAVE','clave_portal'),('ACTIVO','activo'),('FECHA REGISTRO','fecha_registro')])
    exportar_tabla_excel('02_VACACIONES_SALDOS_LOCAL.xlsx', 'vacaciones_saldos', [
        ('DNI','dni'),('TRABAJADOR','trabajador'),('EMPRESA','empresa'),('AREA','area'),('JEFE','jefe'),('JEFE DNI','jefe_dni'),('FECHA INGRESO','fecha_ingreso'),('I_PERIODO','periodo_inicio'),('F_PERIODO','periodo_fin'),('DIAS GANADOS','dias_ganados'),('DIAS GOZADOS','dias_gozados'),('SALDO','saldo'),('PERIODO','periodo'),('FECHA CARGA','fecha_carga')])
    exportar_tabla_excel('03_VACACIONES_SOLICITUDES_LOCAL.xlsx', 'vacaciones_solicitudes', [
        ('ID','id'),('DNI','dni'),('TRABAJADOR','trabajador'),('JEFE DNI','jefe_dni'),('FECHA INICIO','fecha_inicio'),('FECHA FIN','fecha_fin'),('DIAS','dias'),('MOTIVO','motivo'),('ESTADO','estado'),('FECHA SOLICITUD','fecha_solicitud'),('PERIODO DETALLE','periodo_detalle'),('PERIODO IDS','periodo_ids'),('COMENTARIO JEFE','comentario_jefe'),('COMENTARIO GH','comentario_gh')])


def restaurar_trabajadores_desde_excel_si_db_vacia():
    """Si Render/local reinicia con BD vacía, recupera trabajadores desde el Excel local."""
    path = EXCEL_LOCAL_DIR / '01_TRABAJADORES_LOCAL.xlsx'
    if not path.exists():
        return
    try:
        with db() as con:
            total = con.execute('SELECT COUNT(*) FROM trabajadores').fetchone()[0]
            if total > 1:
                return
            wb = load_workbook(path, data_only=True); ws = wb.active
            headers=[clean(c.value).upper() for c in ws[1]]
            def idx(n): return headers.index(n) if n in headers else -1
            for row in ws.iter_rows(min_row=2, values_only=True):
                dni=normalizar_dni(row[idx('DNI')] if idx('DNI')>=0 else '')
                if not dni: continue
                fecha_nac=excel_cell_fecha(row[idx('FECHA NACIMIENTO')] if idx('FECHA NACIMIENTO')>=0 else '')
                fecha_ing=excel_cell_fecha(row[idx('FECHA INGRESO')] if idx('FECHA INGRESO')>=0 else '')
                clave=clean(row[idx('CLAVE')] if idx('CLAVE')>=0 else '') or (re.sub(r'\D','', fecha_nac) or dni)
                con.execute("INSERT OR REPLACE INTO trabajadores(dni,nombre,correo,cargo,area,jefe_dni,jefe_nombre,empresa,planilla,fecha_nacimiento,fecha_ingreso,usuario_portal,clave_portal,activo,fecha_registro) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,1,?)", (dni,clean(row[idx('TRABAJADOR')] if idx('TRABAJADOR')>=0 else ''),clean(row[idx('CORREO')] if idx('CORREO')>=0 else '').lower(),clean(row[idx('CARGO')] if idx('CARGO')>=0 else ''),clean(row[idx('AREA')] if idx('AREA')>=0 else ''),normalizar_dni(row[idx('JEFE INMEDIATO')] if idx('JEFE INMEDIATO')>=0 else row[idx('JEFE DNI')] if idx('JEFE DNI')>=0 else ''),clean(row[idx('JEFE NOMBRE')] if idx('JEFE NOMBRE')>=0 else ''),clean(row[idx('EMPRESA')] if idx('EMPRESA')>=0 else 'AQUANQA') or 'AQUANQA',clean(row[idx('PLANILLA')] if idx('PLANILLA')>=0 else ''),fecha_nac,fecha_ing,dni,clave,now_txt()))
            con.commit()
    except Exception as e:
        print('No se pudo restaurar trabajadores desde Excel local', e)

def normalizar_dni(v):
    d = re.sub(r"\D", "", str(v or ""))
    return d[-8:].zfill(8) if d else ""


def safe_periodo(p):
    return re.sub(r"[^A-Za-z0-9_\- ]", "", clean(p))[:50] or datetime.now(APP_TZ).strftime("%Y-%m")





def periodo_year_value(v):
    """Normaliza I_PERIODO/F_PERIODO para trabajar solo con años: 2025, 2026, etc."""
    if v is None:
        return ''
    if hasattr(v, 'year'):
        return str(v.year)
    txt = clean(v)
    m = re.search(r'(20\d{2}|19\d{2})', txt)
    return m.group(1) if m else txt

def periodo_anual_texto(inicio='', fin=''):
    """Devuelve periodo en formato 2025/2026 usando fechas de inicio y fin."""
    def year(v):
        if v is None:
            return ''
        if hasattr(v, 'year'):
            return str(v.year)
        txt = clean(v)
        m = re.search(r'(20\d{2}|19\d{2})', txt)
        return m.group(1) if m else ''
    yi, yf = year(inicio), year(fin)
    if yi and yf and yi != yf:
        return f'{yi}/{yf}'
    return yi or yf or clean(inicio) or clean(fin) or datetime.now(APP_TZ).strftime('%Y')

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
    conn.create_function('normalizar_dni_sql', 1, normalizar_dni)
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
        # Migraciones livianas para versiones anteriores
        for col, ddl in [
            ('planilla', 'ALTER TABLE trabajadores ADD COLUMN planilla TEXT'),
            ('foto_ruta', 'ALTER TABLE trabajadores ADD COLUMN foto_ruta TEXT'),
            ('fecha_nacimiento', 'ALTER TABLE trabajadores ADD COLUMN fecha_nacimiento TEXT'),
            ('fecha_ingreso', 'ALTER TABLE trabajadores ADD COLUMN fecha_ingreso TEXT'),
            ('empresa_login', 'ALTER TABLE trabajadores ADD COLUMN empresa_login TEXT'),
            ('usuario_portal', 'ALTER TABLE trabajadores ADD COLUMN usuario_portal TEXT'),
            ('clave_portal', 'ALTER TABLE trabajadores ADD COLUMN clave_portal TEXT'),
            ('jefe_dni', 'ALTER TABLE trabajadores ADD COLUMN jefe_dni TEXT'),
            ('jefe_nombre', 'ALTER TABLE trabajadores ADD COLUMN jefe_nombre TEXT'),
        ]:
            try: con.execute(ddl)
            except Exception: pass
        for col, ddl in [
            ('fecha_ingreso', 'ALTER TABLE vacaciones_saldos ADD COLUMN fecha_ingreso TEXT'),
            ('periodo_inicio', 'ALTER TABLE vacaciones_saldos ADD COLUMN periodo_inicio TEXT'),
            ('periodo_fin', 'ALTER TABLE vacaciones_saldos ADD COLUMN periodo_fin TEXT'),
            ('jefe_dni', 'ALTER TABLE vacaciones_saldos ADD COLUMN jefe_dni TEXT'),
        ]:
            try: con.execute(ddl)
            except Exception: pass
        for col, ddl in [
            ('estado', "ALTER TABLE documentos ADD COLUMN estado TEXT DEFAULT 'Pendiente'"),
            ('comentario_rechazo', 'ALTER TABLE documentos ADD COLUMN comentario_rechazo TEXT'),
            ('fecha_aceptacion', 'ALTER TABLE documentos ADD COLUMN fecha_aceptacion TEXT'),
            ('fecha_firma', 'ALTER TABLE documentos ADD COLUMN fecha_firma TEXT'),
            ('fecha_aprobacion', 'ALTER TABLE documentos ADD COLUMN fecha_aprobacion TEXT'),
            ('fecha_lectura', 'ALTER TABLE documentos ADD COLUMN fecha_lectura TEXT'),
            ('leido_por', 'ALTER TABLE documentos ADD COLUMN leido_por TEXT'),
        ]:
            try: con.execute(ddl)
            except Exception: pass
        con.execute('''
        CREATE TABLE IF NOT EXISTS eventos_documento(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            documento_id INTEGER,
            dni TEXT,
            evento TEXT,
            fecha TEXT,
            detalle TEXT
        )''')
        con.execute('''
        CREATE TABLE IF NOT EXISTS app_config(
            clave TEXT PRIMARY KEY,
            valor TEXT
        )''')
        con.execute('''
        CREATE TABLE IF NOT EXISTS login_intentos(
            dni TEXT PRIMARY KEY,
            intentos INTEGER DEFAULT 0,
            bloqueado INTEGER DEFAULT 0,
            ultima_fecha TEXT
        )''')

        con.execute('''
        CREATE TABLE IF NOT EXISTS vacaciones_saldos(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            dni TEXT,
            trabajador TEXT,
            empresa TEXT,
            area TEXT,
            jefe TEXT,
            jefe_dni TEXT,
            fecha_ingreso TEXT,
            periodo_inicio TEXT,
            periodo_fin TEXT,
            dias_ganados REAL DEFAULT 0,
            dias_gozados REAL DEFAULT 0,
            saldo REAL DEFAULT 0,
            periodo TEXT,
            fecha_carga TEXT,
            uploaded_by TEXT
        )''')
        try:
            cols_info = con.execute("PRAGMA table_info(vacaciones_saldos)").fetchall()
            has_id = any(c[1] == 'id' for c in cols_info)
            dni_is_pk = any(c[1] == 'dni' and c[5] == 1 for c in cols_info)
            if (not has_id) or dni_is_pk:
                con.execute('ALTER TABLE vacaciones_saldos RENAME TO vacaciones_saldos_old')
                con.execute('''CREATE TABLE vacaciones_saldos(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    dni TEXT, trabajador TEXT, empresa TEXT, area TEXT, jefe TEXT, jefe_dni TEXT,
                    fecha_ingreso TEXT, periodo_inicio TEXT, periodo_fin TEXT,
                    dias_ganados REAL DEFAULT 0, dias_gozados REAL DEFAULT 0, saldo REAL DEFAULT 0,
                    periodo TEXT, fecha_carga TEXT, uploaded_by TEXT
                )''')
                con.execute('''INSERT INTO vacaciones_saldos(dni,trabajador,empresa,area,jefe,jefe_dni,fecha_ingreso,periodo_inicio,periodo_fin,dias_ganados,dias_gozados,saldo,periodo,fecha_carga,uploaded_by)
                    SELECT dni,trabajador,empresa,area,jefe,jefe_dni,fecha_ingreso,periodo_inicio,periodo_fin,dias_ganados,dias_gozados,saldo,periodo,fecha_carga,uploaded_by FROM vacaciones_saldos_old''')
                con.execute('DROP TABLE vacaciones_saldos_old')
        except Exception:
            pass
        try:
            con.execute('CREATE UNIQUE INDEX IF NOT EXISTS idx_vac_saldos_dni_periodo ON vacaciones_saldos(dni, periodo_inicio, periodo_fin)')
        except Exception:
            pass
        for col, ddl in [
            ('fecha_ingreso', 'ALTER TABLE vacaciones_saldos ADD COLUMN fecha_ingreso TEXT'),
            ('periodo_inicio', 'ALTER TABLE vacaciones_saldos ADD COLUMN periodo_inicio TEXT'),
            ('periodo_fin', 'ALTER TABLE vacaciones_saldos ADD COLUMN periodo_fin TEXT'),
            ('jefe_dni', 'ALTER TABLE vacaciones_saldos ADD COLUMN jefe_dni TEXT'),
        ]:
            try: con.execute(ddl)
            except Exception: pass

        con.execute('''
        CREATE TABLE IF NOT EXISTS vacaciones_solicitudes(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            dni TEXT,
            trabajador TEXT,
            jefe_dni TEXT,
            fecha_inicio TEXT,
            fecha_fin TEXT,
            dias REAL,
            motivo TEXT,
            estado TEXT DEFAULT 'Pendiente jefe',
            comentario_jefe TEXT,
            comentario_gh TEXT,
            fecha_solicitud TEXT,
            fecha_jefe TEXT,
            fecha_gh TEXT
        )''')
        for col, ddl in [
            ('jefe_dni', 'ALTER TABLE vacaciones_solicitudes ADD COLUMN jefe_dni TEXT'),
            ('periodo_detalle', 'ALTER TABLE vacaciones_solicitudes ADD COLUMN periodo_detalle TEXT'),
            ('periodo_ids', 'ALTER TABLE vacaciones_solicitudes ADD COLUMN periodo_ids TEXT'),
        ]:
            try: con.execute(ddl)
            except Exception: pass

        con.execute('''
        CREATE TABLE IF NOT EXISTS contratacion_docs(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            dni TEXT,
            trabajador TEXT,
            empresa TEXT,
            etapa TEXT,
            tipo_doc TEXT,
            estado TEXT DEFAULT 'Generado',
            archivo_nombre TEXT,
            ruta_archivo TEXT,
            fecha_registro TEXT,
            uploaded_by TEXT
        )''')
        con.execute('''
        CREATE TABLE IF NOT EXISTS contratacion_tipos(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT,
            descripcion TEXT,
            etapa TEXT,
            obligatorio INTEGER DEFAULT 1,
            activo INTEGER DEFAULT 1
        )''')
        if not con.execute("SELECT 1 FROM contratacion_tipos LIMIT 1").fetchone():
            base_tipos=[('104','CONTRATO TRABAJADOR','Incorporación'),('619','CONTRATO TRABAJADOR (RENOVACIÓN)','Renovación'),('524','ANEXO DE RIESGOS','Incorporación'),('797','BOLETÍN SIS. PENSIONARIO','Incorporación'),('664','CARGO DE ENTREGA','Incorporación'),('382','CARTA DE COMPROMISO','Incorporación'),('805','ACUERDO PREFERENCIAL','Incorporación'),('809','ELECCIÓN DE BENEFICIOS SOCIALES','Incorporación')]
            con.executemany("INSERT INTO contratacion_tipos(codigo,descripcion,etapa,obligatorio,activo) VALUES(?,?,?,?,1)", [(a,b,c,1) for a,b,c in base_tipos])
        asegurar_carpetas_documentales()
        # Datos demo seguros
        if not con.execute("SELECT 1 FROM usuarios_admin WHERE usuario='admin'").fetchone():
            con.execute("INSERT INTO usuarios_admin(usuario,clave_hash,nombre,rol) VALUES(?,?,?,?)",
                        ("admin", generate_password_hash("admin123"), "Administrador PRIZE", "admin"))
        if not con.execute("SELECT 1 FROM trabajadores WHERE dni='74324033'").fetchone():
            con.execute("INSERT INTO trabajadores(dni,nombre,correo,cargo,area,empresa,activo,fecha_registro) VALUES(?,?,?,?,?,?,?,?)",
                        ("74324033", "AZABACHE LUJAN, OMAR EDUARDO", "omar@demo.com", "Analista", "RR.HH.", "AQUANQA", 1, now_txt()))
        # Normaliza registros antiguos de demostración para que no aparezca PRIZE SUPERFRUITS al trabajador.
        try:
            con.execute("UPDATE trabajadores SET empresa='AQUANQA' WHERE UPPER(COALESCE(empresa,''))='PRIZE SUPERFRUITS'")
        except Exception:
            pass
        # Repara solicitudes antiguas que quedaron sin jefe_dni: toma el jefe desde saldos o ficha trabajadores.
        try:
            con.execute("""
                UPDATE vacaciones_solicitudes
                   SET jefe_dni = COALESCE(
                       (SELECT s.jefe_dni FROM vacaciones_saldos s
                         WHERE s.dni = vacaciones_solicitudes.dni
                           AND COALESCE(s.jefe_dni,'')<>''
                         ORDER BY s.periodo_inicio, s.periodo_fin LIMIT 1),
                       (SELECT tr.jefe_dni FROM trabajadores tr
                         WHERE tr.dni = vacaciones_solicitudes.dni
                           AND COALESCE(tr.jefe_dni,'')<>'' LIMIT 1),
                       jefe_dni
                   )
                 WHERE COALESCE(jefe_dni,'')=''
            """)
        except Exception:
            pass
        # Seguridad: cualquier solicitud registrada con inicio anterior a hoy queda anulada para no consumir saldo.
        try:
            hoy_txt = hoy_lima().isoformat()
            con.execute("""UPDATE vacaciones_solicitudes
                           SET estado='Anulado - fecha anterior a hoy',
                               comentario_gh=COALESCE(comentario_gh,'') || ' | Anulado automáticamente por fecha anterior a hoy.'
                         WHERE date(fecha_inicio) < date(?)
                           AND estado NOT LIKE 'Rechazado%'
                           AND estado NOT LIKE 'Anulado%'""", (hoy_txt,))
        except Exception:
            pass
        con.commit()


init_db()
restaurar_trabajadores_desde_excel_si_db_vacia()
respaldar_exceles_locales()

def get_config(clave, default=''):
    with db() as con:
        r = con.execute('SELECT valor FROM app_config WHERE clave=?', (clave,)).fetchone()
    return r['valor'] if r else default

def set_config(clave, valor):
    with db() as con:
        con.execute('INSERT INTO app_config(clave,valor) VALUES(?,?) ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor', (clave, str(valor)))
        con.commit()

def modo_prueba_activo():
    return get_config('modo_prueba', '0') == '1'

def marca_carga(usuario='sistema'):
    usuario = clean(usuario or 'sistema')
    return usuario + (' [MODO PRUEBA]' if modo_prueba_activo() else '')

def reset_intentos_login(dni):
    with db() as con:
        con.execute('DELETE FROM login_intentos WHERE dni=?', (normalizar_dni(dni),))
        con.commit()

def registrar_intento_fallido(dni):
    dni = normalizar_dni(dni)
    with db() as con:
        r = con.execute('SELECT intentos FROM login_intentos WHERE dni=?', (dni,)).fetchone()
        n = (int(r['intentos']) if r else 0) + 1
        bloqueado = 1 if n >= 3 else 0
        con.execute('INSERT INTO login_intentos(dni,intentos,bloqueado,ultima_fecha) VALUES(?,?,?,?) ON CONFLICT(dni) DO UPDATE SET intentos=?, bloqueado=?, ultima_fecha=?', (dni,n,bloqueado,now_txt(),n,bloqueado,now_txt()))
        con.commit()
    return n, bloqueado

def esta_bloqueado(dni):
    with db() as con:
        r = con.execute('SELECT bloqueado,intentos FROM login_intentos WHERE dni=?', (normalizar_dni(dni),)).fetchone()
    return bool(r and int(r['bloqueado'] or 0)==1), (int(r['intentos']) if r else 0)

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


def portal_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("dni") and not session.get("admin_id"):
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


def generar_clave_trabajador(dni, fecha_nac=''):
    """Clave del trabajador: fecha de nacimiento sin / ni guiones (ddmmaaaa)."""
    if hasattr(fecha_nac, 'strftime'):
        return fecha_nac.strftime('%d%m%Y')
    txt = str(fecha_nac or '').strip()
    nums = re.sub(r"\D", "", txt)
    if len(nums) >= 8:
        return nums[:8]
    dni = normalizar_dni(dni)
    return (dni[-4:] + "PRIZE").upper()


def registrar_evento_documento(doc_id, dni, evento, detalle=''):
    with db() as con:
        con.execute("INSERT INTO eventos_documento(documento_id,dni,evento,fecha,detalle) VALUES(?,?,?,?,?)", (doc_id, normalizar_dni(dni), evento, now_txt(), clean(detalle)))
        if evento in ['Abierto','Leído','Recibido']:
            con.execute("UPDATE documentos SET fecha_lectura=?, leido_por=? WHERE id=? AND (fecha_lectura IS NULL OR fecha_lectura='')", (now_txt(), normalizar_dni(dni), doc_id))
        con.commit()


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


def parse_fecha_any(v):
    txt = clean(v)
    if not txt:
        return None
    if isinstance(v, datetime):
        return v.date()
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"]:
        try:
            return datetime.strptime(txt.split()[0], fmt).date()
        except Exception:
            pass
    return None

def hoy_lima():
    """Fecha actual en Perú para validar vacaciones sin depender del navegador."""
    return datetime.now(APP_TZ).date()

def fecha_iso_segura(v):
    d = parse_fecha_any(v)
    return d.isoformat() if d else ''

def periodos_desde_ingreso(dni=None, tipo=None, max_meses=72):
    """Devuelve periodos desde la fecha de ingreso del trabajador hasta hoy."""
    inicio = None
    if dni:
        t = get_trabajador(dni)
        if t and 'fecha_ingreso' in t.keys():
            inicio = parse_fecha_any(t['fecha_ingreso'])
    hoy = datetime.now(APP_TZ).date()
    if not inicio:
        inicio = hoy.replace(day=1)
    inicio = inicio.replace(day=1)
    out=[]; y=inicio.year; m=inicio.month
    while (y < hoy.year or (y == hoy.year and m <= hoy.month)) and len(out) < max_meses:
        out.append(f"{y:04d}-{m:02d}")
        m += 1
        if m == 13:
            y += 1; m = 1
    docs = periodos_disponibles(dni=dni, tipo=tipo)
    return sorted(set(out + docs), reverse=True)


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
    # Copia automática a carpeta documental organizada para respaldo físico.
    try:
        label_auto, icon_auto, cat_auto = ALL_TIPOS.get(tipo, (tipo, '📄', categoria))
        root_name = {'pago':'DOCUMENTOS DE PAGO','empresa':'DOCUMENTOS DE LA EMPRESA','personal':'DOCUMENTOS PERSONALES'}.get(cat_auto,'DOCUMENTOS')
        auto_base = DOCUMENTOS_BASE_DIR / root_name / slug_folder(label_auto)
        if tipo == 'Normal' and 'seman' in clean(detalle).lower():
            auto_base = auto_base / 'SEMANAL'
        elif tipo == 'Normal':
            auto_base = auto_base / 'MENSUAL'
        auto_base = auto_base / periodo
        if dni: auto_base = auto_base / dni
        auto_base.mkdir(parents=True, exist_ok=True)
        auto_path = auto_base / final
        if str(auto_path) != str(path):
            auto_path.write_bytes(path.read_bytes())
    except Exception:
        pass
    with db() as con:
        con.execute("""
        INSERT INTO documentos(dni,categoria,tipo,periodo,detalle,observacion,archivo_nombre,ruta_archivo,extension,fecha_subida,uploaded_by)
        VALUES(?,?,?,?,?,?,?,?,?,?,?)
        """, (dni, categoria, tipo, periodo, clean(detalle), clean(observacion), original, str(path), ext, now_txt(), uploaded_by))
        con.commit()
    return str(path)


def inferir_tipo_desde_ruta(path: Path):
    texto = str(path).lower()
    reglas = [("util", "Utilidad"), ("vacacion", "Vacaciones"), ("normal", "Normal"), ("cts", "CTS"), ("liquid", "Liquidación"), ("grat", "Gratificación"), ("constancia", "Constancia Gratificación"), ("contrato", "Contrato de Trabajo"), ("sst", "Reglamento de SST"), ("interno", "Reglamento Interno"), ("conducta", "Código de Conducta"), ("politica", "Políticas"), ("políticas", "Políticas"), ("comunicado", "Comunicados")]
    for clave, tipo in reglas:
        if clave in texto:
            return tipo
    return "Otros"


def inferir_periodo_desde_ruta(path: Path):
    texto = str(path)
    m = re.search(r"(20\d{2})[-_ ]?(0[1-9]|1[0-2])", texto)
    if m: return f"{m.group(1)}-{m.group(2)}"
    m = re.search(r"(0[1-9]|1[0-2])[-_ ]?(20\d{2})", texto)
    if m: return f"{m.group(2)}-{m.group(1)}"
    m = re.search(r"(20\d{2})", texto)
    if m: return m.group(1)
    return datetime.now(APP_TZ).strftime("%Y-%m")


def documento_ya_indexado(path: Path):
    with db() as con:
        return con.execute("SELECT 1 FROM documentos WHERE ruta_archivo=?", (str(path),)).fetchone() is not None


def extraer_texto_pdf(path: Path, max_paginas=2):
    """Extrae texto de las primeras páginas para detectar DNI dentro de la boleta."""
    if path.suffix.lower() != '.pdf' or PdfReader is None:
        return ''
    try:
        reader = PdfReader(str(path))
        partes = []
        for page in reader.pages[:max_paginas]:
            try:
                partes.append(page.extract_text() or '')
            except Exception:
                pass
        return '\n'.join(partes)[:8000]
    except Exception:
        return ''


def detectar_dni_en_archivo(path: Path, dni_obj=''):
    """Busca DNI primero en ruta/nombre y, si es PDF, también dentro del contenido."""
    dni_obj = normalizar_dni(dni_obj) if dni_obj else ''
    texto_ruta = str(path)
    if dni_obj and dni_obj in texto_ruta:
        return dni_obj, 'ruta/nombre'
    m = re.search(r"(?<!\d)(\d{8})(?!\d)", texto_ruta)
    if m:
        return m.group(1), 'ruta/nombre'
    texto_pdf = extraer_texto_pdf(path)
    if texto_pdf:
        if dni_obj and dni_obj in re.sub(r'\D', '', texto_pdf):
            return dni_obj, 'contenido PDF'
        patrones = [
            r"(?:DNI|D\.?N\.?I\.?|DOC(?:UMENTO)?|COD(?:IGO)?|IDENTIDAD)\s*[:º°#-]?\s*(\d{8})",
            r"(?<!\d)(\d{8})(?!\d)",
        ]
        for pat in patrones:
            mm = re.search(pat, texto_pdf, flags=re.I)
            if mm:
                return normalizar_dni(mm.group(1)), 'contenido PDF'
    return '', ''


def detalle_auto_desde_ruta(path: Path):
    txt = str(path).lower()
    if 'semanal' in txt or 'semana' in txt:
        return 'Boleta semanal - Importado automáticamente desde carpeta'
    if 'mensual' in txt or 'mes' in txt:
        return 'Boleta mensual - Importado automáticamente desde carpeta'
    return 'Importado automáticamente desde carpeta'


def registrar_archivo_existente(path: Path, dni: str, tipo: str, uploaded_by="auto", fuente_dni='ruta/nombre'):
    uploaded_by = marca_carga(uploaded_by)
    if documento_ya_indexado(path): return False
    ext = path.suffix.lower()
    if ext not in EXT_ALLOWED: return False
    label, icon, categoria = ALL_TIPOS.get(tipo, (tipo, "📄", "personal"))
    dni = normalizar_dni(dni) if categoria != "empresa" else ""
    periodo = inferir_periodo_desde_ruta(path)
    detalle = detalle_auto_desde_ruta(path)
    obs = f"Detectado automáticamente por {fuente_dni}. Ruta: {path.parent}"
    with db() as con:
        con.execute("""
        INSERT INTO documentos(dni,categoria,tipo,periodo,detalle,observacion,archivo_nombre,ruta_archivo,extension,fecha_subida,uploaded_by)
        VALUES(?,?,?,?,?,?,?,?,?,?,?)
        """, (dni, categoria, tipo, periodo, detalle, obs, path.name, str(path), ext, now_txt(), uploaded_by))
        con.commit()
    return True


def sincronizar_documentos_carpeta(dni=None, devolver_resumen=False):
    # Detecta documentos desde DOCUMENTOS_PRIZE_AUTO, incluyendo:
    # DOCUMENTOS DE PAGO / BOLETAS NORMAL / SEMANAL.
    # Si el DNI no está en el nombre, lee el PDF e intenta encontrarlo dentro del texto.
    asegurar_carpetas_documentales()
    base_dirs = []
    for b in [DOCUMENTOS_BASE_DIR, BASE_DIR / "documentos_auto"]:
        if b.exists() and b.is_dir() and b not in base_dirs:
            base_dirs.append(b)
    total = 0
    revisados = omitidos = duplicados = sin_dni = sin_trabajador = 0
    dni_obj = normalizar_dni(dni) if dni else ""
    for base in base_dirs:
        for path in base.rglob("*"):
            if not path.is_file() or path.suffix.lower() not in EXT_ALLOWED:
                continue
            revisados += 1
            if documento_ya_indexado(path):
                duplicados += 1; continue
            tipo = inferir_tipo_desde_ruta(path)
            categoria = ALL_TIPOS.get(tipo, ("", "", "personal"))[2]
            dni_detectado, fuente = detectar_dni_en_archivo(path, dni_obj)
            if categoria != "empresa" and not dni_detectado:
                sin_dni += 1; omitidos += 1; continue
            if dni_obj and categoria != "empresa" and dni_detectado != dni_obj:
                omitidos += 1; continue
            if categoria != "empresa":
                trab = get_trabajador(dni_detectado)
                if not trab or int(trab['activo'] or 0) != 1:
                    sin_trabajador += 1; omitidos += 1; continue
            try:
                if registrar_archivo_existente(path, dni_detectado, tipo, uploaded_by="auto carpeta local", fuente_dni=fuente or 'carpeta'):
                    total += 1
            except Exception:
                omitidos += 1
    resumen = {'nuevos': total, 'revisados': revisados, 'duplicados': duplicados, 'omitidos': omitidos, 'sin_dni': sin_dni, 'sin_trabajador': sin_trabajador, 'base': str(DOCUMENTOS_BASE_DIR)}
    return resumen if devolver_resumen else total

# =============================
# ESTILOS Y LAYOUT
# =============================
BASE = r'''
<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{{ title }}</title>
<style>
:root{--txt:#eef2f7;--mut:#a8b0bb;--yellow:#ffd23f;--yellow2:#ffb21a;--dark:#15181d;--panel:#1e2025;--panel2:#171a20;--line:#343a43;--shadow:0 24px 60px rgba(0,0,0,.35)}
*{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;font-family:Inter,Segoe UI,Arial,sans-serif;color:var(--txt);background:#15181d;font-weight:650}a{text-decoration:none;color:inherit}.hidden{display:none!important}.btn,.btn-blue,.btn-green,.btn-red{border:1px solid #3a414b;border-radius:14px;padding:12px 18px;background:#20242b;color:#eef2f7;font-weight:950;cursor:pointer;display:inline-flex;align-items:center;gap:8px;box-shadow:0 10px 22px rgba(0,0,0,.18);transition:.16s}.btn:hover,.btn-blue:hover,.btn-green:hover{transform:translateY(-1px);box-shadow:0 16px 34px rgba(0,0,0,.28)}.btn-blue{background:linear-gradient(135deg,#292e36,#1f232b);border-color:#4a525e;color:var(--yellow)}.btn-green{background:linear-gradient(135deg,var(--yellow2),var(--yellow));border:0;color:#1d1f24}.btn-red{background:#361a24;border-color:#7f1d1d;color:#fecaca}.flash{padding:13px 16px;border-radius:16px;margin:10px 0;background:#2a2a1b;border:1px solid #9a7b16;color:#ffefaa}.flash.err{background:#3a1720;border-color:#7f1d1d;color:#fecaca}.input,select,textarea{width:100%;border:1px solid #3a414b;border-radius:14px;padding:13px 14px;background:#111418;color:#f8fafc;font:inherit;outline:none}option{background:#111418;color:#f8fafc}.input:focus,select:focus,textarea:focus{border-color:var(--yellow);box-shadow:0 0 0 4px rgba(255,210,63,.12)}
/* LOGIN - estilo imagen negra/amarilla */
.login-body{min-height:100vh;display:grid;place-items:center;padding:20px;position:relative;overflow:hidden;background:linear-gradient(rgba(22,25,29,.86),rgba(22,25,29,.90)),radial-gradient(circle at 7% 4%,#ffd23f 0 23%,transparent 23.2%),radial-gradient(circle at 94% -2%,#ffd23f 0 11%,transparent 11.2%),radial-gradient(circle at 72% 112%,#ffd23f 0 20%,transparent 20.2%),linear-gradient(135deg,#2a2e33,#111418)}.login-card{width:min(92vw,500px);background:linear-gradient(180deg,rgba(25,28,33,.98),rgba(29,33,38,.95));border:1px solid rgba(255,255,255,.10);border-radius:18px;padding:38px 42px 0;box-shadow:0 38px 90px rgba(0,0,0,.52);overflow:hidden;position:relative}.login-card:before{content:"";position:absolute;left:-72px;bottom:-58px;width:365px;height:150px;background:linear-gradient(135deg,#2e4f86,#5d83e6);border-radius:50% 50% 0 0;transform:rotate(-8deg);opacity:.95}.login-card:after{content:"";position:absolute;right:-78px;bottom:-52px;width:350px;height:145px;background:linear-gradient(135deg,#253849,#475b6f);border-radius:50% 50% 0 0;transform:rotate(8deg);opacity:.92}.login-inner{position:relative;z-index:2}.login-logo{text-align:center}.login-logo img{max-width:145px;max-height:105px;object-fit:contain;background:rgba(255,255,255,.92);border-radius:10px;padding:7px;filter:drop-shadow(0 14px 24px rgba(0,0,0,.45))}.login-title{text-align:center;margin:20px 0 30px;color:#aeb7c3}.login-title h1{margin:0 0 7px;color:#fff;font-size:24px;letter-spacing:.5px;text-transform:uppercase}.login-title b{color:#98a4b3}.login-card .field label{display:none}.login-input{display:flex;align-items:center;gap:13px;background:transparent;border-bottom:1px solid rgba(226,232,240,.40);padding:0 6px;margin-bottom:22px;transition:.18s}.login-input:focus-within{border-bottom-color:var(--yellow);box-shadow:0 10px 0 -9px rgba(255,210,63,.9)}.login-input input{border:0;padding:15px 8px;width:100%;font:inherit;outline:none;background:transparent;color:#fff;font-weight:900}.login-input input::placeholder{color:#cbd5e1}.login-card .btn-green{width:auto;justify-content:center;font-size:15px;margin:8px 0 74px;padding:14px 34px;border-radius:28px;background:linear-gradient(135deg,var(--yellow2),var(--yellow));color:#212529;border:0;box-shadow:0 14px 30px rgba(255,178,26,.35)}.login-links{text-align:center;margin-top:-48px;padding-bottom:24px;position:relative;z-index:3}.login-links a{color:#dbeafe;font-size:13px;font-weight:900}
/* APP - dashboard ejecutivo */
.app{display:grid;grid-template-columns:320px 1fr;min-height:100vh;background:#15181d;transition:grid-template-columns .22s ease}.app.side-collapsed{grid-template-columns:86px 1fr}.side{background:linear-gradient(180deg,#1e2024,#171a1f 72%,#111318);color:#f2f4f8;position:sticky;top:0;height:100vh;overflow:auto;transition:.25s;width:320px;z-index:5;box-shadow:12px 0 35px rgba(0,0,0,.34);border-right:1px solid #33373d}.side.collapsed{width:86px}.side-top{height:54px;display:flex;align-items:center;justify-content:space-between;padding:0 14px;background:#17191e;border-bottom:1px solid rgba(255,255,255,.07);position:sticky;top:0;z-index:3}.toggle{cursor:pointer;background:transparent;border:0;color:white;font-size:21px}.brand{padding:28px 16px 22px;text-align:center}.brand img{max-width:150px;max-height:95px;background:rgba(255,255,255,.90);border-radius:16px;object-fit:contain;padding:8px;box-shadow:0 14px 30px rgba(0,0,0,.35);border:1px solid rgba(255,210,63,.28)}.brand p{color:#c8cdd6;font-size:14px;margin-top:18px}.side.collapsed .brand p,.side.collapsed .label,.side.collapsed .chev,.side.collapsed .subtxt,.side.collapsed .side-user{display:none}.side.collapsed .brand{padding:20px 8px}.side.collapsed .brand img{max-width:55px;max-height:55px;border-radius:14px;padding:4px}.menu-group{margin:10px 12px;border-radius:12px;overflow:hidden}.menu-title{width:100%;border:1px solid rgba(255,255,255,.06);display:flex;align-items:center;gap:12px;background:linear-gradient(135deg,#22252b,#1b1e24);color:#eef2f7;padding:15px 14px;font-size:15px;font-weight:1000;cursor:pointer;text-align:left;border-radius:12px}.menu-title:hover{background:linear-gradient(135deg,#2b2f36,#23272f)}.menu-group.force-open .menu-title{background:linear-gradient(135deg,var(--yellow2),var(--yellow));color:#181a1f;box-shadow:0 14px 30px rgba(255,210,63,.20)}.menu-title .chev{margin-left:auto;transition:.18s}.menu-group.closed .chev{transform:rotate(-90deg)}.submenu{background:transparent;padding:9px 0;max-height:720px;transition:max-height .28s ease,padding .18s ease}.menu-group.closed .submenu{max-height:0;padding:0;overflow:hidden}.menu-item{display:flex;align-items:center;gap:13px;padding:13px 18px 13px 40px;color:#dce3ed;font-weight:900;font-size:14px;border-left:4px solid transparent;transition:.13s;border-radius:10px;margin:2px 0}.menu-item:hover{background:#242830;border-left-color:var(--yellow)}.menu-item.active,.menu-title.active{background:linear-gradient(135deg,#34302a,#2a2926);border-left-color:var(--yellow);color:#fff}.side.collapsed .menu-title{justify-content:center;padding:18px 10px}.side.collapsed .menu-item{padding:16px 10px;justify-content:center}.side.collapsed .submenu{display:none}.main{min-width:0;padding:0 34px 50px;overflow:auto;background:radial-gradient(circle at 92% -8%,rgba(255,210,63,.22),transparent 22%),radial-gradient(circle at 100% 96%,rgba(255,210,63,.12),transparent 28%),#15181d}.hero{margin:0 -34px 24px;padding:26px 34px 28px;background:radial-gradient(circle at 72% 0%,rgba(255,210,63,.20),transparent 32%),linear-gradient(120deg,#15181d 0%,#111418 62%,#24282d 100%);border-bottom:1px solid #31363d}.topbar{display:flex;align-items:center;justify-content:space-between;gap:12px}.topbar h1{margin:0;font-size:34px;letter-spacing:-1px;color:#fff}.topbar h1 .accent{color:var(--yellow)}.subtitle{color:#aeb7c3;font-size:16px;margin-top:7px}.grid{display:grid;grid-template-columns:repeat(12,1fr);gap:18px}.card{background:linear-gradient(145deg,#202329,#181b20);border:1px solid #303640;border-radius:18px;box-shadow:0 22px 55px rgba(0,0,0,.25);padding:22px;color:#eef2f7}.mini{grid-column:span 4;display:flex;align-items:center;justify-content:space-between}.mini b{font-size:28px;color:var(--yellow)}.ico{width:56px;height:56px;border-radius:16px;display:grid;place-items:center;background:linear-gradient(135deg,var(--yellow),var(--yellow2));font-size:24px;color:#17191e;box-shadow:0 12px 26px rgba(255,210,63,.18)}.span-12{grid-column:span 12}.span-8{grid-column:span 8}.span-4{grid-column:span 4}.span-6{grid-column:span 6}.span-3{grid-column:span 3}.doc-grid{display:grid;grid-template-columns:repeat(4,minmax(220px,1fr));gap:14px}.doc-card{background:linear-gradient(145deg,#24272d,#1b1f25);border:1px solid #343a43;border-radius:16px;padding:18px;min-height:158px;transition:.16s;position:relative;overflow:hidden}.doc-card:before{content:"";position:absolute;right:-34px;top:-34px;width:86px;height:86px;background:rgba(255,210,63,.17);border-radius:50%}.doc-card h3{margin:0 0 12px;font-size:17px;color:#fff}.doc-card p{margin:0 0 14px;color:#c0c8d2;font-weight:500;line-height:1.45}.doc-card:hover{transform:translateY(-2px);border-color:var(--yellow);box-shadow:0 16px 30px rgba(0,0,0,.25)}.table-wrap{overflow:auto;border:1px solid #343a43;border-radius:14px}table{width:100%;border-collapse:collapse;background:#171a20;color:#eaf3ff}th,td{text-align:left;padding:13px 14px;border-bottom:1px solid #2c323a;vertical-align:top}th{background:#111418;color:var(--yellow);font-size:13px;text-transform:uppercase}tr:hover td{background:#20242b}.form-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;align-items:end}.detail-box{background:linear-gradient(135deg,#202329,#171a20);border:1px solid #343a43;border-radius:16px;padding:15px}.detail-box small{display:block;color:#aeb7c3;margin-bottom:4px}.period-row{display:flex;gap:12px;align-items:end;flex-wrap:wrap}.mobile-head{display:none}.side-user{margin:26px 14px 14px;padding-top:20px;border-top:1px solid rgba(255,255,255,.08);display:flex;align-items:center;gap:11px;color:#e5e7eb}.avatar{width:44px;height:44px;border-radius:50%;display:grid;place-items:center;background:var(--yellow);color:#15181d;font-weight:1000}
@media(max-width:1000px){.app,.app.side-collapsed{grid-template-columns:1fr}.side{position:fixed;left:-335px;width:315px}.side.open{left:0}.side.collapsed{left:-335px}.mobile-head{display:flex;position:sticky;top:0;z-index:20;background:#17191e;color:white;padding:12px 14px;align-items:center;justify-content:space-between;border-bottom:1px solid #343a43}.main{padding:0 14px 30px}.hero{margin:0 -14px 18px;padding:20px 16px}.doc-grid{grid-template-columns:1fr}.mini,.span-8,.span-4,.span-3{grid-column:span 12}.form-grid{grid-template-columns:1fr}.topbar{align-items:flex-start;flex-direction:column}.topbar h1{font-size:24px}.subtitle{font-size:13px}.card{border-radius:16px;padding:17px}.login-card{padding:32px 28px 0}.login-card .btn-green{width:100%}}@media(min-width:1001px) and (max-width:1350px){.doc-grid{grid-template-columns:repeat(2,1fr)}}

/* === RETOQUE PRO ADMIN / FORMULARIOS === */
.login-card{border-radius:26px;background:linear-gradient(180deg,rgba(24,27,32,.98),rgba(16,18,23,.96));backdrop-filter:blur(8px)}
.login-input{border:1px solid rgba(255,255,255,.10);border-radius:16px;background:rgba(255,255,255,.055);box-shadow:inset 0 0 0 1px rgba(255,255,255,.02)}
.login-input:focus-within{background:rgba(255,210,63,.10);border-color:rgba(255,210,63,.72);box-shadow:0 0 0 4px rgba(255,210,63,.12)}
.login-input input{color:#fff;background:transparent}.login-input input:-webkit-autofill{-webkit-box-shadow:0 0 0 1000px #202329 inset!important;-webkit-text-fill-color:#fff!important}
.form-grid{grid-template-columns:repeat(12,1fr);align-items:end}.form-grid .field{grid-column:span 3}.form-grid .field:nth-child(4n+1){grid-column:span 3}.form-grid button,.form-grid .btn,.form-grid .btn-blue,.form-grid .btn-green{grid-column:span 3;justify-content:center;height:54px}.field label{display:block;margin-bottom:8px;color:#eaf0f7;font-size:13px;letter-spacing:.3px}.field input,.field select,.field textarea,.input,select,textarea{background:#0f1319;border:1px solid #3b414b;color:#fff;border-radius:14px;min-height:48px;padding:12px 14px;font-weight:800}.field input:focus,.field select:focus,.field textarea:focus,.input:focus,select:focus,textarea:focus{border-color:var(--yellow);box-shadow:0 0 0 4px rgba(255,210,63,.12);outline:none}.card form{gap:18px}.alert-card{background:linear-gradient(145deg,#202329,#16191e 65%,rgba(255,210,63,.06));}.alert-item{display:grid;grid-template-columns:48px 1fr auto;gap:12px;align-items:center;padding:13px 0;border-top:1px solid #323740}.alert-item:first-of-type{border-top:0}.bell{width:40px;height:40px;border-radius:14px;display:grid;place-items:center;background:linear-gradient(135deg,var(--yellow),var(--yellow2));box-shadow:0 12px 22px rgba(255,210,63,.18)}.alert-item span,.muted,.empty-note{color:#b8c0cb}.mini-btn{padding:9px 13px;border-radius:12px}.admin-hero{border-radius:0 0 24px 24px;margin-bottom:20px}.side .brand img{background:linear-gradient(145deg,#f7f7f7,#d8d8d8);mix-blend-mode:normal}.side .brand{background:radial-gradient(circle at 50% 28%,rgba(255,210,63,.10),transparent 44%)}
@media(max-width:1000px){.form-grid{grid-template-columns:1fr}.form-grid .field,.form-grid button,.form-grid .btn,.form-grid .btn-blue,.form-grid .btn-green{grid-column:span 1;width:100%}.alert-item{grid-template-columns:42px 1fr}.alert-item a{grid-column:1 / -1;justify-content:center}.side.open{box-shadow:0 0 0 999px rgba(0,0,0,.55),12px 0 35px rgba(0,0,0,.34)}}

.status-pill{display:inline-flex;padding:7px 10px;border-radius:999px;background:#242a32;border:1px solid #3b414b;color:#ffd23f;font-weight:1000;white-space:nowrap}.actions{display:flex;gap:8px;flex-wrap:wrap}.modal{position:fixed;inset:0;background:rgba(0,0,0,.70);z-index:80;display:grid;place-items:center;padding:18px}.modal-card{width:min(520px,96vw);background:#1d2128;border:1px solid #3a414b;border-radius:18px;padding:22px;box-shadow:var(--shadow)}.profile-row{display:flex;align-items:center;gap:18px;flex-wrap:wrap}.profile-img{width:92px;height:92px;border-radius:50%;object-fit:cover;background:#fff;padding:4px;border:3px solid var(--yellow)}.profile-form{flex:1;min-width:240px}.sub-mini{padding-left:58px!important;font-size:13px!important;opacity:.92}.bars{display:grid;gap:12px}.bar-row{display:grid;grid-template-columns:190px 1fr 46px;gap:12px;align-items:center}.bar-row span{height:18px;background:#111418;border-radius:999px;overflow:hidden;border:1px solid #343a43}.bar-row i{display:block;height:100%;background:linear-gradient(90deg,var(--yellow2),var(--yellow));border-radius:999px}.bar-row em{font-style:normal;color:#ffd23f;font-weight:1000}input[type=file]{max-width:100%;white-space:normal;overflow:hidden}.field{min-width:0}.card{min-width:0}@media(max-width:1000px){body{overflow-x:hidden}.app{overflow-x:hidden}.main{width:100%;overflow-x:hidden}.card{padding:15px}.form-grid{display:grid!important;grid-template-columns:1fr!important}.form-grid .field,.form-grid button,.form-grid a{grid-column:1!important;width:100%;min-width:0}.field input,.field select,.field textarea,input[type=file]{width:100%;max-width:100%;font-size:14px}.table-wrap{max-width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch}.table-wrap table{min-width:760px}.bar-row{grid-template-columns:1fr}.profile-row{align-items:flex-start}.mobile-head{height:54px}.hero{overflow:hidden}.detail-box{grid-column:span 12!important}}.btn-warn{background:linear-gradient(135deg,#ffce4a,#ff9f1c);border:0;color:#171a20;border-radius:14px;padding:10px 14px;font-weight:1000;cursor:pointer;box-shadow:0 10px 22px rgba(255,178,26,.18)}.btn-danger{background:linear-gradient(135deg,#48131f,#7f1d1d);border:1px solid #ef4444;color:#fee2e2;border-radius:14px;padding:10px 14px;font-weight:1000;cursor:pointer}.st-aprobado{background:#113327!important;border-color:#2dd4bf!important;color:#9fffe8!important}.st-rechazado{background:#3f1520!important;border-color:#ef4444!important;color:#fecaca!important}.st-firmado{background:#182844!important;border-color:#60a5fa!important;color:#bfdbfe!important}.st-aceptado{background:#302a12!important;border-color:#facc15!important;color:#fde68a!important}.row-approved{background:linear-gradient(90deg,rgba(45,212,191,.08),transparent)}.row-rejected{background:linear-gradient(90deg,rgba(239,68,68,.10),transparent)}.nested{margin:4px 0}.nested>.menu-item{width:100%;border:0}.submenu .menu-group{margin:6px 0 6px 22px}.submenu .menu-group .menu-title{padding:13px 14px;font-size:14px;border-radius:10px}.submenu .submenu .menu-item{padding-left:48px}.nested.closed .submenu{display:none}.menu-group.closed .submenu{display:none}.menu-group .chev{margin-left:auto}.menu-group.closed .chev{transform:rotate(-90deg)}
/* Marcador de ubicación limpio: no pinta blanco, solo línea y brillo lateral */
button.menu-item{font:inherit;text-align:left;background:transparent;border-top:0;border-right:0;border-bottom:0;appearance:none;-webkit-appearance:none;width:100%;}
.menu-item.active,.menu-item.parent-active{background:linear-gradient(90deg,rgba(255,210,63,.16),rgba(255,210,63,.035) 42%,transparent)!important;border-left-color:var(--yellow)!important;color:#fff!important;box-shadow:inset 4px 0 0 var(--yellow);}
.menu-item.sub-mini.active{background:rgba(255,210,63,.12)!important;color:#fff!important;border-left-color:var(--yellow)!important;box-shadow:inset 4px 0 0 var(--yellow);}
.menu-group.force-open>.menu-title{background:linear-gradient(135deg,var(--yellow2),var(--yellow));color:#181a1f;}
.nested.force-open>.menu-item.parent-active{background:linear-gradient(90deg,rgba(255,210,63,.14),rgba(255,210,63,.04),transparent)!important;color:#fff!important;}
.nested>.menu-item:focus,.nested>.menu-item:active{background:linear-gradient(90deg,rgba(255,210,63,.16),rgba(255,210,63,.035),transparent)!important;color:#fff!important;outline:none!important;}

.module-tabs{display:grid;grid-template-columns:repeat(3,minmax(220px,1fr));gap:16px}.module-tile{padding:22px;border:1px solid #3a414b;border-radius:18px;background:linear-gradient(145deg,#24272d,#191d23)}.module-tile h2{margin:0 0 8px}.badge-green{display:inline-flex;background:#49a916;color:#fff;border-radius:7px;padding:7px 12px;font-weight:1000}.badge-orange{display:inline-flex;background:#ff8f2d;color:#fff;border-radius:7px;padding:7px 12px;font-weight:1000}.adapta-note{background:#fff;color:#142033;border-radius:14px;padding:14px;border-left:5px solid #ff8f2d}.adapta-table table{background:#fff;color:#102033}.adapta-table th{background:#f4f5f7;color:#102033;text-transform:none;font-size:14px}.adapta-table td{border-bottom:1px solid #e4e7eb}.adapta-table tr:nth-child(even) td{background:#ededed}.adapta-table tr:hover td{background:#ffe9d6}@media(max-width:1000px){.module-tabs{grid-template-columns:1fr}}

/* === AJUSTE PRO: MENÚ UNIFORME + DASHBOARDS POR GESTIÓN === */
.menu-title,.menu-item{min-height:54px;box-sizing:border-box}
.submenu>.menu-item{padding:14px 18px 14px 36px;margin:4px 0}
.menu-group.nested .menu-title,.menu-group.nested>.menu-item{font-size:14px}
.gestion-cards{display:grid;grid-template-columns:repeat(3,minmax(260px,1fr));gap:18px;margin-bottom:18px}
.gestion-card{min-height:190px;display:flex;gap:18px;align-items:flex-start;position:relative;overflow:hidden}
.gestion-card:after{content:"";position:absolute;inset:auto -35px -35px auto;width:120px;height:120px;border-radius:999px;background:rgba(255,210,63,.10)}
.gestion-card.green{border-color:rgba(57,196,99,.40)}.gestion-card.purple{border-color:rgba(160,98,255,.42)}
.gestion-icon{width:64px;height:64px;border-radius:18px;display:grid;place-items:center;font-size:30px;background:linear-gradient(135deg,var(--yellow),var(--yellow2));color:#17191e;box-shadow:0 16px 30px rgba(0,0,0,.26)}
.gestion-card.green .gestion-icon{background:linear-gradient(135deg,#27b862,#6ee78f)}.gestion-card.purple .gestion-icon{background:linear-gradient(135deg,#7f43c7,#b77cff);color:#fff}
.dashboard-panel{grid-column:span 4}.dashboard-panel .mini-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-top:12px}.dash-metric{background:#171b21;border:1px solid #343a43;border-radius:14px;padding:14px;display:flex;justify-content:space-between;align-items:center}.dash-metric b{font-size:24px;color:var(--yellow)}
@media(max-width:1100px){.gestion-cards{grid-template-columns:1fr}.dashboard-panel{grid-column:span 12}.span-6{grid-column:span 12}}



/* === PRIZE PRO 2026: DASHBOARD ADMIN IGUAL A REFERENCIA === */
.app{grid-template-columns:270px minmax(0,1fr);background:radial-gradient(circle at 70% -10%,rgba(255,210,63,.08),transparent 26%),#0f141a!important}.app.side-collapsed{grid-template-columns:82px minmax(0,1fr)}.main{padding:20px 24px 32px!important;background:linear-gradient(180deg,#0f141a,#121820)!important}.side{width:270px!important;padding:8px 6px!important;background:linear-gradient(180deg,#111720,#171c23)!important;border-right:1px solid rgba(255,255,255,.08)}.side.collapsed{width:82px!important}.app.side-collapsed .side{width:82px!important}.side-top{display:flex!important;height:48px!important;background:#111720!important}.side-top .label{font-size:13px!important}.brand{display:none}.side-user{display:none}.menu-group{margin:8px 0!important}.menu-title,.menu-item{min-height:46px!important;border-radius:10px!important;font-size:14px!important;font-weight:950!important;padding:12px 16px!important;gap:12px!important}.submenu>.menu-item{padding:11px 18px 11px 36px!important;min-height:38px!important;font-size:13px!important;margin:2px 0!important}.submenu .menu-group{margin:6px 0 6px 10px!important}.submenu .menu-group .menu-title{min-height:40px!important;font-size:13px!important;padding:11px 14px!important}.menu-item.active,.menu-item.parent-active{border-left:4px solid var(--yellow)!important;background:linear-gradient(90deg,rgba(255,210,63,.16),rgba(255,210,63,.05))!important;box-shadow:inset 4px 0 0 var(--yellow)!important}.menu-group.force-open>.menu-title{background:linear-gradient(135deg,#ffb21a,#ffd23f)!important;color:#111820!important;box-shadow:0 12px 28px rgba(255,178,26,.25)!important}.menu-group.nested.force-open>.menu-title{background:rgba(255,210,63,.08)!important;color:#f8fafc!important;border-left:4px solid var(--yellow)!important;box-shadow:inset 4px 0 0 var(--yellow)!important}.menu-group.nested.force-open>.menu-title .chev{color:#fff}.admin-shell{max-width:1560px;margin:0 auto}.admin-header{display:flex;align-items:flex-start;justify-content:space-between;gap:20px;margin:0 0 18px}.admin-title h1{margin:0 0 4px;font-size:27px;line-height:1.15}.admin-title .role{font-size:17px;font-weight:1000;color:var(--yellow);margin-bottom:18px}.admin-title p{margin:0;color:#e5edf8;font-weight:750}.hambox{width:34px;height:34px;display:grid;place-items:center;border-radius:10px;background:#151c24;border:1px solid rgba(255,255,255,.05);margin-right:14px}.admin-title-row{display:flex;align-items:flex-start}.top-actions{display:flex;align-items:center;gap:13px}.top-icon{width:39px;height:39px;border-radius:10px;background:#151c24;border:1px solid rgba(255,255,255,.05);display:grid;place-items:center;position:relative;font-size:19px}.top-icon i{position:absolute;right:-4px;top:-8px;background:#ff4d5c;color:white;border-radius:999px;font-size:11px;min-width:20px;height:20px;display:grid;place-items:center;font-style:normal}.admin-chip{display:flex;align-items:center;gap:10px;font-weight:950}.admin-chip .a{width:38px;height:38px;border-radius:50%;background:linear-gradient(135deg,#ffb21a,#ffd23f);color:#111;display:grid;place-items:center}.gestion-cards{grid-template-columns:repeat(3,minmax(240px,1fr))!important;gap:16px!important;background:rgba(255,255,255,.025);border:1px solid rgba(255,255,255,.05);border-radius:18px;padding:18px}.gestion-card{min-height:160px!important;padding:26px!important;align-items:center!important}.gestion-card h2{font-size:20px;margin:0 0 12px}.gestion-card p{min-height:52px;line-height:1.5}.gestion-card .btn-warn,.gestion-card .btn-green,.gestion-card .btn-blue{margin-top:8px;min-width:150px;justify-content:center;border-radius:8px;padding:10px 15px}.gestion-card.green .btn-green{background:#163024!important;color:#43d96d!important;border:1px solid rgba(67,217,109,.4)!important}.gestion-card.purple .btn-blue{background:#241a33!important;color:#bc85ff!important;border:1px solid rgba(188,133,255,.45)!important}.gestion-icon{width:64px!important;height:64px!important;border-radius:14px!important}.dashboards-admin{display:grid;grid-template-columns:repeat(3,minmax(240px,1fr));gap:16px;margin-top:16px}.dashboard-panel{grid-column:auto!important;padding:22px!important}.dashboard-panel h2{font-size:16px;margin-bottom:18px}.dashboard-panel .mini-grid{grid-template-columns:repeat(2,1fr)!important;gap:14px!important}.dash-metric{min-height:86px!important;align-items:flex-start!important;position:relative;display:block!important}.dash-metric span{display:block;font-size:12px;color:#f2f6fb;margin-bottom:10px}.dash-metric b{font-size:20px!important;color:#fff!important}.dash-metric .mi{position:absolute;right:13px;bottom:13px;width:38px;height:38px;border-radius:9px;display:grid;place-items:center;background:linear-gradient(135deg,#ffb21a,#ffd23f);color:#101418}.dashboard-panel.green .dash-metric .mi{background:linear-gradient(135deg,#27b862,#6ee78f);color:#fff}.dashboard-panel.purple .dash-metric .mi{background:linear-gradient(135deg,#7f43c7,#b77cff);color:#fff}.dashboard-panel .full-link{margin-top:16px;width:100%;justify-content:space-between;border-radius:8px;padding:12px 16px}.dashboard-panel.green .full-link{background:#142c23!important;color:#44d96d!important;border:1px solid rgba(68,217,109,.38)!important}.dashboard-panel.purple .full-link{background:#241831!important;color:#bd86ff!important;border:1px solid rgba(189,134,255,.42)!important}.admin-footer{display:flex;justify-content:space-between;color:#aab4c1;font-size:12px;margin:24px 4px 0}.test-panel-hidden{margin-top:16px}.admin-section-title{font-size:18px;margin:0 0 14px}.card{background:linear-gradient(145deg,#181e26,#14191f)!important;border-color:#313946!important}@media(max-width:1200px){.dashboards-admin,.gestion-cards{grid-template-columns:1fr!important}.app{grid-template-columns:280px 1fr}.admin-header{flex-direction:column}.top-actions{align-self:flex-end}}@media(max-width:1000px){.main{padding-top:74px!important}.app{grid-template-columns:1fr!important}.side{padding-top:60px!important}.dashboards-admin,.gestion-cards{grid-template-columns:1fr!important}}

/* === MEJORA VISUAL VACACIONES: solicitud y mis solicitudes === */
.vac-request-card{padding:28px!important;border-radius:22px!important;overflow:hidden;position:relative}.vac-request-card:before{content:"";position:absolute;right:-80px;top:-90px;width:230px;height:230px;border-radius:50%;background:rgba(255,210,63,.10);pointer-events:none}.vac-head{display:flex;justify-content:space-between;gap:18px;align-items:flex-start;margin-bottom:20px}.vac-head h2{margin:0 0 8px;font-size:26px}.vac-help{color:#c2ccd8;font-weight:800;line-height:1.45;margin:0}.period-list{display:grid;grid-template-columns:repeat(3,minmax(230px,1fr));gap:14px;margin-top:10px}.period-card{position:relative;display:grid!important;grid-template-columns:auto 1fr;gap:12px;align-items:flex-start;margin:0!important;padding:18px!important;border:1px solid #34404d;border-radius:18px;background:linear-gradient(145deg,#171c23,#11161c);cursor:pointer;min-height:104px;box-shadow:0 12px 26px rgba(0,0,0,.16);transition:.15s}.period-card:hover{transform:translateY(-1px);border-color:rgba(255,210,63,.7);box-shadow:0 16px 35px rgba(0,0,0,.28)}.period-card input{width:18px;height:18px;accent-color:#ffd23f;margin-top:4px}.period-main{display:grid;gap:8px}.period-years{font-size:18px;color:#fff;font-weight:1000;letter-spacing:.2px}.period-meta{display:flex;flex-wrap:wrap;gap:8px}.period-badge{background:#202732;border:1px solid #3a4654;border-radius:999px;padding:6px 10px;color:#d9e5f2;font-size:12px;font-weight:950}.period-badge strong{color:var(--yellow)}.period-card input:disabled~.period-main{opacity:.45}.vac-form-row{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-top:18px}.vac-form-row .field{min-width:0}.vac-submit-row{display:grid;grid-template-columns:1fr 1fr;gap:16px;align-items:end;margin-top:16px}.check-card{height:100%;display:flex!important;align-items:center;gap:10px;padding:15px 16px!important;border:1px solid #34404d;border-radius:16px;background:#11161c;color:#eaf3ff;font-weight:950}.check-card input{accent-color:#ffd23f;width:18px;height:18px}.vac-submit-row .btn-green{height:52px;justify-content:center;font-size:15px}.sol-cards{display:grid;gap:14px}.sol-card{display:grid;grid-template-columns:1.12fr 1.55fr .55fr 1.65fr 1.2fr 1fr;gap:14px;align-items:center;background:linear-gradient(145deg,#171c23,#11161d);border:1px solid #33404d;border-radius:18px;padding:18px 20px;box-shadow:0 12px 30px rgba(0,0,0,.18)}.sol-card.head{background:#0b1015;color:var(--yellow);font-size:13px;font-weight:1000;text-transform:uppercase;box-shadow:none;border-radius:18px 18px 8px 8px}.sol-card:not(.head){border-left:5px solid var(--yellow)}.sol-card b{color:#fff;font-size:16px}.sol-card .dias b{display:inline-grid;place-items:center;min-width:46px;height:38px;border-radius:12px;background:rgba(255,210,63,.12);color:var(--yellow);border:1px solid rgba(255,210,63,.28)}.sol-card .coment{color:#dbe4ee;font-weight:850}.sol-empty{padding:22px;border:1px dashed #3b4552;border-radius:16px;color:#b8c0cb}.local-note{margin-top:10px;color:#b8c0cb;font-size:13px}@media(max-width:1100px){.period-list{grid-template-columns:1fr}.vac-form-row,.vac-submit-row{grid-template-columns:1fr}}@media(max-width:900px){.sol-card,.sol-card.head{grid-template-columns:1fr}.sol-card.head{display:none}.sol-card{gap:9px}.sol-card>div:before{content:attr(data-label);display:block;color:var(--yellow);font-size:11px;text-transform:uppercase;margin-bottom:3px}.vac-head{display:block}}
</style>
<script>
function side(){return document.querySelector('.side')}
function appShell(){return document.querySelector('.app')}
function saveSideScroll(){const s=side(); if(s){localStorage.setItem('sideScroll',s.scrollTop||0)}}
function restoreSideScroll(){const s=side(); if(s){s.scrollTop=parseInt(localStorage.getItem('sideScroll')||'0')}}
function toggleSide(){const s=side(), a=appShell(); if(!s)return; if(window.innerWidth<1000){s.classList.toggle('open')}else{const c=!s.classList.contains('collapsed'); s.classList.toggle('collapsed',c); if(a)a.classList.toggle('side-collapsed',c); localStorage.setItem('sideCollapsed',c?'1':'0')}}
function toggleGroup(id){const g=document.getElementById(id); if(!g)return; g.classList.toggle('closed'); localStorage.setItem('group_'+id,g.classList.contains('closed')?'1':'0')}
function initSide(){const s=side(), a=appShell(); if(!s)return; const c=localStorage.getItem('sideCollapsed')==='1' && window.innerWidth>=1000; s.classList.toggle('collapsed',c); if(a)a.classList.toggle('side-collapsed',c); document.querySelectorAll('.menu-group[data-group]').forEach(g=>{const id=g.id; const saved=localStorage.getItem('group_'+id); if(saved==='1' && !g.classList.contains('force-open')) g.classList.add('closed')}); if(!location.hash){setTimeout(restoreSideScroll,60)}; document.querySelectorAll('.menu-item').forEach(a=>a.addEventListener('click',()=>{saveSideScroll(); if(window.innerWidth<1000){const s=side(); if(s)s.classList.remove('open')}}));}
function filterCards(){const q=(document.getElementById('cardSearch')?.value||'').toLowerCase();document.querySelectorAll('.doc-card').forEach(c=>{c.style.display=c.innerText.toLowerCase().includes(q)?'block':'none'})}
window.addEventListener('DOMContentLoaded',()=>{initSide(); if(location.hash){document.querySelectorAll('.menu-item').forEach(x=>{if(x.getAttribute('href')&&x.getAttribute('href').endsWith(location.hash)) x.classList.add('active')}); setTimeout(()=>{document.querySelector(location.hash)?.scrollIntoView({block:'start'});},120)}});window.addEventListener('beforeunload',saveSideScroll)
</script></head><body>{{ body|safe }}</body></html>
'''


def render_page(content, title="Portal de Documentos PRIZE", active="Inicio"):
    user_label = session.get('admin_nombre') or session.get('nombre') or 'Usuario PRIZE'
    primer_nombre = user_label.split()[0] if user_label else 'Usuario'
    body = f'''
    <div class="mobile-head"><button class="toggle" onclick="toggleSide()">☰</button><b>PRIZE Documentos</b><a href="/logout">Salir</a></div>
    <div class="app"><aside class="side"><div class="side-top"><button class="toggle" title="Expandir / contraer panel" onclick="toggleSide()">☰</button><b class="label">PRIZE RRHH</b><button class="toggle" title="Expandir / contraer panel" onclick="toggleSide()">☰</button></div>
      <div class="brand"><img src="{logo_url()}" alt="PRIZE"><p>Documentos PRIZE</p></div>{sidebar(active)}<div class="side-user"><div class="avatar">👤</div><div><b>{primer_nombre}</b><br><small>{'Administrador' if session.get('admin_id') else 'Trabajador'}</small></div></div></aside><main class="main">{flashes()}{content}</main></div>'''
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
    active_type = str(active).split(':', 1)[0]
    cls = "menu-item active" if active_type == tipo else "menu-item"
    return f"<a class='{cls}' onclick='saveSideScroll()' href='{url_for('panel_tipo', tipo=tipo)}'><span>{icon}</span><span class='label'>{label}</span></a>"


def sidebar(active):
    active_txt = str(active or '')
    active_type = active_txt.split(':', 1)[0]
    active_sub = active_txt.split(':', 1)[1] if ':' in active_txt else ''
    pago_parts=[]
    for k,l,i in TIPOS_PAGO:
        if k=='Normal':
            sub_open = ' force-open' if active_type == k else ''
            base_cls = 'menu-item parent-active' if active_type == k else 'menu-item'
            cls_mensual = 'menu-item sub-mini active' if active_type == k and active_sub == 'Mensual' else 'menu-item sub-mini'
            cls_semanal = 'menu-item sub-mini active' if active_type == k and active_sub == 'Semanal' else 'menu-item sub-mini'
            pago_parts.append(f"<div id='grp_normal' data-group='normal' class='menu-group nested{sub_open}'><button type='button' class='{base_cls}' onclick=\"toggleGroup('grp_normal')\"><span>{i}</span><span class='label'>{l}</span><span class='chev'>∨</span></button><div class='submenu'>")
            pago_parts.append(f"<a class='{cls_mensual}' onclick='saveSideScroll()' href='{url_for('panel_tipo', tipo=k, sub='Mensual')}'><span>📅</span><span class='label'>Normal mensual</span></a>")
            pago_parts.append(f"<a class='{cls_semanal}' onclick='saveSideScroll()' href='{url_for('panel_tipo', tipo=k, sub='Semanal')}'><span>🗓️</span><span class='label'>Normal semanal</span></a></div></div>")
        else:
            pago_parts.append(item(k,l,i,active))
    pago = ''.join(pago_parts)
    emp = ''.join(item(k,l,i,active) for k,l,i in TIPOS_EMPRESA)
    per = ''.join(item(k,l,i,active) for k,l,i in TIPOS_PERSONALES)
    def gclass(keys):
        return 'menu-group force-open' if active_type in keys else 'menu-group'
    pago_cls = gclass([k for k,_,_ in TIPOS_PAGO])
    emp_cls = gclass([k for k,_,_ in TIPOS_EMPRESA])
    per_cls = gclass([k for k,_,_ in TIPOS_PERSONALES])
    admin = ""
    if session.get('admin_id'):
        admin_keys = ['Admin','Trabajadores','Usuarios','Modulo documentos','Subir documentos','Gestion Vacacional','Gestion Contratacion','Modo prueba'] + [k for k,_,_ in TIPOS_PAGO] + [k for k,_,_ in TIPOS_EMPRESA] + [k for k,_,_ in TIPOS_PERSONALES]
        admin_cls = 'menu-group force-open' if active_type in admin_keys else 'menu-group'
        cls_dash = 'menu-item active' if active == 'Admin' else 'menu-item'
        cls_trab = 'menu-item active' if active == 'Trabajadores' else 'menu-item'
        cls_docs = 'menu-item active' if active == 'Subir documentos' else 'menu-item'
        cls_users = 'menu-item active' if active == 'Usuarios' else 'menu-item'
        cls_moddocs = 'menu-item active' if active == 'Modulo documentos' else 'menu-item'
        cls_vac = 'menu-item active' if active == 'Gestion Vacacional' else 'menu-item'
        cls_con = 'menu-item active' if active == 'Gestion Contratacion' else 'menu-item'
        cls_test = 'menu-item active' if active == 'Modo prueba' else 'menu-item'
        docs_mod_keys = [k for k,_,_ in TIPOS_PAGO] + [k for k,_,_ in TIPOS_EMPRESA] + [k for k,_,_ in TIPOS_PERSONALES] + ['Modulo documentos','Subir documentos']
        docs_mod_cls = 'menu-group nested force-open' if active_type in docs_mod_keys else 'menu-group nested'
        vac_cls = 'menu-group nested force-open' if active == 'Gestion Vacacional' else 'menu-group nested'
        con_cls = 'menu-group nested force-open' if active == 'Gestion Contratacion' else 'menu-group nested'
        docs_head = 'menu-title' + (' active' if active_type in docs_mod_keys else '')
        vac_head = 'menu-title' + (' active' if active == 'Gestion Vacacional' else '')
        con_head = 'menu-title' + (' active' if active == 'Gestion Contratacion' else '')
        admin = f"""
        <div id='grp_admin' data-group='admin' class='{admin_cls}'>
          <button type='button' class='menu-title' onclick="toggleGroup('grp_admin')"><span>⚙️</span><span class='label'>Administrador</span><span class='chev'>∨</span></button>
          <div class='submenu'>
            <a class='{cls_dash}' onclick='saveSideScroll()' href='/admin'><span>📊</span><span class='label'>Dashboard</span></a>
            <div id='grp_modulo_documentos' data-group='modulo_documentos' class='{docs_mod_cls}'>
              <button type='button' class='{docs_head}' onclick="toggleGroup('grp_modulo_documentos')"><span>🗃️</span><span class='label'>1. Gestión Documental</span><span class='chev'>∨</span></button>
              <div class='submenu'>
                <a class='{cls_moddocs}' onclick='saveSideScroll()' href='/admin/modulo/documentos'><span>📊</span><span class='label'>Dashboard</span></a>
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
                <a class='{cls_docs}' onclick='saveSideScroll()' href='/admin/documentos'><span>⬆️</span><span class='label'>Subir / gestionar documentos</span></a>
              </div>
            </div>
            <div id='grp_vacacional' data-group='vacacional' class='{vac_cls}'>
              <button type='button' class='{vac_head}' onclick="toggleGroup('grp_vacacional')"><span>🏖️</span><span class='label'>2. Gestión Vacacional</span><span class='chev'>∨</span></button>
              <div class='submenu'>
                <a class='{cls_vac}' onclick='saveSideScroll()' href='/admin/vacaciones'><span>📊</span><span class='label'>Dashboard vacacional</span></a>
                <a class='menu-item' onclick='saveSideScroll()' href='/admin/vacaciones#cargar-saldos'><span>🗓️</span><span class='label'>Saldos Vacacionales</span></a>
                <a class='menu-item' onclick='saveSideScroll()' href='/admin/vacaciones#solicitudes'><span>📄</span><span class='label'>Solicitudes de Vacaciones</span></a>
                <a class='menu-item' onclick='saveSideScroll()' href='/admin/vacaciones#aprobaciones'><span>✅</span><span class='label'>Aprobaciones</span></a>
                <a class='menu-item' onclick='saveSideScroll()' href='/admin/vacaciones#reportes'><span>📑</span><span class='label'>Reportes</span></a>
              </div>
            </div>
            <div id='grp_contratacion' data-group='contratacion' class='{con_cls}'>
              <button type='button' class='{con_head}' onclick="toggleGroup('grp_contratacion')"><span>🧾</span><span class='label'>3. Gestión Contratación</span><span class='chev'>∨</span></button>
              <div class='submenu'>
                <a class='{cls_con}' onclick='saveSideScroll()' href='/admin/contratacion'><span>📊</span><span class='label'>Dashboard contratación</span></a>
                <a class='menu-item' onclick='saveSideScroll()' href='/admin/contratacion#procesos'><span>💼</span><span class='label'>Procesos de Contratación</span></a>
                <a class='menu-item' onclick='saveSideScroll()' href='/admin/contratacion#candidatos'><span>👥</span><span class='label'>Candidatos</span></a>
                <a class='menu-item' onclick='saveSideScroll()' href='/admin/contratacion#contratos'><span>📄</span><span class='label'>Contratos</span></a>
                <a class='menu-item' onclick='saveSideScroll()' href='/admin/contratacion#plantillas'><span>📋</span><span class='label'>Plantilla Documentos</span></a>
                <a class='menu-item' onclick='saveSideScroll()' href='/admin/contratacion#descargas'><span>⬇️</span><span class='label'>Descargas</span></a>
                <a class='menu-item' onclick='saveSideScroll()' href='/admin/contratacion#maestros'><span>🧩</span><span class='label'>Datos Maestros</span></a>
              </div>
            </div>
            <a class='{cls_trab}' onclick='saveSideScroll()' href='/admin/trabajadores'><span>👥</span><span class='label'>Trabajadores</span></a>
            <a class='{cls_users}' onclick='saveSideScroll()' href='/admin/usuarios'><span>🔐</span><span class='label'>Usuarios y claves</span></a>
            <a class='{cls_test}' onclick='saveSideScroll()' href='/admin/modo_prueba'><span>🧪</span><span class='label'>Modo prueba y limpieza</span></a>
          </div>
        </div>"""
    user_docs_keys = [k for k,_,_ in TIPOS_PAGO] + [k for k,_,_ in TIPOS_EMPRESA] + [k for k,_,_ in TIPOS_PERSONALES]
    user_docs_cls = 'menu-group force-open' if active_type in user_docs_keys else 'menu-group'
    documentos_generales = '' if session.get('admin_id') else f"""
      <div id='grp_user_documental' data-group='user_documental' class='{user_docs_cls}'>
        <button type='button' class='menu-title' onclick="toggleGroup('grp_user_documental')"><span>🗃️</span><span class='label'>Gestión Documental</span><span class='chev'>∨</span></button>
        <div class='submenu'><a class='menu-item' onclick='saveSideScroll()' href='/panel'><span>📊</span><span class='label'>Dashboard documental</span></a>
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
        </div>
      </div>"""
    user_gestiones = '' if session.get('admin_id') else f"""
      <div id='grp_user_vacacional' data-group='user_vacacional' class='menu-group {'force-open' if active == 'Gestion Vacacional' else ''}'>
        <button type='button' class='menu-title {'active' if active == 'Gestion Vacacional' else ''}' onclick="toggleGroup('grp_user_vacacional')"><span>🏖️</span><span class='label'>Gestión Vacacional</span><span class='chev'>∨</span></button>
        <div class='submenu'><a class='menu-item {'active' if active == 'Gestion Vacacional' else ''}' onclick='saveSideScroll()' href='/vacaciones/mi_solicitud'><span>📊</span><span class='label'>Dashboard vacacional</span></a><a class='menu-item {'active' if active == 'Gestion Vacacional' else ''}' onclick='saveSideScroll()' href='/vacaciones/mi_solicitud#solicitar'><span>🗓️</span><span class='label'>Saldo y solicitud</span></a><a class='menu-item {'active' if active == 'Gestion Vacacional' else ''}' onclick='saveSideScroll()' href='/vacaciones/aprobaciones_jefe'><span>✅</span><span class='label'>Aprobaciones jefe</span></a></div>
      </div>
      <div id='grp_user_contrato' data-group='user_contrato' class='menu-group {'force-open' if active == 'Gestion Contratacion' else ''}'>
        <button type='button' class='menu-title {'active' if active == 'Gestion Contratacion' else ''}' onclick="toggleGroup('grp_user_contrato')"><span>🧾</span><span class='label'>Gestión Contrato</span><span class='chev'>∨</span></button>
        <div class='submenu'><a class='menu-item {'active' if active == 'Gestion Contratacion' else ''}' onclick='saveSideScroll()' href='/contratacion/mis_documentos'><span>📊</span><span class='label'>Dashboard contrato</span></a><a class='menu-item {'active' if active == 'Gestion Contratacion' else ''}' onclick='saveSideScroll()' href='/contratacion/mis_documentos#mis-contratos'><span>📄</span><span class='label'>Mis contratos</span></a></div>
      </div>"""
    return f"""
    <nav>
      {documentos_generales}
      {user_gestiones}
      {admin}
      <div id='grp_cuenta' data-group='cuenta' class='menu-group {'force-open' if active == 'Inicio' else ''}'>
        <button type='button' class='menu-title {'active' if active == 'Inicio' else ''}' onclick="toggleGroup('grp_cuenta')"><span>👤</span><span class='label'>Mi cuenta</span><span class='chev'>∨</span></button>
        <div class='submenu'><a class='menu-item {'active' if active == 'Inicio' else ''}' onclick='saveSideScroll()' href='/panel'><span>🏠</span><span class='label'>Inicio</span></a><a class='menu-item' href='/logout'><span>🚪</span><span class='label'>Salir</span></a></div>
      </div>
    </nav>"""


def empresas_disponibles_login():
    """Empresas visibles en el login, tomadas de la columna EMPRESA de trabajadores activos."""
    empresas = []
    try:
        with db() as con:
            rows = con.execute("SELECT DISTINCT empresa FROM trabajadores WHERE activo=1 AND COALESCE(TRIM(empresa),'')<>'' ORDER BY empresa").fetchall()
        for r in rows:
            for raw in clean(r['empresa']).replace('|','/').replace(';','/').replace(',','/').split('/'):
                e = clean(raw)
                if not e:
                    continue
                if e.upper() == 'PRIZE SUPERFRUITS':
                    e = 'AQUANQA'
                if e not in empresas:
                    empresas.append(e)
    except Exception:
        pass
    return empresas or ['AQUANQA']

def login_template(admin=False, error=""):
    action = url_for('admin_login') if admin else url_for('login')
    title = "Administrador PRIZE RRHH" if admin else "Portal PRIZE RRHH"
    sub = "Control y trazabilidad documental" if admin else "Ingrese DNI, clave y empresa asignada"
    if admin:
        fields = """
          <div class='field'><label>Usuario</label><div class='login-input'>👤<input name='usuario' placeholder='Ingrese su usuario' required></div></div>
          <div class='field'><label>Clave</label><div class='login-input'>🔒<input name='clave' type='password' placeholder='Ingrese su clave' required></div></div>
        """
    else:
        opts = ''.join([f"<option value='{e}'>{e}</option>" for e in empresas_disponibles_login()])
        fields = f"""
          <div class='field'><label>Empresa</label><div class='login-input'>🏢<select name='empresa' required style='width:100%;background:transparent;color:#fff;border:0;padding:15px 8px;font-weight:900'><option value=''>Seleccione empresa</option>{opts}</select></div></div>
          <div class='field'><label>DNI</label><div class='login-input'>🪪<input name='dni' maxlength='8' placeholder='Ingrese su DNI' required></div></div>
          <div class='field'><label>Clave</label><div class='login-input'>🔑<input name='correo' type='password' placeholder='Fecha nacimiento: ddmmaaaa' required></div></div>
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
    # Logo genérico PRIZE RRHH: sin texto fijo de empresa.
    svg = """<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 360 140'><defs><linearGradient id='g' x1='0' x2='1'><stop offset='0' stop-color='#ffd23f'/><stop offset='1' stop-color='#ff9f1c'/></linearGradient></defs><rect width='360' height='140' rx='26' fill='white'/><circle cx='63' cy='68' r='34' fill='url(#g)'/><text x='105' y='78' font-family='Segoe UI,Arial' font-size='58' font-style='italic' font-weight='900' fill='#111827'>PRIZE</text><text x='112' y='112' font-family='Arial' font-size='24' font-weight='900' fill='#2b668d'>RRHH</text></svg>"""
    return app.response_class(svg, mimetype='image/svg+xml')

# =============================
# LOGIN USUARIO
# =============================
@app.route('/', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        dni = normalizar_dni(request.form.get('dni'))
        clave = clean(request.form.get('correo')).lower()
        bloqueado, intentos_previos = esta_bloqueado(dni)
        if bloqueado:
            return login_template(False, "Usuario bloqueado por 3 intentos fallidos. Solicite desbloqueo al administrador.")
        t = get_trabajador(dni)
        clave_correcta = generar_clave_trabajador(dni, t['fecha_nacimiento'] if t and 'fecha_nacimiento' in t.keys() else '').lower() if t else ''
        clave_guardada = clean(t['clave_portal'] if t and 'clave_portal' in t.keys() else '').lower()
        if not t or not int(t['activo'] or 0) or (clave != clave_correcta and clave != clave_guardada):
            n, b = registrar_intento_fallido(dni)
            if b:
                return login_template(False, "Usuario bloqueado por 3 intentos fallidos. Solicite desbloqueo al administrador.")
            return login_template(False, f"DNI o clave incorrecta. Clave = fecha nacimiento sin / (ddmmaaaa). Intento {n}/3.")
        empresa_login = clean(request.form.get('empresa'))
        empresas_permitidas = []
        for raw in clean(t['empresa'] if t and 'empresa' in t.keys() else '').replace('|','/').replace(';','/').replace(',','/').split('/'):
            e = clean(raw)
            if e.upper() == 'PRIZE SUPERFRUITS':
                e = 'AQUANQA'
            if e and e not in empresas_permitidas:
                empresas_permitidas.append(e)
        if not empresas_permitidas:
            empresas_permitidas = ['AQUANQA']
        if empresa_login not in empresas_permitidas:
            registrar_intento_fallido(dni)
            return login_template(False, 'La empresa seleccionada no pertenece al trabajador. Seleccione la empresa registrada en la columna EMPRESA.')
        reset_intentos_login(dni)
        session.clear(); session['dni'] = dni; session['nombre'] = t['nombre']; session['empresa'] = empresa_login
        return redirect(url_for('panel'))
    return login_template(False)

@app.route('/logout')
def logout():
    session.clear(); return redirect(url_for('login'))


@app.route('/seleccionar_empresa', methods=['GET','POST'])
@worker_required
def seleccionar_empresa():
    dni=session['dni']; t=get_trabajador(dni)
    emp_real = clean(t['empresa']) if t and 'empresa' in t.keys() else ''
    # Solo mostrar la(s) empresa(s) que vienen desde la columna EMPRESA del trabajador.
    # Se elimina cualquier opción fija como PRIZE SUPERFRUITS.
    empresas=[]
    for raw in emp_real.replace('|','/').replace(';','/').replace(',','/').split('/'):
        e=clean(raw)
        if not e: continue
        # corrección defensiva para bases antiguas/demostración
        if e.upper() == 'PRIZE SUPERFRUITS':
            e = 'AQUANQA'
        if e not in empresas:
            empresas.append(e)
    if not empresas:
        empresas=['AQUANQA']
    if request.method=='POST':
        emp=clean(request.form.get('empresa'))
        if emp not in empresas: emp=empresas[0]
        session['empresa']=emp
        flash(f'Empresa seleccionada: {emp}', 'ok')
        return redirect(url_for('panel'))
    opts=''.join([f"<option value='{e}'>{e}</option>" for e in empresas])
    content=f"""<div class='login-body'><form class='login-card' method='post'><div class='login-inner'>
      <div class='login-logo'><img src='{logo_url()}'></div><div class='login-title'><h1>Elegir empresa</h1><b>{t['nombre'] if t else dni}</b></div>
      <div class='field'><label>Empresa asignada</label><div class='login-input'>🏢<select name='empresa' style='width:100%;background:transparent;color:#fff;border:0;padding:15px;font-weight:900'>{opts}</select></div></div>
      <button class='btn-green'>Ingresar al portal</button></div></form></div>"""
    return render_template_string(BASE, body=content, title='Elegir empresa')

@app.route('/panel')
@worker_required
def panel():
    dni = session['dni']; sincronizar_documentos_carpeta(dni); t = get_trabajador(dni)
    if not t:
        flash('No se encontró tu trabajador activo. Vuelve a iniciar sesión o contacta a RRHH.', 'err')
        return redirect(url_for('logout'))
    docs = listar_documentos(dni=dni, limit=999)
    ultimo = docs[0]['tipo'] if docs else 'Sin documento'
    cards = ''.join(doc_card(k,l,i) for k,l,i in (TIPOS_PAGO+TIPOS_EMPRESA+TIPOS_PERSONALES))
    dashboard_gestiones = """
      <div class='card span-12'><h2>Dashboards de gestión</h2><div class='module-tabs'>
        <a class='module-tile' href='/panel'><h2>📁 Gestión Documental</h2><p class='muted'>Documentos, pagos, empresa y personales.</p></a>
        <a class='module-tile' href='/vacaciones/mi_solicitud'><h2>🏖️ Gestión Vacacional</h2><p class='muted'>Saldo, solicitud y seguimiento de aprobaciones.</p></a>
        <a class='module-tile' href='/contratacion/mis_documentos'><h2>🧾 Gestión Contrato</h2><p class='muted'>Contratos, anexos y documentos laborales.</p></a>
      </div></div>
    """
    content = f"""
    <div class='hero'><div class='topbar'><div><h1>Portal Documental <span class='accent'>PRIZE</span></h1><div class='subtitle'>{t['nombre']} · DNI {t['dni']} · {session.get('empresa') or t['empresa']}</div></div><div style='display:flex;gap:10px;align-items:center'><span class='btn'>● Activo</span><a class='btn-blue' href='/panel'>Ver todo</a></div></div></div>
    <section class='grid'><div class='card mini'><div><span>Documentos</span><br><b>{len(docs)}</b></div><div class='ico'>🗂️</div></div><div class='card mini'><div><span>Último tipo</span><br><b>{ultimo}</b></div><div class='ico'>📄</div></div><div class='card mini'><div><span>Estado</span><br><b>Activo</b></div><div class='ico'>✅</div></div>{dashboard_gestiones}<div class='card span-12 profile-card'><div><h2>Mi perfil y foto</h2><p class='muted'>Actualiza tu foto para que el portal quede como panel profesional.</p></div><div class='profile-row'><img class='profile-img' src='{url_for('foto_trabajador', dni=dni) if t['foto_ruta'] else logo_url()}'><form method='post' action='/mi_foto' enctype='multipart/form-data' class='form-grid profile-form'><div class='field'><label>Foto personal</label><input type='file' name='foto' accept='.png,.jpg,.jpeg,.webp' required></div><button class='btn-green'>Cargar foto</button></form></div></div>
    <div class='card span-12'><div style='display:flex;justify-content:space-between;gap:12px;align-items:center;flex-wrap:wrap'><h2>Accesos por pestaña</h2><input id='cardSearch' onkeyup='filterCards()' class='input' style='max-width:310px' placeholder='Buscar pestaña...'></div><div class='doc-grid'>{cards}</div></div>
    <div class='card span-12'><h2>🔔 Notificaciones</h2>{notificaciones_trabajador(dni)}</div><div class='card span-12'><h2>Últimos documentos</h2>{tabla_docs(docs)}</div></section>"""
    return render_page(content, active='Inicio')


def notificaciones_trabajador(dni):
    with db() as con:
        rows = con.execute("SELECT evento,fecha,detalle FROM eventos_documento WHERE dni=? ORDER BY id DESC LIMIT 10", (normalizar_dni(dni),)).fetchall()
    if not rows:
        return "<p class='muted'>Sin notificaciones por ahora.</p>"
    return "".join([f"<div class='alert-item'><div class='bell'>🔔</div><div><b>{r['evento']}</b><br><span>{r['fecha']} · {r['detalle'] or ''}</span></div></div>" for r in rows])

def doc_card(k,l,i):
    return f"<div class='doc-card'><h3>{i} {l}</h3><p>Consulta, filtra por periodo y revisa el detalle del documento.</p><a class='btn-blue' href='{url_for('panel_tipo', tipo=k)}'>Abrir</a></div>"

@app.route('/documentos/<tipo>')
@portal_required
def panel_tipo(tipo):
    if tipo not in ALL_TIPOS: abort(404)
    asegurar_carpetas_documentales(tipo)
    is_admin = bool(session.get('admin_id'))
    dni = session.get('dni', '')
    if is_admin:
        sincronizar_documentos_carpeta()
    else:
        sincronizar_documentos_carpeta(dni)
    label, icon, categoria = ALL_TIPOS[tipo]
    periodo = clean(request.args.get('periodo'))
    sub = clean(request.args.get('sub'))
    pers = periodos_disponibles(dni=None if is_admin else dni, tipo=tipo) if is_admin else periodos_desde_ingreso(dni, tipo)
    docs = listar_documentos(dni=None if is_admin else dni, tipo=tipo, periodo=periodo or None, limit=999)
    if tipo == 'Normal' and sub:
        docs = [d for d in docs if sub.lower() in clean(d['detalle']).lower()]
    opts = "<option value=''>Todos los periodos</option>" + ''.join([f"<option {'selected' if p==periodo else ''}>{p}</option>" for p in pers])
    detalle = detalle_tipo(tipo, docs)
    upload_extra = ""
    if tipo in ['Otros','Contrato Personal'] and not is_admin:
        upload_extra = f"""
        <div class='card span-12'><h2>Adjuntar nuevo documento personal</h2><form method='post' action='/subir_personal' enctype='multipart/form-data' class='form-grid'>
        <input type='hidden' name='tipo' value='{tipo}'><div class='field'><label>Periodo</label><input name='periodo' value='{datetime.now(APP_TZ).strftime('%Y-%m')}'></div><div class='field'><label>Detalle</label><input name='detalle' placeholder='Ej: Certificado, solicitud, evidencia'></div><div class='field'><label>Archivo</label><input type='file' name='archivo' accept='.pdf,.png,.jpg,.jpeg,.webp,.doc,.docx,.xls,.xlsx' required></div><div class='field'><label>Observación</label><textarea name='observacion' rows='2' placeholder='Comentario u observación'></textarea></div><button class='btn-green'>Subir documento</button></form></div>"""
    content = f"""
    <div class='hero'><div class='topbar'><div><h1>{icon} {label}</h1><div class='subtitle'>Consulta filtrada por pestaña y periodo seleccionado.</div></div><a class='btn-blue' href='{url_for('admin') if is_admin else url_for('panel')}'>Volver</a></div></div>
    <section class='grid'><div class='card mini'><div><span>Total</span><br><b>{len(docs)}</b></div><div class='ico'>{icon}</div></div><div class='card mini'><div><span>Periodo</span><br><b>{periodo or 'Todos'}</b></div><div class='ico'>📅</div></div><div class='card mini'><div><span>Filtro</span><br><b>{tipo}{' - '+sub if sub else ''}</b></div><div class='ico'>🔎</div></div>
    <div class='card span-12'><form method='get' class='period-row'><div class='field'><label>Elegir periodo</label><select name='periodo'>{opts}</select></div><button class='btn-blue'>Aplicar filtro</button><a class='btn' href='{url_for('panel_tipo', tipo=tipo)}'>Limpiar</a></form></div>
    <div class='card span-12'><h2>Detalle de {label}</h2>{detalle}</div>{upload_extra}<div class='card span-12'><h2>Listado</h2>{tabla_docs(docs)}</div></section>"""
    return render_page(content, active=(f'{tipo}:{sub}' if tipo == 'Normal' and sub else tipo))

@app.route('/subir_personal', methods=['POST'])
@worker_required
def subir_personal():
    try:
        guardar_documento(request.files.get('archivo'), session['dni'], clean(request.form.get('tipo')) or 'Otros', request.form.get('periodo'), request.form.get('detalle'), request.form.get('observacion'), session.get('dni'))
        flash('Documento personal subido correctamente.', 'ok')
    except Exception as e:
        flash(f'No se pudo subir: {e}', 'error')
    return redirect(url_for('panel_tipo', tipo=clean(request.form.get('tipo')) or 'Otros'))


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
    headers = "<tr><th>Tipo</th><th>DNI</th><th>Trabajador</th><th>Periodo</th><th>Detalle</th><th>Observación</th><th>Estado</th><th>Cargado por</th><th>Fecha</th><th>Archivo</th><th>Acciones</th></tr>"
    if not rows:
        return f"<div class='table-wrap'><table>{headers}<tr><td colspan='11'>No hay documentos en esta pestaña.</td></tr></table></div>"
    body = ''
    is_admin = bool(session.get('admin_id'))
    dni_sess = session.get('dni')
    with db() as con:
        nombres = {r['dni']: r['nombre'] for r in con.execute("SELECT dni,nombre FROM trabajadores").fetchall()}
    for r in rows:
        rid = r['id']; estado = r['estado'] if 'estado' in r.keys() and r['estado'] else 'Pendiente'
        row_cls = 'row-approved' if estado == 'Aprobado' else ('row-rejected' if estado == 'Rechazado' else '')
        pill_cls = 'status-pill ' + ('st-aprobado' if estado == 'Aprobado' else 'st-rechazado' if estado == 'Rechazado' else 'st-firmado' if estado == 'Firmado' else 'st-aceptado' if estado == 'Aceptado' else '')
        ver = f"<a class='btn-blue' target='_blank' href='{url_for('ver_doc', doc_id=rid)}'>Ver/Descargar</a>"
        acciones = []
        # Trabajador: puede aceptar/rechazar/firmar, y solo eliminar documentos personales propios.
        if r['categoria'] in ['pago','empresa','personal'] and dni_sess and (r['dni'] == dni_sess or r['categoria']=='empresa'):
            if estado not in ['Aceptado','Firmado','Aprobado','Rechazado']:
                acciones.append(f"<a class='btn-green mini-btn' href='{url_for('flujo_doc', doc_id=rid, accion='aceptar')}'>Aceptar</a>")
                acciones.append(f"<button class='btn-danger mini-btn' onclick=\"showReject({rid})\">Rechazar</button>")
            if estado in ['Aceptado','Firmado']:
                acciones.append(f"<a class='btn-blue mini-btn' href='{url_for('flujo_doc', doc_id=rid, accion='firmar')}'>Firmar</a>")
        # Administrador: aprueba o rechaza. Eliminar queda disponible solo para admin o personal propio del trabajador.
        if is_admin and r['categoria'] in ['pago','personal','empresa']:
            if estado != 'Aprobado':
                acciones.append(f"<a class='btn-warn mini-btn' href='{url_for('flujo_doc', doc_id=rid, accion='aprobar')}'>Aprobar</a>")
            if estado != 'Rechazado':
                acciones.append(f"<button class='btn-danger mini-btn' onclick=\"showReject({rid})\">Rechazar</button>")
        if is_admin or (dni_sess and r['dni'] == dni_sess and r['categoria'] == 'personal'):
            acciones.append(f"<a class='btn-red mini-btn' onclick=\"return confirm('¿Eliminar este documento?')\" href='{url_for('eliminar_doc', doc_id=rid)}'>Eliminar</a>")
        dni_val = r['dni'] or 'EMPRESA'
        trabajador = nombres.get(r['dni'], 'Documento general' if r['categoria']=='empresa' else '-')
        cargado_por = r['uploaded_by'] or 'sistema'
        if cargado_por == 'auto': cargado_por = 'Carpeta automática'
        body += f"<tr class='{row_cls}'><td>{r['tipo']}</td><td>{dni_val}</td><td>{trabajador}</td><td>{r['periodo'] or ''}</td><td>{r['detalle'] or '-'}</td><td>{r['observacion'] or '-'}</td><td><span class='{pill_cls}'>{estado}</span></td><td><b>{cargado_por}</b></td><td>{r['fecha_subida']}</td><td>{ver}</td><td><div class='actions'>{''.join(acciones) or '-'}</div></td></tr>"
    modal = """<div id='rejectBox' class='modal hidden'><form method='post' id='rejectForm' class='modal-card'><h2>Rechazar documento</h2><label>Comentario obligatorio</label><textarea name='comentario' required rows='4' placeholder='Indique el motivo del rechazo'></textarea><div class='actions'><button class='btn-red'>Rechazar</button><button type='button' class='btn' onclick='hideReject()'>Cancelar</button></div></form></div><script>function showReject(id){let m=document.getElementById('rejectBox'),f=document.getElementById('rejectForm');f.action='/documento/'+id+'/rechazar';m.classList.remove('hidden')}function hideReject(){document.getElementById('rejectBox').classList.add('hidden')}</script>"""
    return f"<div class='table-wrap'><table><thead>{headers}</thead><tbody>{body}</tbody></table></div>{modal}"

@app.route('/documento/<int:doc_id>/eliminar')
def eliminar_doc(doc_id):
    with db() as con:
        r = con.execute("SELECT * FROM documentos WHERE id=?", (doc_id,)).fetchone()
        if not r: abort(404)
        dni_sess = session.get('dni')
        permitido = bool(session.get('admin_id')) or (dni_sess and r['dni'] == dni_sess and r['categoria'] == 'personal')
        if not permitido: abort(403)
        try:
            path = Path(r['ruta_archivo'])
            if path.exists(): path.unlink()
        except Exception:
            pass
        con.execute("DELETE FROM documentos WHERE id=?", (doc_id,))
        con.commit()
    flash('Documento eliminado correctamente.', 'ok')
    return redirect(request.referrer or url_for('panel'))

@app.route('/documento/<int:doc_id>/<accion>')
def flujo_doc(doc_id, accion):
    if accion not in ['aceptar','firmar','aprobar']: abort(404)
    with db() as con:
        r = con.execute("SELECT * FROM documentos WHERE id=?", (doc_id,)).fetchone()
        if not r: abort(404)
        dni_sess = session.get('dni')
        if accion in ['aceptar','firmar'] and (not dni_sess or (r['categoria']!='empresa' and r['dni'] != dni_sess)):
            abort(403)
        if accion == 'aprobar' and not session.get('admin_id'):
            abort(403)
        if accion == 'aceptar':
            con.execute("UPDATE documentos SET estado='Aceptado', fecha_aceptacion=? WHERE id=?", (now_txt(), doc_id))
            flash('Documento aceptado. Ahora puede firmarlo.', 'ok')
        elif accion == 'firmar':
            con.execute("UPDATE documentos SET estado='Firmado', fecha_firma=? WHERE id=?", (now_txt(), doc_id))
            flash('Documento firmado correctamente.', 'ok')
        elif accion == 'aprobar':
            con.execute("UPDATE documentos SET estado='Aprobado', fecha_aprobacion=? WHERE id=?", (now_txt(), doc_id))
            con.execute("INSERT INTO eventos_documento(documento_id,dni,evento,fecha,detalle) VALUES(?,?,?,?,?)", (doc_id, r['dni'] or '', 'Aprobado', now_txt(), 'Aprobado por administrador'))
            flash('Documento aprobado por administrador.', 'ok')
        con.commit()
    return redirect(request.referrer or url_for('panel'))

@app.route('/documento/<int:doc_id>/rechazar', methods=['POST'])
def rechazar_doc(doc_id):
    comentario = clean(request.form.get('comentario'))
    with db() as con:
        r = con.execute("SELECT * FROM documentos WHERE id=?", (doc_id,)).fetchone()
        if not r: abort(404)
        dni_sess = session.get('dni')
        is_admin = bool(session.get('admin_id'))
        if not is_admin and (not dni_sess or (r['categoria']!='empresa' and r['dni'] != dni_sess)): abort(403)
        con.execute("UPDATE documentos SET estado='Rechazado', comentario_rechazo=?, observacion=? WHERE id=?", (comentario, comentario, doc_id))
        con.execute("INSERT INTO eventos_documento(documento_id,dni,evento,fecha,detalle) VALUES(?,?,?,?,?)", (doc_id, r['dni'] or dni_sess or '', 'Rechazado', now_txt(), comentario))
        con.commit()
    flash('Documento rechazado. Se registró notificación para el trabajador.', 'ok')
    return redirect(request.referrer or url_for('panel'))

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
    if session.get('dni'):
        try: registrar_evento_documento(doc_id, session.get('dni'), 'Abierto', 'Trabajador abrió/descargó el documento')
        except Exception: pass
    return send_file(path, as_attachment=False, download_name=r['archivo_nombre'])



def alertas_admin(limit=8):
    with db() as con:
        return con.execute("""
            SELECT d.*, t.nombre AS trabajador
            FROM documentos d
            LEFT JOIN trabajadores t ON t.dni=d.dni
            ORDER BY d.id DESC
            LIMIT ?
        """, (limit,)).fetchall()

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
    sincronizar_documentos_carpeta()
    desde = clean(request.args.get('desde'))
    hasta = clean(request.args.get('hasta'))
    with db() as con:
        trabajadores = con.execute("SELECT COUNT(*) FROM trabajadores").fetchone()[0]
        docs = con.execute("SELECT COUNT(*) FROM documentos").fetchone()[0]
        emp = con.execute("SELECT COUNT(*) FROM documentos WHERE categoria='empresa'").fetchone()[0]
        leidos = con.execute("SELECT COUNT(*) FROM documentos WHERE fecha_lectura IS NOT NULL AND fecha_lectura<>''").fetchone()[0]
        aprobados = con.execute("SELECT COUNT(*) FROM documentos WHERE estado='Aprobado'").fetchone()[0]
        rechazados = con.execute("SELECT COUNT(*) FROM documentos WHERE estado='Rechazado'").fetchone()[0]
        ult = con.execute("SELECT * FROM documentos ORDER BY id DESC LIMIT 12").fetchall()
    alerts = alertas_admin(8)
    with db() as con:
        chart_rows = con.execute("SELECT tipo, COUNT(*) c FROM documentos GROUP BY tipo ORDER BY c DESC LIMIT 8").fetchall()
        fechas_docs = con.execute("SELECT fecha_subida FROM documentos").fetchall()
    hoy_dt = datetime.now(APP_TZ).date()
    doc_dia = doc_semana = doc_mes = doc_rango = 0
    desde_dt = parse_fecha_any(desde) if desde else None
    hasta_dt = parse_fecha_any(hasta) if hasta else None
    for rr in fechas_docs:
        try:
            dd = datetime.strptime((rr['fecha_subida'] or '')[:10], '%d/%m/%Y').date()
            if dd == hoy_dt: doc_dia += 1
            if (hoy_dt - dd).days <= 7: doc_semana += 1
            if dd.year == hoy_dt.year and dd.month == hoy_dt.month: doc_mes += 1
            if (not desde_dt or dd >= desde_dt) and (not hasta_dt or dd <= hasta_dt): doc_rango += 1
        except Exception:
            pass
    maxc = max([x['c'] for x in chart_rows] or [1])
    chart_html = ''.join([f"<div class='bar-row'><b>{x['tipo']}</b><span><i style='width:{max(6, int(x['c']*100/maxc))}%'></i></span><em>{x['c']}</em></div>" for x in chart_rows]) or "<p class='muted'>Sin información para graficar.</p>"
    alert_items = ''.join([f"<div class='alert-item'><div class='bell'>🔔</div><div><b>{(a['trabajador'] or a['dni'] or 'Documento empresa')}</b><br><span>{a['tipo']} · {a['periodo'] or 'Sin periodo'} · {a['fecha_subida']} · Cargado por: {a['uploaded_by'] or 'sistema'}</span></div><a class='btn-blue mini-btn' target='_blank' href='{url_for('ver_doc', doc_id=a['id'])}'>Ver</a></div>" for a in alerts]) or "<div class='empty-note'>Aún no hay documentos cargados.</div>"
    with db() as con:
        ind_rows = con.execute("""
            SELECT tipo,
                   COUNT(*) total,
                   SUM(CASE WHEN estado='Aprobado' THEN 1 ELSE 0 END) aprobados,
                   SUM(CASE WHEN estado='Rechazado' THEN 1 ELSE 0 END) rechazados,
                   SUM(CASE WHEN fecha_lectura IS NOT NULL AND fecha_lectura<>'' THEN 1 ELSE 0 END) leidos
            FROM documentos
            GROUP BY tipo
            ORDER BY tipo
        """).fetchall()
    ind_html = ''.join([f"<tr><td>{r['tipo']}</td><td>{r['total']}</td><td><span class='status-pill st-aprobado'>{r['aprobados']}</span></td><td><span class='status-pill st-rechazado'>{r['rechazados']}</span></td><td><span class='status-pill'>{r['leidos']}</span></td></tr>" for r in ind_rows]) or "<tr><td colspan='5'>Sin documentos.</td></tr>"
    modo_txt = 'ACTIVO' if modo_prueba_activo() else 'INACTIVO'
    with db() as con:
        vac_saldos = con.execute("SELECT COUNT(*) FROM vacaciones_saldos").fetchone()[0]
        vac_solicitudes = con.execute("SELECT COUNT(*) FROM vacaciones_solicitudes").fetchone()[0]
        vac_pendientes = con.execute("SELECT COUNT(*) FROM vacaciones_solicitudes WHERE estado LIKE 'Pendiente%'").fetchone()[0]
        vac_aprobadas = con.execute("SELECT COUNT(*) FROM vacaciones_solicitudes WHERE estado LIKE 'Aprobado%'").fetchone()[0]
        con_docs = con.execute("SELECT COUNT(*) FROM contratacion_docs").fetchone()[0]
        con_tipos = con.execute("SELECT COUNT(*) FROM contratacion_tipos").fetchone()[0]
    content = f"""
    <div class='admin-shell'>
      <div class='admin-header'>
        <div class='admin-title-row'>
          <button class='hambox' onclick='toggleSide()'>☰</button>
          <div class='admin-title'>
            <h1>Centro de Control</h1>
            <div class='role'>Administrador</div>
            <p>Bienvenido al panel de administración. Seleccione una gestión para comenzar.</p>
          </div>
        </div>
        <div class='top-actions'>
          <div class='top-icon'>🔔<i>0</i></div>
          <div class='top-icon'>☰<i>0</i></div>
          <div class='admin-chip'><span class='a'>A</span><span>Administrador⌄</span></div>
        </div>
      </div>

      <div class='gestion-cards'>
        <div class='card gestion-card'>
          <div class='gestion-icon'>📁</div>
          <div><h2>Gestión Documental</h2><p class='muted'>Administre y controle todos los documentos de la organización.</p><a class='btn-warn' href='/admin/modulo/documentos'>Ir al Dashboard <span>→</span></a></div>
        </div>
        <div class='card gestion-card green'>
          <div class='gestion-icon'>☂️</div>
          <div><h2>Gestión Vacacional</h2><p class='muted'>Administre saldos y solicitudes de vacaciones de los trabajadores.</p><a class='btn-green' href='/admin/vacaciones'>Ir al Dashboard <span>→</span></a></div>
        </div>
        <div class='card gestion-card purple'>
          <div class='gestion-icon'>📄</div>
          <div><h2>Gestión Contratación</h2><p class='muted'>Gestione procesos de contratación, candidatos y contratos.</p><a class='btn-blue' href='/admin/contratacion'>Ir al Dashboard <span>→</span></a></div>
        </div>
      </div>

      <div class='dashboards-admin'>
        <div class='card dashboard-panel'>
          <h2>📁 Dashboard - Gestión Documental</h2>
          <div class='mini-grid'>
            <div class='dash-metric'><span>Trabajadores</span><b>{trabajadores}</b><em class='mi'>👥</em></div>
            <div class='dash-metric'><span>Documentos</span><b>{docs}</b><em class='mi'>📄</em></div>
            <div class='dash-metric'><span>Recibidos / Abiertos</span><b>{leidos}</b><em class='mi'>👁️</em></div>
            <div class='dash-metric'><span>Aprobados</span><b>{aprobados}</b><em class='mi'>✓</em></div>
            <div class='dash-metric'><span>Rechazados</span><b>{rechazados}</b><em class='mi'>−</em></div>
            <div class='dash-metric'><span>Empresas</span><b>{emp}</b><em class='mi'>🏢</em></div>
          </div>
          <a class='btn-warn full-link' href='/admin/modulo/documentos'>Ver Dashboard Completo <span>→</span></a>
        </div>
        <div class='card dashboard-panel green'>
          <h2>☂️ Dashboard - Gestión Vacacional</h2>
          <div class='mini-grid'>
            <div class='dash-metric'><span>Saldos Registrados</span><b>{vac_saldos}</b><em class='mi'>🗓️</em></div>
            <div class='dash-metric'><span>Solicitudes</span><b>{vac_solicitudes}</b><em class='mi'>📄</em></div>
            <div class='dash-metric'><span>Pendientes</span><b>{vac_pendientes}</b><em class='mi'>⏱️</em></div>
            <div class='dash-metric'><span>Aprobadas</span><b>{vac_aprobadas}</b><em class='mi'>✓</em></div>
            <div class='dash-metric'><span>Rechazadas</span><b>0</b><em class='mi'>−</em></div>
            <div class='dash-metric'><span>En Proceso</span><b>{vac_pendientes}</b><em class='mi'>…</em></div>
          </div>
          <a class='btn-green full-link' href='/admin/vacaciones'>Ver Dashboard Completo <span>→</span></a>
        </div>
        <div class='card dashboard-panel purple'>
          <h2>🧾 Dashboard - Gestión Contratación</h2>
          <div class='mini-grid'>
            <div class='dash-metric'><span>Procesos Activos</span><b>{con_docs}</b><em class='mi'>💼</em></div>
            <div class='dash-metric'><span>Candidatos</span><b>0</b><em class='mi'>👥</em></div>
            <div class='dash-metric'><span>Contratos Activos</span><b>{con_docs}</b><em class='mi'>📄</em></div>
            <div class='dash-metric'><span>Contratos por Vencer</span><b>0</b><em class='mi'>⏱️</em></div>
            <div class='dash-metric'><span>Contratos Vencidos</span><b>0</b><em class='mi'>−</em></div>
            <div class='dash-metric'><span>Contratos Finalizados</span><b>0</b><em class='mi'>✓</em></div>
          </div>
          <a class='btn-blue full-link' href='/admin/contratacion'>Ver Dashboard Completo <span>→</span></a>
        </div>
      </div>

      <div class='admin-footer'><span>© 2026 PRIZE - Superfruits</span><span>Versión 1.0.0</span></div>
    </div>
    """
    return render_page(content, active='Admin')

@app.route('/admin/trabajadores', methods=['GET','POST'])
@admin_required
def admin_trabajadores():
    if request.method == 'POST':
        if 'excel' in request.files and request.files['excel'].filename:
            f = request.files['excel']; path = UPLOAD_DIR / f"base_{now_file()}_{secure_filename(f.filename)}"; f.save(path)
            wb = load_workbook(path, data_only=True); ws = wb.active
            headers = [clean(c.value).upper().replace('TRABAJADOR','NOMBRE').replace('FECHA NACIMIENTO','FECHA_NACIMIENTO').replace('FECHA INGRESO','FECHA_INGRESO').replace('JEFE INMEDIATO','JEFE_INMEDIATO').replace('JEFE DNI','JEFE_INMEDIATO') for c in ws[1]]
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
                    jefe_dni = normalizar_dni(row[idx('JEFE_INMEDIATO')] if idx('JEFE_INMEDIATO')>=0 else '')
                    jefe_nombre = ''
                    empresa = clean(row[idx('EMPRESA')] if idx('EMPRESA')>=0 else 'AQUANQA')
                    fecha_nac_raw = row[idx('FECHA_NACIMIENTO')] if idx('FECHA_NACIMIENTO')>=0 else ''
                    fecha_nac = excel_cell_fecha(fecha_nac_raw)
                    planilla = clean(row[idx('PLANILLA')] if idx('PLANILLA')>=0 else '')
                    fecha_ing = excel_cell_fecha(row[idx('FECHA_INGRESO')] if idx('FECHA_INGRESO')>=0 else '')
                    clave = generar_clave_trabajador(dni, fecha_nac)
                    con.execute("INSERT OR REPLACE INTO trabajadores(dni,nombre,correo,cargo,area,jefe_dni,jefe_nombre,empresa,planilla,fecha_nacimiento,fecha_ingreso,usuario_portal,clave_portal,activo,fecha_registro) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,1,?)", (dni,nombre,correo,cargo,area,jefe_dni,jefe_nombre,empresa,planilla,fecha_nac,fecha_ing,dni,clave,now_txt()))
                    n+=1
                con.commit()
            respaldar_exceles_locales()
            flash(f'Base cargada correctamente: {n} trabajadores. Respaldo actualizado en REGISTROS_EXCEL_LOCAL.', 'ok')
        else:
            dni=normalizar_dni(request.form.get('dni'))
            with db() as con:
                fecha_nac=clean(request.form.get('fecha_nacimiento')); clave=generar_clave_trabajador(dni, fecha_nac); con.execute("INSERT OR REPLACE INTO trabajadores(dni,nombre,correo,cargo,area,jefe_dni,jefe_nombre,empresa,planilla,fecha_nacimiento,fecha_ingreso,usuario_portal,clave_portal,activo,fecha_registro) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,1,?)", (dni,clean(request.form.get('nombre')),clean(request.form.get('correo')).lower(),clean(request.form.get('cargo')),clean(request.form.get('area')),normalizar_dni(request.form.get('jefe_dni')),clean(request.form.get('jefe_nombre')),clean(request.form.get('empresa')) or 'AQUANQA',clean(request.form.get('planilla')),fecha_nac,fecha_sin_hora(request.form.get('fecha_ingreso')),dni,clave,now_txt()))
                con.commit()
            respaldar_exceles_locales()
            flash('Trabajador guardado y respaldo Excel actualizado.', 'ok')
        return redirect(url_for('admin_trabajadores'))
    with db() as con:
        rows = con.execute("SELECT * FROM trabajadores ORDER BY nombre LIMIT 300").fetchall()
    table = ''.join([f"<tr><td>{r['dni']}</td><td>{r['nombre']}</td><td>{r['correo']}</td><td>{r['cargo'] or ''}</td><td>{r['empresa'] or ''}</td><td>{r['jefe_dni'] if 'jefe_dni' in r.keys() and r['jefe_dni'] else ''}</td><td>{r['planilla'] if 'planilla' in r.keys() and r['planilla'] else ''}</td></tr>" for r in rows])
    content = f"""
    <div class='topbar'><div><h1>Trabajadores</h1><div class='subtitle'>Carga manual o masiva por Excel.</div><div class='local-note'>Respaldo local automático: REGISTROS_EXCEL_LOCAL / 01_TRABAJADORES_LOCAL.xlsx</div></div></div><section class='grid'>
    <div class='card span-12'><h2>Nuevo trabajador</h2><form method='post' class='form-grid'><div class='field'><label>DNI</label><input name='dni' required></div><div class='field'><label>Trabajador</label><input name='nombre' required></div><div class='field'><label>Correo</label><input name='correo' type='email' required></div><div class='field'><label>Cargo</label><input name='cargo'></div><div class='field'><label>Área</label><input name='area'></div><div class='field'><label>Empresa</label><select name='empresa'><option>AQUANQA</option><option>AQUANCA II</option></select></div><div class='field'><label>Jefe inmediato DNI</label><input name='jefe_dni' placeholder='DNI del jefe'></div><div class='field'><label>Jefe nombre</label><input name='jefe_nombre' placeholder='Opcional'></div><div class='field'><label>Planilla</label><input name='planilla'></div><div class='field'><label>Fecha nacimiento</label><input name='fecha_nacimiento' placeholder='dd/mm/aaaa'></div><div class='field'><label>Fecha de ingreso</label><input name='fecha_ingreso' placeholder='dd/mm/aaaa'></div><button class='btn-green'>Guardar + crear usuario</button></form></div>
    <div class='card span-12'><h2>Carga Excel</h2><p class='muted'>Plantilla oficial: EMPRESA / DNI / TRABAJADOR / CARGO / AREA / JEFE INMEDIATO / PLANILLA / CORREO / FECHA NACIMIENTO. Crea usuario masivo con DNI y clave automática.</p><form method='post' enctype='multipart/form-data' class='form-grid'><div class='field'><label>Excel plantilla masiva</label><input type='file' name='excel' accept='.xlsx' required></div><button class='btn-blue'>Importar Excel</button><a class='btn-green' href='/admin/plantilla_trabajadores'>Descargar plantilla</a></form></div>
    <div class='card span-12'><h2>Listado</h2><div class='table-wrap'><table><tr><th>DNI</th><th>Nombre</th><th>Correo</th><th>Cargo</th><th>Empresa</th><th>Jefe DNI</th><th>Planilla</th></tr>{table}</table></div></div></section>"""
    return render_page(content, active='Trabajadores')

@app.route('/admin/plantilla_trabajadores')
@admin_required
def plantilla_trabajadores():
    path = PERSIST_DIR / 'PLANTILLA_CARGA_MASIVA_TRABAJADORES.xlsx'
    wb = Workbook(); ws = wb.active; ws.title = 'TRABAJADORES'
    headers = ['EMPRESA','DNI','TRABAJADOR','CARGO','AREA','JEFE INMEDIATO','PLANILLA','CORREO','FECHA NACIMIENTO','FECHA INGRESO']
    ws.append(headers)
    ws.append(['AQUANQA','74324033','APELLIDOS Y NOMBRES','Analista','RR.HH.','43043999','PLANILLA 01','correo@empresa.com','01/01/1990','01/05/2024'])
    for i, h in enumerate(headers, 1):
        font = copy(ws.cell(1, i).font); font.bold = True; ws.cell(1, i).font = font
        ws.column_dimensions[chr(64+i)].width = 24
    wb.save(path)
    return send_file(path, as_attachment=True, download_name='PLANTILLA_CARGA_MASIVA_TRABAJADORES.xlsx')

@app.route('/foto/<dni>')
def foto_trabajador(dni):
    t = get_trabajador(dni)
    if not t or not t['foto_ruta']: abort(404)
    path = Path(t['foto_ruta'])
    if not path.exists(): abort(404)
    return send_file(path, as_attachment=False)

@app.route('/mi_foto', methods=['POST'])
@worker_required
def mi_foto():
    f = request.files.get('foto')
    if not f or not f.filename:
        flash('Seleccione una foto.', 'error'); return redirect(url_for('panel'))
    ext = Path(secure_filename(f.filename)).suffix.lower()
    if ext not in ['.png','.jpg','.jpeg','.webp']:
        flash('Formato de foto no permitido.', 'error'); return redirect(url_for('panel'))
    folder = UPLOAD_DIR / 'fotos'; folder.mkdir(parents=True, exist_ok=True)
    path = folder / f"{session['dni']}_{now_file()}{ext}"; f.save(path)
    with db() as con:
        con.execute('UPDATE trabajadores SET foto_ruta=? WHERE dni=?', (str(path), session['dni'])); con.commit()
    flash('Foto actualizada correctamente.', 'ok')
    return redirect(url_for('panel'))


@app.route('/admin/usuarios')
@admin_required
def admin_usuarios():
    with db() as con:
        rows = con.execute("SELECT t.dni,t.nombre,t.correo,t.empresa,t.cargo,t.fecha_nacimiento,t.usuario_portal,t.clave_portal,t.activo,COALESCE(l.intentos,0) intentos,COALESCE(l.bloqueado,0) bloqueado FROM trabajadores t LEFT JOIN login_intentos l ON l.dni=t.dni ORDER BY t.nombre LIMIT 10000").fetchall()
    trs=[]
    for r in rows:
        dni = r['dni']
        clave = r['clave_portal'] or generar_clave_trabajador(r['dni'], r['fecha_nacimiento'])
        estado = '🔒 Bloqueado' if int(r['bloqueado'] or 0)==1 else '✅ Activo'
        desbloq = f" <a class='btn-green mini-btn' href='/admin/usuario/{dni}/desbloquear'>Desbloquear</a>" if int(r['bloqueado'] or 0)==1 or int(r['intentos'] or 0)>0 else ''
        trs.append(f"<tr><td>{dni}</td><td>{r['nombre']}</td><td>{r['usuario_portal'] or dni}</td><td><b>{clave}</b></td><td>{estado}<br><small>Intentos: {int(r['intentos'] or 0)}</small></td><td>{r['empresa'] or ''}</td><td><a class='btn-blue mini-btn' href='/admin/usuario/{dni}/reset'>Regenerar</a>{desbloq} <a class='btn-red mini-btn' onclick='return confirm(\"¿Eliminar trabajador/usuario?\")' href='/admin/usuario/{dni}/eliminar'>Eliminar</a></td></tr>")
    table=''.join(trs)
    content=f"""
    <div class='hero'><div class='topbar'><div><h1>Usuarios y contraseñas</h1><div class='subtitle'>Control para más de 10 mil trabajadores. Usuario = DNI; clave = fecha de nacimiento sin / (ddmmaaaa).</div></div><a class='btn-green' href='/admin/plantilla_trabajadores'>Plantilla masiva</a></div></div>
    <section class='grid'><div class='card span-12'><h2>Listado de accesos</h2><p class='muted'>El trabajador ingresa con usuario = DNI y clave = fecha nacimiento sin / (ddmmaaaa).</p><div class='table-wrap'><table><tr><th>DNI</th><th>Trabajador</th><th>Usuario</th><th>Clave</th><th>Estado login</th><th>Empresa</th><th>Opciones</th></tr>{table}</table></div></div></section>"""
    return render_page(content, active='Usuarios')

@app.route('/admin/usuario/<dni>/reset')
@admin_required
def admin_usuario_reset(dni):
    t=get_trabajador(dni)
    if not t: abort(404)
    clave=generar_clave_trabajador(dni, t['fecha_nacimiento'] if 'fecha_nacimiento' in t.keys() else '')
    with db() as con:
        con.execute("UPDATE trabajadores SET usuario_portal=?, clave_portal=? WHERE dni=?", (normalizar_dni(dni), clave, normalizar_dni(dni))); con.execute('DELETE FROM login_intentos WHERE dni=?',(normalizar_dni(dni),)); con.commit()
    flash('Usuario regenerado correctamente.', 'ok')
    return redirect(url_for('admin_usuarios'))


@app.route('/admin/usuario/<dni>/desbloquear')
@admin_required
def admin_usuario_desbloquear(dni):
    reset_intentos_login(dni)
    flash('Usuario desbloqueado. Ya puede ingresar con DNI y fecha de nacimiento sin /.', 'ok')
    return redirect(url_for('admin_usuarios'))

@app.route('/admin/usuario/<dni>/eliminar')
@admin_required
def admin_usuario_eliminar(dni):
    with db() as con:
        con.execute("DELETE FROM trabajadores WHERE dni=?", (normalizar_dni(dni),)); con.commit()
    flash('Trabajador/usuario eliminado.', 'ok')
    return redirect(url_for('admin_usuarios'))

@app.route('/admin/documentos', methods=['GET','POST'])
@admin_required
def admin_documentos():
    if request.method == 'POST':
        tipo = clean(request.form.get('tipo'))
        dni = normalizar_dni(request.form.get('dni'))
        periodo = request.form.get('periodo')
        detalle = request.form.get('detalle')
        obs = request.form.get('observacion')
        per_norm = clean(request.form.get('periodicidad_normal'))
        if tipo == 'Normal' and per_norm and per_norm.lower() not in clean(detalle).lower():
            detalle = (clean(detalle) + ' - ' + per_norm).strip(' -')
        files = request.files.getlist('archivos')
        ok=0
        try:
            for f in files:
                if f and f.filename:
                    guardar_documento(f, dni, tipo, periodo, detalle, obs, marca_carga(session.get('admin_user','admin'))); ok += 1
            flash(f'Carga completada: {ok} archivo(s).', 'ok')
        except Exception as e:
            flash(f'Error en carga: {e}', 'error')
        return redirect(url_for('admin_documentos', tipo=tipo))
    tipo = clean(request.args.get('tipo')) or 'Utilidad'
    buscar = clean(request.args.get('buscar'))
    periodo = clean(request.args.get('periodo'))
    sub = clean(request.args.get('sub'))
    tipo_options = ''.join([f"<option value='{k}' {'selected' if k==tipo else ''}>{l}</option>" for k,l,i in TIPOS_PAGO+TIPOS_EMPRESA+TIPOS_PERSONALES])
    pers = periodos_disponibles(tipo=tipo)
    periodo_options = "<option value=''>Todos</option>" + ''.join([f"<option {'selected' if p==periodo else ''}>{p}</option>" for p in pers])
    rows = listar_documentos(tipo=tipo if tipo else None, periodo=periodo or None, buscar=buscar, limit=500)
    content = f"""
    <div class='hero'><div class='topbar'><div><h1>Subir y gestionar documentos</h1><div class='subtitle'>Administrador: pago, empresa y documentos personales.</div></div><a class='btn-warn' href='/admin/sincronizar'>Actualizar / detectar PDFs</a><a class='btn-blue' href='/admin/crear_carpetas'>Crear carpetas + detectar</a></div></div><section class='grid'>
    <div class='card span-12'><h2>📁 Carpeta local automática</h2><p class='muted'>Ruta actual: <b>{DOCUMENTOS_BASE_DIR}</b><br>Coloca PDFs en DOCUMENTOS DE PAGO / BOLETAS NORMAL / SEMANAL o MENSUAL y presiona <b>Actualizar / detectar PDFs</b>. Solo se cargarán trabajadores activos. El DNI se detecta por nombre/ruta y también leyendo el contenido del PDF.</p><div class='actions'><a class='btn-warn' href='/admin/sincronizar'>Actualizar / detectar PDFs</a><a class='btn-blue' href='/admin/crear_carpetas'>Crear estructura</a></div></div><div class='card span-12'><h2>Carga de documentos</h2><form method='post' enctype='multipart/form-data' class='form-grid'><div class='field'><label>Tipo</label><select name='tipo'>{tipo_options}</select></div><div class='field'><label>DNI trabajador</label><input name='dni' placeholder='Vacío si es documento de empresa'></div><div class='field'><label>Periodo</label><input name='periodo' value='{datetime.now(APP_TZ).strftime('%Y-%m')}' list='periodos'></div><div class='field'><label>Detalle</label><input name='detalle' placeholder='Ej: Boleta semanal / Política actualizada'></div><div class='field'><label>Boleta Normal</label><select name='periodicidad_normal'><option value=''>No aplica</option><option>Mensual</option><option>Semanal</option></select></div><div class='field'><label>Archivos</label><input type='file' name='archivos' accept='.pdf,.png,.jpg,.jpeg,.webp,.doc,.docx,.xls,.xlsx' multiple required></div><div class='field'><label>Observación</label><textarea name='observacion' rows='2'></textarea></div><button class='btn-green'>Subir</button></form></div>
    <div class='card span-12'><h2>Filtros</h2><form method='get' class='form-grid'><div class='field'><label>Tipo</label><select name='tipo'>{tipo_options}</select></div><div class='field'><label>Periodo</label><select name='periodo'>{periodo_options}</select></div><div class='field'><label>Buscar</label><input name='buscar' value='{buscar}' placeholder='DNI, detalle, observación'></div><button class='btn-blue'>Filtrar</button><a class='btn' href='/admin/documentos'>Limpiar</a></form></div>
    <div class='card span-12'><h2>Listado</h2>{tabla_docs(rows)}</div></section>"""
    return render_page(content, active='Subir documentos')



def dias_entre_texto(fi, ff):
    a=parse_fecha_any(fi); b=parse_fecha_any(ff)
    if not a or not b: return 0
    return max((b-a).days+1, 0)


def _ids_periodos_texto(valor):
    return [int(x) for x in str(valor or '').replace(',', '|').split('|') if str(x).strip().isdigit()]


def dias_reservados_periodos(con, dni, excluir_id=None):
    """Devuelve días ya comprometidos por periodo para no usar el mismo saldo 2 veces.
    Considera solicitudes pendientes y aprobadas; ignora rechazadas/anuladas.
    """
    params = [dni]
    extra = ''
    if excluir_id:
        extra = ' AND id<>?'
        params.append(excluir_id)
    filas = con.execute(f"""
        SELECT id, dias, estado, periodo_ids
        FROM vacaciones_solicitudes
        WHERE dni=?
          AND COALESCE(estado,'') NOT LIKE 'Rechazado%'
          AND COALESCE(estado,'') NOT LIKE 'Anulado%'
          {extra}
    """, params).fetchall()
    usados = {}
    for row in filas:
        ids = _ids_periodos_texto(row['periodo_ids'] if 'periodo_ids' in row.keys() else '')
        if not ids:
            continue
        dias = float(row['dias'] or 0)
        por_periodo = dias / max(len(ids), 1)
        for pid in ids:
            usados[pid] = usados.get(pid, 0) + por_periodo
    return usados


def saldo_disponible_real(row, usados_por_periodo):
    return max(float(row['saldo'] or 0) - float(usados_por_periodo.get(int(row['id']), 0)), 0)

def obtener_jefe_dni_trabajador(con, trabajador_dni, periodo_ids=None):
    """Obtiene el DNI del jefe inmediato de forma robusta.
    Prioridad: periodo seleccionado en saldos -> cualquier saldo del trabajador -> ficha trabajadores.
    Esto evita que una solicitud quede sin aprobador cuando el jefe viene de la plantilla de saldos.
    """
    trabajador_dni = normalizar_dni(trabajador_dni)
    periodo_ids = [int(x) for x in (periodo_ids or []) if str(x).isdigit()]
    if periodo_ids:
        marks = ','.join(['?'] * len(periodo_ids))
        row = con.execute(f"""
            SELECT jefe_dni FROM vacaciones_saldos
            WHERE dni=? AND id IN ({marks}) AND COALESCE(jefe_dni,'')<>''
            ORDER BY periodo_inicio, periodo_fin LIMIT 1
        """, [trabajador_dni] + periodo_ids).fetchone()
        if row:
            jd = normalizar_dni(row['jefe_dni'])
            if jd: return jd
    row = con.execute("""
        SELECT jefe_dni FROM vacaciones_saldos
        WHERE dni=? AND COALESCE(jefe_dni,'')<>''
        ORDER BY periodo_inicio, periodo_fin LIMIT 1
    """, (trabajador_dni,)).fetchone()
    if row:
        jd = normalizar_dni(row['jefe_dni'])
        if jd: return jd
    row = con.execute("SELECT jefe_dni FROM trabajadores WHERE dni=?", (trabajador_dni,)).fetchone()
    if row and 'jefe_dni' in row.keys():
        jd = normalizar_dni(row['jefe_dni'])
        if jd: return jd
    return ''

def sql_solicitudes_jefe(extra_where=''):
    return f"""
        SELECT vs.* FROM vacaciones_solicitudes vs
        LEFT JOIN trabajadores tr ON tr.dni = vs.dni
        WHERE (
            normalizar_dni_sql(COALESCE(vs.jefe_dni,'')) = ?
            OR normalizar_dni_sql(COALESCE(tr.jefe_dni,'')) = ?
            OR EXISTS (
                SELECT 1 FROM vacaciones_saldos s
                WHERE s.dni = vs.dni
                  AND normalizar_dni_sql(COALESCE(s.jefe_dni,'')) = ?
            )
        )
        {extra_where}
    """

@app.route('/admin/modulo/documentos')
@admin_required
def admin_modulo_documentos():
    with db() as con:
        total=con.execute('SELECT COUNT(*) c FROM documentos').fetchone()['c']
        pendientes=con.execute("SELECT COUNT(*) c FROM documentos WHERE COALESCE(estado,'Pendiente') IN ('Pendiente','Aceptado','Firmado')").fetchone()['c']
        aprob=con.execute("SELECT COUNT(*) c FROM documentos WHERE estado='Aprobado'").fetchone()['c']
    content=f"""
    <div class='hero'><div class='topbar'><div><h1>Gestión <span class='accent'>Documental</span></h1><div class='subtitle'>Concentra todo lo ya implementado: cargas, PDFs, carpetas locales, aceptación/firma/aprobación y trazabilidad.</div></div><a class='btn-green' href='/admin/documentos'>Entrar a documentos</a></div></div>
    <section class='grid'><div class='card mini'><div><h3>Total documentos</h3><b>{total}</b></div><div class='ico'>🗃️</div></div><div class='card mini'><div><h3>Pendientes</h3><b>{pendientes}</b></div><div class='ico'>⏳</div></div><div class='card mini'><div><h3>Aprobados</h3><b>{aprob}</b></div><div class='ico'>✅</div></div>
    <div class='card span-12'><div class='module-tabs'><a class='module-tile' href='/admin/documentos'><h2>📤 Subir documentos</h2><p class='muted'>Pago, empresa y personales.</p></a><a class='module-tile' href='/admin/sincronizar'><h2>🔎 Detectar PDFs</h2><p class='muted'>Lee carpetas locales y detecta DNI.</p></a><a class='module-tile' href='/admin/crear_carpetas'><h2>📁 Crear carpetas</h2><p class='muted'>Estructura automática DOCUMENTOS_PRIZE_AUTO.</p></a></div></div></section>"""
    return render_page(content, active='Modulo documentos')

@app.route('/admin/vacaciones', methods=['GET','POST'])
@admin_required
def admin_vacaciones():
    if request.method=='POST':
        f=request.files.get('excel')
        ok=0
        if f and f.filename:
            path=UPLOAD_DIR/'vacaciones'; path.mkdir(parents=True, exist_ok=True)
            x=path/(now_file()+'_'+secure_filename(f.filename)); f.save(x)
            wb=load_workbook(x, data_only=True); ws=wb.active
            headers=[str(c.value or '').strip().upper() for c in ws[1]]
            def val(row, names):
                for n in names:
                    if n in headers: return row[headers.index(n)].value
                return ''
            with db() as con:
                for row in ws.iter_rows(min_row=2):
                    dni=normalizar_dni(val(row,['DNI','DOCUMENTO','CODIGO','CÓDIGO']))
                    if not dni: continue
                    trabajador=clean(val(row,['TRABAJADOR','NOMBRE','APELLIDOS Y NOMBRES']))
                    gan=float(val(row,['DIAS GANADOS','DÍAS GANADOS','GANADOS']) or 0)
                    saldo=float(val(row,['SALDO','SALDO VACACIONAL']) or gan)
                    trabajador_db = con.execute('SELECT * FROM trabajadores WHERE dni=?', (dni,)).fetchone()
                    if trabajador_db:
                        trabajador = trabajador or trabajador_db['nombre']
                        fecha_ing = trabajador_db['fecha_ingreso'] if 'fecha_ingreso' in trabajador_db.keys() else ''
                    else:
                        fecha_ing = ''
                    p_ini=periodo_year_value(val(row,['I_PERIODO','PERIODO INICIO','INICIO PERIODO','FECHA INICIO PERIODO']))
                    p_fin=periodo_year_value(val(row,['F_PERIODO','PERIODO FIN','FIN PERIODO','FECHA FIN PERIODO']))
                    periodo=clean(val(row,['PERIODO','PERÍODO'])) or periodo_anual_texto(p_ini, p_fin)
                    jefe_raw=clean(val(row,['JEFE DNI','DNI JEFE','JEFE INMEDIATO','JEFE']))
                    jefe_dni=normalizar_dni(jefe_raw)
                    jefe_nombre=''
                    if jefe_dni:
                        jr=con.execute('SELECT nombre FROM trabajadores WHERE dni=?', (jefe_dni,)).fetchone()
                        jefe_nombre = jr['nombre'] if jr else jefe_raw
                    else:
                        jr=con.execute('SELECT dni,nombre FROM trabajadores WHERE UPPER(nombre)=UPPER(?)', (jefe_raw,)).fetchone()
                        jefe_dni = jr['dni'] if jr else ''
                        jefe_nombre = jr['nombre'] if jr else jefe_raw
                    con.execute('INSERT INTO vacaciones_saldos(dni,trabajador,empresa,area,jefe,jefe_dni,fecha_ingreso,periodo_inicio,periodo_fin,dias_ganados,dias_gozados,saldo,periodo,fecha_carga,uploaded_by) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(dni,periodo_inicio,periodo_fin) DO UPDATE SET trabajador=excluded.trabajador,empresa=excluded.empresa,area=excluded.area,jefe=excluded.jefe,jefe_dni=excluded.jefe_dni,fecha_ingreso=excluded.fecha_ingreso,dias_ganados=excluded.dias_ganados,dias_gozados=0,saldo=excluded.saldo,periodo=excluded.periodo,fecha_carga=excluded.fecha_carga,uploaded_by=excluded.uploaded_by', (dni,trabajador,clean(val(row,['EMPRESA'])),clean(val(row,['AREA','ÁREA'])),jefe_nombre,jefe_dni,fecha_ing,p_ini,p_fin,gan,0,saldo,periodo,now_txt(),marca_carga(session.get('admin_user','admin'))))
                    ok+=1
                con.commit()
                respaldar_exceles_locales()
        flash(f'Saldos vacacionales cargados/actualizados: {ok}. Respaldo Excel local actualizado.','ok')
        return redirect(url_for('admin_vacaciones'))
    q_sol=clean(request.args.get('q_sol'))
    q_sal=clean(request.args.get('q_sal'))
    with db() as con:
        params=[]; where=''
        if q_sal:
            where='WHERE dni LIKE ? OR UPPER(trabajador) LIKE UPPER(?)'
            params=[q_sal+'%', '%'+q_sal+'%']
        saldos=con.execute(f'SELECT * FROM vacaciones_saldos {where} ORDER BY trabajador, periodo_inicio LIMIT 500', params).fetchall()
        params=[]; where=''
        if q_sol:
            where='WHERE dni LIKE ? OR UPPER(trabajador) LIKE UPPER(?)'
            params=[q_sol+'%', '%'+q_sol+'%']
        solicitudes=con.execute(f'SELECT * FROM vacaciones_solicitudes {where} ORDER BY id DESC LIMIT 500', params).fetchall()
    sal=''.join([f"<tr><td>{r['dni']}</td><td>{r['trabajador']}</td><td>{r['empresa'] or ''}</td><td>{r['area'] or ''}</td><td>{r['jefe_dni'] if 'jefe_dni' in r.keys() else ''}</td><td>{r['jefe'] or ''}</td><td>{r['periodo_inicio'] or ''}</td><td>{r['periodo_fin'] or ''}</td><td>{r['dias_ganados']}</td><td><b>{r['saldo']}</b></td></tr>" for r in saldos])
    sol=''.join([f"<tr><td>{r['id']}</td><td>{r['dni']}</td><td>{r['trabajador']}</td><td>{r['jefe_dni'] if 'jefe_dni' in r.keys() else ''}</td><td>{r['fecha_inicio']} al {r['fecha_fin']}</td><td>{r['dias']}</td><td>{r['periodo_detalle'] if 'periodo_detalle' in r.keys() and r['periodo_detalle'] else ''}</td><td><span class='status-pill'>{r['estado']}</span></td><td class='actions'><a class='btn-green mini-btn' href='/admin/vacaciones/{r['id']}/jefe/aprobar'>Apr. jefe</a><a class='btn-green mini-btn' href='/admin/vacaciones/{r['id']}/gh/aprobar'>Apr. GTH</a><a class='btn-red mini-btn' href='/admin/vacaciones/{r['id']}/rechazar'>Rechazar</a></td></tr>" for r in solicitudes])
    content=f"""
    <div class='hero'><div class='topbar'><div><h1>Gestión <span class='accent'>Vacacional</span></h1><div class='subtitle'>Administrador carga saldos; usuario solicita goce; flujo: jefe inmediato → Gestión del Talento Humano.</div></div><a class='btn-green' href='/admin/vacaciones/plantilla'>Descargar plantilla</a></div></div>
    <section class='grid'><div id='aprobaciones' class='card mini'><div><h3>Pendientes de aprobación</h3><b>{len([r for r in solicitudes if 'Pendiente' in (r['estado'] or '')])}</b></div><div class='ico'>✅</div></div><div class='card mini'><div><h3>Saldos registrados</h3><b>{len(saldos)}</b></div><div class='ico'>🗓️</div></div><div class='card mini'><div><h3>Solicitudes totales</h3><b>{len(solicitudes)}</b></div><div class='ico'>📄</div></div><div id='cargar-saldos' class='card span-12'><h2>🏖️ Saldos Vacacionales</h2><form method='post' enctype='multipart/form-data' class='form-grid'><div class='field'><label>Excel saldos</label><input type='file' name='excel' accept='.xlsx' required></div><button class='btn-green'>Subir saldos</button></form><p class='muted'>Columnas: EMPRESA, DNI, TRABAJADOR, AREA, JEFE INMEDIATO (DNI), I_PERIODO, F_PERIODO, DIAS GANADOS, SALDO. No usar FECHA INGRESO ni PERIODO ni DÍAS GOZADOS.</p></div>
    <div id='solicitudes' class='card span-12'><h2>📄 Solicitudes de vacaciones</h2><form method='get' class='form-grid'><div class='field'><label>Buscar por DNI o apellidos</label><input name='q_sol' value='{q_sol}' placeholder='Ej: 473 o QUINTANA'></div><button class='btn-green'>Filtrar solicitudes</button><a class='btn' href='/admin/vacaciones#solicitudes'>Limpiar</a></form><div class='table-wrap'><table><tr><th>ID</th><th>DNI</th><th>Trabajador</th><th>DNI jefe</th><th>Rango</th><th>Días</th><th>Periodo usado</th><th>Estado</th><th>Acciones</th></tr>{sol or '<tr><td colspan=9>No hay solicitudes.</td></tr>'}</table></div></div>
    <div id='reportes' class='card span-12'><h2>📑 Reporte de saldos cargados</h2><form method='get' class='form-grid'><div class='field'><label>Buscar por DNI o apellidos</label><input name='q_sal' value='{q_sal}' placeholder='Ej: 473 o QUINTANA'></div><button class='btn-green'>Filtrar saldos</button><a class='btn' href='/admin/vacaciones#reportes'>Limpiar</a></form><div class='table-wrap'><table><tr><th>DNI</th><th>Trabajador</th><th>Empresa</th><th>Área</th><th>DNI jefe</th><th>Jefe</th><th>I_Periodo</th><th>F_Periodo</th><th>Ganados</th><th>Saldo</th></tr>{sal or '<tr><td colspan=10>No hay saldos cargados.</td></tr>'}</table></div></div></section>"""
    return render_page(content, active='Gestion Vacacional')

@app.route('/admin/vacaciones/plantilla')
@admin_required
def admin_vacaciones_plantilla():
    path=PERSIST_DIR/'PLANTILLA_SALDOS_VACACIONALES.xlsx'
    wb=Workbook(); ws=wb.active; ws.title='SALDOS'
    headers=['EMPRESA','DNI','TRABAJADOR','AREA','JEFE INMEDIATO','I_PERIODO','F_PERIODO','DIAS GANADOS','SALDO']
    ws.append(headers); ws.append(['AQUANQA','74324033','APELLIDOS Y NOMBRES','RRHH','43043999','2025','2026',30,30])
    # Formato simple: I_PERIODO y F_PERIODO son SOLO AÑOS, no fechas completas.
    for i,h in enumerate(headers,1):
        ws.column_dimensions[chr(64+i)].width=24
    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='166534')
        cell.alignment = Alignment(horizontal='center')
    for col in ('F','G'):
        for row in range(2, 502):
            ws[f'{col}{row}'].number_format = '@'
    ws.freeze_panes='A2'
    wb.save(path); return send_file(path, as_attachment=True, download_name='PLANTILLA_SALDOS_VACACIONALES.xlsx')

@app.route('/admin/vacaciones/<int:sid>/<rol>/<accion>')
@admin_required
def admin_vacaciones_accion(sid, rol, accion):
    if accion=='aprobar' and rol=='jefe': estado='Pendiente GTH'; col='fecha_jefe'
    elif accion=='aprobar' and rol=='gh': estado='Aprobado GTH'; col='fecha_gh'
    else: estado='Rechazado'; col='fecha_gh'
    with db() as con:
        con.execute(f'UPDATE vacaciones_solicitudes SET estado=?, {col}=? WHERE id=?', (estado, now_txt(), sid)); con.commit(); respaldar_exceles_locales()
    flash('Solicitud actualizada.', 'ok'); return redirect(url_for('admin_vacaciones'))

@app.route('/vacaciones/mi_solicitud', methods=['GET','POST'])
@worker_required
def trabajador_vacaciones():
    dni=session['dni']; t=get_trabajador(dni)
    with db() as con:
        saldos_usuario=con.execute('SELECT * FROM vacaciones_saldos WHERE dni=? ORDER BY periodo_inicio, periodo_fin',(dni,)).fetchall()
        saldo=saldos_usuario[0] if saldos_usuario else None
    if request.method=='POST':
        fi_raw=clean(request.form.get('fecha_inicio')); ff_raw=clean(request.form.get('fecha_fin'))
        fi_date=parse_fecha_any(fi_raw); ff_date=parse_fecha_any(ff_raw)
        hoy=hoy_lima()
        hoy_iso=hoy.isoformat()
        if not fi_raw or not ff_raw or not fi_date or not ff_date:
            flash('No se registró la solicitud: debe seleccionar fechas válidas.', 'err')
            return redirect(url_for('trabajador_vacaciones'))
        # BLOQUEO DEFINITIVO EN SERVIDOR: aunque manipulen el HTML o escriban la fecha manualmente.
        if fi_date < hoy:
            flash(f'No se registró la solicitud: la fecha de inicio {fi_date.strftime("%d/%m/%Y")} es anterior a hoy {hoy.strftime("%d/%m/%Y")}.', 'err')
            return redirect(url_for('trabajador_vacaciones'))
        if ff_date < hoy:
            flash(f'No se registró la solicitud: la fecha fin {ff_date.strftime("%d/%m/%Y")} es anterior a hoy {hoy.strftime("%d/%m/%Y")}.', 'err')
            return redirect(url_for('trabajador_vacaciones'))
        if ff_date < fi_date:
            flash('No se registró la solicitud: la fecha fin no puede ser menor que la fecha inicio.', 'err')
            return redirect(url_for('trabajador_vacaciones'))
        dias_anticipacion = (fi_date - hoy).days
        if dias_anticipacion > 90:
            flash(f'No se registró la solicitud: solo se permiten solicitudes con máximo 90 días de anticipación. Inicio seleccionado: {fi_date.strftime("%d/%m/%Y")}.', 'err')
            return redirect(url_for('trabajador_vacaciones'))
        fi=fi_date.isoformat(); ff=ff_date.isoformat(); dias=dias_entre_texto(fi,ff)
        adelanto = '1' if request.form.get('adelanto') else ''
        periodo_ids = [int(x) for x in request.form.getlist('periodos') if str(x).isdigit()]
        if not periodo_ids:
            flash('Seleccione con check el periodo que usará para gozar vacaciones.', 'err')
            return redirect(url_for('trabajador_vacaciones'))
        with db() as con:
            marks=','.join(['?']*len(periodo_ids))
            saldos_sel=con.execute(f'SELECT * FROM vacaciones_saldos WHERE dni=? AND id IN ({marks}) ORDER BY periodo_inicio, periodo_fin', [dni]+periodo_ids).fetchall()
            usados = dias_reservados_periodos(con, dni)
        saldo_disponible = sum(saldo_disponible_real(r, usados) for r in saldos_sel)
        if dias <= 0:
            flash('Rango de fechas inválido. Revisa inicio y fin.', 'err')
            return redirect(url_for('trabajador_vacaciones'))
        if not saldos_sel or saldo_disponible <= 0:
            flash('No se registró la solicitud: el/los periodo(s) seleccionado(s) no tienen saldo disponible o ya fueron usados en otra solicitud.', 'err')
            return redirect(url_for('trabajador_vacaciones'))
        if dias > saldo_disponible:
            flash(f'No se registró la solicitud: solicita {dias} día(s), pero el saldo real disponible es {saldo_disponible}.', 'err')
            return redirect(url_for('trabajador_vacaciones'))
        estado = 'Pendiente jefe'
        motivo_base = clean(request.form.get('motivo'))
        periodo_detalle = ' | '.join([f"{r['periodo_inicio']}-{r['periodo_fin']} (saldo disponible {saldo_disponible_real(r, usados)})" for r in saldos_sel])
        periodo_ids_txt = '|'.join(str(x) for x in periodo_ids)
        if adelanto:
            motivo_base = (motivo_base + ' | ' if motivo_base else '') + 'Solicitud marcada como comentario especial; validada dentro del saldo disponible.'
        with db() as con:
            jefe_dni = obtener_jefe_dni_trabajador(con, dni, periodo_ids)
            if not jefe_dni:
                flash('No se registró la solicitud: este trabajador no tiene DNI de jefe inmediato. Cargue la plantilla de trabajadores o saldos con JEFE INMEDIATO = DNI del jefe.', 'err')
                return redirect(url_for('trabajador_vacaciones'))
            # Segundo candado justo antes de grabar: evita registrar fechas pasadas aunque el formulario haya sido alterado.
            if parse_fecha_any(fi) < hoy_lima() or parse_fecha_any(ff) < hoy_lima():
                flash('No se registró la solicitud: las fechas no pueden ser anteriores a hoy.', 'err')
                return redirect(url_for('trabajador_vacaciones'))
            con.execute('INSERT INTO vacaciones_solicitudes(dni,trabajador,jefe_dni,fecha_inicio,fecha_fin,dias,motivo,estado,fecha_solicitud,periodo_detalle,periodo_ids) VALUES(?,?,?,?,?,?,?,?,?,?,?)',(dni,t['nombre'] if t else '',jefe_dni,fi,ff,dias,motivo_base,estado,now_txt(),periodo_detalle,periodo_ids_txt)); con.commit(); respaldar_exceles_locales()
        flash('Solicitud registrada. Pasará por jefe inmediato y Gestión del Talento Humano.','ok')
        return redirect(url_for('trabajador_vacaciones'))
    with db() as con:
        saldos_usuario=con.execute('SELECT * FROM vacaciones_saldos WHERE dni=? ORDER BY periodo_inicio, periodo_fin',(dni,)).fetchall()
        saldo=saldos_usuario[0] if saldos_usuario else None
        solicitudes=con.execute('SELECT * FROM vacaciones_solicitudes WHERE dni=? ORDER BY id DESC',(dni,)).fetchall()
        por_aprobar=con.execute(sql_solicitudes_jefe("AND vs.estado='Pendiente jefe' ORDER BY vs.id DESC"),(dni,dni,dni,dni)).fetchall()
    sol=''.join([f"<div class='sol-card'><div data-label='Fecha'><b>{r['fecha_solicitud']}</b></div><div data-label='Rango'><b>{r['fecha_inicio']} al {r['fecha_fin']}</b></div><div data-label='Días' class='dias'><b>{r['dias']}</b></div><div data-label='Periodo usado'><b>{r['periodo_detalle'] if 'periodo_detalle' in r.keys() and r['periodo_detalle'] else '-'}</b></div><div data-label='Estado'><span class='status-pill'>{r['estado']}</span></div><div data-label='Comentario' class='coment'>{r['motivo'] or '-'}</div></div>" for r in solicitudes])
    sol_aprobar=''.join([f"<tr><td>{r['fecha_solicitud']}</td><td>{r['dni']}</td><td>{r['trabajador']}</td><td>{r['fecha_inicio']} al {r['fecha_fin']}</td><td>{r['dias']}</td><td><span class='status-pill'>{r['estado']}</span></td><td class='actions'><a class='btn-green mini-btn' href='/vacaciones/aprobar_jefe/{r['id']}'>Aprobar</a><a class='btn-red mini-btn' href='/vacaciones/rechazar_jefe/{r['id']}'>Rechazar</a></td></tr>" for r in por_aprobar])
    with db() as con:
        usados_periodos = dias_reservados_periodos(con, dni)
    saldo_val = sum(saldo_disponible_real(r, usados_periodos) for r in saldos_usuario)
    periodos_html = ''.join([f"<label class='period-card {'disabled' if saldo_disponible_real(r, usados_periodos) <= 0 else ''}'><input type='checkbox' name='periodos' value='{r['id']}' {'disabled' if saldo_disponible_real(r, usados_periodos) <= 0 else ''}><span class='period-main'><span class='period-years'>{r['periodo_inicio'] or ''} - {r['periodo_fin'] or ''}</span><span class='period-meta'><span class='period-badge'>Ganados: <strong>{r['dias_ganados']}</strong></span><span class='period-badge'>Saldo real: <strong>{saldo_disponible_real(r, usados_periodos)}</strong></span><span class='period-badge'>Usado/Pendiente: <strong>{usados_periodos.get(int(r['id']), 0)}</strong></span></span></span></label>" for r in saldos_usuario]) or '<p class=\'muted\'>No tiene periodos vacacionales cargados.</p>'
    content=f"""
    <div class='hero'><div class='topbar'><div><h1>Gestión <span class='accent'>Vacacional</span></h1><div class='subtitle'>Consulta tu saldo, valida días disponibles y registra solicitudes.</div></div></div></div>
    <section class='grid'><div class='card mini'><div><h3>Saldo disponible</h3><b>{saldo_val}</b></div><div class='ico'>🏖️</div></div><div class='card mini'><div><h3>Días ganados</h3><b>{sum(float(r['dias_ganados'] or 0) for r in saldos_usuario)}</b></div><div class='ico'>📈</div></div><div class='card mini'><div><h3>Periodos</h3><b>{len(saldos_usuario)}</b></div><div class='ico'>📅</div></div><div class='card mini'><div><h3>Fecha ingreso</h3><b>{fecha_sin_hora(t['fecha_ingreso'] if t and 'fecha_ingreso' in t.keys() else '') or '-'}</b></div><div class='ico'>🗓️</div></div>
    <div class='card span-12' style='{"display:block" if sol_aprobar else "display:none"}'><h2>✅ Solicitudes pendientes por aprobar como jefe inmediato</h2><p class='muted'>Te aparecen aquí solo los trabajadores que tienen tu DNI como jefe inmediato en la plantilla de saldos.</p><div class='table-wrap'><table><tr><th>Fecha</th><th>DNI</th><th>Trabajador</th><th>Rango</th><th>Días</th><th>Estado</th><th>Acciones</th></tr>{sol_aprobar or '<tr><td colspan=7>No tienes solicitudes pendientes por aprobar.</td></tr>'}</table></div></div>
    <div id='solicitar' class='card span-12 vac-request-card'><div class='vac-head'><div><h2>🗓️ Nueva solicitud</h2><p class='vac-help'>Marca el periodo que vas a utilizar. Puedes seleccionar más de uno cuando el descanso consuma saldos acumulados.</p></div></div><form method='post' id='formSolicitudVacaciones'><div class='field'><label>Periodos disponibles</label><div class='period-list'>{periodos_html}</div></div><div class='vac-form-row'><div class='field'><label>Inicio</label><input type='date' id='fecha_inicio_vac' name='fecha_inicio' min='{hoy_lima().isoformat()}' required></div><div class='field'><label>Fin</label><input type='date' id='fecha_fin_vac' name='fecha_fin' min='{hoy_lima().isoformat()}' required></div><div class='field'><label>Motivo / comentario</label><input name='motivo' placeholder='Goce vacacional'></div></div><div class='vac-submit-row'><label class='check-card'><input type='checkbox' name='adelanto' value='1'> Requiere revisión especial</label><button class='btn-green'>Registrar solicitud</button></div></form><script>
(function(){{
  const hoyISO = '{hoy_lima().isoformat()}';
  const f = document.getElementById('formSolicitudVacaciones');
  const ini = document.getElementById('fecha_inicio_vac');
  const fin = document.getElementById('fecha_fin_vac');

  function valorISO(campo){{
    if(!campo || !campo.value) return '';
    const v = campo.value.trim();
    // Navegadores modernos devuelven YYYY-MM-DD.
    if(/^\d{{4}}-\d{{2}}-\d{{2}}$/.test(v)) return v;
    // Soporte adicional si el navegador deja escribir DD/MM/AAAA.
    const m = v.match(/^(\d{{1,2}})[\/\-](\d{{1,2}})[\/\-](\d{{4}})$/);
    if(m){{
      const d = m[1].padStart(2,'0');
      const mo = m[2].padStart(2,'0');
      return `${{m[3]}}-${{mo}}-${{d}}`;
    }}
    return '';
  }}

  [ini, fin].forEach(x => {{
    if(x){{
      x.setAttribute('min', hoyISO);
      x.setAttribute('autocomplete', 'off');
      // No mostrar alertas mientras la fecha está incompleta; solo validar al registrar.
    }}
  }});

  if(f){{ f.addEventListener('submit', function(e){{
    const iniISO = valorISO(ini);
    const finISO = valorISO(fin);
    if(!iniISO || !finISO){{
      e.preventDefault(); alert('Seleccione fecha de inicio y fin válidas.'); return false;
    }}
    if(iniISO < hoyISO || finISO < hoyISO){{
      e.preventDefault(); alert('No se registró: la fecha de inicio y fin no pueden ser anteriores a hoy (' + hoyISO + ').'); return false;
    }}
    if(finISO < iniISO){{
      e.preventDefault(); alert('La fecha fin no puede ser menor que inicio.'); return false;
    }}
  }});}}
}})();
</script></div>
    <div class='card span-12'><h2>Mis solicitudes</h2><div class='sol-cards'><div class='sol-card head'><div>Fecha</div><div>Rango</div><div>Días</div><div>Periodo usado</div><div>Estado</div><div>Comentario</div></div>{sol or "<div class='sol-empty'>No hay solicitudes registradas.</div>"}</div></div></section>"""
    return render_page(content, active='Gestion Vacacional')


@app.route('/vacaciones/aprobar_jefe/<int:sid>')
@worker_required
def vacaciones_aprobar_jefe(sid):
    dni=session['dni']
    with db() as con:
        r=con.execute("SELECT * FROM vacaciones_solicitudes WHERE id=?", (sid,)).fetchone()
        jefe_solicitud = normalizar_dni(r['jefe_dni'] if r and 'jefe_dni' in r.keys() else '') if r else ''
        jefe_trabajador = obtener_jefe_dni_trabajador(con, r['dni']) if r else ''
        if not r or (jefe_solicitud != dni and jefe_trabajador != dni):
            flash('No autorizado: esta solicitud no corresponde a tu aprobación como jefe inmediato.', 'err')
            return redirect(url_for('trabajador_vacaciones'))
        con.execute("UPDATE vacaciones_solicitudes SET estado='Pendiente GTH', fecha_jefe=?, comentario_jefe=? WHERE id=?", (now_txt(), 'Aprobado por jefe inmediato', sid))
        con.commit()
    flash('Solicitud aprobada por jefe inmediato. Ahora queda pendiente de GTH.', 'ok')
    return redirect(url_for('trabajador_vacaciones'))

@app.route('/vacaciones/rechazar_jefe/<int:sid>')
@worker_required
def vacaciones_rechazar_jefe(sid):
    dni=session['dni']
    with db() as con:
        r=con.execute("SELECT * FROM vacaciones_solicitudes WHERE id=?", (sid,)).fetchone()
        jefe_solicitud = normalizar_dni(r['jefe_dni'] if r and 'jefe_dni' in r.keys() else '') if r else ''
        jefe_trabajador = obtener_jefe_dni_trabajador(con, r['dni']) if r else ''
        if not r or (jefe_solicitud != dni and jefe_trabajador != dni):
            flash('No autorizado: esta solicitud no corresponde a tu aprobación como jefe inmediato.', 'err')
            return redirect(url_for('trabajador_vacaciones'))
        con.execute("UPDATE vacaciones_solicitudes SET estado='Rechazado por jefe', fecha_jefe=?, comentario_jefe=? WHERE id=?", (now_txt(), 'Rechazado por jefe inmediato', sid))
        con.commit()
    flash('Solicitud rechazada por jefe inmediato.', 'ok')
    return redirect(url_for('trabajador_vacaciones'))


@app.route('/vacaciones/aprobaciones_jefe')
@worker_required
def vacaciones_aprobaciones_jefe():
    dni=session['dni']; t=get_trabajador(dni)
    with db() as con:
        rows=con.execute(sql_solicitudes_jefe("ORDER BY CASE WHEN vs.estado='Pendiente jefe' THEN 0 ELSE 1 END, vs.id DESC"), (dni,dni,dni,dni)).fetchall()
    pendientes=sum(1 for r in rows if (r['estado'] or '') == 'Pendiente jefe')
    aprobadas=sum(1 for r in rows if 'GTH' in (r['estado'] or '') or 'Aprobado' in (r['estado'] or ''))
    rechazadas=sum(1 for r in rows if 'Rechazado' in (r['estado'] or ''))
    cards=[]
    for r in rows:
        acciones = ""
        if (r['estado'] or '') == 'Pendiente jefe':
            acciones = f"<div class='actions'><a class='btn-green mini-btn' href='/vacaciones/aprobar_jefe/{r['id']}'>Aprobar</a><a class='btn-red mini-btn' href='/vacaciones/rechazar_jefe/{r['id']}'>Rechazar</a></div>"
        cards.append(f"""
        <div class='sol-card'>
          <div data-label='Fecha'><b>{r['fecha_solicitud']}</b></div>
          <div data-label='Trabajador'><b>{r['dni']} - {r['trabajador']}</b></div>
          <div data-label='Rango'><b>{r['fecha_inicio']} al {r['fecha_fin']}</b></div>
          <div data-label='Días' class='dias'><b>{r['dias']}</b></div>
          <div data-label='Periodo usado'><b>{r['periodo_detalle'] if 'periodo_detalle' in r.keys() and r['periodo_detalle'] else '-'}</b></div>
          <div data-label='Estado'><span class='status-pill'>{r['estado']}</span></div>
          <div data-label='Acción'>{acciones or '<span class="muted">Sin acción pendiente</span>'}</div>
        </div>""")
    content=f"""
    <div class='hero'><div class='topbar'><div><h1>Aprobaciones <span class='accent'>Jefe inmediato</span></h1><div class='subtitle'>Este panel aparece para cualquier trabajador que figure como JEFE INMEDIATO por DNI en la plantilla de saldos.</div></div></div></div>
    <section class='grid'>
      <div class='card mini'><div><h3>Jefe</h3><b>{t['nombre'] if t else dni}</b></div><div class='ico'>👤</div></div>
      <div class='card mini'><div><h3>Pendientes</h3><b>{pendientes}</b></div><div class='ico'>⏳</div></div>
      <div class='card mini'><div><h3>Aprobadas / GTH</h3><b>{aprobadas}</b></div><div class='ico'>✅</div></div>
      <div class='card mini'><div><h3>Rechazadas</h3><b>{rechazadas}</b></div><div class='ico'>🚫</div></div>
      <div class='card span-12'><h2>Solicitudes asignadas a tu DNI</h2><p class='muted'>Si no aparece una solicitud, revise que en la plantilla de saldos el campo JEFE INMEDIATO tenga exactamente tu DNI.</p><div class='sol-cards'><div class='sol-card head'><div>Fecha</div><div>Trabajador</div><div>Rango</div><div>Días</div><div>Periodo usado</div><div>Estado</div><div>Acción</div></div>{''.join(cards) or "<div class='sol-empty'>No tienes solicitudes asignadas para aprobar.</div>"}</div></div>
    </section>"""
    return render_page(content, active='Gestion Vacacional')

@app.route('/contratacion/mis_documentos')
@worker_required
def trabajador_contratacion():
    dni=session['dni']; t=get_trabajador(dni)
    with db() as con:
        docs=con.execute('SELECT * FROM contratacion_docs WHERE dni=? ORDER BY id DESC', (dni,)).fetchall()
    rows=''.join([f"<tr><td>{r['tipo_doc']}</td><td>{r['etapa']}</td><td><span class='status-pill'>{r['estado']}</span></td><td>{r['fecha_registro']}</td><td><a class='btn-blue mini-btn' target='_blank' href='/contratacion/ver/{r['id']}'>Ver</a></td></tr>" for r in docs])
    content=f"""
    <div class='hero'><div class='topbar'><div><h1>Gestión de <span class='accent'>Contrato</span></h1><div class='subtitle'>Visualiza y descarga tus contratos, anexos y documentos de incorporación o renovación.</div></div></div></div>
    <section class='grid'><div class='card mini'><div><h3>Trabajador</h3><b>{t['nombre'] if t else dni}</b></div><div class='ico'>👤</div></div><div class='card mini'><div><h3>Empresa</h3><b>{session.get('empresa') or (t['empresa'] if t else '')}</b></div><div class='ico'>🏢</div></div><div class='card mini'><div><h3>Documentos</h3><b>{len(docs)}</b></div><div class='ico'>🧾</div></div>
    <div id='mis-contratos' class='card span-12'><h2>Mis documentos contractuales</h2><div class='table-wrap'><table><tr><th>Documento</th><th>Etapa</th><th>Estado</th><th>Fecha</th><th>Acción</th></tr>{rows or '<tr><td colspan=5>No hay documentos de contratación cargados.</td></tr>'}</table></div></div></section>"""
    return render_page(content, active='Gestion Contratacion')

@app.route('/contratacion/ver/<int:cid>')
@worker_required
def ver_contratacion(cid):
    with db() as con:
        r=con.execute('SELECT * FROM contratacion_docs WHERE id=?', (cid,)).fetchone()
    if not r or r['dni'] != session.get('dni'): abort(404)
    path=Path(r['ruta_archivo'])
    if not path.exists(): abort(404)
    return send_file(path, as_attachment=False, download_name=r['archivo_nombre'])

@app.route('/admin/contratacion', methods=['GET','POST'])
@admin_required
def admin_contratacion():
    if request.method=='POST':
        f=request.files.get('archivo'); dni=normalizar_dni(request.form.get('dni')); trab=get_trabajador(dni); tipo=clean(request.form.get('tipo_doc')); etapa=clean(request.form.get('etapa')) or 'Incorporación'
        if f and f.filename:
            folder=UPLOAD_DIR/'contratacion'/dni; folder.mkdir(parents=True, exist_ok=True)
            name=now_file()+'_'+secure_filename(f.filename); path=folder/name; f.save(path)
            with db() as con:
                con.execute('INSERT INTO contratacion_docs(dni,trabajador,empresa,etapa,tipo_doc,estado,archivo_nombre,ruta_archivo,fecha_registro,uploaded_by) VALUES(?,?,?,?,?,?,?,?,?,?)',(dni, trab['nombre'] if trab else '', trab['empresa'] if trab else '', etapa, tipo, 'Generado', f.filename, str(path), now_txt(), marca_carga(session.get('admin_user','admin')))); con.commit()
            flash('Documento de contratación registrado.','ok')
        return redirect(url_for('admin_contratacion'))
    with db() as con:
        tipos=con.execute('SELECT * FROM contratacion_tipos ORDER BY etapa, descripcion').fetchall()
        docs=con.execute('SELECT * FROM contratacion_docs ORDER BY id DESC LIMIT 300').fetchall()
        trabajadores=con.execute('SELECT dni,nombre FROM trabajadores ORDER BY nombre LIMIT 500').fetchall()
    opt_tipo=''.join([f"<option value='{r['descripcion']}'>{r['codigo']} - {r['descripcion']} ({r['etapa']})</option>" for r in tipos])
    opt_trab=''.join([f"<option value='{r['dni']}'>{r['dni']} - {r['nombre']}</option>" for r in trabajadores])
    tipos_rows=''.join([f"<tr><td>✎ 🗑</td><td><span class='badge-green'>{'Activo' if r['activo'] else 'Inactivo'}</span></td><td>{r['codigo']}</td><td>{r['descripcion']}</td><td>{r['etapa']}</td><td><span class='badge-green'>✓</span></td></tr>" for r in tipos])
    docs_rows=''.join([f"<tr><td>🔍 📄</td><td>{r['dni']}</td><td>{r['trabajador']}</td><td>{r['tipo_doc']}</td><td><span class='badge-green'>F</span></td><td>{r['fecha_registro']}</td></tr>" for r in docs])
    content=f"""
    <div class='hero'><div class='topbar'><div><h1>Gestión <span class='accent'>Contratación</span></h1><div class='subtitle'>Inspirado en Adapta: flujos, carga masiva, maestros, tipos de documento, archivos del trabajador y descargas.</div></div></div></div>
    <section class='grid'><div id='procesos' class='card span-12'><div class='adapta-note'><b>Flujo preparado:</b> trabajador → documentos requeridos por etapa → carga/validación → firmado/archivado → descarga.</div></div>
    <div id='candidatos' class='card span-6'><h2>👥 Candidatos</h2><p class='muted'>Sección preparada para registrar postulantes, evaluación y estado del proceso.</p></div><div id='plantillas' class='card span-6'><h2>📋 Plantilla Documentos</h2><p class='muted'>Sección preparada para formatos de contrato, anexos, acuerdos y documentos requeridos por etapa.</p></div><div id='contratos' class='card span-12'><h2>📄 Subir documento de contratación / Contratos</h2><form method='post' enctype='multipart/form-data' class='form-grid'><div class='field'><label>Trabajador</label><input name='dni' list='trabajadores_list' required><datalist id='trabajadores_list'>{opt_trab}</datalist></div><div class='field'><label>Etapa</label><select name='etapa'><option>Incorporación</option><option>Renovación</option><option>Cese</option></select></div><div class='field'><label>Tipo documento</label><select name='tipo_doc'>{opt_tipo}</select></div><div class='field'><label>Archivo</label><input type='file' name='archivo' required></div><button class='btn-green'>Subir / archivar</button></form></div>
    <div id='maestros' class='card span-12 adapta-table'><h2>🧩 Tipos de documento por etapa</h2><div class='table-wrap'><table><tr><th></th><th>Estado</th><th>Código</th><th>Tipo Doc</th><th>Stage</th><th>Mandatorio</th></tr>{tipos_rows}</table></div></div>
    <div id='descargas' class='card span-12 adapta-table'><h2>⬇️ Archivos trabajador / Descargas</h2><div class='table-wrap'><table><tr><th></th><th>Código</th><th>Apellidos y Nombres</th><th>Tipo Documento</th><th>Estado Doc</th><th>Fecha Envío</th></tr>{docs_rows or '<tr><td colspan=6>No hay archivos.</td></tr>'}</table></div></div></section>"""
    return render_page(content, active='Gestion Contratacion')


@app.route('/admin/crear_carpetas')
@admin_required
def admin_crear_carpetas():
    asegurar_carpetas_documentales()
    total = sincronizar_documentos_carpeta()
    flash(f'Carpeta local creada/actualizada: {DOCUMENTOS_BASE_DIR}. Coloca allí los PDFs y presiona Sincronizar. Documentos detectados automáticamente: {total}.', 'ok')
    return redirect(request.referrer or url_for('admin'))

@app.route('/admin/sincronizar')
@admin_required
def admin_sincronizar():
    resumen = sincronizar_documentos_carpeta(devolver_resumen=True)
    flash(f"Sincronización completada. Nuevos: {resumen['nuevos']} | Revisados: {resumen['revisados']} | Duplicados: {resumen['duplicados']} | Sin DNI: {resumen['sin_dni']} | Sin trabajador activo: {resumen['sin_trabajador']} | Ruta: {DOCUMENTOS_BASE_DIR}", 'ok')
    return redirect(url_for('admin_documentos'))


@app.route('/admin/modo_prueba')
@admin_required
def admin_modo_prueba():
    modo_txt = 'ACTIVO' if modo_prueba_activo() else 'INACTIVO'
    content = f"""
    <div class='hero'><div class='topbar'><div><h1>Modo prueba <span class='accent'>y limpieza general</span></h1><div class='subtitle'>Esta opción aplica a todas las gestiones. Actívala solo cuando vayas a probar con usuarios y administrador.</div></div></div></div>
    <section class='grid'>
      <div class='card span-12'><h2>🧪 Control general de pruebas</h2><p class='muted'>Todo lo cargado quedará marcado como [MODO PRUEBA]. Luego puedes limpiarlo sin tocar la información real.</p><div class='actions'><a class='btn-green' href='/admin/modo_prueba/toggle'>Modo prueba: {modo_txt}</a><a class='btn-danger' onclick='return confirm("¿Borrar documentos y eventos de MODO PRUEBA?")' href='/admin/modo_prueba/limpiar'>Limpiar pruebas</a><a class='btn-blue' href='/admin/desbloquear_usuarios'>Desbloquear usuarios</a></div></div>
    </section>"""
    return render_page(content, active='Modo prueba')

@app.route('/admin/modo_prueba/toggle')
@admin_required
def admin_modo_prueba_toggle():
    set_config('modo_prueba', '0' if modo_prueba_activo() else '1')
    flash('Modo prueba actualizado.', 'ok')
    return redirect(url_for('admin'))

@app.route('/admin/modo_prueba/limpiar')
@admin_required
def admin_modo_prueba_limpiar():
    borrados = 0
    with db() as con:
        rows = con.execute("SELECT id,ruta_archivo FROM documentos WHERE uploaded_by LIKE '%MODO PRUEBA%'").fetchall()
        ids = [r['id'] for r in rows]
        for r in rows:
            try:
                p = Path(r['ruta_archivo'])
                if p.exists() and str(p).startswith(str(UPLOAD_DIR)):
                    p.unlink()
            except Exception:
                pass
        if ids:
            q = ','.join(['?']*len(ids))
            con.execute(f'DELETE FROM eventos_documento WHERE documento_id IN ({q})', ids)
            con.execute(f'DELETE FROM documentos WHERE id IN ({q})', ids)
            borrados = len(ids)
        con.commit()
    flash(f'Modo prueba limpiado. Documentos de prueba borrados: {borrados}.', 'ok')
    return redirect(url_for('admin'))

@app.route('/admin/desbloquear_usuarios')
@admin_required
def admin_desbloquear_usuarios():
    with db() as con:
        con.execute('DELETE FROM login_intentos')
        con.commit()
    flash('Usuarios desbloqueados. Los intentos fallidos fueron reiniciados.', 'ok')
    return redirect(url_for('admin'))

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
