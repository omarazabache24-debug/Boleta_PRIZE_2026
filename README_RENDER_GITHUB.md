# PRIZE - Portal Web/App de Boletas

Proyecto integrado listo para GitHub y Render.

## Archivos principales
- `app.py`: sistema completo integrado.
- `requirements.txt`: dependencias.
- `Procfile`: comando Render.
- `runtime.txt`: versión Python sugerida.

## Render
Build command:
```bash
pip install -r requirements.txt
```
Start command:
```bash
gunicorn app:app
```

## Variables recomendadas
- `SECRET_KEY`: clave segura del sistema.
- `APP_TIMEZONE`: `America/Lima`
- `EMPRESA_NOMBRE`: `PRIZE SUPERFRUITS`
- `DATABASE_URL`: opcional, si conectas PostgreSQL en Render.
- `PERSIST_DIR`: opcional. En Render con disco persistente usa `/data`.

## Login demo
Admin:
- usuario: `admin`
- clave: `admin123`

Trabajador demo:
- DNI: `74324033`
- correo: `omar@demo.com`

## Carga de trabajadores
Desde el panel Admin > Trabajadores, sube un Excel con estas columnas:
- DNI
- NOMBRE o TRABAJADOR
- CORREO
- CARGO
- AREA
- EMPRESA
- PLANILLA

## Carga de PDFs
Desde Admin > Boletas, sube PDFs nombrados con DNI, por ejemplo:
- `74324033.pdf`
- `boleta_74324033_utilidad.pdf`

## Endpoints API integrados
- `/api/health`
- `/api/login`
- `/api/boleta/<dni>`
- `/api/pdf/<dni>`
