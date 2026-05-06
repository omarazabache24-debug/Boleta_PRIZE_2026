# Portal de Documentos PRIZE - Render/GitHub

## Archivos
- `app.py`: aplicación completa Flask.
- `logo_prize.png`: logo detectado automáticamente.
- `requirements.txt`: dependencias livianas sin pandas.
- `Procfile`: comando Render.
- `runtime.txt`: versión Python.

## Render
Build command:
```bash
pip install -r requirements.txt
```
Start command:
```bash
gunicorn app:app
```

Variables opcionales:
- `SECRET_KEY`: clave privada.
- `PERSIST_DIR=/data` si usas disco persistente.
- `APP_TIMEZONE=America/Lima`

## Accesos demo
Administrador: `admin` / `admin123`
Trabajador: DNI `74324033` / correo `omar@demo.com`

## Excel de trabajadores
Columnas recomendadas: `DNI`, `NOMBRE`, `CORREO`, `CARGO`, `AREA`, `EMPRESA`.
