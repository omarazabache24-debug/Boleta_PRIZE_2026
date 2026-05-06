# BOLETAS PRIZE - Render/GitHub corregido

## Corrección aplicada
Se retiró `pandas` porque Render intentaba compilarlo con Python 3.14 y fallaba durante el build. La carga Excel ahora usa `openpyxl`, que no requiere compilar pandas.

## Archivos obligatorios
- `app.py`
- `requirements.txt`
- `Procfile`
- `runtime.txt`
- `.python-version`

## Render
Build Command:
```bash
pip install -r requirements.txt
```

Start Command:
```bash
gunicorn app:app
```

Variables recomendadas:
- `SECRET_KEY`: cualquier clave segura
- `APP_TIMEZONE`: `America/Lima`
- `EMPRESA_NOMBRE`: `PRIZE SUPERFRUITS`

## Acceso demo
Admin:
- usuario: `admin`
- clave: `admin123`

Trabajador demo:
- DNI: `74324033`
- correo: `omar@demo.com`
