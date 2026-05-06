# Portal de Documentos PRIZE

## Archivos para subir a GitHub / Render
- app.py
- requirements.txt
- Procfile
- runtime.txt
- logo_prize.png

## Render
Build Command:
```bash
pip install -r requirements.txt
```
Start Command:
```bash
gunicorn app:app
```

## Variables opcionales
- SECRET_KEY: clave secreta del sistema
- PERSIST_DIR: /data si usas disco persistente en Render
- APP_TIMEZONE: America/Lima

## Accesos demo
Administrador: admin / admin123
Trabajador: DNI 74324033 / correo omar@demo.com

## Mejoras incluidas
- Panel izquierdo contraíble/expandible.
- Grupos de menú de pago, empresa y personales que muestran/ocultan submenús.
- Se conserva pestaña activa y scroll del panel al navegar.
- Login oscuro con colores PRIZE.
- Logo automático desde la carpeta del proyecto.
- Filtro por tipo y periodo.
- Carga de documentos personales y documentos de empresa.
