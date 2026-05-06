# BOLETAS PRIZE - Portal Documentario PRO

## Archivos para subir a GitHub
- app.py
- requirements.txt
- Procfile
- runtime.txt
- logo_prize.png
- carpeta static/logo_prize.png

## Render
Build Command:
```bash
pip install -r requirements.txt
```

Start Command:
```bash
gunicorn app:app
```

## Variables recomendadas
- SECRET_KEY = clave_segura
- APP_TIMEZONE = America/Lima
- EMPRESA_NOMBRE = PRIZE SUPERFRUITS
- DATABASE_URL = si usas PostgreSQL de Render

## Accesos demo
Admin:
- usuario: admin
- clave: admin123

Trabajador demo:
- DNI: 74324033
- correo: omar@demo.com

## Mejoras incluidas
- Panel izquierdo tipo Nisira/DMHT.
- Documentos de pago: utilidades, vacaciones, normal, constancia gratificación, CTS, liquidación, gratificación.
- Documentos empresa y documentos personales.
- Login oscuro con logo PRIZE.
- Reconoce logo automáticamente si existe como logo_prize.png, logo.png, prize.png o PRIZE.png en la misma carpeta del app.py, static, uploads o persistencia.
- Sin pandas para evitar error de compilación en Render.
