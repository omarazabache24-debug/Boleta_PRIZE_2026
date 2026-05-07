PORTAL DE DOCUMENTOS PRIZE

Login administrador:
usuario: admin
clave: admin123

Login trabajador demo:
DNI: 74324033
correo: omar@demo.com

Variables opcionales en Render:
SECRET_KEY=coloca_una_clave_larga
PERSIST_DIR=/data
APP_TIMEZONE=America/Lima
DOCUMENTOS_BASE_DIR=/data/documentos_auto

Para cargar documentos automáticamente por carpetas:
- Coloca PDFs/archivos en una carpeta llamada documentos_auto junto al app.py, o configura DOCUMENTOS_BASE_DIR.
- El sistema detecta DNI de 8 dígitos en el nombre o ruta del archivo.
- También infiere tipo por palabras en carpeta/archivo: utilidad, vacaciones, cts, gratificacion, liquidacion, normal, contrato, reglamento, politicas, etc.
