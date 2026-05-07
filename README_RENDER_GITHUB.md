# Portal de Documentos PRIZE

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
- `SECRET_KEY`: clave segura de sesión.
- `PERSIST_DIR=/data`: usar si tienes Render Disk.
- `DOCUMENTOS_BASE_DIR=/data/documentos_auto`: carpeta para importar documentos automáticamente.
- `APP_TIMEZONE=America/Lima`

## Login demo
Trabajador:
- DNI: `74324033`
- Correo: `omar@demo.com`

Admin:
- Usuario: `admin`
- Clave: `admin123`

## Importación automática por carpeta
Coloca PDFs, imágenes o archivos Office dentro de `documentos_auto` o define `DOCUMENTOS_BASE_DIR`.
El sistema detecta DNI de 8 dígitos en el nombre o ruta del archivo.

Ejemplos:
```text
documentos_auto/Utilidades/2026/74324033_boleta_utilidades_2026.pdf
documentos_auto/CTS/05_2026/74324033_cts.pdf
documentos_auto/Reglamento_SST/reglamento_sst.pdf
```

También puedes entrar como admin y usar **Sincronizar carpeta**.
