# Portal de Documentos PRIZE

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
- SECRET_KEY
- PERSIST_DIR=/data
- APP_TIMEZONE=America/Lima
- DOCUMENTOS_BASE_DIR=/data/documentos_auto

Demo trabajador: DNI 74324033 / omar@demo.com
Admin: admin / admin123
