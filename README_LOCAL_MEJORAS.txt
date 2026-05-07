MEJORAS APLICADAS
=================
1. Dashboard administrador con rango Desde/Hasta para visualizar indicadores.
2. Botón Modo prueba y botón Limpiar pruebas.
   - Cuando el modo prueba está ACTIVO, las cargas quedan marcadas como [MODO PRUEBA].
   - Limpiar pruebas borra solo documentos/eventos marcados como prueba.
3. Bloqueo de trabajadores por 3 intentos fallidos. El admin no se bloquea.
   - Desde Dashboard: botón Desbloquear usuarios.
4. Botón Actualizar / detectar PDFs.
   - Detecta documentos desde DOCUMENTOS_PRIZE_AUTO junto al app.py.
   - Ejemplo:
     DOCUMENTOS_PRIZE_AUTO\DOCUMENTOS DE PAGO\BOLETAS NORMAL\SEMANAL
   - Solo indexa PDFs/archivos permitidos si detecta DNI y el trabajador está activo.
5. Tabla con detalle de quién cargó el documento.

IMPORTANTE PARA USO LOCAL
=========================
Ejecuta app.py desde la carpeta donde está este proyecto. El sistema creará/usará:
DOCUMENTOS_PRIZE_AUTO
