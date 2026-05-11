MEJORAS APLICADAS - PERSISTENCIA Y EXCEL LOCAL

1. Se agregó carpeta REGISTROS_EXCEL_LOCAL.
2. Cada carga o actualización de trabajadores, saldos y solicitudes actualiza automáticamente:
   - 01_TRABAJADORES_LOCAL.xlsx
   - 02_VACACIONES_SALDOS_LOCAL.xlsx
   - 03_VACACIONES_SOLICITUDES_LOCAL.xlsx
3. Si la BD queda vacía, el sistema intenta recuperar trabajadores desde 01_TRABAJADORES_LOCAL.xlsx.
4. La pantalla Mis solicitudes fue rediseñada con tarjetas oscuras y encabezado profesional.
5. Fecha ingreso ahora se muestra sin 00:00:00.
6. Para Render: si usas disco persistente, configurar PERSIST_DIR=/data. En modo local queda en la misma carpeta del proyecto.
