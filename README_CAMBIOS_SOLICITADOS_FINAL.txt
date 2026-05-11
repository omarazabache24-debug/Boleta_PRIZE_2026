CAMBIOS APLICADOS - PRIZE RRHH

1. Selector de empresa del trabajador:
   - Se quitó PRIZE SUPERFRUITS.
   - Ahora solo muestra la empresa cargada en la columna EMPRESA del trabajador.
   - Si existían datos antiguos PRIZE SUPERFRUITS en la BD demo, se normalizan a AQUANQA.

2. Menú del usuario trabajador:
   - Se reorganizó para que tenga Gestión Documental.
   - Dentro de Gestión Documental están Documentos de pago, Documentos de la empresa y Documentos personales.
   - Se mantienen Gestión Vacacional y Gestión Contrato como acciones del trabajador.

3. Gestión vacacional del trabajador:
   - Antes de registrar solicitud, valida saldo disponible.
   - Si tiene saldo suficiente, registra como solicitud normal.
   - Si no tiene saldo, no registra, salvo que marque Adelanto de vacaciones.
   - Las solicitudes por adelanto quedan identificadas en estado y comentario.

4. Corrección de navegación:
   - Se ajustó el menú para que cada pestaña permanezca en su módulo correspondiente y no salte a otra sección.

5. Plantilla de saldos:
   - Se mantiene PERIODO en formato 2025/2026.
   - No se usa días gozados en plantilla de carga.
