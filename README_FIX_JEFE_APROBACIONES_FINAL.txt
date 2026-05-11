FIX FINAL - APROBACIONES JEFE INMEDIATO

1. La solicitud guarda jefe_dni usando prioridad:
   - Jefe inmediato del periodo seleccionado en vacaciones_saldos.
   - Cualquier saldo cargado del trabajador.
   - Ficha de trabajadores.

2. El panel /vacaciones/aprobaciones_jefe ahora busca solicitudes por:
   - vacaciones_solicitudes.jefe_dni
   - trabajadores.jefe_dni
   - vacaciones_saldos.jefe_dni

3. Se agregó reparación automática para solicitudes antiguas sin jefe_dni.

4. Se mantiene bloqueo de fechas pasadas al registrar vacaciones.
