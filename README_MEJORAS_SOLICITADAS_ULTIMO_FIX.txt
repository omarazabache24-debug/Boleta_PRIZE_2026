MEJORAS APLICADAS

1. Solicitud vacacional:
   - Ya no permite registrar vacaciones con fecha inicio menor a la fecha actual.
   - Ya no permite fecha fin menor a fecha inicio.
   - Mantiene validación de saldo: si no tiene saldo o pide más días que el saldo, no registra.

2. Plantilla de saldos vacacionales:
   - Se quitó la columna PERIODO.
   - Se agregaron columnas I_PERIODO y F_PERIODO.
   - No se usa FECHA INGRESO en la plantilla; la fecha se toma desde la base de trabajadores y se muestra en el portal del trabajador.
   - JEFE INMEDIATO debe ir con DNI y se detecta desde la base de trabajadores.

3. Inicio / Volver:
   - Se reforzó /panel para evitar Internal Server Error si el trabajador no existe o perdió sesión.
   - Ahora redirige de forma controlada y muestra aviso.

4. Dashboards:
   - Se agregaron accesos tipo dashboard para Gestión Documental, Gestión Vacacional y Gestión Contrato en el portal usuario.
   - Se mejoraron enlaces del menú lateral del trabajador para mostrar dashboard y accesos por gestión.
