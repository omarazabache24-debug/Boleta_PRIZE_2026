MEJORAS APLICADAS - VACACIONES / JEFE INMEDIATO

1. Plantilla de saldos vacacionales actualizada:
   - Se quitó FECHA INGRESO de la plantilla.
   - JEFE INMEDIATO debe colocarse como DNI del jefe.
   - El sistema detecta el nombre del jefe desde la base de trabajadores.

2. Carga de saldos:
   - Lee EMPRESA, DNI, TRABAJADOR, AREA, JEFE INMEDIATO, PERIODO, DIAS GANADOS y SALDO.
   - La fecha de ingreso ya no se toma de la plantilla de vacaciones; se toma desde la base de trabajadores.

3. Solicitudes del trabajador:
   - Al registrar vacaciones se guarda el DNI del jefe inmediato.
   - Si no tiene saldo o si solicita más días que el saldo disponible, no se registra.
   - En el portal del trabajador se muestra la fecha de ingreso.

4. Aprobaciones del jefe inmediato:
   - Si el usuario logueado es jefe inmediato, verá una sección para aprobar o rechazar solicitudes de sus trabajadores.
   - Si no es jefe, solo verá el seguimiento de sus propias solicitudes.
   - Si el jefe también solicita vacaciones, podrá ver su propio seguimiento en “Mis solicitudes”.

5. Flujo:
   - Trabajador registra solicitud.
   - Jefe inmediato aprueba o rechaza.
   - Si aprueba, queda Pendiente GTH.
   - GTH/Admin aprueba final.
