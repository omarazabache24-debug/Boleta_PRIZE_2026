CAMBIOS IMPLEMENTADOS - PRIZE RRHH

1. Login trabajador:
   - Usuario: DNI.
   - Clave: fecha de nacimiento sin barras, formato ddmmaaaa.
   - Ejemplo: 01/02/1990 => 01021990.

2. Selección de empresa:
   - Luego del login se muestra pantalla para elegir empresa según la columna EMPRESA del trabajador.
   - La empresa queda guardada en sesión y se muestra en el panel.

3. Plantilla de saldos vacacionales:
   - Se retiró DIAS GOZADOS.
   - Columnas actuales: EMPRESA, DNI, TRABAJADOR, AREA, JEFE INMEDIATO, FECHA INGRESO, PERIODO INICIO, PERIODO FIN, DIAS GANADOS, SALDO.
   - El periodo se arma como INICIO al FIN.

4. Corrección de navegación:
   - Gestión vacaciones queda separada de boletas de vacaciones.
   - Gestión contrato queda activa para usuario trabajador.
   - Botón Inicio permanece en /panel, no regresa al login salvo que no haya sesión.

5. Portal trabajador:
   - Gestión vacaciones: visualiza saldo, periodo y registra solicitud.
   - Gestión contrato: visualiza/descarga contratos y anexos cargados por el administrador.

6. Seguridad:
   - Se mantiene bloqueo por 3 intentos fallidos.
   - Se conserva el acceso administrador.
