FIX FECHA SOLICITUD VACACIONES DEFINITIVO

- Se corrigió la validación del navegador para que NO muestre alerta cuando el campo date está incompleto.
- Se valida correctamente YYYY-MM-DD y también DD/MM/AAAA si algún navegador lo permite.
- La comparación se hace siempre contra formato ISO YYYY-MM-DD para evitar falsos bloqueos.
- Se mantiene validación del servidor: no registra inicio o fin menor a la fecha actual de Perú.
