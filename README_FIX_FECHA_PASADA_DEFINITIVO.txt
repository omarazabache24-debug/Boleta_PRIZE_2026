FIX DEFINITIVO - BLOQUEO DE FECHAS PASADAS EN VACACIONES

1. Se corrigió la validación del registro de solicitudes de vacaciones.
2. Ahora el servidor bloquea cualquier FECHA INICIO menor a la fecha actual de Perú.
3. También bloquea FECHA FIN menor a la fecha actual de Perú.
4. Se agregó doble validación antes de guardar en la base de datos.
5. Se agregó bloqueo visual en el formulario HTML con min dinámico y validación JavaScript.
6. Si existían solicitudes antiguas con fecha inicio anterior a hoy, quedan anuladas automáticamente para que no consuman saldo.

Resultado: aunque el usuario modifique el HTML del navegador, escriba manualmente o intente forzar el POST, la solicitud no se registra si la fecha es anterior a hoy.
