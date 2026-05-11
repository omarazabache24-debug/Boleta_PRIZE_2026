CORRECCION ERROR 500 - CARGA SALDOS VACACIONALES

Se corrigio el error interno al cargar saldos vacacionales.

Cambios:
1. Correccion SQL en actualizacion de jefe inmediato.
2. Carga de saldos mas resistente: ya no depende de ON CONFLICT.
3. Acepta encabezados como DIAS GANADO:, DIAS GANADOS, DÍAS GANADOS.
4. Convierte numeros aunque vengan como texto o con coma decimal.
5. Mantiene vinculo trabajador -> JEFE INMEDIATO por DNI.
6. Migra solicitudes pendientes al jefe correcto despues de cargar saldos.

Subir este ZIP completo a Render/GitHub.
