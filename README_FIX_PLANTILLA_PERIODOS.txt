CORRECCIÓN APLICADA - PLANTILLA SALDOS VACACIONALES

1. La plantilla de carga queda con estas columnas exactas:
   EMPRESA | DNI | TRABAJADOR | AREA | JEFE INMEDIATO | I_PERIODO | F_PERIODO | DIAS GANADOS | SALDO

2. I_PERIODO y F_PERIODO ahora se manejan solo como años.
   Ejemplo válido:
   I_PERIODO: 2025
   F_PERIODO: 2026

3. Se retiró el ejemplo anterior con fechas completas 2025-01-01 / 2026-12-31.

4. El sistema también normaliza la carga: si por error alguien sube una fecha, internamente toma solo el año para evitar errores.
