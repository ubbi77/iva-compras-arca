# IVA Compras · TIEMPO → ARCA

Un solo archivo: **IVA_Compras_TIEMPO_ARCA.html**. Se abre con doble clic (Chrome o Edge).
No necesita internet ni instalar nada. Nada sale de la PC.

## Uso
1. **Abrir Excel TIEMPO** (o arrastrarlo a la ventana): el .XLS de compras que exporta TIEMPO.
2. Revisar la planilla. Todas las celdas se pueden editar. Punto de color por fila:
   verde = OK · naranja = aviso · rojo = error · gris = excluido. Clic en el punto = detalle
   con los avisos y las líneas exactas que van al TXT.
3. **Generar TXT ARCA** → muestra totales por tipo y por alícuota y descarga los 2 archivos
   (`libro_iva_digital_compras_cbte_AAAAMM.txt` y `..._alicuotas_AAAAMM.txt`).

El trabajo se guarda solo en el navegador de esa PC. Para pasarlo a otra PC o tener copia:
**Guardar trabajo** (.json) y después **Abrir trabajo**. **Abrir TXT ARCA** permite volver a
abrir un libro ya generado para corregirlo.

## Reglas de conversión (verificadas con septiembre 2026)
| TIEMPO | ARCA |
|---|---|
| FC A/B/C, NC A, LI A, TK A | 001 / 006 / 011, 003, 063, 081 |
| columna1 | Neto gravado (si hay 2 alícuotas se separa desde el IVA de cada una) |
| columna2 / 3 / 4 | IVA 21% / 10,5% / 27% |
| columna5 | Exento · columna6 Perc. IIBB · columna8 Imp. municipales |
| columna9 | No gravado (código de operación "N") |
| columna11 | Percepción IVA · columna12 Imp. internos · columna14 Perc. otros imp. nacionales |
| columna16 | Total |
| LI con CUIT del exterior (50-…) | Se excluye (despacho de importación: ARCA ya lo trae) |

Comprobante A sin IVA → alícuota 0% (0003). B y C → sin alícuotas. Notas de crédito en positivo.

## Para desarrollo
`core.js` (lógica) + `app.html` (pantalla). `python3 build.py` los une en el .html final.
