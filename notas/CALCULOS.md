# CALCULOS — Instrucción 1: caso base (estadero abierto entre semana, 2027)

## Qué hace cada paso
1. **Lee los supuestos** de `datos/supuestos.xlsx` (solo lectura). Ningún valor está escrito en el código.
2. **Días abiertos del mes:** cuenta cuántos miércoles y jueves tiene cada mes de 2027 (patrón `1100111`: 1 = cerrado, 0 = abierto).
3. **Clientes del mes** = días abiertos × personas del área × captura por día abierto, redondeado al entero más cercano (mitades hacia arriba).
4. **Ingresos** = clientes × precio por persona.
5. **Costos variables** = clientes × costo por cliente.
6. **Utilidad del mes** = ingresos − costos variables − costos fijos adicionales.
7. **Punto de equilibrio** = costos fijos ÷ (precio − costo por cliente) = 9.000.000 ÷ 27.000 = 333,3, es decir 334 clientes al mes.
8. **Reserva necesaria** = el valor más negativo de la utilidad acumulada desde enero, en positivo (0 si nunca es negativa).

## Cifras de control (valores inventados)
- Clientes del año: 5.200 · Utilidad del año: 32.400.000 · Equilibrio: 334 clientes al mes · Meses con pérdida: 0 · Reserva: 0.
- Verificado: Python, `modelo.xlsx` (93 fórmulas, 0 errores) y dashboard muestran las mismas cifras; los 12 meses suman los totales.

## Archivos creados
`BRIEF.md`, `datos/supuestos.xlsx`, `calculo/modelo.py`, `salidas/modelo.xlsx`, `salidas/index.html`, `notas/CALCULOS.md`.

## Supuestos que parecen extraños (sin cambiarlos)
- Los 6 valores son inventados; no hay ningún Dato con fuente pública. La población del área (50.000) debería salir del DANE o del propietario.
- Los meses no varían por temporada: solo cambia el número de días abiertos. Con datos reales conviene un factor de temporada.
- La reserva no incluye inversión inicial (adecuaciones, contratación previa).
- Falta `assets/` con el escudo y la paleta: el dashboard usa colores propios.

## Uso de IA
La IA escribió el modelo, el Excel y el dashboard. Los supuestos y la decisión son de quien usa el análisis.

# Instrucción 2 — ¿Le creemos al pronóstico?
Historia de clientes por año (INVENTADA, `datos/limpios/historia-clientes.csv`): 2023 = 41.000; 2024 = 47.000; 2025 = 49.000.
Se finge estar a finales de 2024 y se predice 2025 de dos formas, comparando con el real (49.000):
- Seguir la tendencia (2024 × 2024 ÷ 2023) = 53.878 → error +10,0 %.
- Repetir 2024 = 47.000 → error −4,1 %.
Conclusión: el método más simple falló menos, por eso el caso base no supone crecimiento. Con una sola prueba y datos inventados, esto es solo ilustrativo.

# Instrucción 3 — Tres escenarios
Solo cambia la demanda (−30 %, 0 %, +20 %, leídas de `supuestos.xlsx`); precio, costos y días no cambian. Hoja «Escenarios» de `modelo.xlsx`, con fórmulas.
| Escenario | Utilidad del año | Meses con pérdida | Reserva necesaria |
|---|---|---|---|
| Pesimista (−30 %) | −9.720.000 | 10 | 10.170.000 |
| Base | 32.400.000 | 0 | 0 |
| Optimista (+20 %) | 60.480.000 | 0 | 0 |
La reserva es el punto más bajo de la utilidad acumulada desde enero. Verificado: Python y Excel coinciden (177 fórmulas, 0 errores) y el escenario base iguala el caso base.

# Instrucción 4 — ¿Qué supuesto pesa más?
Cada supuesto se mueve al mínimo y al máximo de su «Rango posible», uno a la vez, con los demás en el caso base (cambio en la utilidad del año):
| Supuesto | Rango | Si sale mal | Si sale bien |
|---|---|---|---|
| Captura | ± 50 % | −70.200.000 | +70.200.000 |
| Personas en el área | ± 25 % | −35.046.000 | +35.154.000 |
| Precio por persona | ± 10 % | −23.400.000 | +23.400.000 |
| Costos fijos al mes | ± 10 % | −10.800.000 | +10.800.000 |
| Costo por cliente | ± 10 % | −9.360.000 | +9.360.000 |
La captura pesa más y es la más débil (inventada, sin evidencia), pero se puede medir antes de invertir con un piloto, una preventa o una encuesta. Estas cifras se calculan en `modelo.py`; no están en `modelo.xlsx`.

# Instrucción 5 — ¿Y si…?
Cinco controles (variación de demanda de −40 % a +20 %, captura, precio, costo y costos fijos) recalculan en vivo la utilidad, los meses con pérdida y la gráfica. Pruebas: valores iniciales = 32.400.000; demanda −30 % = −9.720.000 con 10 meses en pérdida (igual a `modelo.xlsx`); precio +1.000 sube la utilidad. El botón «Volver al caso base» restablece todo.

# Instrucción 6 — Decisión
Sección al comienzo del dashboard y `salidas/MEMO.md` (mismo texto), solo con cifras de `modelo.xlsx`. Se agregó la «Captura de equilibrio» (0,077 %) en la hoja Cálculo. Recomendación: abrir con un piloto de tres meses; reserva necesaria 10.170.000 (escenario pesimista). Excel: 178 fórmulas, 0 errores.
