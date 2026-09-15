# AGSA — herramientas HTML

Aplicaciones de una sola página, sin backend: cada archivo `.html` es autónomo y se abre
con doble clic. Los datos se guardan en el `localStorage` del navegador.

- `Carga_Capacidad_AGSA.html` — carga de capacidad de las líneas de tubo.
- `ABC_Estrategias_AGSA.html` — cálculo ABC y estrategias MTS / ATO / MTO.

## Reglas para `ABC_Estrategias_AGSA.html`

1. **Los datos de AGSA van incrustados en el archivo y no se pierden.**
   Viven en `<script id="agsa-seed" type="application/json">` (maestro de materiales,
   familias/redondos, reglas de medida y pedidos de venta reales, en formato compacto `v:2`).
   Al abrir el archivo por primera vez se cargan solos, y el botón «Restaurar datos de AGSA»
   de la pantalla de Inicio los vuelve a dejar como están.
   **Cualquier HTML nuevo de este proyecto debe conservar ese bloque tal cual**; para
   actualizarlo, exportar un backup desde la propia herramienta (ya usa el mismo formato
   compacto) y sustituir el contenido del script.

2. **Las cantidades de los pedidos están en toneladas** y se muestran con dos decimales.
   Los precios son €/t y los pesos calculados (Suma de Peso, Tn Mes, Tn Semana, percentiles
   de MMPP) van también en toneladas.

3. **El periodo de cálculo es común**: se elige en Parametrización ABC, Calculos MMPP Tubos
   y Calculos Tubos (fecha de pedido o de entrega y meses desde/hasta) y se guarda en
   `PARAM().periodo`.

4. **Nomenclatura del maestro**: el ABC del maestro es el código de tres posiciones (`Aa1`)
   y respeta mayúsculas y minúsculas; las estrategias del maestro son `TMTS` / `TATO` /
   `TMTO` y se comparan sin la `T` inicial.

5. El `localStorage` guarda el estado con la misma serialización compacta; si no cabe,
   se guarda todo menos los pedidos (que vuelven del bloque incrustado) y se avisa.

## Comprobación

No hay build ni tests automáticos: el archivo se valida abriéndolo en un navegador
(Playwright con `PLAYWRIGHT_BROWSERS_PATH=/opt/pw-browsers`) y comprobando que no haya
errores de consola, que los cálculos cuadren y que las pantallas rendericen.
