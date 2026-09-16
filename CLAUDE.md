# AGSA — herramientas HTML

Aplicaciones de una sola página, sin backend: cada archivo `.html` es autónomo y se abre
con doble clic. Los datos se guardan en el `localStorage` del navegador.

- `Carga_Capacidad_AGSA.html` — carga de capacidad de las líneas de tubo.
- `ABC_Estrategias_AGSA.html` — cálculo ABC y estrategias MTS / ATO / MTO.

## Reglas para `ABC_Estrategias_AGSA.html`

1. **El maestro de AGSA va incrustado en el archivo y no se pierde.**
   Vive en `<script id="agsa-seed" type="application/json">`: maestro de materiales,
   familias/redondos y reglas de medida (formato compacto `v:2`). Se carga solo al abrir
   el archivo por primera vez y el botón «Restaurar maestro de AGSA» de Inicio lo repone.
   **Cualquier HTML nuevo de este proyecto debe conservar ese bloque tal cual**; para
   actualizarlo, exportar un backup desde la herramienta (usa el mismo formato) y sustituir
   el contenido del script, quitando los inputs.
   Los **pedidos, previsiones y fabricaciones no se incrustan**: se cargan desde Excel en
   cada sesión y se conservan en el navegador mientras quepan.

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

6. **Temporales**: cada medida tiene, **por máquina**, un mínimo de producción (30 por
   defecto, 100 en TUBOS_17), un temporal y un periodo de repetición (1 cada ciclo,
   2 cada dos, 3 cada tres, 0 nunca), guardados en `regla.cfg[maquina]`. Los colores del
   subciclo salen de ahí: verde 1, amarillo 2, naranja 3, rojo 0.
   En «Calculo Temporales», **Subciclos variable 30d** y **Revisar Temporal** usan fórmulas
   **provisionales** (total 30d por debajo del mínimo, y desviación de la fabricación MTS
   frente al temporal superior al 10 %) a la espera de las definitivas del cliente.

## Comprobación

No hay build ni tests automáticos: el archivo se valida abriéndolo en un navegador
(Playwright con `PLAYWRIGHT_BROWSERS_PATH=/opt/pw-browsers`) y comprobando que no haya
errores de consola, que los cálculos cuadren y que las pantallas rendericen.
