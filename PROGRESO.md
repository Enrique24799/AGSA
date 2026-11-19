# Progreso — AGSA

## Ahora mismo
- Repositorio recién organizado (2026-11-19).

## Pendiente / próximos pasos
- (añadir aquí)

## Origen de los archivos
Repo personal `Enrique24799/AGSA`: ramas main, abc-production-strategies-tool, funny-rubin y happy-hawking.

## Historial (resumen de conversaciones anteriores)
**[17–18/06] Carga de Capacidad AGSA** (`Carga_Capacidad_AGSA.html`)
- Herramienta de carga de capacidad de las líneas de tubo, adaptada de la de Ondupet.
- La tabla «Detalle por Máquina» se sustituye por un mapa de calor de saturación.

**[15–16/09] ABC y estrategias** (`ABC_Estrategias_AGSA.html`, rama `abc-production-strategies-tool`)
- Cálculo ABC y estrategias MTS / ATO / MTO, con el maestro de AGSA incrustado en el HTML (`agsa-seed`). Cualquier HTML nuevo debe conservar ese bloque.
- Cantidades en toneladas, con dos decimales; precios en €/t.
- Periodo de cálculo común a todas las pantallas (`PARAM().periodo`).
- Nomenclatura: ABC de tres posiciones (`Aa1`); estrategias `TMTS` / `TATO` / `TMTO`.
- Temporales por medida y máquina: mínimo de producción (30, o 100 en TUBOS_17), temporal y periodo de repetición (1, 2, 3 o 0). Colores: verde 1, amarillo 2, naranja 3, rojo 0.
- Fórmulas provisionales: «Subciclos variable 30d» y «Revisar Temporal» (desviación MTS frente al temporal > 10 %).

**[02–07/10] Definiciones de planificación AGSA** (`docs/Definiciones_Planificacion_AGSA.docx`)
- **Fecha errónea.** Se compara la fecha de entrega con una fecha propuesta, aplicando la primera regla que se cumpla:
  1. Stock libre ≥ 90 % el día de creación: creación + 2 días de margen logístico.
  2. Familia planificada: fin de la familia en su primer ciclo, con un mínimo de 5 días, + 2.
  3. Familia sin planificar y pedido ≤ 30 Tn: fecha errónea «Sí».
  4. Familia sin planificar y pedido > 30 Tn: fin del ciclo actual (si le quedan ≥ 5 días) o del siguiente, + 2.
  - Resultado «Sin datos» si no hay foto de stock, el material no tiene familia o la máquina no tiene ciclo.
  - Campo `Origen_Propuesta` con la regla aplicada.
- **Tipo de atraso:** Expedición, Comercial, Planificación, Producción, Pendiente de planificar o Sin Atraso. Se calcula con stock reservado (≥ 90 %), fecha de reserva, fecha de disponibilidad y primera fecha de disponibilidad.
- La fecha errónea se fija el día en que entra el pedido y no se recalcula.

**[08/10] CMI – cartera de AGSA en Qlik** (`Cartera_AGSA.qvd`)
- Primera disponibilidad, fecha de reserva y fecha propuesta en la cartera.
- La cartera guarda una foto diaria (`Fecha_Dato`); hay que filtrar por la última.
- Diagnóstico de claves con `MSKA` y `AtrasoPedidoPlanMaster`.
- `vCorteCartera` permite empezar la serie limpia.

**[22/09–11/11] Facturación AGSA en Qlik**
- Proyección de facturación con expedido, cartera, previsión y Peso MTO.
- Tolerancia de tubo `vTolTubo` (0,06 → 0): con 0, el pendiente de facturación es igual a la cartera.
- Materiales ficticios `PESO_MTO_TUBO`, `PESO_MTO_CHAPA` y `PESO_MTO_PERFIL`, para que el Peso MTO filtre por `Tipo_Material_2`.
- `Familia_Linea` queda solo como clave interna del accuracy.
- Centros C412 y C414.

**[28/10–19/11] Ciclos de 30 a 21 días en Tubos** (rama `happy-hawking`, carpeta `analisis/`)
- Metodología (Word), Excel de coste de transformación y escenarios HTML para M8, M16, M17 y M20 (M28 fuera).
- 12 → 17,14 ciclos al año. Las horas de cambio pasan de 2.026 a 2.894 (+868 h).
- **Escenario 1** (equipos actuales): 71,31 → 77,15 €/t (+8,2 %), con 275 k€ de coste fijo sin absorber.
- **Escenario 2** (refuerzo de turnos): 71,31 + X / 51.110 €/t. Sale más barato mientras el refuerzo cueste menos de unos 298 k€ al año (344 €/h).
- **Temporales T1 2027** (`Temporales_2027_T1.xlsx`):
  - el espesor 3 mm de 60x40 y 50x50 pasa de M8 a M28; en M8 se quedan los > 3 mm;
  - las medidas de M16 pasan a M28;
  - Z180 se traslada a Z200 solo en espesores de 2 y 3 mm.
- **Mínimos de fabricación** (`Minimos_Fabricacion.html`):
  - coste del cambio = horas × operarios × €/h + merma de arranque × (materia prima − chatarra);
  - lote mínimo = el mayor entre el mínimo por margen y el mínimo por merma, redondeado a múltiplos de 5;
  - pendiente: valores reales y criterio definitivo (margen neto o de contribución).

**[17/11] Excel: comparar fecha de albarán y fecha de entrega**
- `=ENTERO(I2)` para quitar la hora a la fecha de albarán.
- `FECHA(DERECHA;EXTRAE;IZQUIERDA)` para pasar a fecha real la de entrega que está como texto.
- Diferencia en días y etiqueta «En plazo» / «Retraso».

**[18/11] SAP: sección de cada ubicación**
- Se recomendó el campo estándar Área de almacenamiento (`LGBER`), en WM (LS01N, LS02N, LS11) o EWM.
- Alternativas: área de abastecimiento a producción (PK05) o un campo Z.
- Pendiente: confirmar si se usa WM, EWM o solo almacenes de MM.
