# Memoria de conversaciones con Claude Code

Resumen de lo definido en cada conversación (junio–noviembre de 2026), agrupado por sección.
Fuentes: la última respuesta de cada conversación y los documentos y README de los repositorios
AGSA, PERSEIDA, FERROMALLAS y ONDUPET. Fecha de elaboración: 19/11/2026.

> No incluye nombres de operarios, bajas ni datos de colaboradores concretos: el repositorio es público.

**Mapa de sistemas por empresa** (definido en la presentación al CIO, 11–12/11/2026):

| Empresa | ERP | Planificación (GP) |
|---|---|---|
| Perseida | SAP | GPP (secuenciación, notificación, camiones, pantalla de cambios, carga de capacidad, propuesta de camiones, PRAE) |
| AGSA | SAP | GPB (secuenciación, notificación, camiones, pantalla de cambios, algoritmo de combinación). Aún sin carga de capacidad, propuesta de camiones ni PRAE |
| Balboa | GESAG | GPB |
| Ondupet / Ondupack | — | GIO (escritorio) |
| Ferromallas y Lasao | GESAG | GPx previsto para final de 2026 |

---

## 1. AGSA (tubos)

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

---

## 2. PERSEIDA (cosmética / higiene; SAP + GPP)

**[15/06] Gantt de tiempos de cambio** (`gantt-tiempo-cambio.html`)
- Una fila por línea; turnos M 06–14, T 14–22 y N 22–06.
- Cambio de producto y cambio de formato; filtro por sección (Cosmética / Higiene).

**[15/06] Pantalla de ventas, propuesta de carga** (`OUTPUT_PANTALLA_VENTAS.html`)
- Asigna pedido-posición a camión, con estado Completo / Parcial / Backlog.
- Modo «solo stock» o «stock + planificado».
- Volumen del camión = nº total de palets.

**[16/06] Ajuste de palets** (`AJUSTE_PALETS.html`, transacción `ZPAL_AJUS`) y **seguimiento de palets por cliente**
- Los palets libres se agrupan por código de palet (CLP).
- Acciones: bloqueo manual, alta, ajuste (0 unidades = baja) y devolución.
- Nomenclatura: CLP `0496xxxx`, matrícula `1000xxxxxx`, componentes `709xxxxx` y packaging `049xxxxx`.
- Seguimiento por cliente: tipo alquiler / compra, sin precios.

**[23/06] Pantalla de camiones** (`PANTALLA_CAMIONES.html`)
- Rejilla de camiones y «Nuevo camión» en una pestaña nueva.
- Fecha de disponibilidad con semáforo.
- Pendiente: decidir si el semáforo va también en la rejilla principal.

**[24–26/06] Calendario de fabricaciones** (`FABRICACIONES.html`)
- Turnos por reactor, planificación, órdenes con reactor y fechas.
- Trazabilidad entre envasado y fabricación.

**[01–28/07] Gestión de envasados** (`GESTION_DE_ENVASADOS.html`)
- Control de producción y stock bloqueado por motivo.
- Necesidades de reproceso con fecha de rotura y orden de reproceso (jerarquía lote / matrícula).

**[07/07–06/11] PRAE: asignación de turnos y polivalencia**
- Motor de turnos 5T / 3T, regímenes especiales, P1/P2/P3 y jefes de línea; edición masiva de la matriz de polivalencia.
- Reglas: sin pasos noche→mañana ni tarde→mañana, máximo 6 días seguidos, el festivo pausa el ciclo 5T y el jefe solo cubre puesto de operario si no queda ninguno.
- Siempre hay que cargar el histórico real de la semana anterior.
- Pendiente:
  - reglas para personas nuevas sin P1;
  - comodines que no queden atados a un turno;
  - puesto de jefe vacío (dos operarias o hueco);
  - viernes solo de mañana;
  - ajustar el calendario a la demanda real.

**[27/08] Verificación de disponibilidad (CO09 / COMAC)**
- Las confirmaciones ATP quedan congeladas por orden de verificación, no por fecha de necesidad.
- Se redistribuyen con COMAC.
- Verificar solo en la liberación o programar COMAC periódicamente.

**[16/09–17/11] Material multipaís y traslados**
- `flujos-material-multipais.html`, actual frente a objetivo:
  - entre países, del centro 12 a 12FR;
  - entre almacenes, de 1202 a 12PO.
- `PANTALLA_NECESIDADES_TRASLADO.html`: solpeds de traslado y generación de camión de 33 palets.
- Manuales PPT:
  - RMMDDIBE y MM02 para las áreas 12_12PO y 12_12FV, con clave de aprovisionamiento especial Z9;
  - ampliación de materiales a 12FR, 12NL, 12BE y 12DE.
- Tablas de áreas de planificación: MDMA, MDLV, MDLG y MDLW.
- **[29/09] Ámbito de verificación de un centro solo de traslados** (probablemente de este proyecto): OVZ2 / OVZ9, grupo propio en `MARC-MTVFP` y stock en traslado.

**[05–06/10] Notificación de fabricaciones en reactores** (`notificacion-fabricacion-reactores.html`)
- Notificación única con reparto en unidades de almacenamiento.
- Etiqueta de 148×105 mm que se abre en una pestaña nueva.

**[08/10] Facturación y maestro de materiales de Perseida en Qlik**
- Previsión asociada a un cliente por defecto.
- Agrupación de clientes (`Id_` / `Nombre_Agrupacion_Cliente`), que se cruza por `Key`.
- `DROP FIELDS CLC_AGRCL` de `MARA` para evitar asociaciones erróneas.

**[08–27/10] Reserva de cargas** (`RESERVA_CARGAS.html` + Google Apps Script)
- FEDEFARMA, BIDAFARMA y UNNEFA reservan día y hora de carga con nº de palets; la reserva desaparece para los demás.
- Perfiles CLIENTE, PLANIFICADOR y VISUALIZACION.
- Contraseñas de 4 cifras fuera del repositorio; bloqueo de 1 hora tras 5 fallos.
- Aviso por correo a planificación.

**[23/10 y 27/10] Integración GPP–MES**
- Detectar órdenes finalizadas por error: pedir a MES cantidades y duración; reglas de cierre sospechoso; estado «pendiente de validar»; consulta inversa de toda la actividad.
- Pausar órdenes sin tocar MES:
  - tramos de ejecución (`ORDEN_TRAMO`) y foto de cada lectura de MES;
  - máquina de estados con Pausada, Pendiente de cierre y Revisión;
  - anular tramo y reasignar unidades;
  - orden de continuación (`-R1`);
  - bandeja de incidencias.
- Pendiente de confirmar con MES: si deja reabrir órdenes, cómo se ve la corrección y qué dispara GPP.

**[06/11] Crear varias órdenes de envasado** (`CREAR_ORDENES_ENVASADO.html`)
- El mismo pop-up con «Orden i de N», ◀ ▶ y «Usar estos datos en todas las líneas».
- Valida todas las órdenes antes de crear ninguna; llama a `crearOrdenEnvasado()` una vez por línea y muestra un resumen.

**[12/11] CMI – % de servicio** (probablemente Perseida)
- Campo `Carga_Propia` (S / N) por mapeo de clientes que organizan su propia carga.
- Cuentan para el material expedido, pero no para el % de servicio.

---

## 3. SIKNO

**[09–17/09] SCOPE · SIKNO** (rama `sikno-scope-docs` de FERROMALLAS)
- Guía funcional de pantallas e indicadores (`SCOPE_SIKNO_Guia_Funcional_Pantallas_KPIs.docx`).
- SCOPE es un simulador de consecuencias, no un optimizador.
- **Funcionamiento:**
  - ciclo configurar → lanzar el motor → resultados → decidir;
  - motor de referencia v5, con 6 tablas de resultado;
  - nivel de servicio como semáforo y buffers de subcontratas.
- **11 pantallas:** Dashboard, Maestro y Calendario de Secciones, Rutas, Buffer Subcontratas, Prioridades, Carga por Secciones, Plan de Operaciones, Atraso de Pedidos, Alertas y CTP Comercial.
- **Auditoría de entradas:** Materiales, Órdenes, Pedidos, Secciones y Rutas. Se detectaron 3 fallos:
  - `velocidad_valor` y `tiempo_estándar` no cuadran;
  - `prioridad_manual_hasta` no caduca nunca;
  - el tercero no se recuperó completo.

---

## 4. ONDUPET (packaging; GIO)

**[28/05–17/06] Carga de capacidad OPET** (`Carga_Capacidad_OPET.html`)
- Ciclos y familias con velocidad en uds/min, datos maestros en Supabase.
- Pools B/C con su impacto, reservas de temporal y mapa de calor por Termoformado / Extrusión.
- Documento de metodología.

**[05/06] CTP** (rama `compassionate-clarke`)
- Opción B: insertar al final de la familia; solo el exceso sobre el temporal desplaza a las siguientes.
- Opción C: forzar la fecha de entrega hacia atrás, insertando antes de la orden que bloquea.
- Información MTO / MTS en la hoja de órdenes.

**[10–17/06] Balance de masas** (`BALANCE_MASAS.html`)
- Merma intrínseca por el teórico de la lista de materiales; pesadas extrínsecas.
- Baja automática: el sobrante del lote se imputa a la última orden antes del cambio de secuencia (se quitó la regla de 24 h).

**[12/06] Seguimiento de pedidos v5.2**
- Documentación funcional con las fórmulas de los KPIs: atraso real y teórico, cobertura, planificación y motivos.
- Se eliminaron las pestañas Stock y CTP.

**[12–15/06] Reservas de stock** (`RESERVAS_ONDUPET.html`)
- Prioridades: P1 manual, P2 urgente, P3 clientes / referencias, P4/P5 horizonte.
- Algoritmo V8 en JavaScript.
- [20/10] Segunda revisión del algoritmo, interrumpida por el límite de uso.

**[01/10–29/10] Balance de masas para dirección** (`Balance_Masas_Direccion.html`)
- Desviación en kg por orden y material, sin euros.
- Umbrales editables: ±3 % sin desviación, ±6 % crítico.
- Filtros de máquina, fechas, tipo (bobina / auxiliar) y pedido.

**[11/11] Secuenciador de bobinas de extrusión** (`Secuenciador_Bobinas_Ondupet.html`)
- Cambios:
  - 90 min por cambio de calidad, más la matriz de colores;
  - 90 min si el ancho varía más de ±20 mm;
  - 30 min por cambio de micraje.
- Reordena sin cambiar de máquina, con el lote en curso fijo; prioriza evitar stock negativo y luego minimizar cambios.
- Con los datos reales, las bobinas en negativo bajan de 49 a 37.
- Conclusión: a EREMA EXT.2 le falta capacidad.

**[10–17/11] Plan to Release (P2R) en Ondupet** (rama `cl-grupo-packaging-presentation`, app DICE)
- Catálogo de hallazgos (HT) y palancas (P2R): una palanca por hallazgo.
- HT-04 se fusiona en HT-13 (sobrecobertura) y P2R-03 en P2R-01 («Gestión de la demanda y coberturas»).
- Roturas separadas en PT y MMPP; HT-11 con KPI propio.
- KPIs iniciales: OTIF, cambios / producción, roturas y días de stock. Últimos 12 meses; objetivo, el mejor trimestre.
- Las paradas por falta de material alimentan HT-06 y HT-09; las horas de línea paradas, HT-10.
- Pendiente: responsable y fecha de puesta en marcha de las 17 intervenciones.

---

## 5. BALBOA (acería; GESAG + GPB)

**[27/07] Facturación de Balboa** (`scripts/BALBOA_Facturacion.qvs`)
- Expedido y cartera, con QVD propio.

**[19/10] Stock de Getafe en Qlik**
- Se añade a `STOCK` con el campo `Sociedad`, en el histórico `STOCK_Balboa.qvd`.
- Cartera, tipo de atraso, días de stock y máquina siguen solo con Balboa.
- Relleno de ceros de los productos estrella por sociedad.

**[27–29/10] Arreglo de paquetes y no conformidades**
- Lead time de no conformidades: movimiento al taller → arreglo, con mediana de 1,3 días.
- Análisis de NC repetidas en varias coladas: casi siempre es el mismo número reutilizado. Cada NC se identifica por colada + nº.

**[13–17/11] CMI – cambios de familia en el tren de perfiles** (consulta `GPB`)
- La secuencia IPE → IPEA → IPEAA cuenta 1 cambio.
- HEA ↔ HEB cuenta cambio solo con distinta medida.
- `Fin_Laminacion` = «YA FABRICADO» si lo reservado o con nota ≥ pendiente − 1,5 Tn.

---

## 6. OMEGA

**[26/08] Análisis del flujo de stock** (`Omega_AnalisisFlujoStock_v1.html`, FERROMALLAS)
- La sociedad se configuró para reparar (avisos y órdenes de reparación con IW42, MB1A y DP90) pero fabrica.
- Sin notificación de producto terminado, en el ejemplo de 700 uds solo 50 quedan en sistema. Hay materia prima fantasma y el MRP no se puede usar.

**[16/11] Flujo de reparación: actual frente a objetivo** (`Omega_proceso_reparacion.html`)
- Objetivo: listas de materiales y hojas de ruta como requisito previo; separar la solicitud de reparación de la fabricación para stock con orden CO (KO01 / KO88).
- Riesgos:
  - coste 0 frente al precio estándar;
  - DP90 sin importe;
  - la orden CO no genera necesidades MRP.
- Pendiente: confirmar qué es exactamente la «orden CO».

---

## 7. MARCELIANO MARTÍN

- No he encontrado ninguna conversación ni documento identificado como Marceliano Martín.
- La presentación al CIO lista proyectos de «MM», sin más detalle.
- Puede que algunas conversaciones de SAP del apartado 10 (por ejemplo, el piloto C312) sean de esta empresa. Indícamelo y las muevo.

---

## 8. MGT / PDA (papel–cartón)

**[29–30/10] Integración papel–cartón MGT–Ondupack** (rama `gifted-lovelace`)
- Situación actual: coordinación por personas y correo, y 7 responsabilidades sin dueño.
- **Escenario 1:** integración SAP con sistema pull (mínimos y máximos; el pedido de compra en Ondupack crea el de venta en MGT).
- **Escenario 2:** departamento común con Planificador Master y dos secuenciadores.
- 6 decisiones de dirección.
- Entregables: presentación, guion para Teams e informe Word de 7 páginas.

**[03–10/11] Plan to Release MGT y Ondupack Almendralejo**
- Presentación P2R para dirección, con agentes por empresa.
- Consolidado P2R con KPIs dados de alta en DICE.
- Intervenciones de MGT y Ondupack.

**[10/11] Flujo de planificación por planta** (`Flujo_Planificacion_Plantas.pptx`)
- As-is, organigrama y automatizaciones.
- Papel = MGT · PDA · Papresa, con tres secuenciadores; el de PDA está pendiente de nombre.
- El algoritmo de combinación de papel cuelga de «Bobinadora».

---

## 9. DEPARTAMENTO (planificación corporativa · CL Grupo Industrial)

**[17–18/08] Presentación del Departamento de Planificación** (`Departamento_Planificacion.pptx`)
- Tres roles: Functional Lead, Technical Lead y Solution Developer.
- Reuniones: weekly funcional, weekly técnica y daily, más la mensual y el 1:1 trimestral.
- Plan de desarrollo individual y DAFO.
- Pendiente: DAFO por persona y confirmar cadencias.

**[02/09] Reorganización del equipo tras una salida**
- Propuesta de organización y tabla de decisiones.
- Pendiente: decidir si se prepara como página visual para presentarla.

**[28–29/09] Fin de colaboración de un colaborador externo y carta de recomendación**
- Correo de formalización con preaviso de 60 días, en registro mercantil, no laboral.
- Carta de recomendación en clave de prestación de servicios.

**[22–26/10] Carta de recomendación de CL Grupo Industrial**
- Maquetada en PDF y Word, con membrete corporativo.
- Recomendación: la firma el responsable directo y se añade el sello de la empresa.

**[08/10–09/11] Modelo de planificación corporativo CL** (`modelo_planificacion/`)
- Corporativo define; cada planta ejecuta.
- Herramienta HTML con notas, flujos as-is / to-be, RACI, organigramas, políticas, KPIs y plan de acción.
- Kit en Excel.
- Presentación «Nuevo modelo de planificación» con Celeste.

**[10–12/11] Presentación del departamento para el CIO**
- ERP y GP por empresa (tabla del principio).
- 22 proyectos de 2026 agrupados por empresa.
- Pendiente:
  - qué hacer con la diapositiva 10 (Carga de Capacidad en AGSA): rotularla como en desarrollo, presentarla como vista previa o quitarla;
  - el estado de 3 proyectos: algoritmo de reservas de Ondupet, y pantallas de notificación y de tiempos de cambio de Perseida.

**[18/11] Documento funcional de parámetros MRP** (job MRP en planta)
- Word de 11 páginas sobre la plantilla de Fijación de SolPed.
- Contenido: NETCH / NETPL / NEUPL, lista MRP, modo de planificación y variante del job en SM37.
- Pendiente de revisar: lista MRP 1 → 2 y modo 3 → 1.

---

## 10. Pendiente de asignar a una sección

**[08/06] Speed Master KPI y necesidades de MMPP** (FERROMALLAS, `Ferromallas_CalendarioProduccion_v20.html`)
- Necesidades de materia prima por diámetro de alambrón para mallas.
- Maestro de velocidades y tiempos de cambio.

**[08/07 y 01/09] Qlik: horas de operario, de marcha y de GMAO** (MES, `id_sociedad = 2`)
- Fechas sin pasar por texto y filtros SARGable.
- Mismo nombre de máquina en las dos tablas.
- Columnas cualificadas en GMAO.

**[05/10–06/11] SAP: MRP con punto de pedido y fecha de rotura por orden**
- Mensaje 10 (adelantar) con punto de pedido.
- Fecha de rotura por orden: suma de las cantidades de las órdenes anteriores, sin mirar sus fechas.
- Excel con 10 casos de prueba.

**[17–18/11] Fijación de SolPed por pedido de SalesForce y cobertura ABC** (piloto C312; materiales ZERT)
- Casos:
  - 1, pedido directo: kilos redondeados al lote mínimo;
  - 2, grupo Z2: kilos tal cual;
  - 3, grupo 40: con lote mínimo, solo si la cobertura de SAP < máxima del ABC (A 75, B 180, C 365 días) y el lanzamiento ≤ hoy + cobertura.
- SolPed creada ya fijada y con referencia al pedido; después, MRP individual en modo «borrar y crear de nuevo».
- Word de 7 páginas.
- Pendiente: si un caso 3 fuera de rango se vuelve a revisar más adelante.

**Excluidas por ser personales:** comparación de seguros y ayudas en Extremadura, apuntes de inglés B2 y oposiciones.
