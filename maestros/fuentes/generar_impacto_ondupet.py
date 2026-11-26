"""Impacto en € de los hallazgos P2R de Ondupet contra la perfección (12 meses: nov-2025 a oct-2026).

El € de cada hallazgo se calcula contra la perfección (OTIF 100 %, 0 agotados, 0 paradas por falta de material,
0 cambios, saturación 100 % de las horas con turno, stock mínimo). El objetivo realista se pone después en la palanca.

Estructura:
- Supuestos: los parámetros de coste y precio. No hay ninguno en los datos: son supuestos editables (celdas
  amarillas) que hay que sustituir por los reales.
- Bolsillos: el volumen físico de 12 meses de cada KPI (calculado con los Excel Ahorros_E2E_Ondupet) y su € con
  fórmula sobre los supuestos.
- Reparto: el % de cada bolsillo que va a cada hallazgo (propuesta editable) y el € resultante.
- Resumen: € por hallazgo.
- Datos que faltan, Avisos de datos y Propuesta HT-06 (carretillero y bobina).

Uso: python generar_impacto_ondupet.py [salida.xlsx]
"""
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from generar_intervenciones_ondupet import F, NAVY, cabecera, celda, fill, titulo

AMARILLO, GRIS_F, VERDE, ROJO = 'FFF4C2', 'EEF1F4', 'E3F3EA', 'FCE4E4'
PERIODO = 'nov-2025 a oct-2026'
SUP, PEND, DEF = 'Supuesto', 'Pendiente', 'Definido'

# (código, parámetro, valor, unidad, estado, quién lo tiene, se usa en)
SUPUESTOS = [
    ('S01', 'Coste hora de termo parado con el turno abierto (personal directo, energía en espera y fijos)', 100, '€/h', SUP,
     'Control de gestión (tarifa del centro de coste) y RRHH (personas por termo y turno)', 'Faltas, cambios, espera de personal'),
    ('S02', 'Coste fijo hora de termo sin turno (subactividad: amortización y estructura, sin personal)', 50, '€/h', SUP,
     'Control de gestión (SAP CO, centros de coste)', 'Saturación (HT-03)'),
    ('S03', 'Precio medio de venta de PT', 3.00, '€/kg', SUP, 'Control de gestión / SAP SD (facturación 12 m)',
     'OTIF: coste financiero'),
    ('S04', 'Margen de contribución de PT (precio − material − embalaje − energía variable − transporte)', 1.00, '€/kg',
     SUP, 'Control de gestión (SAP CO-PA), por sector si es posible', 'OTIF y agotados'),
    ('S05', 'Coste estándar de PT (valor del stock)', 2.20, '€/kg', SUP, 'SAP CO (coste estándar por material)',
     'Sobrecobertura (HT-13)'),
    ('S06', 'Coste de capital (CMPC)', 0.08, '%/año', SUP, 'Dirección financiera', 'OTIF financiero y HT-13'),
    ('S07', 'Coste de almacén por palet', 36, '€/palet·año', SUP, 'Logística / Control de gestión (propio y externo)',
     'Sobrecobertura (HT-13)'),
    ('S08', '% de la venta servida tarde que se pierde o queda en riesgo', 0.05, '%', SUP,
     'Dirección Comercial (clientes perdidos o que bajan volumen por servicio)', 'OTIF'),
    ('S09', '% de lo no servido (posiciones < 90 % servidas) que es venta perdida', 0.50, '%', SUP,
     'Comercial / SAP SD (motivo de rechazo de la posición)', 'OTIF'),
    ('S10', 'Penalizaciones y abonos por retraso o servicio incompleto (12 m)', 0, '€/año', PEND,
     'Finanzas / Comercial (notas de abono con motivo)', 'OTIF'),
    ('S11', 'Portes urgentes o expediciones fuera de ruta por retraso (12 m)', 0, '€/año', PEND,
     'Logística (facturas del transportista)', 'OTIF'),
    ('S12', 'Compras urgentes y portes urgentes por faltas de material (12 m)', 0, '€/año', PEND,
     'Compras (SAP MM: pedidos urgentes y sobreprecio)', 'Faltas (HT-06)'),
    ('S13', 'Factor de infrarregistro de agotados (1 = lo registrado es todo)', 1.00, 'factor', SUP,
     'Comercial (sep-2026 no tiene ningún registro)', 'Agotados'),
    ('S14', 'Tiempo objetivo de cambio de bloque (molde)', 4, 'h/cambio', DEF, 'Planificación (definido)', 'Cambios'),
    ('S15', 'Tiempo objetivo de cambio de altura (insertos)', 1, 'h/cambio', DEF, 'Planificación (definido)', 'Cambios'),
    ('S16', 'Tiempo objetivo de cambio de calidad', 1, 'h/cambio', DEF, 'Planificación (definido)', 'Cambios'),
    ('S17', 'Merma de arranque por cambio de bloque', 100, 'kg/cambio', SUP, 'Producción / Calidad', 'Cambios'),
    ('S18', 'Merma de arranque por cambio de altura', 30, 'kg/cambio', SUP, 'Producción / Calidad', 'Cambios'),
    ('S19', 'Merma por cambio de calidad (purga y transición)', 50, 'kg/cambio', SUP, 'Producción / Calidad', 'Cambios'),
    ('S20', 'Valor neto de la merma (bobina − molido recuperado)', 0.70, '€/kg', SUP, 'SAP CO / Compras', 'Cambios'),
    ('S21', 'Provisión por obsolescencia del stock sin demanda > 8 meses', 0.50, '% del coste', SUP,
     'Control de gestión (política de provisión)', 'HT-13 (puntual)'),
]
FILA_SUP = 5
REF = {s[0]: f'Supuestos!$C${FILA_SUP + i}' for i, s in enumerate(SUPUESTOS)}

HT = [('HT-01', 'Roturas de stock'), ('HT-03', 'Optimización carga capacidad'), ('HT-05', 'Fechas de entrega poco realistas'),
      ('HT-06', 'Optimización gestión faltas'), ('HT-08', 'Optimización secuenciación'),
      ('HT-09', 'Optimización urgencias secuenciación'), ('HT-11', 'Optimización servicio'), ('HT-13', 'Sobrecobertura de stock')]
DESTINOS = [h[0] for h in HT] + ['Fuera de P2R']
ANUAL, PUNTUAL = 'Anual', 'Puntual'


def f(expr):
    """Sustituye {S01}… por la celda del supuesto."""
    return '=' + expr.format(**REF)


# (id, bolsillo/KPI, componente, volumen, unidad, volumen 2, unidad 2, fórmula € con {V} y {V2}, tipo, fuente y método,
#  reparto {destino: %}, por qué ese reparto)
BOLSILLOS = [
    # ------------------------------------------------------------------ OTIF 79,2 % por líneas (KPI de HT-05 y HT-11)
    ('B01', 'OTIF (79,2 %)', 'Margen en riesgo por los kg servidos tarde', 2_028_000, 'kg tarde', None, None,
     '{V}*{S04}*{S08}', ANUAL,
     'Servicio: 1.237 de 5.960 líneas tarde (albarán > fecha comprometida); kg = Ctd Expedida × peso del Maestro, sin bobinas.',
     {'HT-01': 28.5, 'HT-05': 11.1, 'HT-11': 60.4},
     'Atribución causal por kg: HT-01 = MTS sin stock el día antes (ATP), HT-05 = plazo comprometido menor que el '
     'realizable (5 d MTS, 21 d MTO), HT-11 = el resto (había stock y plazo). Solo se imputa el exceso sobre la tasa base de retraso.'),
    ('B02', 'OTIF (79,2 %)', 'Venta perdida por servicio incompleto (posiciones < 90 % servidas)', 114_000, 'kg no servidos',
     None, None, '{V}*{S04}*{S09}', ANUAL,
     'Servicio: 167 posiciones por debajo del 90 % (11,95 M uds). In full 95,7 % al 100 %.',
     {'HT-11': 100}, 'Había pedido y no se completó: fallo de plan o de seguimiento.'),
    ('B03', 'OTIF (79,2 %)', 'Coste financiero del cobro retrasado', 16_686_000, 'kg·día de retraso', None, None,
     '{V}*{S03}*{S06}/365', ANUAL, 'Servicio: Σ kg tarde × días de retraso (media 8,2 días, mediana 4).',
     {'HT-01': 32.2, 'HT-05': 11.5, 'HT-11': 56.3}, 'Misma atribución causal, con t·día como inductor.'),
    ('B04', 'OTIF (79,2 %)', 'Penalizaciones, abonos y portes urgentes', 1_237, 'líneas tarde', None, None,
     '{S10}+{S11}', ANUAL, 'Pendiente de Finanzas y Logística: 0 € hasta tener el dato.',
     {'HT-01': 22.2, 'HT-05': 13.6, 'HT-11': 64.2}, 'Misma atribución causal, con líneas como inductor.'),
    # ------------------------------------------------------------------ Agotados (KPI roturas PT)
    ('B05', 'Agotados (177 t registradas)', 'Rotura MTS: venta perdida sin stock', 56_022, 'kg perdidos', None, None,
     '{V}*{S04}*{S13}', ANUAL,
     'Agotados 12 m: 34 registros MTS perdidos (declarados o sin pedido en ±10 días) y el 50 % de 15 parciales MTS.',
     {'HT-01': 100}, 'Rotura de stock MTS.'),
    ('B06', 'Agotados (177 t registradas)', 'Plazo MTO no aceptado por el cliente', 42_476, 'kg perdidos', None, None,
     '{V}*{S04}*{S13}', ANUAL, 'Agotados de referencias MTO o bobina: 11 perdidos y el 50 % de 3 parciales.',
     {'HT-05': 100}, 'No es rotura: el cliente no acepta la fecha que se le ofrece.'),
    ('B07', 'Agotados (177 t registradas)', 'Venta perdida por falta de ABS', 17_398, 'kg perdidos', None, None,
     '{V}*{S04}*{S13}', ANUAL, 'Agotados: 4 registros «no se puede fabricar al no tener stock de ABS».',
     {'HT-06': 100}, 'Falta de material.'),
    ('B08', 'Agotados (177 t registradas)', 'Venta perdida por falta de capacidad', 3_900, 'kg perdidos', None, None,
     '{V}*{S04}*{S13}', ANUAL, 'Agotados: 2 registros «no teníamos capacidad en termo 5».', {'HT-03': 100}, 'Capacidad.'),
    ('B09', 'Agotados (177 t registradas)', 'Parciales MTS (la otra mitad)', 1_896, 'kg perdidos', None, None,
     '{V}*{S04}*{S13}', ANUAL, 'Agotados: el 50 % de los 15 parciales MTS que no va a HT-01.',
     {'HT-05': 50, 'HT-11': 50}, 'Pedido de otra cantidad: plazo o servicio.'),
    # ------------------------------------------------------------------ Paradas por falta de material (KPI de HT-06)
    ('B10', 'Paradas por falta de material (1,6 % del calendario)', 'Espera por falta de bobinas fabricadas', 687.7,
     'h de termo', None, None, '{V}*{S01}', ANUAL,
     'Paradas: 239 registros (133 episodios). El 72 % ocurre con la extrusora en marcha fabricando otra cosa y a mitad de pedido.',
     {'HT-06': 72, 'HT-08': 15, 'Fuera de P2R': 13},
     'HT-06 72 %: la extrusora estaba disponible; falla la cobertura de bobina frente al plan del termo. HT-08 15 %: '
     'espera dentro de un cambio o con la extrusora cambiando de pedido (secuencias no sincronizadas). Fuera 13 %: '
     'avería o fallo de la extrusora. A decidir.'),
    ('B11', 'Paradas por falta de material (1,6 % del calendario)', 'Faltas de cajas, ABS y «fallo de planificación» de material',
     58.3, 'h de termo', None, None, '{V}*{S01}', ANUAL,
     'Paradas: faltan cajas 30,9 h, falta ABS 2,3 h y 25,2 h de «PP-fallo planificación» cuyo texto es una falta.',
     {'HT-06': 100}, 'Falta de material.'),
    ('B12', 'Paradas por falta de material (1,6 % del calendario)', 'Falta de abastecimiento del carretillero y afines',
     76.3, 'h de termo', None, None, '{V}*{S01}', ANUAL,
     'Paradas: OP-falta abastecimiento carretillero 66,9 h, parada ayuda carretillero 5,6 h y saturación de expediciones 3,9 h.',
     {'HT-06': 100}, 'Decidido: va a HT-06 (aprovisionamiento a línea; HT-10 no está asignado en Ondupet).'),
    ('B13', 'Paradas por falta de material (1,6 % del calendario)', 'Compras y portes urgentes por faltas', None, '€/año',
     None, None, '{S12}', ANUAL, 'Pendiente de Compras: 0 € hasta tener el dato.', {'HT-06': 100}, 'Falta de material.'),
    # ------------------------------------------------------------------ Cambios (KPI nº cambios / producción)
    ('B14', 'Cambios (1,52 por M uds)', 'Cambios de bloque (molde)', 363, 'cambios', None, None,
     '{V}*({S14}*{S01}+{S17}*{S20})', ANUAL,
     'Producciones reconstruida en orden cronológico por termo (los flags del Excel cuentan 1.013). Tiempo objetivo × coste hora + merma.',
     {'HT-08': 65, 'HT-09': 35},
     'Comparando la secuencia ejecutada con la foto del plan de cada viernes: 48 % de las horas según plan, 17,5 % '
     'urgencias directas, 10 % retornos tras urgencia y 24 % resecuenciado sin urgencia visible (mitad y mitad). Rango de HT-09: 20–48 %.'),
    ('B15', 'Cambios (1,52 por M uds)', 'Cambios de altura (insertos)', 558, 'cambios', None, None,
     '{V}*({S15}*{S01}+{S18}*{S20})', ANUAL, 'Producciones reconstruida (los flags del Excel cuentan 1.006).',
     {'HT-08': 65, 'HT-09': 35}, 'Igual que los cambios de bloque.'),
    ('B16', 'Cambios (1,52 por M uds)', 'Cambios de calidad sin cambio de bloque', 177, 'cambios', None, None,
     '{V}*({S16}*{S01}+{S19}*{S20})', ANUAL,
     'Producciones reconstruida (los flags del Excel cuentan 336). Misma regla que el Excel: calidad distinta con el mismo bloque.',
     {'HT-08': 65, 'HT-09': 35}, 'Igual que los cambios de bloque.'),
    # ------------------------------------------------------------------ Saturación (KPI % saturación)
    ('B17', 'Saturación (marcha 71,9 % del calendario)', 'Termos sin turno o sin personal, y arranques de semana', 593.0,
     'h de termo', None, None, '{V}*{S02}', ANUAL,
     'Paradas: PP-sin turno no hay personal 524,6 h, plantilla no presente, arranque y paro de fin de semana. Casi todo en T1 y T2.',
     {'HT-03': 100}, 'Horas sin carga. Confirmar con Producción si «sin turno» es falta de carga (HT-03) o absentismo (fuera de P2R).'),
    ('B18', 'Saturación (marcha 71,9 % del calendario)', 'Termos sin pedidos', 54.1, 'h de termo', None, None,
     '{V}*{S02}', ANUAL, 'Paradas: PP-planing sin pedidos (T1, T5 y T6).', {'HT-03': 100}, 'Horas sin carga.'),
    ('B19', 'Saturación (marcha 71,9 % del calendario)', 'Espera de cambio de molde por falta de personal', 104.0,
     'h de termo', None, None, '{V}*{S01}', ANUAL,
     'Paradas: los maquinistas están con averías o cambios en otros termos (2 maquinistas por turno).',
     {'HT-03': 75, 'HT-08': 25},
     'El cuello es el nº de maquinistas por turno (HT-03); el 25 % se evitaría escalonando los cambios en la secuencia (HT-08).'),
    ('B20', 'Saturación (marcha 71,9 % del calendario)', 'Extrusoras sin turno o sin personal', 86.0, 'h de extrusora', None,
     None, '{V}*{S02}', ANUAL, 'Paradas de E1 y E2.', {'HT-03': 100}, 'Horas sin carga.'),
    # ------------------------------------------------------------------ Stock de PT (KPI días de stock PT)
    ('B21', 'Stock de PT (31,7 días; 856 t de media)', 'Exceso sobre la cobertura técnica (ciclo/2 + 7 días; MTO 7 días)',
     276_147, 'kg', 2_530, 'palets', '{V}*{S05}*{S06}+{V2}*{S07}', ANUAL,
     'Dias_Stock, media de las fotos fiables 27-mar a 24-jul-2026, por material frente a la demanda de 12 m. Incluye 39 t '
     'sin ninguna venta en 12 m. 65 materiales explican el 80 %.',
     {'HT-13': 100}, 'Sobrecobertura.'),
    ('B22', 'Stock de PT (31,7 días; 856 t de media)', 'Stock de ciclo por fabricar en campañas (hasta la perfección de 7 días)',
     402_088, 'kg', 3_051, 'palets', '{V}*{S05}*{S06}+{V2}*{S07}', ANUAL,
     'Diferencia entre la perfección (7 días de cobertura) y la cobertura técnica: el ciclo medio ponderado es de 55 días.',
     {'HT-13': 100},
     'Parte del KPI de días de stock contra la perfección. Alternativa a decidir: tratarlo como contrapartida de HT-08 '
     '(lotes grandes para evitar cambios) o dejarlo fuera, porque «0 cambios» y «0 stock de ciclo» no pueden darse a la vez.'),
    ('B23', 'Stock de PT (31,7 días; 856 t de media)', 'Obsolescencia: stock sin demanda desde hace más de 8 meses',
     46_600, 'kg', None, None, '{V}*{S05}*{S21}', PUNTUAL,
     'Dias_Stock: 56 materiales con 365 días de cobertura en todas las fotos fiables (MTO 28,1 t, MTS 18,5 t).',
     {'HT-13': 100}, 'Pérdida puntual (provisión), no anual.'),
]
FILA_B = 5


def hoja_supuestos(wb):
    ws = wb.active
    ws.title = 'Supuestos'
    titulo(ws, 'Supuestos de coste y precio · Impacto P2R Ondupet',
           'Los Excel de Ondupet no traen costes ni precios. Las celdas amarillas son SUPUESTOS para poder calcular: '
           'cámbialos por el dato real y todo el libro se recalcula. «Definido» = valor acordado por Planificación.')
    cabecera(ws, 4, ['Código', 'Parámetro', 'Valor', 'Unidad', 'Estado', 'Quién lo tiene', 'Se usa en'],
             [8, 70, 11, 13, 12, 52, 30])
    for i, (cod, par, val, uni, est, quien, uso) in enumerate(SUPUESTOS):
        r = FILA_SUP + i
        fmt = '0%' if uni.startswith('%') else '#,##0.00' if isinstance(val, float) and val < 10 else '#,##0'
        for j, v in enumerate([cod, par, val, uni, est, quien, uso], 1):
            c = celda(ws, r, j, v, centrada=j in (1, 3, 4, 5), bold=j == 3, fmt=fmt if j == 3 else None)
            if j == 3:
                c.fill = fill(AMARILLO if est != DEF else GRIS_F)
    ult = FILA_SUP + len(SUPUESTOS) - 1
    for valor, color in ((SUP, AMARILLO), (PEND, ROJO), (DEF, VERDE)):
        ws.conditional_formatting.add(f'E{FILA_SUP}:E{ult}', FormulaRule(formula=[f'$E{FILA_SUP}="{valor}"'], fill=fill(color)))
    ws.freeze_panes = 'C5'


def hoja_bolsillos(wb):
    ws = wb.create_sheet('Bolsillos')
    titulo(ws, f'Bolsillos de € por KPI · volumen de 12 meses ({PERIODO}) contra la perfección',
           'Volumen físico calculado con Ahorros_E2E_Ondupet (Servicio, Agotados, Dias_Stock, Paradas, Producciones e '
           'Historico_Planif_Semanal). El € sale de multiplicar el volumen por los supuestos. «Puntual» no se repite cada año.')
    cab = ['ID', 'Bolsillo (KPI)', 'Componente', 'Volumen 12 m', 'Unidad', 'Volumen 2', 'Unidad 2', '€', 'Tipo',
           'Fuente y método']
    cabecera(ws, 4, cab, [6, 30, 46, 14, 16, 11, 9, 13, 9, 90])
    for i, (bid, bol, comp, vol, uni, vol2, uni2, expr, tipo, fuente, _, _) in enumerate(BOLSILLOS):
        r = FILA_B + i
        formula = '=' + expr.format(V=f'D{r}', V2=f'F{r}', **REF)
        vals = [bid, bol, comp, vol, uni, vol2, uni2, formula, tipo, fuente]
        for j, v in enumerate(vals, 1):
            fmt = '#,##0.0' if j in (4, 6) and isinstance(v, float) else '#,##0' if j in (4, 6) else '#,##0 "€"' if j == 8 else None
            celda(ws, r, j, v, centrada=j in (1, 4, 5, 6, 7, 8, 9), fmt=fmt, bold=j == 8)
    ult = FILA_B + len(BOLSILLOS) - 1
    t = ult + 1
    celda(ws, t, 3, 'Total anual', bold=True)
    celda(ws, t, 8, f'=SUMIF(I{FILA_B}:I{ult},"{ANUAL}",H{FILA_B}:H{ult})', centrada=True, fmt='#,##0 "€"', bold=True)
    celda(ws, t + 1, 3, 'Total puntual', bold=True)
    celda(ws, t + 1, 8, f'=SUMIF(I{FILA_B}:I{ult},"{PUNTUAL}",H{FILA_B}:H{ult})', centrada=True, fmt='#,##0 "€"', bold=True)
    # subtotales por bolsillo
    r0 = t + 3
    cabecera(ws, r0, ['', 'Bolsillo (KPI)', '', '', '', '', '', '€ anual'], [6, 30, 46, 14, 16, 11, 9, 13])
    bolsillos = list(dict.fromkeys(b[1] for b in BOLSILLOS))
    for k, b in enumerate(bolsillos):
        r = r0 + 1 + k
        celda(ws, r, 2, b, bold=True)
        celda(ws, r, 8, f'=SUMIFS(H{FILA_B}:H{ult},B{FILA_B}:B{ult},B{r},I{FILA_B}:I{ult},"{ANUAL}")', centrada=True,
              fmt='#,##0 "€"', bold=True)
    ws.freeze_panes = 'D5'
    return ult


def hoja_reparto(wb, ult_b):
    ws = wb.create_sheet('Reparto')
    titulo(ws, 'Reparto de cada bolsillo entre hallazgos (propuesta)',
           'Celdas amarillas: % propuesto con los datos, editable. Cada fila tiene que sumar 100 %. «Fuera de P2R» = '
           'causas de mantenimiento u operación. A la derecha, el € que resulta para cada hallazgo.')
    nd = len(DESTINOS)
    cab = ['ID', 'Componente', '€ del bolsillo', 'Tipo'] + [f'% {d}' for d in DESTINOS] + ['Σ %'] + \
          [f'€ {d}' for d in DESTINOS] + ['Por qué este reparto']
    anchos = [6, 44, 13, 9] + [9] * nd + [8] + [11] * nd + [90]
    cabecera(ws, 4, cab, anchos)
    c_pct0, c_sum = 5, 5 + nd
    c_eur0 = c_sum + 1
    for i, (bid, _, comp, *_resto) in enumerate(BOLSILLOS):
        reparto, porque = BOLSILLOS[i][10], BOLSILLOS[i][11]
        r = 5 + i
        rb = FILA_B + i
        celda(ws, r, 1, bid, centrada=True)
        celda(ws, r, 2, comp)
        celda(ws, r, 3, f'=Bolsillos!H{rb}', centrada=True, fmt='#,##0 "€"', bold=True)
        celda(ws, r, 4, f'=Bolsillos!I{rb}', centrada=True)
        for k, d in enumerate(DESTINOS):
            v = reparto.get(d)
            c = celda(ws, r, c_pct0 + k, v / 100 if v else None, centrada=True, fmt='0.0%')
            c.fill = fill(AMARILLO)
            col_pct = get_column_letter(c_pct0 + k)
            celda(ws, r, c_eur0 + k, f'=IF({col_pct}{r}="","",$C{r}*{col_pct}{r})', centrada=True, fmt='#,##0')
        p0, p1 = get_column_letter(c_pct0), get_column_letter(c_sum - 1)
        celda(ws, r, c_sum, f'=SUM({p0}{r}:{p1}{r})', centrada=True, fmt='0%', bold=True)
        celda(ws, r, c_eur0 + nd, porque)
    ult = 4 + len(BOLSILLOS)
    sc = get_column_letter(c_sum)
    ws.conditional_formatting.add(f'{sc}5:{sc}{ult}', FormulaRule(formula=[f'ABS(${sc}5-1)>0.0001'], fill=fill(ROJO),
                                                                 font=Font(name=F, bold=True, color='B42318')))
    # totales por hallazgo
    t = ult + 1
    celda(ws, t, 2, 'Total anual por hallazgo', bold=True)
    celda(ws, t + 1, 2, 'Total puntual por hallazgo', bold=True)
    for k in range(nd):
        col = get_column_letter(c_eur0 + k)
        celda(ws, t, c_eur0 + k, f'=SUMIF($D$5:$D${ult},"{ANUAL}",{col}5:{col}{ult})', centrada=True, fmt='#,##0', bold=True)
        celda(ws, t + 1, c_eur0 + k, f'=SUMIF($D$5:$D${ult},"{PUNTUAL}",{col}5:{col}{ult})', centrada=True, fmt='#,##0',
              bold=True)
    ws.freeze_panes = ws.cell(5, 3)
    return {d: (get_column_letter(c_eur0 + k), t) for k, d in enumerate(DESTINOS)}


KPI_HT = {
    'HT-01': ('Roturas PT (agotados) · parte del OTIF', '177 t registradas de agotados en 12 m (1,8 % de lo expedido) · '
              '28,5 % del kg tarde era MTS sin stock'),
    'HT-03': ('% saturación', 'Marcha 71,9 % del calendario 24×7; horas sin carga 647 h en termos y 86 h en extrusoras'),
    'HT-05': ('OTIF', '79,2 % por líneas · 11 % del kg tarde se debe a plazos irrealizables · 42 t de agotados MTO'),
    'HT-06': ('Paradas por falta de material', '822 h de termo parado (bobina 688, carretillero 76, cajas y ABS 58)'),
    'HT-08': ('Nº cambios / producción', '1.098 cambios (363 bloque, 558 altura, 177 calidad) = 2.187 h objetivo; 1,52 por M uds'),
    'HT-09': ('Cambios en el programa cerrado', '~35 % de los cambios no seguía la secuencia del plan del viernes'),
    'HT-11': ('OTIF', '79,2 % por líneas · 60 % del kg tarde tenía stock y plazo suficiente · 114 t sin servir'),
    'HT-13': ('Días de stock PT', '31,7 días de media (856 t) · 678 t por encima de la perfección (7 días) · 47 t sin demanda'),
}


def hoja_resumen(wb, cols):
    ws = wb.create_sheet('Resumen', 0)
    titulo(ws, f'Impacto en € de los hallazgos P2R de Ondupet contra la perfección · {PERIODO}',
           'Con los supuestos de la hoja Supuestos (celdas amarillas) y el reparto propuesto de la hoja Reparto. '
           'HT-12, HT-14 y HT-15 (tiempo de persona) se calculan aparte.')
    cabecera(ws, 4, ['Hallazgo', 'Nombre', 'KPI', 'Valor 12 m', '€ anual', '% del total', '€ puntual'],
             [10, 36, 34, 80, 13, 10, 12])
    n = len(HT)
    for i, (hid, nom) in enumerate(HT):
        r = 5 + i
        col, t = cols[hid]
        kpi, val = KPI_HT[hid]
        vals = [hid, nom, kpi, val, f'=Reparto!{col}{t}', f'=E{r}/$E${5 + n + 1}', f'=Reparto!{col}{t + 1}']
        for j, v in enumerate(vals, 1):
            celda(ws, r, j, v, centrada=j in (1, 5, 6, 7), bold=j in (1, 5),
                  fmt='#,##0 "€"' if j in (5, 7) else '0%' if j == 6 else None)
    r = 5 + n
    col, t = cols['Fuera de P2R']
    for j, v in enumerate(['—', 'Fuera de P2R (averías de extrusora)', '', '', f'=Reparto!{col}{t}', f'=E{r}/$E${r + 1}',
                           f'=Reparto!{col}{t + 1}'], 1):
        celda(ws, r, j, v, centrada=j in (1, 5, 6, 7), fmt='#,##0 "€"' if j in (5, 7) else '0%' if j == 6 else None)
    tot = r + 1
    celda(ws, tot, 2, 'Total', bold=True)
    celda(ws, tot, 5, f'=SUM(E5:E{r})', centrada=True, fmt='#,##0 "€"', bold=True)
    celda(ws, tot, 7, f'=SUM(G5:G{r})', centrada=True, fmt='#,##0 "€"', bold=True)
    nota = tot + 2
    for k, texto in enumerate([
        'Lectura: es lo que cuesta al año no estar en la perfección en cada KPI. El objetivo de la palanca (p. ej. pasar de '
        '79 % a 85 % de OTIF) se aplicará después sobre estos importes.',
        'Circulante que se liberaría si el stock estuviera en la perfección (caja, una sola vez): ver «Circulante liberable».',
        'Los € dependen sobre todo de cuatro supuestos: coste hora de termo (S01), margen de contribución (S04), % de venta '
        'tarde que se pierde (S08) y coste estándar del stock (S05). Son los primeros datos que hay que pedir.',
    ]):
        c = ws.cell(nota + k, 1, texto)
        c.font = Font(name=F, size=9, italic=True, color='5B6B78')
        ws.merge_cells(start_row=nota + k, start_column=1, end_row=nota + k, end_column=7)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.row_dimensions[nota + k].height = 26
    r_circ = nota + 4
    celda(ws, r_circ, 2, 'Circulante liberable (stock por encima de la perfección × coste estándar)', bold=True)
    i21 = [b[0] for b in BOLSILLOS].index('B21')
    i22 = [b[0] for b in BOLSILLOS].index('B22')
    celda(ws, r_circ, 5, f'=(Bolsillos!D{FILA_B + i21}+Bolsillos!D{FILA_B + i22})*{REF["S05"]}', centrada=True,
          fmt='#,##0 "€"', bold=True)
    ws.freeze_panes = 'C5'


DATOS = [
    # (prioridad, dato, unidad y detalle, quién, sustituye a, hallazgos)
    ('Imprescindible', 'Coste hora de termo parado: personas directas por termo y turno × coste hora cargado, más energía en espera y fijos',
     '€/h por termo', 'Control de gestión (tarifa del centro de coste) + RRHH', 'S01', 'HT-06, HT-08, HT-09, HT-03'),
    ('Imprescindible', 'Coste fijo de subactividad por hora de termo (amortización, mantenimiento fijo, estructura)', '€/h por termo',
     'Control de gestión (SAP CO, centros de coste)', 'S02', 'HT-03'),
    ('Imprescindible', 'Margen de contribución de PT por sector (PROTEINA, AGRICULTURA, INDUSTRIA) o por referencia', '€/kg',
     'Control de gestión (SAP CO-PA)', 'S04', 'HT-01, HT-05, HT-11, HT-06, HT-03'),
    ('Imprescindible', 'Coste estándar de PT por material', '€/kg o €/1000 uds', 'SAP CO (coste estándar)', 'S05', 'HT-13'),
    ('Imprescindible', '% de la venta servida tarde que se pierde o queda en riesgo (histórico de clientes que bajan o se van)',
     '%', 'Dirección Comercial', 'S08', 'HT-05, HT-11, HT-01'),
    ('Imprescindible', 'Penalizaciones, abonos y descuentos por retraso o servicio incompleto, 12 m', '€ y nº', 'Finanzas / Comercial',
     'S10', 'HT-05, HT-11'),
    ('Imprescindible', 'Portes urgentes o expediciones fuera de ruta, 12 m', '€ y nº', 'Logística', 'S11', 'HT-05, HT-11'),
    ('Imprescindible', 'Compras urgentes y portes urgentes por faltas de material, 12 m', '€ y nº', 'Compras (SAP MM)', 'S12', 'HT-06'),
    ('Mejora', 'Precio medio de venta por sector', '€/kg', 'Control de gestión / SAP SD', 'S03', 'OTIF (coste financiero)'),
    ('Mejora', 'Coste de capital', '%/año', 'Dirección financiera', 'S06', 'HT-13, OTIF'),
    ('Mejora', 'Coste de almacén propio y externo, y palets en almacén externo', '€/palet·año', 'Logística / Control de gestión',
     'S07', 'HT-13'),
    ('Mejora', 'Motivo de cierre de las posiciones no completas (venta perdida o ajuste pactado)', 'sí/no por posición',
     'SAP SD (motivo de rechazo)', 'S09', 'HT-11'),
    ('Mejora', 'Merma de arranque por tipo de cambio, o consumo real de bobina por orden', 'kg/cambio', 'Producción / Calidad; SAP PP',
     'S17–S19', 'HT-08, HT-09'),
    ('Mejora', 'Valor de la bobina y del molido recuperado', '€/kg', 'SAP CO / Compras', 'S20', 'HT-08, HT-09'),
    ('Mejora', 'Política de provisión por obsolescencia y destrucciones de los últimos 12 m', '% y €', 'Control de gestión',
     'S21', 'HT-13'),
    ('Mejora', 'Validar si los agotados registrados son todos (sep-2026 no tiene ninguno)', 'factor', 'Comercial', 'S13', 'HT-01'),
    ('Para cerrar el reparto', 'Significado de «PP-SIN TURNO NO HAY PERSONAL»: falta de carga o falta de plantilla', '—',
     'Producción / RRHH', '—', 'HT-03 (525 h)'),
    ('Para cerrar el reparto', 'Por qué las esperas de bobina pasan de 0 h (mar–may 2026) a 238 h (oct 2026): cambio real o de codificación',
     '—', 'Producción', '—', 'HT-06, HT-08'),
    ('Para cerrar el reparto', 'Stock de PT de nov-2025 a mar-2026 y de ago a oct-2026 (el de Dias_Stock está congelado del 25-jul al 05-nov)',
     'uds por material y mes', 'SAP (MB5B / MC.9) o BI', '—', 'HT-13'),
    ('Para cerrar el reparto', 'Significado de «Fecha entrega» (en cliente o de expedición) y días de tránsito por cliente', 'días',
     'Logística', '—', 'HT-05, HT-11'),
]


def hoja_datos(wb):
    ws = wb.create_sheet('Datos que faltan')
    titulo(ws, 'Datos que faltan para pasar de supuestos a € reales',
           '«Imprescindible»: cambia mucho el resultado. «Mejora»: se puede trabajar con el supuesto. «Para cerrar el reparto»: '
           'no es un coste, pero decide a qué hallazgo va el €.')
    cabecera(ws, 4, ['Prioridad', 'Dato', 'Unidad y detalle', 'Quién lo tiene', 'Sustituye a', 'Hallazgos'],
             [20, 80, 22, 42, 12, 30])
    for i, fila in enumerate(DATOS):
        for j, v in enumerate(fila, 1):
            celda(ws, 5 + i, j, v, centrada=j in (1, 5), bold=j == 1)
    ult = 4 + len(DATOS)
    for valor, color in (('Imprescindible', ROJO), ('Mejora', AMARILLO), ('Para cerrar el reparto', GRIS_F)):
        ws.conditional_formatting.add(f'A5:A{ult}', FormulaRule(formula=[f'$A5="{valor}"'], fill=fill(color)))
    ws.freeze_panes = 'B5'


AVISOS = [
    ('Producciones · Resumen_Cambios', 'Los flags Cambio_Bloque, Cambio_Altura y Cambio_Calidad comparan cada fila con la '
     'siguiente en el orden del fichero (por pmIdFab, no por fecha) y cada fila es un segmento de turno. Cuentan 2,6 veces de '
     'más: 1.013 / 1.006 / 336 cambios en 12 m frente a 363 / 558 / 177 reales. Hay que ordenar por máquina y pmFecIni.'),
    ('Dias_Stock', 'Solo hay fotos desde el 26-mar-2026, y el stock de PT está congelado del 25-jul al 05-nov-2026 (los días '
     'sí se recalculan). Se usa la media de las fotos fiables de abril a julio.'),
    ('Dias_Stock', '365 días significa «sin demanda en el horizonte», no obsoleto: 192 de las 225 t con 365 días tuvieron '
     'ventas en los últimos 12 meses.'),
    ('Servicio', '«Peso Total Kg» se calcula con la cantidad pedida y se repite en cada entrega parcial (suma 45 % de más). '
     'Se usa Ctd Expedida × peso del Maestro. Estrategia y ABC vienen vacías y se completan desde el Maestro.'),
    ('Servicio', 'El retraso se mide contra la fecha comprometida, que a veces ya está renegociada: 49 agotados con «fecha de '
     'entrega ofrecida» aparecen después como servidos en plazo. Si «Fecha entrega» es de llegada al cliente, el OTIF real '
     'bajaría hasta el 63,6 %.'),
    ('Servicio', 'Bobinas PET vendidas a terceros: peso 0 y unidades que no cuadran. Solo se cuentan como líneas (OTIF 36 %).'),
    ('Agotados', 'Estado = 1 en todos los registros; el motivo está en Comentario. Septiembre de 2026 no tiene ningún registro '
     'y un solo usuario hace el 76 %: es un suelo, no la realidad.'),
    ('Agotados', 'En 24 de 29 agotados MTS con foto de stock había stock físico suficiente: el problema es sobre todo de ATP o '
     'reservas a otros pedidos, no de falta física.'),
    ('Paradas', 'Registra el 100 % del calendario 24×7 sin huecos, así que la saturación se puede calcular. pParadaGrupo no '
     'sirve para clasificar (el mismo motivo tiene varios grupos); se clasifica por mpNombre.'),
    ('Paradas', 'Cambios de patrón a validar: «sin turno» suma 315 h en nov-2025 y luego cae a 0, justo cuando suben las '
     'esperas de bobina (0 h en mar–may 2026 y 238 h en oct-2026).'),
    ('Resumen_Paradas', 'La tabla dinámica cubre 23 meses (ene-2025 a nov-2026), no 12, e infla el carretillero (226 h solo en '
     'el primer trimestre de 2025).'),
    ('Historico_Planif_Semanal', 'Fotos de todos los viernes. La columna Calidad está vacía y las fotos anteriores al '
     '12-dic-2025 están incompletas. El plan usa tiempos de cambio de la mitad de los reales.'),
]


def hoja_avisos(wb):
    ws = wb.create_sheet('Avisos de datos')
    titulo(ws, 'Avisos sobre los datos de Ahorros_E2E_Ondupet', 'Lo que hay que corregir o tener en cuenta al usar estas hojas.')
    cabecera(ws, 4, ['Hoja', 'Aviso'], [26, 140])
    for i, (h, a) in enumerate(AVISOS):
        celda(ws, 5 + i, 1, h, bold=True)
        celda(ws, 5 + i, 2, a)


PROPUESTA = [
    ('P2R-05', 'Automatización', 'Llamada de material a línea desde la secuencia',
     'Con la secuencia de cada termo, el sistema genera cada turno la lista de bobinas, cajas y ABS que necesita cada termo '
     'en las próximas horas y la envía al terminal del carretillero, ordenada por hora de necesidad. El carretillero confirma '
     'la entrega y la etiqueta de la bobina se valida al montarla.',
     'Paradas «OP-falta abastecimiento carretillero» y afines: 76 h/año. Por texto: 48 % falta de tiempo o de capacidad del '
     'carretillero, 17 % material no disponible, 17 % bobina equivocada o mal etiquetada.',
     'Horas de termo parado por aprovisionamiento a línea'),
    ('P2R-05', 'Negocio', 'Estándar de reposición a pie de máquina y carretilleros por turno',
     'Definir el stock mínimo a pie de máquina por termo (bobinas y cajas para X horas), la regla de reposición y el número de '
     'carretilleros por turno según los cambios previstos en la secuencia.',
     'Las esperas se concentran en los turnos con varios cambios a la vez.', 'Paradas por aprovisionamiento a línea'),
    ('P2R-05', 'Automatización', 'Cobertura de bobina por orden de termo',
     'Antes de arrancar una orden en un termo, el sistema comprueba que hay bobina (ST) fabricada o planificada en la extrusora '
     'para terminarla, y avisa si la orden se quedará sin bobina a mitad.',
     'Espera por falta de bobinas fabricadas: 688 h/año. El 72 % ocurre con la extrusora en marcha fabricando otra cosa y el '
     '79 % de los episodios a mitad de pedido.', 'Horas de termo parado por falta de bobina'),
]


def hoja_propuesta(wb):
    ws = wb.create_sheet('Propuesta HT-06')
    titulo(ws, 'Propuesta: aprovisionamiento a línea (carretillero) y bobina dentro de HT-06',
           'El carretillero se imputa a HT-06. P2R-05 no tiene intervenciones en Ondupet: estas tres son la propuesta para darlas de alta.')
    cabecera(ws, 4, ['Palanca', 'De qué es', 'Intervención', 'Qué hay que hacer', 'Por qué (datos)', 'KPI'],
             [9, 15, 34, 70, 60, 30])
    for i, fila in enumerate(PROPUESTA):
        for j, v in enumerate(fila, 1):
            celda(ws, 5 + i, j, v, centrada=j in (1, 2), bold=j == 3)


def construir(salida):
    wb = Workbook()
    hoja_supuestos(wb)
    ult_b = hoja_bolsillos(wb)
    cols = hoja_reparto(wb, ult_b)
    hoja_resumen(wb, cols)
    hoja_datos(wb)
    hoja_avisos(wb)
    hoja_propuesta(wb)
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    wb.active = 0
    wb.save(salida)


if __name__ == '__main__':
    out = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).resolve().parents[1] / 'Impacto_Hallazgos_Ondupet.xlsx')
    construir(out)
    print(out)
