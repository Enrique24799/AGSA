"""Hoja de indicadores (KPIs) de Plan to Release por empresa para el consolidado de P2R.

Los comunes llevan el mismo nombre en las tres empresas, aunque se midan contra datos distintos; la hoja calcula
cuáles coinciden. «Se puede medir hoy» resume lo que dicen los E2E de cada planta.
"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font

from generar_intervenciones_ondupet import F, NAVY, cabecera, celda, fill, titulo

OTIF, PREV, PERD = 'OTIF', 'Previsión de la demanda', 'Demanda perdida'
SAT, CAMB, ESTAB, ADH = 'Saturación del cuello de botella', 'Horas de cambio', 'Estabilidad del programa', 'Adherencia al programa'

# (KPI, definición en la empresa, unidad, sentido, frecuencia, fuente, palanca, hallazgo, medible hoy, qué falta)
KPIS = {
    'Ondupet': [
        (OTIF, '% de líneas entregadas completas en la fecha del cliente (pasará a la fecha confirmada cuando se use)', '%', '↑',
         'Semanal', 'GIO y albaranes de SAP', 'P2R-08 · P2R-04', 'HT-P2R-11 · HT-P2R-05', 'Sí',
         'Usar la fecha confirmada del pedido.'),
        (PREV, 'Precisión (1 − WMAPE) y sesgo de la previsión cerrada frente a lo expedido, por familia', '%', '↑', 'Mensual',
         'GIO', 'P2R-01', 'HT-P2R-01', 'Parcial', 'Fijar fórmula y umbral; hoy se mira en la reunión sin KPI formal.'),
        (PERD, '€ / mes de agotados y de pedidos rechazados o reducidos', '€/mes', '↓', 'Mensual', 'GIO (registro diario de agotados)',
         'P2R-01', 'HT-P2R-01', 'Parcial', 'Registrar los agotados todos los días y con motivo.'),
        (SAT, 'Carga planificada / capacidad disponible de extrusión y termoformado', '%', 'banda 85–100 %', 'Semanal',
         'GIO (cálculo de % de saturación, ACC-03)', 'P2R-02', 'HT-P2R-03', 'Parcial', 'Que ACC-03 calcule la saturación por máquina.'),
        (CAMB, 'Horas de cambio de molde y altura por semana', 'h/semana', '↓', 'Semanal', 'Partes de producción', 'P2R-06',
         'HT-P2R-08', 'Por confirmar', 'Confirmar que los partes registran el tiempo de cambio.'),
        (ESTAB, 'Cambios de última hora dentro del programa por semana', 'nº/semana', '↓', 'Semanal', 'GIO', 'P2R-06', 'HT-P2R-09',
         'No', 'Definir la zona congelada y registrar cada cambio con su causa.'),
        (ADH, '% de órdenes fabricadas en el día y turno programados', '%', '↑', 'Semanal', 'GIO frente a partes de producción',
         'P2R-08 · P2R-06', 'HT-P2R-12', 'Por confirmar', 'Comparar programa y partes de producción por orden.'),
        ('Cambios de previsión tras el cierre', '€ y nº de cambios de la previsión del mes en curso después del día de cierre',
         '€/mes', '↓', 'Mensual', 'GIO (registro de cambios, ACC-01)', 'P2R-01', 'HT-P2R-01', 'No',
         'Día de cierre y bloqueo con registro de cambios.'),
        ('Cobertura de producto terminado', 'Stock de PT / consumo medio diario, por segmento ABC', 'días', '↓', 'Mensual',
         'GIO y SAP', 'P2R-03', 'HT-P2R-04', 'Parcial', 'Segmentar por ABC/XYZ (ACC-06).'),
    ],
    'Ondupack Almendralejo': [
        (OTIF, '% de líneas entregadas completas en la fecha propuesta por el CTP', '%', '↑', 'Semanal', 'GIO', 'P2R-08',
         'HT-P2R-11', 'Sí', 'Que Dirección y Comercial aprueben el KPI oficial (ACC-08).'),
        (PREV, '% de la demanda repetitiva que tiene previsión (pasará a precisión y sesgo cuando haya previsión)', '%', '↑',
         'Mensual', 'GIO', 'P2R-01', 'HT-P2R-01', 'Sí', 'Hoy es 0 %.'),
        (PERD, '€ / mes de pedidos rechazados, cantidades reducidas o fechas no aceptadas por capacidad', '€/mes', '↓', 'Mensual',
         'GIO (ACC-02)', 'P2R-01', 'HT-P2R-01', 'No', 'No se registra en el sistema.'),
        (SAT, 'Carga planificada / capacidad disponible de la 924, con la velocidad real', '%', 'banda 85–100 %', 'Semanal',
         'GIO (carga-capacidad)', 'P2R-02', 'HT-P2R-03', 'Parcial', 'Hoy usa la velocidad objetivo, que la 924 no alcanza.'),
        (CAMB, 'Horas de cambio en onduladora (canal y calidad) y en converting (escenarios SMED) por semana', 'h/semana', '↓',
         'Semanal', 'GIO (control de producción y escenarios SMED)', 'P2R-06', 'HT-P2R-08', 'Sí', ''),
        (ESTAB, 'Cambios del cliente aceptados con el pedido ya programado, por semana', 'nº/semana', '↓', 'Semanal',
         'GIO (trazabilidad del pedido)', 'P2R-06', 'HT-P2R-09', 'Parcial', 'Registrar la causa de cada reprogramación.'),
        (ADH, '% de órdenes fabricadas en las fechas previstas de control de producción', '%', '↑', 'Semanal',
         'GIO (control de producción)', 'P2R-08 · P2R-06', 'HT-P2R-12', 'Sí', ''),
        ('Velocidad real y OEE de la 924', 'Velocidad real / objetivo y OEE (rendimiento × disponibilidad × calidad)', '%', '↑',
         'Semanal', 'GIO y Click (históricos)', 'P2R-02', 'HT-P2R-03', 'Sí', ''),
        ('Órdenes liberadas con material pendiente', 'Órdenes autorizadas sin todo el material, por semana', 'nº/semana', '↓',
         'Semanal', 'GIO', 'P2R-05', 'HT-P2R-06', 'Sí', ''),
        ('Compras y transportes urgentes', 'Número y € de compras y transportes urgentes por faltas de materia prima', '€/mes',
         '↓', 'Mensual', 'SAP', 'P2R-05', 'HT-P2R-06', 'Por confirmar', 'Marcar en SAP las compras urgentes.'),
        ('Cobertura de papel', 'Stock de papel / consumo medio diario, por calidad', 'días', 'banda', 'Mensual',
         'SAP e inventario físico', 'P2R-05', 'HT-P2R-07', 'Parcial', 'El stock solo es fiable tras el inventario mensual.'),
        ('Ocupación del almacén intermedio', 'm² ocupados / m² disponibles de plancha (WIP)', '%', '↓', 'Diaria',
         'GIO (saturación WIP)', 'P2R-06', 'HT-P2R-08', 'Sí', ''),
    ],
    'MGT': [
        (OTIF, '% de líneas entregadas completas: externos contra la fecha pedida, Ondupack contra la fecha del ciclo', '%', '↑',
         'Semanal', 'SAP', 'P2R-08 · P2R-04', 'HT-P2R-11 · HT-P2R-05', 'Parcial',
         'No hay fecha confirmada y Ondupack cambia las fechas después.'),
        (PREV, '% de la demanda (Ondupack y clientes repetitivos) que tiene previsión por calidad y ancho', '%', '↑', 'Por ciclo',
         'GIO', 'P2R-01', 'HT-P2R-01', 'Sí', 'Hoy es 0 %.'),
        (PERD, '€ / mes de pedidos rechazados o fechas no aceptadas por capacidad', '€/mes', '↓', 'Mensual',
         'Salesforce (O2C)', 'P2R-01', 'HT-P2R-01', 'No', 'No se sabe si se registran las ventas perdidas.'),
        (SAT, 'Carga con cartera / capacidad de la máquina a 3 turnos, por ciclo y calidad', '%', 'banda 85–100 %', 'Por ciclo',
         'GIO y partes de producción', 'P2R-02', 'HT-P2R-03', 'No', 'Hoy se carga a capacidad infinita.'),
        (CAMB, 'Horas de cambio de calidad y ancho por ciclo', 'h/ciclo', '↓', 'Por ciclo', 'Partes de producción', 'P2R-06',
         'HT-P2R-08', 'Por confirmar', 'Confirmar que los partes registran el tiempo de cambio.'),
        (ESTAB, 'Urgencias encajadas en el ciclo, también las de Ondupack fuera de plazo, por semana', 'nº/semana', '↓',
         'Semanal', 'GIO', 'P2R-06', 'HT-P2R-09', 'No', 'Registrar cada urgencia con su origen.'),
        (ADH, '% de órdenes fabricadas en el ciclo y día previstos', '%', '↑', 'Semanal', 'GIO frente a partes de producción',
         'P2R-08 · P2R-06', 'HT-P2R-12', 'Por confirmar', 'Comparar plan del ciclo y partes de producción.'),
        ('Stock de Ondupack en MGT', '€ y días de cobertura de las referencias de Ondupack en el almacén de MGT', '€ · días', '↓',
         'Semanal', 'SAP', 'P2R-03', 'HT-P2R-04', 'Sí', ''),
        ('Cobertura y exactitud del stock de materiales', 'Días de cobertura de materias primas y auxiliares y % de descuadre en el '
         'cuadre quincenal', 'días · %', '↓', 'Mensual', 'SAP', 'P2R-05', 'HT-P2R-07', 'Parcial',
         'Registrar las entradas de báscula en el momento.'),
        ('Liberación sin comprobación en sistema', '% de órdenes liberadas con comprobación manual del material o entregadas en papel',
         '%', '↓', 'Semanal', 'GIO', 'P2R-07', 'HT-P2R-10', 'Sí', 'Hoy es el 100 %.'),
        ('Coordinación diaria con Ondupack', 'Correos al día para coordinar stock y cargas con Ondupack', 'nº/día', '↓', 'Diaria',
         'Correo', 'P2R-08', 'HT-P2R-11', 'Sí', 'Hoy son unos 16.'),
        ('Desviación pedido – previsión de Ondupack', '|pedido − previsión| / previsión de papel de Ondupack, por ciclo', '%',
         '↓', 'Por ciclo', 'GIO', 'P2R-01', 'HT-P2R-01', 'No', 'Ondupack aún no envía previsión (su ACC-06).'),
    ],
}
ORDEN = ('Ondupet', 'MGT', 'Ondupack Almendralejo')


def hoja_kpis(wb):
    ws = wb.create_sheet('KPIs')
    titulo(ws, 'Indicadores de Plan to Release por empresa',
           'Los comunes llevan el mismo nombre en las tres empresas aunque cada una los mida con sus datos; la columna «Coincide» '
           'lo calcula. «Se puede medir hoy» resume los E2E de cada planta.')
    cab = ['Empresa', 'KPI', 'Coincide', 'Definición en la empresa', 'Unidad', 'Sentido', 'Frecuencia', 'Fuente', 'Palanca',
           'Hallazgo', 'Se puede medir hoy', 'Qué falta para medirlo']
    fila = 6
    cabecera(ws, fila, cab, [20, 30, 15, 58, 10, 13, 11, 30, 15, 22, 13, 44])
    r = fila + 1
    filas = [(emp, k) for emp in ORDEN for k in KPIS[emp]]
    ult = fila + len(filas)
    for emp, (kpi, defi, uni, sen, frec, fuente, pal, hall, hoy, falta) in filas:
        coincide = (f'=IF(COUNTIF($B${fila + 1}:$B${ult},B{r})=3,"En las 3",'
                    f'IF(COUNTIF($B${fila + 1}:$B${ult},B{r})=2,"En 2","Específico"))')
        vals = [emp, kpi, coincide, defi, uni, sen, frec, fuente, pal, hall, hoy, falta or None]
        for j, v in enumerate(vals, 1):
            celda(ws, r, j, v, centrada=j in (3, 5, 6, 7, 11), bold=j == 2)
        r += 1
    for cond, color in (('"Sí"', 'E3F3EA'), ('"Parcial"', 'FFF1D6'), ('"Por confirmar"', 'FFF1D6'), ('"No"', 'FCE4E4')):
        ws.conditional_formatting.add(f'K{fila + 1}:K{ult}', FormulaRule(formula=[f'$K{fila + 1}={cond}'], fill=fill(color)))
    ws.conditional_formatting.add(f'C{fila + 1}:C{ult}', FormulaRule(formula=[f'$C{fila + 1}="En las 3"'],
                                                                  font=Font(name=F, bold=True, color=NAVY)))
    # resumen por empresa
    rng = lambda col: f'${col}${fila + 1}:${col}${ult}'
    for k, emp in enumerate(ORDEN):
        col = 1 + 3 * k
        a = ws.cell(4, col, emp)
        a.font = Font(name=F, size=10, bold=True, color=NAVY)
        b = ws.cell(4, col + 1, f'=COUNTIF({rng("A")},"{emp}")&" KPIs · medibles hoy: "&COUNTIFS({rng("A")},"{emp}",{rng("K")},"Sí")')
        b.font = Font(name=F, size=10, color='5B6B78')
    ws.freeze_panes = ws.cell(fila + 1, 3)
    ws.auto_filter.ref = f'A{fila}:L{ult}'
    ws.sheet_view.showGridLines = False
    return ws
