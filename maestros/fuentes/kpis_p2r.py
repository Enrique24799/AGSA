"""Indicadores (KPIs) de Plan to Release para el consolidado de P2R.

En DICE la definición del KPI es de grupo y vale para las 19 empresas, así que el catálogo lleva nombres genéricos y un
único nombre, unidad, sentido, periodicidad y agregación por KPI (los campos de la ficha «Nuevo KPI»). La hoja por
empresa dice cómo lo mide cada una y toma unidad, sentido y periodicidad del catálogo.
"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font

from generar_intervenciones_ondupet import F, NAVY, cabecera, celda, fill, titulo

SUBE, BAJA = 'mejor si sube', 'mejor si baja'
SUMA, MEDIA = 'suma', 'media'
SEM, MES = 'semanal', 'mensual'

OTIF = 'OTIF'
ERR = 'Error de previsión (WMAPE)'
COBPREV = 'Cobertura de previsión'
CAMBPREV = 'Cambios de previsión tras el cierre'
PERD = 'Demanda perdida'
SAT = 'Saturación del cuello de botella'
OEE = 'OEE del cuello de botella'
CAMB = 'Horas de cambio'
PROG = 'Cambios en el programa cerrado'
ADH = 'Adherencia al programa'
COBPT = 'Cobertura de producto terminado'
COBMP = 'Cobertura de materias primas'
EXACT = 'Exactitud del stock de materiales'
STKIC = 'Stock intercompany'
LIBMAT = 'Órdenes liberadas con material pendiente'
LIBSIS = 'Órdenes liberadas sin comprobación en sistema'
URG = 'Compras y transportes urgentes'
OCUP = 'Ocupación del almacén intermedio'
CORR = 'Correos de coordinación intercompany'

# Ficha «Nuevo KPI»: (nombre, unidad, sentido, periodicidad, agregación de dimensiones, qué mide para todo el grupo)
CATALOGO = [
    (OTIF, '%', SUBE, SEM, MEDIA, 'Líneas de pedido entregadas completas y en fecha / líneas entregadas.'),
    (ERR, '%', BAJA, MES, MEDIA, 'Σ|real − previsión| / Σ real, por familia. Solo donde hay previsión; el sesgo se mira en el '
     'mismo análisis.'),
    (COBPREV, '%', SUBE, MES, MEDIA, 'Demanda que tiene previsión / demanda total. Para empresas que aún no prevén.'),
    (CAMBPREV, '€', BAJA, MES, SUMA, '€ de la previsión del mes en curso que se cambian después del día de cierre.'),
    (PERD, '€', BAJA, MES, SUMA, '€ de pedidos rechazados, reducidos o con fecha no aceptada, y de agotados.'),
    (SAT, '%', SUBE, SEM, MEDIA, 'Carga planificada / capacidad disponible del recurso cuello de botella. Objetivo 85–100 %: '
     'por encima es sobrecarga.'),
    (OEE, '%', SUBE, SEM, MEDIA, 'Disponibilidad × rendimiento (velocidad real / objetivo) × calidad del cuello de botella.'),
    (CAMB, 'h', BAJA, SEM, SUMA, 'Horas de cambio de formato, calidad o ancho en el periodo.'),
    (PROG, 'nº', BAJA, SEM, SUMA, 'Cambios y urgencias metidos en el programa después de cerrarlo.'),
    (ADH, '%', SUBE, SEM, MEDIA, 'Órdenes fabricadas en el día o turno programado / órdenes programadas.'),
    (COBPT, 'días', BAJA, MES, MEDIA, 'Stock de producto terminado / consumo medio diario.'),
    (COBMP, 'días', BAJA, MES, MEDIA, 'Stock de materias primas / consumo medio diario.'),
    (EXACT, '%', BAJA, MES, MEDIA, 'Descuadre entre el stock del sistema y el recuento físico / stock.'),
    (STKIC, '€', BAJA, SEM, SUMA, 'Stock que la empresa guarda para otras empresas del grupo.'),
    (LIBMAT, '%', BAJA, SEM, MEDIA, 'Órdenes liberadas sin todo el material disponible / órdenes liberadas.'),
    (LIBSIS, '%', BAJA, SEM, MEDIA, 'Órdenes liberadas con comprobación manual del material o en papel / órdenes liberadas.'),
    (URG, '€', BAJA, MES, SUMA, '€ de compras y transportes urgentes por faltas de material.'),
    (OCUP, '%', BAJA, SEM, MEDIA, 'Espacio ocupado / espacio disponible del almacén de semielaborado.'),
    (CORR, 'nº', BAJA, SEM, SUMA, 'Correos enviados para coordinar stock, cargas y fechas con otras empresas del grupo.'),
]

# (KPI del catálogo, cómo lo mide la empresa, fuente, palanca, hallazgo, medible hoy, qué falta)
KPIS = {
    'Ondupet': [
        (OTIF, 'Contra la fecha del cliente (pasará a la fecha confirmada cuando se use)', 'GIO y albaranes de SAP',
         'P2R-08 · P2R-04', 'HT-P2R-11 · HT-P2R-05', 'Sí', 'Usar la fecha confirmada del pedido.'),
        (ERR, 'Previsión cerrada frente a lo expedido, por familia; con su sesgo', 'GIO', 'P2R-01', 'HT-P2R-01', 'Parcial',
         'Fijar fórmula y umbral; hoy se mira en la reunión sin KPI formal.'),
        (PERD, 'Agotados y pedidos rechazados o reducidos', 'GIO (registro diario de agotados)', 'P2R-01', 'HT-P2R-01',
         'Parcial', 'Registrar los agotados todos los días y con motivo.'),
        (SAT, 'Extrusión y termoformado', 'GIO (cálculo de % de saturación, ACC-03)', 'P2R-02', 'HT-P2R-03', 'Parcial',
         'Que ACC-03 calcule la saturación por máquina.'),
        (CAMB, 'Cambios de molde y altura', 'Partes de producción', 'P2R-06', 'HT-P2R-08', 'Por confirmar',
         'Confirmar que los partes registran el tiempo de cambio.'),
        (PROG, 'Cambios de última hora dentro del programa', 'GIO', 'P2R-06', 'HT-P2R-09', 'No',
         'Definir la zona congelada y registrar cada cambio con su causa.'),
        (ADH, 'Órdenes fabricadas en el día y turno programados', 'GIO frente a partes de producción', 'P2R-08 · P2R-06',
         'HT-P2R-12', 'Por confirmar', 'Comparar programa y partes de producción por orden.'),
        (CAMBPREV, 'Cambios de la previsión del mes en curso después del día de cierre', 'GIO (registro de cambios, ACC-01)',
         'P2R-01', 'HT-P2R-01', 'No', 'Día de cierre y bloqueo con registro de cambios.'),
        (COBPT, 'Por segmento ABC', 'GIO y SAP', 'P2R-03', 'HT-P2R-04', 'Parcial', 'Segmentar por ABC/XYZ (ACC-06).'),
    ],
    'Ondupack Almendralejo': [
        (OTIF, 'Contra la fecha propuesta por el CTP', 'GIO', 'P2R-08', 'HT-P2R-11', 'Sí',
         'Que Dirección y Comercial aprueben el KPI oficial (ACC-08).'),
        (COBPREV, 'Demanda repetitiva con previsión', 'GIO', 'P2R-01', 'HT-P2R-01', 'Sí', 'Hoy es 0 %.'),
        (PERD, 'Pedidos rechazados, cantidades reducidas o fechas no aceptadas por capacidad', 'GIO (ACC-02)', 'P2R-01',
         'HT-P2R-01', 'No', 'No se registra en el sistema.'),
        (SAT, 'La 924, con la velocidad real', 'GIO (carga-capacidad)', 'P2R-02', 'HT-P2R-03', 'Parcial',
         'Hoy usa la velocidad objetivo, que la 924 no alcanza.'),
        (CAMB, 'Onduladora (canal y calidad) y converting (escenarios SMED)', 'GIO (control de producción y escenarios SMED)',
         'P2R-06', 'HT-P2R-08', 'Sí', ''),
        (PROG, 'Cambios del cliente aceptados con el pedido ya programado', 'GIO (trazabilidad del pedido)', 'P2R-06',
         'HT-P2R-09', 'Parcial', 'Registrar la causa de cada reprogramación.'),
        (ADH, 'Órdenes fabricadas en las fechas previstas de control de producción', 'GIO (control de producción)',
         'P2R-08 · P2R-06', 'HT-P2R-12', 'Sí', ''),
        (OEE, 'La 924', 'GIO y Click (históricos)', 'P2R-02', 'HT-P2R-03', 'Sí', ''),
        (LIBMAT, 'Órdenes autorizadas sin todo el material', 'GIO', 'P2R-05', 'HT-P2R-06', 'Sí', ''),
        (URG, 'Por faltas de materia prima', 'SAP', 'P2R-05', 'HT-P2R-06', 'Por confirmar', 'Marcar en SAP las compras urgentes.'),
        (COBMP, 'Papel, por calidad', 'SAP e inventario físico', 'P2R-05', 'HT-P2R-07', 'Parcial',
         'El stock solo es fiable tras el inventario mensual.'),
        (OCUP, 'Plancha en el almacén intermedio (WIP), en m²', 'GIO (saturación WIP)', 'P2R-06', 'HT-P2R-08', 'Sí', ''),
    ],
    'MGT': [
        (OTIF, 'Externos contra la fecha pedida; Ondupack contra la fecha del ciclo', 'SAP', 'P2R-08 · P2R-04',
         'HT-P2R-11 · HT-P2R-05', 'Parcial', 'No hay fecha confirmada y Ondupack cambia las fechas después.'),
        (COBPREV, 'Demanda de Ondupack y de clientes repetitivos con previsión, por calidad y ancho', 'GIO', 'P2R-01',
         'HT-P2R-01', 'Sí', 'Hoy es 0 %.'),
        (PERD, 'Pedidos rechazados o fechas no aceptadas por capacidad', 'Salesforce (O2C)', 'P2R-01', 'HT-P2R-01', 'No',
         'No se sabe si se registran las ventas perdidas.'),
        (SAT, 'Máquina de papel a 3 turnos, por ciclo y calidad', 'GIO y partes de producción', 'P2R-02', 'HT-P2R-03', 'No',
         'Hoy se carga a capacidad infinita.'),
        (CAMB, 'Cambios de calidad y ancho', 'Partes de producción', 'P2R-06', 'HT-P2R-08', 'Por confirmar',
         'Confirmar que los partes registran el tiempo de cambio.'),
        (PROG, 'Urgencias encajadas en el ciclo, también las de Ondupack fuera de plazo', 'GIO', 'P2R-06', 'HT-P2R-09', 'No',
         'Registrar cada urgencia con su origen.'),
        (ADH, 'Órdenes fabricadas en el ciclo y día previstos', 'GIO frente a partes de producción', 'P2R-08 · P2R-06',
         'HT-P2R-12', 'Por confirmar', 'Comparar plan del ciclo y partes de producción.'),
        (STKIC, 'Referencias de Ondupack en el almacén de MGT', 'SAP', 'P2R-03', 'HT-P2R-04', 'Sí', ''),
        (COBMP, 'Materias primas y auxiliares', 'SAP', 'P2R-05', 'HT-P2R-07', 'Parcial',
         'Registrar las entradas de báscula en el momento.'),
        (EXACT, 'En el cuadre quincenal', 'SAP', 'P2R-05', 'HT-P2R-07', 'Parcial',
         'Registrar las entradas de báscula en el momento.'),
        (LIBSIS, 'Comprobación manual del material o entrega en papel', 'GIO', 'P2R-07', 'HT-P2R-10', 'Sí', 'Hoy es el 100 %.'),
        (CORR, 'Con Ondupack, para stock y cargas', 'Correo', 'P2R-08', 'HT-P2R-11', 'Sí', 'Hoy son unos 16 al día.'),
        (ERR, 'Pedido de papel de Ondupack frente a su previsión, por ciclo', 'GIO', 'P2R-01', 'HT-P2R-01', 'No',
         'Ondupack aún no envía previsión (su ACC-06).'),
    ],
}
ORDEN = ('Ondupet', 'MGT', 'Ondupack Almendralejo')
CAT = 'Catálogo de KPIs'
CAT_FILA = 4


def hoja_catalogo(wb):
    ws = wb.create_sheet(CAT)
    titulo(ws, 'Catálogo de KPIs de Plan to Release',
           'Definición de grupo, válida para las 19 empresas: los campos de la ficha «Nuevo KPI» de DICE. '
           'Cómo lo mide cada empresa está en la hoja KPIs.')
    cab = ['Nombre', 'Unidad', 'Sentido', 'Periodicidad', 'Agregación de dimensiones', 'Qué mide', 'Empresas que lo usan']
    cabecera(ws, CAT_FILA, cab, [40, 9, 15, 13, 16, 70, 14])
    for i, (nombre, uni, sen, per, agr, que) in enumerate(CATALOGO):
        r = CAT_FILA + 1 + i
        for j, v in enumerate([nombre, uni, sen, per, agr, que], 1):
            celda(ws, r, j, v, centrada=j in (2, 3, 4, 5), bold=j == 1)
        celda(ws, r, 7, f'=COUNTIF(KPIs!$B:$B,A{r})', centrada=True)
    ult = CAT_FILA + len(CATALOGO)
    for valor, color in ((SUBE, 'E3F3EA'), (BAJA, 'FCE4E4')):
        ws.conditional_formatting.add(f'C{CAT_FILA + 1}:C{ult}',
                                      FormulaRule(formula=[f'$C{CAT_FILA + 1}="{valor}"'], fill=fill(color)))
    ws.freeze_panes = ws.cell(CAT_FILA + 1, 2)
    return ws


def hoja_kpis(wb):
    hoja_catalogo(wb)
    nombres = {c[0] for c in CATALOGO}
    for emp, filas in KPIS.items():
        for k in filas:
            if k[0] not in nombres:
                raise ValueError(f'{emp}: «{k[0]}» no está en el catálogo')
    ws = wb.create_sheet('KPIs')
    titulo(ws, 'Indicadores de Plan to Release por empresa',
           'El nombre, la unidad, el sentido y la periodicidad vienen del catálogo de grupo; cada empresa dice cómo lo mide. '
           '«Coincide» cuenta en cuántas de las tres está. «Se puede medir hoy» resume los E2E de cada planta.')
    cab = ['Empresa', 'KPI', 'Coincide', 'Cómo lo mide la empresa', 'Unidad', 'Sentido', 'Periodicidad', 'Fuente', 'Palanca',
           'Hallazgo', 'Se puede medir hoy', 'Qué falta para medirlo']
    fila = 6
    cabecera(ws, fila, cab, [20, 34, 13, 52, 9, 15, 13, 30, 15, 22, 13, 44])
    c_ini, c_fin = CAT_FILA + 1, CAT_FILA + len(CATALOGO)

    def buscar(col, r):
        return (f"=INDEX('{CAT}'!${col}${c_ini}:${col}${c_fin},"
                f"MATCH($B{r},'{CAT}'!$A${c_ini}:$A${c_fin},0))")

    filas = [(emp, k) for emp in ORDEN for k in KPIS[emp]]
    ult = fila + len(filas)
    for i, (emp, (kpi, como, fuente, pal, hall, hoy, falta)) in enumerate(filas):
        r = fila + 1 + i
        coincide = (f'=IF(COUNTIF($B${fila + 1}:$B${ult},B{r})=3,"En las 3",'
                    f'IF(COUNTIF($B${fila + 1}:$B${ult},B{r})=2,"En 2","Específico"))')
        vals = [emp, kpi, coincide, como, buscar('B', r), buscar('C', r), buscar('D', r), fuente, pal, hall, hoy,
                falta or None]
        for j, v in enumerate(vals, 1):
            celda(ws, r, j, v, centrada=j in (3, 5, 6, 7, 11), bold=j == 2)
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
