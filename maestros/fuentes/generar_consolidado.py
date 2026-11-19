"""Genera el Excel consolidado de P2R con lo que hay dado de alta en la aplicación (DICE Studio).

Hojas: fases del esqueleto, personas, dedicación, catálogo de hallazgos, catálogo de palancas y una hoja por empresa
(hallazgos de la celda, palanca que los cura e intervenciones).
Los datos se transcriben de las capturas de la aplicación; las descripciones de hallazgos y palancas vienen
del maestro (generar_maestro_p2r.py).

Uso: python generar_consolidado.py [salida.xlsx]
"""
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font

from generar_intervenciones_ondupet import F, cabecera, celda, fill, titulo
from hojas_empresa import hoja_empresa, hoja_fichas
from fichas_palancas import hojas_fichas
from kpis_p2r import hoja_kpis

FECHA = '19/11/2026'
EMPRESAS = ['Ondupet', 'Ondupack', 'MGT']
F2P, P2P = 'Forecast to Plan', 'Plan to Produce'
DEM, SOP, NEC = 'Demanda y previsión (Forecast to Plan)', 'S&OP y Plan Maestro (Forecast to Plan)', 'Necesidades de Fabricación y MRP'
PRO, LIB, SEG = 'Programación', 'Liberación', 'Seguimiento y Control'

# (orden, fase, tramo, descripción, opcional, áreas de negocio, actividades)
FASES = [
    (1, DEM, F2P, 'De Histórico y Señal de Mercado --> Demanda Consensuada', '—', 'todas', 7),
    (3, SOP, F2P, 'De Demanda Global --> Equilibrio Demanda, Capacidad y Recursos', '—', 'todas', 7),
    (5, NEC, P2P, 'De Plan Maestro --> Órdenes Previsionales de Fabricación y Solicitudes de Compra', '—', 'todas', 8),
    (6, PRO, P2P, 'De Orden Previsional --> Secuencia por Línea y Turno', '—', 'todas', 6),
    (7, LIB, P2P, 'De Orden Programada --> Orden Liberada Ejecutable', '—', 'todas', 4),
    (8, SEG, P2P, 'De Plan Ejecutado --> Desviaciones Explicadas con Acción', '—', 'todas', 9),
]

# (persona, id, email, categoría, roles, censo en, área de negocio, jornada h/año, coste anual k€, Σ dedicación, licencias)
PERSONAS = [
    ('Victoriano Pérez Carrasco', 'ID-20', 'vperez@clgrupoindustrial.com', 'Técnico especialista',
     'Planificación · Planificador de la producción', 'Ondupet', 'Packaging', 1720, 60.2, 1.00, 'sin licencias'),
    ('Bárbara Ayuste Tejera', 'ID-22', 'bayuste@ondupet.es', 'Administrativo',
     'Compras · Category Manager; Compras · Tramitadores de compras', 'Ondupet', 'Packaging', 1720, 37.8, 0.25, 'SAP ECC · usuario'),
    ('Manuel Jesús Trinidad Montero', 'ID-24', 'mjtrinidad@ondupet.es', 'Responsable / mando intermedio',
     'Comercial · Comercial con potestad', 'Ondupet', 'Packaging', 1720, 82.6, 0.15, 'sin licencias'),
    ('Monica Gordillo Guerrero', 'ID-21', 'mgordillo@ondupet.es', 'Administrativo',
     'Comercial · Comercial con potestad; Comercial · Administración de ventas', 'Ondupet', 'Packaging', 1720, 37.8, 0.20,
     'sin licencias'),
    ('Antonio Nogales Carrión', 'ID-23', 'anogales@ondupet.es', 'Responsable / mando intermedio',
     'Compras · Category Manager; Compras · Tramitadores de compras', 'Ondupet', 'Packaging', 1720, 82.6, 0.0, ''),
    ('Miguel Ángel Rosado Martínez', 'sin ID', 'planificacion@ondupack.com', 'Técnico especialista',
     'Planificación · Planificador de la producción; Logística · Responsable Expediciones', 'Ondupack', 'Cartón', 1720, 60.2, 0.0, ''),
    ('Tomás Buenavida Santiago', 'sin ID', 'tbuenavida@ondupack.com', 'Administrativo',
     'Planificación · Planificador de la producción', 'Ondupack', 'Cartón', 1720, 37.8, 0.0, ''),
    ('Maria Teresa Espinosa Paredes', 'sin ID', 'tere_att@ondupack.com', 'Técnico especialista',
     'Comercial', 'Ondupack', 'Cartón', 1720, 60.2, 0.10, ''),
    ('José María Baquero Jaen', 'sin ID', 'jbaquero@ondupack.com', 'Responsable / mando intermedio',
     'Producción · Responsable Producción', 'Ondupack', 'Cartón', 1720, 82.6, 0.0, ''),
    ('Angélica Lemus Fernández', 'sin ID', 'alemus@ondupack.com', 'Responsable / mando intermedio',
     'Compras · Category Manager', 'Ondupack', 'Cartón', 1720, 82.6, 0.0, ''),
    ('Karim Karach Boujtat', 'sin ID', 'kkarach@ondupack.com', 'Responsable / mando intermedio',
     'Producción · Responsable Producción', 'Navalmoral', 'Cartón', 1720, 82.6, 0.0, ''),
    ('Joan Carles Girones Campaña', 'sin ID', 'jc.girones@matiasgomatomas.com', 'Técnico especialista',
     'Planificación · Planificador de la producción', 'MGT', 'Papel', 1720, 60.2, 0.0, ''),
    ('Marta Boyo Martinez', 'sin ID', 'm.boyo@matiasgomatomas.com', 'Técnico especialista',
     'Compras · Category Manager', 'MGT', 'Papel', 1720, 60.2, 0.0, ''),
    ('Albert Fontanet Lopez', 'sin ID', 'a.fontanet@matiasgomatomas.com', 'Responsable / mando intermedio',
     'Comercial', 'MGT', 'Papel', 1720, 82.6, 0.0, ''),
]

# (persona, celda, fase, % de dedicación)
DEDICACION = [
    ('Bárbara Ayuste Tejera', 'P2R × Ondupet', NEC, 0.25),
    ('Manuel Jesús Trinidad Montero', 'P2R × Ondupet', SOP, 0.15),
    ('Monica Gordillo Guerrero', 'P2R × Ondupet', SOP, 0.20),
    ('Victoriano Pérez Carrasco', 'P2R × Ondupet', NEC, 0.30),
    ('Victoriano Pérez Carrasco', 'P2R × Ondupet', LIB, 0.25),
    ('Victoriano Pérez Carrasco', 'P2R × Ondupet', PRO, 0.25),
    ('Victoriano Pérez Carrasco', 'P2R × Ondupet', SEG, 0.15),
    ('Victoriano Pérez Carrasco', 'P2R × Ondupet', DEM, 0.05),
]
CELDAS = ['P2R × Ondupet', 'P2R × Ondupack', 'P2R × MGT']

S = '✓'
# (id, nombre, fase, descripción, familia, naturaleza, métrica, se cura con, caza inversa, aplicado en Ondupet/Ondupack/MGT)
HALLAZGOS = [
    ('HT-P2R-01', 'Roturas de stock', DEM,
     'Falta producto terminado que debería estar en stock (infracobertura): hay agotados y pedidos que no se sirven o se '
     'sirven tarde.',
     'Oportunidad', '—', '€ de ventas / año en líneas de pedido con agotado de producto terminado', 'P2R-01', '—', (S, S, S)),
    ('HT-P2R-03', 'Optimización carga capacidad', SOP,
     'La carga no se equilibra con la capacidad: hay turnos pagados sin carga suficiente y, a la vez, líneas sobrecargadas.',
     'Ineficiencia', 'Transformación', 'Horas de línea / año de turnos no ajustados a la carga', 'P2R-02', '—', (S, S, S)),
    ('HT-P2R-05', 'Fechas de entrega poco realistas', SOP,
     'La fecha al cliente se promete sin comprobar disponibilidad ni capacidad y no se renegocia cuando no se llega.',
     'Oportunidad', '—', '€ de ventas / año en pedidos entregados después de la fecha pedida', 'P2R-04', 'Ondupack',
     (S, '', S)),
    ('HT-P2R-06', 'Optimización gestión faltas', NEC,
     'Las máquinas paran por falta de materia prima, auxiliares o semielaborado porque las faltas se detectan tarde. '
     'Incluye las horas de máquina paradas.',
     'Ineficiencia', 'Fuga de P&L', 'Horas de máquina / año paradas por falta de material', 'P2R-05', 'MGT', (S, S, '')),
    ('HT-P2R-07', 'Optimización gestión necesidades', NEC,
     'Las necesidades de materias primas y semielaborado no salen del plan o usan parámetros desactualizados: sobra stock '
     'de unos materiales y falta de otros.',
     'Ineficiencia', 'Circulante', '€ de stock de materias primas y semielaborado por encima de la cobertura objetivo',
     'P2R-05', 'Ondupet', ('', S, S)),
    ('HT-P2R-08', 'Optimización secuenciación', PRO,
     'La secuencia no minimiza los cambios (molde, altura, calidad, ancho): se hacen más cambios por unidad producida de '
     'los necesarios.',
     'Ineficiencia', 'Transformación', 'Horas de línea / año en cambios evitables', 'P2R-06', '—', (S, S, S)),
    ('HT-P2R-09', 'Optimización urgencias secuenciación', PRO,
     'Sin zona congelada ni reglas de prioridad, las urgencias rompen el programa cerrado y añaden cambios y '
     'resecuenciaciones.',
     'Ineficiencia', 'Transformación', 'Horas de línea / año en cambios metidos en el programa cerrado', 'P2R-06', '—',
     (S, S, S)),
    ('HT-P2R-10', 'Optimización aprovisionamiento línea', LIB,
     'Se libera sin confirmar material, utillaje y aprovisionamiento a línea, y la orden arranca con lo que falta.',
     'Ineficiencia', 'Transformación', 'Horas de línea / año perdidas en arranques fallidos', 'P2R-07',
     'Ondupack, Ondupet', ('', '', S)),
    ('HT-P2R-11', 'Optimización servicio', SEG,
     'El servicio no se mide ni se gestiona: los atrasos se ven cuando ya han ocurrido y no dejan causa.',
     'Oportunidad', '—', '€ de ventas / año en pedidos fuera de OTIF', 'P2R-08', '—', (S, S, S)),
    ('HT-P2R-12', 'Secuenciación manual', PRO,
     'Tiempo del planificador secuenciando y resecuenciando a mano.',
     'Ineficiencia', 'Esfuerzo', 'Horas de persona / año secuenciando a mano', 'P2R-06', 'MGT, Ondupack', (S, '', '')),
    ('HT-P2R-13', 'Sobrecobertura de stock', SOP,
     'Hay producto terminado con más días de stock que su objetivo: se fabrica de más o contra una previsión inflada.',
     'Ineficiencia', 'Circulante', '€ de stock de producto terminado por encima de la cobertura objetivo', 'P2R-01',
     'MGT, Ondupack', (S, '', '')),
    ('HT-P2R-14', 'Gestión manual servicio', SEG,
     'Tiempo del planificador en el seguimiento manual del servicio: reportes de atrasos, avisos y ajustes pedido a pedido.',
     'Ineficiencia', 'Esfuerzo', 'Horas de persona / año en seguimiento manual del servicio', 'P2R-08', 'MGT, Ondupack',
     (S, '', '')),
    ('HT-P2R-15', 'Gestión manual faltas', NEC,
     'Tiempo del planificador y de Compras revisando y persiguiendo las faltas de material una a una.',
     'Ineficiencia', 'Esfuerzo', 'Horas de persona / año gestionando faltas a mano', 'P2R-05', 'MGT, Ondupack',
     (S, '', '')),
]

# (código, palanca, qué cambia, base expuesta, fase as-is, cura tipos, KPI, aplicada en Ondupet/Ondupack/MGT)
# P2R-04 y P2R-07 salen truncados en la aplicación: nombre por confirmar.
PALANCAS = [
    ('P2R-01', 'Gestión de coberturas',
     'Cada referencia tiene una cobertura objetivo según su segmento y su demanda (previsión cerrada y consensuada), y se '
     'corrige antes de que falte (rotura) o sobre (sobrecobertura).',
     '€ de ventas / año con agotado · € de stock de producto terminado por encima de la cobertura objetivo',
     DEM, 'HT-P2R-01, HT-P2R-13', 'Roturas de PT · días de stock de PT', (S, S, S)),
    ('P2R-02', 'Plan de capacidad',
     'Carga y capacidad se equilibran cada mes, con datos: turnos, reservas de capacidad e intercompany se deciden con '
     'escenarios, no con urgencias.',
     'Horas de línea / año de turnos no ajustados a la carga', SOP, 'HT-P2R-03', '% de saturación del cuello de botella',
     (S, S, S)),
    ('P2R-04', 'Promesa de fecha mediante CTP',
     'La fecha al cliente se confirma con disponibilidad y capacidad reales (CTP) y se renegocia en cuanto se sabe que no se '
     'llega.',
     '€ de ventas / año en pedidos no entregados en la fecha pedida', SOP, 'HT-P2R-05', 'OTIF contra la fecha pedida',
     (S, '', S)),
    ('P2R-05', 'Análisis de las excepciones',
     'Las necesidades se calculan a tiempo y con parámetros vigentes, y el planificador solo atiende las faltas y excepciones '
     'críticas, antes de que paren la máquina.',
     'Horas de máquina / año paradas por falta de material · € de stock de materias primas y semielaborado por encima de la '
     'cobertura objetivo · horas de persona / año gestionando faltas',
     NEC, 'HT-P2R-06, HT-P2R-07, HT-P2R-15', 'Paradas por falta de material · días de stock MMPP y ST', (S, S, S)),
    ('P2R-06', 'Programación estable y optimizada',
     'El programa respeta la capacidad, minimiza los cambios y solo se modifica dentro de reglas: zona congelada y '
     'prioridades.',
     'Horas de línea / año en cambios evitables y en cambios del programa cerrado · horas de persona / año secuenciando',
     PRO, 'HT-P2R-08, HT-P2R-09, HT-P2R-12', 'Cambios / producción · cambios en el programa cerrado', (S, S, S)),
    ('P2R-07', 'Liberación con confirmación',
     'Una orden solo se libera con material, utillaje y aprovisionamiento a línea confirmados, y dentro de su ventana.',
     'Horas de línea / año perdidas en arranques fallidos', LIB, 'HT-P2R-10', '% de órdenes liberadas sin confirmación',
     ('', '', S)),
    ('P2R-08', 'Control del servicio',
     'Servicio y adherencia se miden solos, los riesgos se avisan antes y cada desviación tiene causa y una acción que vuelve '
     'a la planificación.',
     '€ de ventas / año en pedidos fuera de OTIF · horas de persona / año en seguimiento manual', SEG, 'HT-P2R-11, HT-P2R-14',
     'OTIF contra la fecha confirmada', (S, S, S)),
]

AVISO = 'FCE4E4'

# KPI y fórmula sugeridos por tipo (tokens de la aplicación: valor, obj). Resultado en €/año; las constantes son de
# ejemplo y se ajustan en cada celda.
KPI_FORMULA = {
    'HT-P2R-01': ('Roturas PT ↓ (% líneas con agotado)', '(valor-obj)/100*8000000*0.25'),
    'HT-P2R-03': ('% Saturación ↑ (horas de turno pagadas sin carga)', '(obj-valor)/100*6000*150'),
    'HT-P2R-05': ('OTIF ↑ (contra fecha pedida)', '(obj-valor)/100*8000000*0.25'),
    'HT-P2R-06': ('Paradas por falta de material ↓ (h/mes)', '(valor-obj)*12*150'),
    'HT-P2R-07': ('Días de stock MMPP y ST ↓', '(valor-obj)*8000*0.08'),
    'HT-P2R-08': ('Nº cambios / producción ↓', '(valor-obj)*20000*1.5*150'),
    'HT-P2R-09': ('Cambios en el programa cerrado ↓ (nº/semana)', '(valor-obj)*52*1.5*150'),
    'HT-P2R-10': ('% órdenes liberadas sin confirmación ↓ (en DICE pone «Cambios en el programa cerrado»)',
                  '(valor-obj)/100*3000*0.5*150'),
    'HT-P2R-11': ('OTIF ↑ (contra fecha confirmada)', '(obj-valor)/100*8000000*0.25'),
    'HT-P2R-12': ('Tiempo persona ↓ (h/año)', '(valor-obj)*35'),
    'HT-P2R-13': ('Días de stock PT ↓', '(valor-obj)*20000*0.08'),
    'HT-P2R-14': ('Tiempo persona ↓ (h/año)', '(valor-obj)*35'),
    'HT-P2R-15': ('Tiempo persona ↓ (h/año)', '(valor-obj)*35'),
}


def construir(salida):
    wb = Workbook()
    wf = wb.active
    wf.title = 'Fases'
    wper, wded, wh, wp = (wb.create_sheet(n) for n in ('Personas', 'Dedicación', 'Hallazgos', 'Palancas'))
    f_ult, per_ult, ded_ult = 4 + len(FASES), 4 + len(PERSONAS), 4 + len(DEDICACION)
    h_ult, p_ult = 4 + len(HALLAZGOS), 4 + len(PALANCAS)
    rng = lambda hoja, col, ult: f"'{hoja}'!${col}$5:${col}${ult}" if ' ' in hoja else f'{hoja}!${col}$5:${col}${ult}'

    # --- Fases ---
    titulo(wf, 'Esqueleto estándar · P2R · Plan to Release',
           f'Transcrito de la aplicación el {FECHA}. Repositorio: sin repositorio · 7 palancas lo usan · 0 con uno propio. '
           'Variantes de P2R: 0 (todas las celdas usan el esqueleto estándar).')
    cabecera(wf, 4, ['Orden', 'Fase', 'Tramo', 'Descripción (de → a)', 'Opcional', 'Áreas de negocio', 'Actividades',
                     'Hallazgos del catálogo', 'Palancas del catálogo'], [8, 40, 17, 62, 10, 16, 12, 14, 14])
    for k, (orden, fase, tramo, desc, opc, areas, acts) in enumerate(FASES):
        r = 5 + k
        vals = [orden, fase, tramo, desc, opc, areas, acts,
                f'=COUNTIF({rng("Hallazgos", "C", h_ult)},B{r})', f'=COUNTIF({rng("Palancas", "F", p_ult)},B{r})']
        for j, v in enumerate(vals, 1):
            celda(wf, r, j, v, centrada=j in (1, 3, 5, 6, 7, 8, 9))
    t = f_ult + 1
    celda(wf, t, 1, 'Total', bold=True)
    for col in 'GHI':
        celda(wf, t, ' ABCDEFGHI'.index(col), f'=SUM({col}5:{col}{f_ult})', centrada=True, bold=True)
    wf.freeze_panes = 'A5'

    # --- Personas ---
    titulo(wper, 'Personas por empresa',
           f'Transcrito de la aplicación el {FECHA} (área de negocio Packaging, 14 personas). Coste anual estimado. '
           '«Σ dedicación» es el dato de la aplicación; «Dedicación detallada» suma la hoja Dedicación.')
    cab = ['Persona', 'ID', 'Email', 'Categoría', 'Roles', 'Censo en', 'Área de negocio', 'Jornada h/año', 'Coste anual (k€)',
           'Σ Dedicación', 'Coste imputado (k€)', 'Dedicación detallada', 'Sin detallar', 'Licencias']
    cabecera(wper, 4, cab, [30, 9, 32, 26, 44, 12, 15, 11, 12, 11, 13, 12, 11, 18])
    for k, p in enumerate(PERSONAS):
        r = 5 + k
        nom, pid, mail, cat, roles, censo, area, jornada, coste, ded, lic = p
        vals = [nom, pid, mail, cat, roles, censo, area, jornada, coste, ded, f'=I{r}*J{r}',
                f'=SUMIF({rng("Dedicación", "A", ded_ult)},A{r},{rng("Dedicación", "E", ded_ult)})', f'=J{r}-L{r}', lic or None]
        for j, v in enumerate(vals, 1):
            fmt = '#,##0' if j == 8 else '#,##0.0' if j in (9, 11) else '0%' if j in (10, 12, 13) else None
            celda(wper, r, j, v, centrada=j in (2, 6, 8, 9, 10, 11, 12, 13), fmt=fmt)
    t = per_ult + 1
    celda(wper, t, 1, f'Σ {len(PERSONAS)} personas', bold=True)
    for col, fmt in (('H', '#,##0'), ('I', '#,##0.0'), ('K', '#,##0.0')):
        celda(wper, t, ' ABCDEFGHIJKLMN'.index(col), f'=SUM({col}5:{col}{per_ult})', centrada=True, fmt=fmt, bold=True)
    wper.conditional_formatting.add(f'M5:M{per_ult}', FormulaRule(formula=['ABS($M5)>0.001'], fill=fill(AVISO),
                                                                   font=Font(name=F, bold=True, color='B42318')))
    wper.freeze_panes = 'B5'
    wper.auto_filter.ref = f'A4:N{per_ult}'

    # --- Dedicación ---
    titulo(wded, 'Dedicación por celda y fase',
           'Una fila por persona, celda (proceso × empresa) y fase. Por ahora solo está el detalle de P2R × Ondupet.')
    cabecera(wded, 4, ['Persona', 'Censo en', 'Celda', 'Fase', '% dedicación', 'Coste imputado (k€)'], [30, 12, 18, 40, 13, 16])
    for k, (nom, celda_, fase, pct) in enumerate(DEDICACION):
        r = 5 + k
        vals = [nom, f'=INDEX({rng("Personas", "F", per_ult)},MATCH(A{r},{rng("Personas", "A", per_ult)},0))', celda_, fase, pct,
                f'=E{r}*INDEX({rng("Personas", "I", per_ult)},MATCH(A{r},{rng("Personas", "A", per_ult)},0))']
        for j, v in enumerate(vals, 1):
            celda(wded, r, j, v, centrada=j in (2, 3, 5, 6), fmt='0%' if j == 5 else '#,##0.0' if j == 6 else None)
    wded.freeze_panes = 'A5'
    wded.auto_filter.ref = f'A4:F{ded_ult}'
    r0 = ded_ult + 3
    cabecera(wded, r0, ['Resumen por celda', 'Personas', 'Coste imputado (k€)'], [30, 12, 18])
    for k, c in enumerate(CELDAS):
        r = r0 + 1 + k
        celda(wded, r, 1, c)
        celda(wded, r, 2, f'=SUMPRODUCT(({rng("Dedicación", "C", ded_ult)}=A{r})/'
                          f'COUNTIFS({rng("Dedicación", "A", ded_ult)},{rng("Dedicación", "A", ded_ult)},'
                          f'{rng("Dedicación", "C", ded_ult)},{rng("Dedicación", "C", ded_ult)}))', centrada=True, fmt='0')
        celda(wded, r, 3, f'=SUMIF({rng("Dedicación", "C", ded_ult)},A{r},{rng("Dedicación", "F", ded_ult)})', centrada=True,
              fmt='#,##0.0')

    # --- Hallazgos ---
    titulo(wh, 'Catálogo de tipos de hallazgo · P2R',
           f'Transcrito de la aplicación el {FECHA} (13 tipos). Descripción, métrica, KPI y fórmula del maestro; en la '
           'fórmula, valor y obj son los tokens de la aplicación y las constantes son de ejemplo. '
           '«Empresas» se calcula con las columnas de cada empresa.')
    cab = ['ID', 'Nombre', 'Fase', 'Descripción', 'Familia', 'Naturaleza', 'Métrica (volumen expuesto)', 'KPI sugerido',
           'Fórmula sugerida (€/año)', 'Se cura con', *EMPRESAS, 'Empresas', 'Caza inversa']
    cabecera(wh, 4, cab, [12, 30, 34, 60, 13, 15, 36, 34, 26, 18, 10, 10, 10, 10, 13])
    for k, (hid, nom, fase, desc, fam, nat, met, cura, caza, apl) in enumerate(HALLAZGOS):
        r = 5 + k
        kpi, formula = KPI_FORMULA[hid]
        vals = [hid, nom, fase, desc, fam, nat, met, kpi, formula, cura, *[a or None for a in apl],
                f'=COUNTIF(K{r}:M{r},"{S}")', caza]
        for j, v in enumerate(vals, 1):
            c = celda(wh, r, j, v, centrada=j in (1, 5, 6, 11, 12, 13, 14, 15))
            if j == 9:
                c.font = Font(name='Consolas', size=10, color='00263A')
    wh.conditional_formatting.add(f'A5:O{h_ult}', FormulaRule(formula=['$E5="Oportunidad"'], fill=fill('FFF1D6')))
    wh.freeze_panes = 'C5'
    wh.auto_filter.ref = f'A4:O{h_ult}'

    # --- Palancas ---
    titulo(wp, 'Catálogo de palancas · P2R',
           f'Transcrito de la aplicación el {FECHA} (7 palancas). €/año de referencia: vacío (el € sale de cada empresa) · '
           'confianza: por cuantificar. «Qué cambia», KPI y base expuesta del maestro.')
    cab = ['Código', 'Palanca', 'Qué cambia', 'KPI', 'Base expuesta', 'Fase as-is', '€/año ref.', 'Confianza', 'Cura tipos',
           *EMPRESAS, 'Aplicada en']
    cabecera(wp, 4, cab, [10, 32, 58, 30, 52, 34, 10, 15, 22, 10, 10, 10, 11])
    for k, (cod, nom, que, base, fase, cura, kpi, apl) in enumerate(PALANCAS):
        r = 5 + k
        vals = [cod, nom, que, kpi, base, fase, '—', 'Por cuantificar', cura, *[a or None for a in apl],
                f'=COUNTIF(J{r}:L{r},"{S}")']
        for j, v in enumerate(vals, 1):
            celda(wp, r, j, v, centrada=j in (1, 7, 8, 10, 11, 12, 13))
    wp.freeze_panes = 'C5'
    wp.auto_filter.ref = f'A4:M{p_ult}'

    palancas_cat = {p[0]: (p[1], p[2]) for p in PALANCAS}
    for empresa in ('Ondupet', 'MGT', 'Ondupack Almendralejo'):
        hoja_empresa(wb, empresa, palancas_cat)
    hoja_fichas(wb)
    hojas_fichas(wb)
    hoja_kpis(wb)

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    wb.save(salida)


if __name__ == '__main__':
    out = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).resolve().parents[1] / 'P2R_Consolidado.xlsx')
    construir(out)
    print(out)
