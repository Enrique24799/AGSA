#!/usr/bin/env python3
"""Coste de los KPIs de Plan to Release de Ondupet (situación ideal), una hoja por KPI, con fórmulas.

Lee un JSON con los datos ya agregados de los Excel de ahorros (paradas, producciones, stock, agotados y servicio) y
genera el libro. Los supuestos (margen €/t, coste de transformación, coberturas objetivo, bandas de servicio...) están
en la hoja Parámetros: cambiar una celda recalcula todo.

Uso: python3 generar_costes_kpis_ondupet.py <datos.json> <salida.xlsx>
Los datos llevan clientes y volúmenes: no se guardan en el repositorio.
"""
import json
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

F = 'Arial'
NAVY, AZUL, HIELO, GRIS = '00263A', '98C9EB', 'EAF3FA', '5B6B78'
AMARILLO = 'FFF2CC'
THIN = Side(style='thin', color='D2DCE5')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
EUR, TN, H, PCT, ENT = '#,##0 "€"', '#,##0.0', '#,##0', '0.0%', '#,##0'


def lista_si_no(ws):
    dv = DataValidation(type='list', formula1='"Sí,No"', allow_blank=False)
    ws.add_data_validation(dv)
    return dv


def fill(c):
    return PatternFill('solid', fgColor=c)


def titulo(ws, texto, sub):
    ws['A1'] = texto
    ws['A1'].font = Font(name=F, size=14, bold=True, color=NAVY)
    ws['A2'] = sub
    ws['A2'].font = Font(name=F, size=9, italic=True, color=GRIS)
    ws.sheet_view.showGridLines = False


def cabecera(ws, fila, col, textos, anchos=None):
    for j, t in enumerate(textos):
        c = ws.cell(fila, col + j, t)
        c.font = Font(name=F, size=9, bold=True, color='FFFFFF')
        c.fill = fill(NAVY)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = BOX
        if anchos:
            ws.column_dimensions[L(col + j)].width = anchos[j]
    ws.row_dimensions[fila].height = 30


def celda(ws, fila, col, v, fmt=None, bold=False, entrada=False, wrap=False, centro=False, color=None):
    c = ws.cell(fila, col, v)
    c.font = Font(name=F, size=9, bold=bold, color='0000FF' if entrada else (color or '000000'))
    if entrada:
        c.fill = fill(AMARILLO)
    if fmt:
        c.number_format = fmt
    c.alignment = Alignment(wrap_text=wrap, vertical='center', horizontal='center' if centro else None)
    c.border = BOX
    return c


def criterio(ws, fila, ncols, texto, alto):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
    c = ws.cell(fila, 1, texto)
    c.font = Font(name=F, size=9, color=NAVY)
    c.fill = fill(HIELO)
    c.alignment = Alignment(wrap_text=True, vertical='top')
    ws.row_dimensions[fila].height = alto


def nombre(wb, n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


def subtitulo(ws, fila, texto, col=1):
    c = ws.cell(fila, col, texto)
    c.font = Font(name=F, size=10, bold=True, color=NAVY)


# ---------------------------------------------------------------- parámetros
PARAMS = [
    ('Margen', 'Margen de contribución', 1000, '€/t', EUR,
     'Valor de ejemplo: pon el vuestro. Valora las ventas perdidas (agotados y servicio) y la alternativa de paradas y '
     'cambios.'),
    ('CosteTransf', 'Coste de transformación', 350, '€/t', EUR,
     'En tu Excel aparecen 350 €/t (paradas) y 423 €/t (cambios): confirma cuál.'),
    ('PctFijo', 'Parte fija del coste de transformación', 1, '%', PCT,
     'Lo que se paga aunque el termo esté parado (personal, amortización, estructura). Quita la energía si va incluida.'),
    ('CostePT', 'Coste del producto terminado en stock', 1800, '€/t', EUR,
     'Valor de ejemplo: material + transformación. Valora el exceso de stock.'),
    ('CosteCapital', 'Coste de capital', 0.08, '%/año', PCT, 'Coste financiero de tener el stock inmovilizado.'),
    ('CosteAlmacen', 'Almacenaje y obsolescencia', 0.07, '%/año', PCT, 'Espacio, manipulación y riesgo de no venderlo.'),
    ('IdealCambios', 'Situación ideal de cambios', 'Mínimo mensual', 'lista', None,
     '«Mínimo mensual»: cada bloque, altura y calidad se fabrica una sola vez al mes en cada termo. «Cero»: cuentan todos '
     'los cambios.'),
    ('HBloque', 'Duración de un cambio de bloque', 0.96, 'h', '0.00',
     'Media de los partes: 970 h de cambio de molde ÷ 1.013 cambios. Tu supuesto anterior: 4 h.'),
    ('HAltura', 'Duración de un cambio de altura', 0.64, 'h', '0.00',
     'Media de los partes: 647 h de cambio de insertos ÷ 1.006 cambios. Tu supuesto anterior: 1 h.'),
    ('HCalidad', 'Duración de un cambio de calidad', 0.30, 'h', '0.00',
     'Media de los partes: 100 h de cambio de saca por calidad ÷ 336 cambios. Tu supuesto anterior: 1 h.'),
    ('Tolerancia', 'Tolerancia de retraso', 0, 'días',
     '0', 'Un albarán cuenta como tarde si sale más de estos días después de la fecha de entrega.'),
]
OBJ_STOCK = [('MTS', 'A', 15), ('MTS', 'B', 20), ('MTS', 'C', 30), ('MTO', 'A', 7), ('MTO', 'B', 7), ('MTO', 'C', 7),
             ('-', '-', 30)]
BANDAS = [(0, 0.10), (0.70, 0.05), (0.80, 0.03), (0.90, 0.01), (0.95, 0)]


def hoja_parametros(wb):
    ws = wb.create_sheet('Parámetros')
    titulo(ws, 'Parámetros', 'Las celdas amarillas son supuestos: cámbialos y se recalcula todo el libro.')
    cabecera(ws, 4, 1, ['Parámetro', 'Valor', 'Unidad', 'Nota'], [38, 16, 10, 95])
    for i, (n, txt, v, uni, fmt, nota) in enumerate(PARAMS):
        r = 5 + i
        celda(ws, r, 1, txt, bold=True)
        celda(ws, r, 2, v, fmt, entrada=True, centro=True)
        celda(ws, r, 3, uni, centro=True)
        celda(ws, r, 4, nota, wrap=True)
        nombre(wb, n, f"'Parámetros'!$B${r}")
    dv = DataValidation(type='list', formula1='"Mínimo mensual,Cero"')
    ws.add_data_validation(dv)
    dv.add(f'B{5 + [p[0] for p in PARAMS].index("IdealCambios")}')

    r0 = 5 + len(PARAMS) + 2
    subtitulo(ws, r0, 'Cobertura objetivo del stock de producto terminado (días)')
    cabecera(ws, r0 + 1, 1, ['Clave', 'Días objetivo', 'Estrategia', 'Nota'])
    for i, (e, abc, dias) in enumerate(OBJ_STOCK):
        r = r0 + 2 + i
        celda(ws, r, 1, f'{e}|{abc}')
        celda(ws, r, 2, dias, ENT, entrada=True, centro=True)
        celda(ws, r, 3, f'{e} {abc}' if e != '-' else 'Sin clasificar', centro=True)
        celda(ws, r, 4, 'MTO: lo fabricado contra pedido debería salir en pocos días.' if e == 'MTO' and abc == 'A'
              else '', wrap=True)
    a, b = r0 + 2, r0 + 1 + len(OBJ_STOCK)
    nombre(wb, 'ObjClave', f"'Parámetros'!$A${a}:$A${b}")
    nombre(wb, 'ObjDias', f"'Parámetros'!$B${a}:$B${b}")

    r1 = b + 3
    subtitulo(ws, r1, 'Venta futura que se pierde según el nivel de servicio del cliente (kg a tiempo)')
    cabecera(ws, r1 + 1, 1, ['Servicio desde', '% de su venta que se pierde', '', 'Nota'])
    for i, (desde, perd) in enumerate(BANDAS):
        r = r1 + 2 + i
        celda(ws, r, 1, desde, PCT, entrada=True, centro=True)
        celda(ws, r, 2, perd, PCT, entrada=True, centro=True)
        hasta = BANDAS[i + 1][0] if i + 1 < len(BANDAS) else None
        celda(ws, r, 4, f'Clientes con servicio entre {desde:.0%} y {hasta:.0%}' if hasta is not None
              else f'Clientes con servicio de {desde:.0%} o más', wrap=True)
    a, b = r1 + 2, r1 + 1 + len(BANDAS)
    nombre(wb, 'BandaDesde', f"'Parámetros'!$A${a}:$A${b}")
    nombre(wb, 'Bandas', f"'Parámetros'!$A${a}:$B${b}")

    r2 = b + 2
    criterio(ws, r2, 4, 'Situación ideal = planificación ideal: 0 paradas por falta de material, cambios mínimos, stock en '
             'su cobertura objetivo, 0 ventas perdidas por planificación y 100 % de los kg a tiempo. El valor de cada '
             'hoja es el coste del hallazgo, no el ahorro: el ahorro saldrá de aplicar a cada palanca su % de mejora.', 42)
    ws.freeze_panes = 'A5'


# ---------------------------------------------------------------- paradas
CLASES_PARADA = [
    ('ESPERA POR FALTA DE BOBINAS FABRICADAS', 'Sí', 'Bobinas de extrusión propias: sincronización o capacidad de la '
     'extrusión. Se concentra entre junio y octubre de 2026.'),
    ('FALTAN CAJAS PARA EMBALAR', 'Sí', 'Material auxiliar comprado.'),
    ('FALTA ABS', 'Sí', 'Materia prima comprada.'),
    ('OP-FALTA ABASTECIMIENTO CARRETILLERO', 'No', 'Aprovisionamiento a línea (liberación), no falta de material. Ponlo '
     'en «Sí» si quieres contarlo aquí.'),
    ('PP-ESPERA CAMBIO DE MOLDE//PEDIDO FALTA PERSONAL', 'No', 'Es personal en el cambio: se cuenta en la hoja Cambios.'),
]


def hoja_paradas(wb, d):
    ws = wb.create_sheet('Paradas')
    termos = d['termos']
    titulo(ws, 'Paradas por falta de material', f'Termos 1 a 6 · {d["periodo"]} · horas de paro de la hoja Paradas')
    criterio(ws, 3, 10,
             'Qué se valora: las horas en que un termo está parado esperando material. Situación ideal: 0 h.\n'
             'Criterio: son horas de máquina pagadas que no producen. Se valora la parte fija del coste de transformación '
             'que corresponde a ese tiempo: coste de transformación × parte fija × t fabricadas × horas perdidas ÷ '
             '(horas de marcha + horas perdidas). Las horas de paro están sumadas termo a termo, así que la productividad '
             'es la de un termo (t ÷ horas de marcha), no la de los seis juntos.\n'
             'Alternativa (informativa, no se suma): si los termos estuvieran saturados, esas toneladas se habrían vendido '
             '→ t perdidas × margen.', 78)
    fila = 5
    subtitulo(ws, fila, 'Horas de paro por clase')
    cabecera(ws, fila + 1, 1, ['Clase de paro'] + [t.title() for t in termos] + ['Total h', '¿Cuenta?', 'Nota'],
             [46] + [9] * 6 + [10, 10, 70])
    dv = lista_si_no(ws)
    r0 = fila + 2
    for i, (cl, cuenta, nota) in enumerate(CLASES_PARADA):
        r = r0 + i
        celda(ws, r, 1, cl)
        for j, t in enumerate(termos):
            celda(ws, r, 2 + j, round(d['clases'][cl][t], 1), H)
        celda(ws, r, 8, f'=SUM(B{r}:G{r})', H, bold=True)
        celda(ws, r, 9, cuenta, entrada=True, centro=True)
        dv.add(f'I{r}')
        celda(ws, r, 10, nota, wrap=True)
        ws.row_dimensions[r].height = 26
    r1 = r0 + len(CLASES_PARADA) - 1
    rt = r1 + 1
    celda(ws, rt, 1, 'Horas que cuentan', bold=True)
    for j in range(6):
        col = L(2 + j)
        celda(ws, rt, 2 + j, f'=SUMPRODUCT({col}{r0}:{col}{r1}*($I${r0}:$I${r1}="Sí"))', H, bold=True)
    celda(ws, rt, 8, f'=SUM(B{rt}:G{rt})', H, bold=True)

    fila = rt + 3
    subtitulo(ws, fila, 'Marcha y producción por termo')
    cabecera(ws, fila + 1, 1, ['Termo', 'Horas de marcha', 't fabricadas', 't/h de marcha'])
    m0 = fila + 2
    for i, t in enumerate(termos):
        r = m0 + i
        celda(ws, r, 1, t.title())
        celda(ws, r, 2, round(d['marcha_h'][t], 1), H)
        celda(ws, r, 3, round(d['tn'][t], 1), TN)
        celda(ws, r, 4, f'=C{r}/B{r}', '0.000')
    m1 = m0 + len(termos) - 1
    mt = m1 + 1
    celda(ws, mt, 1, 'Total termos', bold=True)
    celda(ws, mt, 2, f'=SUM(B{m0}:B{m1})', H, bold=True)
    celda(ws, mt, 3, f'=SUM(C{m0}:C{m1})', TN, bold=True)
    celda(ws, mt, 4, f'=C{mt}/B{mt}', '0.000', bold=True)
    nombre(wb, 'HorasMarcha', f'Paradas!$B${mt}')
    nombre(wb, 'TnFabricadas', f'Paradas!$C${mt}')
    nombre(wb, 'Productividad', f'Paradas!$D${mt}')

    fila = mt + 3
    subtitulo(ws, fila, 'Cálculo')
    cabecera(ws, fila + 1, 1, ['Concepto', 'Valor', 'Unidad'])
    filas = [
        ('Horas perdidas que cuentan', f'=H{rt}', H, 'h'),
        ('Horas de marcha de los 6 termos', '=HorasMarcha', H, 'h'),
        ('t fabricadas', '=TnFabricadas', TN, 't'),
        ('Productividad de un termo', '=Productividad', '0.000', 't/h'),
        ('t que se dejaron de fabricar', None, TN, 't'),
        ('Coste de transformación de lo fabricado', '=TnFabricadas*CosteTransf', EUR, '€'),
        ('Coste del hallazgo', None, EUR, '€/año'),
        ('Alternativa con margen (informativa)', None, EUR, '€'),
    ]
    c0 = fila + 2
    for i, (txt, f, fmt, uni) in enumerate(filas):
        r = c0 + i
        if txt.startswith('t que'):
            f = f'=B{c0}*Productividad'
        if txt == 'Coste del hallazgo':
            f = f'=CosteTransf*PctFijo*TnFabricadas*B{c0}/(HorasMarcha+B{c0})'
        if txt.startswith('Alternativa'):
            f = f'=B{c0 + 4}*Margen'
        celda(ws, r, 1, txt, bold=txt == 'Coste del hallazgo')
        celda(ws, r, 2, f, fmt, bold=txt == 'Coste del hallazgo')
        celda(ws, r, 3, uni, centro=True)
    nombre(wb, 'ParadasHoras', f'Paradas!$B${c0}')
    nombre(wb, 'ParadasCoste', f'Paradas!$B${c0 + 6}')
    nombre(wb, 'ParadasAlt', f'Paradas!$B${c0 + 7}')

    fila = c0 + len(filas) + 2
    subtitulo(ws, fila, 'Horas por mes, todos los termos (de tu tabla dinámica, informativo)')
    meses = d['mensual_meses']
    clases = list(d['mensual'])
    cabecera(ws, fila + 1, 1, ['Mes'] + [c.title() for c in clases] + ['Total'])
    ws.row_dimensions[fila + 1].height = 62
    for i, m in enumerate(meses):
        r = fila + 2 + i
        celda(ws, r, 1, m.replace('_', '-'), centro=True)
        for j, cl in enumerate(clases):
            celda(ws, r, 2 + j, round(d['mensual'][cl][i], 1), H)
        celda(ws, r, 2 + len(clases), f'=SUM(B{r}:{L(1 + len(clases))}{r})', H, bold=True)
    ws.freeze_panes = 'A5'


# ---------------------------------------------------------------- cambios
ESPERAS_CAMBIO = [('PP-ESPERA CAMBIO DE MOLDE//PEDIDO FALTA PERSONAL', 'Sí'),
                  ('AV-HERRAMIENTA O MOLDE NO LISTO EN ZONA CAMBIO', 'Sí')]
REGISTRADAS = [('CAMBIO MOLDE', 'bloque'), ('CAMBIO MOLDE - NUEVA REFERENCIA', 'bloque'),
               ('CAMBIO DE INSERTOS (ALTURA DE MOLDE)', 'altura'), ('CAMBIO DE PEDIDO- CAMBIO DE SACA POR CALIDAD', 'calidad'),
               ('CAMBIO Nº DE PEDIDO//PADDING', 'otros'), ('PP-LIMPIEZA MOLINO CAMBIO', 'otros')]


def hoja_cambios(wb, d):
    ws = wb.create_sheet('Cambios')
    titulo(ws, 'Cambios de formato', f'Termos 1 a 6 · {d["periodo"]} · cambios de la hoja Producciones')
    criterio(ws, 3, 9,
             'Qué se valora: las horas de cambio (bloque, altura y calidad) por encima de las que necesitaría una '
             'planificación ideal, más las esperas ligadas al cambio. Situación ideal (parámetro): «Mínimo mensual» = cada '
             'bloque, altura y calidad se fabrica una sola vez al mes en cada termo; cambios mínimos = bloques distintos, '
             'alturas distintas − bloques distintos y calidades distintas. «Cero» = cuentan todos los cambios.\n'
             'Horas = cambios de más × duración media de cada tipo (parámetros, sacada de los partes).\n'
             'Criterio de €: el mismo que en paradas (parte fija del coste de transformación del tiempo perdido). '
             'Alternativa informativa: t que se dejan de fabricar × margen.', 78)
    filas_d = sorted(d['cambios'].items(), key=lambda kv: (kv[0].split('|')[0], kv[0].split('|')[1]))
    # detalle a la derecha del resumen: columnas K en adelante
    dc = 11
    subtitulo(ws, 5, 'Detalle por termo y mes', col=dc)
    cab = ['Termo', 'Mes', 't fabricadas', 'Cambios de bloque', 'Cambios de altura', 'Cambios de calidad',
           'Bloques distintos', 'Alturas distintas', 'Calidades distintas']
    cabecera(ws, 6, dc, cab, [10, 9, 11, 10, 10, 10, 10, 10, 10])
    d0 = 7
    for i, (k, v) in enumerate(filas_d):
        r = d0 + i
        maq, mes = k.split('|')
        vals = [maq.title(), mes.replace('_', '-'), round(v['kg'] / 1000, 1), v['cb'], v['ca'], v['cc'], v['nbl'],
                v['ntr'], v['ncal']]
        for j, x in enumerate(vals):
            celda(ws, r, dc + j, x, TN if j == 2 else (ENT if j > 2 else None), centro=j < 2)
    d1 = d0 + len(filas_d) - 1
    col = lambda j: f'${L(dc + j)}${d0}:${L(dc + j)}${d1}'

    subtitulo(ws, 5, 'Cambios de más frente al ideal')
    cabecera(ws, 6, 1, ['Tipo de cambio', 'Cambios reales', 'Cambios ideales', 'Cambios de más', 'h por cambio', 'Horas'],
             [34, 12, 12, 12, 11, 11])
    tipos = [('Bloque', 3, f'SUM({col(6)})', 'HBloque'),
             ('Altura', 4, f'SUM({col(7)})-SUM({col(6)})', 'HAltura'),
             ('Calidad', 5, f'SUM({col(8)})', 'HCalidad')]
    for i, (t, jr, ideal, hn) in enumerate(tipos):
        r = 7 + i
        celda(ws, r, 1, t, bold=True)
        celda(ws, r, 2, f'=SUM({col(jr)})', ENT)
        celda(ws, r, 3, f'=IF(IdealCambios="Cero",0,{ideal})', ENT)
        celda(ws, r, 4, f'=MAX(0,B{r}-C{r})', ENT)
        celda(ws, r, 5, f'={hn}', '0.00')
        celda(ws, r, 6, f'=D{r}*E{r}', H)
    dv = lista_si_no(ws)
    subtitulo(ws, 11, 'Esperas ligadas al cambio (ideal: 0 h)')
    cabecera(ws, 12, 1, ['Clase de paro', 'Horas', '¿Cuenta?'])
    for i, (cl, cuenta) in enumerate(ESPERAS_CAMBIO):
        r = 13 + i
        celda(ws, r, 1, cl, wrap=True)
        celda(ws, r, 2, round(sum(d['clases'][cl].values()), 1), H)
        celda(ws, r, 3, cuenta, entrada=True, centro=True)
        dv.add(f'C{r}')
        ws.row_dimensions[r].height = 24
    subtitulo(ws, 16, 'Cálculo')
    cabecera(ws, 17, 1, ['Concepto', 'Valor', 'Unidad'])
    calc = [('Horas de cambio de más', '=SUM(F7:F9)', H, 'h'),
            ('Horas de espera ligadas al cambio', '=SUMPRODUCT(B13:B14*(C13:C14="Sí"))', H, 'h'),
            ('Horas perdidas', '=B18+B19', H, 'h'),
            ('t que se dejaron de fabricar', '=B20*Productividad', TN, 't'),
            ('Coste del hallazgo', '=CosteTransf*PctFijo*TnFabricadas*B20/(HorasMarcha+B20)', EUR, '€/año'),
            ('Alternativa con margen (informativa)', '=B21*Margen', EUR, '€')]
    for i, (txt, f, fmt, uni) in enumerate(calc):
        r = 18 + i
        celda(ws, r, 1, txt, bold=txt == 'Coste del hallazgo')
        celda(ws, r, 2, f, fmt, bold=txt == 'Coste del hallazgo')
        celda(ws, r, 3, uni, centro=True)
    nombre(wb, 'CambiosHoras', 'Cambios!$B$20')
    nombre(wb, 'CambiosCoste', 'Cambios!$B$22')
    nombre(wb, 'CambiosAlt', 'Cambios!$B$23')

    subtitulo(ws, 26, 'De dónde sale la duración de cada cambio (horas registradas en los partes)')
    cabecera(ws, 27, 1, ['Clase de paro', 'Horas', 'Tipo', 'h por cambio real'])
    reales = {'bloque': 'B7', 'altura': 'B8', 'calidad': 'B9'}
    for i, (cl, tipo) in enumerate(REGISTRADAS):
        r = 28 + i
        celda(ws, r, 1, cl, wrap=True)
        celda(ws, r, 2, round(sum(d['clases'][cl].values()), 1), H)
        celda(ws, r, 3, tipo, centro=True)
        celda(ws, r, 4, f'=B{r}/{reales[tipo]}' if tipo in reales else '—', '0.00', centro=tipo not in reales)
    r = 28 + len(REGISTRADAS)
    celda(ws, r, 1, 'Total registrado', bold=True)
    celda(ws, r, 2, f'=SUM(B28:B{r - 1})', H, bold=True)
    ws.freeze_panes = 'A5'


# ---------------------------------------------------------------- stock
def hoja_stock(wb, d):
    st = d['stock']
    wd = wb.create_sheet('Datos stock')
    titulo(wd, 'Datos de stock', f'Media diaria de las fotos del {st["desde"]} al {st["hasta"]} ({st["n"]} días), '
                                 'producto terminado con peso en el maestro')
    cab = ['Tipo', 'Estrategia', 'ABC', 'Días de cobertura', 'kg medio', 'kg ÷ días (medio)', 'Días objetivo',
           'Exceso kg']
    cabecera(wd, 4, 1, cab, [10, 11, 7, 12, 12, 14, 12, 12])
    filas = sorted((k.split('|'), v) for k, v in st['agg'].items())
    for i, ((tipo, es, abc, dia), (kg, kgd, _)) in enumerate(filas):
        r = 5 + i
        for j, x in enumerate([tipo, es, abc, int(dia), round(kg, 1), round(kgd, 4)]):
            celda(wd, r, 1 + j, x, None if j < 3 else (ENT if j == 3 else '#,##0.0'), centro=j < 4)
        celda(wd, r, 7, f'=IFERROR(INDEX(ObjDias,MATCH(B{r}&"|"&C{r},ObjClave,0)),INDEX(ObjDias,MATCH("-|-",ObjClave,0)))', ENT)
        celda(wd, r, 8, f'=IF(D{r}>G{r},E{r}-G{r}*F{r},0)', '#,##0')
    n1 = 4 + len(filas)
    wd.freeze_panes = 'A5'
    rng = lambda c: f"'Datos stock'!${c}$5:${c}${n1}"

    ws = wb.create_sheet('Stock')
    titulo(ws, 'Días de stock · exceso sobre la cobertura objetivo',
           f'Producto terminado · fotos diarias del {st["desde"]} al {st["hasta"]} (no hay datos anteriores)')
    criterio(ws, 3, 9,
             'Qué se valora: el stock de producto terminado por encima de la cobertura objetivo de su segmento. Situación '
             'ideal: cada referencia en su cobertura objetivo (parámetros); MTO casi sin stock.\n'
             'Exceso de una referencia en un día = stock − consumo diario × días objetivo, si su cobertura supera el '
             'objetivo. Se hace la media de todos los días del periodo. Cobertura = «Días de stock sin entrada ponderado» '
             '(stock actual ÷ consumo).\n'
             'Criterio de €: coste anual de mantener el exceso = exceso (t) × coste del producto terminado (€/t) × (coste '
             'de capital + almacenaje y obsolescencia).', 78)
    cab = ['Tipo', 'Estrategia', 'ABC', 'Días objetivo', 'Stock medio (t)', 'Stock con más de 90 días (t)',
           'Exceso medio (t)', 'Valor del exceso (€)', 'Coste anual (€)']
    cabecera(ws, 5, 1, cab, [10, 11, 7, 11, 13, 15, 13, 15, 14])
    segs = sorted({(t, e, a) for (t, e, a, _), _ in filas}, key=lambda x: (x[0] != 'Cesta', x[1] == '-', x[1], x[2]))
    for i, (t, e, a) in enumerate(segs):
        r = 6 + i
        crit = f'{rng("A")},"{t}",{rng("B")},"{e}",{rng("C")},"{a}"'
        celda(ws, r, 1, t)
        celda(ws, r, 2, e, centro=True)
        celda(ws, r, 3, a, centro=True)
        celda(ws, r, 4, f'=IFERROR(INDEX(ObjDias,MATCH(B{r}&"|"&C{r},ObjClave,0)),INDEX(ObjDias,MATCH("-|-",ObjClave,0)))', ENT)
        celda(ws, r, 5, f'=SUMIFS({rng("E")},{crit})/1000', TN)
        celda(ws, r, 6, f'=SUMIFS({rng("E")},{crit},{rng("D")},">90")/1000', TN)
        celda(ws, r, 7, f'=SUMIFS({rng("H")},{crit})/1000', TN)
        celda(ws, r, 8, f'=G{r}*CostePT', EUR)
        celda(ws, r, 9, f'=H{r}*(CosteCapital+CosteAlmacen)', EUR)
    r1 = 5 + len(segs)
    rt = r1 + 1
    celda(ws, rt, 1, 'Total', bold=True)
    for c in 'EFGHI':
        celda(ws, rt, ' ABCDEFGHI'.index(c), f'=SUM({c}6:{c}{r1})', EUR if c in 'HI' else TN, bold=True)
    nombre(wb, 'StockExceso', f'Stock!$G${rt}')
    nombre(wb, 'StockCoste', f'Stock!$I${rt}')
    cero = sum(st['cero_kg'].values()) / 1000
    sinp = st['sin_peso']
    mil = lambda x: f'{x:,.0f}'.replace(',', '.')
    notas = [f'No se cuenta el stock con cobertura 0 (≈ {mil(cero)} t de media): tiene existencias pero no consumo '
             'calculado, así que no se puede saber si sobra.',
             f'Bobinas sin peso en el maestro: ≈ {mil(sinp.get("Bobina", 0))} unidades de media no incluidas; si esas '
             f'unidades fueran kg, serían ≈ {mil(sinp.get("Bobina", 0) / 1000)} t.',
             'Materias primas y auxiliares no están incluidas: el maestro no tiene su peso ni su coste.']
    for i, n in enumerate(notas):
        c = ws.cell(rt + 2 + i, 1, '• ' + n)
        c.font = Font(name=F, size=9, italic=True, color=GRIS)
    ws.freeze_panes = 'A6'


# ---------------------------------------------------------------- agotados
DESTINOS = ['Rotura de stock', 'Falta de capacidad', 'Falta de materia prima', 'Fallo en la entrega',
            'Excluido (no es planificación)']
MAPA_CAUSAS = [('Sin stock de producto terminado', 'Rotura de stock'),
               ('Fecha de entrega ofrecida', 'Según estrategia: MTS → rotura de stock · MTO → falta de capacidad'),
               ('Falta de capacidad', 'Falta de capacidad'),
               ('Falta de materia prima (ABS)', 'Falta de materia prima'),
               ('Fallo en la entrega', 'Fallo en la entrega'),
               ('Fuera de porfolio', 'Excluido (no es planificación)'),
               ('Precio o competencia', 'Excluido (no es planificación)'),
               ('Pedido mínimo', 'Excluido (no es planificación)'),
               ('Prueba', 'Excluido (no es planificación)'),
               ('Sin comentario', 'Excluido (no es planificación)'),
               ('Otros', 'Excluido (no es planificación)')]


def hoja_agotados(wb, d):
    ws = wb.create_sheet('Agotados')
    ag = d['agotados']
    titulo(ws, 'Agotados · ventas perdidas en consulta', f'Registro de agotados · {d["periodo"]}')
    criterio(ws, 3, 13,
             'Qué se valora: las ventas que se pierden en la consulta del cliente (registro de agotados). Situación ideal: '
             '0 ventas perdidas por planificación.\n'
             'Criterio de €: t perdidas × margen. Reparto por causa: «Fecha de entrega ofrecida» va por estrategia (MTS → '
             'rotura de stock, porque tendría que haber stock; MTO → falta de capacidad, porque el plazo sale largo). El '
             'resto según el comentario, con la tabla de causas (editable). Se excluyen las causas que no son de '
             'planificación: porfolio, precio y pedido mínimo.\n'
             'Peso: del maestro por referencia; si no está, la media del bloque; en bobinas la cantidad se toma como kg.',
             78)
    subtitulo(ws, 5, 'Resumen por destino')
    cabecera(ws, 6, 1, ['Destino', 'Registros', 'Unidades', 't', 'Coste (€)'], [34, 10, 13, 9, 13])
    n0, n1 = 22, 22 + len(ag) - 1
    rng = lambda c: f'${c}${n0}:${c}${n1}'
    for i, dest in enumerate(DESTINOS):
        r = 7 + i
        celda(ws, r, 1, dest, bold=True)
        celda(ws, r, 2, f'=COUNTIF({rng("L")},A{r})', ENT)
        celda(ws, r, 3, f'=SUMIF({rng("L")},A{r},{rng("E")})', ENT)
        celda(ws, r, 4, f'=SUMIF({rng("L")},A{r},{rng("H")})', TN)
        celda(ws, r, 5, f'=SUMIF({rng("L")},A{r},{rng("M")})', EUR)
        nombre(wb, f'Agot{i}T', f'Agotados!$D${r}')
        nombre(wb, f'Agot{i}E', f'Agotados!$E${r}')
    rt = 7 + len(DESTINOS)
    celda(ws, rt, 1, 'Total', bold=True)
    for c, fmt in zip('BCDE', (ENT, ENT, TN, EUR)):
        celda(ws, rt, ' ABCDE'.index(c), f'=SUM({c}7:{c}{rt - 1})', fmt, bold=True)

    subtitulo(ws, 5, 'Tabla de causas (editable)', col=7)
    cabecera(ws, 6, 7, ['Causa', 'Destino'])
    ws.merge_cells(start_row=6, start_column=8, end_row=6, end_column=10)
    for i, (ca, de) in enumerate(MAPA_CAUSAS):
        r = 7 + i
        celda(ws, r, 7, ca)
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=10)
        celda(ws, r, 8, de, entrada=ca != 'Fecha de entrega ofrecida')
    mapa = f'$G$7:$H${6 + len(MAPA_CAUSAS)}'

    subtitulo(ws, n0 - 2, 'Detalle de los registros')
    cab = ['Mes', 'Fecha de entrega', 'Cliente', 'Referencia', 'Unidades', 'kg por unidad', 'Fuente del peso', 't',
           'Estrategia', 'Comentario', 'Causa', 'Destino', 'Coste (€)']
    cabecera(ws, n0 - 1, 1, cab)
    for col, w in zip('ABCDEFGHIJKLM', [34, 10, 26, 34, 11, 9, 20, 8, 9, 50, 26, 26, 11]):
        ws.column_dimensions[col].width = max(ws.column_dimensions[col].width or 0, w)
    for i, a in enumerate(sorted(ag, key=lambda x: (x['mes'], x['fecha']))):
        r = n0 + i
        vals = [a['mes'].replace('_', '-'), a['fecha'], a['cliente'], a['ref'], a['uds'], a['peso'], a['fuente'], None,
                a['estr'] or '-', a['coment'], a['causa']]
        for j, x in enumerate(vals):
            celda(ws, r, 1 + j, x, {4: ENT, 5: '0.000'}.get(j), centro=j in (0, 1, 8))
        celda(ws, r, 8, f'=IFERROR(E{r}*F{r}/1000,0)', TN)
        celda(ws, r, 12, f'=IF(K{r}="Fecha de entrega ofrecida",IF(I{r}="MTS","Rotura de stock","Falta de capacidad"),'
                         f'IFERROR(VLOOKUP(K{r},{mapa},2,0),"Excluido (no es planificación)"))')
        celda(ws, r, 13, f'=IF(L{r}="Excluido (no es planificación)",0,H{r}*Margen)', EUR)
    ws.freeze_panes = 'A5'


# ---------------------------------------------------------------- servicio
def hoja_servicio(wb, d):
    sv = d['servicio']
    wd = wb.create_sheet('Datos servicio')
    titulo(wd, 'Datos de servicio', f'Albaranes · {d["periodo"]} · un albarán por fila')
    cab = ['Pedido-posición', 'Material', 'Código cliente', 'Cliente', 'Sector', 'Estrategia', 'Fecha de pedido',
           'Fecha de entrega', 'Fecha de albarán', 'Días de retraso', 'Tarde', 'Cantidad', 'kg', 'Nota']
    cabecera(wd, 4, 1, cab, [14, 40, 10, 34, 13, 10, 11, 11, 11, 10, 7, 11, 11, 22])
    for i, s in enumerate(sorted(sv, key=lambda x: (x['falb'] or '', x['pos'] or ''))):
        r = 5 + i
        fd = lambda x: date.fromisoformat(x) if x else None
        vals = [s['pos'], s['mat'], s['cod'], s['cliente'], s['sector'], s['estr'] or '-', fd(s['fped']), fd(s['fent']),
                fd(s['falb'])]
        for j, x in enumerate(vals):
            c = celda(wd, r, 1 + j, x, 'dd/mm/yyyy' if j >= 6 else None, centro=j in (2, 5))
        celda(wd, r, 10, f'=IF(AND(H{r}<>"",I{r}<>""),I{r}-H{r},0)', '0')
        celda(wd, r, 11, f'=IF(J{r}>Tolerancia,1,0)', '0', centro=True)
        celda(wd, r, 12, s['ctd'], ENT)
        celda(wd, r, 13, round(s['kg'], 1), '#,##0.0')
        celda(wd, r, 14, s['nota'])
    n1 = 4 + len(sv)
    wd.freeze_panes = 'A5'
    rng = lambda c: f"'Datos servicio'!${c}$5:${c}${n1}"

    ws = wb.create_sheet('Servicio')
    titulo(ws, 'Servicio · venta futura perdida', f'Albaranes · {d["periodo"]}')
    criterio(ws, 3, 11,
             'Qué se valora: la venta futura que se pierde porque el cliente, con el servicio que ha recibido, deja de '
             'pedir una parte de su volumen. No se valora lo retrasado (al final se expidió) ni los agotados (van en su '
             'hoja). Situación ideal: 100 % de los kg a tiempo.\n'
             'Criterio: para cada cliente, nivel de servicio = kg a tiempo ÷ kg expedidos (en kg y no en albaranes, porque '
             'los pedidos grandes fallan más); % de pérdida según la tabla de bandas (parámetros); t en riesgo = t '
             'expedidas del cliente × % de pérdida; € = t en riesgo × margen.\n'
             'Un albarán es «tarde» si sale más días después de la fecha de entrega que la tolerancia (parámetro). En '
             'bobinas sin peso se toma la cantidad como kg.', 78)
    subtitulo(ws, 5, 'Resumen')
    cabecera(ws, 6, 1, ['Concepto', 'Valor'], [44, 14])
    res = [('Albaranes', f'=COUNT({rng("J")})', ENT), ('Albaranes tarde', f'=SUM({rng("K")})', ENT),
           ('Servicio por albaranes', '=1-B8/B7', PCT),
           ('t expedidas', f'=SUM({rng("M")})/1000', TN),
           ('t tarde', f'=SUMIF({rng("K")},1,{rng("M")})/1000', TN),
           ('Servicio en kg', '=1-B11/B10', PCT),
           ('t de venta futura en riesgo', None, TN),
           ('Coste (€/año)', None, EUR)]
    for i, (txt, f, fmt) in enumerate(res):
        r = 7 + i
        celda(ws, r, 1, txt, bold=txt.startswith('Coste'))
        celda(ws, r, 2, f, fmt, bold=txt.startswith('Coste'))
    clientes = sorted({s['cod']: s for s in sv}.values(), key=lambda s: -sum(x['kg'] for x in sv if x['cod'] == s['cod']))
    c0 = 30
    c1 = c0 + len(clientes) - 1
    ws['B13'] = f'=SUM(J{c0}:J{c1})'
    ws['B14'] = f'=SUM(K{c0}:K{c1})'
    for ref, fmt in (('B13', TN), ('B14', EUR)):
        ws[ref].number_format = fmt
        ws[ref].font = Font(name=F, size=9, bold=ref == 'B14')
        ws[ref].border = BOX
    nombre(wb, 'ServT', 'Servicio!$B$13')
    nombre(wb, 'ServCoste', 'Servicio!$B$14')

    subtitulo(ws, 5, 'Por banda de servicio del cliente', col=4)
    cabecera(ws, 6, 4, ['Servicio desde', 'hasta', '% de pérdida', 'Clientes', 't expedidas', 't en riesgo', 'Coste (€)'],
             [11, 9, 11, 9, 12, 12, 13])
    nb = len(BANDAS)
    for i in range(nb):
        r = 7 + i
        celda(ws, r, 4, f'=INDEX(BandaDesde,{i + 1})', PCT, centro=True)
        celda(ws, r, 5, f'=INDEX(BandaDesde,{i + 2})' if i + 1 < nb else 1.0000001, PCT, centro=True)
        celda(ws, r, 6, f'=INDEX(Bandas,{i + 1},2)', PCT, centro=True)
        rg = lambda c: f'${c}${c0}:${c}${c1}'
        cond = f'{rg("H")},">="&D{r},{rg("H")},"<"&E{r}'
        celda(ws, r, 7, f'=COUNTIFS({cond})', ENT)
        celda(ws, r, 8, f'=SUMIFS({rg("F")},{cond})', TN)
        celda(ws, r, 9, f'=SUMIFS({rg("J")},{cond})', TN)
        celda(ws, r, 10, f'=SUMIFS({rg("K")},{cond})', EUR)

    subtitulo(ws, 14, 'Días de retraso de los albaranes tarde (informativo)', col=4)
    cabecera(ws, 15, 4, ['Días de retraso', '', '', 'Albaranes', 't'])
    ws.merge_cells(start_row=15, start_column=4, end_row=15, end_column=6)
    tramos = [('1 a 2', 1, 2), ('3 a 7', 3, 7), ('8 a 15', 8, 15), ('Más de 15', 16, 100000)]
    for i, (txt, a, b) in enumerate(tramos):
        r = 16 + i
        celda(ws, r, 4, txt, centro=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
        cond = f'{rng("J")},">="&{a},{rng("J")},"<="&{b},{rng("K")},1'
        celda(ws, r, 7, f'=COUNTIFS({cond})', ENT)
        celda(ws, r, 8, f'=SUMIFS({rng("M")},{cond})/1000', TN)

    subtitulo(ws, c0 - 2, 'Por cliente')
    cab = ['Código', 'Cliente', 'Sector', 'Albaranes', 'Albaranes tarde', 't expedidas', 't tarde', 'Servicio en kg',
           '% de pérdida', 't en riesgo', 'Coste (€)']
    cabecera(ws, c0 - 1, 1, cab)
    for col, w in zip('ABCDEFGHIJK', [44, 34, 13, 11, 12, 12, 9, 12, 11, 11, 13]):
        ws.column_dimensions[col].width = max(ws.column_dimensions[col].width or 0, w)
    for i, s in enumerate(clientes):
        r = c0 + i
        celda(ws, r, 1, s['cod'], centro=True)
        celda(ws, r, 2, s['cliente'])
        celda(ws, r, 3, s['sector'])
        celda(ws, r, 4, f'=COUNTIF({rng("C")},A{r})', ENT)
        celda(ws, r, 5, f'=SUMIF({rng("C")},A{r},{rng("K")})', ENT)
        celda(ws, r, 6, f'=SUMIF({rng("C")},A{r},{rng("M")})/1000', TN)
        celda(ws, r, 7, f'=SUMIFS({rng("M")},{rng("C")},A{r},{rng("K")},1)/1000', TN)
        celda(ws, r, 8, f'=IF(F{r}>0,1-G{r}/F{r},1)', PCT)
        celda(ws, r, 9, f'=VLOOKUP(H{r},Bandas,2,TRUE)', PCT)
        celda(ws, r, 10, f'=F{r}*I{r}', TN)
        celda(ws, r, 11, f'=J{r}*Margen', EUR)
    ws.freeze_panes = 'A5'


# ---------------------------------------------------------------- resumen
def hoja_resumen(wb):
    ws = wb.create_sheet('Resumen', 0)
    titulo(ws, 'Coste de los KPIs de Plan to Release · Ondupet',
           'Periodo 11/2025–10/2026 (stock: 26/03–31/10/2026) · situación ideal = planificación ideal · el valor es el '
           'coste del hallazgo, no el ahorro')
    cab = ['KPI', 'Qué se valora', 'Situación ideal', 'Criterio de valoración', 'Cantidad', 'Unidad', 'Coste (€/año)',
           'Alternativa con margen (€)', 'Hallazgos que comparten el KPI']
    cabecera(ws, 4, 1, cab, [30, 42, 24, 42, 11, 7, 14, 15, 40])
    filas = [
        ('Paradas por falta de material', 'Horas de termo paradas esperando bobinas, cajas o ABS', '0 h',
         'Parte fija del coste de transformación del tiempo perdido', '=ParadasHoras', 'h', '=ParadasCoste',
         '=ParadasAlt', 'HT-06 (cajas y ABS) · HT-03 y sincronización extrusión–termo (bobinas)'),
        ('Cambios', 'Horas de cambio por encima del ideal y esperas del cambio', '=IdealCambios',
         'Parte fija del coste de transformación del tiempo perdido', '=CambiosHoras', 'h', '=CambiosCoste', '=CambiosAlt',
         'HT-08 · HT-09'),
        ('Días de stock', 'Stock de producto terminado por encima de la cobertura objetivo',
         'Cobertura objetivo por segmento', 'Exceso × coste del PT × (capital + almacenaje)', '=StockExceso', 't',
         '=StockCoste', None, 'HT-13 · HT-04'),
        ('Agotados · rotura de stock', 'Ventas perdidas en consulta: MTS sin stock', '0 ventas perdidas',
         't × margen', '=Agot0T', 't', '=Agot0E', None, 'HT-01'),
        ('Agotados · falta de capacidad', 'Ventas perdidas en consulta: MTO con plazo largo o sin capacidad',
         '0 ventas perdidas', 't × margen', '=Agot1T', 't', '=Agot1E', None, 'HT-03'),
        ('Agotados · falta de materia prima', 'Ventas perdidas por falta de ABS', '0 ventas perdidas', 't × margen',
         '=Agot2T', 't', '=Agot2E', None, 'HT-06'),
        ('Agotados · fallo en la entrega', 'Ventas perdidas tras fallar una entrega', '0 ventas perdidas', 't × margen',
         '=Agot3T', 't', '=Agot3E', None, 'HT-11'),
        ('Servicio · venta futura perdida', 'Parte de las compras de cada cliente que deja de pedir por su nivel de '
         'servicio', '100 % de los kg a tiempo', 't en riesgo × margen', '=ServT', 't', '=ServCoste', None,
         'HT-05 · HT-11 · HT-09'),
    ]
    for i, fila in enumerate(filas):
        r = 5 + i
        for j, v in enumerate(fila):
            fmt = {4: '#,##0', 6: EUR, 7: EUR}.get(j)
            celda(ws, r, 1 + j, v, fmt, bold=j in (0, 6), wrap=j in (1, 2, 3, 8), centro=j == 5)
        ws.row_dimensions[r].height = 32
    rt = 5 + len(filas)
    celda(ws, rt, 1, 'Total orientativo', bold=True)
    celda(ws, rt, 7, f'=SUM(G5:G{rt - 1})', EUR, bold=True)
    c = ws.cell(rt + 2, 1, '• Total antes del reparto entre hallazgos: algunos KPI comparten causa (por ejemplo, la espera '
                           'de bobinas y los agotados por capacidad, o los cambios de plan con cambios y servicio).')
    c.font = Font(name=F, size=9, italic=True, color=GRIS)
    c = ws.cell(rt + 3, 1, '• La alternativa con margen de paradas y cambios no se suma: es el valor si esas toneladas se '
                           'hubieran vendido (máquina saturada), en lugar del coste de transformación.')
    c.font = Font(name=F, size=9, italic=True, color=GRIS)
    c = ws.cell(rt + 4, 1, '• Margen, costes y coberturas son supuestos de la hoja Parámetros (celdas amarillas).')
    c.font = Font(name=F, size=9, italic=True, color=GRIS)
    ws.freeze_panes = 'A5'


def main(entrada, salida):
    d = json.load(open(entrada))
    wb = Workbook()
    wb.remove(wb.active)
    hoja_parametros(wb)
    hoja_paradas(wb, d)
    hoja_cambios(wb, d)
    hoja_stock(wb, d)
    hoja_agotados(wb, d)
    hoja_servicio(wb, d)
    hoja_resumen(wb)
    orden = ['Resumen', 'Parámetros', 'Paradas', 'Cambios', 'Stock', 'Agotados', 'Servicio', 'Datos stock',
             'Datos servicio']
    wb._sheets = [wb[n] for n in orden]
    wb.active = 0
    wb.save(salida)
    print(salida)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
