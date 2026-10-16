"""Genera el maestro genérico de hallazgos y palancas de Plan to Release (P2R).

Uso: python generar_maestro_p2r.py [salida.xlsx]
"""
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

F = 'Arial'
NAVY, TEAL, TINT, GREY, BORDER = '00263A', '244C5A', 'EAF3FA', '5B6B78', 'D7DFE5'
OPOR, WARN = 'FFF1D6', 'FCE4E4'

FAM_I, FAM_O = 'Ineficiencia (coste)', 'Oportunidad (ingreso expuesto)'
ESF, FUGA, CIRC, TRANS, ING = ('Esfuerzo (ya contado dentro de personas)', 'Fuga de P&L', 'Circulante × CMPC',
                               'Transformación (horas de línea, mermas)', 'Ingreso expuesto')
NATURALEZAS = [ESF, FUGA, CIRC, TRANS, ING]
TIPOS = ['Política', 'Dato', 'Auto', 'Algoritmo', 'IA', 'Consulta', 'Eliminar', 'Sistema']
FASES = ['Demanda y previsión', 'S&OP y Plan Maestro', 'Necesidades y Aprovisionamiento', 'Programación',
         'Liberación', 'Seguimiento y Control', 'Transversal del proceso']

COMPARTE_INGRESO = 'Comparte ingreso con H14, H15, H16 y H39: cuantificar solo uno por celda.'

# (id, orden de fase, hallazgo, familia, naturaleza, métrica, nota)
HALLAZGOS = [
    ('H01', 1, 'No hay previsión estadística: es el criterio de Comercial o una copia del año anterior', FAM_I, CIRC,
     '€ de previsión mensual sin modelo', ''),
    ('H02', 1, 'La previsión no tiene día de cierre y se cambia dentro del mes en curso', FAM_I, CIRC,
     'Cambios de previsión dentro del mes en curso / mes', ''),
    ('H03', 1, 'No hay consenso de demanda entre Comercial y Planificación, o hay reunión pero no decide', FAM_I, CIRC,
     '€ de previsión mensual sin consensuar', ''),
    ('H04', 1, 'La previsión se usa como objetivo de ventas o presupuesto, no como demanda esperada', FAM_I, CIRC,
     'Sesgo (%) × € de previsión', ''),
    ('H05', 1, 'La demanda perdida (agotados, pedidos rechazados) no se registra y el histórico sale sesgado', FAM_I, CIRC,
     'Agotados o rechazos / mes sin registrar', 'Sin este dato, la previsión estadística (PL01) aprende el sesgo.'),
    ('H06', 1, 'No se miden precisión ni sesgo, o se miden sin umbral', FAM_I, CIRC,
     'Familias sin KPI de precisión', ''),
    ('H07', 1, 'Las desviaciones de previsión no llevan análisis de causa ni acción', FAM_I, CIRC,
     'Desviaciones fuera de umbral sin acción / mes', ''),
    ('H08', 1, 'El horizonte de previsión es más corto que el plazo de las decisiones que alimenta (turnos, compras de plazo largo, intercompany)',
     FAM_I, TRANS, '€ de compras o turnos decididos sin previsión', ''),
    ('H09', 2, 'No hay S&OP: el desequilibrio entre demanda y capacidad no se decide con dirección y se arregla abajo, con urgencias',
     FAM_I, FUGA, 'Horas extra y turnos no planificados / mes', 'Puede solapar con H10 y H29 (mismas horas extra): cuantificar una vez.'),
    ('H10', 2, 'La capacidad no se planifica con datos: turnos y balanceo se deciden por intuición, sin cruzar carga con capacidad',
     FAM_I, TRANS, 'Horas de línea / mes', ''),
    ('H11', 2, 'El plan maestro no tiene aprobación formal ni zona congelada y se cambia después de publicarlo', FAM_I, TRANS,
     'Cambios dentro de la zona congelada / semana', ''),
    ('H12', 2, 'No hay política de stock por segmento: ABC/XYZ, MTS/MTO y stock de seguridad sin definir o sin revisar', FAM_I, CIRC,
     'Materiales sin estrategia vigente', 'La política la decide Planificación (P2R); el dato se guarda en R2M.'),
    ('H13', 2, 'Stock de producto terminado excesivo u obsoleto por fabricar contra una previsión inflada', FAM_I, CIRC,
     '€ de stock sin movimiento > 6 meses', 'Si el stock se deprecia o se destruye, esa pérdida es Fuga de P&L.'),
    ('H14', 2, 'Roturas en materiales contra stock: hay pedidos que no se pueden servir', FAM_O, ING,
     '€ de pedidos no servidos / mes', COMPARTE_INGRESO),
    ('H15', 2, 'La fecha al cliente se promete sin comprobar stock ni capacidad', FAM_O, ING,
     '€ de pedidos / mes confirmados con fecha no viable', COMPARTE_INGRESO),
    ('H16', 2, 'Cuando se sabe que no se llega, la fecha no se renegocia con el cliente', FAM_O, ING,
     '€ de pedidos / mes entregados tarde sin aviso', COMPARTE_INGRESO),
    ('H17', 2, 'La cartera se revisa dos veces: Comercial y Planificación, cada uno por su lado', FAM_I, ESF,
     'Horas / día', ''),
    ('H18', 2, 'Cálculos repetitivos a mano: reservas de capacidad (temporales), coberturas, cargas', FAM_I, ESF,
     'Horas / mes', 'Es un cálculo fijo: mejor una regla automática que un agente.'),
    ('H19', 2, 'Intercompany (opcional): la planta que suministra se entera por un pedido urgente, sin plan conjunto ni mín/máx',
     FAM_I, TRANS, 'Pedidos intercompany urgentes / mes', 'Solo para las divisiones que se suministran entre sí.'),
    ('H20', 3, 'La propuesta del sistema se corrige a mano de forma sistemática y sin registrar el motivo', FAM_I, ESF,
     'Correcciones / día', 'Baja mucho al cerrar la previsión (PL02).'),
    ('H21', 3, 'Parámetros de planificación desactualizados: lote mínimo y múltiplo, plazos, stock de seguridad', FAM_I, CIRC,
     '% de materiales sin revisar en 12 meses', ''),
    ('H22', 3, 'El MRP corre con poca frecuencia y las faltas se detectan tarde', FAM_I, TRANS,
     'Faltas / mes detectadas con menos margen que el plazo del proveedor', ''),
    ('H23', 3, 'Las faltas se revisan todas a mano, sin criticidad ni gestión por excepción', FAM_I, ESF,
     'Líneas de falta revisadas / semana', ''),
    ('H24', 3, 'Compras y transportes urgentes por faltas no anticipadas', FAM_I, FUGA,
     '€ de sobrecoste urgente / mes', ''),
    ('H25', 3, 'Producto terminado y semielaborado se planifican por separado, sin sincronizar', FAM_I, CIRC,
     '€ de stock de semielaborado', ''),
    ('H26', 3, 'La agrupación de necesidades en lotes o combinaciones se hace a mano', FAM_I, TRANS,
     '% de merma de combinación', ''),
    ('H27', 3, 'Materias primas y auxiliares sin política de stock: se compra a ojo o por el mínimo del proveedor', FAM_I, CIRC,
     '€ de materia prima sin consumo > 6 meses', ''),
    ('H28', 4, 'La secuencia se hace a mano, sin optimizar los cambios (molde, formato, altura, color)', FAM_I, TRANS,
     'Horas de cambio / semana', ''),
    ('H29', 4, 'Se programa sin capacidad finita: se carga más de lo que cabe', FAM_I, TRANS,
     'Horas cargadas por encima de la capacidad / semana', ''),
    ('H30', 4, 'El programa no tiene zona congelada y los cambios de última hora obligan a resecuenciar', FAM_I, TRANS,
     'Resecuenciaciones / semana', ''),
    ('H31', 4, 'Urgencias y prioridades sin reglas: se decide por llamada o se escala a dirección', FAM_I, ESF,
     'Escalados / semana', ''),
    ('H32', 4, 'Se usa siempre la ruta por defecto, sin reequilibrar con la carga', FAM_I, TRANS,
     'Horas en ruta no óptima / mes', ''),
    ('H33', 4, 'El programa no tiene en cuenta a las personas (polivalencias, turnos) ni las paradas planificadas', FAM_I, TRANS,
     'Órdenes reprogramadas por personal o parada / mes', ''),
    ('H34', 5, 'Se libera sin comprobar material en línea, utillaje, ficha y personal', FAM_I, TRANS,
     'Arranques fallidos / mes', ''),
    ('H35', 5, 'No hay ventana de liberación: se libera todo de golpe o en el último momento', FAM_I, CIRC,
     'Días entre liberación e inicio', ''),
    ('H36', 5, 'El aprovisionamiento a línea no va sincronizado con la secuencia', FAM_I, TRANS,
     'Minutos de espera de material por orden', ''),
    ('H37', 5, 'La orden sale con una ficha técnica o una versión de producto no vigente', FAM_I, TRANS,
     'Órdenes con ficha no vigente / mes', ''),
    ('H38', 6, 'No hay KPIs formales de servicio (OTIF a fecha confirmada) ni de adherencia al programa', FAM_O, ING,
     'Pedidos / mes sin medir', 'Habilitador: no suma € por sí solo. Piso «sin declarar».'),
    ('H39', 6, 'Los atrasos se detectan cuando ya han ocurrido, sin aviso de pedidos en riesgo', FAM_O, ING,
     '€ de pedidos atrasados / mes', COMPARTE_INGRESO),
    ('H40', 6, 'Las desviaciones del plan no tienen análisis de causa y lo aprendido no vuelve a la previsión ni a los parámetros',
     FAM_I, TRANS, 'Desviaciones sin causa / mes', ''),
    ('H41', 6, 'Los informes de seguimiento se hacen a mano', FAM_I, ESF, 'Horas / semana', ''),
    ('H42', 7, 'La planificación se concentra en una persona, sin suplente ni procedimiento escrito', FAM_I, ESF,
     'Roles por persona', ''),
    ('H43', 7, 'No hay RACI: no está claro quién decide la previsión, la fecha y la prioridad', FAM_I, ESF,
     'Decisiones escaladas / mes', ''),
    ('H44', 7, 'Se planifica en Excel paralelos o con sistemas sin integrar, con doble carga', FAM_I, ESF,
     'Horas / semana de doble carga', ''),
]

# (id, palanca, tipo principal, tipo secundario, evoluciona a IA, hallazgos, KPI, intervención tipo)
PALANCAS = [
    ('PL01', 'Previsión estadística por familia y material, a 12–18 meses', 'Auto', '', 'Sí', ['H01', 'H08'],
     'Precisión (1–WMAPE)',
     'Modelo por familia y material con histórico limpio (incluida la demanda perdida) y estacionalidad; Comercial solo revisa las excepciones antes del cierre.'),
    ('PL02', 'Calendario de demanda con día de cierre y mes en curso congelado', 'Política', '', 'No', ['H02'],
     'Cambios tras el cierre',
     'Publicar el calendario anual con el día de cierre y bloquear en el sistema la previsión del mes en curso; los cambios posteriores, solo con aprobación y motivo.'),
    ('PL03', 'Revisión mensual de demanda con consenso: se planifica la demanda esperada, no el objetivo', 'Política', '', 'No',
     ['H03', 'H04'], 'Sesgo · consenso en fecha',
     'Reunión mensual Comercial–Planificación con orden del día fijo, acta y decisiones; previsión y presupuesto se cargan por separado.'),
    ('PL04', 'Registro de la demanda perdida con su motivo', 'Dato', '', 'No', ['H05'], '% de agotados registrados',
     'Definir qué es un agotado o un rechazo y registrarlo a diario en el sistema, con su motivo.'),
    ('PL05', 'KPI de precisión y sesgo por familia, con umbral y plan de acción', 'Auto', '', 'No', ['H06', 'H07'],
     'Familias fuera de umbral con acción',
     'Cálculo automático mensual por familia, umbral por segmento y plan de acción para cada familia fuera de umbral.'),
    ('PL06', 'S&OP mensual con escenarios valorados en € y reunión ejecutiva', 'Política', '', 'No', ['H09', 'H19'],
     'Horas extra no planificadas · cumplimiento del plan',
     'Ciclo mensual: revisión de demanda, revisión de suministro y reunión ejecutiva; cada escenario (turnos, stock, subcontratación, intercompany) valorado en €.'),
    ('PL07', 'Plan de capacidad: carga frente a capacidad por línea, con regla de turnos', 'Auto', '', 'Sí', ['H10'],
     'Ocupación real frente a la planificada',
     'Cruce automático de previsión cerrada × rutas × tiempos de ciclo contra el calendario; umbrales de ocupación para abrir o cerrar turnos.'),
    ('PL08', 'Zonas congeladas en el plan maestro y en el programa', 'Política', '', 'No', ['H11', 'H30'],
     'Cambios en zona congelada',
     'Definir los horizontes (congelado, firme, libre) del plan y del programa; las excepciones se aprueban y quedan registradas.'),
    ('PL09', 'Segmentación ABC/XYZ y política de stock por segmento, revisada cada trimestre', 'Política', 'Auto', 'No',
     ['H12', 'H13', 'H14', 'H27'], 'Cobertura · roturas · stock obsoleto',
     'Criterios ABC por valor y XYZ por variabilidad, matriz que asigna MTS/MTO y stock de seguridad, y recálculo trimestral automático.'),
    ('PL10', 'Promesa de fecha con disponibilidad real: stock y capacidad libre (ATP/CTP)', 'Auto', '', 'No', ['H15'],
     '% de pedidos confirmados con fecha viable',
     'Consulta de disponibilidad (stock y capacidad libre) al confirmar el pedido: la fecha la propone el sistema, no se teclea.'),
    ('PL11', 'Política de cartera: una sola revisión conjunta, dos fechas en el pedido y aviso al cliente', 'Política', 'Eliminar',
     'No', ['H16', 'H17'], 'OTIF a fecha confirmada',
     'Una revisión diaria conjunta; fecha solicitada y fecha confirmada en cada pedido; plazo fijo de confirmación y aviso al cliente si no se llega.'),
    ('PL12', 'Automatizar cálculos e informes repetitivos', 'Auto', '', 'No', ['H18', 'H41'], 'Horas liberadas',
     'Pasar a reglas del sistema los cálculos e informes que hoy se hacen en Excel cada día, semana o mes.'),
    ('PL13', 'Plan intercompany con mín/máx y pedidos programados', 'Política', '', 'No', ['H19'],
     'Pedidos intercompany urgentes',
     'Mín/máx por material entre plantas, pedidos programados y revisión semanal conjunta.'),
    ('PL14', 'Necesidades por excepción: alertas por criticidad y motivo obligatorio en cada corrección', 'Auto', '', 'Sí',
     ['H20', 'H23'], 'Correcciones / día',
     'El sistema marca solo las excepciones, ordenadas por criticidad; cada corrección manual exige un motivo de una lista cerrada.'),
    ('PL15', 'Revisión periódica de los parámetros de planificación contra la realidad', 'Dato', '', 'No', ['H21'],
     '% de parámetros vigentes',
     'Revisión semestral de lotes, plazos y stock de seguridad contra los datos reales, con un responsable por parámetro.'),
    ('PL16', 'MRP diario con alertas tempranas de faltas', 'Auto', '', 'No', ['H22', 'H24'],
     'Faltas detectadas a tiempo · € urgentes',
     'MRP diario (nocturno) con alerta de faltas priorizada por fecha de necesidad y por pedido de cliente.'),
    ('PL17', 'Planificación multinivel sincronizada: producto terminado, semielaborado y materias primas', 'Sistema', '', 'No',
     ['H25'], 'Stock de semielaborado',
     'Planificar los tres niveles en el mismo cálculo, con trazabilidad de cada necesidad hasta su origen.'),
    ('PL18', 'Optimizador de agrupación y combinación', 'Algoritmo', '', 'No', ['H26'], '% de merma',
     'Algoritmo de agrupación por reglas de lote y combinación que minimiza la merma y los cambios.'),
    ('PL19', 'Secuenciador con capacidad finita, matriz de cambios y rutas alternativas', 'Algoritmo', '', 'Sí',
     ['H28', 'H29', 'H32'], 'Horas de cambio · adherencia',
     'Matriz de tiempos de cambio por línea y capacidad finita; propone la secuencia y las rutas alternativas, y el programador la aprueba.'),
    ('PL20', 'Reglas de prioridad y de escalado', 'Política', '', 'No', ['H31'], 'Escalados / semana',
     'Criterios escritos (cliente, fecha, margen, penalización) y un árbitro definido para los conflictos.'),
    ('PL21', 'Programar teniendo en cuenta personas y mantenimiento', 'Dato', 'Sistema', 'No', ['H33'],
     'Reprogramaciones por personal o parada',
     'Cargar en el programa la matriz de polivalencias, los turnos y las paradas preventivas.'),
    ('PL22', 'Semáforo de liberación 24–72 h antes, con ventana de liberación', 'Auto', '', 'No',
     ['H34', 'H35', 'H36', 'H37'], 'Arranques fallidos',
     'Comprobación automática de material en línea, utillaje, ficha vigente y personal; la orden se libera en verde y dentro de su ventana.'),
    ('PL23', 'Cuadro de mando de servicio y adherencia, con alertas de pedidos en riesgo', 'Auto', '', 'No', ['H38', 'H39'],
     'OTIF · avisos anticipados',
     'OTIF a fecha confirmada y adherencia al programa calculados solos, con alerta de los pedidos en riesgo antes de que se retrasen.'),
    ('PL24', 'Análisis de causa raíz de las desviaciones, que vuelve a la previsión y a los parámetros', 'Política', '', 'Sí',
     ['H07', 'H40'], 'Desviaciones con causa',
     'Clasificar cada desviación por causa (lista cerrada), asignar una acción con responsable y devolver lo aprendido a la previsión y a los parámetros.'),
    ('PL25', 'Copiloto de consulta: «¿cuándo sirvo X?», «¿por qué va tarde X?»', 'Consulta', '', 'No', ['H15', 'H17', 'H39'],
     'Consultas resueltas sin llamada',
     'Asistente que responde sobre plazos y estado de pedidos con los datos del sistema, sin modificar nada.'),
    ('PL26', 'Gobierno: RACI, titular y suplente por rol, procedimiento escrito', 'Política', '', 'No', ['H42', 'H43'],
     'Roles con suplente',
     'RACI por actividad, titular y suplente por rol y un procedimiento escrito por fase.'),
    ('PL27', 'Planificar dentro del sistema: sin Excel paralelos y con los sistemas integrados', 'Sistema', '', 'No', ['H44'],
     'Horas de doble carga',
     'Inventariar los Excel paralelos, llevar su lógica al sistema e integrar los sistemas para cargar cada dato una sola vez.'),
]

# (id, proceso destino, hallazgo, familia, naturaleza, métrica, palanca en ese proceso, tipo)
OTROS = [
    ('X01', 'R2M', 'El alta de material no pasa por Planificación (estrategia, ruta, parámetros)', FAM_I, CIRC,
     'Altas / mes sin validar por Planificación', 'Validación de Planificación en el alta de material', 'Política'),
    ('X02', 'R2M', 'Duplicados en el maestro, sin criterio de qué es un material', FAM_I, ESF,
     'Materiales duplicados', 'Criterio de material y auditoría de duplicados', 'Dato'),
    ('X03', 'R2M', 'El calendario de turnos y paradas, las velocidades y los tiempos de cambio no coinciden con la realidad',
     FAM_I, TRANS, '% de recursos con datos sin revisar', 'Maestro de recursos de fabricación mantenido', 'Dato'),
    ('X04', 'P2P', 'Plazos de proveedor desactualizados y confirmaciones sin registrar', FAM_I, CIRC,
     '% de líneas con fecha confirmada distinta de la prevista', 'Seguimiento de confirmaciones de proveedor', 'Auto'),
    ('X05', 'O2C', 'Cartera con errores (cantidades mayores que el pedido, fechas imposibles) y pedidos servidos sin saldar',
     FAM_I, CIRC, 'Líneas de cartera erróneas o servidas sin saldar', 'Validación de cargas y saldado por regla', 'Auto'),
]

# (n.º, hallazgo creado en Ondupet, id del maestro, nota)
ONDUPET_H = [
    (1, 'Sin cierre de previsiones: se modifica a mitad de mes', 'H02', ''),
    (2, 'No se hace carga de capacidad (se decide por sensaciones)', 'H10', ''),
    (3, 'No se calculan ABC ni estrategias', 'H12', 'Antes lo mandé a R2M; queda mejor en P2R porque falta la política, no el dato.'),
    (4, 'Temporales calculados a mano cada mes', 'H18', ''),
    (5, 'No se modifican las fechas al cliente aunque se sepa que no se llega', 'H16', 'Familia: Oportunidad (ingreso expuesto).'),
    (6, 'Saldado de pedidos manual; a veces no se hace', 'X05', 'Va a O2C × Ondupet con «Detectado en: P2R».'),
    (7, 'Corrección manual de desviaciones de demanda', 'H20', ''),
    (8, 'Revisión de faltas manual', 'H23', ''),
    (9, 'Secuenciación a mano', 'H28', ''),
]
# (n.º, palanca creada en Ondupet, id principal, complementarias)
ONDUPET_P = [
    ('P1', 'Previsión cerrada y congelada', 'PL02', 'PL05'),
    ('P2', 'Capacidad decidida con datos', 'PL07', 'PL12'),
    ('P3', 'Segmentación y estrategias al día', 'PL09', ''),
    ('P4', 'Fecha comprometida viable', 'PL11', 'PL10, PL23'),
    ('P5', 'Saldado por regla', 'X05', ''),
    ('P6', 'Necesidades y faltas por excepción', 'PL14', 'PL16'),
    ('P7', 'Secuencia optimizada', 'PL19', 'PL08, PL20'),
]

THIN = Side(style='thin', color=BORDER)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='top', wrap_text=True)


def fill(c):
    return PatternFill('solid', fgColor=c)


def titulo(ws, texto, sub):
    ws['A1'] = texto
    ws['A1'].font = Font(name=F, size=14, bold=True, color=NAVY)
    ws['A2'] = sub
    ws['A2'].font = Font(name=F, size=9, italic=True, color=GREY)


def tabla(ws, fila, cabeceras, anchos, filas, centradas=()):
    """Escribe cabecera y filas desde `fila`; devuelve la última fila escrita."""
    for j, (h, w) in enumerate(zip(cabeceras, anchos), 1):
        c = ws.cell(fila, j, h)
        c.font, c.fill, c.alignment, c.border = Font(name=F, size=10, bold=True, color='FFFFFF'), fill(NAVY), CENTER, BOX
        ws.column_dimensions[c.column_letter].width = w
    for i, valores in enumerate(filas, fila + 1):
        for j, v in enumerate(valores, 1):
            c = ws.cell(i, j, v)
            c.font = Font(name=F, size=10, color=NAVY)
            c.alignment = CENTER if j in centradas else WRAP
            c.border = BOX
    ultima = fila + len(filas)
    ws.freeze_panes = ws.cell(fila + 1, 1)
    ws.auto_filter.ref = f'A{fila}:{ws.cell(fila, len(cabeceras)).column_letter}{ultima}'
    return ultima


def lista(wb, ws, rango, fuente):
    dv = DataValidation(type='list', formula1=fuente, allow_blank=True)
    dv.error, dv.errorTitle = 'Elige un valor de la lista (hoja Listas).', 'Valor no válido'
    ws.add_data_validation(dv)
    dv.add(rango)


def construir(salida):
    wb = Workbook()
    ley = wb.active
    ley.title = 'Leyenda'
    wh, wp, wr, wo, wn, wres = (wb.create_sheet(n) for n in
                                ('Hallazgos', 'Palancas', 'Relación', 'Otros procesos', 'Ondupet', 'Resumen'))
    wl = wb.create_sheet('Listas')

    # --- Listas (fuente de las validaciones) ---
    for j, (cab, vals) in enumerate([('Fase', FASES), ('Familia', [FAM_I, FAM_O]), ('Naturaleza', NATURALEZAS),
                                     ('Tipo de palanca', TIPOS), ('Sí/No', ['Sí', 'No'])], 1):
        c = wl.cell(1, j, cab)
        c.font, c.fill = Font(name=F, size=10, bold=True, color='FFFFFF'), fill(NAVY)
        wl.column_dimensions[c.column_letter].width = 42
        for i, v in enumerate(vals, 2):
            wl.cell(i, j, v).font = Font(name=F, size=10, color=NAVY)
    rng = {'fase': f'Listas!$A$2:$A${len(FASES) + 1}', 'fam': 'Listas!$B$2:$B$3',
           'nat': f'Listas!$C$2:$C${len(NATURALEZAS) + 1}', 'tipo': f'Listas!$D$2:$D${len(TIPOS) + 1}',
           'sino': 'Listas!$E$2:$E$3'}

    # --- Hallazgos ---
    titulo(wh, 'Maestro de hallazgos · Plan to Release (P2R)',
           'Genérico para cualquier empresa. El anclaje concreto (paso) y la cuantificación se ponen al usarlo en cada celda.')
    pal_de = {}
    for p in PALANCAS:
        for h in p[5]:
            pal_de.setdefault(h, []).append(p[0])
    n_rel = sum(len(p[5]) for p in PALANCAS)
    rel_ult = 4 + n_rel
    filas = []
    for k, (hid, fo, txt, fam, nat, met, nota) in enumerate(HALLAZGOS):
        r = 5 + k
        filas.append([hid, 'P2R', fo, FASES[fo - 1], txt, fam, nat, met, ', '.join(pal_de.get(hid, [])),
                      f'=COUNTIF(Relación!$C$5:$C${rel_ult},A{r})', nota])
    h_ult = tabla(wh, 4, ['ID', 'Proceso', 'Orden', 'Fase', 'Hallazgo', 'Familia', 'Naturaleza', 'Métrica (volumen expuesto)',
                          'Palancas', 'Nº palancas', 'Notas'],
                  [7, 9, 7, 24, 58, 22, 26, 32, 16, 10, 40], filas, centradas=(1, 2, 3, 10))
    lista(wb, wh, f'D5:D{h_ult}', rng['fase'])
    lista(wb, wh, f'F5:F{h_ult}', rng['fam'])
    lista(wb, wh, f'G5:G{h_ult}', rng['nat'])
    wh.conditional_formatting.add(f'A5:K{h_ult}', FormulaRule(formula=[f'$F5="{FAM_O}"'], fill=fill(OPOR)))
    wh.conditional_formatting.add(f'J5:J{h_ult}', FormulaRule(formula=['$J5=0'], fill=fill(WARN),
                                                              font=Font(name=F, bold=True, color='B42318')))

    # --- Palancas ---
    titulo(wp, 'Maestro de palancas · Plan to Release (P2R)',
           'La palanca dice qué cambia; la intervención tipo, cómo suele hacerse. Responsable y fechas se fijan en cada celda.')
    filas = []
    for k, (pid, nom, t1, t2, ia, hs, kpi, como) in enumerate(PALANCAS):
        r = 5 + k
        filas.append([pid, 'P2R', nom, t1, t2 or None, ia, ', '.join(hs), f'=COUNTIF(Relación!$A$5:$A${rel_ult},A{r})',
                      kpi, como])
    p_ult = tabla(wp, 4, ['ID', 'Proceso', 'Palanca', 'Tipo principal', 'Tipo secundario', 'Evoluciona a IA',
                          'Hallazgos que resuelve', 'Nº hallazgos', 'KPI', 'Intervención tipo (cómo)'],
                  [7, 9, 50, 14, 14, 12, 22, 11, 30, 62], filas, centradas=(1, 2, 4, 5, 6, 8))
    lista(wb, wp, f'D5:E{p_ult}', rng['tipo'])
    lista(wb, wp, f'F5:F{p_ult}', rng['sino'])

    # --- Relación (una fila por par palanca–hallazgo) ---
    titulo(wr, 'Relación palanca ↔ hallazgo',
           'Una fila por par, lista para importar. Los textos se buscan en las hojas Palancas y Hallazgos.')
    filas = []
    k = 0
    for p in PALANCAS:
        for h in p[5]:
            r = 5 + k
            filas.append([p[0], f'=INDEX(Palancas!$C$5:$C${p_ult},MATCH(A{r},Palancas!$A$5:$A${p_ult},0))', h,
                          f'=INDEX(Hallazgos!$E$5:$E${h_ult},MATCH(C{r},Hallazgos!$A$5:$A${h_ult},0))',
                          f'=INDEX(Hallazgos!$D$5:$D${h_ult},MATCH(C{r},Hallazgos!$A$5:$A${h_ult},0))'])
            k += 1
    tabla(wr, 4, ['Palanca', 'Texto de la palanca', 'Hallazgo', 'Texto del hallazgo', 'Fase del hallazgo'],
          [10, 50, 10, 60, 28], filas, centradas=(1, 3))

    # --- Otros procesos ---
    titulo(wo, 'Detectados en P2R, pero van al maestro de otro proceso',
           'Créalos en el proceso destino con «Detectado en: P2R»: el coste suma donde se arregla.')
    filas = [[x[0], x[1], 'P2R', x[2], x[3], x[4], x[5], x[6], x[7]] for x in OTROS]
    o_ult = tabla(wo, 4, ['ID', 'Proceso destino', 'Detectado en', 'Hallazgo', 'Familia', 'Naturaleza',
                          'Métrica (volumen expuesto)', 'Palanca en ese proceso', 'Tipo'],
                  [7, 11, 11, 58, 22, 26, 32, 40, 12], filas, centradas=(1, 2, 3, 9))
    lista(wb, wo, f'E5:E{o_ult}', rng['fam'])
    lista(wb, wo, f'F5:F{o_ult}', rng['nat'])
    lista(wb, wo, f'I5:I{o_ult}', rng['tipo'])

    # --- Ondupet: dónde cae lo ya creado ---
    titulo(wn, 'Ondupet · dónde cae lo que ya está creado',
           'Para re-vincular los hallazgos y palancas creados directamente en la celda P2R × Ondupet.')
    look_h = (lambda r: f'=IFERROR(INDEX(Hallazgos!$E$5:$E${h_ult},MATCH(C{r},Hallazgos!$A$5:$A${h_ult},0)),'
                        f'INDEX(\'Otros procesos\'!$D$5:$D${o_ult},MATCH(C{r},\'Otros procesos\'!$A$5:$A${o_ult},0)))')
    filas = [[n, txt, mid, look_h(5 + i), nota] for i, (n, txt, mid, nota) in enumerate(ONDUPET_H)]
    ult = tabla(wn, 4, ['Nº', 'Hallazgo creado', 'ID maestro', 'Hallazgo del maestro', 'Nota'],
                [8, 50, 12, 60, 50], filas, centradas=(1, 3))
    r0 = ult + 3
    wn.cell(r0 - 1, 1, 'Palancas').font = Font(name=F, size=12, bold=True, color=NAVY)
    look_p = (lambda r: f'=IFERROR(INDEX(Palancas!$C$5:$C${p_ult},MATCH(C{r},Palancas!$A$5:$A${p_ult},0)),'
                        f'INDEX(\'Otros procesos\'!$H$5:$H${o_ult},MATCH(C{r},\'Otros procesos\'!$A$5:$A${o_ult},0)))')
    filas = [[n, txt, mid, look_p(r0 + 1 + i), comp or None] for i, (n, txt, mid, comp) in enumerate(ONDUPET_P)]
    tabla(wn, r0, ['Nº', 'Palanca creada', 'ID maestro', 'Palanca del maestro', 'Complementarias'],
          [8, 50, 12, 60, 50], filas, centradas=(1, 3))
    wn.freeze_panes = 'A5'
    wn.auto_filter.ref = None

    # --- Resumen (todo con fórmulas sobre las hojas maestras) ---
    titulo(wres, 'Resumen del maestro', 'Se recalcula solo si añades o cambias filas en Hallazgos y Palancas.')
    H = lambda col: f'Hallazgos!${col}$5:${col}${h_ult}'
    hdr = lambda r, vals: [(lambda c: (setattr(c, 'font', Font(name=F, size=10, bold=True, color='FFFFFF')),
                                       setattr(c, 'fill', fill(NAVY)), setattr(c, 'alignment', CENTER),
                                       setattr(c, 'border', BOX)))(wres.cell(r, j, v)) for j, v in enumerate(vals, 1)]

    def celda(r, j, v, bold=False):
        c = wres.cell(r, j, v)
        c.font, c.border = Font(name=F, size=10, bold=bold, color=NAVY), BOX
        c.alignment = WRAP if j == 1 else CENTER
        return c

    hdr(4, ['Fase', 'Ineficiencia', 'Oportunidad', 'Total'])
    for i, fase in enumerate(FASES, 5):
        celda(i, 1, fase)
        celda(i, 2, f'=COUNTIFS({H("D")},A{i},{H("F")},"{FAM_I}")')
        celda(i, 3, f'=COUNTIFS({H("D")},A{i},{H("F")},"{FAM_O}")')
        celda(i, 4, f'=B{i}+C{i}', bold=True)
    t = 5 + len(FASES)
    celda(t, 1, 'Total', bold=True)
    for j, col in enumerate('BCD', 2):
        celda(t, j, f'=SUM({col}5:{col}{t - 1})', bold=True)

    r = t + 3
    hdr(r, ['Naturaleza', 'Hallazgos'])
    for i, nat in enumerate(NATURALEZAS, r + 1):
        celda(i, 1, nat)
        celda(i, 2, f'=COUNTIF({H("G")},A{i})')
    r = r + len(NATURALEZAS) + 3
    hdr(r, ['Tipo principal de palanca', 'Palancas', 'Evolucionan a IA'])
    for i, tp in enumerate(TIPOS, r + 1):
        celda(i, 1, tp)
        celda(i, 2, f'=COUNTIF(Palancas!$D$5:$D${p_ult},A{i})')
        celda(i, 3, f'=COUNTIFS(Palancas!$D$5:$D${p_ult},A{i},Palancas!$F$5:$F${p_ult},"Sí")')
    r = r + len(TIPOS) + 3
    hdr(r, ['Control de calidad', 'Valor'])
    celda(r + 1, 1, 'Hallazgos sin palanca (debe ser 0)')
    celda(r + 1, 2, f'=COUNTIF(Hallazgos!$J$5:$J${h_ult},0)', bold=True)
    celda(r + 2, 1, 'Palancas sin hallazgo (debe ser 0)')
    celda(r + 2, 2, f'=COUNTIF(Palancas!$H$5:$H${p_ult},0)', bold=True)
    celda(r + 3, 1, 'Pares palanca–hallazgo')
    celda(r + 3, 2, f'=COUNTA(Relación!$A$5:$A${rel_ult})', bold=True)
    for col, w in zip('ABCD', (40, 14, 14, 10)):
        wres.column_dimensions[col].width = w

    # --- Leyenda ---
    ley.column_dimensions['A'].width = 3
    ley.column_dimensions['B'].width = 30
    ley.column_dimensions['C'].width = 100
    ley['B2'] = 'Maestro de hallazgos y palancas · Plan to Release (P2R)'
    ley['B2'].font = Font(name=F, size=18, bold=True, color=NAVY)
    ley['B3'] = 'Catálogo genérico para cualquier empresa del grupo · 16/10/2026'
    ley['B3'].font = Font(name=F, size=11, italic=True, color=TEAL)
    bloques = [
        ('Hojas', [
            ('Hallazgos', f'{len(HALLAZGOS)} hallazgos genéricos por fase, con familia, naturaleza y métrica.'),
            ('Palancas', f'{len(PALANCAS)} palancas con su tipo, los hallazgos que resuelven, KPI e intervención tipo.'),
            ('Relación', 'Una fila por par palanca–hallazgo, lista para importar.'),
            ('Otros procesos', 'Hallazgos que se detectan en P2R pero se arreglan en R2M, P2P u O2C.'),
            ('Ondupet', 'Dónde cae cada hallazgo y palanca ya creados en la celda P2R × Ondupet.'),
            ('Resumen', 'Recuentos por fase, familia, naturaleza y tipo, y controles de calidad.'),
            ('Listas', 'Valores permitidos de los desplegables. Si cambias una lista, cambia aquí.'),
        ]),
        ('Familia', [
            (FAM_I, 'Cuesta dinero hoy: horas, stock, horas de línea, sobrecostes.'),
            (FAM_O, 'Pone en riesgo una venta. No significa «oportunidad de mejora».'),
        ]),
        ('Naturaleza', [
            (ESF, 'Horas de personas. Libera horas, pero no suma € nuevos: ya están en el coste de personas.'),
            (FUGA, 'Sobrecostes y pérdidas directas: urgencias, horas extra, penalizaciones, obsolescencia.'),
            (CIRC, 'Stock o necesidades de más, valorados al coste del capital.'),
            (TRANS, 'Horas de línea perdidas en cambios o paradas, y mermas.'),
            (ING, 'Solo para Oportunidad. Si la app ofrece otra naturaleza para Oportunidad, cámbiala en Listas.'),
        ]),
        ('Tipos de palanca', [
            ('Política', 'Regla o forma de trabajar, sin desarrollo.'),
            ('Dato', 'Corregir o mantener datos.'),
            ('Auto', 'Regla, cálculo o integración, sin IA.'),
            ('Algoritmo', 'Optimizador (secuencia, combinación).'),
            ('IA', 'Agente que propone; la persona aprueba.'),
            ('Consulta', 'Copiloto que solo lee datos y responde.'),
            ('Eliminar', 'El paso desaparece.'),
            ('Sistema', 'Cambio de herramienta o integración entre sistemas.'),
            ('Evoluciona a IA = Sí', 'Se implanta primero como regla y después puede pasar a agente.'),
        ]),
        ('Reglas de uso', [
            ('Anclaje', 'En el maestro solo va la fase. El paso concreto se elige al usar el hallazgo en cada celda.'),
            ('Cuantificación', 'En el maestro solo va la métrica. El valor y el tipo de piso se ponen en cada celda.'),
            ('Ingreso una vez', 'H14, H15, H16 y H39 exponen el mismo ingreso: en cada celda, cuantifica solo uno.'),
            ('Habilitadores', 'H38 no suma € por sí solo: déjalo con el piso «sin declarar».'),
            ('Añadir filas', 'Nueva fila al final de Hallazgos o Palancas con el siguiente ID, y su par en Relación. '
                             'Resumen avisa si queda un hallazgo sin palanca.'),
            ('Celdas a editar', 'Todas las columnas de texto. No edites «Nº palancas», «Nº hallazgos» ni los textos de '
                                'Relación y Ondupet: son fórmulas.'),
            ('Color', 'Las filas amarillas son de familia Oportunidad. Un «Nº palancas» en rojo es un hallazgo sin palanca.'),
        ]),
    ]
    r = 5
    for titulo_b, items in bloques:
        ley.cell(r, 2, titulo_b).font = Font(name=F, size=12, bold=True, color=NAVY)
        r += 1
        for a, b in items:
            ca, cb = ley.cell(r, 2, a), ley.cell(r, 3, b)
            ca.font, cb.font = Font(name=F, size=10, bold=True, color=TEAL), Font(name=F, size=10, color=NAVY)
            ca.alignment = cb.alignment = WRAP
            r += 1
        r += 1
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    wb.save(salida)


if __name__ == '__main__':
    out = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).resolve().parents[1] / 'Maestro_P2R_Hallazgos_Palancas.xlsx')
    construir(out)
    print(out)
