"""Genera las intervenciones propuestas para las palancas aplicadas en P2R × Ondupet.

Las columnas siguen el formulario «Nueva intervención» de la aplicación. Las cajas son las del flujo
as-is de Ondupet tal como están dibujadas.

Uso: python generar_intervenciones_ondupet.py [salida.xlsx]
"""
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

F = 'Arial'
NAVY, GREY, BORDER = '00263A', '5B6B78', 'D7DFE5'
COLOR_TIPO = {'Negocio': 'EEF1F4', 'Automatización': 'E3F3EA', 'Agente': 'EEE9FB', 'Consulta': 'E4EEFB'}

PALANCAS = [
    ('P2R-01', 'Gestión de la demanda', 'HT-P2R-01'),
    ('P2R-02', 'S&OP y plan de capacidad', 'HT-P2R-03'),
    ('P2R-03', 'Política de inventario por segmento', 'HT-P2R-04, HT-P2R-05'),
    ('P2R-04', 'Promesa de fecha fiable', 'HT-P2R-05'),
    ('P2R-06', 'Programación estable y optimizada', 'HT-P2R-08, HT-P2R-09'),
    ('P2R-08', 'Control del servicio y mejora continua', 'HT-P2R-11, HT-P2R-12'),
]

# Cajas del flujo as-is de Ondupet (fase, caja)
CAJAS = [
    ('Demanda y previsión', 'Introducción Previsiones Sistema'),
    ('Demanda y previsión', 'Comprobación de Accuracy Mes en Curso'),
    ('Demanda y previsión', '¿Buen % de Accuracy?'),
    ('Demanda y previsión', 'Corrección Previsiones Futuros Meses'),
    ('S&OP y Plan Maestro', 'Plan de Fabricación → Decisión de turnos y balanceo de materiales'),
    ('S&OP y Plan Maestro', 'Temporales → Reserva de Ciclo Por Familia'),
    ('S&OP y Plan Maestro', 'Cartera Pedidos'),
    ('Necesidades y Aprovisionamiento', 'Propuesta de Fabricación PT'),
    ('Necesidades y Aprovisionamiento', 'Agrupación de Necesidades y Propuesta de Fabricación ST'),
    ('Necesidades y Aprovisionamiento', 'MRP Semanal → Faltas de MMPP y auxiliares'),
    ('Programación', 'Secuencia Por Línea → Ciclo · Molde · Altura'),
    ('Programación', '¿Cambios de última hora?'),
    ('Programación', 'Recuenciar urgencias y prioridades'),
    ('Liberación', '¿Utillaje y Aprovisionamiento a línea Correcto?'),
    ('Liberación', 'Liberación de la Orden'),
    ('Seguimiento y Control', 'Seguimiento Diario Cumplimiento Servicio'),
    ('Seguimiento y Control', '¿Hay Atrasos?'),
    ('Seguimiento y Control', 'Ajustes Planificación y Pedidos'),
]
NUEVA = 'Nueva tarea'

D = lambda d, m, a: date(a, m, d)


def I(pal, tipo, nombre, que, inicio, fin, resultado, criterio, cajas=(), comite='', cob=None, exc=None,
      depende='', decidir='', detalle=''):
    return dict(pal=pal, tipo=tipo, nombre=nombre, que=que, ini=inicio, fin=fin, res=resultado, crit=criterio,
                cajas=list(cajas), comite=comite, cob=cob, exc=exc, dep=depende, dec=decidir, det=detalle)


INTERVENCIONES = [
    # ---------------------------------------------------------------- P2R-01 · Gestión de la demanda
    I('P2R-01', 'Negocio', 'Calendario de previsión con día de cierre',
      'Fijar un día de cierre mensual de la previsión y congelar el mes en curso. Después del cierre, cualquier cambio '
      'necesita la aprobación de Planificación y queda registrado con su motivo. Publicar el calendario del año.',
      D(1, 2, 2027), D(31, 3, 2027),
      'La previsión del mes en curso deja de cambiarse sin aviso a mitad de mes.',
      'Primer cierre en fecha el 31/03/2027 y 0 cambios del mes en curso sin aprobación durante 3 meses seguidos.',
      decidir='Día de cierre · quién aprueba las excepciones', detalle='D-PL02'),
    I('P2R-01', 'Negocio', 'Reunión mensual de consenso de demanda',
      'Reunión mensual Comercial–Planificación antes del cierre, con orden del día fijo (desviaciones, cambios y riesgos), '
      'acta y decisiones. Se acuerda quién responde del stock fabricado por cambios fuera de plazo.',
      D(1, 3, 2027), D(30, 4, 2027),
      'La previsión que entra al plan está consensuada y cada cambio fuera de plazo tiene un responsable.',
      '100 % de los meses con reunión y acta antes del cierre.',
      depende='Calendario de previsión con día de cierre', decidir='Quién responde del stock fabricado por cambios',
      detalle='D-PL03'),
    I('P2R-01', 'Negocio', 'Registro diario de agotados y demanda perdida',
      'Definir qué es un agotado y qué es un pedido rechazado, y registrarlos cada día en el sistema con su motivo, '
      'para que el histórico de demanda deje de estar distorsionado.',
      D(1, 3, 2027), D(30, 4, 2027),
      'El histórico de demanda recoge lo que se dejó de vender.',
      '≥ 95 % de los agotados registrados con motivo en la auditoría mensual.',
      decidir='Definición de agotado · quién registra', detalle='D-PL04'),
    I('P2R-01', 'Automatización', 'Bloqueo del mes en curso y registro de cambios en GIO',
      'GIO bloquea la previsión del mes en curso tras el cierre. Cuando se aprueba una excepción, guarda quién la cambia, '
      'cuándo y por qué.',
      D(1, 4, 2027), D(30, 4, 2027),
      'Ningún cambio de previsión fuera de plazo pasa sin aprobación ni sin dejar rastro.',
      '100 % de los cambios posteriores al cierre con aprobación y motivo registrados.',
      cajas=[('Introducción Previsiones Sistema', 'Se simplifica')],
      comite='La previsión se introduce hasta el día de cierre; después el sistema la bloquea y cada excepción queda trazada.',
      cob=1.0, exc=0.05, depende='Calendario de previsión con día de cierre', detalle='D-PL02'),
    I('P2R-01', 'Automatización', 'KPI automático de precisión y sesgo por familia',
      'Calcular cada mes la precisión (1–WMAPE) y el sesgo por familia, con un umbral por segmento. El rombo «¿Buen % de '
      'Accuracy?» lo resuelve la regla y solo las familias fuera de umbral pasan a revisión.',
      D(5, 4, 2027), D(28, 5, 2027),
      'La revisión de accuracy deja de ser manual y se centra en las familias que fallan.',
      'KPI publicado el día 2 de cada mes sin trabajo manual; cada familia fuera de umbral con una acción.',
      cajas=[('Comprobación de Accuracy Mes en Curso', 'La hace el sistema'), ('¿Buen % de Accuracy?', 'La hace el sistema')],
      comite='La comprobación de accuracy la calcula el sistema y la decisión «¿buen accuracy?» la toma una regla con umbral.',
      cob=1.0, exc=0.20, decidir='Umbral de precisión por segmento', detalle='D-PL05'),
    I('P2R-01', 'Agente', 'Agente de análisis de desviaciones',
      'Cada mes analiza las familias fuera de umbral, explica la causa probable (agotado, pedido puntual, cliente nuevo, '
      'cambio de Comercial) y propone la corrección de los meses siguientes. Planificación y Comercial la aprueban en la '
      'reunión de consenso.',
      D(1, 11, 2027), D(31, 3, 2028),
      'Toda desviación relevante tiene causa y una corrección propuesta antes de la reunión.',
      '≥ 80 % de las propuestas aceptadas sin rehacer y sesgo dentro del umbral 3 meses seguidos.',
      cajas=[('Corrección Previsiones Futuros Meses', 'Se simplifica')],
      comite='La corrección de previsiones la prepara un agente con su causa; las personas solo la aprueban.',
      cob=0.80, exc=0.20, depende='KPI automático de precisión y sesgo por familia (6 meses de datos)', detalle='D-PL24'),
    I('P2R-01', 'Agente', 'Agente de previsión base',
      'Antes del cierre genera una previsión estadística por familia y material (histórico, estacionalidad y agotados). '
      'Comercial deja de cargar la previsión desde cero y solo ajusta las excepciones.',
      D(1, 1, 2028), D(30, 6, 2028),
      'La previsión parte de un modelo y Comercial aporta solo lo que el modelo no sabe.',
      'Precisión del agente igual o mejor que la de Comercial 3 meses seguidos; ≥ 80 % de referencias sin ajuste manual.',
      cajas=[('Introducción Previsiones Sistema', 'Se simplifica')],
      comite='La previsión la prepara un agente y Comercial revisa solo las excepciones.',
      cob=0.80, exc=0.20, depende='Registro diario de agotados (12 meses de histórico limpio)', detalle='D-PL01'),

    # ---------------------------------------------------------------- P2R-02 · S&OP y plan de capacidad
    I('P2R-02', 'Negocio', 'Capacidad real de extrusión y termoformado',
      'Levantar con Producción el calendario de turnos y paradas y validar los tiempos de ciclo por línea de extrusión y '
      'termoformado. El dato se mantiene en R2M; aquí se pide y se usa.',
      D(1, 4, 2027), D(28, 5, 2027),
      'Hay un dato fiable de capacidad para cruzarlo con la carga.',
      '100 % de las líneas con calendario y tiempos de ciclo validados.',
      decidir='Quién mantiene el calendario en R2M', detalle='Maestro X03 (R2M)'),
    I('P2R-02', 'Negocio', 'Revisión mensual de capacidad con dirección de planta',
      'Tras el cierre de la previsión, Planificación presenta la carga frente a la capacidad y la dirección de planta '
      'decide turnos y balanceo con umbrales de ocupación acordados. Es el S&OP de planta mientras llega el de grupo.',
      D(3, 5, 2027), D(30, 6, 2027),
      'Turnos y balanceo se deciden una vez al mes, con datos y con dirección.',
      '100 % de los meses con decisión de turnos documentada antes del inicio de mes.',
      depende='Capacidad real de extrusión y termoformado',
      decidir='Umbrales de ocupación para abrir o cerrar turno', detalle='D-PL06'),
    I('P2R-02', 'Automatización', 'Carga frente a capacidad automática',
      'Al cerrar la previsión, el sistema cruza previsión × rutas × tiempos de ciclo con el calendario y muestra la '
      'ocupación de cada línea de extrusión y termoformado, marcando las que pasan del umbral.',
      D(1, 6, 2027), D(30, 7, 2027),
      'Ya no se meten necesidades sin saber si caben.',
      'Ocupación real dentro de ±10 % de la planificada y 0 turnos no planificados al mes.',
      cajas=[('Plan de Fabricación → Decisión de turnos y balanceo de materiales', 'Se simplifica')],
      comite='La decisión de turnos y balanceo se toma sobre un cálculo de carga frente a capacidad, no por intuición.',
      cob=1.0, exc=0.15, depende='Capacidad real de extrusión y termoformado · Bloqueo del mes en curso', detalle='D-PL07'),
    I('P2R-02', 'Automatización', 'Temporales automáticos',
      'La reserva de ciclo por familia se calcula sola en cada cierre a partir de la previsión cerrada y los tiempos de '
      'ciclo. Planificación solo revisa los cambios en la reunión de consenso.',
      D(1, 6, 2027), D(30, 7, 2027),
      'Los temporales salen el día del cierre sin cálculo manual.',
      'Temporales publicados el día del cierre 3 meses seguidos sin intervención manual.',
      cajas=[('Temporales → Reserva de Ciclo Por Familia', 'La hace el sistema')],
      comite='Los temporales los calcula el sistema; es un cálculo fijo, así que basta una regla y no hace falta un agente.',
      cob=1.0, exc=0.10, depende='Calendario de previsión con día de cierre', detalle='D-PL12'),
    I('P2R-02', 'Agente', 'Agente de escenarios de capacidad',
      'Simula alternativas de turnos, balanceo entre líneas y stock, y propone el plan con su impacto en horas y euros. '
      'Planificación lo aprueba y lo lleva a la revisión mensual.',
      D(1, 11, 2027), D(31, 3, 2028),
      'Cada mes hay un plan de capacidad propuesto con su coste, listo para decidir.',
      '≥ 70 % de los planes aprobados sin rehacer; horas extra no planificadas a la baja 3 meses seguidos.',
      cajas=[('Plan de Fabricación → Decisión de turnos y balanceo de materiales', 'Se simplifica')],
      comite='El plan de turnos y balanceo lo propone un agente con escenarios valorados; Planificación decide.',
      cob=0.70, exc=0.30, depende='Carga frente a capacidad automática · migración de GIO a web', detalle='D-PL07'),

    # ---------------------------------------------------------------- P2R-03 · Política de inventario por segmento
    I('P2R-03', 'Negocio', 'Criterios de segmentación y estrategia de stock',
      'Definir ABC por valor, XYZ por variabilidad y la matriz que asigna a cada segmento MTS o MTO y su stock de '
      'seguridad. Comercial valida las estrategias.',
      D(1, 3, 2027), D(31, 3, 2027),
      'Cada material tiene una estrategia y un stock objetivo justificados.',
      'Matriz aprobada por Planificación y Comercial.',
      decidir='Cortes de ABC y XYZ · nivel de servicio por segmento', detalle='D-PL09'),
    I('P2R-03', 'Negocio', 'Responsable del stock sin movimiento',
      'Revisión mensual con Comercial del producto terminado sin movimiento: se decide venderlo, reprocesarlo o '
      'depreciarlo, y se imputa a quien lo originó cuando viene de un cambio de previsión fuera de plazo.',
      D(1, 4, 2027), D(31, 5, 2027),
      'El stock fabricado y no vendido tiene dueño y decisión.',
      '100 % del stock sin movimiento > 6 meses con decisión y responsable.',
      depende='Reunión mensual de consenso de demanda', detalle='D-PL09'),
    I('P2R-03', 'Automatización', 'ABC y estrategias recalculados cada trimestre',
      'Ejecutar la transacción de ABC que ya tiene GIO y programar el recálculo trimestral, avisando de los materiales '
      'que cambian de segmento para que Planificación los revise.',
      D(1, 3, 2027), D(30, 4, 2027),
      'El ABC y las estrategias están siempre al día sin trabajo manual.',
      'Recálculo trimestral sin intervención y 100 % de materiales con segmento vigente.',
      comite='No cambia ninguna caja: actualiza el dato que usan el plan de fabricación y la propuesta.',
      cob=1.0, exc=0.10, depende='Criterios de segmentación y estrategia de stock', detalle='D-PL09'),
    I('P2R-03', 'Agente', 'Agente de stock por segmento',
      'Cada mes revisa la cobertura de cada segmento, detecta exceso, obsoleto y riesgo de rotura en MTS, y propone la '
      'acción: ajustar el stock de seguridad, cambiar la estrategia o liquidar. Planificación aprueba.',
      D(1, 1, 2028), D(30, 4, 2028),
      'El stock se corrige antes de que sobre o falte.',
      '€ de stock fuera de política −20 % en 6 meses y 0 roturas en materiales A-MTS.',
      cajas=[('Revisión mensual de stock por segmento', NUEVA)],
      comite='Nueva tarea mensual en S&OP y Plan Maestro: un agente revisa el stock por segmento y propone acciones.',
      cob=0.80, exc=0.20, depende='ABC y estrategias recalculados cada trimestre', detalle='D-PL09'),

    # ---------------------------------------------------------------- P2R-04 · Promesa de fecha fiable
    I('P2R-04', 'Negocio', 'Política de fecha confirmada',
      'Cada pedido lleva fecha solicitada y fecha confirmada. Planificación confirma en un plazo fijo y, si sabe que no '
      'se llega, Comercial avisa al cliente con la nueva fecha; el aviso lo ejecuta O2C.',
      D(1, 3, 2027), D(30, 4, 2027),
      'La fecha que tiene el cliente es una fecha que Planificación puede cumplir.',
      '≥ 90 % de los pedidos con fecha confirmada dentro del plazo.',
      decidir='Plazo de confirmación · plazo de aviso al cliente', detalle='D-PL11'),
    I('P2R-04', 'Automatización', 'Propuesta de fecha con temporales y capacidad libre',
      'Al entrar el pedido, el sistema propone la fecha confirmada a partir de los temporales y la capacidad libre. '
      'Planificación solo revisa las que no encajan.',
      D(1, 11, 2027), D(29, 2, 2028),
      'La fecha se propone con datos, no se teclea la que pide el cliente.',
      '≥ 85 % de las fechas aceptadas sin cambios y OTIF a fecha confirmada en objetivo.',
      cajas=[('Cartera Pedidos', 'Se simplifica')],
      comite='En la cartera, la fecha la propone el sistema con la disponibilidad real.',
      cob=0.85, exc=0.15, depende='Temporales automáticos · migración de GIO a web', detalle='D-PL10'),
    I('P2R-04', 'Consulta', 'Copiloto de plazos para Comercial',
      'Comercial pregunta «¿cuándo puedo servir X?» y el copiloto responde con temporales, capacidad libre y cartera, sin '
      'llamar a Planificación. Solo lee datos.',
      D(1, 1, 2028), D(31, 3, 2028),
      'Comercial resuelve sola la mayoría de consultas de plazo.',
      '−70 % de llamadas de plazo a Planificación.',
      cajas=[('Cartera Pedidos', 'Se simplifica')],
      comite='Las consultas de plazo se resuelven con un copiloto; Planificación solo atiende las difíciles.',
      cob=0.80, exc=0.20, depende='Propuesta de fecha con temporales y capacidad libre', detalle='D-PL25'),
    I('P2R-04', 'Agente', 'Agente de renegociación de fechas',
      'Cuando el programa muestra que un pedido no llega a su fecha confirmada, calcula la nueva fecha viable y prepara el '
      'aviso al cliente para que Comercial lo revise y lo envíe por O2C.',
      D(1, 4, 2028), D(30, 6, 2028),
      'Ningún pedido llega tarde sin que el cliente lo sepa antes.',
      '≥ 90 % de los pedidos que no llegan avisados antes de su fecha.',
      cajas=[('Ajustes Planificación y Pedidos', 'Se simplifica')],
      comite='Cuando un pedido no llega, el agente propone la nueva fecha y el aviso; Comercial lo envía.',
      cob=0.80, exc=0.20, depende='Política de fecha confirmada · Cuadro de mando de servicio con alertas',
      detalle='D-PL11'),

    # ---------------------------------------------------------------- P2R-06 · Programación estable y optimizada
    I('P2R-06', 'Negocio', 'Zona congelada y reglas de prioridad',
      'Fijar una zona congelada en el programa y escribir las reglas de prioridad (cliente, fecha confirmada, margen y '
      'penalización). Si no hay acuerdo, decide la dirección de planta aplicando esas reglas, no por teléfono.',
      D(1, 4, 2027), D(28, 5, 2027),
      'Los cambios de última hora pasan por reglas conocidas.',
      'Reglas publicadas y aplicadas en el 100 % de las urgencias.',
      decidir='Horas de la zona congelada (48–72 h) · árbitro', detalle='D-PL08, D-PL20'),
    I('P2R-06', 'Negocio', 'Matriz de tiempos de cambio',
      'Medir con Producción los tiempos de cambio de molde y de altura entre referencias de cada línea. El dato se '
      'mantiene en R2M y es la base del optimizador.',
      D(1, 4, 2027), D(28, 5, 2027),
      'Se sabe cuánto cuesta cada cambio y se puede optimizar.',
      'Matriz completa para el 100 % de las líneas.',
      detalle='D-PL19 · maestro X03 (R2M)'),
    I('P2R-06', 'Automatización', 'Control de la zona congelada en GIO',
      'GIO bloquea los cambios dentro de la zona congelada salvo aprobación, y registra quién los pide y por qué.',
      D(1, 6, 2027), D(30, 7, 2027),
      'La zona congelada se respeta sin depender de la memoria de nadie.',
      '≤ 2 cambios aprobados por semana dentro de la zona congelada.',
      cajas=[('¿Cambios de última hora?', 'La hace el sistema')],
      comite='La decisión «¿cambio de última hora?» la controla el sistema: dentro de la zona congelada solo con aprobación.',
      cob=1.0, exc=0.10, depende='Zona congelada y reglas de prioridad', detalle='D-PL08'),
    I('P2R-06', 'Automatización', 'Optimizador de secuencia',
      'Propone la secuencia por línea minimizando los cambios de molde y altura y respetando las fechas confirmadas; el '
      'programador la revisa y la aprueba. Las urgencias se recalculan con las reglas de prioridad.',
      D(1, 11, 2027), D(31, 3, 2028),
      'La secuencia deja de hacerse a mano y la resecuenciación manual desaparece.',
      '−20 % de horas de cambio y adherencia al programa ≥ 90 %.',
      cajas=[('Secuencia Por Línea → Ciclo · Molde · Altura', 'Se simplifica'),
             ('Recuenciar urgencias y prioridades', 'Desaparece')],
      comite='La secuencia la propone un optimizador y el programador la aprueba; la resecuenciación manual desaparece.',
      cob=0.85, exc=0.15, depende='Zona congelada y reglas de prioridad · Matriz de tiempos de cambio · migración de GIO a web',
      detalle='D-PL19'),
    I('P2R-06', 'Agente', 'Agente de urgencias',
      'Ante una petición de cambio, calcula el impacto (qué pedidos se retrasan y cuántos cambios extra supone), aplica '
      'las reglas de prioridad y propone la decisión. Si hay conflicto, la escala al árbitro con el impacto calculado.',
      D(1, 4, 2028), D(30, 6, 2028),
      'Cada urgencia se decide con su impacto delante.',
      '100 % de las urgencias con impacto calculado y −50 % de escalados.',
      cajas=[('¿Cambios de última hora?', 'Se simplifica')],
      comite='Las urgencias las evalúa un agente con su impacto; solo los conflictos llegan a la dirección de planta.',
      cob=0.75, exc=0.25, depende='Optimizador de secuencia', detalle='D-PL20'),

    # ---------------------------------------------------------------- P2R-08 · Control del servicio y mejora continua
    I('P2R-08', 'Negocio', 'KPI de servicio y de adherencia',
      'Definir el OTIF contra fecha confirmada y la adherencia al programa (fórmula, fuente y frecuencia) y revisarlos '
      'cada semana con Comercial.',
      D(1, 3, 2027), D(30, 4, 2027),
      'Se sabe qué % del servicio se cumple y contra qué fecha.',
      'KPIs publicados cada semana desde mayo de 2027.',
      depende='Política de fecha confirmada', decidir='Objetivo de OTIF', detalle='D-PL23'),
    I('P2R-08', 'Negocio', 'Revisión mensual de causas de atraso',
      'Cada atraso se clasifica con una lista cerrada de causas (previsión, capacidad, material, secuencia o cambio del '
      'cliente). Cada mes se revisan las principales con una acción y un responsable, y lo aprendido vuelve a la '
      'previsión o a los parámetros.',
      D(3, 5, 2027), D(30, 6, 2027),
      'Los atrasos dejan de repetirse por la misma causa.',
      '≥ 90 % de los atrasos con causa y cada causa principal con acción.',
      depende='KPI de servicio y de adherencia', detalle='D-PL24'),
    I('P2R-08', 'Automatización', 'Cuadro de mando de servicio con alertas',
      'OTIF, atrasos y pedidos en riesgo se calculan solos cada día y sustituyen al reporte diario de cargas atrasadas. '
      'Avisa a Comercial de los pedidos que van a llegar tarde antes de que ocurra.',
      D(1, 4, 2027), D(30, 6, 2027),
      'Los atrasos se ven antes de que pasen y el reporte manual desaparece.',
      'Reporte diario manual eliminado y ≥ 80 % de los atrasos avisados antes de la fecha.',
      cajas=[('Seguimiento Diario Cumplimiento Servicio', 'La hace el sistema'), ('¿Hay Atrasos?', 'La hace el sistema')],
      comite='El seguimiento diario lo hace el sistema y avisa de los pedidos en riesgo antes del atraso.',
      cob=1.0, exc=0.10, depende='KPI de servicio y de adherencia', detalle='D-PL23'),
    I('P2R-08', 'Consulta', 'Copiloto «¿por qué va tarde el pedido X?»',
      'Responde a Comercial y a Planificación el estado y la causa de un pedido con los datos del sistema, sin cambiar '
      'nada.',
      D(1, 11, 2027), D(31, 1, 2028),
      'Las preguntas sobre pedidos se resuelven sin buscar en varias pantallas ni llamar.',
      '≥ 80 % de las consultas resueltas sin intervención de Planificación.',
      cajas=[('Seguimiento Diario Cumplimiento Servicio', 'Se simplifica')],
      comite='Un copiloto responde por qué va tarde un pedido; Planificación solo atiende los casos raros.',
      cob=0.80, exc=0.20, depende='Cuadro de mando de servicio con alertas', detalle='D-PL25'),
    I('P2R-08', 'Agente', 'Agente de causa raíz',
      'Clasifica cada atraso por causa, propone la acción y devuelve lo aprendido a la previsión y a los parámetros (por '
      'ejemplo, ajustar un tiempo de ciclo o la previsión de una familia). Planificación aprueba.',
      D(1, 1, 2028), D(30, 6, 2028),
      'El ciclo aprende solo: cada atraso deja una corrección propuesta.',
      '≥ 90 % de los atrasos con causa asignada por el agente y causas repetidas −30 % en 6 meses.',
      cajas=[('Ajustes Planificación y Pedidos', 'Se simplifica')],
      comite='Los ajustes tras un atraso los prepara un agente con su causa; Planificación los aprueba.',
      cob=0.75, exc=0.25, depende='Revisión mensual de causas de atraso · Cuadro de mando de servicio con alertas',
      detalle='D-PL24'),
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


def cabecera(ws, fila, cabeceras, anchos):
    for j, (h, w) in enumerate(zip(cabeceras, anchos), 1):
        c = ws.cell(fila, j, h)
        c.font, c.fill, c.alignment, c.border = Font(name=F, size=10, bold=True, color='FFFFFF'), fill(NAVY), CENTER, BOX
        ws.column_dimensions[c.column_letter].width = w


def celda(ws, r, j, v, centrada=False, fmt=None, bold=False):
    c = ws.cell(r, j, v)
    c.font, c.border = Font(name=F, size=10, color=NAVY, bold=bold), BOX
    c.alignment = CENTER if centrada else WRAP
    if fmt:
        c.number_format = fmt
    return c


def construir(salida, empresa='Ondupet', palancas=PALANCAS, cajas_flujo=CAJAS, intervenciones=INTERVENCIONES,
              fase_nuevas=None, segunda=None):
    """segunda = (planta, meses): añade fechas desplazadas para una segunda planta con su propia celda."""
    wb = Workbook()
    wi = wb.active
    wi.title = 'Intervenciones'
    wp = wb.create_sheet('Por palanca')
    wc = wb.create_sheet('Cajas del flujo')

    nombre_pal = {p[0]: p[1] for p in palancas}
    hall_pal = {p[0]: p[2] for p in palancas}
    fase_nuevas = fase_nuevas or {}
    cont = {}

    # --- Intervenciones (orden del formulario) ---
    titulo(wi, f'Intervenciones propuestas · P2R × {empresa}',
           'Una fila por intervención, con los campos en el orden del formulario «Nueva intervención». '
           'Responsable: lo trae la palanca. Confianza: EST. Las fechas siguen el plan de despliegue 2027–2028.' +
           (f' Las dos últimas columnas desplazan las fechas {segunda[1]} meses para {segunda[0]}, si es una celda aparte.' if segunda else ''))
    cab = ['ID', 'Palanca origen', 'Hallazgos que cura', 'De qué es este trabajo', 'Nombre', 'Qué hay que hacer',
           'Cambia tareas del flujo', 'Cajas del flujo → qué le pasa', 'Qué queremos cambiar (para el comité)',
           'Cobertura objetivo %', 'Excepción humana objetivo %', 'Resultado esperado', 'Criterio de éxito',
           'Empieza', 'Termina (go-live)', 'Confianza', 'Depende de', 'Lo que falta decidir', 'Detalle del maestro']
    anchos = [8, 30, 18, 15, 30, 56, 10, 40, 46, 11, 11, 36, 40, 12, 12, 10, 34, 30, 16]
    if segunda:
        cab += [f'Empieza · {segunda[0]}', f'Termina · {segunda[0]}']
        anchos += [13, 13]
    cabecera(wi, 4, cab, anchos)
    for k, it in enumerate(intervenciones):
        r = 5 + k
        n = cont[it['pal']] = cont.get(it['pal'], 0) + 1
        iid = f"{it['pal'].split('-')[1]}.{n}"
        cajas = '\n'.join(f'{c} → {q}' for c, q in it['cajas'])
        vals = [iid, f"{it['pal']} · {nombre_pal[it['pal']]}", hall_pal[it['pal']], it['tipo'], it['nombre'], it['que'],
                'Sí' if it['cajas'] else 'No', cajas or None, it['comite'] or None, it['cob'], it['exc'], it['res'],
                it['crit'], it['ini'], it['fin'], 'EST', it['dep'] or None, it['dec'] or None, it['det'] or None]
        if segunda:
            vals += [f'=EDATE(N{r},{segunda[1]})', f'=EDATE(O{r},{segunda[1]})']
        for j, v in enumerate(vals, 1):
            fmt = '0%' if j in (10, 11) else 'dd/mm/yyyy' if j in (14, 15, 20, 21) else None
            celda(wi, r, j, v, centrada=j in (1, 4, 7, 10, 11, 14, 15, 16, 20, 21), fmt=fmt)
        wi.cell(r, 4).fill = fill(COLOR_TIPO[it['tipo']])
    i_ult = 4 + len(intervenciones)
    wi.freeze_panes = 'F5'
    wi.auto_filter.ref = f'A4:{wi.cell(4, len(cab)).column_letter}{i_ult}'

    # --- Por palanca (fórmulas sobre Intervenciones) ---
    titulo(wp, 'Resumen por palanca', 'Se recalcula al cambiar la hoja Intervenciones.')
    tipos = ['Negocio', 'Automatización', 'Agente', 'Consulta']
    cabecera(wp, 4, ['Palanca', 'Nombre', 'Hallazgos', 'Intervenciones', *tipos, 'Empieza', 'Último go-live'],
             [10, 36, 22, 13, 11, 15, 10, 10, 12, 14])
    rng = lambda col: f'Intervenciones!${col}$5:${col}${i_ult}'
    for k, (pid, nom, hs) in enumerate(palancas):
        r = 5 + k
        crit = f'"{pid} · {nom}"'
        vals = [pid, nom, hs, f'=COUNTIF({rng("B")},{crit})',
                *[f'=COUNTIFS({rng("B")},{crit},{rng("D")},"{t}")' for t in tipos],
                f'=_xlfn.MINIFS({rng("N")},{rng("B")},{crit})', f'=_xlfn.MAXIFS({rng("O")},{rng("B")},{crit})']
        for j, v in enumerate(vals, 1):
            celda(wp, r, j, v, centrada=j != 2 and j != 3, fmt='dd/mm/yyyy' if j >= 9 else None)
    t = 5 + len(palancas)
    celda(wp, t, 1, 'Total', bold=True)
    for j in range(4, 9):
        col = wp.cell(4, j).column_letter
        celda(wp, t, j, f'=SUM({col}5:{col}{t - 1})', centrada=True, bold=True)
    wp.freeze_panes = 'A5'

    # --- Cajas del flujo: qué intervención toca cada caja ---
    titulo(wc, f'Cajas del flujo as-is de {empresa} que cambian',
           'Lo que el to-be pintará en cada caja. «Nueva tarea» son tareas que añade una intervención.')
    cabecera(wc, 4, ['Fase', 'Caja', 'Nº intervenciones', 'Intervenciones y qué le pasa'], [30, 52, 14, 80])
    nuevas = []
    for it in intervenciones:
        for c, q in it['cajas']:
            if q == NUEVA and c not in nuevas:
                nuevas.append(c)
    todas = cajas_flujo + [(fase_nuevas.get(c, 'S&OP y Plan Maestro') + ' (nueva)', c) for c in nuevas]
    for k, (fase, caja) in enumerate(todas):
        r = 5 + k
        toca = [f"{it['nombre']} ({it['tipo']}) → {q}" for it in intervenciones for c, q in it['cajas'] if c == caja]
        celda(wc, r, 1, fase)
        celda(wc, r, 2, caja)
        celda(wc, r, 3, f'=SUMPRODUCT(--ISNUMBER(SEARCH(B{r}&" → ",Intervenciones!$H$5:$H${i_ult})))', centrada=True)
        celda(wc, r, 4, '\n'.join(toca) or None)
    c_ult = 4 + len(todas)
    wc.conditional_formatting.add(f'A5:D{c_ult}', FormulaRule(formula=['$C5=0'], font=Font(name=F, color='9AA5B0')))
    wc.freeze_panes = 'A5'

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    wb.save(salida)


if __name__ == '__main__':
    out = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).resolve().parents[1] / 'Intervenciones_P2R_Ondupet.xlsx')
    construir(out)
    print(out)
