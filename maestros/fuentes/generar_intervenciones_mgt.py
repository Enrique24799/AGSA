"""Genera las intervenciones propuestas para las palancas de P2R × MGT (Matías Gomá Tomás).

Mismo formato que las de Ondupet y Ondupack. Las cajas son las del flujo as-is de MGT; si al dibujarlo en la
aplicación cambia algún nombre, hay que cambiarlo aquí también.

Uso: python generar_intervenciones_mgt.py [salida.xlsx]
"""
import sys
from pathlib import Path

from generar_intervenciones_ondupet import D, I, NUEVA, construir

PALANCAS = [
    ('P2R-01', 'Gestión de la demanda', 'HT-P2R-01'),
    ('P2R-02', 'S&OP y plan de capacidad', 'HT-P2R-03'),
    ('P2R-03', 'Política de inventario por segmento', 'HT-P2R-04'),
    ('P2R-04', 'Promesa de fecha fiable', 'HT-P2R-05'),
    ('P2R-05', 'Planificación de materiales por excepción', 'HT-P2R-07'),
    ('P2R-06', 'Programación estable y optimizada', 'HT-P2R-08, HT-P2R-09'),
    ('P2R-07', 'Liberación controlada', 'HT-P2R-10'),
    ('P2R-08', 'Control del servicio y mejora continua', 'HT-P2R-11, HT-P2R-12'),
]

CAJAS = [
    ('Demanda y previsión', 'Pedido intercompany del próximo ciclo'),
    ('Demanda y previsión', 'Encajar la urgencia según la prioridad de Ondupack'),
    ('Demanda y previsión', 'Pedido de cliente externo'),
    ('Demanda y previsión', 'Revisar capacidad para la fecha pedida'),
    ('S&OP y Plan Maestro', 'Ciclo de producción de las 4 calidades'),
    ('S&OP y Plan Maestro', 'Duración por calidad (temporales)'),
    ('S&OP y Plan Maestro', 'Carga a capacidad infinita · 3 turnos a velocidad máxima'),
    ('S&OP y Plan Maestro', '¿Pedidos sin capacidad para su fecha?'),
    ('S&OP y Plan Maestro', 'Revisar con Comercial y avisar al cliente'),
    ('Necesidades y Aprovisionamiento', 'Propuesta de fabricación de PT por pedido'),
    ('Necesidades y Aprovisionamiento', 'Bobina madre por gramaje y calidad'),
    ('Necesidades y Aprovisionamiento', 'MRP de auxiliares en SAP'),
    ('Necesidades y Aprovisionamiento', 'MRP de materias primas en Excel'),
    ('Necesidades y Aprovisionamiento', 'Solicitud de compra y pedido'),
    ('Programación', 'Algoritmo de combinación por ciclos'),
    ('Programación', 'Agrupar por calidad y ancho'),
    ('Programación', '¿Entra una urgencia?'),
    ('Programación', 'Encajar la urgencia a mano'),
    ('Liberación', 'Comprobar a mano que hay material'),
    ('Liberación', '¿Material disponible?'),
    ('Liberación', 'Liberar e imprimir la orden'),
    ('Liberación', 'Seguir la llegada del material'),
    ('Seguimiento y Control', 'Revisar incidencias por llamada'),
    ('Seguimiento y Control', 'Enviar el stock a Ondupack cada día'),
    ('Seguimiento y Control', 'Ondupack decide qué se carga al día siguiente'),
    ('Seguimiento y Control', 'Cierre manual del pedido'),
]
FASE_NUEVAS = {'Previsión del próximo ciclo': 'Demanda y previsión',
               'Revisión mensual del stock intercompany': 'S&OP y Plan Maestro'}

INTERVENCIONES = [
    # ---------------------------------------------------------------- P2R-01 · Gestión de la demanda
    I('P2R-01', 'Negocio', 'Previsión de papel de Ondupack por ciclo',
      'Acordar con Ondupack que, en cada ciclo, envíe la previsión de papel de los ciclos siguientes por calidad y ancho, '
      'además del pedido firme 10 días antes. Es la otra mitad de la intervención «Previsión de papel compartida con '
      'MGT/PDA» de Ondupack.',
      D(4, 10, 2027), D(17, 12, 2027),
      'MGT planifica sus ciclos con lo que va a pedir Ondupack, no solo con lo que ya ha pedido.',
      'Previsión recibida en el 100 % de los ciclos y desviación pedido–previsión < 15 %.',
      decidir='Horizonte de la previsión (2–3 ciclos) · formato', detalle='D-PL13'),
    I('P2R-01', 'Negocio', 'Previsión de los clientes externos repetitivos',
      'Identificar con Comercial los clientes externos con consumo repetitivo y acordar con ellos una previsión no '
      'vinculante por calidad y ancho, revisada en cada ciclo.',
      D(2, 11, 2027), D(31, 1, 2028),
      'La demanda externa repetitiva entra en el ciclo antes de que llegue el pedido.',
      'Previsión de ≥ 60 % del volumen externo repetitivo en cada ciclo.',
      decidir='Criterio de cliente repetitivo · quién la recoge (Comercial o Joan Carles)', detalle='D-PL03'),
    I('P2R-01', 'Negocio', 'Registro de las ventas perdidas',
      'Definir qué es una venta perdida (pedido rechazado, cantidad reducida o fecha no aceptada por capacidad) y '
      'registrarla con su motivo. Se captura en Salesforce, que es O2C; aquí se usa.',
      D(4, 10, 2027), D(30, 11, 2027),
      'Se sabe cuánta demanda se pierde y por qué.',
      '100 % de los pedidos rechazados o reducidos con motivo.',
      detalle='D-PL04'),
    I('P2R-01', 'Automatización', 'Previsión del ciclo cargada en GIO',
      'La previsión de Ondupack y de los clientes repetitivos se carga en GIO para que la planificación del ciclo use '
      'cartera más previsión; cada pedido que llega consume previsión.',
      D(1, 12, 2027), D(25, 2, 2028),
      'El ciclo deja de planificarse solo con la cartera.',
      'Ciclos planificados con cartera y previsión desde marzo de 2028.',
      cajas=[('Previsión del próximo ciclo', NUEVA), ('Ciclo de producción de las 4 calidades', 'Se simplifica')],
      comite='Nueva tarea de previsión en Demanda; el ciclo de las 4 calidades se planifica con cartera más previsión.',
      cob=1.0, exc=0.10,
      depende='Previsión de papel de Ondupack por ciclo · Previsión de los clientes externos repetitivos', detalle='D-PL01'),
    I('P2R-01', 'Agente', 'Agente de previsión por calidad y ancho',
      'Propone la previsión de los próximos ciclos por calidad y ancho a partir de la previsión de Ondupack, el histórico '
      'de pedidos y las ventas perdidas. Planificación la valida.',
      D(1, 6, 2028), D(22, 12, 2028),
      'La previsión de cada ciclo sale preparada y Planificación solo revisa las excepciones.',
      'Precisión ≥ 80 % por calidad y ≥ 80 % de las líneas sin ajuste manual.',
      cajas=[('Previsión del próximo ciclo', 'Se simplifica')],
      comite='La previsión de cada ciclo la prepara un agente; Planificación revisa las excepciones.',
      cob=0.80, exc=0.20, depende='Previsión del ciclo cargada en GIO · Registro de las ventas perdidas (6 meses)',
      detalle='D-PL01'),

    # ---------------------------------------------------------------- P2R-02 · S&OP y plan de capacidad
    I('P2R-02', 'Negocio', 'Política de llenado de capacidad',
      'Decidir con Dirección qué se fabrica cuando sobra capacidad (referencias repetitivas con mínimo y máximo, sobre todo '
      'las de Ondupack) y cuándo compensa parar, comparando el coste de parada con el coste del stock. Sustituye a la '
      'regla de «3 turnos siempre».',
      D(4, 10, 2027), D(17, 12, 2027),
      'El stock que se genera para llenar la máquina es una decisión económica, no una consecuencia.',
      'Política aprobada y € de stock sin pedido dentro del límite acordado.',
      decidir='Coste de parada frente a coste de stock · referencias con las que se llena', detalle='D-PL06'),
    I('P2R-02', 'Negocio', 'Revisión mensual de capacidad y ciclos',
      'Reunión mensual de Joan Carles con Comercial (Albert) y Dirección: carga frente a capacidad de los próximos ciclos, '
      'duración de cada calidad y pedidos que no caben.',
      D(2, 11, 2027), D(31, 1, 2028),
      'Los pedidos que no caben se ven y se hablan con el cliente antes de lanzar el ciclo.',
      'Decisión de ciclos documentada cada mes y 0 pedidos sin capacidad detectados después de lanzar el ciclo.',
      depende='Política de llenado de capacidad', detalle='D-PL06'),
    I('P2R-02', 'Automatización', 'Carga a capacidad finita por ciclo',
      'GIO carga los pedidos a capacidad finita por ciclo y calidad y marca los que no caben al entrar, en lugar de cargar '
      'a capacidad infinita y revisarlos después.',
      D(1, 12, 2027), D(31, 3, 2028),
      'Se sabe al entrar el pedido si cabe en el ciclo.',
      '100 % de los pedidos sin capacidad marcados al entrar.',
      cajas=[('Carga a capacidad infinita · 3 turnos a velocidad máxima', 'Se simplifica'),
             ('¿Pedidos sin capacidad para su fecha?', 'La hace el sistema')],
      comite='La carga se hace a capacidad finita y el sistema marca solo los pedidos que no caben.',
      cob=1.0, exc=0.15, depende='Política de llenado de capacidad', detalle='D-PL07'),
    I('P2R-02', 'Automatización', 'Propuesta de duración de los ciclos por calidad',
      'El sistema propone la duración de cada calidad en el próximo ciclo a partir de la cartera y la previsión. Joan '
      'Carles la ajusta.',
      D(3, 1, 2028), D(31, 3, 2028),
      'La duración de cada calidad deja de decidirse sobre la marcha.',
      'Propuesta aceptada sin cambios en ≥ 80 % de los ciclos.',
      cajas=[('Duración por calidad (temporales)', 'La hace el sistema')],
      comite='La duración de cada calidad la propone el sistema; el planificador la ajusta.',
      cob=1.0, exc=0.20, depende='Previsión del ciclo cargada en GIO', detalle='D-PL12'),
    I('P2R-02', 'Agente', 'Agente de escenarios de ciclo',
      'Simula alternativas de duración y orden de las calidades, paradas y stock, y propone el ciclo con su impacto en '
      'coste y servicio. Planificación lo aprueba en la revisión mensual.',
      D(1, 6, 2028), D(22, 12, 2028),
      'Cada ciclo llega con un plan propuesto y valorado.',
      '≥ 70 % de los planes aprobados sin rehacer.',
      cajas=[('Ciclo de producción de las 4 calidades', 'Se simplifica')],
      comite='El ciclo de las 4 calidades lo propone un agente con escenarios valorados; Planificación decide.',
      cob=0.70, exc=0.30, depende='Carga a capacidad finita · Propuesta de duración de los ciclos', detalle='D-PL07'),

    # ---------------------------------------------------------------- P2R-03 · Política de inventario por segmento
    I('P2R-03', 'Negocio', 'Pull system con Ondupack',
      'Definir con Ondupack qué referencias se reponen por pull, con un mínimo y un máximo en MGT. Ondupack retira cada día '
      'de ese stock y MGT lo repone al combinar sus ciclos.',
      D(4, 10, 2027), D(17, 12, 2027),
      'El stock de Ondupack en MGT tiene límites y deja de crecer sin control.',
      'Referencias pull con mínimo y máximo aprobados y stock intercompany dentro del máximo.',
      decidir='Referencias, mínimos y máximos · de quién es el stock', detalle='D-PL09, D-PL13'),
    I('P2R-03', 'Negocio', 'Ajuste del pedido intercompany a lo fabricado',
      'Cuando se fabrica más que el pedido de Ondupack dentro de una tolerancia, el pedido se modifica para que se lleve '
      'todo lo fabricado. La modificación se hace en O2C.',
      D(4, 10, 2027), D(30, 11, 2027),
      'Lo fabricado de más deja de quedarse en el almacén de MGT.',
      '0 excedentes de pedidos intercompany sin asignar a más de 30 días.',
      decidir='Tolerancia de sobrefabricación', detalle='D-PL09'),
    I('P2R-03', 'Automatización', 'Reposición automática por mínimo y máximo',
      'GIO genera la necesidad de reposición de las referencias pull cuando el stock baja del mínimo, y la mete en la '
      'cartera del próximo ciclo.',
      D(3, 1, 2028), D(31, 3, 2028),
      'La reposición de las referencias pull no depende de un pedido manual.',
      '100 % de las roturas de mínimo con reposición en el siguiente ciclo.',
      cajas=[('Propuesta de fabricación de PT por pedido', 'Se simplifica')],
      comite='Además de por pedido, el sistema propone fabricar para reponer las referencias pull.',
      cob=1.0, exc=0.10, depende='Pull system con Ondupack', detalle='D-PL09'),
    I('P2R-03', 'Agente', 'Agente de stock intercompany',
      'Cada mes revisa la cobertura de las referencias de Ondupack en MGT, propone ajustar mínimos y máximos y avisa del '
      'stock que se queda sin salida.',
      D(1, 6, 2028), D(22, 12, 2028),
      'El stock intercompany se ajusta antes de que envejezca.',
      '€ de stock intercompany −20 % en 6 meses sin roturas en referencias pull.',
      cajas=[('Revisión mensual del stock intercompany', NUEVA)],
      comite='Nueva tarea mensual en S&OP y Plan Maestro: un agente revisa el stock intercompany y propone ajustes.',
      cob=0.80, exc=0.20, depende='Reposición automática por mínimo y máximo', detalle='D-PL09'),

    # ---------------------------------------------------------------- P2R-04 · Promesa de fecha fiable
    I('P2R-04', 'Negocio', 'Fecha confirmada por Planificación',
      'Cada pedido externo recibe una fecha confirmada por Planificación en un plazo fijo, distinta de la solicitada. Los '
      'pedidos de Ondupack llevan una fecha firme por ciclo, y cambiarla dentro del ciclo requiere acuerdo.',
      D(4, 10, 2027), D(30, 11, 2027),
      'Las fechas de los pedidos, también las de Ondupack, son fechas que se pueden cumplir.',
      '≥ 90 % de los pedidos con fecha confirmada en plazo.',
      decidir='Plazo de confirmación · reglas de cambio para Ondupack', detalle='D-PL11'),
    I('P2R-04', 'Automatización', 'Propuesta de fecha con la capacidad libre del ciclo',
      'Al entrar el pedido, el sistema propone la primera fecha posible según el ciclo de su calidad y la capacidad libre. '
      'Es el CTP que ya usa Ondupack en GIO, adaptado a ciclos de papel.',
      D(1, 2, 2028), D(31, 5, 2028),
      'La fecha se propone con datos en lugar de revisarse a mano.',
      '≥ 85 % de las fechas propuestas aceptadas sin cambios.',
      cajas=[('Revisar capacidad para la fecha pedida', 'La hace el sistema')],
      comite='La fecha la propone el sistema con la capacidad del ciclo, como ya se hace en Ondupack.',
      cob=0.85, exc=0.15, depende='Carga a capacidad finita por ciclo', detalle='D-PL10'),
    I('P2R-04', 'Consulta', 'Copiloto de plazos para Comercial',
      'Comercial pregunta «¿cuándo sale la próxima tirada de esta calidad y ancho?» y el copiloto responde con el plan de '
      'ciclos y la capacidad libre, sin llamar a Planificación.',
      D(1, 6, 2028), D(29, 9, 2028),
      'Comercial resuelve sola las consultas de plazo.',
      '−60 % de consultas de plazo a Planificación.',
      cajas=[('Revisar capacidad para la fecha pedida', 'Se simplifica')],
      comite='Las consultas de plazo se resuelven con un copiloto.',
      cob=0.80, exc=0.20, depende='Propuesta de fecha con la capacidad libre del ciclo', detalle='D-PL25'),

    # ---------------------------------------------------------------- P2R-05 · Materiales por excepción
    I('P2R-05', 'Negocio', 'Planificación de la bobina madre',
      'Levantar y documentar cómo se planifica hoy la bobina madre por gramaje y calidad, quién lo hace y con qué dato.',
      D(4, 10, 2027), D(29, 10, 2027),
      'La planificación del semielaborado deja de ser una caja negra.',
      'Proceso documentado y con responsable.',
      decidir='Quién planifica la bobina madre', detalle='D-PL17'),
    I('P2R-05', 'Negocio', 'Entradas de báscula registradas en el momento',
      'Asignar quién registra en el sistema cada entrada de báscula en el momento, para que el consumo no llegue antes que '
      'la entrada y el stock no se descuadre. El registro es del proceso de inventario; aquí se usa.',
      D(4, 10, 2027), D(30, 11, 2027),
      'El stock de auxiliares y materias primas es fiable sin cuadrarlo dos veces al mes.',
      'Diferencias de stock < 2 % sin cuadre manual.',
      detalle='Proceso de inventario y almacén'),
    I('P2R-05', 'Automatización', 'MRP de materias primas en SAP',
      'Pasar el MRP de materias primas del Excel a SAP, como ya se hace con los auxiliares, alimentado con el plan de '
      'ciclos y con lead time, lote mínimo y stock de seguridad completos.',
      D(1, 12, 2027), D(31, 3, 2028),
      'Las necesidades de materias primas salen del sistema y no de una hoja personal.',
      '100 % de las materias primas planificadas en SAP.',
      cajas=[('MRP de materias primas en Excel', 'La hace el sistema')],
      comite='El MRP de materias primas lo calcula SAP con el plan de ciclos; desaparece el Excel.',
      cob=0.95, exc=0.05, depende='Entradas de báscula registradas en el momento', detalle='D-PL15, D-PL16'),
    I('P2R-05', 'Agente', 'Agente de compras de materiales',
      'Revisa cada día los mensajes del MRP, propone las solicitudes de compra y avisa cuando el consumo se separa del '
      'plan. Marta Boyo y Joan Carles aprueban; el pedido sigue en P2P.',
      D(1, 6, 2028), D(22, 12, 2028),
      'Las solicitudes de compra salen preparadas y solo se revisan las excepciones.',
      '≥ 80 % de las solicitudes propuestas aceptadas sin cambios.',
      cajas=[('Solicitud de compra y pedido', 'Se simplifica')],
      comite='Las solicitudes de compra las prepara un agente a partir del MRP; las personas aprueban.',
      cob=0.80, exc=0.20, depende='MRP de materias primas en SAP', detalle='D-PL14'),

    # ---------------------------------------------------------------- P2R-06 · Programación estable y optimizada
    I('P2R-06', 'Automatización', 'Cierre de SIGPA (algoritmo de combinación)',
      'Cerrar las pruebas del algoritmo de combinación en MGT y PDA, como ya prevé el plan de sistemas.',
      D(11, 1, 2027), D(31, 3, 2027),
      'El algoritmo está validado y se puede usar como herramienta estándar.',
      'Pruebas cerradas y algoritmo en producción.',
      cajas=[('Algoritmo de combinación por ciclos', 'Se simplifica')],
      comite='El algoritmo de combinación queda validado para usarse en cada ciclo.',
      cob=0.80, exc=0.20, detalle='D-PL18'),
    I('P2R-06', 'Automatización', 'Algoritmo con lo ya fabricado y fecha límite de cartera',
      'El algoritmo tiene en cuenta lo ya fabricado y el stock pull, y una fecha límite de cartera por ciclo. Su salida pasa '
      'a ser el estándar y la agrupación manual por calidad y ancho desaparece.',
      D(2, 11, 2027), D(25, 2, 2028),
      'La combinación sale del algoritmo y no de una agrupación manual.',
      '≥ 90 % de los ciclos con la combinación del algoritmo y merma de recorte a la baja.',
      cajas=[('Algoritmo de combinación por ciclos', 'Se simplifica'), ('Agrupar por calidad y ancho', 'Desaparece')],
      comite='La combinación la hace el algoritmo con lo ya fabricado; la agrupación manual desaparece.',
      cob=0.90, exc=0.10, depende='Cierre de SIGPA · Pull system con Ondupack', detalle='D-PL18, D-PL19'),
    I('P2R-06', 'Negocio', 'Zona congelada del ciclo y reglas de urgencia',
      'Fijar la fecha límite de cartera de cada ciclo (para Ondupack ya son 10 días) y una zona congelada. Definir qué es '
      'una urgencia, también la de Ondupack fuera de plazo, y quién decide.',
      D(4, 10, 2027), D(30, 11, 2027),
      'Las urgencias entran al ciclo con reglas conocidas.',
      '≤ 2 urgencias por ciclo dentro de la zona congelada y todas aprobadas.',
      decidir='Días de la zona congelada · quién decide las urgencias', detalle='D-PL08, D-PL20'),
    I('P2R-06', 'Automatización', 'Alertas de planificación',
      'El sistema avisa de los pedidos en riesgo, del impacto de una urgencia en el ciclo y de los cambios dentro de la zona '
      'congelada. Hoy no hay ninguna alerta.',
      D(3, 1, 2028), D(31, 3, 2028),
      'Las urgencias y los riesgos se ven sin revisar el plan a mano.',
      '≥ 80 % de los atrasos avisados antes de la fecha.',
      cajas=[('Encajar la urgencia a mano', 'Se simplifica')],
      comite='El sistema avisa del impacto de cada urgencia; el planificador decide con esa información.',
      cob=1.0, exc=0.20, depende='Zona congelada del ciclo y reglas de urgencia', detalle='D-PL23'),
    I('P2R-06', 'Agente', 'Agente de urgencias',
      'Ante una urgencia, comprueba si cabe en el ciclo y qué pedidos y combinaciones se ven afectados, aplica las reglas '
      'y propone la decisión. Si hay conflicto, la escala con el impacto calculado.',
      D(1, 6, 2028), D(22, 12, 2028),
      'Cada urgencia se decide con su impacto delante.',
      '100 % de las urgencias con impacto calculado.',
      cajas=[('Encajar la urgencia a mano', 'Se simplifica')],
      comite='Las urgencias las evalúa un agente; solo los conflictos llegan a Comercial.',
      cob=0.75, exc=0.25, depende='Alertas de planificación', detalle='D-PL20'),

    # ---------------------------------------------------------------- P2R-07 · Liberación controlada
    I('P2R-07', 'Negocio', 'Criterio de liberación',
      'Definir qué tiene que estar disponible para liberar una orden (materias primas, auxiliares y bobina madre) y que la '
      'orden llegue a producción en formato digital.',
      D(4, 10, 2027), D(30, 11, 2027),
      'Se libera con un criterio conocido y no según quién lo compruebe.',
      'Criterio publicado y aplicado al 100 % de las órdenes.',
      detalle='D-PL22'),
    I('P2R-07', 'Automatización', 'Disponibilidad de material por orden',
      'El sistema muestra en cada orden si el material está y cuándo llega, a partir del MRP y de los pedidos de compra. '
      'Sustituye la comprobación manual en varios sitios.',
      D(3, 1, 2028), D(28, 4, 2028),
      'El planificador ve en una pantalla qué órdenes puede liberar.',
      'Comprobación manual eliminada y 0 órdenes liberadas sin material.',
      cajas=[('Comprobar a mano que hay material', 'La hace el sistema'), ('¿Material disponible?', 'La hace el sistema'),
             ('Seguir la llegada del material', 'Se simplifica')],
      comite='La disponibilidad de material por orden la calcula el sistema; el seguimiento de llegadas se simplifica.',
      cob=1.0, exc=0.10, depende='MRP de materias primas en SAP', detalle='D-PL22'),
    I('P2R-07', 'Automatización', 'Orden digital a producción',
      'La orden liberada llega a producción en pantalla, desde GIO, en lugar de imprimirse en papel.',
      D(1, 12, 2027), D(25, 2, 2028),
      'Producción trabaja con la orden vigente y los cambios se ven al momento.',
      '0 órdenes en papel.',
      cajas=[('Liberar e imprimir la orden', 'Se simplifica')],
      comite='La orden se libera en digital; desaparece la impresión en papel.',
      cob=1.0, exc=0.05, decidir='Pantalla en línea o tablet · qué ve producción', detalle='D-PL22'),

    # ---------------------------------------------------------------- P2R-08 · Control del servicio
    I('P2R-08', 'Negocio', 'Pilar de servicio semanal',
      'Reunión semanal fija entre MGT, Ondupack y Comercial: servicio, pedidos pendientes, atrasos y sus causas, y restos de '
      'pedidos antiguos por cerrar.',
      D(4, 10, 2027), D(30, 11, 2027),
      'Las incidencias se revisan con un día fijo y cada causa sale con una acción.',
      'Reunión celebrada el 100 % de las semanas y cada causa principal con acción.',
      depende='KPI de servicio', detalle='D-PL24'),
    I('P2R-08', 'Negocio', 'KPI de servicio',
      'Definir el OTIF contra fecha confirmada para clientes externos y para Ondupack, y publicarlo cada semana.',
      D(4, 10, 2027), D(30, 11, 2027),
      'Se sabe qué % del servicio se cumple, también hacia el grupo.',
      'KPI publicado cada semana desde diciembre de 2027.',
      depende='Fecha confirmada por Planificación', decidir='Objetivo de OTIF externo e intercompany', detalle='D-PL23'),
    I('P2R-08', 'Automatización', 'Visibilidad de stock y cargas para Ondupack',
      'Ondupack ve en línea el stock de MGT y la producción prevista, y pide las cargas desde ahí, en lugar de los unos '
      '16 correos diarios. La carga la sigue ejecutando T2D.',
      D(3, 1, 2028), D(31, 3, 2028),
      'La coordinación diaria de entregas deja de hacerse por correo.',
      '−80 % de correos de coordinación de entregas.',
      cajas=[('Enviar el stock a Ondupack cada día', 'Desaparece'),
             ('Ondupack decide qué se carga al día siguiente', 'Se simplifica')],
      comite='Ondupack consulta el stock y pide las cargas en línea; desaparece el envío diario de stock por correo.',
      cob=1.0, exc=0.10, depende='Pull system con Ondupack', detalle='D-PL13'),
    I('P2R-08', 'Automatización', 'Cuadro de mando de servicio',
      'OTIF, atrasos y pedidos en riesgo se calculan solos cada día, para externos y para Ondupack.',
      D(3, 1, 2028), D(31, 3, 2028),
      'Los atrasos se ven antes de que llegue la llamada.',
      'Cuadro de mando diario sin trabajo manual.',
      cajas=[('Revisar incidencias por llamada', 'La hace el sistema')],
      comite='El seguimiento de incidencias lo hace el sistema y avisa antes del atraso.',
      cob=1.0, exc=0.15, depende='KPI de servicio', detalle='D-PL23'),
    I('P2R-08', 'Agente', 'Agente de coordinación de cargas con Ondupack',
      'Lee las peticiones de Ondupack, comprueba stock y plan, y propone la carga del día siguiente dentro del mínimo y '
      'máximo acordados. Joan Carles aprueba.',
      D(1, 6, 2028), D(22, 12, 2028),
      'La coordinación diaria con Ondupack sale preparada y solo se revisan las excepciones.',
      '≥ 80 % de las cargas propuestas aceptadas sin cambios.',
      cajas=[('Ondupack decide qué se carga al día siguiente', 'Se simplifica')],
      comite='Las cargas diarias a Ondupack las prepara un agente; Planificación las aprueba.',
      cob=0.80, exc=0.20, depende='Visibilidad de stock y cargas para Ondupack', detalle='D-PL13'),
]

if __name__ == '__main__':
    out = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).resolve().parents[1] / 'Intervenciones_P2R_MGT.xlsx')
    construir(out, empresa='MGT', palancas=PALANCAS, cajas_flujo=CAJAS, intervenciones=INTERVENCIONES,
              fase_nuevas=FASE_NUEVAS)
    print(out)
