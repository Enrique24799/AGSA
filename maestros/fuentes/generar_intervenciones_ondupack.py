"""Genera las intervenciones propuestas para las palancas de P2R × Ondupack (Almendralejo y Navalmoral).

Mismo formato que las de Ondupet. Las cajas son las del flujo as-is de Ondupack; si al dibujarlo en la
aplicación cambia algún nombre, hay que cambiarlo aquí también.

Uso: python generar_intervenciones_ondupack.py [salida.xlsx]
"""
import sys
from pathlib import Path

from generar_intervenciones_ondupet import D, I, NUEVA, construir

PALANCAS = [
    ('P2R-01', 'Gestión de la demanda', 'HT-P2R-13'),
    ('P2R-02', 'S&OP y plan de capacidad', 'HT-P2R-03'),
    ('P2R-05', 'Planificación de materiales por excepción', 'HT-P2R-06, HT-P2R-07'),
    ('P2R-06', 'Programación estable y optimizada', 'HT-P2R-08, HT-P2R-09'),
    ('P2R-08', 'Control del servicio y mejora continua', 'HT-P2R-11, HT-P2R-12'),
    ('P2R-03', 'Política de inventario por segmento', 'HT-P2R-04 (a confirmar)'),
]

CAJAS = [
    ('Demanda y previsión', 'Alta del pedido en GIO'),
    ('Demanda y previsión', 'Simulación CTP · primera fecha posible por ruta'),
    ('Demanda y previsión', 'Ruta alternativa o negociar la fecha'),
    ('S&OP y Plan Maestro', 'Cartera planificada y no planificada'),
    ('S&OP y Plan Maestro', 'Agrupar pedidos repetitivos'),
    ('S&OP y Plan Maestro', 'Carga frente a capacidad con escenarios de 2 a 5 turnos'),
    ('S&OP y Plan Maestro', '¿Saturación > 100 %?'),
    ('S&OP y Plan Maestro', 'Replanificar ruta, fecha o turnos'),
    ('Necesidades y Aprovisionamiento', 'Balance de plancha (WIP) y de producto terminado'),
    ('Necesidades y Aprovisionamiento', 'Necesidad de papel por consumo pasado'),
    ('Necesidades y Aprovisionamiento', 'Compra de papel y auxiliares'),
    ('Programación', 'Combinación y secuencia de onduladora por ciclos de canal'),
    ('Programación', 'Secuencia de converting con escenarios SMED'),
    ('Programación', 'Sincronizar con la plancha disponible'),
    ('Programación', '¿Alerta o cambio de ruta o fecha?'),
    ('Programación', 'Replanificar ruta o fecha'),
    ('Liberación', '¿Material, cartón, ruta, máquina y capacidad OK?'),
    ('Liberación', 'Autorizar la liberación con material pendiente'),
    ('Seguimiento y Control', 'Cierre automático al terminar converting'),
    ('Seguimiento y Control', 'Seguimiento de fechas atrasadas'),
    ('Seguimiento y Control', 'Ajustes Planificación y Pedidos'),
    ('Seguimiento y Control', 'Reunión semanal de planificación'),
]
FASE_NUEVAS = {'Previsión de referencias repetitivas': 'Demanda y previsión',
               'Registro de la demanda perdida': 'Demanda y previsión'}

INTERVENCIONES = [
    # ---------------------------------------------------------------- P2R-01 · Gestión de la demanda
    I('P2R-01', 'Negocio', 'Origen de la demanda y clientes repetitivos',
      'Clasificar la demanda de los últimos 12 meses por origen (pedido firme, pedido de stock e intercompany) e '
      'identificar los clientes y referencias con consumo repetitivo, que serán los primeros en tener previsión.',
      D(1, 6, 2027), D(30, 7, 2027),
      'Se sabe qué parte de la demanda se puede prever y por dónde empezar.',
      '% de la demanda por origen publicado y lista de referencias repetitivas aprobada con Comercial.',
      decidir='Criterio de «repetitivo» (frecuencia y variabilidad)', detalle='D-PL01'),
    I('P2R-01', 'Negocio', 'Previsión no vinculante de la demanda repetitiva',
      'Acordar con Comercial, y con los clientes que lo acepten, una previsión mensual no vinculante de las '
      'referencias repetitivas, con día de cierre y una revisión mensual de consenso.',
      D(1, 7, 2027), D(30, 9, 2027),
      'La demanda repetitiva entra al plan antes de que llegue el pedido.',
      'Previsión cerrada en fecha 3 meses seguidos para ≥ 60 % del volumen repetitivo.',
      depende='Origen de la demanda y clientes repetitivos',
      decidir='Día de cierre · horizonte (3–6 meses) · qué clientes participan', detalle='D-PL02, D-PL03'),
    I('P2R-01', 'Automatización', 'Registro de la demanda perdida en GIO',
      'Cuando el cliente no acepta la fecha que propone el CTP, reduce la cantidad o se rechaza un pedido por '
      'capacidad, GIO obliga a registrar el motivo.',
      D(1, 6, 2027), D(31, 8, 2027),
      'Se sabe cuánta demanda se pierde y por qué.',
      '100 % de los pedidos rechazados o reducidos con motivo registrado.',
      cajas=[('Ruta alternativa o negociar la fecha', 'Se simplifica'), ('Registro de la demanda perdida', NUEVA)],
      comite='Al negociar la fecha, el sistema registra la demanda que se pierde y su motivo.',
      cob=1.0, exc=0.05, detalle='D-PL04'),
    I('P2R-01', 'Automatización', 'Previsión visible en carga-capacidad y CTP',
      'Cargar en GIO la previsión cerrada de las referencias repetitivas para que la carga-capacidad y el CTP la '
      'vean junto a la cartera; cada pedido que llega consume previsión.',
      D(1, 9, 2027), D(31, 12, 2027),
      'La capacidad y las fechas tienen en cuenta lo que va a llegar, no solo lo que ya ha llegado.',
      'Carga-capacidad con previsión publicada cada mes y fechas CTP estables al entrar los pedidos previstos.',
      cajas=[('Previsión de referencias repetitivas', NUEVA),
             ('Carga frente a capacidad con escenarios de 2 a 5 turnos', 'Se simplifica')],
      comite='Nueva tarea de previsión en Demanda; la carga-capacidad pasa a ver cartera más previsión.',
      cob=1.0, exc=0.10, depende='Previsión no vinculante de la demanda repetitiva', detalle='D-PL01'),
    I('P2R-01', 'Agente', 'Agente de previsión de la demanda repetitiva',
      'Cada mes propone la previsión de las referencias repetitivas a partir del histórico de pedidos y de la demanda '
      'perdida. Comercial ajusta solo las excepciones antes del cierre.',
      D(1, 4, 2028), D(30, 9, 2028),
      'La previsión parte de un modelo y Comercial solo aporta lo que el modelo no sabe.',
      'Precisión ≥ 80 % en las referencias repetitivas y ≥ 80 % de ellas sin ajuste manual.',
      cajas=[('Previsión de referencias repetitivas', 'Se simplifica')],
      comite='La previsión de la demanda repetitiva la prepara un agente; Comercial revisa las excepciones.',
      cob=0.80, exc=0.20,
      depende='Registro de la demanda perdida (12 meses) · migración de GIO a web en Ondupack', detalle='D-PL01'),

    # ---------------------------------------------------------------- P2R-02 · S&OP y plan de capacidad
    I('P2R-02', 'Negocio', 'Capacidad real de la 924',
      'Medir con Producción la velocidad y el OEE reales de la 924 frente a los objetivos y acordar la capacidad que '
      'se puede planificar mientras no se corrija. La velocidad objetivo se mantiene en R2M.',
      D(1, 6, 2027), D(31, 7, 2027),
      'No se planifica en la 924 más de lo que realmente produce.',
      'Velocidad planificada de la 924 dentro de ±5 % de la real durante 2 meses.',
      decidir='Velocidad planificable de la 924 hasta que se resuelva', detalle='Maestro X03 (R2M)'),
    I('P2R-02', 'Negocio', 'Revisión mensual de capacidad y cuellos de botella',
      'Reunión mensual Planificación–Producción–Comercial con la carga de los próximos 2–3 meses por máquina: decide '
      'turnos con los escenarios de 2 a 5, qué entra en la 924 y qué se desvía por ruta alternativa.',
      D(1, 7, 2027), D(30, 9, 2027),
      'Turnos y cuellos de botella se deciden con antelación, no al ver las fechas atrasadas.',
      'Decisión de turnos documentada antes de cada mes y 0 semanas con la 924 por encima del 100 % sin plan.',
      depende='Capacidad real de la 924', detalle='D-PL06'),
    I('P2R-02', 'Automatización', 'Carga con velocidad real y aviso anticipado de saturación',
      'La carga-capacidad usa la velocidad y el OEE reales del histórico cuando se separan del objetivo, y avisa con '
      'semanas de antelación de la saturación de máquinas y del almacén intermedio.',
      D(1, 9, 2027), D(30, 11, 2027),
      'La sobresaturación se ve antes de que se note en las fechas o en el almacén intermedio.',
      '≥ 80 % de las saturaciones avisadas con más de 2 semanas de antelación.',
      cajas=[('¿Saturación > 100 %?', 'Se simplifica')],
      comite='El sistema anticipa la saturación con velocidades reales; la decisión se toma con tiempo.',
      cob=1.0, exc=0.15, depende='Capacidad real de la 924', detalle='D-PL07'),
    I('P2R-02', 'Agente', 'Agente de escenarios de capacidad',
      'Simula turnos, rutas alternativas para descargar la 924 y agrupaciones, y propone el plan con su impacto en '
      'servicio y coste. Planificación lo aprueba en la revisión mensual.',
      D(1, 4, 2028), D(30, 9, 2028),
      'Cada saturación llega con un plan propuesto y valorado.',
      '≥ 70 % de los planes aprobados sin rehacer.',
      cajas=[('Replanificar ruta, fecha o turnos', 'Se simplifica')],
      comite='La replanificación de ruta, fecha o turnos la propone un agente con escenarios; Planificación decide.',
      cob=0.70, exc=0.30,
      depende='Carga con velocidad real · migración de GIO a web en Ondupack', detalle='D-PL07'),

    # ---------------------------------------------------------------- P2R-05 · Materiales por excepción
    I('P2R-05', 'Negocio', 'Maestro MRP y política de reposición',
      'Completar lead time, lote mínimo, stock de seguridad y calendario de proveedor para papel y auxiliares, '
      'incluido el ciclo de MGT/PDA para el papel intercompany. Decidir qué auxiliares van por MRP y cuáles por '
      'mínimo y máximo. El dato se mantiene en R2M.',
      D(1, 6, 2027), D(31, 8, 2027),
      'Las necesidades de materiales se pueden calcular con parámetros fiables.',
      '100 % de papeles y auxiliares con parámetros completos.',
      decidir='Qué auxiliares van por MRP y cuáles por mínimo y máximo', detalle='D-PL15'),
    I('P2R-05', 'Negocio', 'Tabla de papeles sustitutos',
      'Aprobar con Producción y Calidad qué papel puede sustituir a cuál, y en qué condiciones, para que la sustitución '
      'deje de decidirse a mano en cada falta.',
      D(1, 7, 2027), D(30, 9, 2027),
      'Ante una falta se sabe al momento qué sustituto vale.',
      'Tabla aprobada para el 100 % de las calidades activas.',
      decidir='Quién aprueba una sustitución fuera de la tabla', detalle='D-PL14'),
    I('P2R-05', 'Automatización', 'Necesidad de papel calculada desde el plan',
      'Calcular cada día la necesidad neta de papel a partir de la cartera planificada, la previsión y la lista de '
      'materiales, en lugar del consumo pasado, y avisar de las faltas por orden y por fecha.',
      D(1, 9, 2027), D(31, 12, 2027),
      'Las faltas de materia prima se ven antes de que paren la producción.',
      '−50 % de órdenes liberadas con material pendiente en 6 meses.',
      cajas=[('Necesidad de papel por consumo pasado', 'La hace el sistema')],
      comite='La necesidad de papel la calcula el sistema con el plan y avisa de las faltas por orden.',
      cob=0.90, exc=0.10, depende='Maestro MRP y política de reposición',
      decidir='Sistema donde corre el cálculo (GIO o SAP)', detalle='D-PL16, D-PL17'),
    I('P2R-05', 'Automatización', 'Stock de papel fiable sin esperar al inventario',
      'Declarar el consumo real de papel por orden, que ya se puede trazar en GPO, en lugar de calcularlo por diferencia '
      'en el inventario mensual. El registro lo hace el proceso de inventario; aquí se usa.',
      D(1, 9, 2027), D(30, 11, 2027),
      'El cálculo de necesidades parte de un stock real cada día.',
      'Diferencia de inventario mensual de papel < 2 %.',
      comite='No cambia ninguna caja: mejora el stock que usa el cálculo de necesidades.',
      cob=1.0, exc=0.05, detalle='Proceso de inventario y almacén'),
    I('P2R-05', 'Automatización', 'Previsión de papel compartida con MGT/PDA',
      'En cada ciclo, Ondupack envía a MGT/PDA la previsión de papel de los próximos ciclos por calidad y ancho, además '
      'del pedido firme 10 días antes. El pedido sigue yendo por P2P.',
      D(1, 10, 2027), D(31, 12, 2027),
      'La papelera planifica su ciclo con una previsión, no solo con pedidos.',
      'Previsión enviada en el 100 % de los ciclos y desviación pedido–previsión < 15 %.',
      cajas=[('Compra de papel y auxiliares', 'Se simplifica')],
      comite='Junto al pedido de papel, Ondupack comparte su previsión con MGT/PDA; resuelve también el «sin previsión» de MGT.',
      cob=1.0, exc=0.10, depende='Necesidad de papel calculada desde el plan', detalle='D-PL13'),
    I('P2R-05', 'Agente', 'Agente de faltas de materia prima',
      'Cada día revisa las faltas por orden, propone la acción (adelantar el pedido a MGT, usar un sustituto aprobado o '
      'mover la orden) y prepara el aviso a Compras. Planificación aprueba.',
      D(1, 4, 2028), D(30, 9, 2028),
      'Cada falta llega a Planificación con la solución propuesta.',
      '≥ 80 % de las propuestas aceptadas y −30 % de horas de máquina paradas por falta de materia prima.',
      cajas=[('Autorizar la liberación con material pendiente', 'Se simplifica')],
      comite='Las faltas de materia prima las gestiona un agente que propone la solución; Planificación y Compras deciden.',
      cob=0.80, exc=0.20,
      depende='Necesidad de papel calculada desde el plan · Tabla de papeles sustitutos', detalle='D-PL14'),

    # ---------------------------------------------------------------- P2R-06 · Programación estable y optimizada
    I('P2R-06', 'Negocio', 'Zona congelada y canal único de cambios',
      'Fijar una zona congelada antes de onduladora: dentro de ella, los cambios de cantidad o fecha del cliente '
      'necesitan la aprobación de Planificación. Los cambios llegan por un canal único y no por correos sueltos; '
      'la recepción la hace O2C.',
      D(1, 6, 2027), D(31, 7, 2027),
      'El programa de los próximos días deja de moverse por cambios del cliente.',
      '≤ 2 cambios aprobados por semana dentro de la zona congelada.',
      decidir='Horas de la zona congelada · quién aprueba', detalle='D-PL08'),
    I('P2R-06', 'Negocio', 'Criterio de urgencia y causas de reprogramación',
      'Escribir qué es una urgencia, qué clientes no pueden fallar y quién autoriza desplazar pedidos, y crear la '
      'lista cerrada de causas de reprogramación.',
      D(1, 6, 2027), D(31, 7, 2027),
      'Las urgencias se resuelven con reglas conocidas, no hablando caso a caso.',
      'Criterio publicado y 100 % de urgencias resueltas con él.',
      detalle='D-PL20'),
    I('P2R-06', 'Automatización', 'Control de cambios y causa de reprogramación en GIO',
      'GIO bloquea los cambios dentro de la zona congelada salvo aprobación y obliga a elegir la causa al cambiar la '
      'ruta o la fecha de un pedido.',
      D(1, 8, 2027), D(30, 9, 2027),
      'Cada reprogramación queda registrada con su causa.',
      '100 % de las reprogramaciones con causa.',
      cajas=[('¿Alerta o cambio de ruta o fecha?', 'Se simplifica'), ('Replanificar ruta o fecha', 'Se simplifica')],
      comite='El sistema protege la zona congelada y guarda por qué se reprograma cada pedido.',
      cob=1.0, exc=0.10, depende='Zona congelada y canal único de cambios · Criterio de urgencia', detalle='D-PL08'),
    I('P2R-06', 'Automatización', 'Algoritmo de combinación de onduladora',
      'Propone las combinaciones por canal minimizando trim y cambios, y agrupa las referencias que hoy se repiten en '
      'la misma semana. El planificador las revisa y aprueba. Ya está en el plan de sistemas.',
      D(1, 3, 2027), D(30, 9, 2027),
      'La combinación deja de hacerse a mano y no se repiten maestros en la semana.',
      '−20 % de cambios en onduladora, merma de trim a la baja y 0 maestros repetidos en la semana sin motivo.',
      cajas=[('Combinación y secuencia de onduladora por ciclos de canal', 'Se simplifica'),
             ('Agrupar pedidos repetitivos', 'La hace el sistema')],
      comite='La combinación de onduladora la propone un algoritmo y agrupa las repeticiones; el planificador aprueba.',
      cob=0.85, exc=0.15, depende='Capacidad real de la 924 · velocidades, OEE y matrices de cambio estables (R2M)',
      detalle='D-PL18, D-PL19'),
    I('P2R-06', 'Automatización', 'Secuenciador de converting con matriz SMED',
      'Propone la secuencia de converting minimizando los cambios de troquel, clichés y tintas con los escenarios SMED '
      'que ya existen, sincronizada con la plancha disponible.',
      D(1, 11, 2027), D(31, 3, 2028),
      'La secuencia de converting minimiza cambios sin trabajo manual.',
      '−15 % de horas de cambio en converting y adherencia al programa ≥ 90 %.',
      cajas=[('Secuencia de converting con escenarios SMED', 'Se simplifica')],
      comite='La secuencia de converting la propone el sistema con la matriz SMED; el planificador aprueba.',
      cob=0.85, exc=0.15, depende='Algoritmo de combinación de onduladora · migración de GIO a web en Ondupack',
      detalle='D-PL19'),
    I('P2R-06', 'Agente', 'Agente de urgencias y cambios',
      'Ante un cambio del cliente o una urgencia, calcula el impacto (pedidos desplazados, cambios extra y saturación del '
      'almacén intermedio), aplica el criterio de urgencia y propone la decisión; si hay conflicto, la escala.',
      D(1, 4, 2028), D(30, 9, 2028),
      'Cada urgencia se decide con su impacto delante.',
      '100 % de las urgencias con impacto calculado y −50 % de escalados.',
      cajas=[('¿Alerta o cambio de ruta o fecha?', 'Se simplifica')],
      comite='Las urgencias las evalúa un agente con su impacto; solo los conflictos se escalan.',
      cob=0.75, exc=0.25, depende='Control de cambios y causa de reprogramación en GIO', detalle='D-PL20'),

    # ---------------------------------------------------------------- P2R-08 · Control del servicio
    I('P2R-08', 'Negocio', 'KPI oficial de servicio',
      'Que Dirección y Comercial decidan el KPI oficial (recomendado: OTIF contra fecha confirmada por el CTP, pedido '
      'completo) y su objetivo, y publicarlo cada semana.',
      D(1, 6, 2027), D(31, 7, 2027),
      'Hay una sola forma de medir el servicio en las dos plantas.',
      'KPI aprobado y publicado cada semana desde agosto de 2027.',
      decidir='KPI y fecha de referencia · objetivo', detalle='D-PL23'),
    I('P2R-08', 'Negocio', 'Reunión semanal con KPIs y causas',
      'La reunión semanal de planificación Ondupack/Navalmoral incorpora el OTIF, la adherencia y las principales causas '
      'de atraso y de reprogramación, cada una con su acción y su responsable.',
      D(1, 8, 2027), D(30, 9, 2027),
      'La reunión semanal decide acciones sobre las causas, no solo sobre los tickets.',
      'Cada causa principal del mes con acción y responsable.',
      depende='KPI oficial de servicio · Control de cambios y causa de reprogramación en GIO', detalle='D-PL24'),
    I('P2R-08', 'Automatización', 'Cuadro de mando de servicio y adherencia',
      'OTIF, adherencia al programa y pedidos en riesgo se calculan solos cada día desde GIO, para las dos plantas.',
      D(1, 9, 2027), D(30, 11, 2027),
      'El seguimiento de atrasos deja de ser una consulta manual.',
      'Cuadro de mando diario sin trabajo manual y ≥ 80 % de atrasos avisados antes de la fecha.',
      cajas=[('Seguimiento de fechas atrasadas', 'La hace el sistema')],
      comite='El seguimiento de fechas atrasadas lo hace el sistema y avisa de los pedidos en riesgo.',
      cob=1.0, exc=0.10, depende='KPI oficial de servicio', detalle='D-PL23'),
    I('P2R-08', 'Automatización', 'Propuesta de actualización de velocidades y OEE',
      'Cada mes compara la velocidad y el OEE reales con los objetivos y propone actualizar los que se desvían más del '
      'umbral, como la 924. La actualización se ejecuta en R2M con la actualización masiva que ya existe.',
      D(1, 10, 2027), D(31, 12, 2027),
      'El plan aprende de la realidad: los objetivos de velocidad no se quedan desfasados.',
      '100 % de las máquinas con velocidad objetivo dentro de ±5 % de la real.',
      comite='No cambia ninguna caja: mantiene al día el dato con el que se planifica.',
      cob=1.0, exc=0.20, depende='Capacidad real de la 924', decidir='Umbral de desviación', detalle='D-PL24'),
    I('P2R-08', 'Consulta', 'Copiloto de estado de pedidos',
      'Responde a Call center y a Comercial «¿cuándo sale el pedido X?» y «¿por qué va tarde?» con los datos de GIO, '
      'sin cambiar nada.',
      D(1, 4, 2028), D(30, 6, 2028),
      'Las consultas de estado dejan de llegar a Planificación.',
      '−60 % de consultas de estado a Planificación.',
      cajas=[('Seguimiento de fechas atrasadas', 'Se simplifica')],
      comite='Un copiloto responde el estado de los pedidos; Planificación solo atiende los casos raros.',
      cob=0.80, exc=0.20, depende='Cuadro de mando de servicio y adherencia', detalle='D-PL25'),
    I('P2R-08', 'Agente', 'Agente de causa raíz',
      'Clasifica cada atraso y desviación (velocidad, OEE, falta de papel, cambio del cliente o saturación), propone la '
      'acción y el ajuste de parámetros. Planificación aprueba.',
      D(1, 4, 2028), D(30, 9, 2028),
      'Cada atraso deja una corrección propuesta y las causas dejan de repetirse.',
      '≥ 90 % de los atrasos con causa y causas repetidas −30 % en 6 meses.',
      cajas=[('Ajustes Planificación y Pedidos', 'Se simplifica')],
      comite='Los ajustes tras un atraso los prepara un agente con su causa; Planificación los aprueba.',
      cob=0.75, exc=0.25, depende='Reunión semanal con KPIs y causas · Cuadro de mando de servicio y adherencia',
      detalle='D-PL24'),

    # ---------------------------------------------------------------- P2R-03 · solo si se confirma HT-P2R-04
    I('P2R-03', 'Negocio', 'Política de stock anticipado y sobrantes',
      'Decidir cuánto stock adelantado por agrupación se permite por referencia y qué se hace con los sobrantes al '
      'cerrar la orden: asignarlos al próximo pedido, venderlos o darlos de baja.',
      D(1, 9, 2027), D(31, 10, 2027),
      'El stock que generan las agrupaciones y los sobrantes tiene límite y destino.',
      '100 % del stock anticipado dentro del límite por referencia.',
      decidir='Crear solo si se confirma HT-P2R-04 · límite de cobertura por referencia', detalle='D-PL09'),
    I('P2R-03', 'Automatización', 'Control del stock anticipado y de sobrantes',
      'GIO marca el stock generado por agrupación o por sobrante y avisa cuando supera la cobertura permitida.',
      D(1, 11, 2027), D(31, 1, 2028),
      'El stock anticipado se ve y se corrige antes de que envejezca.',
      '€ de stock sin movimiento > 3 meses a la baja 3 meses seguidos.',
      cajas=[('Agrupar pedidos repetitivos', 'Se simplifica'), ('Cierre automático al terminar converting', 'Se simplifica')],
      comite='El sistema vigila el stock que generan las agrupaciones y los sobrantes.',
      cob=1.0, exc=0.10, depende='Política de stock anticipado y sobrantes',
      decidir='Crear solo si se confirma HT-P2R-04', detalle='D-PL09'),
]

if __name__ == '__main__':
    out = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).resolve().parents[1] / 'Intervenciones_P2R_Ondupack.xlsx')
    construir(out, empresa='Ondupack', palancas=PALANCAS, cajas_flujo=CAJAS, intervenciones=INTERVENCIONES,
              fase_nuevas=FASE_NUEVAS, segunda=('Navalmoral', 3))
    print(out)
