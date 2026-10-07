#!/usr/bin/env python3
"""Propuesta inicial del modelo de planificación corporativo de CL Grupo Industrial (packaging).

Es la única fuente de los datos que precarga la herramienta HTML y el Excel del kit.
Ejecutar este script regenera datos_base.json.
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import datos_equipo as equipo  # noqa: E402

VERSION = 1

EMPRESAS = [
    dict(id='CORP', nombre='Corporativo · Planificación', tipo='Corporativo', erp='SAP',
         herramienta='SAP · GIO · SIGPA', director='', responsable='', personasHoy='',
         particularidades='Define el modelo y gobierna la planificación de todas las plantas.', notas=''),
    dict(id='OPET', nombre='Ondupet', tipo='PET · extrusión y termoformado', erp='SAP',
         herramienta='GIO (escritorio; migración a web propuesta para 2027)', director='', responsable='', personasHoy='',
         particularidades='Ciclos por molde (termoformado) y por calidad (extrusión); MRP en SAP integrado con GIO.',
         notas=''),
    dict(id='OPACK_ALM', nombre='Ondupack Almendralejo', tipo='Cartón ondulado · onduladora y converting', erp='SAP',
         herramienta='Aplicación de escritorio (migración a web propuesta para 2027)', director='', responsable='',
         personasHoy='', particularidades='Combiplan para combinar pedidos; consume papel intercompany.', notas=''),
    dict(id='OPACK_NAV', nombre='Ondupack Navalmoral', tipo='Cartón ondulado · onduladora y converting', erp='SAP',
         herramienta='Aplicación de escritorio (migración a web propuesta para 2027)', director='', responsable='',
         personasHoy='', particularidades='Consume papel intercompany.', notas=''),
    dict(id='MGT', nombre='MGT', tipo='Papelera', erp='SAP', herramienta='SIGPA (web)', director='', responsable='',
         personasHoy='', particularidades='Algoritmo de combinación en SIGPA en pruebas; suministra papel a Ondupack.',
         notas=''),
    dict(id='PDA', nombre='PDA', tipo='Papelera', erp='SAP', herramienta='SIGPA (web)', director='', responsable='',
         personasHoy='', particularidades='Suministra papel a Ondupack.', notas=''),
    dict(id='PAPRESA', nombre='Papresa', tipo='Papelera · nueva incorporación', erp='ERP propio · SAP previsto 01/01/2027',
         herramienta='Por confirmar', director='', responsable='', personasHoy='',
         particularidades='Trazabilidad compra-venta en un HTML conectado a los distintos ERP hasta su paso a SAP.',
         notas=''),
]

# ámbito: Corporativo / Planta / Otras áreas
ROLES = [
    dict(id='DIR', corto='Dir. Planificación', nombre='Director/a de Planificación', ambito='Corporativo',
         mision='Dueño del modelo de planificación del grupo: define cómo se planifica, quién lo hace y con qué objetivos.',
         responsabilidades='Aprobar procesos, políticas y roles · Fijar objetivos y KPIs por planta · Presidir el S&OP '
                           'del grupo · Asignar personas y suplencias · Priorizar mejoras y desarrollos.',
         decide='Políticas, estructura de roles, asignación de personas, objetivos y prioridades de desarrollo.',
         kpis='OTIF del grupo · Madurez de procesos · % integración intercompany.'),
    dict(id='PO', corto='Resp. procesos', nombre='Responsable de procesos (Functional Lead)', ambito='Corporativo',
         mision='Diseña, documenta y mantiene los procedimientos estándar y vela por que cada planta los cumpla.',
         responsabilidades='Manual de planificación · Auditorías de madurez · Formación y certificación de '
                           'planificadores · Interlocución con las plantas.',
         decide='Cómo se ejecuta cada actividad; aprueba las excepciones locales al estándar.',
         kpis='Madurez de procesos · Adherencia al plan.'),
    dict(id='MD', corto='Dato maestro', nombre='Responsable de dato maestro', ambito='Corporativo',
         mision='Garantiza que el dato de planificación es único, completo y fiable en todas las plantas.',
         responsabilidades='Altas y cambios de datos de planificación · Velocidades, tiempos de cambio y estrategias · '
                           'Auditoría de calidad del dato · Mapeo de materiales intercompany.',
         decide='Aprobación de cualquier cambio en el dato maestro de planificación.',
         kpis='Calidad del dato maestro.'),
    dict(id='SOP', corto='S&OP / Intercompany', nombre='Coordinador/a de S&OP e intercompany', ambito='Corporativo',
         mision='Integra demanda y suministro del grupo y coordina los flujos de papel hacia las cartoneras.',
         responsabilidades='Preparar y conducir el S&OP mensual · Balanceo de carga entre plantas · Mínimos y máximos '
                           'intercompany · Seguimiento del % de integración.',
         decide='Reparto de capacidad entre plantas dentro de las políticas; escalados intercompany.',
         kpis='Precisión de la previsión · % integración intercompany · OTIF del grupo.'),
    dict(id='SIS', corto='Sistemas', nombre='Sistemas de planificación (Technical Lead y desarrollo)',
         ambito='Corporativo',
         mision='Construye y mantiene las herramientas de planificación comunes a todas las plantas (SAP, GIO, SIGPA).',
         responsabilidades='Desarrollos y mejoras · Despliegues controlados · Soporte de segundo nivel · '
                           'Documentación técnica.',
         decide='Arquitectura técnica y pasos a producción.',
         kpis='Plazo de entrega de mejoras · Incidencias resueltas en plazo.'),
    dict(id='RPP', corto='Resp. planif. planta', nombre='Responsable de planificación de planta', ambito='Planta',
         mision='Responde del plan de la planta y de que se ejecute según el estándar corporativo.',
         responsabilidades='Aprobar el plan semanal · Resolver conflictos de prioridad · Representar a la planta en '
                           'el S&OP · Coordinar al equipo de planificación de la planta.',
         decide='Plan semanal y prioridades dentro de las políticas; escala lo que las supera.',
         kpis='OTIF · Adherencia y estabilidad del plan.'),
    dict(id='PDC', corto='Demanda y cartera', nombre='Planificador/a de demanda y cartera', ambito='Planta',
         mision='Convierte previsión y pedidos en necesidades de fabricación y compromete fechas fiables.',
         responsabilidades='Previsión y demanda perdida · Cartera diaria y cierre de pedidos · Propuesta de '
                           'fabricación · Fechas al cliente y reservas de stock.',
         decide='Fecha comprometida dentro del plan y propuesta de fabricación.',
         kpis='Precisión de la previsión · Salud de la cartera · OTIF.'),
    dict(id='PPR', corto='Programación', nombre='Programador/a de producción', ambito='Planta',
         mision='Secuencia la producción para cumplir el plan con el menor coste de cambios.',
         responsabilidades='Plan semanal y secuencia (ciclos, combinación) · Lanzamiento de órdenes · Replanificación '
                           'ante incidencias · Carga de capacidad.',
         decide='Secuencia y lanzamiento dentro del horizonte congelado.',
         kpis='Adherencia al plan · Horas de cambio · Saturación por línea.'),
    dict(id='PMT', corto='Materiales', nombre='Planificador/a de materiales', ambito='Planta',
         mision='Asegura que semielaborado, papel y materias primas llegan a tiempo y en la cantidad necesaria.',
         responsabilidades='MRP · Necesidades de semielaborado y materias primas · Flujo intercompany y recepciones · '
                           'Balance de masas.',
         decide='Propuestas de aprovisionamiento dentro de las políticas.',
         kpis='Cobertura de stock · Roturas · % integración intercompany.'),
    dict(id='DPL', corto='Dir. planta', nombre='Dirección de planta', ambito='Otras áreas',
         mision='Dirige la operación de la planta y apoya la ejecución del plan.',
         responsabilidades='Recursos de producción · Conflictos operativos · OEE.',
         decide='Recursos operativos de la planta.', kpis='OEE · Coste de producción.'),
    dict(id='COM', corto='Comercial', nombre='Comercial / Atención al cliente', ambito='Otras áreas',
         mision='Aporta la visión de mercado y es la voz del cliente.',
         responsabilidades='Información de demanda · Comunicación de fechas y retrasos al cliente.',
         decide='Prioridad comercial dentro de las reglas acordadas.', kpis='Servicio al cliente.'),
    dict(id='PRO', corto='Producción', nombre='Producción (jefes de turno)', ambito='Otras áreas',
         mision='Ejecuta el plan en máquina e informa de lo real.',
         responsabilidades='Cumplir la secuencia · Registrar producción, paradas y consumos · Avisar de incidencias.',
         decide='Ajustes operativos dentro del turno.', kpis='OEE · Adherencia al plan.'),
    dict(id='LOG', corto='Logística', nombre='Logística / Expediciones', ambito='Otras áreas',
         mision='Expide los pedidos y gestiona los almacenes.',
         responsabilidades='Cargas vinculadas a pedido · Saturación de almacén.', decide='Organización de cargas.',
         kpis='Expediciones en fecha.'),
    dict(id='CMP', corto='Compras', nombre='Compras', ambito='Otras áreas',
         mision='Gestiona proveedores externos y el aprovisionamiento intercompany.',
         responsabilidades='Pedidos de compra · Seguimiento de proveedores.', decide='Proveedor y condiciones.',
         kpis='Plazo y fiabilidad de suministro.'),
    dict(id='MNT', corto='Mantenimiento', nombre='Mantenimiento', ambito='Otras áreas',
         mision='Planifica las paradas y garantiza la disponibilidad de las máquinas.',
         responsabilidades='Calendario de paradas · Avisos de averías.', decide='Ventanas de mantenimiento.',
         kpis='Disponibilidad.'),
]

# (id, nombre, [(id, nombre, descripción, frecuencia, RACI)])
# RACI: 'ROL:LETRA' separados por espacios; letras R, A, A/R, C, I. Exactamente una A por actividad.
PROCESOS = [
    ('G', 'Gobierno y S&OP', [
        ('G.1', 'Modelo de procesos y manual de planificación',
         'Definir, documentar y mantener cómo se ejecuta cada actividad en todas las plantas.', 'Continua',
         'DIR:A PO:R SIS:C RPP:C DPL:I'),
        ('G.2', 'Políticas de planificación',
         'MTS/MTO, stock de seguridad, horizonte congelado, prioridades, lotes, mínimos y máximos.', 'Trimestral',
         'DIR:A PO:R SOP:C RPP:C COM:C DPL:C PDC:I PPR:I PMT:I'),
        ('G.3', 'S&OP mensual del grupo',
         'Revisión de demanda, de suministro y reunión ejecutiva para decidir capacidad y prioridades.', 'Mensual',
         'DIR:A SOP:R RPP:C COM:C DPL:C CMP:C PO:I'),
        ('G.4', 'Objetivos y revisión de KPIs',
         'Fijar objetivos por planta y revisar el cuadro de mando común.', 'Mensual',
         'DIR:A PO:R RPP:C DPL:C PDC:I PPR:I PMT:I'),
        ('G.5', 'Roles, personas, suplencias y formación',
         'Asignar cada rol con nombre y apellidos, con titular y suplente, y su plan de formación.', 'Trimestral',
         'DIR:A PO:R RPP:C DPL:C PDC:I PPR:I PMT:I'),
        ('G.6', 'Gestión de cambios y prioridad de desarrollos',
         'Canal único de peticiones, priorización y paso a producción de mejoras.', 'Semanal',
         'DIR:A PO:R SIS:R RPP:C DPL:I'),
    ]),
    ('P1', 'Dato maestro', [
        ('1.1', 'Datos de planificación del material',
         'Altas y mantenimiento de los datos que usa la planificación (unidades, rutas, máquinas por defecto).',
         'Diaria', 'MD:A/R PMT:C PPR:C RPP:I SIS:I'),
        ('1.2', 'Segmentación ABC/XYZ y estrategia MTS/MTO',
         'Clasificar materiales y asignar su estrategia de planificación con criterio común.', 'Trimestral',
         'PO:A MD:R SOP:C RPP:C COM:C PDC:I PMT:I'),
        ('1.3', 'Velocidades, tiempos de cambio y rutas',
         'Mantener los estándares de máquina a partir de los datos reales registrados.', 'Mensual',
         'MD:A/R PPR:R PRO:C RPP:I'),
        ('1.4', 'Calendario de turnos, días y paradas',
         'Turnos activos, días festivos y paradas de mantenimiento por máquina.', 'Semanal',
         'RPP:A PPR:R PRO:C MNT:C MD:I PDC:I'),
        ('1.5', 'Auditoría de calidad del dato',
         'Detectar duplicados, datos incompletos e incoherencias y corregirlos.', 'Mensual',
         'PO:A MD:R RPP:C DIR:I'),
    ]),
    ('P2', 'Demanda', [
        ('2.1', 'Previsión de demanda', 'Generar y ajustar la previsión por material o familia.', 'Mensual',
         'RPP:A PDC:R COM:C SOP:C PMT:I PPR:I'),
        ('2.2', 'Consenso de la previsión con Comercial',
         'Revisar la previsión con Comercial y cerrar el plan de demanda.', 'Mensual',
         'SOP:A PDC:R COM:C RPP:C DIR:I'),
        ('2.3', 'Demanda perdida (agotados)',
         'Registrar los pedidos que no entran por falta de stock, precio u otros motivos.', 'Diaria',
         'RPP:A PDC:R COM:C SOP:I'),
        ('2.4', 'Precisión y sesgo de la previsión', 'Medir el error de la previsión y corregir sus causas.',
         'Mensual', 'SOP:A PDC:R RPP:C COM:C DIR:I'),
    ]),
    ('P3', 'Cartera', [
        ('3.1', 'Revisión diaria de la cartera', 'Pendiente, stock disponible y cargas de cada pedido.', 'Diaria',
         'RPP:A PDC:R COM:C PPR:I'),
        ('3.2', 'Cargas vinculadas a pedidos', 'Asociar cada carga o expedición a su pedido.', 'Diaria',
         'RPP:A LOG:R PDC:C COM:I'),
        ('3.3', 'Cierre (saldado) de pedidos', 'Cerrar los pedidos completados o anulados.', 'Semanal',
         'RPP:A PDC:R COM:C LOG:I'),
        ('3.4', 'Prioridad de pedidos en conflicto',
         'Decidir qué se sirve primero cuando no llega para todo; los conflictos entre plantas van a S&OP.',
         'Semanal', 'RPP:A PDC:R COM:C PPR:C SOP:C DPL:I'),
    ]),
    ('P4', 'Necesidades', [
        ('4.1', 'Ejecución y revisión del MRP', 'Lanzar el MRP y revisar sus propuestas y excepciones.', 'Diaria',
         'RPP:A PMT:R MD:C PPR:I CMP:I'),
        ('4.2', 'Propuesta de fabricación de producto terminado',
         'Qué fabricar según previsión, cartera, stock y órdenes ya lanzadas (MTS y MTO).', 'Diaria',
         'RPP:A PDC:R PMT:C PPR:C COM:I'),
        ('4.3', 'Necesidades de semielaborado',
         'Bobina, lámina o cartón necesarios para el producto terminado planificado.', 'Diaria',
         'RPP:A PMT:R PPR:C PRO:I'),
        ('4.4', 'Necesidades de materias primas y auxiliares',
         'Cálculo de necesidades y traslado a Compras.', 'Semanal', 'RPP:A PMT:R CMP:C SOP:I'),
    ]),
    ('P5', 'Capacidad', [
        ('5.1', 'Carga de capacidad y saturación por línea',
         '% de ocupación de cada máquina según lo lanzado y la previsión.', 'Semanal',
         'RPP:A PPR:R PRO:C SOP:I DPL:I'),
        ('5.2', 'Balanceo entre líneas y plantas',
         'Mover carga entre máquinas o plantas con antelación.', 'Mensual', 'SOP:A RPP:R PPR:C DPL:C DIR:I COM:I'),
        ('5.3', 'Saturación de almacenes', 'Proyección de ocupación de almacén de semielaborado y producto terminado.',
         'Semanal', 'RPP:A PPR:R LOG:C PRO:I'),
    ]),
    ('P6', 'Programación', [
        ('6.1', 'Plan de producción semanal', 'Plan congelado de la semana, aprobado por la planta.', 'Semanal',
         'RPP:A PPR:R PRO:C PDC:C PMT:C DPL:I COM:I'),
        ('6.2', 'Secuenciación, ciclos y combinación',
         'Orden de fabricación que minimiza cambios (moldes, calidades, combinaciones).', 'Diaria',
         'RPP:A PPR:R PRO:C PDC:I'),
        ('6.3', 'Lanzamiento de órdenes a máquina', 'Liberar las órdenes a producción.', 'Diaria',
         'RPP:A PPR:R PDC:C PRO:I'),
        ('6.4', 'Replanificación ante incidencias', 'Averías, faltas de material o urgencias.', 'Diaria',
         'RPP:A PPR:R PRO:C MNT:C PDC:C COM:I'),
        ('6.5', 'Cronograma de cambios', 'Proyección de cambios de molde, formato o calidad por máquina.', 'Semanal',
         'RPP:A PPR:R PRO:C MNT:C'),
    ]),
    ('P7', 'Entrega', [
        ('7.1', 'Compromiso de fecha al cliente', 'Consulta de plazos sobre la capacidad real (CTP).', 'Diaria',
         'RPP:A PDC:R PPR:C COM:I'),
        ('7.2', 'Reserva de stock a pedidos', 'Asignar el stock disponible a cada pedido por fechas.', 'Diaria',
         'RPP:A PDC:R LOG:C COM:I'),
        ('7.3', 'Retrasos y aviso al cliente', 'Detectar retrasos y avisar al cliente antes de que ocurran.', 'Diaria',
         'RPP:A PDC:R COM:R PPR:C DPL:I'),
        ('7.4', 'Coordinación con expediciones', 'Que lo planificado salga en la fecha comprometida.', 'Diaria',
         'RPP:A PDC:R LOG:R COM:I'),
    ]),
    ('P8', 'Control', [
        ('8.1', 'Plan frente a real (adherencia)', 'Seguimiento de lo producido frente a lo planificado.', 'Semanal',
         'RPP:A PPR:R PRO:C PO:I DPL:I'),
        ('8.2', 'OEE y rendimiento por máquina', 'Disponibilidad, rendimiento y calidad.', 'Semanal',
         'DPL:A PRO:R PPR:C RPP:I PO:I'),
        ('8.3', 'Balance de masas', 'Consumos reales frente a teóricos para detectar descuadres de stock.',
         'Semanal', 'RPP:A PMT:R PRO:C MD:C PO:I'),
        ('8.4', 'Causa raíz de desviaciones', 'Analizar por qué no se cumplió el plan y corregir el origen.',
         'Semanal', 'RPP:A PPR:R PRO:C PDC:C PMT:C PO:I'),
    ]),
    ('IC', 'Intercompany', [
        ('I.1', 'Mapeo de materiales y mínimos/máximos',
         'Tabla de correspondencia cartonera ↔ papelera y puntos de pedido.', 'Mensual',
         'SOP:A MD:R PMT:R RPP:C CMP:I'),
        ('I.2', 'Necesidad y pedidos de compra/venta',
         'Trigger de necesidad y pedidos de compra y venta mapeados.', 'Diaria', 'SOP:A PMT:R CMP:C COM:C RPP:I'),
        ('I.3', 'Planificación en papelera contra pedidos de stock',
         'La papelera planifica contra stock y reserva el material para las cartoneras.', 'Semanal',
         'RPP:A PPR:R SOP:C PMT:I'),
        ('I.4', 'Recepciones y % de integración', 'Seguimiento de lo recibido de MGT, PDA y Papresa.', 'Semanal',
         'SOP:A PMT:R CMP:C LOG:C DIR:I'),
        ('I.5', 'Control y cierre de pedidos intercompany', 'Apertura y cierre controlado de pedidos entre plantas.',
         'Semanal', 'SOP:A PMT:R CMP:C COM:C RPP:I'),
    ]),
]

POLITICAS = [
    ('Segmentación ABC/XYZ', 'Volumen (ABC) y variabilidad (XYZ) con el mismo criterio en todas las plantas; '
     'recálculo trimestral.', 'Umbrales de cada planta', 'MD'),
    ('Estrategia MTS/MTO', 'Regla común por segmento para decidir qué se fabrica contra stock y qué contra pedido; '
     'las excepciones las aprueba corporativo.', 'Excepciones por planta', 'PO'),
    ('Stock de seguridad y cobertura', 'Cobertura objetivo en días por segmento, revisada en el S&OP.',
     'Días de cobertura por segmento', 'SOP'),
    ('Horizonte congelado', 'Ventana en la que el plan no se cambia sin aprobación del responsable de planta.',
     'Días por tipo de planta', 'DIR'),
    ('Prioridad de pedidos', 'Orden común: fecha comprometida, cliente estratégico, antigüedad. Los conflictos entre '
     'plantas se deciden en el S&OP.', '', 'DIR'),
    ('Mínimos y máximos intercompany', 'Los propone la cartonera, los valida S&OP y se revisan cada mes.',
     'Por referencia', 'SOP'),
    ('Lote mínimo de fabricación', 'Lote mínimo por familia según el coste del cambio.', 'Por familia y máquina',
     'PO'),
    ('Nivel de servicio objetivo', 'OTIF objetivo por segmento de cliente o material.', 'Objetivo por planta', 'DIR'),
    ('Reglas de secuenciación', 'Ciclos por molde o calidad y combinación de pedidos según el estándar de cada tipo '
     'de planta.', 'Familias y ciclos de cada planta', 'PO'),
    ('Cambios en el dato maestro', 'Todo cambio se solicita al responsable de dato maestro, que lo aprueba y deja '
     'trazado.', '', 'MD'),
]

KPIS = [
    ('OTIF', 'Pedidos entregados en la fecha y cantidad comprometidas.',
     'Pedidos en fecha y completos / pedidos entregados', 'Mensual', 'RPP'),
    ('Adherencia al plan', 'Lo producido según el plan semanal.',
     'Órdenes fabricadas en su semana y cantidad / órdenes planificadas', 'Semanal', 'PPR'),
    ('Estabilidad del plan', 'Cambios dentro del horizonte congelado.',
     'Órdenes modificadas en horizonte congelado / órdenes del plan', 'Semanal', 'RPP'),
    ('Precisión de la previsión', 'Error y sesgo de la previsión.', '1 − MAPE; sesgo = Σ(previsión − real) / Σ real',
     'Mensual', 'SOP'),
    ('Cobertura de stock', 'Días de consumo que cubre el stock.', 'Stock / consumo medio diario', 'Semanal', 'PMT'),
    ('Horas de cambio', 'Tiempo de máquina dedicado a cambios.', 'Σ horas de cambio por línea', 'Semanal', 'PPR'),
    ('OEE', 'Eficiencia global de los equipos.', 'Disponibilidad × rendimiento × calidad', 'Semanal', 'PRO'),
    ('Saturación por línea', 'Ocupación prevista de cada máquina.', 'Horas cargadas / horas disponibles', 'Semanal',
     'PPR'),
    ('% integración intercompany', 'Recepciones de SQ y DU servidas por MGT, PDA o Papresa.',
     'Recepciones intercompany / recepciones totales', 'Mensual', 'SOP'),
    ('Salud de la cartera', 'Pedidos vencidos o sin cerrar.', 'Pedidos vencidos o abiertos indebidamente / cartera',
     'Semanal', 'PDC'),
    ('Calidad del dato maestro', 'Materiales con datos completos y sin duplicados.',
     'Materiales correctos / materiales activos', 'Mensual', 'MD'),
    ('Madurez del proceso', 'Procesos alineados con el estándar corporativo.',
     'Actividades alineadas / actividades aplicables', 'Trimestral', 'PO'),
]

CADENCIA = [
    ('S&OP ejecutivo del grupo', 'Mensual (1.ª semana)', '90 min',
     'Dir. Planificación, S&OP, responsables de planta, Comercial, Dirección General',
     'Plan de demanda consensuado, plan de suministro, KPIs',
     'Decisiones de capacidad, prioridades y riesgos'),
    ('Revisión de demanda', 'Mensual', '60 min', 'S&OP, planificadores de demanda, Comercial',
     'Previsión, demanda perdida, precisión', 'Plan de demanda consensuado'),
    ('Revisión de suministro', 'Mensual', '60 min', 'S&OP, responsables de planta, programación, materiales',
     'Carga de capacidad, coberturas, intercompany', 'Plan de suministro y escalados'),
    ('Comité semanal de plan (por planta)', 'Semanal', '45 min',
     'Responsable de planta, demanda y cartera, programación, materiales, Producción',
     'Cartera, carga, incidencias de la semana', 'Plan semanal congelado'),
    ('Daily de planificación (por planta)', 'Diaria', '15 min', 'Demanda y cartera, programación, Producción',
     'Control de producción, roturas, urgencias', 'Ajustes del día'),
    ('Revisión intercompany', 'Semanal', '30 min',
     'S&OP, materiales de las cartoneras, responsables de las papeleras',
     'Mínimos/máximos, pedidos abiertos, recepciones', 'Acciones de suministro'),
    ('Comité de planificación corporativo', 'Mensual', '60 min',
     'Dir. Planificación, procesos, dato maestro, S&OP, sistemas',
     'Avance del despliegue, cambios, peticiones de desarrollo', 'Decisiones sobre el modelo y prioridades'),
    ('Auditoría de madurez', 'Trimestral', '2 h por planta', 'Responsable de procesos, responsable de planta',
     'Notas de proceso y matriz de madurez', 'Plan de acción de la planta'),
]

DECISIONES = [
    ('Línea de dependencia de los planificadores de planta',
     'A) Jerárquica de corporativo con puesto en planta · B) Funcional de corporativo y jerárquica de planta',
     'A: es la coherente con el mandato de centralizar.', 'Dirección General y RR. HH.'),
    ('Horizonte congelado estándar', 'Uno por tipo de planta (PET, cartón, papel) o uno común',
     'Uno por tipo de planta, igual para todas las plantas del mismo tipo.', 'Dir. Planificación'),
    ('Qué puede decidir la planta sin escalar', 'Umbrales de cambio de plan, prioridad y fechas',
     'Fijar umbrales claros y un plazo máximo de respuesta de corporativo.', 'Dir. Planificación'),
    ('Herramienta única de planificación', 'SIGPA (web) en todas las plantas o convivencia con GIO',
     'Alinear con los proyectos 2027 de migración a web.', 'Dir. Planificación y Sistemas'),
    ('Planta piloto y fecha de arranque', 'Ondupet, Ondupack o papeleras',
     'Ondupet, por ser la planta con mayor madurez.', 'Dir. Planificación'),
    ('Papresa hasta su paso a SAP', 'Aplicar ya roles y cadencia o esperar a enero de 2027',
     'Aplicar ya roles y cadencia; las herramientas, tras el paso a SAP.', 'Dir. Planificación'),
]

PLAN = [
    ('0 · Diagnóstico', 'Quién planifica hoy en cada planta, qué hace y cuánto tiempo dedica; rellenar esta herramienta.',
     'Notas por proceso y empresa; directorio de personas.', 'Mes 1–2', 'PO'),
    ('1 · Diseño del modelo', 'RACI, fichas de rol, políticas, cadencia y KPIs; validación con la dirección y las '
     'plantas.', 'Manual de planificación v1 y RACI con nombres.', 'Mes 2–3', 'DIR'),
    ('2 · Piloto', 'Asignación nominal, formación, primer ciclo semanal y primer S&OP en la planta piloto.',
     'Piloto en marcha y KPIs de partida.', 'Mes 3–5', 'PO'),
    ('3 · Despliegue', 'Resto de plantas: Ondupack (Almendralejo y Navalmoral), MGT y PDA; Papresa tras su paso a SAP.',
     'Modelo implantado en todas las plantas.', 'Mes 5–9', 'DIR'),
    ('4 · Consolidación', 'S&OP del grupo, auditorías trimestrales y planificación en red papel → cartón.',
     'Ciclo de mejora continua.', 'Continuo', 'DIR'),
]

# Organigrama to-be propuesto (el as-is lo rellena cada planta). Sigue la opción A de la decisión abierta:
# el equipo de planificación depende jerárquicamente de corporativo y funcionalmente de la dirección de planta.
# (id, depende de, puesto, área, rol, empresa externa, dependencia funcional, notas)
ORG_TOBE_PLANTA = [
    ('t1', '', 'Director/a de Planificación', 'Dirección', 'DIR', 'CORP', '', 'Corporativo.'),
    ('t2', 't1', 'Responsable de planificación de planta', 'Planificación', 'RPP', '', 't6',
     'Dependencia jerárquica de corporativo y funcional de la dirección de planta (opción A, pendiente de decisión).'),
    ('t3', 't2', 'Planificador/a de demanda y cartera', 'Planificación', 'PDC', '', '', ''),
    ('t4', 't2', 'Programador/a de producción', 'Planificación', 'PPR', '', '', ''),
    ('t5', 't2', 'Planificador/a de materiales', 'Planificación', 'PMT', '', '', ''),
    ('t6', '', 'Dirección de planta', 'Dirección', 'DPL', '', '', ''),
]
ORG_TOBE_CORP = [
    ('c1', '', 'Dirección General', 'Dirección', '', '', '', ''),
    ('c2', 'c1', 'Director/a de Planificación', 'Planificación', 'DIR', '', '', ''),
    ('c3', 'c2', 'Responsable de procesos (Functional Lead)', 'Planificación', 'PO', '', '', ''),
    ('c4', 'c2', 'Responsable de dato maestro', 'Planificación', 'MD', '', '', ''),
    ('c5', 'c2', 'Coordinador/a de S&OP e intercompany', 'Planificación', 'SOP', '', '', ''),
    ('c6', 'c2', 'Technical Lead', 'Planificación', 'SIS', '', '', ''),
    ('c7', 'c6', 'Solution Developer', 'Planificación', 'SIS', '', '', ''),
    ('c8', 'c2', 'Responsables de planificación de planta', 'Planificación', 'RPP', '', '',
     'Uno por planta; ver el organigrama de cada empresa.'),
]


def build_org():
    out = {}
    for e in EMPRESAS:
        tpl = ORG_TOBE_CORP if e['id'] == 'CORP' else ORG_TOBE_PLANTA
        out[e['id']] = dict(asis=[], tobe=[dict(id=i, parent=p, puesto=pu, persona='', area=ar, rol=ro, empresa=ex,
                                                funcional=fu, dedicacion='', notas=no)
                                           for i, p, pu, ar, ro, ex, fu, no in tpl])
    return out


PLANT_ROLES = ['RPP', 'PDC', 'PPR', 'PMT']
CORP_ROLES = ['DIR', 'PO', 'MD', 'SOP', 'SIS']


def parse_raci(spec):
    out = {}
    for tok in spec.split():
        rol, letra = tok.split(':')
        out[rol] = letra
    return out


def build():
    procesos, raci = [], {}
    for pid, pname, acts in PROCESOS:
        items = []
        for aid, aname, desc, freq, spec in acts:
            items.append(dict(id=aid, nombre=aname, descripcion=desc, frecuencia=freq, herramienta=''))
            raci[aid] = parse_raci(spec)
            letters = list(raci[aid].values())
            assert sum(l in ('A', 'A/R') for l in letters) == 1, aid
            assert any(l in ('R', 'A/R') for l in letters), aid
            assert all(r in {x['id'] for x in ROLES} for r in raci[aid]), aid
        procesos.append(dict(id=pid, nombre=pname, actividades=items))
    ids_rol = {x['id'] for x in ROLES}
    assert all(r in ids_rol for f in equipo.PERSONAS.values() for r, _, _ in f)
    assert all(n[5] in ids_rol | {''} for f in equipo.ORG_ASIS.values() for n in f)
    personas, n = [], 0
    for e in EMPRESAS:
        filas = equipo.PERSONAS.get(e['id']) or [
            (rol, '', '') for rol in (CORP_ROLES if e['id'] == 'CORP' else PLANT_ROLES)]
        for rol, titular, dedicacion in filas:
            n += 1
            personas.append(dict(id=f'p{n}', empresa=e['id'], rol=rol, titular=titular, suplente='', puesto='',
                                 dedicacion=dedicacion, email='', notas=''))
    organigramas = build_org()
    for emp, nodos in equipo.ORG_ASIS.items():
        organigramas[emp]['asis'] = [dict(id=i, parent=p, puesto=pu, persona=pe, area=ar, rol=ro, empresa='',
                                          funcional='', dedicacion='', notas='') for i, p, pu, pe, ar, ro in nodos]
    return dict(
        meta=dict(version=VERSION, modelo='Modelo de planificación corporativo · CL Grupo Industrial',
                  autor='', actualizado=''),
        empresas=EMPRESAS,
        roles=ROLES,
        procesos=procesos,
        raci=raci,
        personas=personas,
        organigramas=organigramas,
        notas={},
        politicas=[dict(id=f'pol{i}', nombre=a, estandar=b, parametros=c, responsable=d, estado='Propuesta',
                        notas='') for i, (a, b, c, d) in enumerate(POLITICAS, 1)],
        kpis=[dict(id=f'kpi{i}', nombre=a, definicion=b, formula=c, frecuencia=d, responsable=e, objetivo='',
                   notas='') for i, (a, b, c, d, e) in enumerate(KPIS, 1)],
        cadencia=[dict(id=f'cad{i}', nombre=a, frecuencia=b, duracion=c, participantes=d, entradas=e, salidas=f,
                       notas='') for i, (a, b, c, d, e, f) in enumerate(CADENCIA, 1)],
        decisiones=[dict(id=f'dec{i}', nombre=a, opciones=b, recomendacion=c, responsable=d, fecha='',
                         estado='Abierta', resolucion='') for i, (a, b, c, d) in enumerate(DECISIONES, 1)],
        plan=[dict(id=f'fase{i}', nombre=a, actividades=b, entregables=c, plazo=d, responsable=e, estado='Pendiente',
                   notas='') for i, (a, b, c, d, e) in enumerate(PLAN, 1)],
    )


if __name__ == '__main__':
    data = build()
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'datos_base.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    n_act = sum(len(p['actividades']) for p in data['procesos'])
    print(f'OK {out}: {len(data["empresas"])} empresas, {len(data["roles"])} roles, {len(data["procesos"])} procesos, '
          f'{n_act} actividades, {len(data["personas"])} huecos de persona')
