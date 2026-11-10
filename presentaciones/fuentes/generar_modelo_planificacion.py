#!/usr/bin/env python3
"""Presentación para dirección: nuevo modelo de planificación de Packaging (Celeste, organigrama y acciones).

Portada y tres diapositivas con la plantilla corporativa (las utilidades de dibujo vienen de generar_p2r_direccion):
  1. Celeste (DICE), portal único de planificación, con la captura del dashboard de integración papel-cartón.
  2. Organigrama de planificación: as-is por planta y to-be con Planificador Máster.
  3. Acciones para el nuevo organigrama, por planta, sacadas del análisis de P2R.

Uso: python3 presentaciones/fuentes/generar_modelo_planificacion.py
"""
from PIL import Image
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from generar_p2r_direccion import (AZUL, BLANCO, CARD, HIELO, ICONOS, L, LINEA, MUTED, NAVY, R, RAIZ, TEAL, W, Deck,
                                   caja, icono, linea, pastilla, rellenar, rotulo, texto)

SALIDA = RAIZ / 'presentaciones' / 'Nuevo_Modelo_Planificacion.pptx'
CAPTURA = ICONOS.parent / 'capturas' / 'celeste_integracion_papel_carton.png'
PLANTAS = ('MGT', 'PDA', 'Ondupack Almendralejo', 'Ondupack Navalmoral')


# ------------------------------------------------------------------ 1. Celeste
def celeste(d):
    sl = d.contenido_('Celeste · Portal único de planificación', notas=(
        'Celeste (DICE) reúne en un solo portal todo lo que hoy se consulta en correos, Excel y varios sistemas: '
        'necesidades, stock, urgencias e intercompany. Es la herramienta común del nuevo organigrama.'))
    b = caja(sl, L, 1.45, W, 0.62, relleno=NAVY)
    rellenar(b.text_frame, [[('Celeste :  ', {'color': AZUL}),
                             ('una sola pantalla para planificar papel y cartón: necesidades, stock, urgencias e '
                              'intercompany.', {'color': BLANCO})]],
             tam=12.5, negrita=True, ancla=MSO_ANCHOR.MIDDLE, margen=0.3)
    modulos = [
        ('cubos', 'Necesidades', 'Faltas de material y propuestas de compra.'),
        ('almacen', 'Stock', 'Propio y de las otras plantas, con mínimos y máximos.'),
        ('reloj', 'Urgencias', 'Si caben, qué pedidos mueven y quién decide.'),
        ('intercambio', 'Intercompany', 'Previsión, pedidos y cargas entre papeleras y cartoneras.'),
        ('calendario', 'Capacidad y programación', 'Carga por máquina, escenarios y secuencia.'),
        ('panel', 'Servicio', 'OTIF, atrasos y sus causas.'),
    ]
    y0, th, tw, gap = 2.3, 1.02, 2.65, 0.15
    for i, (ic, tit, txt) in enumerate(modulos):
        x = L + (i % 2) * (tw + gap)
        y = y0 + (i // 2) * (th + gap)
        caja(sl, x, y, tw, th, relleno=CARD)
        icono(sl, ic, x + 0.15, y + 0.17, d=0.42)
        texto(sl, x + 0.68, y + 0.13, tw - 0.8, 0.5, tit, tam=10.5, negrita=True, ancla=MSO_ANCHOR.MIDDLE,
              interlineado=1.0)
        texto(sl, x + 0.15, y + 0.64, tw - 0.3, 0.36, txt, tam=9, color=TEAL, interlineado=1.0)
    # captura de Celeste, centrada en el hueco de la derecha con su pie
    fx, fw = L + 2 * tw + gap + 0.3, R - (L + 2 * tw + gap + 0.3)
    fh = 3 * th + 2 * gap
    with Image.open(CAPTURA) as im:
        ih = fw * im.height / im.width
    pie = 0.25
    yi = y0 + (fh - (ih + 0.08 + pie)) / 2
    pic = sl.shapes.add_picture(str(CAPTURA), Inches(fx), Inches(yi), Inches(fw), Inches(ih))
    pic.name = 'Captura de Celeste'
    pic._element.nvPicPr.cNvPr.set('descr', 'Captura de Celeste: dashboard de integración papel-cartón')
    pic.line.color.rgb = LINEA
    pic.line.width = Pt(0.75)
    texto(sl, fx, yi + ih + 0.08, fw, pie, 'Integración papel-cartón: entregas, proveedores, plazos y estado de los '
          'pedidos intercompany.', tam=9, cursiva=True, color=MUTED, alin=PP_ALIGN.CENTER)
    yb = y0 + fh + 0.18
    caja(sl, L, yb, W, 0.45, relleno=HIELO)
    icono(sl, 'check', L + 0.12, yb + 0.075, d=0.3)
    texto(sl, L + 0.55, yb, W - 0.7, 0.45, [[('Sustituye a los correos, los Excel y las consultas en varios sistemas, ',
                                              {'negrita': True}),
                                             ('en MGT, PDA, Ondupack Almendralejo y Ondupack Navalmoral.', {})]],
          tam=10, ancla=MSO_ANCHOR.MIDDLE)


# ------------------------------------------------------------------ 2. Organigrama
def organigrama(d):
    sl = d.contenido_('Organigrama · As-is y to-be', 'Papel y cartón', notas=(
        'Hoy cada planta planifica con su propio equipo. En el to-be se planifica por material: un planificador de '
        'papel y otro de cartón, coordinados por un Planificador Máster del que depende aprovisionamiento. Comercial '
        'queda al mismo nivel que el Máster.'))
    # --- as-is
    ax, aw = L, 5.75
    rotulo(sl, ax, 1.45, aw, 'As-is · cada planta con su propio equipo', color=MUTED)
    cw, cg = (aw - 3 * 0.12) / 4, 0.12
    for g, (grupo, idx) in enumerate((('Papel', (0, 1)), ('Cartón', (2, 3)))):
        gx = ax + idx[0] * (cw + cg)
        gw = 2 * cw + cg
        texto(sl, gx, 1.78, gw, 0.22, grupo.upper(), tam=8.5, negrita=True, color=TEAL, alin=PP_ALIGN.CENTER)
        linea(sl, gx + 0.05, 2.03, gx + gw - 0.05, 2.03, color=AZUL, grosor=1.5)
    roles = ('Secuenciador', 'Logística', 'Comercial', 'Aprovisionamiento')
    for k, planta in enumerate(PLANTAS):
        x = ax + k * (cw + cg)
        c = caja(sl, x, 2.15, cw, 0.6, relleno=NAVY)
        rellenar(c.text_frame, planta, tam=9, negrita=True, color=BLANCO, alin=PP_ALIGN.CENTER,
                 ancla=MSO_ANCHOR.MIDDLE, margen=0.05, interlineado=1.0)
        for j, rol in enumerate(roles):
            r = caja(sl, x, 2.88 + j * 0.58, cw, 0.48, relleno=CARD, borde=LINEA)
            rellenar(r.text_frame, rol, tam=9, color=NAVY, alin=PP_ALIGN.CENTER, ancla=MSO_ANCHOR.MIDDLE,
                     margen=0.05, interlineado=1.0)
    yb = 2.88 + 4 * 0.58 + 0.12
    c = caja(sl, ax, yb, aw, 6.35 - yb, relleno=HIELO)
    rellenar(c.text_frame, [[('Hoy: ', {'negrita': True}),
                             ('cada planta tiene su secuenciador, su coordinador de cargas, su comercial y su '
                              'aprovisionamiento, y la coordinación entre plantas va por correo.', {})]],
             tam=10, color=TEAL, ancla=MSO_ANCHOR.MIDDLE, margen=0.2, interlineado=1.05)
    # flecha
    f = caja(sl, ax + aw + 0.12, 3.7, 0.36, 0.42, relleno=AZUL, forma=MSO_SHAPE.CHEVRON)
    f.adjustments[0] = 0.5
    # --- to-be
    tx = ax + aw + 0.6
    tw = R - tx
    rotulo(sl, tx, 1.45, tw, 'To-be · planificación única por material', color=TEAL)
    ny, nh = 1.85, 0.8
    mx, mw = tx + 0.55, 2.6
    m = caja(sl, mx, ny, mw, nh, relleno=NAVY)
    rellenar(m.text_frame, [[('Planificador Máster', {'negrita': True, 'tam': 12})],
                            [('Papel y cartón · todas las plantas', {'tam': 8.5, 'color': AZUL})]],
             color=BLANCO, alin=PP_ALIGN.CENTER, ancla=MSO_ANCHOR.MIDDLE, interlineado=1.05)
    cx, cwid = R - 1.95, 1.95
    c = caja(sl, cx, ny, cwid, nh, relleno=AZUL)
    rellenar(c.text_frame, [[('Comercial', {'negrita': True, 'tam': 12})], [('', {'tam': 8.5})]],
             color=NAVY, alin=PP_ALIGN.CENTER, ancla=MSO_ANCHOR.MIDDLE, interlineado=1.05)
    linea(sl, mx + mw, ny + nh / 2, cx, ny + nh / 2, color=TEAL, grosor=1.25, discontinua=True)
    hijos = [('Planificador de papel', 'MGT · PDA · Papresa'), ('Planificador de cartón', 'Almendralejo · Navalmoral'),
             ('Aprovisionamiento', 'de cada planta')]
    hy, hh, hg = 3.3, 0.9, 0.15
    hw = (tw - 2 * hg) / 3
    bus = ny + nh + 0.25
    centros = [tx + i * (hw + hg) + hw / 2 for i in range(3)]
    linea(sl, mx + mw / 2, ny + nh, mx + mw / 2, bus, color=NAVY, grosor=1.25)
    linea(sl, centros[0], bus, centros[-1], bus, color=NAVY, grosor=1.25)
    for i, (tit, sub) in enumerate(hijos):
        x = tx + i * (hw + hg)
        linea(sl, centros[i], bus, centros[i], hy, color=NAVY, grosor=1.25)
        h = caja(sl, x, hy, hw, hh, relleno=BLANCO, borde=NAVY)
        rellenar(h.text_frame, [[(tit, {'negrita': True, 'tam': 10.5})], [(sub, {'tam': 8.5, 'color': MUTED})]],
                 color=NAVY, alin=PP_ALIGN.CENTER, ancla=MSO_ANCHOR.MIDDLE, margen=0.08, interlineado=1.05)
    cy = hy + hh + 0.3
    c = caja(sl, tx, cy, tw, 6.35 - cy, relleno=CARD)
    puntos = [
        ('Se planifica por material, ', 'no por planta: papel (MGT y PDA) y cartón (Almendralejo y Navalmoral).'),
        ('La secuenciación y la coordinación de cargas ', 'pasan a los planificadores, con el algoritmo y los agentes.'),
        ('Comercial y Planificador Máster ', 'acuerdan juntos prioridades y estrategias.'),
        ('Todos trabajan sobre Celeste, ', 'con la misma información.'),
    ]
    texto(sl, tx + 0.2, cy + 0.15, tw - 0.4, 6.35 - cy - 0.3,
          [[(a, {'negrita': True, 'color': NAVY}), (b, {})] for a, b in puntos], tam=9.5, color=TEAL,
          interlineado=1.0, espacio=4)


# ------------------------------------------------------------------ 3. Acciones
# (palanca, texto, destacada, tiene agente)
ACCIONES = {
    'papel': [
        ('P2R-06', [('Afinar el algoritmo de combinación: ', True),
                    ('que tenga en cuenta lo ya planificado y una fecha límite de cartera.', False)], True, False),
        ('P2R-01', [('Demanda: ', True), ('previsión de papel de Ondupack por ciclo y de los clientes repetitivos.',
                                          False)], False, False),
        ('P2R-02', [('Capacidad: ', True), ('política de llenado y revisión mensual de capacidad y ciclos.', False)],
         False, True),
        ('P2R-03', [('Inventario: ', True), ('reposición pull con Ondupack, con mínimo y máximo por referencia.', False)],
         False, True),
        ('P2R-04', [('Fecha: ', True), ('copiloto de plazos para Comercial.', False)], False, True),
        ('P2R-05', [('Materiales: ', True), ('MRP de materias primas en SAP y propuestas de compra.', False)], False,
         True),
        ('P2R-06', [('Programación: ', True), ('urgencias con reglas.', False)], False, True),
        ('P2R-07', [('Liberación: ', True), ('disponibilidad de material por orden en el sistema.', False)], False,
         False),
        ('P2R-08', [('Servicio: ', True), ('KPI y pilar semanal; stock y cargas visibles para Ondupack.', False)],
         False, True),
    ],
    'carton': [
        ('P2R-06', [('Implantar el algoritmo de combinación ', True),
                    ('de onduladora en Almendralejo y Navalmoral.', False)], True, False),
        ('P2R-01', [('Demanda: ', True), ('previsión de clientes y referencias repetitivos en el sistema y registro de '
                                          'la demanda perdida.', False)], False, False),
        ('P2R-02', [('Capacidad: ', True), ('revisión mensual de capacidad y cuellos de botella, con escenarios.', False)],
         False, True),
        ('P2R-05', [('Materiales: ', True), ('papel calculado desde el plan, stock fiable y previsión compartida con '
                                             'MGT, PDA y Papresa.', False)], False, True),
        ('P2R-06', [('Programación: ', True), ('urgencias y cambios con reglas.', False)], False, True),
        ('P2R-08', [('Servicio: ', True), ('KPI oficial, reunión semanal con causas y estado de pedidos.', False)],
         False, True),
    ],
}


def acciones(d):
    sl = d.contenido_('Acciones para el nuevo organigrama', 'Por planta', notas=(
        'Las acciones salen del análisis de Plan to Release (palancas e intervenciones). El algoritmo de combinación '
        'es la pieza clave: se afina en las papeleras y se implanta en las cartoneras. El icono marca las acciones '
        'con agente.'))
    gap = 0.3
    cw = (W - gap) / 2
    cols = [('industria', 'Papel · MGT y PDA', '8 palancas', 'papel'),
            ('caja', 'Cartón · Ondupack Almendralejo y Navalmoral', '5 palancas', 'carton')]
    fh, fg = 0.41, 0.05
    for k, (ic, tit, sub, clave) in enumerate(cols):
        x = L + k * (cw + gap)
        icono(sl, ic, x, 1.45, d=0.46)
        texto(sl, x + 0.6, 1.43, cw - 0.6, 0.3, tit, tam=12, negrita=True)
        texto(sl, x + 0.6, 1.72, cw - 0.6, 0.22, sub, tam=8.5, color=MUTED)
        for i, (cod, partes, destacada, agente) in enumerate(ACCIONES[clave]):
            y = 2.08 + i * (fh + fg)
            caja(sl, x, y, cw, fh, relleno=NAVY if destacada else CARD)
            pastilla(sl, x + 0.1, y + (fh - 0.22) / 2, 0.66, 0.22, cod, relleno=AZUL if destacada else NAVY,
                     color=NAVY if destacada else BLANCO, tam=7)
            color = BLANCO if destacada else TEAL
            runs = [(t, {'negrita': neg, 'color': (AZUL if destacada else NAVY) if neg else color}) for t, neg in partes]
            texto(sl, x + 0.86, y, cw - 1.3, fh, [runs], tam=9.5, color=color, ancla=MSO_ANCHOR.MIDDLE,
                  interlineado=1.0)
            if agente:
                sl.shapes.add_picture(str(ICONOS / 'robot_n.png'), Inches(x + cw - 0.34), Inches(y + (fh - 0.22) / 2),
                                      Inches(0.22), Inches(0.22))
    # común, bajo la columna de cartón
    x = L + cw + gap
    yc = 2.08 + len(ACCIONES['carton']) * (fh + fg) + 0.15
    hc = 2.08 + len(ACCIONES['papel']) * (fh + fg) - fg - yc
    c = caja(sl, x, yc, cw, hc, relleno=HIELO)
    rellenar(c.text_frame, [[('Común a todas las plantas', {'negrita': True, 'color': NAVY, 'tam': 10.5})],
                            [('Celeste como portal único (necesidades, stock, urgencias e intercompany) y la nueva '
                              'organización: Planificador Máster y planificadores de papel y cartón.', {})]],
             tam=9.5, color=TEAL, ancla=MSO_ANCHOR.MIDDLE, margen=0.2, interlineado=1.05, espacio=3)
    # pie
    yp = 2.08 + len(ACCIONES['papel']) * (fh + fg) + 0.04
    sl.shapes.add_picture(str(ICONOS / 'robot_n.png'), Inches(L), Inches(yp + 0.03), Inches(0.18), Inches(0.18))
    texto(sl, L + 0.25, yp, W - 0.25, 0.24, [[('incluye agente  ·  ', {}),
                                              ('PDA y Navalmoral: mismas acciones que MGT y Almendralejo, pendientes de '
                                               'validar en cada planta.', {'cursiva': True})]], tam=8.5, color=MUTED,
          ancla=MSO_ANCHOR.MIDDLE)


def main():
    d = Deck()
    d.portada('Planificación Packaging', 'Celeste, organigrama y acciones por planta', ' Noviembre 2026')
    celeste(d)
    organigrama(d)
    acciones(d)
    d.guardar(SALIDA)
    print(SALIDA)


if __name__ == '__main__':
    main()
