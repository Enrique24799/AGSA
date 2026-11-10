#!/usr/bin/env python3
"""Presentación de tres diapositivas sobre el flujo de planificación de papel y cartón.

  1. Flujo as-is: fases de Plan to Release y títulos de las actividades (sin pasos) en papel (MGT y PDA) y cartón
     (Ondupack Almendralejo y Navalmoral), con el secuenciador de cada planta.
  2. El mismo flujo con las automatizaciones y agentes necesarios en cada planta (intervenciones de los maestros,
     más afinar el algoritmo de combinación en papel e implantarlo en cartón).
  3. Organigrama to-be: Planificador Máster y planificadores de papel y cartón.

Usa la plantilla corporativa y las utilidades de generar_p2r_direccion.
Uso: python3 presentaciones/fuentes/generar_flujo_secuenciadores.py
"""
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches

from generar_p2r_direccion import (AZUL, BLANCO, CARD, GRIS, HIELO, ICONOS, L, LINEA, MUTED, NAVY, R, RAIZ, TEAL, W, Deck,
                                   caja, icono, linea, rellenar, rotulo, texto)

SALIDA = RAIZ / 'presentaciones' / 'Flujo_Planificacion_Plantas.pptx'

FASES = [('demanda', 'Demanda y previsión'), ('balanza', 'S&OP y Plan Maestro'),
         ('cubos', 'Necesidades de fabricación y MRP'), ('calendario', 'Programación'), ('play', 'Liberación'),
         ('panel', 'Seguimiento y control')]
VACIA = 'Previsión y precisión'  # actividad que existe en el flujo pero hoy no se hace

# carril: (clave, rótulo, plantas, secuenciadores [(iniciales, nombre, planta)], actividades por fase)
CARRILES = [
    ('papel', 'Papel', 'MGT · PDA', [('JC', 'Joan Carles Girones', 'MGT y PDA')], [
        ['Demanda intercompany', 'Demanda externa', VACIA],
        ['Plan de fabricación', 'Carga de capacidad'],
        ['Propuesta de fabricación', 'MRP'],
        ['Programación'],
        ['Liberación de la orden'],
        ['Seguimiento del servicio', 'Entregas a Ondupack y cierre'],
    ]),
    ('carton', 'Cartón', 'Almendralejo · Navalmoral',
     [('MA', 'Miguel Ángel Rosado', 'Almendralejo'), ('TB', 'Tomás Buenavida', 'Navalmoral')], [
         ['Entrada de pedidos y fecha', VACIA],
         ['Cartera', 'Carga capacidad'],
         ['Balance de plancha y PT', 'Papel y auxiliares'],
         ['Onduladora', 'Converting'],
         ['Liberación'],
         ['Seguimiento del servicio', 'Revisión semanal'],
     ]),
]

AUTO, AGENTE, ALGORITMO = 'auto', 'agente', 'algoritmo'
# automatizaciones por carril y fase: [(actividad, [(tipo, texto)])]
AUTOMATIZACIONES = {
    'papel': [
        [('Demanda externa', [(AGENTE, 'Copiloto de plazos para Comercial')])],
        [('Plan de fabricación', [(AGENTE, 'Agente de stock intercompany')]),
         ('Carga de capacidad', [(AGENTE, 'Agente de escenarios de ciclo')])],
        [('MRP', [(AUTO, 'MRP de materias primas en SAP'), (AGENTE, 'Agente de compras de materiales')])],
        [('Programación', [(ALGORITMO, 'Afinar el algoritmo de combinación'), (AGENTE, 'Agente de urgencias')])],
        [('Liberación de la orden', [(AUTO, 'Disponibilidad de material por orden')])],
        [('Entregas a Ondupack y cierre', [(AUTO, 'Stock y cargas visibles para Ondupack'),
                                           (AGENTE, 'Agente de coordinación de cargas')])],
    ],
    'carton': [
        [('Entrada de pedidos y fecha', [(AUTO, 'Registro de la demanda perdida')])],
        [('Carga capacidad', [(AGENTE, 'Agente de escenarios de capacidad')])],
        [('Papel y auxiliares', [(AUTO, 'Necesidad de papel desde el plan'), (AUTO, 'Stock de papel fiable'),
                                 (AGENTE, 'Agente de faltas de materia prima')])],
        [('Onduladora', [(ALGORITMO, 'Implantar el algoritmo de combinación')]),
         ('Converting', [(AGENTE, 'Agente de urgencias y cambios')])],
        [],
        [('Seguimiento del servicio', [(AUTO, 'Copiloto de estado de pedidos'), (AGENTE, 'Agente de causa raíz')])],
    ],
}

# geometría común a las dos diapositivas de flujo: una tarjeta alta por fase (como «Cómo funciona hoy, fase a fase»)
# y dos carriles, papel y cartón, con su cabecera a la izquierda
XH, WH = L, 1.9                       # cabecera de carril
X0 = L + WH + 0.15                    # primera fase
GAP = 0.12
CW = (R - X0 - 5 * GAP) / 6           # ancho de cada fase
YT, YB = 1.45, 6.35                   # tarjetas de fase
YC = 2.5                              # fin de la cabecera de las tarjetas
HL, GL = 1.87, 0.11                   # alto de cada carril y separación


def borde_discontinuo(shape):
    ln = shape.line._get_or_add_ln()
    ln.append(ln.makeelement(qn('a:prstDash'), {'val': 'dash'}))


def fases(sl):
    for i, (ic, fase) in enumerate(FASES):
        x = X0 + i * (CW + GAP)
        caja(sl, x, YT, CW, YB - YT, relleno=CARD)
        icono(sl, ic, x + 0.12, YT + 0.12, d=0.42)
        texto(sl, x + 0.62, YT + 0.16, 0.4, 0.3, str(i + 1), tam=9, color=GRIS, negrita=True)
        texto(sl, x + 0.12, YT + 0.6, CW - 0.2, 0.42, fase, tam=10, negrita=True, interlineado=1.0)
        yl = YC + HL + GL / 2
        linea(sl, x + 0.1, yl, x + CW - 0.1, yl, color=LINEA, grosor=1)


def carriles(sl):
    """Cabecera de cada carril con el secuenciador de cada planta; devuelve la y de cada carril."""
    ys = []
    for k, (_, rot, plantas, personas, _) in enumerate(CARRILES):
        y = YC + k * (HL + GL)
        ys.append(y)
        caja(sl, XH, y, WH, HL, relleno=HIELO)
        icono(sl, 'industria' if k == 0 else 'caja', XH + 0.12, y + 0.12, d=0.38)
        rotulo(sl, XH + 0.58, y + 0.1, 1.2, rot)
        texto(sl, XH + 0.58, y + 0.3, WH - 0.64, 0.42, plantas, tam=9.5, negrita=True, interlineado=1.0)
        for j, (ini, nom, planta) in enumerate(personas):
            yp = y + 0.78 + j * 0.55
            caja(sl, XH + 0.12, yp, 0.4, 0.4, relleno=NAVY, forma=MSO_SHAPE.OVAL)
            texto(sl, XH + 0.12, yp, 0.4, 0.4, ini, tam=8.5, negrita=True, color=BLANCO, alin=PP_ALIGN.CENTER,
                  ancla=MSO_ANCHOR.MIDDLE)
            texto(sl, XH + 0.6, yp - 0.03, WH - 0.65, 0.5,
                  [[(nom, {'negrita': True})], [(f'Secuenciador · {planta}', {'tam': 7.5, 'color': MUTED})]],
                  tam=8.5, interlineado=1.0)
    return ys


def flujo_asis(d):
    sl = d.contenido_('Flujo de planificación · As-is', 'As-is', notas=(
        'Las seis fases de Plan to Release y las actividades de cada una, sin el detalle de los pasos. Papel (MGT y '
        'PDA) y cartón (Almendralejo y Navalmoral) siguen el mismo esquema; en cada planta influye su secuenciador.'))
    fases(sl)
    ys = carriles(sl)
    for (_, _, _, _, acts), y in zip(CARRILES, ys):
        for i, lista in enumerate(acts):
            x = X0 + i * (CW + GAP)
            for j, act in enumerate(lista):
                yy = y + 0.04 + j * 0.5
                vacia = act == VACIA
                c = caja(sl, x + 0.08, yy, CW - 0.16, 0.42, relleno=CARD if vacia else BLANCO,
                         borde=GRIS if vacia else LINEA)
                if vacia:
                    borde_discontinuo(c)
                rellenar(c.text_frame, act, tam=8.5, color=MUTED if vacia else NAVY, cursiva=vacia,
                         alin=PP_ALIGN.CENTER, ancla=MSO_ANCHOR.MIDDLE, margen=0.05, interlineado=1.0)
    # leyenda en el hueco sobre las cabeceras de carril
    c = caja(sl, XH + 0.05, YT + 0.3, 0.42, 0.28, relleno=CARD, borde=GRIS)
    borde_discontinuo(c)
    texto(sl, XH + 0.55, YT + 0.22, WH - 0.55, 0.44, 'Existe en el flujo, pero hoy no se hace', tam=7.5, color=MUTED,
          cursiva=True, ancla=MSO_ANCHOR.MIDDLE, interlineado=1.0)


def flujo_automatizaciones(d):
    sl = d.contenido_('Flujo de planificación · Automatizaciones', 'To-be', notas=(
        'El mismo flujo con lo que hay que automatizar en cada planta: automatizaciones de sistema y agentes que '
        'proponen y dejan la decisión a la persona. El algoritmo de combinación se afina en papel y se implanta en '
        'cartón.'))
    fases(sl)
    ys = carriles(sl)
    for (clave, *_), y in zip(CARRILES, ys):
        for i, grupos in enumerate(AUTOMATIZACIONES[clave]):
            x = X0 + i * (CW + GAP)
            yy = y + 0.02
            if not grupos:
                texto(sl, x, y + 0.1, CW, 0.3, 'Sin automatización', tam=8, cursiva=True, color=MUTED,
                      alin=PP_ALIGN.CENTER)
            for act, items in grupos:
                texto(sl, x + 0.08, yy, CW - 0.16, 0.25, act, tam=7.5, negrita=True, color=TEAL,
                      ancla=MSO_ANCHOR.MIDDLE, interlineado=1.0)
                yy += 0.25
                for tipo, txt in items:
                    alg = tipo == ALGORITMO
                    caja(sl, x + 0.06, yy, CW - 0.12, 0.42, relleno=NAVY if alg else BLANCO,
                         borde=None if alg else LINEA)
                    ic = 'robot' if tipo == AGENTE else 'engranajes'
                    sl.shapes.add_picture(str(ICONOS / f'{ic}_{"b" if alg else "n"}.png'), Inches(x + 0.11),
                                          Inches(yy + 0.12), Inches(0.17), Inches(0.17))
                    texto(sl, x + 0.33, yy, CW - 0.42, 0.42, txt, tam=7.5, negrita=alg, color=BLANCO if alg else NAVY,
                          ancla=MSO_ANCHOR.MIDDLE, interlineado=1.0)
                    yy += 0.47
    # leyenda en el hueco sobre las cabeceras de carril
    for k, (ic, lab) in enumerate((('engranajes', 'Automatización'), ('robot', 'Agente'))):
        yl = YT + 0.25 + k * 0.26
        sl.shapes.add_picture(str(ICONOS / f'{ic}_n.png'), Inches(XH + 0.08), Inches(yl + 0.02), Inches(0.17),
                              Inches(0.17))
        texto(sl, XH + 0.32, yl, 1.6, 0.22, lab, tam=8, color=MUTED, ancla=MSO_ANCHOR.MIDDLE)


def organigrama(d):
    sl = d.contenido_('Organigrama to-be', 'To-be', notas=(
        'La planificación se organiza por material: un planificador de papel y otro de cartón, coordinados por el '
        'Planificador Máster.'))
    mw, mh = 4.2, 1.15
    mx, my = L + (W - mw) / 2, 1.75
    m = caja(sl, mx, my, mw, mh, relleno=NAVY)
    icono(sl, 'usuario', mx + 0.3, my + (mh - 0.6) / 2, d=0.6, fondo=AZUL, variante='n')
    texto(sl, mx + 1.1, my, mw - 1.3, mh, [[('Planificador Máster', {'negrita': True, 'tam': 16})],
                                           [('Papel y cartón · todas las plantas', {'tam': 10, 'color': AZUL})]],
          color=BLANCO, ancla=MSO_ANCHOR.MIDDLE, interlineado=1.05, espacio=2)
    hijos = [('industria', 'Planificador de papel', ['MGT', 'PDA', 'Papresa']),
             ('caja', 'Planificador de cartón', ['Ondupack Almendralejo', 'Ondupack Navalmoral'])]
    hw, hh, hy = 4.6, 1.15, 3.75
    xs = [L + 0.6, R - 0.6 - hw]
    bus = my + mh + 0.4
    linea(sl, mx + mw / 2, my + mh, mx + mw / 2, bus, color=NAVY, grosor=1.5)
    linea(sl, xs[0] + hw / 2, bus, xs[1] + hw / 2, bus, color=NAVY, grosor=1.5)
    for (ic, tit, plantas), x in zip(hijos, xs):
        linea(sl, x + hw / 2, bus, x + hw / 2, hy, color=NAVY, grosor=1.5)
        caja(sl, x, hy, hw, hh, relleno=BLANCO, borde=NAVY)
        icono(sl, ic, x + 0.3, hy + (hh - 0.6) / 2, d=0.6)
        texto(sl, x + 1.1, hy, hw - 1.3, hh, tit, tam=15, negrita=True, ancla=MSO_ANCHOR.MIDDLE)
        # plantas que planifica
        pw, pg = (hw - (len(plantas) - 1) * 0.12) / len(plantas), 0.12
        for j, planta in enumerate(plantas):
            c = caja(sl, x + j * (pw + pg), hy + hh + 0.25, pw, 0.5, relleno=HIELO)
            rellenar(c.text_frame, planta, tam=9.5, negrita=True, color=NAVY, alin=PP_ALIGN.CENTER,
                     ancla=MSO_ANCHOR.MIDDLE, margen=0.05, interlineado=1.0)


def main():
    d = Deck()
    flujo_asis(d)
    flujo_automatizaciones(d)
    organigrama(d)
    d.guardar(SALIDA)
    print(SALIDA)


if __name__ == '__main__':
    main()
