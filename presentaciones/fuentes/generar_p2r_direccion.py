#!/usr/bin/env python3
"""Presentación para dirección: Plan to Release en MGT y Ondupack Almendralejo (situación actual, to-be y equipo).

Usa la plantilla corporativa de Planificacion_Packaging_CL.pptx: clona su portada y su separador y monta el contenido
sobre su diseño con cabecera. Los recuentos de intervenciones salen de los maestros de P2R.

Uso: python3 presentaciones/fuentes/generar_p2r_direccion.py
"""
import copy
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

AQUI = Path(__file__).resolve().parent
RAIZ = AQUI.parent.parent
sys.path.insert(0, str(RAIZ / 'maestros' / 'fuentes'))
from hojas_empresa import INTERVENCIONES  # noqa: E402

PLANTILLA = RAIZ / 'presentaciones' / 'Planificacion_Packaging_CL.pptx'
SALIDA = RAIZ / 'presentaciones' / 'P2R_MGT_Ondupack_Direccion.pptx'
ICONOS = AQUI / 'iconos'


def rgb(h):
    return RGBColor.from_string(h)


NAVY, AZUL, TEAL, CARD, GRIS, HIELO, MUTED, BLANCO, LINEA = map(
    rgb, ('00263A', '98C9EB', '244C5A', 'F5F7F8', '9AA9B4', 'EAF3FA', '5B6B78', 'FFFFFF', 'D2DCE5'))
FUENTE = 'Montserrat'
L, R = 0.6, 12.73  # márgenes del contenido (pulgadas)
W = R - L


# ------------------------------------------------------------------ utilidades de dibujo
def caja(sl, x, y, w, h, relleno=None, borde=None, forma=MSO_SHAPE.ROUNDED_RECTANGLE, radio=0.07, nombre=None):
    s = sl.shapes.add_shape(forma, Inches(x), Inches(y), Inches(w), Inches(h))
    if relleno is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = relleno
    if borde is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = borde
        s.line.width = Pt(0.75)
    if forma == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = min(0.5, radio / min(w, h))
    s.shadow.inherit = False
    if nombre:
        s.name = nombre
    return s


def rellenar(tf, parrafos, tam=11, color=NAVY, negrita=False, cursiva=False, alin=PP_ALIGN.LEFT, ancla=MSO_ANCHOR.TOP,
             interlineado=None, espacio=0, margen=0.0):
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(margen)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = ancla
    if isinstance(parrafos, str):
        parrafos = [parrafos]
    for i, par in enumerate(parrafos):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = alin
        if interlineado:
            p.line_spacing = interlineado
        if espacio and i < len(parrafos) - 1:
            p.space_after = Pt(espacio)
        runs = [(par, {})] if isinstance(par, str) else par
        for txt, o in runs:
            r = p.add_run()
            r.text = txt
            f = r.font
            f.name = FUENTE
            f.size = Pt(o.get('tam', tam))
            f.bold = o.get('negrita', negrita)
            f.italic = o.get('cursiva', cursiva)
            f.color.rgb = o.get('color', color)
            if o.get('espaciado'):
                r._r.get_or_add_rPr().set('spc', str(o['espaciado']))
            if o.get('simbolo'):  # flecha de Wingdings, como en la versión revisada
                rpr = r._r.get_or_add_rPr()
                rpr.append(rpr.makeelement(qn('a:sym'), {'typeface': 'Wingdings', 'pitchFamily': '2', 'charset': '2'}))
    return tf


def texto(sl, x, y, w, h, parrafos, nombre=None, **kw):
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    rellenar(tb.text_frame, parrafos, **kw)
    if nombre:
        tb.name = nombre
    return tb


def rotulo(sl, x, y, w, txt, color=TEAL, alin=PP_ALIGN.LEFT):
    """Etiqueta pequeña en versalitas, como en la plantilla."""
    return texto(sl, x, y, w, 0.22, [[(txt.upper(), {'espaciado': 100})]], tam=8.5, negrita=True, color=color, alin=alin)


def pastilla(sl, x, y, w, h, txt, relleno=AZUL, color=NAVY, tam=8, borde=None):
    s = caja(sl, x, y, w, h, relleno=relleno, borde=borde, radio=h / 2)
    rellenar(s.text_frame, [[(txt, {'espaciado': 50})]], tam=tam, negrita=True, color=color, alin=PP_ALIGN.CENTER,
             ancla=MSO_ANCHOR.MIDDLE)
    return s


def icono(sl, nombre, x, y, d=0.48, fondo=NAVY, variante='b', escala=0.5):
    caja(sl, x, y, d, d, relleno=fondo, forma=MSO_SHAPE.OVAL)
    t = d * escala
    sl.shapes.add_picture(str(ICONOS / f'{nombre}_{variante}.png'), Inches(x + (d - t) / 2), Inches(y + (d - t) / 2),
                          Inches(t), Inches(t))


def linea(sl, x1, y1, x2, y2, color=LINEA, grosor=0.75, discontinua=False):
    c = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(grosor)
    if discontinua:
        ln = c.line._get_or_add_ln()
        d = ln.makeelement(qn('a:prstDash'), {'val': 'dash'})
        ln.append(d)
    return c


# ------------------------------------------------------------------ diapositivas base
class Deck:
    def __init__(self):
        self.prs = Presentation(str(PLANTILLA))
        self.originales = len(self.prs.slides)
        self.portada_src = self.prs.slides[0]
        self.separador_src = self.prs.slides[2]
        self.contenido = next(l for l in self.prs.slide_layouts if l.name == '2_Imagen y contenido 3')
        self.n = 0

    def _clonar(self, src):
        dst = self.prs.slides.add_slide(src.slide_layout)
        for ph in list(dst.placeholders):
            ph._element.getparent().remove(ph._element)
        for sh in src.shapes:
            el = copy.deepcopy(sh._element)
            for blip in el.iter(qn('a:blip')):
                rid = blip.get(qn('r:embed'))
                rel = src.part.rels[rid]
                blip.set(qn('r:embed'), dst.part.relate_to(rel.target_part, rel.reltype))
            dst.shapes._spTree.insert_element_before(el, 'p:extLst')
        self.n += 1
        return dst

    def portada(self, titulo, subtitulo, pie):
        sl = self._clonar(self.portada_src)
        tf = next(sh for sh in sl.shapes if sh.name == 'Título').text_frame
        for p, txt in zip(tf.paragraphs, (titulo, subtitulo, pie)):
            p.runs[0].text = txt
            for r in p.runs[1:]:
                r.text = ''
        return sl

    def separador(self, num, titulo, subtitulo, tam_titulo=34):
        sl = self._clonar(self.separador_src)
        por_nombre = {sh.name: sh for sh in sl.shapes}
        por_nombre['Oval 5'].text_frame.paragraphs[0].runs[0].text = num
        r = por_nombre['TextBox 6'].text_frame.paragraphs[0].runs[0]
        r.text = titulo
        r.font.size = Pt(tam_titulo)
        por_nombre['TextBox 7'].text_frame.paragraphs[0].runs[0].text = subtitulo
        return sl

    def contenido_(self, titulo, etiqueta=None, notas=None):
        sl = self.prs.slides.add_slide(self.contenido)
        self.n += 1
        tf = sl.shapes.title.text_frame
        tf.text = titulo
        f = tf.paragraphs[0].runs[0].font
        f.name, f.size, f.bold, f.color.rgb = FUENTE, Pt(23), True, BLANCO
        if etiqueta:
            pastilla(sl, 8.05, 0.34, 2.45, 0.36, etiqueta.upper(), tam=8.5)
        texto(sl, 12.13, 6.62, 0.6, 0.25, str(self.n), tam=9, color=GRIS, alin=PP_ALIGN.RIGHT)
        if notas:
            sl.notes_slide.notes_text_frame.text = notas
        return sl

    def guardar(self):
        lst = self.prs.slides._sldIdLst
        for sid in list(lst)[:self.originales]:
            self.prs.part.drop_rel(sid.rId)
            lst.remove(sid)
        self.prs.save(str(SALIDA))


# ------------------------------------------------------------------ contenido
FASES = [('demanda', 'Demanda y previsión'), ('balanza', 'S&OP y Plan Maestro'),
         ('cubos', 'Necesidades de fabricación y MRP'), ('calendario', 'Programación'), ('play', 'Liberación'),
         ('panel', 'Seguimiento y control')]

ASIS = {
    'MGT': dict(
        stats=[('0 %', 'de la demanda tiene previsión, ni la externa ni la de Ondupack'),
               ('3 turnos', 'fijos a velocidad máxima, cargando a capacidad infinita'),
               ('Numerosos', 'correos al día para coordinar stock y cargas con Ondupack'),
               ('4 de 13', 'indicadores de Plan to Release se pueden medir hoy')],
        fases=[('Coordinación con Ondupack, que pide 10 días antes del inicio del ciclo. Sin previsión ni fecha '
                'propuesta.', 'alto'),
               ('Ciclo de 4 calidades con la cartera, a capacidad infinita, duración fija de los temporales de cada '
                'calidad.', 'alto'),
               ('MRP de materias primas en Excel; el de auxiliares, en SAP, se descuadra debido a stock no fiable.',
                'medio'),
               ('Algoritmo de combinación sin lo ya planificado; urgencias encajadas a mano.', 'alto'),
               ('El material se comprueba a mano en varios sitios y la orden se entrega en papel.', 'bajo'),
               ('Sin KPI de servicio: incidencias por llamada y Ondupack decide qué se carga.', 'medio')],
        funciona='planificación en SIGPA (web) y algoritmo de combinación por ciclos, en pruebas.'),
    'Ondupack': dict(
        stats=[('0 %', 'de la demanda repetitiva tiene previsión integrada en el sistema'),
               ('Nº 1', 'problema: la falta de materia prima; el papel se pide por consumo pasado'),
               ('Sin KPI', 'oficial de servicio: falta decidir cuál y contra qué fecha'),
               ('7 de 12', 'indicadores de Plan to Release se pueden medir hoy')],
        fases=[('Alta en GIO y fecha con CTP. Previsión fuera del sistema y sin registro de la demanda perdida.',
                'medio'),
               ('Carga-capacidad con escenarios de turnos y alerta de saturación, pero la sobrecarga se detecta tarde.',
                'alto'),
               ('Papel por consumo histórico, sustituciones a mano y consumos por inventario mensual.', 'alto'),
               ('Combinación de onduladora y converting sin optimizar; cambios del cliente por email.', 'alto'),
               ('Se libera con material pendiente y Compras persigue la falta.', 'sin'),
               ('Cierre automático y reunión semanal, pero sin KPI oficial de servicio.', 'medio')],
        funciona='fecha con CTP, carga-capacidad con alertas de saturación, OEE por máquina, escenarios SMED y '
                 'bloqueos con responsable y SLA.'),
}

# (código, palanca, qué cambia, tiene agente)
TOBE = {
    'MGT': [
        ('P2R-01', 'Gestión de la demanda',
         'Previsión de papel de Ondupack por ciclo, calidad y ancho, y de los clientes externos repetitivos.', False),
        ('P2R-02', 'S&OP y plan de capacidad',
         'Política de llenado y revisión mensual: turnos y ciclo ajustados a la cartera, con escenarios.', True),
        ('P2R-03', 'Política de inventario',
         'Reposición pull con Ondupack, con mínimo y máximo por referencia, y control del stock intercompany.', True),
        ('P2R-04', 'Promesa de fecha fiable',
         'Fecha propuesta al entrar el pedido, con un copiloto de plazos para Comercial.', True),
        ('P2R-05', 'Materiales por excepción',
         'MRP de materias primas en SAP, fuera del Excel, y propuestas de compra automáticas.', True),
        ('P2R-06', 'Programación estable',
         'Algoritmo que tiene en cuenta lo ya planificado y una fecha límite de cartera; urgencias con reglas.', True),
        ('P2R-07', 'Liberación controlada',
         'Disponibilidad de material por orden en el sistema, sin comprobaciones a mano.', False),
        ('P2R-08', 'Control del servicio',
         'KPI y pilar de servicio semanal; Ondupack ve el stock y las cargas sin correos.', True),
    ],
    'Ondupack Almendralejo': [
        ('P2R-01', 'Gestión de la demanda',
         'Previsión no vinculante de los clientes y referencias repetitivos, cargada en GIO para la carga y el CTP, '
         'y demanda perdida registrada.', False),
        ('P2R-02', 'S&OP y plan de capacidad',
         'Revisión mensual de capacidad y cuellos de botella con análisis de escenarios.', True),
        ('P2R-05', 'Materiales por excepción',
         'Necesidad de papel calculada desde el plan, stock fiable sin esperar al inventario y previsión de papel '
         'compartida con MGT, PDA y Papresa.', True),
        ('P2R-06', 'Programación estable',
         'Algoritmo de combinación de onduladora; urgencias y cambios gestionados con reglas.', True),
        ('P2R-08', 'Control del servicio',
         'KPI oficial de servicio y reunión semanal con causas; estado de pedidos y causa raíz automáticos.', True),
    ],
}

def contar(empresa):
    tipos = [it[2] for it in INTERVENCIONES[empresa]]
    return len(tipos), sum(t == 'agente' for t in tipos)


# ------------------------------------------------------------------ diapositivas
FLECHA = ''  # flecha de Wingdings (se dibuja con a:sym), como en la versión revisada
JUST = PP_ALIGN.JUSTIFY


def personas_resumen(lista):
    return [[(nombre + ' ', {}), (f'{FLECHA} {rol}', {'simbolo': True})] for nombre, rol in lista]


def resumen(d):
    n = contar('MGT')[0] + contar('Ondupack Almendralejo')[0]
    sl = d.contenido_('Resumen ejecutivo', 'Plan to Release', notas=(
        'Mensaje: las dos plantas planifican hoy con la cartera y a mano. El to-be no cambia de sistema: añade previsión, '
        'capacidad real y medición del servicio. Lo que más pesa es el equipo: pocas personas y muy concentradas.'))
    b = caja(sl, L, 1.45, W, 0.78, relleno=NAVY)
    rellenar(b.text_frame, [[('En resumen:  ', {'color': AZUL}),
                             ('las dos plantas planifican solo con la cartera y a mano. El to-be planifica con previsión, '
                              f'capacidad real y servicio medido: 13 palancas y {n} intervenciones.', {'color': BLANCO})]],
             tam=12.5, negrita=True, ancla=MSO_ANCHOR.MIDDLE, margen=0.3, interlineado=1.1)
    cols = [
        ('industria', 'MGT', 'Papelera · suministra papel a Ondupack y a clientes externos', [
            ('Hoy', 'Sin previsión. Carga a capacidad infinita con 3 turnos fijos, stock de Ondupack acumulado en MGT '
                    'y coordinación por correo.'),
            ('To-be', 'Previsión de Ondupack y de clientes repetitivos, ciclo ajustado a la cartera, reposición pull '
                      'con mínimo y máximo y servicio medido.'),
            ('Equipo', personas_resumen(EQUIPO_MGT))]),
        ('caja', 'Ondupack Almendralejo', 'Cartonera · onduladora y converting', [
            ('Hoy', 'Planta madura (CTP, OEE, escenarios SMED, carga capacidad), pero sin previsión, con la falta de '
                    'papel como problema nº 1. Algunos materiales tienen previsión, pero no está integrada en el '
                    'sistema.'),
            ('To-be', 'Previsión de la demanda repetitiva, análisis de escenarios de capacidad, papel calculado desde '
                      'el plan y KPI oficial de servicio.'),
            ('Equipo', personas_resumen(EQUIPO_ONDUPACK))]),
    ]
    cw, gap, y0, h = (W - 0.3) / 2, 0.3, 2.45, 3.95
    for k, (ic, nom, sub, filas) in enumerate(cols):
        x = L + k * (cw + gap)
        caja(sl, x, y0, cw, h, relleno=CARD)
        icono(sl, ic, x + 0.25, y0 + 0.22)
        texto(sl, x + 0.88, y0 + 0.2, cw - 1.1, 0.3, nom, tam=14, negrita=True)
        texto(sl, x + 0.88, y0 + 0.5, cw - 1.1, 0.25, sub, tam=9, color=MUTED, cursiva=True)
        for i, (lab, txt) in enumerate(filas):
            yy = y0 + 0.98 + i * 0.98
            if i:
                linea(sl, x + 0.25, yy - 0.12, x + cw - 0.25, yy - 0.12)
            rotulo(sl, x + 0.25, yy + 0.02, 0.9, lab)
            texto(sl, x + 1.2, yy, cw - 1.45, 0.82, txt, tam=10.5, interlineado=1.05, alin=JUST)


def asis(d, empresa, titulo):
    a = ASIS[empresa]
    sl = d.contenido_(titulo, 'Situación actual', notas=(
        'Arriba, las cifras que resumen la planta. Abajo, cómo funciona hoy cada fase de Plan to Release y el impacto '
        'de sus problemas (escala 1 a 5 de los hallazgos: alto = 4 o 5).'))
    cw, gap = (W - 3 * 0.23) / 4, 0.23
    for i, (num, lab) in enumerate(a['stats']):
        x = L + i * (cw + gap)
        caja(sl, x, 1.45, cw, 1.1, relleno=HIELO)
        texto(sl, x + 0.2, 1.53, cw - 0.4, 0.5, num, tam=26, negrita=True)
        texto(sl, x + 0.2, 2.03, cw - 0.4, 0.45, lab, tam=9, color=TEAL, interlineado=1.0)
    rotulo(sl, L, 2.78, 6, 'Cómo funciona hoy, fase a fase')
    fw, fg, y0, fh = (W - 5 * 0.15) / 6, 0.15, 3.07, 2.8
    estilos = {'alto': ('IMPACTO ALTO', NAVY, BLANCO, None), 'medio': ('IMPACTO MEDIO', AZUL, NAVY, None),
               'bajo': ('IMPACTO BAJO', BLANCO, MUTED, LINEA), 'sin': ('SIN HALLAZGO', BLANCO, MUTED, LINEA)}
    for i, ((ic, fase), (txt, imp)) in enumerate(zip(FASES, a['fases'])):
        x = L + i * (fw + fg)
        caja(sl, x, y0, fw, fh, relleno=CARD)
        icono(sl, ic, x + 0.15, y0 + 0.15, d=0.42)
        texto(sl, x + 0.67, y0 + 0.2, 0.4, 0.3, str(i + 1), tam=9, color=GRIS, negrita=True)
        texto(sl, x + 0.15, y0 + 0.66, fw - 0.3, 0.42, fase, tam=10, negrita=True, interlineado=1.0)
        texto(sl, x + 0.15, y0 + 1.13, fw - 0.3, 1.2, txt, tam=10, color=TEAL, interlineado=1.05, alin=JUST)
        lab, rel, col, bor = estilos[imp]
        pastilla(sl, x + 0.15, y0 + fh - 0.4, fw - 0.3, 0.26, lab, relleno=rel, color=col, tam=7.5, borde=bor)
    yb = y0 + fh + 0.14
    caja(sl, L, yb, W, 0.42, relleno=HIELO)
    icono(sl, 'check', L + 0.12, yb + 0.06, d=0.3, escala=0.5)
    texto(sl, L + 0.55, yb, W - 0.7, 0.42, [[('Ya funciona: ', {'negrita': True}), (a['funciona'], {})]], tam=10,
          ancla=MSO_ANCHOR.MIDDLE)


def tobe(d, empresa, titulo, alto_fila, empezar=None):
    filas = TOBE[empresa]
    n, ag = contar(empresa)
    sl = d.contenido_(titulo, f'To-be · {len(filas)} palancas', notas=(
        'Cada palanca es un cambio en la forma de planificar. El icono marca las palancas que incluyen un agente '
        '(automatización con IA que propone y la persona aprueba).'))
    texto(sl, L, 1.38, W, 0.3, [[(f'{n} intervenciones', {'negrita': True}),
                                 (f':  {n - ag} de negocio y software y {ag} agentes', {})]], tam=11)
    x_que, x_ag = 3.35, 11.75
    rotulo(sl, L + 0.15, 1.8, 1.6, 'Palanca', color=MUTED)
    rotulo(sl, x_que, 1.8, 1.6, 'Qué cambia', color=MUTED)
    rotulo(sl, x_ag, 1.8, 0.9, 'Agente', color=MUTED, alin=PP_ALIGN.CENTER)
    gap = 0.07
    for i, (cod, nom, que, agente) in enumerate(filas):
        y = 2.08 + i * (alto_fila + gap)
        caja(sl, L, y, W, alto_fila, relleno=CARD)
        pastilla(sl, L + 0.15, y + (alto_fila - 0.24) / 2, 0.72, 0.24, cod, relleno=NAVY, color=BLANCO, tam=7.5)
        texto(sl, L + 0.98, y, 1.65, alto_fila, nom, tam=10, negrita=True, ancla=MSO_ANCHOR.MIDDLE, interlineado=1.0)
        texto(sl, x_que, y, x_ag - x_que - 0.25, alto_fila, que, tam=10.5, color=TEAL, ancla=MSO_ANCHOR.MIDDLE,
              interlineado=1.0)
        if agente:
            sl.shapes.add_picture(str(ICONOS / 'robot_n.png'), Inches(x_ag + 0.45 - 0.13),
                                  Inches(y + (alto_fila - 0.26) / 2), Inches(0.26), Inches(0.26))
    if empezar:
        yb = 2.08 + len(filas) * (alto_fila + gap) + 0.12
        rotulo(sl, L, yb, 4, 'Por dónde empezar')
        bw = (W - 2 * 0.2) / 3
        for i, (tit, txt) in enumerate(empezar):
            x = L + i * (bw + 0.2)
            caja(sl, x, yb + 0.3, bw, 0.8, relleno=HIELO)
            caja(sl, x + 0.15, yb + 0.44, 0.34, 0.34, relleno=NAVY, forma=MSO_SHAPE.OVAL)
            texto(sl, x + 0.15, yb + 0.44, 0.34, 0.34, str(i + 1), tam=10, negrita=True, color=AZUL, alin=PP_ALIGN.CENTER,
                  ancla=MSO_ANCHOR.MIDDLE)
            texto(sl, x + 0.62, yb + 0.38, bw - 0.75, 0.66, [[(tit, {'negrita': True})], txt], tam=9.5,
                  interlineado=1.0, espacio=2)


def equipo(d, empresa, titulo, personas, palancas, apoyo, atencion):
    sl = d.contenido_(titulo, f'Equipo · {len(personas)} personas', notas=(
        'Personas de la planta censadas en el proyecto. La barra indica en cuántas palancas participa cada una: '
        'oscuro = la lidera, claro = la apoya.'))
    n = len(personas)
    gap = 0.2
    cw = (W - (n - 1) * gap) / n
    y0, h = 1.45, 3.35
    for i, (ini, nom, rol, lidera, apoya) in enumerate(personas):
        x = L + i * (cw + gap)
        caja(sl, x, y0, cw, h, relleno=CARD)
        caja(sl, x + 0.2, y0 + 0.2, 0.62, 0.62, relleno=NAVY, forma=MSO_SHAPE.OVAL)
        texto(sl, x + 0.2, y0 + 0.2, 0.62, 0.62, ini, tam=13, negrita=True, color=BLANCO, alin=PP_ALIGN.CENTER,
              ancla=MSO_ANCHOR.MIDDLE)
        texto(sl, x + 0.2, y0 + 0.95, cw - 0.4, 0.45, nom, tam=11.5, negrita=True, interlineado=1.0)
        texto(sl, x + 0.2, y0 + 1.42, cw - 0.4, 0.4, rol, tam=9, color=MUTED, cursiva=True, interlineado=1.0)
        # barra de palancas
        sw = min(0.3, (cw - 0.4 - (len(palancas) - 1) * 0.05) / len(palancas))
        codigos = [c for c, _ in palancas]
        for j, p in enumerate(codigos):
            col = NAVY if p in lidera else AZUL if p in apoya else LINEA
            caja(sl, x + 0.2 + j * (sw + 0.05), y0 + 1.92, sw, 0.14, relleno=col, forma=MSO_SHAPE.RECTANGLE)
        part = len([p for p in codigos if p in lidera or p in apoya])
        texto(sl, x + 0.2, y0 + 2.1, cw - 0.4, 0.22, f'En {part} de {len(palancas)} palancas', tam=8.5, color=MUTED)
        nombres = dict(palancas)
        lid = ' · '.join(nombres[p] for p in codigos if p in lidera) or '—'
        apy = ' · '.join(nombres[p] for p in codigos if p in apoya) or '—'
        texto(sl, x + 0.2, y0 + 2.4, cw - 0.4, 0.9, [[('Lidera: ', {'negrita': True}), (lid, {})],
                                                     [('Apoya: ', {'negrita': True}), (apy, {})]],
              tam=9, interlineado=1.0, espacio=3)
    # leyenda
    yl = y0 + h + 0.1
    for k, (col, lab) in enumerate(((NAVY, 'lidera'), (AZUL, 'apoya'), (LINEA, 'no participa'))):
        caja(sl, L + k * 1.35, yl + 0.05, 0.18, 0.12, relleno=col, forma=MSO_SHAPE.RECTANGLE)
        texto(sl, L + k * 1.35 + 0.25, yl, 1.1, 0.22, lab, tam=8.5, color=MUTED, ancla=MSO_ANCHOR.MIDDLE)
    yb, bh = yl + 0.4, 6.4 - (yl + 0.4)
    bw = (W - 0.3) / 2
    caja(sl, L, yb, bw, bh, relleno=HIELO)
    icono(sl, 'equipo', L + 0.2, yb + 0.2, d=0.42)
    texto(sl, L + 0.78, yb + 0.2, bw - 1.0, bh - 0.3, [[('Apoyo de otras áreas', {'negrita': True, 'tam': 11})], apoyo],
          tam=10, color=TEAL, interlineado=1.05, espacio=3)
    caja(sl, L + bw + 0.3, yb, bw, bh, relleno=NAVY)
    icono(sl, 'bandera', L + bw + 0.5, yb + 0.2, d=0.42, fondo=AZUL, variante='n')
    texto(sl, L + bw + 1.08, yb + 0.2, bw - 1.3, bh - 0.3,
          [[('A tener en cuenta', {'negrita': True, 'tam': 11, 'color': AZUL})], atencion],
          tam=10, color=BLANCO, interlineado=1.05, espacio=3)


def intercompany(d):
    sl = d.contenido_('MGT y Ondupack · Intercompany', 'Papel intercompany', notas=(
        'Ondupack es el principal cliente de MGT. Hoy la relación se gestiona por correo y con pedidos firmes; el to-be '
        'añade previsión y reposición pull, y encaja con la integración intercompany en SAP que ya está en marcha.'))
    xh, wh, xa, xt, wt = 2.95, 4.4, 7.5, 8.05, 4.68
    rotulo(sl, xh + 0.2, 1.5, 3, 'Hoy', color=MUTED)
    rotulo(sl, xt + 0.2, 1.5, 3, 'To-be', color=TEAL)
    filas = [
        ('demanda', 'Demanda', 'Solo el pedido firme, 10 días antes del ciclo; las fechas cambian después.',
         'Previsión de papel por ciclo, calidad y ancho, más el pedido firme 10 días antes.'),
        ('almacen', 'Stock', 'Lo fabricado de más se queda en MGT y Ondupack lo retira poco a poco.',
         'Reposición pull con mínimo y máximo por referencia.'),
        ('correo', 'Coordinación', 'Numerosos correos al día; Ondupack decide qué se carga al día siguiente.',
         'Ondupack ve el stock y las cargas en el sistema y un agente propone la carga.'),
        ('panel', 'Medición', 'Sin KPI de servicio a Ondupack y con restos de pedidos antiguos sin cerrar.',
         'OTIF a Ondupack y stock intercompany, revisados cada semana.'),
    ]
    for i, (ic, cat, hoy, tb) in enumerate(filas):
        y = 1.8 + i * 1.0
        icono(sl, ic, L, y + 0.17, d=0.5)
        texto(sl, L + 0.65, y, 1.6, 0.84, cat, tam=12, negrita=True, ancla=MSO_ANCHOR.MIDDLE)
        c = caja(sl, xh, y, wh, 0.84, relleno=CARD)
        rellenar(c.text_frame, hoy, tam=10.5, color=TEAL, ancla=MSO_ANCHOR.MIDDLE, margen=0.2, interlineado=1.05)
        f = caja(sl, xa, y + 0.27, 0.4, 0.3, relleno=AZUL, forma=MSO_SHAPE.CHEVRON)
        f.adjustments[0] = 0.5
        c = caja(sl, xt, y, wt, 0.84, relleno=NAVY)
        rellenar(c.text_frame, tb, tam=10.5, color=BLANCO, negrita=True, ancla=MSO_ANCHOR.MIDDLE, margen=0.2,
                 interlineado=1.05)
    texto(sl, L, 5.95, W, 0.3, [[('Nota: ', {'negrita': True}),
                                 ('encaja con la integración intercompany en SAP (mínimos y máximos y pedidos automáticos) '
                                  'que ya está en marcha con MGT, PDA y Papresa.', {})]], tam=9, cursiva=True, color=MUTED)


def decisiones(d):
    sl = d.contenido_('Qué necesitamos de dirección', 'Decisiones', notas=(
        'Cinco decisiones para arrancar. Las tres primeras son de personas: sin owner ni dedicación el despliegue no '
        'se puede dimensionar ni medir.'))
    items = [
        ('usuario', 'Nombrar el process owner de P2R en MGT',
         'La celda de MGT está sin activar: sin responsable no arranca el despliegue.'),
        ('reloj', 'Asignar la dedicación de cada persona a P2R',
         'Falta en MGT y en Ondupack: sin ella no se dimensiona el equipo ni el ahorro de los agentes.'),
        ('equipo', 'Reforzar la planificación de MGT',
         'Joan Carles es el único planificador y está en las 8 palancas: necesita respaldo o liberar tiempo.'),
        ('diana', 'Decidir el KPI oficial de servicio de Ondupack',
         'Dirección y Comercial: qué OTIF usar y contra qué fecha (la del CTP).'),
        ('euro', 'Validar la siguiente fase: cuantificar en euros',
         'Valorar cada hallazgo y el compromiso de valor de cada palanca para priorizar con números.'),
    ]
    for i, (ic, tit, txt) in enumerate(items):
        y = 1.5 + i * 0.95
        caja(sl, L, y, W, 0.8, relleno=CARD)
        caja(sl, L + 0.18, y + 0.17, 0.46, 0.46, relleno=NAVY, forma=MSO_SHAPE.OVAL)
        texto(sl, L + 0.18, y + 0.17, 0.46, 0.46, str(i + 1), tam=13, negrita=True, color=AZUL, alin=PP_ALIGN.CENTER,
              ancla=MSO_ANCHOR.MIDDLE)
        texto(sl, L + 0.85, y + 0.1, 10.3, 0.6, [[(tit, {'negrita': True, 'tam': 12.5})], [(txt, {'color': TEAL})]],
              tam=10.5, interlineado=1.0, espacio=2, ancla=MSO_ANCHOR.MIDDLE)
        icono(sl, ic, R - 0.68, y + 0.17, d=0.46, fondo=HIELO, variante='n')


# personas (nombre, función) tal como figuran en la versión revisada
EQUIPO_MGT = [('Joan Carles Girones', 'Planificador + Aprovisionamiento Auxiliares'),
              ('Marta Boyo', 'Aprovisionamiento MMPP'), ('Albert Fontanet', 'Comercial')]
EQUIPO_ONDUPACK = [('María Teresa Espinosa', 'Carga Capacidad + Planificación'),
                   ('Miguel Ángel Rosado', 'Secuenciador'), ('Angélica Lemus', 'Aprovisionamiento MMPP')]


def main():
    d = Deck()
    d.portada('Plan to Release', 'MGT y Ondupack Almendralejo · hoy, to-be y equipo',
              'Comité de dirección  |  Noviembre 2026')
    resumen(d)
    rol = dict(EQUIPO_MGT + EQUIPO_ONDUPACK)

    d.separador('01', 'MGT', 'Papelera · suministra papel a Ondupack y a clientes externos')
    asis(d, 'MGT', 'MGT · Situación actual')
    tobe(d, 'MGT', 'MGT · To-be', 0.46)
    pal_mgt = [('P2R-01', 'Demanda'), ('P2R-02', 'Capacidad'), ('P2R-03', 'Inventario'), ('P2R-04', 'Fecha'),
               ('P2R-05', 'Materiales'), ('P2R-06', 'Programación'), ('P2R-07', 'Liberación'), ('P2R-08', 'Servicio')]
    equipo(d, 'MGT', 'MGT · Equipo disponible', [
        ('JC', 'Joan Carles Girones', rol['Joan Carles Girones'],
         {'P2R-02', 'P2R-03', 'P2R-06', 'P2R-07'}, {'P2R-01', 'P2R-04', 'P2R-05', 'P2R-08'}),
        ('MB', 'Marta Boyo', rol['Marta Boyo'], {'P2R-05'}, {'P2R-07'}),
        ('AF', 'Albert Fontanet', rol['Albert Fontanet'], {'P2R-01', 'P2R-04', 'P2R-08'}, {'P2R-02', 'P2R-03'}),
    ], pal_mgt,
        'Ramón, en el algoritmo de combinación · Sistemas / IT en MRP, liberación y visibilidad de stock · Angélica '
        'Lemus (Compras de Ondupack) en lo intercompany.',
        'Joan Carles es el único planificador y está en las 8 palancas. Falta nombrar el process owner de la celda y '
        'cargar la dedicación de cada persona.')

    d.separador('02', 'Ondupack Almendralejo', 'Cartonera · onduladora y converting', tam_titulo=28)
    asis(d, 'Ondupack', 'Ondupack · Situación actual')
    tobe(d, 'Ondupack Almendralejo', 'Ondupack · To-be', 0.5, empezar=[
        ('Capacidad y combinación', 'Escenarios de capacidad y algoritmo de combinación: son la base de la carga y la '
                                    'secuencia.'),
        ('Bloque de materiales', 'Ataca la falta de papel, su problema nº 1.'),
        ('Agentes, al final', 'Cuando el proceso esté estable y Ondupack trabaje en GIO web.'),
    ])
    pal_op = [('P2R-01', 'Demanda'), ('P2R-02', 'Capacidad'), ('P2R-05', 'Materiales'), ('P2R-06', 'Programación'),
              ('P2R-08', 'Servicio')]
    equipo(d, 'Ondupack', 'Ondupack · Equipo disponible', [
        ('MT', 'María Teresa Espinosa', rol['María Teresa Espinosa'], {'P2R-08'}, {'P2R-01'}),
        ('MA', 'Miguel Ángel Rosado', rol['Miguel Ángel Rosado'], {'P2R-01', 'P2R-06'},
         {'P2R-02', 'P2R-05', 'P2R-08'}),
        ('AL', 'Angélica Lemus', rol['Angélica Lemus'], {'P2R-05'}, set()),
    ], pal_op,
        'Karim Karach (Producción, Navalmoral) en capacidad · Sistemas / IT en el algoritmo de combinación. '
        'Angélica apoya además a MGT en lo intercompany.',
        'Miguel Ángel está en las 5 palancas. Solo María Teresa tiene dedicación cargada (10 %). '
        'Process owner: Enrique Tejada.')

    d.separador('03', 'Siguientes pasos', 'Intercompany y decisiones')
    intercompany(d)
    decisiones(d)
    d.guardar()
    print(SALIDA)


if __name__ == '__main__':
    main()
