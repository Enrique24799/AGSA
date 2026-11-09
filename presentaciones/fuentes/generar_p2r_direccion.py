#!/usr/bin/env python3
"""Presentación para dirección: Plan to Release en MGT y Ondupack Almendralejo (situación actual, to-be y equipo).

Usa la plantilla corporativa de Planificacion_Packaging_CL.pptx: clona su portada y su separador y monta el contenido
sobre su diseño con cabecera. Las fechas de la hoja de ruta salen de las intervenciones de los maestros de P2R.

Uso: python3 presentaciones/fuentes/generar_p2r_direccion.py
"""
import calendar
import copy
import sys
from datetime import date, timedelta
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
from fichas_palancas import AGENTES, NEG, _fechas  # noqa: E402
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
               ('16', 'correos al día, aprox., para coordinar stock y cargas con Ondupack'),
               ('4 de 13', 'indicadores de Plan to Release se pueden medir hoy')],
        fases=[('Pedidos por O2C; Ondupack pide 10 días antes del ciclo. Sin previsión ni fecha propuesta.', 'alto'),
               ('Ciclo de 4 calidades con la cartera, a capacidad infinita. Lo sobrante queda en stock.', 'alto'),
               ('MRP de materias primas en Excel; el de auxiliares, en SAP, se descuadra por la báscula.', 'medio'),
               ('Combinación sin lo ya fabricado; calidad y ancho a mano; urgencias encajadas a mano.', 'alto'),
               ('El material se comprueba a mano en varios sitios y la orden se entrega en papel.', 'bajo'),
               ('Sin KPI de servicio: incidencias por llamada y Ondupack decide qué se carga.', 'medio')],
        funciona='planificación en SIGPA (web) y algoritmo de combinación por ciclos, en pruebas.'),
    'Ondupack': dict(
        stats=[('0 %', 'de la demanda repetitiva tiene previsión'),
               ('924', 'es el cuello de botella: no alcanza los metros planificados'),
               ('Nº 1', 'problema: la falta de materia prima; el papel se pide por consumo pasado'),
               ('7 de 12', 'indicadores de Plan to Release se pueden medir hoy')],
        fases=[('Alta en GIO y fecha con CTP. Sin previsión ni registro de la demanda perdida.', 'medio'),
               ('Escenarios de turnos y alerta de saturación, pero con una velocidad que la 924 no alcanza.', 'alto'),
               ('Papel por consumo histórico, sustituciones a mano y consumos por inventario mensual.', 'alto'),
               ('Combinación de onduladora y converting sin optimizar; cambios del cliente por email.', 'alto'),
               ('Se libera con material pendiente y Compras persigue la falta.', 'sin'),
               ('Cierre automático y reunión semanal, pero sin KPI oficial de servicio.', 'medio')],
        funciona='fecha con CTP, alertas de saturación, OEE por máquina, escenarios SMED y bloqueos con responsable y SLA.'),
}

# (código, palanca, qué cambia, responsable, tiene agente)
TOBE = {
    'MGT': [
        ('P2R-01', 'Gestión de la demanda',
         'Previsión de papel de Ondupack por ciclo, calidad y ancho, y de los clientes externos repetitivos.',
         'Albert Fontanet', False),
        ('P2R-02', 'S&OP y plan de capacidad',
         'Política de llenado y revisión mensual: turnos y ciclo ajustados a la cartera, con escenarios.',
         'Joan Carles Girones', True),
        ('P2R-03', 'Política de inventario',
         'Reposición pull con Ondupack, con mínimo y máximo por referencia, y control del stock intercompany.',
         'Joan Carles Girones', True),
        ('P2R-04', 'Promesa de fecha fiable',
         'Fecha propuesta al entrar el pedido, con un copiloto de plazos para Comercial.', 'Albert Fontanet', True),
        ('P2R-05', 'Materiales por excepción',
         'MRP de materias primas en SAP, fuera del Excel, y propuestas de compra automáticas.', 'Marta Boyo', True),
        ('P2R-06', 'Programación estable',
         'Algoritmo que tiene en cuenta lo ya fabricado y una fecha límite de cartera; urgencias con reglas.',
         'Joan Carles Girones', True),
        ('P2R-07', 'Liberación controlada',
         'Disponibilidad de material por orden en el sistema, sin comprobaciones a mano.', 'Joan Carles Girones', False),
        ('P2R-08', 'Control del servicio',
         'KPI y pilar de servicio semanal; Ondupack ve el stock y las cargas sin correos.', 'Albert Fontanet', True),
    ],
    'Ondupack Almendralejo': [
        ('P2R-01', 'Gestión de la demanda',
         'Previsión no vinculante de los clientes y referencias repetitivos, cargada en GIO para la carga y el CTP, '
         'y demanda perdida registrada.', 'Miguel Ángel Rosado', False),
        ('P2R-02', 'S&OP y plan de capacidad',
         'Revisión mensual de capacidad y cuellos de botella con la velocidad real de la 924, con escenarios.',
         'José María Baquero', True),
        ('P2R-05', 'Materiales por excepción',
         'Necesidad de papel calculada desde el plan, stock fiable sin esperar al inventario y previsión de papel '
         'compartida con MGT, PDA y Papresa.', 'Angélica Lemus', True),
        ('P2R-06', 'Programación estable',
         'Algoritmo de combinación de onduladora; urgencias y cambios gestionados con reglas.', 'Miguel Ángel Rosado',
         True),
        ('P2R-08', 'Control del servicio',
         'KPI oficial de servicio y reunión semanal con causas; estado de pedidos y causa raíz automáticos.',
         'Mª Teresa Espinosa', True),
    ],
}

MESES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic']


def tramos(empresa):
    """Por palanca: tramos (inicio, fin) de negocio y software y de agentes, uniendo los que se solapan."""
    res = {}
    for cod, _, tipo, pal, *_ in INTERVENCIONES[empresa]:
        base = (NEG.get((empresa, cod)) or AGENTES[(empresa, cod)])[0]
        ini, fin = _fechas(empresa, base)
        res.setdefault(pal, {'neg': [], 'agente': []})['agente' if tipo == 'agente' else 'neg'].append((ini, fin))
    for pal, d in res.items():
        for k, v in d.items():
            unidos = []
            for ini, fin in sorted(v):
                if unidos and ini <= unidos[-1][1] + timedelta(days=1):
                    unidos[-1] = (unidos[-1][0], max(unidos[-1][1], fin))
                else:
                    unidos.append((ini, fin))
            d[k] = unidos
    return res


def ventana(empresa, pal):
    t = tramos(empresa)[pal]
    todos = t['neg'] + t['agente']
    a, b = min(x[0] for x in todos), max(x[1] for x in todos)
    return f'{MESES[a.month - 1]} {a.year % 100} – {MESES[b.month - 1]} {b.year % 100}'


def contar(empresa):
    tipos = [it[2] for it in INTERVENCIONES[empresa]]
    return len(tipos), sum(t == 'agente' for t in tipos)


# ------------------------------------------------------------------ diapositivas
def resumen(d):
    n_mgt, a_mgt = contar('MGT')
    n_op, a_op = contar('Ondupack Almendralejo')
    sl = d.contenido_('Resumen ejecutivo', 'Plan to Release', notas=(
        'Mensaje: las dos plantas planifican hoy con la cartera y a mano. El to-be no cambia de sistema: añade previsión, '
        'capacidad real y medición del servicio. Lo que más pesa es el equipo: pocas personas y muy concentradas.'))
    b = caja(sl, L, 1.45, W, 0.78, relleno=NAVY)
    rellenar(b.text_frame, [[('En resumen:  ', {'color': AZUL}),
                             ('las dos plantas planifican solo con la cartera y a mano. El to-be planifica con previsión, '
                              f'capacidad real y servicio medido: 13 palancas y {n_mgt + n_op} intervenciones entre 2027 y 2028.',
                              {'color': BLANCO})]], tam=12.5, negrita=True, ancla=MSO_ANCHOR.MIDDLE, margen=0.3,
             interlineado=1.1)
    cols = [
        ('industria', 'MGT', 'Papelera · suministra papel a Ondupack y a clientes externos', [
            ('Hoy', 'Sin previsión. Carga a capacidad infinita con 3 turnos fijos, stock de Ondupack acumulado en MGT '
                    'y coordinación por correo, unos 16 al día.'),
            ('To-be', f'Previsión de Ondupack y de clientes repetitivos, ciclo ajustado a la cartera, reposición pull '
                      f'con mínimo y máximo y servicio medido. 8 palancas y {n_mgt} intervenciones ({a_mgt} agentes).'),
            ('Equipo', '3 personas. Joan Carles, único planificador, está en las 8 palancas y lidera 4. '
                       'Falta nombrar el process owner.')]),
        ('caja', 'Ondupack Almendralejo', 'Cartonera · onduladora y converting', [
            ('Hoy', 'La planta más madura (CTP, OEE, escenarios SMED), pero sin previsión, con la 924 saturada y con '
                    'la falta de papel como problema nº 1.'),
            ('To-be', f'Previsión de la demanda repetitiva, capacidad con la velocidad real de la 924, papel calculado '
                      f'desde el plan y KPI oficial de servicio. 5 palancas y {n_op} intervenciones ({a_op} agentes).'),
            ('Equipo', '5 personas, 2 en Planificación. Miguel Ángel está en las 5 palancas y lidera 2. '
                       'Process owner: Enrique Tejada.')]),
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
            texto(sl, x + 1.2, yy, cw - 1.45, 0.82, txt, tam=10.5, interlineado=1.05)


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
        texto(sl, x + 0.15, y0 + 1.13, fw - 0.3, 1.2, txt, tam=10, color=TEAL, interlineado=1.05)
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
        'Cada palanca es un cambio en la forma de planificar, con un responsable y una ventana de fechas. El icono marca '
        'las palancas que incluyen un agente (automatización con IA que propone y la persona aprueba).'))
    texto(sl, L, 1.38, W, 0.3, [[(f'{n} intervenciones', {'negrita': True}),
                                 (f':  {n - ag} de negocio y software y {ag} agentes', {})]], tam=11)
    cols = [('Palanca', L + 0.15), ('Qué cambia', 3.35), ('Agente', 8.9), ('Responsable', 9.6), ('Cuándo', 11.45)]
    for lab, x in cols:
        rotulo(sl, x, 1.8, 1.6, lab, color=MUTED)
    gap = 0.07
    for i, (cod, nom, que, resp, agente) in enumerate(filas):
        y = 2.08 + i * (alto_fila + gap)
        caja(sl, L, y, W, alto_fila, relleno=CARD)
        pastilla(sl, L + 0.15, y + (alto_fila - 0.24) / 2, 0.72, 0.24, cod, relleno=NAVY, color=BLANCO, tam=7.5)
        texto(sl, L + 0.98, y, 1.65, alto_fila, nom, tam=10, negrita=True, ancla=MSO_ANCHOR.MIDDLE, interlineado=1.0)
        texto(sl, 3.35, y, 5.35, alto_fila, que, tam=10, color=TEAL, ancla=MSO_ANCHOR.MIDDLE, interlineado=1.0)
        if agente:
            sl.shapes.add_picture(str(ICONOS / 'robot_n.png'), Inches(9.05), Inches(y + (alto_fila - 0.24) / 2),
                                  Inches(0.24), Inches(0.24))
        texto(sl, 9.6, y, 1.8, alto_fila, resp, tam=10, ancla=MSO_ANCHOR.MIDDLE, interlineado=1.0)
        texto(sl, 11.45, y, 1.2, alto_fila, ventana(empresa, cod), tam=9.5, color=MUTED, ancla=MSO_ANCHOR.MIDDLE)
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
    texto(sl, L, 1.38, W, 0.3, [[('Responsables: ', {'negrita': True}),
                                 ('Albert Fontanet y Joan Carles Girones en MGT · Angélica Lemus en Ondupack', {})]],
          tam=10.5, color=TEAL)
    xh, wh, xa, xt, wt = 2.95, 4.4, 7.5, 8.05, 4.68
    rotulo(sl, xh + 0.2, 1.82, 3, 'Hoy', color=MUTED)
    rotulo(sl, xt + 0.2, 1.82, 3, 'To-be', color=TEAL)
    filas = [
        ('demanda', 'Demanda', 'Solo el pedido firme, 10 días antes del ciclo; las fechas cambian después.',
         'Previsión de papel por ciclo, calidad y ancho, más el pedido firme 10 días antes.'),
        ('almacen', 'Stock', 'Lo fabricado de más se queda en MGT y Ondupack lo retira poco a poco.',
         'Reposición pull con mínimo y máximo por referencia.'),
        ('correo', 'Coordinación', 'Unos 16 correos al día; Ondupack decide qué se carga al día siguiente.',
         'Ondupack ve el stock y las cargas en el sistema y un agente propone la carga.'),
        ('panel', 'Medición', 'Sin KPI de servicio a Ondupack y con restos de pedidos antiguos sin cerrar.',
         'OTIF a Ondupack y stock intercompany, revisados cada semana.'),
    ]
    for i, (ic, cat, hoy, tb) in enumerate(filas):
        y = 2.1 + i * 0.98
        icono(sl, ic, L, y + 0.17, d=0.5)
        texto(sl, L + 0.65, y, 1.6, 0.84, cat, tam=12, negrita=True, ancla=MSO_ANCHOR.MIDDLE)
        c = caja(sl, xh, y, wh, 0.84, relleno=CARD)
        rellenar(c.text_frame, hoy, tam=10.5, color=TEAL, ancla=MSO_ANCHOR.MIDDLE, margen=0.2, interlineado=1.05)
        f = caja(sl, xa, y + 0.27, 0.4, 0.3, relleno=AZUL, forma=MSO_SHAPE.CHEVRON)
        f.adjustments[0] = 0.5
        c = caja(sl, xt, y, wt, 0.84, relleno=NAVY)
        rellenar(c.text_frame, tb, tam=10.5, color=BLANCO, negrita=True, ancla=MSO_ANCHOR.MIDDLE, margen=0.2,
                 interlineado=1.05)
    texto(sl, L, 6.06, W, 0.3, [[('Nota: ', {'negrita': True}),
                                 ('encaja con la integración intercompany en SAP (mínimos y máximos y pedidos automáticos) '
                                  'que ya está en marcha con MGT, PDA y Papresa.', {})]], tam=9, cursiva=True, color=MUTED)


def hoja_de_ruta(d):
    n = sum(contar(e)[0] for e in TOBE)
    sl = d.contenido_('Hoja de ruta 2027-2028', f'{n} intervenciones', notas=(
        'Fechas de las intervenciones dadas de alta. Primero lo de negocio y software, que ordena el proceso y los datos; '
        'los agentes llegan en 2028, cuando el proceso ya está estable.'))
    x0, x1 = 3.55, R
    mes = (x1 - x0) / 24

    def px(dt):
        dias = calendar.monthrange(dt.year, dt.month)[1]
        return x0 + ((dt.year - 2027) * 12 + dt.month - 1 + (dt.day - 1) / dias) * mes

    # cabecera de tiempo
    for k, anio in enumerate((2027, 2028)):
        xa = x0 + k * 12 * mes
        caja(sl, xa + 0.02, 1.4, 12 * mes - 0.04, 0.26, relleno=NAVY, radio=0.05)
        texto(sl, xa, 1.4, 12 * mes, 0.26, str(anio), tam=9, negrita=True, color=BLANCO, alin=PP_ALIGN.CENTER,
              ancla=MSO_ANCHOR.MIDDLE)
        for q in range(4):
            texto(sl, xa + q * 3 * mes, 1.7, 3 * mes, 0.2, f'T{q + 1}', tam=8.5, color=MUTED, alin=PP_ALIGN.CENTER)
    y_ini, fila = 2.0, 0.26
    grupos = [('Ondupack Almendralejo', 'Ondupack Almendralejo'), ('MGT', 'MGT')]
    filas_tot = sum(len(TOBE[e]) for e, _ in grupos) * fila + len(grupos) * 0.3
    for q in range(9):
        xq = x0 + q * 3 * mes
        linea(sl, xq, 1.92, xq, y_ini + filas_tot + 0.05, color=LINEA, grosor=0.5)
    y = y_ini
    for emp, lab in grupos:
        rotulo(sl, L, y + 0.04, 3, lab)
        y += 0.3
        t = tramos(emp)
        for cod, nom, *_ in TOBE[emp]:
            caja(sl, L, y + 0.02, x1 - L, fila - 0.04, relleno=CARD, forma=MSO_SHAPE.RECTANGLE)
            texto(sl, L + 0.1, y, 2.85, fila, [[(cod + '  ', {'negrita': True}), (nom, {})]], tam=8.5,
                  ancla=MSO_ANCHOR.MIDDLE)
            for tipo, col in (('neg', NAVY), ('agente', AZUL)):
                for a, b in t[cod][tipo]:
                    xa, xb = px(a), px(b + timedelta(days=1))
                    caja(sl, xa, y + 0.06, xb - xa, fila - 0.12, relleno=col, radio=0.04)
            y += fila
    # hito de agentes
    xm = px(date(2028, 4, 1))
    linea(sl, xm, 1.92, xm, y + 0.05, color=TEAL, grosor=1, discontinua=True)
    yl = y + 0.15
    for k, (col, lab) in enumerate(((NAVY, 'Negocio y software'), (AZUL, 'Agentes'))):
        caja(sl, L + k * 2.1, yl + 0.045, 0.3, 0.13, relleno=col, radio=0.03)
        texto(sl, L + k * 2.1 + 0.38, yl, 1.7, 0.22, lab, tam=9, color=MUTED, ancla=MSO_ANCHOR.MIDDLE)
    texto(sl, xm - 4.2, yl, 4.1, 0.22, 'Abril 2028: primeros agentes, en Ondupack', tam=9, color=TEAL, negrita=True,
          alin=PP_ALIGN.RIGHT, ancla=MSO_ANCHOR.MIDDLE)


def decisiones(d):
    sl = d.contenido_('Qué necesitamos de dirección', 'Decisiones', notas=(
        'Cinco decisiones para arrancar. Las tres primeras son de personas: sin owner ni dedicación el despliegue no '
        'se puede dimensionar ni medir.'))
    items = [
        ('usuario', 'Nombrar el process owner de P2R en MGT',
         'La celda de MGT está sin activar: sin responsable no arranca el despliegue.', 'Antes de oct 27'),
        ('reloj', 'Asignar la dedicación de cada persona a P2R',
         'Falta en MGT y en Ondupack: sin ella no se dimensiona el equipo ni el ahorro de los agentes.', 'Ahora'),
        ('equipo', 'Reforzar la planificación de MGT',
         'Joan Carles es el único planificador y está en las 8 palancas: necesita respaldo o liberar tiempo.',
         'Antes de oct 27'),
        ('diana', 'Decidir el KPI oficial de servicio de Ondupack',
         'Dirección y Comercial: qué OTIF usar y contra qué fecha (la del CTP).', 'Jun – jul 27'),
        ('euro', 'Validar la siguiente fase: cuantificar en euros',
         'Valorar cada hallazgo y el compromiso de valor de cada palanca para priorizar con números.', 'Siguiente paso'),
    ]
    rotulo(sl, 10.6, 1.4, 2.1, 'Cuándo', color=MUTED)
    for i, (ic, tit, txt, cuando) in enumerate(items):
        y = 1.7 + i * 0.92
        caja(sl, L, y, W, 0.8, relleno=CARD)
        caja(sl, L + 0.18, y + 0.17, 0.46, 0.46, relleno=NAVY, forma=MSO_SHAPE.OVAL)
        texto(sl, L + 0.18, y + 0.17, 0.46, 0.46, str(i + 1), tam=13, negrita=True, color=AZUL, alin=PP_ALIGN.CENTER,
              ancla=MSO_ANCHOR.MIDDLE)
        texto(sl, L + 0.85, y + 0.1, 8.5, 0.6, [[(tit, {'negrita': True, 'tam': 12.5})], [(txt, {'color': TEAL})]],
              tam=10.5, interlineado=1.0, espacio=2, ancla=MSO_ANCHOR.MIDDLE)
        sl.shapes.add_picture(str(ICONOS / f'{ic}_n.png'), Inches(10.05), Inches(y + 0.26), Inches(0.28), Inches(0.28))
        pastilla(sl, 10.6, y + 0.24, 1.95, 0.32, cuando, relleno=BLANCO, color=NAVY, tam=8.5, borde=LINEA)


def main():
    d = Deck()
    d.portada('Plan to Release', 'MGT y Ondupack Almendralejo · hoy, to-be y equipo',
              'Comité de dirección  |  Noviembre 2026')
    resumen(d)

    d.separador('01', 'MGT', 'Papelera · suministra papel a Ondupack y a clientes externos')
    asis(d, 'MGT', 'MGT · Situación actual')
    tobe(d, 'MGT', 'MGT · To-be', 0.46)
    pal_mgt = [('P2R-01', 'Demanda'), ('P2R-02', 'Capacidad'), ('P2R-03', 'Inventario'), ('P2R-04', 'Fecha'),
               ('P2R-05', 'Materiales'), ('P2R-06', 'Programación'), ('P2R-07', 'Liberación'), ('P2R-08', 'Servicio')]
    equipo(d, 'MGT', 'MGT · Equipo disponible', [
        ('JC', 'Joan Carles Girones', 'Planificador de la producción; atiende también al cliente',
         {'P2R-02', 'P2R-03', 'P2R-06', 'P2R-07'}, {'P2R-01', 'P2R-04', 'P2R-05', 'P2R-08'}),
        ('AF', 'Albert Fontanet', 'Responsable comercial', {'P2R-01', 'P2R-04', 'P2R-08'}, {'P2R-02', 'P2R-03'}),
        ('MB', 'Marta Boyo', 'Compras · Category Manager', {'P2R-05'}, {'P2R-07'}),
    ], pal_mgt,
        'Ramón, en el algoritmo de combinación · Sistemas / IT en MRP, liberación y visibilidad de stock · Angélica '
        'Lemus (Compras de Ondupack) en lo intercompany.',
        'Joan Carles es el único planificador y está en las 8 palancas. Falta nombrar el process owner de la celda y '
        'cargar la dedicación de cada persona.')

    d.separador('02', 'Ondupack Almendralejo', 'Cartonera · onduladora y converting', tam_titulo=28)
    asis(d, 'Ondupack', 'Ondupack · Situación actual')
    tobe(d, 'Ondupack Almendralejo', 'Ondupack · To-be', 0.5, empezar=[
        ('Capacidad real de la 924', 'Velocidad real y algoritmo de combinación en 2027: son la base de la carga y la '
                                     'secuencia.'),
        ('Bloque de materiales', 'Desde sep 27: ataca la falta de papel, su problema nº 1.'),
        ('Agentes desde abril de 2028', 'Cuando Ondupack ya trabaje en GIO web.'),
    ])
    pal_op = [('P2R-01', 'Demanda'), ('P2R-02', 'Capacidad'), ('P2R-05', 'Materiales'), ('P2R-06', 'Programación'),
              ('P2R-08', 'Servicio')]
    equipo(d, 'Ondupack', 'Ondupack · Equipo disponible', [
        ('MA', 'Miguel Ángel Rosado', 'Planificación y Expediciones', {'P2R-01', 'P2R-06'},
         {'P2R-02', 'P2R-05', 'P2R-08'}),
        ('TB', 'Tomás Buenavida', 'Planificación', set(), {'P2R-01', 'P2R-02', 'P2R-05', 'P2R-06'}),
        ('JB', 'José María Baquero', 'Responsable de Producción', {'P2R-02'}, {'P2R-06'}),
        ('AL', 'Angélica Lemus', 'Compras · Category Manager', {'P2R-05'}, set()),
        ('TE', 'Mª Teresa Espinosa', 'Comercial', {'P2R-08'}, {'P2R-01'}),
    ], pal_op,
        'Karim Karach (Producción, Navalmoral) en capacidad · Sistemas / IT en el algoritmo de combinación. '
        'Angélica apoya además a MGT en lo intercompany.',
        'Planificación son 2 personas y Miguel Ángel está en las 5 palancas. Solo Mª Teresa tiene dedicación cargada '
        '(10 %). Process owner: Enrique Tejada.')

    d.separador('03', 'Siguientes pasos', 'Intercompany, hoja de ruta y decisiones')
    intercompany(d)
    hoja_de_ruta(d)
    decisiones(d)
    d.guardar()
    print(SALIDA)


if __name__ == '__main__':
    main()
