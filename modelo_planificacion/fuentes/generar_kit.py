#!/usr/bin/env python3
"""Genera el Excel del modelo de planificación a partir de un JSON (la propuesta base o un export de la herramienta).

Uso: python generar_kit.py [datos.json] [salida.xlsx]
"""
import datetime as dt
import json
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
NAVY, BLUE, TEAL, TINT, GREY, BORDER = '00263A', '98C9EB', '244C5A', 'EAF3FA', '5B6B78', 'D7DFE5'
F = 'Arial'

ESTADOS = {'': 'Sin revisar', 'alineado': 'Alineado con el estándar', 'propio': 'Estándar propio',
           'informal': 'Informal', 'nohace': 'No se hace', 'na': 'No aplica'}
AMBITOS = ['Corporativo', 'Planta', 'Otras áreas']

thin = Side(style='thin', color=BORDER)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)


def fill(c):
    return PatternFill('solid', fgColor=c)


def title(ws, text, sub=None):
    ws['A1'] = text
    ws['A1'].font = Font(name=F, size=14, bold=True, color=NAVY)
    if sub:
        ws['A2'] = sub
        ws['A2'].font = Font(name=F, size=9, italic=True, color=GREY)


def table(ws, row, headers, rows, widths, wrap_cols=None):
    """Escribe una tabla con cabecera azul marino desde la fila `row`. Devuelve la última fila escrita."""
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=j, value=h)
        c.font = Font(name=F, size=10, bold=True, color='FFFFFF')
        c.fill = fill(NAVY)
        c.alignment = CENTER
        c.border = BOX
    for i, r in enumerate(rows, row + 1):
        for j, v in enumerate(r, 1):
            c = ws.cell(row=i, column=j, value=v)
            c.font = Font(name=F, size=10, color=NAVY)
            c.alignment = WRAP
            c.border = BOX
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = ws.cell(row=row + 1, column=1)
    return row + len(rows)


def build(data, out, origen):
    roles = data['roles']
    rid = {r['id']: r for r in roles}
    emp = {e['id']: e for e in data['empresas']}
    acts = [(p, a) for p in data['procesos'] for a in p['actividades']]
    raci = data.get('raci', {})
    wb = Workbook()

    # ---------------------------------------------------------------- portada
    ws = wb.active
    ws.title = 'Portada'
    ws.sheet_view.showGridLines = False
    ws['B2'] = 'Modelo de planificación corporativo · CL Grupo Industrial'
    ws['B2'].font = Font(name=F, size=18, bold=True, color=NAVY)
    ws['B3'] = 'Corporativo define procesos, roles, personas, políticas y KPIs; cada planta ejecuta.'
    ws['B3'].font = Font(name=F, size=11, italic=True, color=TEAL)
    ws['B5'], ws['C5'] = 'Origen de los datos', origen
    ws['B6'], ws['C6'] = 'Generado', dt.date.today().strftime('%d/%m/%Y')
    ws['B7'], ws['C7'] = 'Dónde se edita', 'En la herramienta Modelo_Planificacion_CL.html; este Excel es una foto para revisar y compartir.'
    for r in (5, 6, 7):
        ws[f'B{r}'].font = Font(name=F, size=10, bold=True, color=TEAL)
        ws[f'C{r}'].font = Font(name=F, size=10, color=NAVY)
    ws['B9'] = 'Contenido'
    ws['B9'].font = Font(name=F, size=12, bold=True, color=NAVY)
    hojas = [
        ('Procesos', 'Catálogo de procesos y actividades de planificación', '=COUNTA(Procesos!B:B)-1', 'actividades'),
        ('RACI', 'Matriz RACI por rol, con revisión automática de cada fila', "=COUNTIF(RACI!{col}:{col},\"OK\")", 'filas correctas'),
        ('Roles', 'Fichas de rol: misión, responsabilidades, qué decide y KPIs', '=COUNTA(Roles!A:A)-1', 'roles'),
        ('Asignación', 'Titular y suplente de cada rol en cada empresa', "=COUNTA(Asignación!C:C)-1", 'roles con titular'),
        ('Empresas', 'Perímetro: empresas y plantas', '=COUNTA(Empresas!A:A)-1', 'empresas'),
        ('Políticas', 'Políticas de planificación: estándar corporativo y parámetros por planta', "=COUNTA('Políticas'!A:A)-1", 'políticas'),
        ('KPIs', 'Cuadro de mando común', '=COUNTA(KPIs!A:A)-1', 'KPIs'),
        ('Reuniones', 'Cadencia de planificación del grupo', '=COUNTA(Reuniones!A:A)-1', 'reuniones'),
        ('Decisiones', 'Decisiones que hay que cerrar con dirección', '=COUNTIF(Decisiones!F:F,"Abierta")', 'abiertas'),
        ('Plan', 'Plan de despliegue por fases', '=COUNTA(Plan!A:A)-1', 'fases'),
    ]
    has_notes = any(v for e in data.get('notas', {}).values() for v in e.values())
    if has_notes:
        hojas.insert(1, ('Notas', 'Notas por empresa y actividad: situación actual, quién, problemas y propuesta',
                         '=COUNTA(Notas!E:E)-1-COUNTIF(Notas!E:E,"Sin revisar")', 'actividades revisadas'))
        hojas.insert(2, ('Madurez', 'Situación por empresa y proceso', None, ''))
    org = data.get('organigramas', {})
    has_org = any(v for e in org.values() for v in e.values())
    if has_org:
        pos = [h[0] for h in hojas].index('Asignación') + 1
        hojas.insert(pos, ('Organigramas', 'Organigramas as-is y to-be de cada empresa: puesto, persona y dependencias',
                           '=COUNTA(Organigramas!D:D)-1', 'puestos'))
    if any(any(any(x.values()) for x in v.values()) for v in (data.get('raciAsis') or {}).values()):
        hojas.insert([h[0] for h in hojas].index('RACI') + 1,
                     ('RACI as-is', 'Cómo se hace hoy en cada empresa y diferencias con el estándar', "=COUNTA('RACI as-is'!C:C)-1", 'filas'))
    raci_col = get_column_letter(3 + len(roles) + 1)
    for i, (h, d, f, u) in enumerate(hojas, 10):
        ws[f'B{i}'] = h
        ws[f'B{i}'].hyperlink = f"#'{h}'!A1"
        ws[f'B{i}'].font = Font(name=F, size=10, bold=True, color=TEAL, underline='single')
        ws[f'C{i}'] = d
        ws[f'C{i}'].font = Font(name=F, size=10, color=NAVY)
        if f:
            ws[f'D{i}'] = f.replace('{col}', raci_col)
            ws[f'D{i}'].font = Font(name=F, size=10, bold=True, color=NAVY)
            ws[f'D{i}'].alignment = Alignment(horizontal='right')
            ws[f'E{i}'] = u
            ws[f'E{i}'].font = Font(name=F, size=9, color=GREY)
    r0 = 10 + len(hojas) + 1
    ws[f'B{r0}'] = 'Leyenda RACI'
    ws[f'B{r0}'].font = Font(name=F, size=12, bold=True, color=NAVY)
    for k, (l, t, bg, fg) in enumerate([
            ('A', 'Decide y responde del resultado (una sola por actividad)', NAVY, 'FFFFFF'),
            ('R', 'Lo ejecuta', BLUE, NAVY), ('A/R', 'Decide y además lo ejecuta', NAVY, 'FFFFFF'),
            ('C', 'Se le consulta antes', TINT, TEAL), ('I', 'Se le informa después', 'FFFFFF', GREY)], r0 + 1):
        c = ws[f'B{k}']
        c.value, c.fill, c.font, c.alignment, c.border = l, fill(bg), Font(name=F, size=10, bold=True, color=fg), CENTER, BOX
        ws[f'C{k}'] = t
        ws[f'C{k}'].font = Font(name=F, size=10, color=NAVY)
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 78
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 22

    # ---------------------------------------------------------------- procesos
    ws = wb.create_sheet('Procesos')
    rows = []
    for p, a in acts:
        row = raci.get(a['id'], {})
        ra = [rid[k]['nombre'] for k, v in row.items() if v in ('A', 'A/R') and k in rid]
        rr = [rid[k]['nombre'] for k, v in row.items() if v in ('R', 'A/R') and k in rid]
        rows.append([p['nombre'], a['id'], a['nombre'], a.get('descripcion', ''), a.get('frecuencia', ''),
                     ', '.join(ra), ', '.join(rr)])
    table(ws, 1, ['Proceso', 'ID', 'Actividad', 'Descripción', 'Frecuencia estándar', 'Decide (A)', 'Ejecuta (R)'],
          rows, [18, 7, 38, 60, 14, 30, 40])

    # ---------------------------------------------------------------- RACI
    ws = wb.create_sheet('RACI')
    ordered = [r for g in AMBITOS for r in roles if (r.get('ambito') or 'Otras áreas') == g]
    ncol = 3 + len(ordered)
    # fila 1: grupos; fila 2: roles
    ws.cell(row=1, column=1, value='Matriz RACI').font = Font(name=F, size=14, bold=True, color=NAVY)
    col = 4
    for g in AMBITOS:
        n = sum(1 for r in ordered if (r.get('ambito') or 'Otras áreas') == g)
        if not n:
            continue
        ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + n - 1)
        c = ws.cell(row=2, column=col, value=g.upper())
        c.font, c.fill, c.alignment = Font(name=F, size=9, bold=True, color=TEAL), fill(TINT), CENTER
        col += n
    hdr = ['Proceso', 'ID', 'Actividad'] + [r.get('corto') or r['nombre'] for r in ordered] + ['Revisión']
    for j, h in enumerate(hdr, 1):
        c = ws.cell(row=3, column=j, value=h)
        c.font, c.fill, c.alignment, c.border = Font(name=F, size=9, bold=True, color='FFFFFF'), fill(NAVY), CENTER, BOX
    ws.row_dimensions[3].height = 42
    styles = {'A': (NAVY, 'FFFFFF'), 'A/R': (NAVY, 'FFFFFF'), 'R': (BLUE, NAVY), 'C': (TINT, TEAL), 'I': ('FFFFFF', GREY)}
    r = 4
    first, last = get_column_letter(4), get_column_letter(ncol)
    for p, a in acts:
        row = raci.get(a['id'], {})
        vals = [p['nombre'], a['id'], a['nombre']] + [row.get(x['id'], '') for x in ordered]
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=j, value=v or None)
            c.border = BOX
            if j <= 3:
                c.font, c.alignment = Font(name=F, size=10, color=NAVY), WRAP
            else:
                bg, fg = styles.get(v, ('FFFFFF', NAVY))
                c.font, c.fill, c.alignment = Font(name=F, size=10, bold=True, color=fg), fill(bg), CENTER
        rng = f'{first}{r}:{last}{r}'
        c = ws.cell(row=r, column=ncol + 1,
                    value=f'=IF(COUNTIF({rng},"A")+COUNTIF({rng},"A/R")<>1,"Revisar A",'
                          f'IF(COUNTIF({rng},"R")+COUNTIF({rng},"A/R")=0,"Falta R","OK"))')
        c.font, c.alignment, c.border = Font(name=F, size=9, bold=True, color=TEAL), CENTER, BOX
        r += 1
    dv = DataValidation(type='list', formula1='"R,A,A/R,C,I"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f'{first}4:{last}{r - 1}')
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 6
    ws.column_dimensions['C'].width = 40
    for j in range(4, ncol + 1):
        ws.column_dimensions[get_column_letter(j)].width = 11
    ws.column_dimensions[get_column_letter(ncol + 1)].width = 11
    ws.freeze_panes = 'D4'
    leg = r + 1
    ws.cell(row=leg, column=1, value='A = decide y responde (una por fila) · R = ejecuta · A/R = decide y ejecuta · '
                                     'C = se le consulta · I = se le informa. La columna «Revisión» se recalcula sola.'
            ).font = Font(name=F, size=9, italic=True, color=GREY)

    # ---------------------------------------------------------------- RACI as-is por empresa (si hay datos)
    asis = {e: v for e, v in (data.get('raciAsis') or {}).items() if any(any(x.values()) for x in v.values())}
    if asis:
        ws = wb.create_sheet('RACI as-is')
        ws.cell(row=1, column=1, value='RACI as-is · cómo se hace hoy en cada empresa').font = Font(name=F, size=14, bold=True, color=NAVY)
        hdr = ['Empresa', 'Proceso', 'ID', 'Actividad'] + [x.get('corto') or x['nombre'] for x in ordered] + ['Revisión', 'Dif. con to-be']
        for j, h in enumerate(hdr, 1):
            c = ws.cell(row=3, column=j, value=h)
            c.font, c.fill, c.alignment, c.border = Font(name=F, size=9, bold=True, color='FFFFFF'), fill(NAVY), CENTER, BOX
        ws.row_dimensions[3].height = 42
        r = 4
        f0, f1 = get_column_letter(5), get_column_letter(4 + len(ordered))
        t0, t1 = get_column_letter(4), get_column_letter(3 + len(ordered))
        for e in data['empresas']:
            src = asis.get(e['id'])
            if not src:
                continue
            for k, (p, a) in enumerate(acts):
                row = src.get(a['id'], {})
                if not any(row.values()):
                    continue
                vals = [e['nombre'], p['nombre'], a['id'], a['nombre']] + [row.get(x['id'], '') for x in ordered]
                for j, v in enumerate(vals, 1):
                    c = ws.cell(row=r, column=j, value=v or None)
                    c.border = BOX
                    if j <= 4:
                        c.font, c.alignment = Font(name=F, size=10, color=NAVY), WRAP
                    else:
                        bg, fg = styles.get(v, ('FFFFFF', NAVY))
                        c.font, c.fill, c.alignment = Font(name=F, size=10, bold=True, color=fg), fill(bg), CENTER
                rng = f'{f0}{r}:{f1}{r}'
                c = ws.cell(row=r, column=len(ordered) + 5,
                            value=f'=IF(COUNTIF({rng},"A")+COUNTIF({rng},"A/R")<>1,"Revisar A",'
                                  f'IF(COUNTIF({rng},"R")+COUNTIF({rng},"A/R")=0,"Falta R","OK"))')
                c.font, c.alignment, c.border = Font(name=F, size=9, bold=True, color=TEAL), CENTER, BOX
                tr = 4 + k  # fila de la misma actividad en la hoja RACI (to-be)
                c = ws.cell(row=r, column=len(ordered) + 6, value=f"=SUMPRODUCT(--({rng}<>RACI!{t0}{tr}:{t1}{tr}))")
                c.font, c.alignment, c.border = Font(name=F, size=9, bold=True, color=NAVY), CENTER, BOX
                r += 1
        for j, w in enumerate([22, 16, 6, 40] + [11] * len(ordered) + [11, 12], 1):
            ws.column_dimensions[get_column_letter(j)].width = w
        ws.freeze_panes = 'E4'
        ws.auto_filter.ref = f'A3:{get_column_letter(len(ordered) + 6)}{r - 1}'
        ws.cell(row=r + 1, column=1, value='«Dif. con to-be» cuenta las celdas de la fila que no coinciden con la hoja RACI (estándar corporativo).'
                ).font = Font(name=F, size=9, italic=True, color=GREY)

    # ---------------------------------------------------------------- roles y asignación
    ws = wb.create_sheet('Roles')
    table(ws, 1, ['Rol', 'Nombre corto', 'Ámbito', 'Misión', 'Responsabilidades', 'Qué decide', 'KPIs'],
          [[x['nombre'], x.get('corto', ''), x.get('ambito', ''), x.get('mision', ''), x.get('responsabilidades', ''),
            x.get('decide', ''), x.get('kpis', '')] for x in ordered], [30, 18, 13, 45, 55, 40, 32])

    ws = wb.create_sheet('Asignación')
    rows = [[emp.get(p['empresa'], {}).get('nombre', p['empresa']), rid.get(p['rol'], {}).get('nombre', p['rol']),
             p.get('titular') or None, p.get('suplente') or None, p.get('puesto', ''), p.get('dedicacion', ''),
             p.get('email', ''), p.get('notas', '')] for p in data.get('personas', [])]
    end = table(ws, 1, ['Empresa', 'Rol', 'Titular', 'Suplente', 'Puesto actual', '% dedicación', 'Email', 'Notas'],
                rows, [26, 40, 28, 28, 24, 12, 28, 36])
    for i in range(2, end + 1):
        for col in ('C', 'D'):
            if ws[f'{col}{i}'].value is None:
                ws[f'{col}{i}'].fill = fill('FFF4CC')
    ws.cell(row=end + 2, column=1, value='En amarillo, los huecos por asignar (titular o suplente).'
            ).font = Font(name=F, size=9, italic=True, color=GREY)

    # ---------------------------------------------------------------- organigramas (árbol con sangría)
    if has_org:
        ws = wb.create_sheet('Organigramas')
        rows = []
        lab = lambda x: (x.get('puesto') or 'Puesto') + (' · ' + x['persona'] if x.get('persona') else '')
        for e in data['empresas']:
            for vista, vname in (('asis', 'As-is'), ('tobe', 'To-be')):
                nodes = (org.get(e['id']) or {}).get(vista) or []
                byid = {n['id']: n for n in nodes}
                kids = {}
                for n in nodes:
                    par = n.get('parent') if n.get('parent') in byid and n.get('parent') != n['id'] else ''
                    kids.setdefault(par, []).append(n)
                seen = set()

                def walk(n, d):
                    if n['id'] in seen:
                        return
                    seen.add(n['id'])
                    rows.append([e['nombre'], vname, d + 1, '      ' * d + (n.get('puesto') or ''),
                                 n.get('persona') or 'Vacante', n.get('area', ''),
                                 rid.get(n.get('rol'), {}).get('nombre', ''),
                                 emp.get(n.get('empresa'), {}).get('nombre', '') if n.get('empresa') else '',
                                 lab(byid[n['parent']]) if n.get('parent') in byid else '',
                                 lab(byid[n['funcional']]) if n.get('funcional') in byid else '',
                                 n.get('dedicacion', ''), n.get('notas', '')])
                    for c in kids.get(n['id'], []):
                        walk(c, d + 1)

                for r in kids.get('', []) + nodes:
                    walk(r, 0)
        end = table(ws, 1, ['Empresa', 'Vista', 'Nivel', 'Puesto', 'Persona', 'Área', 'Rol del modelo', 'De otra empresa',
                            'Depende de', 'Dependencia funcional de', '% planificación', 'Notas'],
                    rows, [24, 8, 7, 44, 24, 14, 34, 22, 38, 38, 13, 44])
        ws.auto_filter.ref = f'A1:L{end}'
        for i in range(2, end + 1):
            if ws[f'E{i}'].value == 'Vacante':
                ws[f'E{i}'].font = Font(name=F, size=10, italic=True, color='B4532A')
            if ws[f'F{i}'].value == 'Planificación':
                for col in 'ABCDEFGHIJKL':
                    ws[f'{col}{i}'].fill = fill(TINT)

    if has_notes:
        ws = wb.create_sheet('Notas', 1)
        rows = []
        for e in data['empresas']:
            for p, a in acts:
                n = data['notas'].get(e['id'], {}).get(a['id'], {})
                rows.append([e['nombre'], p['nombre'], a['id'], a['nombre'], ESTADOS.get(n.get('estado', ''), 'Sin revisar'),
                             n.get('quien', ''), n.get('herramienta', ''), n.get('frecuencia', ''), n.get('horas', ''),
                             n.get('situacion', ''), n.get('problemas', ''), n.get('propuesta', ''),
                             n.get('prioridad', ''), n.get('notas', '')])
        table(ws, 1, ['Empresa', 'Proceso', 'ID', 'Actividad', 'Situación', 'Quién lo hace hoy', 'Herramienta',
                      'Frecuencia real', 'Horas/semana', 'Cómo se hace hoy', 'Problemas', 'Propuesta', 'Prioridad',
                      'Notas'], rows, [22, 16, 6, 34, 20, 22, 18, 14, 10, 45, 45, 45, 10, 30])
        ws.auto_filter.ref = f'A1:N{len(rows) + 1}'
        ws = wb.create_sheet('Madurez', 2)
        title(ws, 'Actividades alineadas con el estándar / actividades aplicables',
              'Se calcula desde la hoja Notas.')
        procs = [p['nombre'] for p in data['procesos']]
        for j, h in enumerate(['Empresa'] + procs + ['Total'], 1):
            c = ws.cell(row=4, column=j, value=h)
            c.font, c.fill, c.alignment, c.border = Font(name=F, size=9, bold=True, color='FFFFFF'), fill(NAVY), CENTER, BOX
        for i, e in enumerate(data['empresas'], 5):
            ws.cell(row=i, column=1, value=e['nombre']).font = Font(name=F, size=10, bold=True, color=NAVY)
            for j, pn in enumerate(procs + [None], 2):
                cond_p = f',Notas!$B:$B,"{pn}"' if pn else ''
                aplic = f'(COUNTIFS(Notas!$A:$A,$A{i}{cond_p})-COUNTIFS(Notas!$A:$A,$A{i}{cond_p},Notas!$E:$E,"No aplica"))'
                c = ws.cell(row=i, column=j, value=f'=IFERROR(COUNTIFS(Notas!$A:$A,$A{i}{cond_p},Notas!$E:$E,'
                                                    f'"Alineado con el estándar")/{aplic},0)')
                c.number_format, c.alignment, c.border = '0%', CENTER, BOX
                c.font = Font(name=F, size=10, color=NAVY)
        ws.column_dimensions['A'].width = 28
        for j in range(2, len(procs) + 3):
            ws.column_dimensions[get_column_letter(j)].width = 13

    # ---------------------------------------------------------------- resto de tablas
    rn = lambda k: rid.get(k, {}).get('nombre', k or '')
    specs = [
        ('Empresas', ['Empresa / planta', 'Tipo', 'ERP', 'Herramienta de planificación', 'Director/a de planta',
                      'Responsable de planificación hoy', 'Nº personas planificando', 'Particularidades', 'Notas'],
         [[e.get(k, '') for k in ('nombre', 'tipo', 'erp', 'herramienta', 'director', 'responsable', 'personasHoy',
                                   'particularidades', 'notas')] for e in data['empresas']],
         [26, 30, 22, 32, 22, 26, 12, 50, 30]),
        ('Políticas', ['Política', 'Estándar corporativo', 'Parámetros por planta', 'Responsable', 'Estado', 'Notas'],
         [[x['nombre'], x['estandar'], x['parametros'], rn(x['responsable']), x['estado'], x['notas']]
          for x in data['politicas']], [28, 60, 32, 32, 13, 30]),
        ('KPIs', ['KPI', 'Definición', 'Fórmula', 'Frecuencia', 'Responsable', 'Objetivo', 'Notas'],
         [[x['nombre'], x['definicion'], x['formula'], x['frecuencia'], rn(x['responsable']), x['objetivo'], x['notas']]
          for x in data['kpis']], [26, 42, 48, 12, 32, 24, 26]),
        ('Reuniones', ['Reunión', 'Frecuencia', 'Duración', 'Participantes', 'Entradas', 'Salidas / decisiones', 'Notas'],
         [[x['nombre'], x['frecuencia'], x['duracion'], x['participantes'], x['entradas'], x['salidas'], x['notas']]
          for x in data['cadencia']], [30, 18, 12, 46, 38, 36, 24]),
        ('Decisiones', ['Decisión', 'Opciones', 'Recomendación', 'Quién decide', 'Fecha límite', 'Estado', 'Resolución'],
         [[x['nombre'], x['opciones'], x['recomendacion'], x['responsable'], x['fecha'], x['estado'], x['resolucion']]
          for x in data['decisiones']], [34, 48, 44, 26, 12, 11, 34]),
        ('Plan', ['Fase', 'Actividades', 'Entregables', 'Plazo', 'Responsable', 'Estado', 'Notas'],
         [[x['nombre'], x['actividades'], x['entregables'], x['plazo'], rn(x['responsable']), x['estado'], x['notas']]
          for x in data['plan']], [20, 60, 40, 12, 30, 12, 26]),
    ]
    for name, headers, rows, widths in specs:
        table(wb.create_sheet(name), 1, headers, rows, widths)

    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = NAVY if ws.title in ('Portada', 'RACI') else BLUE
        # impresión: horizontal y ajustada al ancho de la página
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    wb.save(out)
    return out


if __name__ == '__main__':
    actuales = os.path.join(HERE, 'datos_actuales.json')
    src = sys.argv[1] if len(sys.argv) > 1 else (actuales if os.path.exists(actuales) else os.path.join(HERE, 'datos_base.json'))
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, '..', 'Kit_Modelo_Planificacion_CL.xlsx')
    with open(src, encoding='utf-8') as f:
        data = json.load(f)
    meta = data.get('meta', {})
    origen = ('Propuesta inicial' if os.path.basename(src) == 'datos_base.json'
              else f"Datos del equipo · {meta.get('autor') or 'export de la herramienta'} · {meta.get('exportado', '')[:10]}")
    print('OK', os.path.normpath(build(data, out, origen)))
