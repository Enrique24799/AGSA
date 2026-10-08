#!/usr/bin/env python3
"""Genera la herramienta HTML autocontenida a partir de plantilla.html, datos_base.json y el logo."""
import base64
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'Modelo_Planificacion_CL.html')


def main():
    # los datos del equipo (último export) tienen prioridad sobre la propuesta inicial
    fuente = 'datos_actuales.json' if os.path.exists(os.path.join(HERE, 'datos_actuales.json')) else 'datos_base.json'
    with open(os.path.join(HERE, fuente), encoding='utf-8') as f:
        datos = json.load(f)
    datos.setdefault('raciAsis', {})
    datos.setdefault('organigramas', {})
    datos.setdefault('apartados', {})
    datos.setdefault('flujos', {})
    with open(os.path.join(HERE, 'plantilla.html'), encoding='utf-8') as f:
        html = f.read()
    with open(os.path.join(HERE, 'logo_blanco.png'), 'rb') as f:
        logo = base64.b64encode(f.read()).decode('ascii')
    # aportaciones posteriores: se aplican una vez en cada navegador y solo rellenan lo vacío
    pdir = os.path.join(HERE, 'parches')
    parches = []
    for fn in sorted(os.listdir(pdir)) if os.path.isdir(pdir) else []:
        if fn.endswith('.json'):
            with open(os.path.join(pdir, fn), encoding='utf-8') as f:
                parches.append(json.load(f))
    # el JSON va dentro de <script>: se escapa "</" para que no pueda cerrar la etiqueta
    payload = json.dumps(datos, ensure_ascii=False).replace('</', '<\\/')
    payload_p = json.dumps(parches, ensure_ascii=False).replace('</', '<\\/')
    assert html.count('__DATOS__') == 1 and html.count('__LOGO__') == 1 and html.count('__PARCHES__') == 1
    html = html.replace('__DATOS__', payload).replace('__PARCHES__', payload_p).replace('__LOGO__', logo)
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(html)
    print('OK', os.path.normpath(OUT), f'{len(html) / 1024:.0f} KB', 'desde', fuente, f'+ {len(parches)} parches')


if __name__ == '__main__':
    main()
