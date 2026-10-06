#!/usr/bin/env python3
"""Genera la herramienta HTML autocontenida a partir de plantilla.html, datos_base.json y el logo."""
import base64
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'Modelo_Planificacion_CL.html')


def main():
    with open(os.path.join(HERE, 'datos_base.json'), encoding='utf-8') as f:
        datos = json.load(f)
    with open(os.path.join(HERE, 'plantilla.html'), encoding='utf-8') as f:
        html = f.read()
    with open(os.path.join(HERE, 'logo_blanco.png'), 'rb') as f:
        logo = base64.b64encode(f.read()).decode('ascii')
    # el JSON va dentro de <script>: se escapa "</" para que no pueda cerrar la etiqueta
    payload = json.dumps(datos, ensure_ascii=False).replace('</', '<\\/')
    assert html.count('__DATOS__') == 1 and html.count('__LOGO__') == 1
    html = html.replace('__DATOS__', payload).replace('__LOGO__', logo)
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(html)
    print('OK', os.path.normpath(OUT), f'{len(html) / 1024:.0f} KB')


if __name__ == '__main__':
    main()
