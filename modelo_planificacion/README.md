# Modelo de planificación corporativo · CL Grupo Industrial

Kit para centralizar la planificación de packaging en corporativo: **corporativo define** procesos, roles, personas, políticas y KPIs; **cada planta ejecuta**.

| Archivo | Para qué sirve |
|---|---|
| `Modelo_Planificacion_CL.html` | Herramienta de trabajo. Se abre con doble clic (Chrome o Edge), no necesita instalación ni conexión. Guarda lo escrito en el propio navegador y exporta un JSON con todo. |
| `Kit_Modelo_Planificacion_CL.xlsx` | La propuesta inicial en Excel: procesos, flujos (si hay), RACI, roles, asignación, organigramas, empresas, políticas, KPIs, reuniones, decisiones y plan. |
| `fuentes/` | Origen de ambos: `datos_actuales.json` (último export del equipo, tiene prioridad), `datos_base.py` (propuesta inicial), `parches/` (aportaciones posteriores, como el análisis E2E de Ondupet), `plantilla.html`, `build_html.py` y `generar_kit.py`. |

## Flujo de trabajo

1. Cada persona abre la herramienta, pone su nombre y rellena las notas de su empresa y proceso (si su planta tiene un paso que no está en el catálogo, lo añade con **+ Añadir apartado en …**: misma ficha, solo visible en esa planta), y el organigrama de su empresa (pestañas **as-is** y **to-be**: cajitas por puesto con persona, área, rol, dependencia jerárquica y funcional y % de dedicación a planificación).
2. En **Flujos de proceso**, cada empresa dibuja el flujo de cada proceso en dos pestañas independientes, **as-is** y **to-be**: cajas de inicio/fin, tarea, decisión, documento/sistema y nota, unidas con flechas (se arrastra desde el punto del borde de una caja hasta otra), con responsable, sistema, actividad del catálogo y puntos de dolor. «Partir de las actividades» crea el esqueleto del proceso, «Copiar de…» trae el flujo de otra pestaña o empresa, y cada flujo se descarga en PNG o SVG para presentaciones.
3. En **Matriz RACI**, la pestaña **as-is** recoge por empresa quién hace hoy cada actividad y marca en naranja lo que no coincide con el **to-be** (estándar corporativo, único para todas). Cada proceso se despliega o contrae pulsando su cabecera.
4. **Exportar para Claude** descarga un JSON. Si varias personas trabajan por separado, se pueden juntar con **Importar → Combinar** o pasar todos los JSON.
5. Con el JSON se genera el modelo final: Excel con notas y madurez por empresa y proceso, flujos paso a paso, RACI con nombres, fichas de rol, manual de planificación y diapositivas.

## Regenerar

```bash
python3 fuentes/datos_base.py                         # propuesta → fuentes/datos_base.json
python3 fuentes/build_html.py                         # herramienta HTML (con datos_actuales.json si existe)
python3 fuentes/generar_kit.py [export.json] [salida.xlsx]   # Excel (propuesta o un export)
```

Cada archivo de `fuentes/parches/` va dentro del HTML y se incorpora **una sola vez en cada navegador**. Solo rellena lo que esté vacío: notas campo a campo, RACI as-is fila a fila, y flujos y apartados que no existan. Así, quien ya tenga datos guardados recibe la aportación sin perder nada de lo que escribió.
