# Modelo de planificación corporativo · CL Grupo Industrial

Kit para centralizar la planificación de packaging en corporativo: **corporativo define** procesos, roles, personas, políticas y KPIs; **cada planta ejecuta**.

| Archivo | Para qué sirve |
|---|---|
| `Modelo_Planificacion_CL.html` | Herramienta de trabajo. Se abre con doble clic (Chrome o Edge), no necesita instalación ni conexión. Guarda lo escrito en el propio navegador y exporta un JSON con todo. |
| `Kit_Modelo_Planificacion_CL.xlsx` | La propuesta inicial en Excel: procesos, RACI, roles, asignación, organigramas, empresas, políticas, KPIs, reuniones, decisiones y plan. |
| `fuentes/` | Origen de ambos: `datos_actuales.json` (último export del equipo, tiene prioridad), `datos_base.py` (propuesta inicial), `plantilla.html`, `build_html.py` y `generar_kit.py`. |

## Flujo de trabajo

1. Cada persona abre la herramienta, pone su nombre y rellena las notas de su empresa y proceso, y el organigrama de su empresa (pestañas **as-is** y **to-be**: cajitas por puesto con persona, área, rol, dependencia jerárquica y funcional y % de dedicación a planificación).
2. En **Matriz RACI**, la pestaña **as-is** recoge por empresa quién hace hoy cada actividad y marca en naranja lo que no coincide con el **to-be** (estándar corporativo, único para todas).
3. **Exportar para Claude** descarga un JSON. Si varias personas trabajan por separado, se pueden juntar con **Importar → Combinar** o pasar todos los JSON.
4. Con el JSON se genera el modelo final: Excel con notas y madurez por empresa y proceso, RACI con nombres, fichas de rol, manual de planificación y diapositivas.

## Regenerar

```bash
python3 fuentes/datos_base.py                         # propuesta → fuentes/datos_base.json
python3 fuentes/build_html.py                         # herramienta HTML (con datos_actuales.json si existe)
python3 fuentes/generar_kit.py [export.json] [salida.xlsx]   # Excel (propuesta o un export)
```
