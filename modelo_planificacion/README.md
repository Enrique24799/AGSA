# Modelo de planificación corporativo · CL Grupo Industrial

Kit para centralizar la planificación de packaging en corporativo: **corporativo define** procesos, roles, personas, políticas y KPIs; **cada planta ejecuta**.

| Archivo | Para qué sirve |
|---|---|
| `Modelo_Planificacion_CL.html` | Herramienta de trabajo. Se abre con doble clic (Chrome o Edge), no necesita instalación ni conexión. Guarda lo escrito en el propio navegador y exporta un JSON con todo. |
| `Kit_Modelo_Planificacion_CL.xlsx` | La propuesta inicial en Excel: procesos, RACI, roles, asignación, empresas, políticas, KPIs, reuniones, decisiones y plan. |
| `fuentes/` | Origen de ambos: `datos_base.py` (propuesta), `plantilla.html`, `build_html.py` y `generar_kit.py`. |

## Flujo de trabajo

1. Cada persona abre la herramienta, pone su nombre y rellena las notas de su empresa y proceso.
2. **Exportar para Claude** descarga un JSON. Si varias personas trabajan por separado, se pueden juntar con **Importar → Combinar** o pasar todos los JSON.
3. Con el JSON se genera el modelo final: Excel con notas y madurez por empresa y proceso, RACI con nombres, fichas de rol, manual de planificación y diapositivas.

## Regenerar

```bash
python3 fuentes/datos_base.py                         # propuesta → fuentes/datos_base.json
python3 fuentes/build_html.py                         # herramienta HTML
python3 fuentes/generar_kit.py [export.json] [salida.xlsx]   # Excel (propuesta o un export)
```
