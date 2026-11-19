# AGSA — Planificación

Repositorio de **AGSA** (tubos) en la organización de GitHub
**PLANIFICACION-CL**. Aquí va SOLO lo relacionado con **AGSA**.

## Organización de repositorios (un repo por empresa)
| Repo | Empresa | Actividad | ERP · Planificación |
|---|---|---|---|
| AGSA | AGSA | Tubos | SAP · GPB |
| PERSEIDA | Perseida | Cosmética / higiene | SAP · GPP |
| SIKNO | Sikno |  |  |
| MGT-PDA | MGT / PDA | Papel |  |
| ONDUPACK | Ondupack | Cartón | GIO |
| MARCELIANO-MARTIN | Marceliano Martín |  |  |
| ONDUPET | Ondupet | Packaging (termoformado / extrusión) | GIO |
| BALBOA | Balboa | Acería | GESAG · GPB |
| OMEGA | Omega |  | SAP |
| FERROMALLAS | Ferromallas | Mallas | GESAG · GPx (previsto) |
| LASAOMALLAS | Lasaomallas | Mallas | GESAG · GPx (previsto) |
| DEPARTAMENTO | Departamento de Planificación (CL Grupo Industrial) | Todo lo que afecta a varias empresas: modelo corporativo, P2R, presentaciones, plantillas comunes |  |

Si una tarea es de otra empresa, avisar: debe hacerse en su repositorio, no aquí.

## Regla: nuevas secciones
Antes de crear una sección o herramienta nueva (una carpeta nueva de primer nivel),
**preguntar al usuario** si prefiere:
1. Crearla como carpeta dentro de este repo, o
2. Crear un repositorio nuevo para tenerlo mejor organizado.
No crear la carpeta hasta tener respuesta.

## Estructura de este repo
- `carga-capacidad/` → Carga de Capacidad: Carga y saturación de máquinas de tubo
- `abc-estrategias/` → ABC y Estrategias: ABC y estrategias MTS / ATO / MTO
- `analisis-ciclo-21d/` → Análisis Ciclo 21 días: Escenarios, mínimos de fabricación, temporales y coste de transformación
- `docs/` → Documentación: Definiciones de planificación

Cada herramienta va en su propia carpeta. `README.md` es el índice para personas.

## Reglas de `abc-estrategias/ABC_Estrategias_AGSA.html`

1. **El maestro de AGSA va incrustado en el archivo y no se pierde.**
   Vive en `<script id="agsa-seed" type="application/json">`: maestro de materiales,
   familias/redondos y reglas de medida (formato compacto `v:2`). Se carga solo al abrir
   el archivo por primera vez y el botón «Restaurar maestro de AGSA» de Inicio lo repone.
   **Cualquier HTML nuevo de este proyecto debe conservar ese bloque tal cual**; para
   actualizarlo, exportar un backup desde la herramienta (usa el mismo formato) y sustituir
   el contenido del script, quitando los inputs.
   Los **pedidos, previsiones y fabricaciones no se incrustan**: se cargan desde Excel en
   cada sesión y se conservan en el navegador mientras quepan.

2. **Las cantidades de los pedidos están en toneladas** y se muestran con dos decimales.
   Los precios son €/t y los pesos calculados (Suma de Peso, Tn Mes, Tn Semana, percentiles
   de MMPP) van también en toneladas.

3. **El periodo de cálculo es común**: se elige en Parametrización ABC, Calculos MMPP Tubos
   y Calculos Tubos (fecha de pedido o de entrega y meses desde/hasta) y se guarda en
   `PARAM().periodo`.

4. **Nomenclatura del maestro**: el ABC del maestro es el código de tres posiciones (`Aa1`)
   y respeta mayúsculas y minúsculas; las estrategias del maestro son `TMTS` / `TATO` /
   `TMTO` y se comparan sin la `T` inicial.

5. El `localStorage` guarda el estado con la misma serialización compacta; si no cabe,
   se guarda todo menos los pedidos (que vuelven del bloque incrustado) y se avisa.

6. **Temporales**: cada medida tiene, **por máquina**, un mínimo de producción (30 por
   defecto, 100 en TUBOS_17), un temporal y un periodo de repetición (1 cada ciclo,
   2 cada dos, 3 cada tres, 0 nunca), guardados en `regla.cfg[maquina]`. Los colores del
   subciclo salen de ahí: verde 1, amarillo 2, naranja 3, rojo 0.
   En «Calculo Temporales», **Subciclos variable 30d** y **Revisar Temporal** usan fórmulas
   **provisionales** (total 30d por debajo del mínimo, y desviación de la fabricación MTS
   frente al temporal superior al 10 %) a la espera de las definitivas del cliente.

## Comprobación
No hay build ni tests: se valida abriendo el HTML en un navegador (Playwright) y comprobando que no haya errores de consola, que los cálculos cuadren y que las pantallas rendericen.

## Convenciones
- Idioma: español en interfaz, comentarios y mensajes de commit.
- Herramientas en HTML autocontenido (un archivo por herramienta) salvo que se pida otra cosa.
- No subir contraseñas ni claves privadas al repo.
- El trabajo terminado se fusiona en `main`: las sesiones nuevas solo ven `main`.

## Trabajo en curso
@PROGRESO.md

Al terminar cada tarea: actualizar `PROGRESO.md` y hacer commit.
