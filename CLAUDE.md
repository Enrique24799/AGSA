# AGSA — Planificación

Repositorio de la empresa **AGSA** (fabricación de tubos) dentro de la organización
de GitHub **PLANIFICACION-CL**. Aquí va SOLO lo relacionado con AGSA.

## Organización de repositorios (un repo por empresa)
| Repo | Empresa |
|---|---|
| AGSA | AGSA (este repo) |
| PERSEIDA | Perseida |
| SIKNO | Sikno |
| MGT-PDA | MGT/PDA |
| MARCELIANO-MARTIN | Marceliano Martín |
| ONDUPET | Ondupet |
| BALBOA | Balboa |
| OMEGA | Omega |
| DEPARTAMENTO | Departamento |
| FERROMALLAS | Ferromallas |
| LASAOMALLAS | Lasaomallas |
| comun | Plantillas y piezas reutilizables entre empresas |

Si una tarea es de otra empresa, avisar: debe hacerse en su repositorio, no aquí.

## Regla: nuevas secciones
Antes de crear una sección o herramienta nueva (una carpeta nueva de primer nivel),
**preguntar al usuario** si prefiere:
1. Crearla como carpeta dentro de este repo, o
2. Crear un repositorio nuevo para tenerlo mejor organizado.
No crear la carpeta hasta tener respuesta.

## Estructura de este repo
- `carga-capacidad/` → Panel HTML de carga de capacidad de máquinas (tubos).
  Archivo único autocontenido; usa Supabase (datos) y Chart.js (gráficos) por CDN.

Cada herramienta va en su propia carpeta. `README.md` es el índice para personas.

## Convenciones
- Idioma: español en interfaz, comentarios y mensajes de commit.
- Herramientas en HTML autocontenido (un archivo por herramienta) salvo que se pida otra cosa.
- No subir contraseñas ni claves privadas al repo.

## Trabajo en curso
@PROGRESO.md

Al terminar cada tarea: actualizar `PROGRESO.md` y hacer commit.
