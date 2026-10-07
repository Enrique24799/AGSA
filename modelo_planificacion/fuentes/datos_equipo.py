"""Datos introducidos por el equipo en la herramienta (transcritos de sus capturas del 07/10/2026).

Se aplican sobre la propuesta inicial al generar datos_base.json, para que la nueva versión los traiga ya cargados.
Los nombres se respetan tal cual se escribieron.
"""

# Personas y roles asignados: empresa -> [(rol, titular, % dedicación)], en el orden de la pantalla
PERSONAS = {
    'CORP': [
        ('DIR', 'Enrique Tejada', '100 %'),
        ('PO', 'Manuel Rodriguez', '100 %'),
        ('SIS', 'Daniel Leal', '100 %'),
        ('SIS', 'Jesus Amigo', '100 %'),
        ('SIS', 'Vicente Hernandez', '100 %'),
        ('SIS', 'Juan Trigo', '100 %'),
    ],
    'OPET': [
        ('RPP', 'Victoriano Perez', '30 %'),
        ('PDC', 'Victoriano Perez', '30 %'),
        ('PPR', 'Victoriano Perez', '30 %'),
        ('CMP', 'Victoriano Perez', '10 %'),
        ('CMP', 'Barbara Ayuste', '50 %'),
        ('PMT', 'Barbara Ayuste', '50 %'),
        ('DPL', 'Manuel Castro', '100 %'),
        ('COM', 'Manuel Jesus Trinidad', '100 %'),
        ('COM', 'Monica Gordillo', '50 %'),
        ('MD', 'Monica Gordillo', '20 %'),
        ('LOG', 'Monica Gordillo', '30 %'),
        ('PRO', 'Jefes Turno', '100 %'),
        ('SOP', 'Fernando Tolosa', '100 %'),
        ('CMP', 'Antonio Nogales', '100 %'),
    ],
}

# Organigramas as-is: empresa -> [(id, jefe, puesto, persona, área, rol)]
ORG_ASIS = {
    'OPET': [
        ('o1', '', 'Director Planta', 'Manuel Castro', 'Dirección', 'DPL'),
        ('o2', 'o1', 'Responsable Planificación', 'Victoriano Perez', 'Planificación', 'RPP'),
        ('o3', 'o1', 'Responsable Comercial', 'Manuel Jesus Trinidad', 'Comercial', 'COM'),
        ('o4', 'o3', 'Comercial / Logística Externa', 'Monica Gordillo', 'Comercial', 'COM'),
        ('o5', 'o1', 'Responsable Compras', 'Antonio Nogales', 'Compras', 'CMP'),
        ('o6', 'o5', 'Gestor Compras', 'Barbara Ayuste', 'Compras', 'CMP'),
        ('o7', 'o1', 'Responsable Produccion', 'Juan Manuel Barroso', 'Producción', 'PRO'),
        ('o8', 'o7', 'Jefes Turno', 'Jefes Turno', 'Producción', 'PRO'),
    ],
    'CORP': [
        ('a1', '', 'Responsable Planificacion', 'Enrique Tejada', 'Planificación', 'DIR'),
        ('a2', 'a1', 'Responsable Técnico', 'Daniel Leal', 'Planificación', ''),
        ('a3', 'a1', 'Responsable Procesos', 'Manuel Rodríguez', 'Planificación', 'PO'),
        ('a4', 'a3', 'Programador .NET', 'Jesus Amigo', 'Planificación', 'SIS'),
        ('a5', 'a3', 'Programador .NET', 'Vicente Hernández', 'Planificación', 'SIS'),
        ('a6', 'a3', 'Nuevo puesto', '', '', ''),
        ('a7', 'a3', 'Nuevo puesto', '', '', ''),
    ],
}
