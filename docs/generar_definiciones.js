// Genera docs/Definiciones_Planificacion_AGSA.docx
// Uso: npm install docx (una vez) y después: node docs/generar_definiciones.js
// Para añadir un concepto nuevo, añade su apartado en CONTENIDO y vuelve a generar.

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  LevelFormat, Footer, PageNumber,
} = require("docx");

const VERSION = "Versión 1 · 2 de octubre de 2026";
const SALIDA = path.join(__dirname, "Definiciones_Planificacion_AGSA.docx");

// ---------------------------------------------------------------------------
// Estilo
// ---------------------------------------------------------------------------
const AZUL = "1F3864";
const GRIS = "595959";
const ANCHO = 9638; // A4 con márgenes de 2 cm, en DXA

// Texto con formato en línea: `campo` en monoespaciada y **texto** en negrita
function runs(texto, base = {}) {
  return texto
    .split(/(`[^`]+`|\*\*[^*]+\*\*)/g)
    .filter((t) => t.length > 0)
    .map((t) => {
      if (t.startsWith("`")) {
        return new TextRun({ ...base, text: t.slice(1, -1), font: "Consolas", size: (base.size || 22) - 2 });
      }
      if (t.startsWith("**")) {
        return new TextRun({ ...base, text: t.slice(2, -2), bold: true });
      }
      return new TextRun({ ...base, text: t });
    });
}

const p = (texto) => new Paragraph({ children: runs(texto) });
const intro = (texto) => new Paragraph({ keepNext: true, children: runs(texto) });
const h1 = (texto) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(texto)] });
const h2 = (texto) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(texto)] });
const vineta = (texto) => new Paragraph({ numbering: { reference: "vinetas", level: 0 }, children: runs(texto) });
const regla = (texto) => new Paragraph({ numbering: { reference: "reglas", level: 0 }, children: runs(texto) });

const borde = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };

function tabla(cabecera, filas, anchos) {
  const celda = (texto, i, esCabecera) =>
    new TableCell({
      width: { size: anchos[i], type: WidthType.DXA },
      shading: esCabecera ? { fill: "D9E2F3", type: ShadingType.CLEAR, color: "auto" } : undefined,
      margins: { top: 60, bottom: 60, left: 110, right: 110 },
      children: [new Paragraph({ spacing: { after: 0 }, children: runs(texto, { size: 20, bold: esCabecera }) })],
    });
  return new Table({
    width: { size: ANCHO, type: WidthType.DXA },
    columnWidths: anchos,
    borders: { top: borde, bottom: borde, left: borde, right: borde, insideHorizontal: borde, insideVertical: borde },
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: cabecera.map((t, i) => celda(t, i, true)) }),
      ...filas.map((fila) => new TableRow({ cantSplit: true, children: fila.map((t, i) => celda(t, i, false)) })),
    ],
  });
}

const espacio = () => new Paragraph({ spacing: { after: 60 }, children: [] });

// ---------------------------------------------------------------------------
// Contenido
// ---------------------------------------------------------------------------
const CONTENIDO = [
  // --- Cabecera ---
  new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Definiciones de planificación AGSA", bold: true, size: 48, color: AZUL })] }),
  new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Conceptos y campos de la app de planificación en Qlik", size: 26, color: GRIS })] }),
  new Paragraph({
    spacing: { after: 240 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: AZUL, space: 6 } },
    children: [new TextRun({ text: VERSION, size: 20, color: "7F7F7F" })],
  }),
  p("Este documento recoge cómo se calcula cada indicador de planificación de AGSA y qué contiene cada campo. Se amplía cada vez que se define un concepto nuevo."),

  // --- 1. Fecha errónea ---
  h1("1. Fecha errónea"),
  h2("1.1 Qué indica"),
  p("Indica si la fecha de entrega que se dio al pedido era alcanzable cuando se creó. Para saberlo, se compara la fecha de entrega con una fecha propuesta: la fecha más temprana en la que se considera que el pedido se podía servir."),

  h2("1.2 Cómo se calcula la fecha propuesta"),
  p("Se aplican estas reglas en orden y se usa la primera que se cumple:"),
  regla("**Hay stock suficiente.** Si el día de creación del pedido había stock libre del material para cubrir al menos el 90 % de la cantidad del pedido (tolerancia del 10 %), la fecha propuesta es la fecha de creación más 2 días naturales, por el margen logístico."),
  regla("**La familia está planificada.** Si no hay stock suficiente y la familia del material está planificada, la fecha propuesta es el fin de fabricación de la familia en el primer ciclo en que aparece, sea el actual o el siguiente."),
  regla("**La familia no está planificada y el pedido es de 30 Tn o menos.** Si la familia no está planificada en ningún ciclo y el pedido no supera las 30 Tn, no se calcula fecha propuesta y la fecha se considera errónea directamente."),
  regla("**La familia no está planificada y el pedido supera las 30 Tn.** La fecha propuesta es el fin del ciclo actual de la máquina si a ese ciclo le quedan 5 días o más. Si le quedan menos de 5 días, es el fin del ciclo siguiente."),
  p("No se puede calcular la fecha propuesta, y el resultado es «Sin datos», cuando no hay foto de stock del día de creación del pedido, cuando el material no tiene familia asignada o cuando la máquina no tiene planificado el ciclo que haría falta."),

  h2("1.3 Resultado"),
  tabla(
    ["Fecha_Erronea", "Cuándo"],
    [
      ["«Sí»", "La fecha de entrega es anterior a la fecha propuesta, o la familia no está planificada y el pedido es de 30 Tn o menos."],
      ["«No»", "La fecha de entrega es igual o posterior a la fecha propuesta."],
      ["«Sin datos»", "No se puede calcular la fecha propuesta (ver el apartado 1.2)."],
    ],
    [2000, 7638],
  ),
  espacio(),

  h2("1.4 Origen de la fecha propuesta"),
  intro("El campo `Origen_Propuesta` indica qué regla se ha aplicado a cada posición:"),
  tabla(
    ["Origen_Propuesta", "Cuándo se aplica", "Fecha propuesta"],
    [
      ["«Stock»", "Stock libre ≥ 90 % del pedido el día de creación", "Creación + 2 días"],
      ["«Fabricación de la familia»", "Sin stock suficiente y familia planificada", "Fin de la familia en su primer ciclo planificado"],
      ["«Familia sin planificar (pedido pequeño)»", "Familia sin planificar y pedido de 30 Tn o menos", "Ninguna. Fecha errónea «Sí»"],
      ["«Fin del ciclo actual»", "Familia sin planificar, pedido de más de 30 Tn y al ciclo actual le quedan 5 días o más", "Fin del ciclo actual"],
      ["«Fin del ciclo siguiente»", "Familia sin planificar, pedido de más de 30 Tn y al ciclo actual le quedan menos de 5 días", "Fin del ciclo siguiente. Si no está planificado, «Sin datos»"],
      ["«Sin foto de stock»", "No hay foto de stock del día de creación", "Ninguna. «Sin datos»"],
      ["«Material sin familia»", "Sin stock suficiente y el material no tiene familia", "Ninguna. «Sin datos»"],
    ],
    [3000, 3938, 2700],
  ),
  espacio(),

  h2("1.5 Definiciones y criterios"),
  vineta("**Cantidad del pedido:** la de la posición (pedido-posición), en kg. 30 Tn son 30.000 kg."),
  vineta("**Stock libre:** stock de libre utilización del material no asignado a pedidos, según la foto de stock del día de creación del pedido."),
  vineta("**Ciclo actual de una máquina:** el primer ciclo planificado que todavía no ha terminado. **Fin de un ciclo:** el último fin previsto de las fabricaciones de ese ciclo."),
  vineta("**Fin de fabricación de la familia:** el último fin previsto de las fabricaciones de la familia dentro de ese ciclo."),
  vineta("**Planificación que se tiene en cuenta:** fabricaciones pendientes (estatus 0 y 1) de las máquinas 8, 16, 17, 20 y 28."),
  vineta("**Días que le quedan al ciclo:** días naturales desde el día de la recarga hasta el fin del ciclo."),
  vineta("**Momento del cálculo:** el stock se mira el día de creación del pedido, pero la planificación de familias y ciclos es la del día de la recarga. Por eso el resultado de un mismo pedido puede cambiar de un día a otro."),

  // --- 2. Tipo de atraso ---
  h1("2. Tipo de atraso"),
  h2("2.1 Qué indica"),
  p("Clasifica cada posición abierta de la cartera según si va con retraso y, si es así, en qué parte del proceso está la causa."),

  h2("2.2 Cómo se calcula"),
  intro("Primero se mira si la posición está reservada: lo está cuando el stock asignado a ella cubre al menos el 90 % de su cantidad pendiente. Después se aplica la primera fila que se cumple:"),
  tabla(
    ["Reservada", "Condición", "Tipo_Atraso"],
    [
      ["Sí", "La fecha de entrega es hoy o posterior", "«Sin Atraso»"],
      ["Sí", "La fecha de entrega ya ha pasado y la reserva se hizo el día de la entrega o después", "«Atraso Produccion»"],
      ["Sí", "La fecha de entrega ya ha pasado y la reserva se hizo antes de la fecha de entrega", "«Atraso Expedicion»"],
      ["No", "No hay fecha de disponibilidad, o es posterior a la fecha de entrega", "«Atraso Planificacion»"],
      ["No", "La fecha de disponibilidad es igual o anterior a la fecha de entrega", "«Sin Atraso»"],
    ],
    [1450, 5588, 2600],
  ),
  espacio(),

  h2("2.3 Qué significa cada tipo"),
  vineta("**Atraso Produccion:** el material se reservó el día de la entrega o después, así que fabricación llegó tarde."),
  vineta("**Atraso Expedicion:** el material estaba reservado antes de la fecha de entrega, esa fecha ya ha pasado y la posición sigue pendiente. El retraso está en la salida del pedido."),
  vineta("**Atraso Planificacion:** el material no está reservado y planificación no da fecha de disponibilidad o la da después de la fecha de entrega. Incluye posiciones que aún no han vencido: es un atraso previsto."),
  vineta("**Sin Atraso:** ninguno de los casos anteriores."),

  h2("2.4 Criterios"),
  vineta("La cartera solo incluye posiciones abiertas con cantidad pendiente. Si la cantidad pendiente es 0, el pedido se debe cerrar."),
  vineta("Toda reserva tiene fecha de reserva."),
  vineta("Planificación actualiza la fecha de disponibilidad como mínimo al día actual."),

  // --- 3. Situación por material ---
  h1("3. Situación por material"),
  h2("3.1 Qué es"),
  p("Una fila por material activo de AGSA con su stock, la cartera pendiente, las órdenes de fabricación pendientes y la primera fabricación prevista. Se calcula en cada recarga y se guarda en `Situacion_Material_AGSA.qvd`. Todas las cantidades van en toneladas. Los campos se definen en el apartado 4.2."),

  h2("3.2 Materiales incluidos"),
  vineta("**Materiales activos:** status del material en el centro (MMSTA) vacío o «VC». Un material bloqueado no aparece aunque tenga stock o pedidos."),
  vineta("**Solo materiales de AGSA, según la regla de centro:** tubos en C412; flejes, bobinas y chapas en C414; perfiles en cualquiera de los dos."),
  vineta("Si un material aparece en los dos centros, se queda una sola fila."),

  h2("3.3 Cómo se calcula"),
  vineta("**Stock:** última foto de stock de AGSA (unidad KG, lotes que empiezan por 2, tipos de stock «Stock» y «Pedido Cliente»)."),
  vineta("**Pedidos pendientes:** cantidad pendiente de las posiciones abiertas de pedidos de venta."),
  vineta("**Órdenes pendientes:** fabricaciones con estatus 0 y 1 del planificador, en todas las máquinas. Cuenta lo que falta por fabricar: cantidad planificada menos fabricada."),
  vineta("**Libre tras pedidos:** stock total menos pedidos pendientes más órdenes pendientes. Se usa el stock total porque el stock reservado también está dentro de la cantidad pendiente de los pedidos a los que está asignado, así que se compensa."),
  vineta("**Stock total frente a reservado y libre:** el total no es la suma de los dos. La diferencia es el stock bloqueado o en control de calidad que no está asignado a ningún pedido."),
  vineta("**Próxima fabricación:** fin previsto de la primera orden pendiente. Si una orden está en curso o retrasada, la fecha puede estar ya pasada."),

  // --- 4. Diccionario de campos ---
  h1("4. Diccionario de campos"),
  h2("4.1 Cartera de pedidos"),
  intro("Filas con `Tabla` = «Historico_Cartera». Una fila por posición abierta y por día de recarga. El histórico se guarda en `Cartera_AGSA.qvd` desde el 23/07/2026."),
  tabla(
    ["Campo", "Descripción"],
    [
      ["`Fecha_Dato`", "Día de la recarga. Cada recarga guarda una foto de la cartera abierta."],
      ["`Key`", "Año y mes de creación del pedido más el material. Enlaza con el maestro de materiales."],
      ["`Pedido_Posicion`", "Número de pedido y posición en SAP."],
      ["`Centro_Material`", "Centro SAP de la posición."],
      ["`Id_Material`", "Código SAP del material."],
      ["`Id_Cliente`", "Cliente solicitante."],
      ["`Id_Comercial`", "Comercial (grupo de vendedores en SAP)."],
      ["`Cantidad_Pedido`", "Cantidad de la posición, en kg."],
      ["`Importe_Pedido`", "Valor neto de la posición, en euros."],
      ["`Precio_Unitario`", "Importe_Pedido / Cantidad_Pedido, en €/kg."],
      ["`añomes_pedido`", "Año y mes de la fecha del pedido (AAAA_MM)."],
      ["`añomes_entrega`", "Año y mes de la fecha de entrega (AAAA_MM)."],
      ["`Cantidad_Expedida`", "Kg ya expedidos de la posición."],
      ["`Fecha_Creacion`", "Fecha de creación del pedido en SAP."],
      ["`Fecha_Entrega`", "Fecha de entrega de la posición (primer reparto)."],
      ["`Fecha`", "Fecha de la fila en el modelo. En la cartera es la fecha de entrega."],
      ["`Fecha_Salida`", "Fecha de la última salida de mercancía de la posición."],
      ["`Estado`", "En la cartera: «Atraso Pedido Abierto» si la fecha de entrega ya ha pasado y «A tiempo Pedido Abierto» si no."],
      ["`Cantidad_Pendiente`", "Cantidad_Pedido menos Cantidad_Expedida, en kg. Nunca es negativa."],
      ["`Satus_Rechazo`", "«A» si la posición está abierta y tiene cantidad pendiente; «C» si está cerrada. La cartera solo incluye «A»."],
      ["`Tabla`", "«Historico_Cartera»."],
      ["`Reservado`", "Kg de stock asignados a la posición: stock de tipo «Pedido Cliente», sumando libre, bloqueado y en control de calidad."],
      ["`FechaReserva`", "Fecha de la primera reserva de stock para la posición."],
      ["`Fecha_Disponibilidad`", "Fecha en la que planificación prevé tener el material de la posición. Vacía si planificación no ha dado fecha. Planificación la actualiza como mínimo al día actual."],
      ["`Familia`", "Familia del material (material → medida → familia)."],
      ["`Maquina`", "Máquina en la que se fabrica la familia."],
      ["`Foto_Stock`", "1 si existe foto de stock del día de creación del pedido; vacío si no."],
      ["`Stock_Libre_Pedido`", "Kg de stock libre del material el día de creación del pedido. 0 si ese día hay foto pero el material no aparece."],
      ["`Prox_Fin_Familia`", "Fin de fabricación de la familia en el primer ciclo en que está planificada. Vacío si no está planificada."],
      ["`Fin_Ciclo_Actual`", "Fin del ciclo en curso de la máquina."],
      ["`Fin_Ciclo_Proximo`", "Fin del ciclo siguiente de la máquina."],
      ["`Tipo_Atraso`", "Clasificación del atraso (apartado 2)."],
      ["`Origen_Propuesta`", "Regla que ha dado la fecha propuesta (apartado 1.4)."],
      ["`Fecha_Propuesta`", "Fecha más temprana en la que se considera que el pedido se podía servir (apartado 1.2)."],
      ["`Fecha_Erronea`", "«Sí», «No» o «Sin datos» (apartado 1.3)."],
    ],
    [2700, 6938],
  ),
  espacio(),

  h2("4.2 Situación por material"),
  intro("Filas con `Tabla` = «Situacion Material». En el modelo, el nombre, el tipo, la clase ABC y la estrategia llegan a través de `Key`, desde el maestro de materiales. El fichero `Situacion_Material_AGSA.qvd` los lleva todos."),
  tabla(
    ["Campo", "Descripción"],
    [
      ["`Tabla`", "«Situacion Material»."],
      ["`Fecha`", "Día de la recarga."],
      ["`Key`", "Año y mes de la recarga más el material. Enlaza con el maestro de materiales."],
      ["`Id_Material`", "Código SAP del material."],
      ["`Nombre_Material`", "Descripción del material."],
      ["`Tipo_Material`", "«Fleje», «Tubo», «Bobina», «Chapa» o «Perfil», según el código del material. En el modelo es `Tipo_Material_2`, del maestro de materiales."],
      ["`ABC`", "«A», «B» o «C», a partir de la clave ABC de SAP."],
      ["`Estrategia`", "Estrategia de planificación del material en el centro (grupo de planificación de SAP)."],
      ["`Stock_Total_Tn`", "Todo el stock del material: libre, bloqueado y en control de calidad, esté asignado a pedidos o no."],
      ["`Stock_Reservado_Tn`", "Stock asignado a pedidos de cliente."],
      ["`Stock_Libre_Tn`", "Stock de libre utilización no asignado a pedidos."],
      ["`Pedidos_Pendientes_Tn`", "Cantidad pendiente de servir de las posiciones abiertas."],
      ["`Proxima_Fabricacion`", "Fin previsto de la primera orden de fabricación pendiente del material. «SIN FABRICACION» si no tiene ninguna."],
      ["`Ordenes_Pendientes_Tn`", "Cantidad que falta por fabricar en las órdenes pendientes."],
      ["`Tn_Libre_Tras_Pedidos`", "Stock_Total_Tn − Pedidos_Pendientes_Tn + Ordenes_Pendientes_Tn."],
    ],
    [2700, 6938],
  ),
  espacio(),

  // --- 5. Pendiente de confirmar ---
  h1("5. Pendiente de confirmar"),
  vineta("**Límite de 30 Tn:** se aplica a la cantidad de la posición. Si debe ser el total del pedido (todas sus posiciones), hay que cambiar el cálculo."),
  vineta("**Materiales sin familia:** cuando no hay stock suficiente salen «Sin datos» en fecha errónea, no «Sí»."),
  vineta("**Unidad de las órdenes de fabricación:** se asume que la cantidad planificada viene en kg."),
];

// ---------------------------------------------------------------------------
// Documento
// ---------------------------------------------------------------------------
const doc = new Document({
  creator: "Planificación AGSA",
  title: "Definiciones de planificación AGSA",
  styles: {
    default: { document: { run: { font: "Calibri", size: 22 }, paragraph: { spacing: { after: 120, line: 276 } } } },
    paragraphStyles: [
      {
        id: "Pie", name: "Pie", basedOn: "Normal", quickFormat: true,
        run: { size: 18, color: "7F7F7F" },
        paragraph: { spacing: { after: 0 } },
      },
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, color: AZUL },
        paragraph: { spacing: { before: 360, after: 120 }, keepNext: true },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, color: AZUL },
        paragraph: { spacing: { before: 240, after: 80 }, keepNext: true },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "vinetas",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 567, hanging: 283 } } } }],
      },
      {
        reference: "reglas",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 567, hanging: 340 } } } }],
      },
    ],
  },
  sections: [
    {
      properties: { page: { margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              style: "Pie",
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun("Definiciones de planificación AGSA · Página "),
                new TextRun({ children: [PageNumber.CURRENT] }),
                new TextRun(" de "),
                new TextRun({ children: [PageNumber.TOTAL_PAGES] }),
              ],
            }),
          ],
        }),
      },
      children: CONTENIDO,
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(SALIDA, buffer);
  console.log("Generado:", SALIDA);
});
