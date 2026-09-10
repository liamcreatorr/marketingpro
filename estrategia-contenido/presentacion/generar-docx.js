/* Genera el documento de estrategia de contenido de Kenia.
   Reutiliza el sistema de estilos de branding/presentacion/generar-docx.js
   para que los dos entregables se vean como la misma casa.
   Uso: node generar-docx.js "salida.docx"                                  */

const fs = require('fs');
const d = require('docx');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Footer,
  WidthType, ShadingType, BorderStyle, AlignmentType, HeadingLevel,
  PageBreak, LevelFormat
} = d;

/* ---------------- paleta Kenia ---------------- */
const CIAN = "079CB5", PETROLEO = "00677E", NEGRO = "1B1B1B",
      GRIS = "3C3C3C", GRIS_CLARO = "6E7F82", NARANJA = "FF5A00",
      FONDO = "EAF2F4", FONDO_2 = "F2F6F7", LINEA = "C7D9DD";
const SERIF = "Georgia", SANS = "Calibri";
const W = 9648;
const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

function runs(text, base = {}) {
  const out = [];
  for (const part of String(text).split(/(\*\*[^*]+\*\*)/g)) {
    if (!part) continue;
    if (part.startsWith("**") && part.endsWith("**"))
      out.push(new TextRun({ ...base, text: part.slice(2, -2), bold: true }));
    else out.push(new TextRun({ ...base, text: part }));
  }
  return out;
}

const p = (t, o = {}) => new Paragraph({
  spacing: { after: o.after ?? 150, line: 288 },
  children: runs(t, { font: SANS, size: 21, color: o.color || NEGRO }),
});

const eyebrow = (t, color = CIAN) => new Paragraph({
  spacing: { after: 90 },
  children: [new TextRun({ text: t.toUpperCase(), font: SANS, size: 16, bold: true, color, characterSpacing: 30 })],
});

let secN = 0;
function h1(t) {
  secN += 1;
  return [
    new Paragraph({
      spacing: { before: 60, after: 50 },
      children: [new TextRun({ text: String(secN).padStart(2, "0"), font: SANS, size: 18, bold: true, color: CIAN, characterSpacing: 30 })],
    }),
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 0, after: 200 },
      children: [new TextRun({ text: t, font: SERIF, size: 32, bold: true, color: NEGRO })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: CIAN, space: 8 } },
    }),
  ];
}

const h2 = t => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 320, after: 110 },
  children: [new TextRun({ text: t, font: SERIF, size: 24, bold: true, color: PETROLEO })],
});

const h3 = t => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before: 240, after: 80 },
  children: [new TextRun({ text: t, font: SANS, size: 21, bold: true, color: NEGRO })],
});

function boxed(kids, accent, fill) {
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [W],
    borders: { top: NONE, bottom: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE,
               left: { style: BorderStyle.SINGLE, size: 18, color: accent } },
    rows: [new TableRow({ children: [new TableCell({
      width: { size: W, type: WidthType.DXA },
      shading: fill ? { type: ShadingType.CLEAR, fill, color: "auto" } : undefined,
      margins: { top: 170, bottom: 170, left: 270, right: 220 },
      children: kids,
    })] })],
  });
}

function quote(text, attr) {
  const kids = [new Paragraph({
    spacing: { after: attr ? 60 : 0, line: 288 },
    children: runs(text, { font: SERIF, size: 25, color: PETROLEO }),
  })];
  if (attr) kids.push(new Paragraph({
    spacing: { after: 0 },
    children: [new TextRun({ text: attr.toUpperCase(), font: SANS, size: 15, bold: true, color: GRIS_CLARO, characterSpacing: 26 })],
  }));
  return boxed(kids, CIAN, null);
}

function panel(title, lines, accent = CIAN) {
  const kids = [];
  if (title) kids.push(new Paragraph({
    spacing: { after: 110 },
    children: [new TextRun({ text: title.toUpperCase(), font: SANS, size: 16, bold: true, color: accent, characterSpacing: 30 })],
  }));
  lines.forEach((l, i) => {
    const last = i === lines.length - 1;
    if (typeof l === "string") kids.push(new Paragraph({
      spacing: { after: last ? 0 : 130, line: 288 },
      children: runs(l, { font: SANS, size: 21, color: NEGRO }),
    }));
    else kids.push(l);
  });
  return boxed(kids, accent, FONDO);
}

const bullet = (t, level = 0) => new Paragraph({
  numbering: { reference: "vinetas", level },
  spacing: { after: 80, line: 288 },
  children: runs(t, { font: SANS, size: 21, color: NEGRO }),
});

function table(headers, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  if (total !== W) throw new Error("anchos suman " + total + ", deben sumar " + W);
  const head = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: PETROLEO, color: "auto" },
      margins: { top: 110, bottom: 110, left: 150, right: 150 },
      children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({
        text: String(h).toUpperCase(), font: SANS, size: 16, bold: true, color: "FFFFFF", characterSpacing: 22,
      })] })],
    })),
  });
  const body = rows.map((r, ri) => new TableRow({
    children: r.map((c, i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: ri % 2 ? FONDO_2 : "FFFFFF", color: "auto" },
      margins: { top: 110, bottom: 110, left: 150, right: 150 },
      children: [new Paragraph({ spacing: { after: 0, line: 264 }, children: runs(c, { font: SANS, size: 19, color: NEGRO }) })],
    })),
  }));
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: widths,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: LINEA },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: LINEA },
      left: { style: BorderStyle.SINGLE, size: 4, color: LINEA },
      right: { style: BorderStyle.SINGLE, size: 4, color: LINEA },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: LINEA },
      insideVertical: { style: BorderStyle.SINGLE, size: 2, color: LINEA },
    },
    rows: [head, ...body],
  });
}

const gap = (h = 200) => new Paragraph({ spacing: { after: h }, children: [] });
const brk = () => new Paragraph({ children: [new PageBreak()] });

/* Semana del calendario: cabecera propia + tabla de piezas */
function semana(cod, titulo, fechas, trabajo, piezas, clave) {
  const el = [
    new Paragraph({
      spacing: { before: 320, after: 0 },
      children: [
        new TextRun({ text: cod + "   ", font: SANS, size: 18, bold: true, color: clave ? NARANJA : CIAN, characterSpacing: 20 }),
        new TextRun({ text: titulo, font: SERIF, size: 25, bold: true, color: NEGRO }),
        new TextRun({ text: "   " + fechas, font: SANS, size: 19, color: GRIS_CLARO }),
      ],
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: clave ? NARANJA : LINEA, space: 6 } },
    }),
    new Paragraph({ spacing: { before: 140, after: 160, line: 288 }, children: runs(trabajo, { font: SANS, size: 21, color: GRIS }) }),
  ];
  if (piezas.length) el.push(table(["Pieza", "Día", "Qué es", "Embudo"], piezas, [700, 900, 6248, 1800]));
  return el;
}

/* Bloque de pendientes con casilla para marcar */
function pendientes(titulo, quien, filas, accent = NARANJA) {
  return [
    new Paragraph({
      spacing: { before: 300, after: 0 },
      children: [
        new TextRun({ text: titulo, font: SERIF, size: 24, bold: true, color: NEGRO }),
        new TextRun({ text: "   " + quien, font: SANS, size: 18, bold: true, color: accent, characterSpacing: 20 }),
      ],
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: LINEA, space: 6 } },
    }),
    gap(140),
    table(["", "Qué falta", "Quién", "Cuándo"],
      filas.map(f => ["☐", f[0], f[1], f[2]]),
      [400, 5748, 1700, 1800]),
  ];
}

/* ================= CONTENIDO ================= */
const C = [];
const add = (...xs) => xs.forEach(x => Array.isArray(x) ? C.push(...x) : C.push(x));

/* ---------- PORTADA ---------- */
add(
  gap(1000),
  eyebrow("Organic Club · Documento interno"),
  new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Kenia", font: SERIF, size: 100, bold: true, color: NEGRO })] }),
  new Paragraph({ spacing: { after: 260 }, children: [new TextRun({ text: "Estrategia de contenido y plan de lanzamiento", font: SERIF, size: 34, color: GRIS })] }),
  new Paragraph({ spacing: { after: 460 }, indent: { right: 7900 }, children: [],
    border: { bottom: { style: BorderStyle.SINGLE, size: 30, color: NARANJA, space: 2 } } }),
  table(["Campo", "Detalle"], [
    ["Para", "Alonzo García"],
    ["Preparado por", "Organic Club — Liam Libre"],
    ["Cliente", "Ninro Libre — fundador y CEO de Kenia"],
    ["Fecha", "10 de septiembre de 2026"],
    ["Versión", "1.1 — auditada"],
    ["Base", "La plataforma de marca cerrada el 8 de septiembre, las 12 entrevistas a corredores de Caracas, el PESTEL y el análisis de competencia"],
  ], [2100, 7548]),
  gap(420),
  p("Este documento es la estrategia de contenido completa: embudo, relato, pilares, formatos, calendario del primer mes y plan de lanzamiento. Y, sobre todo, **lo que todavía falta y de quién depende**, que está en la sección 16."),
  p("No es la presentación al cliente. Es el documento con el que trabajamos nosotros, así que dice también lo que no sabemos y lo que salió mal en la primera versión."),
  brk(),
);

/* ---------- ÍNDICE ---------- */
add(
  new Paragraph({ spacing: { after: 240 }, children: [new TextRun({ text: "Contenido", font: SERIF, size: 34, bold: true, color: NEGRO })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: CIAN, space: 8 } } }),
  gap(160),
  table(["#", "Sección", "Qué encontrarás"], [
    ["01", "Dónde estamos", "La decisión que ordena el mes y las cuatro cosas del calendario que mandan"],
    ["02", "Lo urgente: los tres días antes del rodaje", "Lo que tiene que estar el viernes 12 o el domingo se improvisa"],
    ["03", "El relato", "Héroe, villano, mentor y la transformación. El problema interno, que es donde se gana"],
    ["04", "Jerarquía emocional", "Las siete emociones ordenadas y las cinco prohibidas"],
    ["05", "El embudo", "La escalera de mensajes, el reparto TOFU/MOFU/BOFU y el CTA del mes sin app"],
    ["06", "Los seis objetivos de contenido", "Los cinco de la lluvia de ideas más el que añadimos"],
    ["07", "Los cinco pilares", "Con sus subtemas, sus pesos y sus guardarraíles"],
    ["08", "Formatos y matriz de producción", "Los seis formatos y la tabla que decide si el mes se puede ejecutar"],
    ["09", "Plan de canales", "Qué va en cada uno, qué no, y lo que decidimos no hacer"],
    ["10", "El calendario del mes", "Las 19 piezas, semana a semana"],
    ["11", "El rodaje del 13 de septiembre", "Los seis bloques y el orden del día"],
    ["12", "El plan de lanzamiento", "Los tres movimientos y el video de lanzamiento"],
    ["13", "Tagline, frases y comunidad", "Lo que hay que aprobar antes del viernes"],
    ["14", "Qué de la lluvia de ideas no va", "Las doce ideas, revisadas una por una"],
    ["15", "Métricas y gobierno", "Metas del mes 1 y cómo se decide qué matar"],
    ["16", "Lo que falta y de quién depende", "27 pendientes, agrupados por responsable"],
    ["17", "Cómo se auditó", "Los 17 errores de la versión 1 y qué se corrigió"],
  ], [640, 3400, 5608]),
  brk(),
);

/* ---------- 01 ---------- */
add(
  h1("Dónde estamos"),
  quote("El mes que viene no vende la app. Construye la lista que la descarga el día uno.", "La decisión que ordena todo"),
  gap(180),
  p("Entre el rodaje del **13 de septiembre** y el lanzamiento de mediados de octubre hay un mes en el que se puede publicar pero **no se puede descargar nada**. Ese mes tiene un solo trabajo: llegar al día del lanzamiento con **400 personas esperando**. Todo lo demás es consecuencia de eso."),

  h2("Las cuatro cosas del calendario que mandan"),
  table(["Hecho", "Qué significa para contenido"], [
    ["Ninro se graba una sola vez", "Viaja a Buenos Aires del 15 al 20 y después el calendario aprieta. **Siete de los trece videos del mes llevan su cara** y todos salen del 13 de septiembre."],
    ["Caracas Rock cae antes del lanzamiento", "4 de octubre. Cuatro de los doce entrevistados la corren y se la pierden como usuarios. Es donde está el público, once días antes: se va a captar, no a vender."],
    ["Octubre es el mes más lluvioso del año", "Unos 228 mm y casi treinta días con lluvia, justo en la ventana donde se mide la retención a 30 días. Es a la vez el mayor riesgo y el mejor territorio de contenido."],
    ["La campaña se ancla a noviembre", "La Media Maratón de Hipereventos. Da fecha, urgencia y un motivo para suscribirse en octubre en vez de «algún día». Tres de doce la corren."],
  ], [2900, 6748]),
  gap(200),
  panel("La regla que manda sobre todo el proyecto", [
    "Cada vez que se predijo el comportamiento del público sin datos, la predicción salió mal. Hay **12 entrevistas** y un PESTEL cerrado.",
    "Por eso, en este documento, **todo lo que es hipótesis va marcado como hipótesis** y todo lo que es hueco abierto está listado en la sección 16 sin rellenar. No se inventa nada para que el plan se vea completo.",
  ]),

  h2("Qué se produce"),
  table(["Formato", "Cantidad", "Detalle"], [
    ["Videos", "13", "12 del mes más el video de lanzamiento"],
    ["Carruseles", "6", "3 del mes más el pack de tres del lanzamiento"],
    ["Stories", "3 a 5 diarias", "Con caja de preguntas fija los martes"],
    ["Presencial", "1 evento", "Caracas Rock, 4 de octubre"],
  ], [2000, 1500, 6148]),
  brk(),
);

/* ---------- 02 ---------- */
add(
  h1("Lo urgente: los tres días antes del rodaje"),
  p("Si algo de esto no está el viernes 12, el domingo se improvisa. Y un rodaje improvisado con la única persona disponible del mes es el peor escenario del plan entero."),
  gap(120),
  table(["Cuándo", "Qué", "Quién"], [
    ["jueves 11", "**Confirmar los 8–10 del box**, con nombre y hora de llegada", "Ninro"],
    ["jueves 11", "**Describir el reel de referencia** que pasó Kurt", "Kurt · Liam"],
    ["viernes 12", "**Aprobar el tagline y las cuatro frases** — se graban a cámara", "Ninro"],
    ["viernes 12", "Ropa sin logos de terceros visibles", "Ninro"],
    ["viernes 12", "Guion cerrado, shot list impresa, tarjetas y baterías", "Nosotros"],
    ["domingo 13", "5:30 — b-roll de hora azul", "Nosotros"],
    ["domingo 13", "6:00 — testimonios del box", "Todos"],
    ["domingo 13", "11:00 — pedir las diez reseñas de tienda", "Ninro"],
  ], [1600, 6248, 1800]),
  gap(200),
  panel("Por qué las reseñas se piden un mes antes", [
    "**La app sale a tiendas sin una sola reseña, sin testimonios y sin casos**, prometiendo que se adapta y que previene lesiones.",
    "La única fuente gratis somos nosotros mismos: los doce entrevistados y la gente del box. Se piden veinte para conseguir diez, y no se paga, no se cambia por nada y no se dicta el texto: Apple y Google penalizan las reseñas incentivadas, y una sanción de tienda en la semana del lanzamiento sería peor que no tener ninguna.",
  ], NARANJA),
  brk(),
);

/* ---------- 03 ---------- */
add(
  h1("El relato"),
  quote("Nadie deja de correr por falta de voluntad. Deja de correr porque sigue un plan que no sabe nada de su vida.", "El relato maestro, en una frase"),
  gap(180),
  p("El arquetipo ya estaba decidido en la plataforma de marca: **el usuario es el héroe, Kenia es el mentor**. Lo que faltaba era bajarlo a estructura de relato, y ahí aparece la pieza que decide este mercado."),
  gap(120),
  table(["Pieza del relato", "En Kenia"], [
    ["El héroe", "Quien ya corre y no tiene quien decida con él"],
    ["Lo que quiere", "Llegar a su meta sin abandonar en el intento"],
    ["Su problema externo", "El plan que sigue no sabe nada de su vida. Cuatro de doce siguen el de un amigo"],
    ["Su problema interno", "**Cree que el que falla es él.** Once de doce no se llaman corredores"],
    ["Su problema de fondo", "Nadie debería tener que ganarse el derecho a entrenar"],
    ["El villano", "No es una persona ni una app: son los dos enemigos"],
    ["El mentor", "Empatía —«¿cómo amaneciste?»— más autoridad —método de autor"],
    ["El fracaso que evita", "Abandonar otra vez, y esta vez creyendo que fue culpa suya"],
    ["La transformación", "De «siempre lo dejo» a «cuando falto, el plan se mueve conmigo»"],
  ], [2900, 6748]),
  gap(200),
  panel("La razón de que esto importe", [
    "Casi todas las marcas de fitness cuentan solo el problema externo —no tienes plan, toma un plan— y por eso todas suenan igual.",
    "**Quien resuelve el problema interno se queda con el cliente. Quien resuelve solo el externo compite en precio contra Garmin, que es gratis.**",
  ]),

  h2("Los tres relatos, y cuánto pesa cada uno"),
  p("La marca cuenta tres historias distintas. Confundirlas es lo que hace que una cuenta suene a producto un día y a diario personal al siguiente."),
  table(["Relato", "Peso", "Para qué sirve", "Riesgo si se pasa"], [
    ["El del usuario", "60 %", "Legitimar y atraer. Es el que trae público nuevo", "Ninguno. Es el correcto"],
    ["El del método", "30 %", "Autoridad. Es lo que sostiene los 19,99 al mes", "Suena a clase. Se arregla con caras, no con más datos"],
    ["El del fundador", "10 %", "Confianza antes de tener usuarios", "**La marca se vuelve él**, y el arquetipo se corre a Héroe"],
  ], [1900, 800, 3474, 3474]),
  gap(200),
  panel("Veto permanente", [
    "La historia de transformación física de Ninro **no se graba, no se menciona y no se guarda**, en ninguna duración del relato.",
    "El relato se construye sobre su cruzada —«el que hace CrossFit no corre»— y sobre el método. La decisión no se reabre.",
  ], NARANJA),

  h2("Las siete estructuras de guion"),
  p("Plantillas que se rellenan. Cada una de las 19 piezas del mes usa una."),
  table(["Estructura", "Para qué", "Error típico"], [
    ["G1 · La negación", "Legitimar. Pregunta, negativas en cadena, el dato que las contradice, la vuelta", "Cerrar con «tú también eres corredor»: conserva el título y solo lo reparte"],
    ["G2 · El día que cambió", "Producto. Un martes normal, una variable cambia, el plan reacciona, el porqué", "Narrar lo que se ve. El producto habla solo; la voz solo explica el porqué"],
    ["G3 · La corrección amable", "Educar. Lo que hace casi todo el mundo, por qué falla, qué cambia, el permiso", "El primer beat en segunda persona. «No hagas esto» pone el sujeto en la persona"],
    ["G4 · El regreso", "La caída, el silencio, el primer gesto, el plan que estaba ahí", "Convertirlo en épica. **El regreso no es heroico: es normal**, y ahí está su fuerza"],
    ["G5 · La comparación honesta", "Convertir. Se pone el otro objeto delante con respeto y se dice qué no puede saber", "Comparar personas en vez de objetos. Se comparan planes, nunca gente"],
    ["G6 · El lugar", "Atraer. Sonido real, el sitio reconocible, una frase, nada más", "Añadir. Funciona por lo que le falta"],
    ["G7 · La respuesta directa", "Autoridad. La pregunta real de alguien, la respuesta, el porqué, el límite", "Saltarse el límite. Decir dónde no aplica es lo que hace creíble el resto"],
  ], [2200, 3724, 3724]),
  brk(),
);

/* ---------- 04 ---------- */
add(
  h1("Jerarquía emocional"),
  panel("La corrección que dieron las entrevistas", [
    "Se dio por hecho que la emoción a atacar era la culpa de descansar. **Es falso y está medido: 10 de 12 descansan sin conflicto** y ya tienen interiorizado que el descanso es parte del entrenamiento.",
    "**La emoción raíz es la culpa del que falló**, y sobre todo el peso de lo que cuesta retomar. Todo el sistema emocional de la marca se construye sobre eso.",
  ], NARANJA),
  gap(220),
  table(["#", "Emoción", "Qué libera", "Frase que la encarna"], [
    ["1", "**Volver sin castigo**", "Que perder tres semanas no sea empezar de cero", "«Tu plan te esperó.»"],
    ["2", "Legitimidad", "Poder entrenar sin sentirse impostor", "«Si corres, ya estás dentro.»"],
    ["3", "Alivio", "De decidir cada día si entrenar, cuánto y cómo", "«Hoy toca esto. Y por esto.»"],
    ["4", "Seguridad", "Del miedo a repetir la lesión", "«Reducimos el riesgo» — nunca «no te vas a lesionar»"],
    ["5", "Ser entendido", "De la soledad del que entrena solo", "«¿Cómo amaneciste hoy?»"],
    ["6", "Pertenencia", "Del ranking que te deja siempre abajo", "«Gente que corre.»"],
    ["7", "Orgullo silencioso", "El de haber sostenido, no el de la foto en la meta", "«Tu mejor 5K. Y llevas siete semanas sin saltarte una.»"],
  ], [400, 2100, 3574, 3574]),

  h2("Las cinco emociones prohibidas"),
  table(["Emoción", "Por qué está vetada"], [
    ["Vergüenza", "Es el valor 1 al revés. Es la palanca más barata del fitness y la marca renunció a ella por escrito"],
    ["Culpa", "El enemigo declarado, y la emoción raíz del público"],
    ["Euforia", "Envejece mal. **La emoción de Kenia tiene que seguir funcionando el martes malo**"],
    ["Superioridad", "No puedes legitimar a uno y burlarte de otro igual"],
    ["Miedo", "Salvo Patricia, cuyo miedo es anterior a la marca: **se reconoce, no se enciende**. «Esto es lo que te rompió» es reconocer; «esto te va a romper» es amenazar"],
  ], [1900, 7748]),
  gap(200),
  panel("La excepción declarada: Gabriel", [
    "La persona 5 —el que va por el sub-3— **no se compra con ninguna de las siete**. A él lo convencen el argumento técnico y lo que Elite añade, y es el único que detecta al instante si el argumento es hueco.",
    "Cualquier pieza dirigida a él se escribe sin capa emocional.",
  ]),
  brk(),
);

/* ---------- 05 ---------- */
add(
  h1("El embudo — que son dos, no uno"),
  p("El embudo de ventas decide **cuánto** contenido de cada tipo se produce. El embudo comunicacional decide **qué se dice**. En Kenia manda el segundo, porque el obstáculo dominante no es «no conozco Kenia» sino «esto no es para mí»."),

  h2("La escalera de mensajes"),
  table(["#", "Peldaño", "Lo que la persona piensa", "Lo que Kenia le dice"], [
    ["1", "Legitimidad", "«Yo no soy corredor. Me falta tiempo, estructura, disciplina»", "**«Si corres, ya estás dentro.»**"],
    ["2", "Problema", "«Corro, pero cada semana improviso o sigo un plan que no me calza»", "**«Tu semana cambia. Tu plan también.»**"],
    ["3", "Mecanismo", "«Vale, ¿y esto qué hace distinto?»", "**«Te pregunta cómo amaneciste antes de mandarte a correr.»**"],
    ["4", "Justificación", "«¿Y por qué pagaría, si Garmin Coach es gratis?»", "**«Un entrenador en Caracas cuesta 30 a 80 dólares al mes. Kenia Coach, 19,99.»**"],
    ["5", "Permanencia", "«Fallé una semana. Ya para qué»", "**«Tu plan te esperó.»**"],
  ], [400, 1500, 3874, 3874]),
  gap(200),
  quote("Una pieza puede saltar un peldaño, nunca dos. Un video que va del 1 al 4 no convierte: ofende.", "La regla operativa"),

  h2("El reparto TOFU / MOFU / BOFU"),
  p("Propusimos 40/40/20. Es correcto **como media del mes, no como cuota semanal**: durante el mes sin app el BOFU no tiene a dónde mandar a nadie."),
  table(["Fase", "TOFU", "MOFU", "BOFU", "Por qué"], [
    ["Mes sin app · 14 sep a 14 oct", "45 %", "45 %", "10 %", "El BOFU no es una pieza: es el CTA al canal en el pie de las de MOFU"],
    ["Semana de lanzamiento · 15 a 18 oct", "0 %", "25 %", "75 %", "La única semana con permiso para vender de frente"],
    ["**Media del mes**", "**37 %**", "**42 %**", "**21 %**", "Es, en la práctica, el 40/40/20 que propusimos"],
    ["Régimen · desde noviembre", "40 %", "40 %", "20 %", "Ya sí semana a semana"],
  ], [2448, 700, 700, 700, 5100]),

  h2("El CTA del mes sin app"),
  quote("Canal de WhatsApp, no landing.", null),
  gap(180),
  bullet("**Es el canal que ya usan.** WhatsApp aparece como canal principal o de grupo en las cinco personas."),
  bullet("**Una landing propia es bloqueable.** El PESTEL documenta más de 200 dominios bloqueados en Venezuela por Conatel, y recomienda literalmente no poner la captación en un solo dominio propio. Un canal de WhatsApp no se bloquea por dominio."),
  bullet("**El día del lanzamiento se avisa a todos de una vez**, con notificación, sin depender del algoritmo de Instagram ni de que abran un correo."),
  gap(200),
  table(["Peldaño de la persona", "CTA que se usa"], [
    ["TOFU — acaba de ver el primer video", "«Aquí publicamos todos los días. Sigue.» Nada más"],
    ["MOFU — ya guardó algo", "«En octubre abre Kenia. En el canal aviso el día exacto»"],
    ["BOFU — ya está esperando", "«El canal ya tiene 300 personas. Los que están ahí entran primero»"],
  ], [3300, 6348]),
  gap(200),
  panel("Lo que el CTA nunca dice en fase 1", [
    "«Descárgala» —no existe—, «prueba gratis» —no está confirmado— y «suscríbete» —no hay dónde—.",
    "**Prometer algo que no se puede cumplir en 30 días rompe el valor 4 de la marca antes de tener un solo usuario.**",
  ], NARANJA),
  brk(),
);

/* ---------- 06 ---------- */
add(
  h1("Los seis objetivos de contenido"),
  p("La lluvia de ideas traía cinco: identidad, atracción, educación, conversión y fidelización. Los cinco se mantienen y entra uno más, que en esta marca no es un lujo."),
  gap(120),
  table(["Objetivo", "Peso", "Qué persigue", "Cómo se mide"], [
    ["Atracción", "25 %", "Llegar a gente que no conoce Kenia", "Alcance de no seguidores · seguidores nuevos"],
    ["**Legitimación** (nuevo)", "20 %", "Que quien corre y no se llama corredor se reconozca dentro", "Comentarios de identificación · compartidos en privado"],
    ["Educación", "20 %", "Autoridad: que Kenia sepa de lo que habla", "Guardados · tiempo de visualización"],
    ["Identidad de marca", "15 %", "Que se entienda qué es Kenia y por qué se llama así", "Respuestas correctas en la caja de preguntas"],
    ["Conversión", "15 %", "Lista de espera, descarga, suscripción", "Contactos en el canal · descargas · suscriptores"],
    ["Fidelización", "5 %", "Que el que entró no se vaya y el que falló vuelva", "Retención a 30 días · reingresos"],
  ], [2100, 700, 3424, 3424]),
  gap(200),
  panel("Por qué añadimos legitimación como objetivo propio", [
    "**Se mide distinto.** Atracción se mide en alcance; legitimación se mide en si la gente se escribe entre ella para mandarse el video. Es una señal cualitativa que hay que ir a buscar.",
    "**Es el único hallazgo verificado que ningún competidor está usando.** Toda la categoría le habla a runners.",
    "**Y si no es objetivo, no se produce.** Es contenido que no vende nada de forma visible, y en cualquier revisión de resultados es el primero que se cae. Ponerlo como objetivo lo protege.",
  ]),
  gap(200),
  panel("La trampa que hay que evitar con este objetivo", [
    "Legitimar no es dar un título. Kenia **no** dice «tú también eres corredor»: eso conserva el título y solo lo reparte.",
    "Kenia dice **«no hace falta serlo para entrenar»**. Le quita el requisito.",
  ], NARANJA),
  brk(),
);

/* ---------- 07 ---------- */
add(
  h1("Los cinco pilares"),
  p("No salen de una lluvia de ideas: salen de cruzar los dos enemigos, los tres mensajes de la pirámide y el territorio que nadie puede copiar desde fuera."),
  gap(120),
  table(["Pilar", "Peso", "Objetivo", "Subtemas principales"], [
    ["1 · Ya estás dentro", "25 %", "Legitimación",
     "«¿Tú te consideras corredor?» · las seis frases textuales · el requisito que siempre está un paso más adelante · el primer día de vuelta · quién decide tu semana"],
    ["2 · Decide contigo", "25 %", "Conversión",
     "La pregunta de la mañana · el día que cambió · el aviso de carga · el plan del pana frente al que se recalcula · tu semana entera en una pantalla"],
    ["3 · El porqué", "20 %", "Educación",
     "Los tipos de sesión · la carga combinada · por qué en Caracas tus ritmos no cuadran · el método keniano · zapatillas por tipo de sesión · lo que cuesta un entrenador"],
    ["4 · Aquí se corre así", "20 %", "Atracción",
     "La hora azul · rutas reconocibles · el clima se entrena · el calor de aquí · las carreras de aquí · el running de Caracas está creciendo"],
    ["5 · La gente de Kenia", "10 %", "Prueba social",
     "Testimonios del box · el regreso · perfiles de clubes · las ligas y el ranking · Ninro respondiendo en público"],
  ], [1900, 700, 1500, 5548]),
  gap(200),
  panel("El pilar «Lifestyle» de la lluvia de ideas no entra", [
    "**Es la voz del héroe.** Lifestyle en fitness significa aspiración: cómo se ve, qué ropa, qué desayuna. Kenia no es el héroe, es el mentor, y la prueba rápida de la plataforma —¿esta foto la firmaría Nike sin dudar?— la falla casi todo el contenido de lifestyle.",
    "**Choca con la política de cuerpo**, porque acaba en cuerpos, comidas y rutinas de gente que se ve bien.",
    "**Y no prueba ningún mensaje.** Llena calendario sin construir argumento, y con 19 piezas al mes y dos personas no sobra espacio para eso.",
    "Lo que sí se salva —el mundo del corredor de Caracas— entra por el pilar 4, que lo trata como **lugar** y no como estilo de vida. La diferencia es si la pieza dice «así es aquí» o «así deberías vivir».",
  ], NARANJA),

  h2("El guardarraíl del pilar 4, que es el de más riesgo"),
  p("La prueba obligatoria de cada pieza: **¿celebra al que salió, o muestra a alguien preparado?** Si celebra, está juzgando al que no salió y no sale."),
  table(["Sí se puede decir", "No se puede decir"], [
    ["«Esta sesión es bajo lluvia a propósito.»", "«Llueve. No es excusa.»"],
    ["«Tu carrera puede ser así. Vamos a que no sea la primera vez.»", "«Los que ganan salen igual.»"],
    ["«Con este calor bajamos el ritmo. No estás perdiendo forma.»", "«Si llovió y no fuiste, fallaste.»"],
  ], [4824, 4824]),

  h2("La prueba de las cuatro preguntas"),
  p("Antes de publicar cualquier pieza. **Una sola respuesta mala y no sale.**"),
  bullet("¿Le habla al héroe, o habla como el héroe? Si la pieza arenga, es voz de héroe."),
  bullet("¿Esta imagen la firmaría Nike sin dudar? Si sí, probablemente está en el arquetipo equivocado."),
  bullet("¿Juzga a alguien? Al que no salió, al que va lento, al que falló."),
  bullet("¿Promete algo que el producto no sostiene? Supervisión humana, «no te vas a lesionar», una prueba gratuita sin confirmar."),
  brk(),
);

/* ---------- 08 ---------- */
add(
  h1("Formatos y matriz de producción"),
  p("Seis formatos serializados, cada uno con nombre propio, portada fija y una frase de apertura que se repite. Con dos personas y 1.000 dólares el primer mes, **es la única forma de sostener 19 piezas**: se decide una vez la estética y cada pieza después cuesta la mitad."),
  gap(120),
  table(["Formato", "Qué es", "Duración", "Pilar"], [
    ["Pantalla", "Grabación de la app, limpia, con texto encima", "15 a 35 s", "2"],
    ["Ninro", "Su cara. Dos registros: set fijo y selfie", "30 a 75 s", "3"],
    ["Calle", "POV, b-roll y sesión de entrenamiento", "12 a 25 s", "4"],
    ["La gente", "Personas reales contestando una pregunta fija", "20 a 50 s", "1 y 5"],
    ["Carrusel", "6 a 10 láminas. El formato de guardado", "—", "3"],
    ["Motion", "El dato como material visual", "10 a 45 s", "2"],
  ], [1600, 5148, 1400, 1500]),

  h2("La matriz de producción — la tabla que decide si el mes se puede ejecutar"),
  table(["Formato", "Piezas", "Cuáles", "¿Ninro?", "¿Salir a grabar?"], [
    ["Carrusel", "6", "C1 a C6", "No", "No"],
    ["Ninro", "5", "V2 V4 V6 V7 V10", "**Sí**", "Sí"],
    ["Pantalla", "3", "V5 V8 V11", "No", "No"],
    ["La gente", "3", "V1 V9 V12", "Parcial", "Sí"],
    ["Calle", "1 más b-roll", "V3", "No", "Sí"],
    ["Motion", "1 más 2 versiones", "V13", "No", "No"],
  ], [1600, 1400, 2648, 2000, 2000]),
  gap(200),
  panel("Lo que esta tabla significa", [
    "**Diez de las diecinueve piezas no dependen de nadie más que de nosotros**: los seis carruseles, las tres de pantalla y el motion. Ese es el colchón. Si el 13 sale mal, el mes se sostiene igual, más pobre pero en pie.",
    "**Lo que no se recupera es la cara de Ninro y los testimonios.** Cinco piezas llevan su formato, y con V1 y V13 son siete de los trece videos del mes los que muestran su cara.",
    "Y una lectura menos obvia: **Calle solo tiene una pieza propia, y no es un descuido.** Una salida de una mañana produce el b-roll de todo el mes. El formato rinde más como cantera que como pieza.",
  ]),

  h2("Mecánica de publicación"),
  table(["Elemento", "Cómo"], [
    ["Reels", "9:16, 1080 × 1920. TOFU de 12 a 25 s, MOFU educativo de 45 a 75 s. **Nada entre 25 y 45: es tierra de nadie**"],
    ["El gancho", "**Los primeros 2 segundos deciden.** Texto en pantalla desde el frame 1, sin intro de logo"],
    ["Texto", "Siempre. Se ve sin sonido. Subtítulos quemados, no automáticos"],
    ["Zona segura", "Nada importante en el 15 % de arriba ni el 25 % de abajo: se lo come la interfaz"],
    ["Carruseles", "4:5, 1080 × 1350. De 6 a 10 láminas. **La lámina 1 tiene que funcionar como si fuera la única**"],
    ["Audio", "Sonido real por defecto. Audio en tendencia solo en piezas de comunidad, nunca en producto"],
    ["Horarios", "6:00 a 7:30 y 19:30 a 21:00. Domingo de 8 a 10. Son hipótesis: se corrigen con datos propios a las tres semanas"],
    ["Reutilización", "**La regla del 1:4.** Cada grabación produce el reel, un corte para stories, una cita para carrusel y un fragmento para el video de lanzamiento"],
  ], [1900, 7748]),
  brk(),
);

/* ---------- 09 ---------- */
add(
  h1("Plan de canales"),
  quote("Un canal principal, un canal de conversión, y el resto es espejo.", null),
  gap(180),
  p("Con dos personas no se sostienen cinco plataformas nativas. Lo que sí se sostiene es producir para una y adaptar, con un canal —WhatsApp— que no es de contenido sino de conversión."),
  gap(120),
  table(["Canal", "Rol", "Qué va", "Qué no va"], [
    ["Instagram", "Principal", "Reels, carruseles, stories diarias, colaboraciones", "Contenido largo"],
    ["WhatsApp", "**Conversión**", "Canal de difusión, que es la lista. Mensajes de Ninro a su grupo del box", "Contenido diario. **Techo de dos mensajes por semana**"],
    ["TikTok", "Espejo", "Los reels sin marca de agua, con 24 h de retraso", "Carruseles. Piezas de precio"],
    ["Strava", "Presencia", "Club de Kenia y rutas de las salidas de grabación", "**Strava no se usa para vender**"],
    ["YouTube", "Aplazado", "Solo Shorts espejo. El largo entra en noviembre", "Video largo en fase 1"],
    ["Web", "Respaldo", "Página simple con el canal y la ficha de tienda", "**Nunca ser el único punto de captación**"],
    ["Facebook", "No", "Solo como ubicación de pauta", "Contenido orgánico"],
  ], [1400, 1300, 3474, 3474]),
  gap(200),
  panel("Aviso de honestidad sobre esta tabla", [
    "La ficha de buyer personas marca los medios de **Daniel, Patricia y Gabriel como hipótesis sin verificar**, y los de Rebeca y Andrés como verificados solo en parte. Lo único verificado es el Instagram de Rebeca y el grupo de WhatsApp del box de Andrés.",
    "Y hay una excepción escrita que conviene no olvidar: **«Gabriel usa poco Instagram para informarse»** — y es la persona que compra Elite todo el año.",
    "**Este plan de canales es la mejor apuesta disponible, no un hecho.** Es lo que la caja de preguntas de los martes tiene que corregir.",
  ], NARANJA),

  h2("Lo que decidimos no hacer, a propósito"),
  table(["No hacemos", "Por qué"], [
    ["Blog y SEO", "Es la jugada de RunMotion y funciona, pero tarda de seis a nueve meses. El lanzamiento es en cinco semanas. Se reconsidera en enero"],
    ["Podcast propio", "Coste de producción alto y audiencia lenta. El fake podcast da el formato sin el compromiso"],
    ["YouTube largo en fase 1", "Daniel y Gabriel lo consumen, pero son dos de cinco personas y el coste por pieza es cinco veces mayor"],
    ["Pauta antes del 22 de septiembre", "No hay creativos probados. Poner 300 dólares sobre contenido sin datos de retención orgánica es quemarlos"],
  ], [2600, 7048]),

  h2("Stories — el esqueleto semanal"),
  p("Tres a cinco al día, también los días sin pieza de feed. **Es donde vive la lista de espera** y es la mitad del alcance real."),
  table(["Día", "Story fija", "Para qué"], [
    ["Lunes", "Tu semana en Kenia: qué se publica", "Ritmo de cuenta"],
    ["Martes", "**Caja de preguntas**", "Alimenta contenido **y cierra el hueco de qué consumen**"],
    ["Miércoles", "Detrás de cámara o pantalla suelta", "Cercanía"],
    ["Jueves", "**Ninro responde** comentarios reales", "La jugada de Runna: el fundador dando la cara"],
    ["Viernes", "Recordatorio del canal, con el número actual", "Conversión"],
    ["Sábado y domingo", "Ruta, clima, comunidad y carreras", "Utilidad y pertenencia"],
  ], [1800, 4424, 3424]),
  brk(),
);

/* ---------- 10 ---------- */
add(
  h1("El calendario del mes"),
  p("Del lunes 14 de septiembre al domingo 18 de octubre. **Está construido en semanas relativas**: si la fecha de lanzamiento se mueve, la semana 5 se desplaza entera y las cuatro anteriores no se tocan."),
  gap(120),
  table(["Semana", "Fechas", "Tema", "Piezas"], [
    ["S0", "dom 13 sep", "Rodaje", "La única ventana con Ninro"],
    ["S1", "14 a 20 sep", "Ya estás dentro", "V1 · V2 · V3 · C1"],
    ["S2", "21 a 27 sep", "El porqué", "V4 · V5 · V6 · C2"],
    ["S3", "28 sep a 4 oct", "Caracas Rock", "V7 · C3 · V8"],
    ["S4", "5 a 11 oct", "Cuenta atrás", "V9 · V10 · V11 · V12"],
    ["S5", "12 a 18 oct", "Lanzamiento", "C4 · V13 · C5 · C6"],
  ], [1000, 1900, 2900, 3848]),
);

add(semana("S1", "Ya estás dentro", "14 a 20 de septiembre",
  "Que quien nunca ha oído hablar de Kenia se reconozca en el primer video. Ninro está en Buenos Aires del 15 al 20: cero producción nueva, todo sale del 13.",
  [
    ["V1", "mar 15", "**«¿Tú te consideras corredor?»** — el montaje de negativas. Es la pieza de apertura de la marca, y cierra con Ninro: «Once de doce me dijeron que no. Y todos corren»", "TOFU"],
    ["V2", "jue 17", "**«¿Por qué Kenia?»** — Ninro en selfie: «todo el mundo me pregunta si la app es de citas»", "TOFU"],
    ["V3", "sáb 19", "**Hora azul** — POV a las 5:30, sin voz. «Nadie deja de correr por falta de voluntad»", "TOFU"],
    ["C1", "dom 20", "**Las seis frases** con las que la gente de Caracas dice que no es corredora, textuales. **Primera vez que se pide el canal**", "MOFU"],
  ]));

add(semana("S2", "El porqué", "21 a 27 de septiembre",
  "Demostrar que Kenia sabe de entrenamiento, y enseñar el producto por primera vez. Es la semana que construye la autoridad que sostiene el precio de octubre. El 22 arranca la pauta.",
  [
    ["V4", "mar 22", "**La carga combinada** — fake podcast. Tres de doce se lesionaron por correr y levantar subiendo carga a la vez, y ninguna app de running pregunta por el gimnasio", "MOFU"],
    ["V5", "jue 24", "**La pregunta de la mañana** — pantalla. **La pieza más importante del mes después del video de lanzamiento**: es el único momento en que el diferencial se ve en vez de contarse", "MOFU"],
    ["V6", "sáb 26", "**Los tipos de sesión** — tier list de cosas, nunca de personas", "TOFU"],
    ["C2", "dom 27", "**Por qué tus ritmos no cuadran** — 900 m de altitud, calor y humedad. **Ninguna app global lo explica aquí**", "MOFU"],
  ]));

add(semana("S3", "Caracas Rock", "28 de septiembre a 4 de octubre",
  "Estar donde está el público. Cuatro de los doce entrevistados corren esa carrera. El domingo 4 no hay pieza de feed: es día de operación, no de publicación.",
  [
    ["V7", "mar 29", "**La semana antes** — «faltan cinco días para tu carrera; lo que hagas ahora ya no te va a hacer más rápido»", "MOFU"],
    ["C3", "jue 1 oct", "**Guía de Caracas Rock** — la semana previa, el día y el lunes de después. La jugada de RunMotion a escala local, y hoy sin usar por nadie aquí", "MOFU"],
    ["V8", "sáb 3", "**Una semana, ordenada** — pantalla: el plan en scroll y sin voz. Lo que Runna hace mejor que nadie", "MOFU"],
    ["—", "dom 4", "**La carrera.** Stories en vivo, QR al canal y la pregunta en meta. El sábado, en el retiro de kits, se capta más que el domingo", "—"],
  ], true));

add(semana("S4", "Cuenta atrás", "5 a 11 de octubre",
  "Convertir todo lo acumulado en lista de espera, y enganchar a los que acaban de correr al siguiente objetivo antes de que se desconecten.",
  [
    ["V9", "lun 5", "**Caracas Rock** — cobertura. «¿Y ahora qué?»", "TOFU"],
    ["V10", "mié 7", "**El lunes después de la carrera** — ataca el churn más grave de los cinco: «ya corrí la carrera». Nombra por primera vez la media de noviembre", "MOFU"],
    ["V11", "vie 9", "**Este plan no sabe que hoy llovió** — el enemigo A, sin nombrar a ningún competidor", "BOFU"],
    ["V12", "dom 11", "**Tu plan te esperó** — el regreso. El beneficio emocional central de la marca", "MOFU"],
  ]));

add(semana("S5", "Lanzamiento", "12 a 18 de octubre",
  "La única semana del trimestre con permiso para vender de frente.",
  [
    ["C4", "lun 12", "**El jueves abre** — cuenta atrás. Los del canal entran primero", "BOFU"],
    ["V13", "jue 15", "**El video de lanzamiento** — 45 segundos. Ver la sección 12", "BOFU"],
    ["C5", "vie 16", "**Qué hace Kenia que tu plan no hace**", "BOFU"],
    ["C6", "dom 18", "**Lo que cuesta** — precios. **No se publica hasta que Kurt confirme el cobro y la prueba gratuita**", "BOFU"],
  ], true));

add(
  gap(300),
  panel("Cómo leer el reparto de este calendario", [
    "**El BOFU de fase 1 es una sola pieza, y está bien.** Cuando no hay nada que descargar, el BOFU no vive en piezas dedicadas: vive en el CTA del pie de cada pieza de MOFU. Nueve de las quince piezas de S1 a S4 lo llevan.",
    "**El mes completo cae en 37/42/21**, que es prácticamente el 40/40/20 que propusimos. El reparto sí se cumple, pero al final del mes y no cada semana.",
  ]),
  brk(),
);

/* ---------- 11 ---------- */
add(
  h1("El rodaje del 13 de septiembre"),
  quote("Siete de los trece videos del mes llevan su cara, y todos salen de este día. Lo que no se grabe el 13 no existe.", null),
  gap(180),
  p("Y una regla que manda sobre todo el rodaje: **todo lo que se grabe tiene que servir a las dos fases** —el mes sin app y el lanzamiento—. Por eso Ninro nunca dice «descárgala» ni «ya está disponible»: eso lo dice un texto en pantalla, que se cambia en montaje."),

  h2("El orden del día"),
  table(["Bloque", "Hora", "Qué", "Prioridad"], [
    ["E-1", "5:30 – 6:00", "**B-roll de hora azul.** Faroles, ciudad vacía, primeras pisadas", "Alta"],
    ["C", "6:00 – 7:30", "**Testimonios del box.** La gente llega primero, cuando aún tiene energía", "**Máxima**"],
    ["D", "7:30 – 8:00", "Micro-entrevistas del enemigo B, a la misma gente", "**Máxima**"],
    ["E-2", "8:00 – 9:00", "El resto del b-roll: sesión de entrenamiento, calle, Caracas de día", "Alta"],
    ["A", "9:00 – 10:30", "**Ninro, set fijo.** Las ocho tomas de autoridad", "**Máxima**"],
    ["B", "10:30 – 11:00", "Ninro, selfie y calle", "Alta"],
    ["F", "11:00 – 11:30", "**Las diez reseñas de tienda**", "**Máxima**"],
  ], [1000, 1700, 5448, 1500]),
  gap(200),
  panel("Dos cosas de este orden que no son arbitrarias", [
    "**La gente se graba antes que Ninro** porque la gente se va y Ninro no. Si el día se cae, lo último que se pierde tiene que ser lo que se puede repetir.",
    "**El bloque E va partido.** La hora azul es a las 5:30, no a las ocho: es el territorio visual 2 y no se puede fingir con luz de media mañana. Media hora de madrugada ahorra una salida entera.",
  ]),

  h2("Bloque C — las seis preguntas de los testimonios"),
  p("Ocho a diez personas, de cuatro a seis minutos cada una. **Es el bloque más importante del día.**"),
  table(["#", "Pregunta", "Qué busca"], [
    ["C1", "¿Cuánto llevas corriendo y cuántas veces por semana?", "Rompe el hielo y da contexto en pantalla"],
    ["C2", "¿De dónde sale el plan que sigues hoy?", "«De un amigo», «me lo armé yo», «con Gemini». **El rival real**"],
    ["C3", "¿Cuál fue la última vez que no pudiste entrenar, y qué pasó?", "Lluvia, fiebre, un hijo, quedarse dormido. **Nunca se dice «excusa»**"],
    ["C4", "¿Y qué pasó después? ¿Cuánto te costó retomar?", "**La pregunta más importante del día.** Es el beneficio emocional central"],
    ["C5", "¿Qué es lo que más te cuesta sostener del entrenamiento?", "El problema en sus palabras"],
    ["C6", "Si mañana tuvieras a alguien que te dijera qué toca hoy, ¿qué le preguntarías?", "Convierte al entrevistado en usuario sin mentir, y sirve en las dos fases"],
  ], [500, 4574, 4574]),
  gap(200),
  panel("Cómo se pregunta", [
    "**Nunca se pregunta «¿por qué fallaste?».** El sujeto es el día, no la persona: «¿qué pasó ese día?».",
    "**Se deja el silencio.** La segunda frase siempre es mejor que la primera.",
    "**Se pide que la respuesta incluya la pregunta**, para que el corte funcione sin que se oiga a quien pregunta.",
    "Y no se les pregunta por su cuerpo, por su peso ni por cómo se veían antes. Política de cuerpo.",
  ]),

  h2("Bloque A — las ocho tomas de Ninro"),
  table(["#", "Qué se le pregunta", "Alimenta"], [
    ["A1", "¿Qué es Kenia, en treinta segundos, para alguien que nunca la ha oído?", "V13 · presentación"],
    ["A2", "¿Por qué se llama Kenia?", "V2"],
    ["A3", "¿Qué le pasa a un plan cuando tu semana se rompe?", "V13 · V11"],
    ["A4", "Mucha gente corre y hace pesas la misma semana. ¿Dónde está el riesgo?", "V4"],
    ["A5", "¿Cuántos tipos de sesión existen de verdad, y para qué sirve cada una?", "V6"],
    ["A6", "La semana antes de una carrera, ¿qué se hace y qué no? ¿Y el lunes de después?", "V7 · V10"],
    ["A7", "En Caracas hace 30 grados y estamos a 900 metros. ¿Eso cómo afecta?", "V6 · C2"],
    ["A8", "¿A quién NO le sirve Kenia?", "V13 · stories"],
  ], [500, 6148, 3000]),
  gap(180),
  p("Y al final del bloque, **las cuatro frases de campaña a cámara**, plano cerrado, tres versiones cada una: seria, más suave y una con media sonrisa. En montaje se elige."),
  gap(200),
  panel("La lista de comprobación de fin de día", [
    "Antes de guardar el equipo, alguien la lee en voz alta: los 8–10 testimonios con C4 incluida · las 15–20 micro-entrevistas · las ocho tomas de Ninro · las cuatro frases en tres versiones · los diez planos de b-roll, con los cordones y con la luz de las 5:30 · el compromiso de reseñas hablado con todos · el mensaje de cesión de imagen mandado a cada persona.",
    "**Y copia de seguridad de las tarjetas en dos sitios distintos, antes de dormir.** No es burocracia: una tarjeta perdida el 13 de septiembre no se recupera con más esfuerzo. Se recupera cancelando el lanzamiento.",
  ], NARANJA),
  brk(),
);

/* ---------- 12 ---------- */
add(
  h1("El plan de lanzamiento"),
  panel("El problema, dicho sin adornos", [
    "**La app sale a tiendas sin una sola reseña, sin testimonios y sin casos**, prometiendo que se adapta y que previene lesiones. Y sale en el mes más lluvioso del año, contra alternativas gratuitas que ya están dentro del reloj que el usuario compró.",
  ], NARANJA),
  gap(220),
  table(["Decisión", "Por qué"], [
    ["**La prueba social se fabrica un mes antes**", "Los diez testimonios y las diez reseñas se graban y se piden el 13 de septiembre, no el 15 de octubre"],
    ["**El lanzamiento no se ancla a sí mismo: se ancla a noviembre**", "La Media Maratón de Hipereventos da fecha, urgencia y un motivo para suscribirse en octubre en vez de «algún día»"],
    ["**El día uno no empieza de cero: empieza con la lista**", "400 personas en el canal que reciben el aviso a la vez, sin depender del algoritmo"],
  ], [3300, 6348]),

  h2("Los tres movimientos"),
  table(["Movimiento", "Cuándo", "Qué hace"], [
    ["A · Captación presencial", "dom 4 oct", "Llenar la lista donde está el público, once días antes"],
    ["B · La cuenta atrás", "5 a 14 oct", "Convertir audiencia en lista, y lista en expectativa"],
    ["C · La apertura", "15 a 18 oct", "Descarga, reseñas el día uno y el ancla de noviembre"],
  ], [2600, 1700, 5348]),

  h2("Movimiento A — Caracas Rock"),
  p("La carrera cae once días antes del lanzamiento y los corredores se la pierden como usuarios. No importa: **no vamos a venderles nada. Vamos a por sus contactos y por su cara.**"),
  table(["Las dos preguntas de la meta", "Para qué"], [
    ["«¿Y ahora qué? ¿Cuál es tu próxima carrera?»", "Alimenta V9 y V10, y **detecta a los que van a la media de noviembre**, que son los suscriptores de octubre"],
    ["«¿Tú te consideras corredor?»", "La misma pregunta en la meta, con el dorsal puesto, y la gente sigue diciendo que no. **Ese contraste es el mejor material de marca que se puede conseguir**"],
  ], [3800, 5848]),
  gap(200),
  panel("Lo que no se hace en Caracas Rock", [
    "No se vende, no se dice el precio, no se pide descargar nada. La app no existe todavía. **Se pide permiso para avisar.**",
    "Y sin un mensaje de cesión de imagen a cada persona grabada, con su nombre, **no se publica ninguna cara**.",
  ], NARANJA),

  h2("Movimiento C — el día 15, hora por hora"),
  table(["Hora", "Qué", "Dónde"], [
    ["6:00", "**Mensaje al canal.** Es lo primero que pasa", "WhatsApp"],
    ["6:30", "El video de lanzamiento", "Reels y TikTok"],
    ["7:00", "Story con el link a tienda, fijada en destacados", "Stories"],
    ["8:00", "Ninro publica en su cuenta personal y en el grupo del box", "Personal"],
    ["A lo largo del día", "**Las diez reseñas se suben a tienda**, pedidas un mes antes", "Tienda"],
    ["12:00", "Mensaje uno a uno a los 12 entrevistados, con nombre", "WhatsApp"],
    ["Todo el día", "**Ninro contesta cada comentario, uno por uno**", "Instagram"],
  ], [2000, 5648, 2000]),
  gap(180),
  p("Lo último no es un detalle. El fundador de Runna contestando dudas en público generó más confianza que cualquier campaña. **Es lo más barato y lo más creíble que tenemos, y el día 15 es el día que más se nota.**"),

  h2("El video de lanzamiento — 45 segundos"),
  panel("Aviso", [
    "**El reel de referencia que pasó Kurt no se pudo revisar**: Instagram bloquea el acceso desde el entorno donde se trabajó. Esta estructura está construida desde cero.",
    "Con que alguien lo vea y conteste tres cosas —qué formato es, cómo arranca los primeros dos segundos, y qué gustó exactamente de él— se calibra en una hora.",
  ], NARANJA),
  gap(220),
  table(["Tiempo", "Qué se ve", "Qué se oye o se lee"], [
    ["0:00–0:03", "Negro, tipografía cian", "**«Nadie deja de correr por falta de voluntad.»**"],
    ["0:03–0:10", "Lluvia sobre asfalto · despertador a las 4:58 · una mano en el gemelo", "Sin voz. «Deja de correr porque sigue un plan que no sabe nada de su vida»"],
    ["0:10–0:17", "Ninro, set fijo, a cámara", "«Todos los planes te dicen qué hacer. Ninguno te pregunta cómo amaneciste»"],
    ["0:17–0:30", "**Pantalla de la app.** La pregunta, la respuesta, la sesión cambiando delante del ojo", "«Te pregunta cómo dormiste, cómo tienes las piernas y cuánto tiempo tienes. Y con eso decide tu día. Y te dice por qué»"],
    ["0:30–0:38", "Cuatro caras del box, cortes de un segundo", "Cuatro negativas del bloque D, y Ninro en off: **«Si corres, ya estás dentro»**"],
    ["0:38–0:45", "Logo, descriptor, tagline y badges de tienda", "**Kenia · Entrenamiento de running** — El plan que decide contigo"],
  ], [1300, 4174, 4174]),
  gap(200),
  table(["Sí", "No"], [
    ["Sonido real: lluvia, pisadas, interfaz", "Música épica de librería con subida"],
    ["Cara de concentración", "Grito, puño al aire, cámara lenta heroica"],
    ["Caracas reconocible: el Ávila, la Cota Mil", "Amanecer genérico de banco de imágenes"],
    ["Cuerpos diversos", "Solo cuerpos atléticos y delgados"],
    ["Un solo CTA, al final", "«Plan personalizado» de titular"],
  ], [4824, 4824]),
  gap(180),
  p("Del mismo montaje salen tres versiones: **la larga de 45 s** para feed y ficha de tienda, **la corta de 15 s** para pauta —el tramo 0:10 a 0:25, que dice qué es en tres segundos— y **la de 20 s solo de pantalla** para el video de vista previa de las tiendas."),

  h2("Y hasta enero — el mapa, no el calendario"),
  table(["Tramo", "Trabajo del contenido", "Riesgo que atiende"], [
    ["23 nov a 6 dic", "**La cosecha del trimestre:** los primeros casos reales de gente que sostuvo de ocho a diez semanas", "El lanzamiento ya no es noticia"],
    ["7 a 20 dic", "**El contenido del que falla en diciembre.** La promesa central se demuestra en vez de prometerse", "**La retención a 60 días**"],
    ["21 dic a 3 ene", "Frecuencia mínima, coste mínimo. Nada de estrenar formatos ni de quemar pauta", "Gastar en la peor ventana del año"],
    ["4 a 18 ene", "**La segunda campaña de captación**, la más barata del año en intención", "Única ventana del semestre comparable a un lanzamiento"],
  ], [1900, 4374, 3374]),
  gap(200),
  panel("Por qué esto es un mapa y no un calendario de piezas", [
    "Bajar noviembre, diciembre y enero a piezas hoy, sin un solo dato de conversión, sería predecir sin datos: el error número uno documentado de este proyecto.",
    "**Las piezas se escriben el 19 de octubre y el 30 de noviembre, con números en la mano.**",
  ]),
  brk(),
);

/* ---------- 13 ---------- */
add(
  h1("Tagline, frases y comunidad"),
  panel("Todo esto es propuesta", [
    "Necesita el visto bueno de Ninro **antes del viernes 12 de septiembre**, porque el tagline y las cuatro frases se graban el 13.",
  ], NARANJA),
  gap(220),

  h2("Tagline"),
  quote("Kenia · Entrenamiento de running\nEl plan que decide contigo", "Recomendado"),
  gap(180),
  p("El descriptor ya dice de qué categoría es; el tagline tiene que decir qué hace distinto. **«Decide contigo» hace las dos cosas a la vez**: nombra la acción del producto y, con el «contigo», mete al usuario dentro sin pedirle que se llame nada."),
  p("Y aguanta las cuatro pruebas prácticas: bio de Instagram, subtítulo de tienda, cierre de reel y dicho en voz alta por Ninro. Las otras dos candidatas no aguantan las cuatro."),
  gap(120),
  table(["Candidata", "Veredicto"], [
    ["**El plan que decide contigo**", "**Recomendada.** Cinco palabras, dice qué hace, y «decide contigo» es literalmente el diferencial que no dice ningún competidor"],
    ["Tu semana cambia. Tu plan también", "Muy clara y muy compartible, pero son dos frases. **Funciona mejor como frase de campaña que como lockup**"],
    ["El plan que sabe cómo fue tu semana", "Cálida y concreta, pero «saber» es pasivo: describe registro, no decisión. Y el registro es higiene, no argumento"],
  ], [3300, 6348]),
  gap(180),
  p("**«Master your pace» se retira**, como ya estaba decidido: está en inglés, es evocativo y no explica. Puede sobrevivir como línea gráfica secundaria si al diseñador le sirve."),

  h2("Las cuatro frases de campaña"),
  p("Reemplazan a las cuatro retiradas. **Cada una cubre un pilar y un peldaño del embudo** — no son cuatro maneras de decir lo mismo."),
  table(["Frase", "Pilar", "Reemplaza a"], [
    ["**Tu semana cambia. Tu plan también.**", "Decide contigo", "«Donde termina tu excusa, empieza tu carrera»"],
    ["**Si corres, ya estás dentro.**", "Ya estás dentro", "«Tu único rival te mira en el espejo»"],
    ["**Tu plan te esperó.**", "El regreso", "«La meta es solo el principio»"],
    ["**El clima se entrena.**", "Aquí se corre así", "«El asfalto no miente»"],
  ], [3300, 2000, 4348]),
  gap(200),
  panel("La condición de la frase 3", [
    "**«Tu plan te esperó» solo se puede publicar si el producto de verdad hace algo razonable el día después de un fallo.**",
    "Si la respuesta de Kurt es que el plan se rompe y hay que empezar de cero, esta frase no sale — y tenemos un problema bastante más grande que una frase de campaña.",
  ], NARANJA),

  h2("El nombre de la comunidad"),
  quote("Gente que corre", "Recomendado"),
  gap(180),
  p("El criterio, que es lo que decide: **tiene que ser algo a lo que te unes, no un título que haya que ganarse**. Eso descarta de entrada «los corredores de Kenia», «runners» y «atletas»."),
  p("Y esta frase es textual de las entrevistas: **tres de doce, al negar que fueran corredores, se describieron exactamente así**. Convertir su forma de excluirse en el nombre del club es toda la estrategia de marca en dos palabras."),
  p("El argumento en una línea, por si alguien pregunta: **no hay que ser corredor para entrar; hay que correr.** Y correr es lo que esas once personas ya hacen."),
  gap(200),
  panel("La alternativa, y por qué no la recomendamos", [
    "«En formación» sale de la persona 5 y es la dirección que la plataforma de marca marcó como más interesante.",
    "**El riesgo es real: conserva la idea de que todavía no llegaste.** Nombra la carencia en vez de disolverla.",
    "Antes de fijar cualquiera de las dos hay que comprobar que el handle está libre y que no colisiona con otro club o marca fitness en Venezuela.",
  ]),
  brk(),
);

/* ---------- 14 ---------- */
add(
  h1("Qué de la lluvia de ideas no va"),
  p("De los doce videos propuestos, **siete van tal cual, cuatro se reencuadran y uno no va**."),
  gap(120),
  table(["Idea original", "Veredicto", "Qué se hace"], [
    ["Motion de presentación", "Va", "Es el video de lanzamiento. Pero el titular no puede ser «plan personalizado»: es lo que Garmin y Nike regalan"],
    ["Ninro hablando de la app", "Va", "Imprescindible. El método es de autor y el autor no aparece por ningún lado"],
    ["Dinámica con las personas", "Se concreta", "«Dinámica» es una intención, no un formato. Pasa a ser «¿Tú te consideras corredor?»"],
    ["Tier list", "Se reencuadra", "**Una tier list ordena y juzga.** Que ordene superficies, condiciones y errores de carga. **Nunca gente ni niveles**"],
    ["Historia de Ninro", "Solo una mitad", "La cruzada sí. Su transformación física no: es tema vetado y la decisión no se reabre"],
    ["Lista de ejercicios de menor a mayor", "Se reencuadra", "Kenia es de running, no de gimnasio. Pasa a «los tipos de sesión y para qué sirve cada una»"],
    ["Fake podcast", "Va", "Con un matiz: formato, no parodia. El tono es serio, el set da el código y el contenido va en serio"],
    ["Sesión de entrenamiento", "Va", "B-roll de todo el trimestre"],
    ["Un POV", "Va", "De lo más barato y lo más diferencial que hay"],
    ["Hablando a cámara en selfie", "Va", "El registro que mejor funciona en redes y el que más humaniza"],
    ["**Comparaciones: estilo de vida antes y ahora**", "**No va**", "Ver el recuadro de abajo"],
    ["«No hagas esto, haz esto»", "Se reencuadra", "«No hagas esto» pone el sujeto en la persona. Pasa a «esto te rompe, y por qué»"],
  ], [3000, 1700, 4948]),
  gap(220),
  panel("La única idea en rojo, y la intención era buena", [
    "Lo que buscaba era mostrar transformación, que es lo que mejor funciona en fitness y lo que Runna hace con sus corredores reales.",
    "**Por qué no se puede: «estilo de vida antes y ahora» es antes/después con otra ropa.** Aunque no se vea el cuerpo, la estructura es la misma y el mensaje que deja también: que había una versión peor de ti. La política de cuerpo de Kenia es una regla dura.",
    "**Lo que sí se puede:** antes y después **de proceso y de marcas**, nunca de físico. «Hace ocho semanas corría 5K sin plan; hoy va por el 10K y sabe por qué toca cada sesión».",
    "**Y lo que funciona igual de bien:** la comparación entre dos planes, no entre dos personas. Es V11, y ataca al enemigo A sin tocar a nadie.",
  ], NARANJA),

  h2("Los pilares propuestos"),
  table(["Pilar propuesto", "Veredicto", "Qué pasa con él"], [
    ["Comunidad", "Va", "Es el pilar 5, La gente de Kenia"],
    ["Reto", "Se reencuadra", "«Reto» tira directo a «te desafío a que salgas», que es «sin excusas» con otra cara. Lo bueno de la idea —las ligas, el ranking, las carreras con fecha— **cabe en los pilares 4 y 5 sin el marco de desafío**"],
    ["Lifestyle", "**No va**", "Es la voz del héroe, choca con la política de cuerpo y no prueba ningún mensaje. Ver la sección 7"],
    ["El cuarto, sin definir", "Resuelto", "Con los pilares 1, 2 y 3, que salen de los enemigos y de los mensajes"],
  ], [2200, 1700, 5748]),

  h2("Los cuatro errores que este proyecto ya está en riesgo de repetir"),
  table(["El error", "Antídoto"], [
    ["Predecir el comportamiento del público sin datos", "Caja de preguntas los martes y revisión con datos propios a las tres semanas. Todo lo que es hipótesis va marcado"],
    ["Vender la parte comoditizada", "«Plan personalizado» y «lleva tus marcas» no pueden ser titular de ninguna pieza"],
    ["Deslizarse a «sin excusas»", "La prueba obligatoria: ¿celebra al que salió o muestra a alguien preparado?"],
    ["Prometer lo que el producto no sostiene", "Las dos ramas del BOFU y el veto sobre el carrusel de precios"],
  ], [3300, 6348]),
  brk(),
);

/* ---------- 15 ---------- */
add(
  h1("Métricas y gobierno"),
  table(["Indicador", "Meta", "Corte", "De dónde sale"], [
    ["**Lista de espera**", "400", "mié 14 oct", "Ver el desglose de abajo"],
    ["Seguidores nuevos", "+700", "dom 18 oct", "Horizonte por cerrar con Ninro"],
    ["**Descargas el día 1**", "150", "jue 15 oct", "37,5 % de la lista. Conservador: la lista es gente que ya dijo que sí"],
    ["Descargas a 4 días", "250", "dom 18 oct", ""],
    ["**Reseñas el día 1**", "10", "jue 15 oct", "Bloque F del rodaje. Se piden 20 para conseguir 10"],
    ["Guardados por carrusel", "≥3 %", "continuo", "Indicador de utilidad. Decide qué se serializa"],
    ["Coste por contacto de pauta", "≤2,50 $", "14 oct", "Unos 220 $ entre 95 contactos atribuidos a pauta"],
  ], [2400, 900, 1500, 4848]),

  h2("De dónde salen los 400"),
  table(["Fuente", "Contactos", "Supuesto"], [
    ["Grupo del box y red de Ninro", "60", "La gente que ya lo conoce. Es lo más seguro de la lista"],
    ["Los 12 entrevistados y sus grupos", "25", "Dieron su tiempo; son los más fáciles de convertir"],
    ["**Caracas Rock**", "**120**", "**El número más frágil: no sabemos cuántos inscritos hay**"],
    ["Contenido orgánico", "100", "Unos 5 contactos por pieza. Conservador"],
    ["Pauta", "95", "Del 22 de septiembre al 14 de octubre, unos 220 $"],
  ], [3000, 1300, 5348]),
  gap(200),
  panel("Dónde está el riesgo", [
    "**120 de 400 dependen de Caracas Rock, y no sabemos cuántos inscritos tiene esa carrera.** Si rinde la mitad, la lista se queda en 340 y el día uno pierde unas 20 descargas.",
    "No es catastrófico, pero conviene que Ninro consiga el dato antes del 25 de septiembre para poder compensar con pauta.",
  ], NARANJA),
  gap(200),
  panel("Y dos avisos sobre el presupuesto y el horizonte", [
    "**El presupuesto de pauta de la fase 1 no son 300 $.** La pauta arranca el 22 de septiembre y la lista se cierra el 14 de octubre: son 22 días, unos **220 $**. El resto del mes se gasta ya en la semana de lanzamiento.",
    "**El horizonte de los 2.500 de audiencia es ambiguo en las fuentes.** La propuesta dice «a 90 días» —unos 833 al mes— y los objetivos SMART van del 8 de septiembre al 8 de enero —unos 615 al mes—. Tomamos el largo y ponemos +700 en el mes 1, pero **conviene que Ninro cierre cuál se va a medir**: entre uno y otro la meta cambia un 35 %.",
  ]),

  h2("Cómo se decide qué matar"),
  p("**Ningún formato se juzga por una pieza.** Cada formato tiene tres publicaciones antes de una decisión: la primera es una prueba, la segunda una lectura, la tercera un veredicto."),
  table(["Decisión", "Cuándo", "Ejemplo"], [
    ["Serializar", "Pasa los umbrales en dos de tres piezas", "Si «por qué tus ritmos no cuadran» funciona, salen el calor, la humedad y el viento"],
    ["Reencuadrar", "Falla el gancho pero el fondo interesa", "Se cambian los primeros dos segundos. **No cuenta como una de las tres**"],
    ["Matar", "Falla dos umbrales en las tres piezas", "Se retira y su peso se reparte entre los que sí funcionan"],
  ], [1800, 3300, 4548]),
  gap(200),
  panel("Los umbrales son provisionales, y a propósito", [
    "Los números de referencia —retención a 3 segundos, guardados, alcance de no seguidores— son de categoría, **no de esta cuenta, que arranca de cero el 15 de septiembre**.",
    "Sirven tres semanas y se tiran: **el 5 de octubre el umbral pasa a ser la mediana de lo que hayamos publicado**, que es el único punto de comparación honesto.",
    "No es una cautela de trámite. El error número uno de este proyecto es predecir sin datos, y una tabla de umbrales inventados es exactamente esa trampa con aspecto de rigor.",
  ], NARANJA),
  gap(200),
  table(["Fecha", "Qué se revisa", "Qué se decide"], [
    ["5 de octubre", "Las primeras 11 piezas", "Recalibrar umbrales · qué formato se refuerza · dónde va la pauta del lanzamiento"],
    ["19 de octubre", "El lanzamiento completo", "El plan de noviembre entero y el reparto de pilares del régimen"],
    ["30 de noviembre", "El trimestre", "Si el 40/40/20 es el correcto y si el precio se revisa con datos reales"],
  ], [1900, 2900, 4848]),
  gap(200),
  panel("La aritmética que conviene tener delante", [
    "900 $ de pauta a 90 días entre 40 suscriptores = **22,5 $ de CAC**, contra una meta comprometida de menos de 25 $. **No hay margen: si la conversión cae del 3 % al 2 %, el CAC se va a 33 $ y se rompe el compromiso.** Con honorarios de agencia cargados, el CAC real por suscriptor es de unos 107 $.",
    "No es un problema de contenido y no se arregla con contenido, pero tiene que estar escrito donde están las metas. De las siete métricas comprometidas, **la única que el contenido mueve directamente es crecimiento de audiencia**.",
  ]),
  brk(),
);

/* ---------- 16 ---------- */
add(
  h1("Lo que falta y de quién depende"),
  p("Veintisiete cosas abiertas, agrupadas por quién las debe. Las del primer bloque bloquean el rodaje del domingo."),
  p("**Ninguna se rellena por nuestra cuenta.** Un hueco marcado como hueco se puede cerrar; un hueco tapado con una suposición se descubre en octubre."),
);

add(pendientes("Bloqueantes de esta semana", "antes del domingo 13", [
  ["Confirmar los 8–10 del box, con nombre y hora. Sin gente no hay testimonios, y sin testimonios la app lanza sin una sola reseña", "Ninro", "jue 11"],
  ["Describir el reel de referencia que pasó Kurt: qué formato es, cómo arrancan los dos primeros segundos y qué gustó de él", "Kurt · Liam", "jue 11"],
  ["Aprobar el tagline «El plan que decide contigo». Se graba a cámara el 13", "Ninro", "vie 12"],
  ["Aprobar las cuatro frases de campaña. Se graban sueltas, tres versiones cada una", "Ninro", "vie 12"],
  ["Aprobar «Gente que corre» como nombre de comunidad, y comprobar que el handle está libre", "Ninro", "vie 12"],
  ["Ropa sin logos de terceros visibles. El manual actual ya tiene adidas y Nike en portada", "Ninro", "vie 12"],
  ["Guion cerrado, shot list impresa, tarjetas formateadas y baterías", "Nosotros", "vie 12"],
]));

add(pendientes("Los huecos de producto", "Kurt", [
  ["¿Cómo paga exactamente un venezolano? Es el riesgo nº 1 del PESTEL: si no se resuelve, se cae el BOFU entero", "Kurt", "30 sep"],
  ["¿Hay prueba gratuita y de cuántos días? Decide el CTA de todo el contenido", "Kurt", "30 sep"],
  ["¿Qué hace el plan el día después de un fallo? Sin esto no se puede publicar «Tu plan te esperó»", "Kurt", "25 sep"],
  ["Fecha exacta del lanzamiento. El calendario asume jueves 15 de octubre", "Kurt", "25 sep"],
  ["Confirmar que las pantallas que grabamos salen en la versión de octubre", "Kurt", "18 sep"],
  ["¿Se integra con reloj o con Strava? Once de doce ya miden con reloj: es la objeción de Daniel", "Kurt", "30 sep"],
  ["Qué de los tres planes entra en la versión de octubre", "Kurt", "30 sep"],
]));

add(pendientes("Datos y decisiones", "Ninro", [
  ["Inscritos reales de Caracas Rock. 120 de los 400 contactos dependen de esa carrera", "Ninro", "25 sep"],
  ["Qué cuentas de running venezolanas tienen audiencia, con nombres exactos", "Ninro", "25 sep"],
  ["Competencia local: entrenadores con marca personal en Instagram", "Ninro", "30 sep"],
  ["Cerrar el horizonte de los 2.500 de audiencia: 90 días o del 8 de septiembre al 8 de enero", "Ninro", "30 sep"],
  ["Registro de marca, dominio y handles. «Kenia» es un topónimo", "Ninro", "—"],
]));

add(pendientes("Producción", "Organic Club", [
  ["Crear el canal de WhatsApp y el club de Strava. El canal es toda la captación del mes", "Nosotros", "lun 14 sep"],
  ["Sistema de portadas por serie, para que la cuadrícula se lea como un índice", "Nosotros", "18 sep"],
  ["Grabar todas las pantallas del mes de una vez: V5, V8 y V11", "Nosotros", "lun 21 sep"],
  ["Moodboard con los siete territorios visuales. El criterio ya está escrito", "Nosotros", "25 sep"],
  ["Recalibrar los umbrales con datos propios", "Nosotros", "5 oct"],
]));

add(pendientes("Lo que sigue abierto en el análisis de marketing", "afecta a esta estrategia", [
  ["Tendencias, FODA, pricing formal, benchmark y socios estratégicos. El PESTEL está cerrado", "Nosotros", "—"],
  ["Decidir si Rebeca entra en pauta: el PESTEL vio 47 % de participación femenina en el Maratón CAF", "Ninro", "19 oct"],
  ["Decidir si el 42 % de corredores nuevos entra en el target. Enero es cuando ese público aparece", "Ninro", "diciembre"],
]));

add(
  gap(300),
  panel("Los tres que de verdad bloquean", [
    "**El reel de Kurt**, porque el video de lanzamiento se monta la semana del 5 de octubre y sin la referencia va a ciegas.",
    "**El cobro**, porque si no se resuelve no hay nada que vender y el contenido pasa a construir audiencia para un lanzamiento posterior. No depende de nadie de este equipo.",
    "**Qué hace el plan tras un fallo**, porque es la promesa central de la marca y hoy no sabemos si el producto la cumple.",
  ], NARANJA),
  brk(),
);

/* ---------- 17 ---------- */
add(
  h1("Cómo se auditó"),
  p("La versión 1 de esta estrategia se revisó entera contra las fuentes antes de darla por buena. **Salieron 17 errores**, todos corregidos. Se deja escrito por la misma razón por la que la marca escribe sus propias correcciones: para no repetirlas."),
  gap(120),
  table(["Qué se comprobó", "Resultado"], [
    ["Fechas y días de la semana del calendario", "**24 de 24 correctas**"],
    ["Referencias cruzadas entre documentos", "**26 de 26 válidas**"],
    ["Cifras de las 12 entrevistas y del PESTEL", "Todas coinciden con la fuente"],
    ["Precios y planes", "Coinciden"],
    ["Palabras y temas vetados", "Solo aparecen marcados como vetados"],
    ["Conteo de piezas y porcentajes", "**3 errores**"],
    ["Aritmética de pauta y coste por contacto", "**2 errores**"],
    ["Alcance contra lo que pedía el traspaso", "**3 huecos**"],
  ], [4824, 4824]),

  h2("Los cuatro que más habrían costado"),
  p("No son los de cifra. Los de cifra se ven leyendo; **estos se leen bien, suenan seguros y solo aparecen cuando se cuenta**."),
  table(["El error", "Por qué importaba"], [
    ["**Pantalla estaba en 2 piezas de 19**", "Cuando el análisis de competencia dice literal que las capturas de la app son «lo primero que hay que producir» — y es el formato que no depende de nadie. Se corrigió pasando V8 a pantalla"],
    ["**La hora azul programada a las 8 de la mañana**", "Es a las 5:30. El documento se contradecía consigo mismo en la misma página, sobre el único día de rodaje del mes"],
    ["**Seis umbrales dados como criterio**", "Sin marcarlos como hipótesis. Es el error nº 1 del proyecto disfrazado de rigor, que es la forma más difícil de detectarlo"],
    ["**Gabriel contado como usuario de Instagram**", "Su ficha dice literal «Instagram lo usa poco para informarse». Es la persona que compra Elite todo el año"],
  ], [3300, 6348]),
  gap(220),
  panel("Lo que sigue sin poder verificarse", [
    "**No son errores: son huecos de información**, y quedan marcados porque decir que están cerrados sería el peor error de todos.",
    "El reel de referencia de Kurt · la prueba gratuita, el cobro y qué hace el plan tras un fallo · la fecha exacta del lanzamiento · **y qué cuentas y formatos consume cada persona, que es el hueco que más duele** y el que la caja de preguntas tiene que cerrar.",
  ], NARANJA),

  gap(500),
  new Paragraph({
    spacing: { before: 300, after: 80 },
    border: { top: { style: BorderStyle.SINGLE, size: 18, color: NEGRO, space: 10 } },
    children: [new TextRun({ text: "Organic Club", font: SERIF, size: 24, bold: true, color: NEGRO })],
  }),
  new Paragraph({
    children: [new TextRun({ text: "Liam Libre · Alonzo García   |   Estrategia de contenido Kenia · versión 1.1 · 10 de septiembre de 2026", font: SANS, size: 19, color: GRIS })],
  }),
  gap(200),
  p("El detalle completo vive en el repositorio, en estrategia-contenido: nueve documentos numerados más el registro de auditoría. Este documento es el resumen operativo; no los sustituye.", { color: GRIS_CLARO }),
);

/* ================= BUILD ================= */
const doc = new Document({
  creator: "Organic Club",
  title: "Kenia — Estrategia de contenido",
  description: "Estrategia de contenido y plan de lanzamiento de Kenia. Organic Club, septiembre de 2026.",
  numbering: {
    config: [
      { reference: "vinetas", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 220 } }, run: { color: CIAN, font: SANS } } },
        { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 220 } }, run: { color: CIAN, font: SANS } } },
      ]},
      { reference: "numerada", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 400, hanging: 260 } }, run: { color: CIAN, bold: true, font: SANS } } },
      ]},
    ],
  },
  styles: { default: { document: { run: { font: SANS, size: 21, color: NEGRO } } } },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 },
                          margin: { top: 1300, bottom: 1200, left: 1296, right: 1296 } } },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { before: 120 },
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: LINEA, space: 8 } },
      children: [new TextRun({ text: "Organic Club  ·  Estrategia de contenido Kenia  ·  v1.1  ·  10 de septiembre de 2026",
                               font: SANS, size: 15, color: GRIS_CLARO })],
    })] }) },
    children: C,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(process.argv[2], buf);
  console.log("escrito:", process.argv[2], (buf.length / 1024).toFixed(0) + " KB");
});
