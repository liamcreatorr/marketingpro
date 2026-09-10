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

const numbered = t => new Paragraph({
  numbering: { reference: "numerada", level: 0 },
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

function persona(rank, nombre, apodo, lede, campos, ayuda, nota) {
  const el = [
    new Paragraph({
      spacing: { before: 340, after: 0 },
      children: [
        new TextRun({ text: rank + "   ", font: SANS, size: 18, bold: true, color: CIAN, characterSpacing: 20 }),
        new TextRun({ text: nombre, font: SERIF, size: 25, bold: true, color: NEGRO }),
        new TextRun({ text: "   " + apodo, font: SERIF, size: 22, italics: true, color: GRIS_CLARO }),
      ],
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: LINEA, space: 6 } },
    }),
    new Paragraph({ spacing: { before: 140, after: 160, line: 288 }, children: runs(lede, { font: SANS, size: 21, color: GRIS }) }),
    table(["Campo", "Detalle"], campos, [2100, 7548]),
    gap(140),
    panel("Cómo lo ayuda Kenia", [ayuda]),
  ];
  if (nota) { el.push(gap(140)); el.push(panel(nota[0], nota.slice(1), NARANJA)); }
  return el;
}

/* ================= CONTENIDO ================= */
const C = [];
const add = (...xs) => xs.forEach(x => Array.isArray(x) ? C.push(...x) : C.push(x));

/* ---------- PORTADA ---------- */
add(
  gap(1000),
  eyebrow("Organic Club · Documento de marca"),
  new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Kenia", font: SERIF, size: 100, bold: true, color: NEGRO })] }),
  new Paragraph({ spacing: { after: 260 }, children: [new TextRun({ text: "Plataforma de marca completa", font: SERIF, size: 34, color: GRIS })] }),
  new Paragraph({ spacing: { after: 460 }, indent: { right: 7900 }, children: [],
    border: { bottom: { style: BorderStyle.SINGLE, size: 30, color: NARANJA, space: 2 } } }),
  table(["Campo", "Detalle"], [
    ["Cliente", "Ninro Libre — fundador y CEO de Kenia"],
    ["Preparado por", "Organic Club — Liam Libre y Alonzo García"],
    ["Fecha", "8 de septiembre de 2026"],
    ["Versión", "1.0 — para aprobación"],
    ["Base", "12 entrevistas a corredores de Caracas, auditoría del manual visual y las sesiones de trabajo con Ninro y con Kurt"],
  ], [2100, 7548]),
  gap(420),
  p("Este documento es la base sobre la que se escribe todo lo demás: la campaña de lanzamiento, el contenido de redes, los textos de la tienda y cualquier pieza que lleve el nombre de Kenia. No es un manual de estilo gráfico. Es lo que la marca dice, a quién se lo dice y cómo lo dice."),
  p("Todo lo que hay aquí está cerrado salvo lo que aparece marcado en la sección 15. Si algo no encaja con cómo ves la marca, dilo sobre este documento: es más barato cambiarlo ahora que después de grabar."),
  brk(),
);

/* ---------- ÍNDICE ---------- */
add(
  new Paragraph({ spacing: { after: 240 }, children: [new TextRun({ text: "Contenido", font: SERIF, size: 34, bold: true, color: NEGRO })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: CIAN, space: 8 } } }),
  gap(160),
  table(["#", "Sección", "Qué encontrarás"], [
    ["01", "La marca en una página", "El resumen con el que ya se puede escribir un anuncio correcto"],
    ["02", "De dónde sale esto", "Las tres fuentes y los cuatro hallazgos que cambiaron la estrategia"],
    ["03", "El nombre", "Qué significa Kenia, la respuesta oficial y el descriptor fijo"],
    ["04", "Círculo dorado", "Por qué existe la marca, cómo trabaja y qué vende"],
    ["05", "Misión y visión", "Y la métrica que prueba la visión"],
    ["06", "Valores", "Cuatro, cada uno con lo que cuesta"],
    ["07", "Posicionamiento", "Los dos enemigos y el mapa competitivo real"],
    ["08", "Público objetivo y buyer personas", "Las cinco fichas, ordenadas por prioridad"],
    ["09", "Beneficios", "Racionales, emocionales y cuál lidera en cada canal"],
    ["10", "Arquetipo y personalidad", "Quién es Kenia dentro de la historia del usuario"],
    ["11", "Tono de voz", "Los cuatro ejes, el tono por momento y las palabras prohibidas"],
    ["12", "Identidad visual", "Lo que se conserva, la jerarquía de color y lo que falta"],
    ["13", "Brief de moodboard", "Los siete territorios visuales"],
    ["14", "Lo que proponemos reforzar", "Ocho propuestas que hoy no existen en la plataforma"],
    ["15", "Decisiones pendientes", "Qué bloquea y quién lo desbloquea"],
    ["16", "Calendario", "Las fechas que mandan"],
  ], [640, 3400, 5608]),
  brk(),
);

/* ---------- 01 ---------- */
add(
  h1("La marca en una página"),
  p("Si alguien solo lee esta sección, tiene que poder escribir un anuncio correcto."),
  quote("El plan se adapta a ti, no tú al plan.", "La bisagra de toda la marca"),
  gap(180),
  table(["Pieza", "Definición"], [
    ["Qué es", "Kenia lleva tus marcas y te entrena con ellas. Registra tu historial, tus ritmos y tu carga, y sobre eso arma un plan que se recalcula cada día."],
    ["Por qué existe", "Nadie deja de correr por falta de voluntad. Deja de correr porque sigue un plan que no sabe nada de su vida, y cuando ese plan se rompe lo hace sentir a él como el que falló."],
    ["Para quién", "Quien ya corre y no tiene quien le diga qué hacer mañana. El nivel no define el público."],
    ["Contra qué", "El plan que te pasó un conocido por WhatsApp. Cuatro de doce corredores entrevistados siguen uno. Ese es el rival real, no otra app."],
    ["Misión", "Ayudamos a cualquier persona que quiera correr a sostener el entrenamiento dentro de su vida real, con un plan que decide con ella cada día, para que llegue a su meta sin abandonar en el intento."],
    ["Visión", "Que en Latinoamérica correr deje de ser cosa de corredores."],
    ["Valores", "Nunca con culpa · Explicamos siempre el porqué · El plan se adapta a ti · Prometemos solo lo que sostenemos"],
    ["Arquetipo", "El usuario es el héroe. Kenia es su mentor: Sabio dominante, Cuidador de apoyo."],
    ["Personalidad", "Atento · Claro · Firme · Honesto · Cercano"],
    ["Descriptor", "Kenia · Entrenamiento de running"],
    ["Color", "El azul es la marca. El naranja es el momento."],
    ["Lo que nunca se dice", "«Para corredores» y «sin excusas», en cualquiera de sus formas."],
  ], [1900, 7748]),
  brk(),
);

/* ---------- 02 ---------- */
add(
  h1("De dónde sale esto"),
  p("Nada de este documento es opinión nuestra sin respaldo. Estas son las tres fuentes."),
  gap(120),
  table(["Fuente", "Qué aportó"], [
    ["12 entrevistas", "Corredores de Caracas de la red del box y de tu entorno, entrevistados en septiembre con un cuestionario de nueve preguntas. Es la base de casi todo lo que sigue."],
    ["El manual visual", "Auditamos las diecisiete piezas del manual existente contra la estrategia del producto. Doce hallazgos, repartidos entre lo que resolvemos nosotros y lo que le toca al diseñador."],
    ["Las sesiones contigo y con Kurt", "Producto, precios, modelo de negocio, qué hace y qué no hace la app, y las correcciones que fuiste haciendo sobre cada versión de este documento."],
  ], [2400, 7248]),

  h2("Los cuatro hallazgos que más cambiaron la estrategia"),
  table(["Hallazgo", "Qué cambió"], [
    ["Solo 1 de 12 se llama «corredor» sin ponerle un adjetivo o un pero",
     "La marca no puede pedir ese título para entrar. Ni el descriptor, ni el tagline, ni la campaña pueden decir «para corredores»: le estarían diciendo a once de cada doce que esto no es para ellos."],
    ["10 de 12 descansan sin culpa",
     "Creíamos que el permiso para descansar era el gran diferencial emocional. No lo es: ya se lo saben. La culpa aparece cuando la vida gana, y sobre todo en lo que cuesta retomar. El mensaje central pasó a ser volver sin castigo."],
    ["Nadie nombró a Runna ni a RunMotion",
     "El competidor real es el plan informal de un conocido, y detrás asoma la IA generalista: alguien ya arma sus planes con Gemini. Cambia el argumento de venta completo."],
    ["El clima es la causa nº 1 de sesión perdida — 4 de 12",
     "Y en la reunión del 8 le diste la vuelta: el problema no es que llueva, es no estar preparado para correr cuando llueve. Eso abrió un territorio de producto y de contenido que ninguna app global va a ocupar aquí."],
  ], [3300, 6348]),
  gap(200),
  panel("Un dato de precio que vale dinero", [
    "Preguntamos cuánto cobra un entrenador de running en Caracas. Los cuatro que lo sabían dieron un rango de **30 a 80 dólares al mes**, y las asesorías personalizadas desde 100. Kenia Coach cuesta 19,99. La comparación ya se puede publicar con dato propio.",
    "Con una advertencia de la misma investigación: **ocho de doce no sabían ese precio**. La comparación no se puede dar por sabida. Hay que enseñarla, no insinuarla.",
  ]),
  brk(),
);

/* ---------- 03 ---------- */
add(
  h1("El nombre"),
  p("El nombre está cerrado. Lo que hacía falta era decidir qué significa y cómo se sostiene."),

  h2("El país aporta el método, no la estética"),
  p("Kenia se abraza como referencia, con una condición: hablamos del **método**, no de la geografía. El vínculo es real, porque entrenaste con método keniano, y eso es sustento, no decoración."),
  table(["Sí se usa", "Nunca se usa"], [
    ["El método", "La bandera"],
    ["La cultura de volumen y paciencia", "Los colores panafricanos"],
    ["La idea de correr como oficio", "Mapas, siluetas del continente y un keniano corriendo como recurso gráfico"],
  ], [4824, 4824]),
  gap(180),
  p("El manual actual tiene un mapa de Kenia en la lámina «About Project». Ahí es donde la referencia se cae en decoración y conviene quitarla."),

  h2("La respuesta oficial cuando alguien pregunte «¿por qué Kenia?»"),
  quote("Kenia es el país donde correr es un oficio. De ahí viene el método con el que entrenamos."),
  gap(180),
  p("Corta, honesta y verificable. No promete origen africano ni presencia allá. Dice de dónde viene el método, que es lo cierto."),

  h2("El nombre necesita ayuda, y hay que dársela"),
  p("En español «Kenia» no evoca running. Lo probaron ustedes mismos: hubo quien leyó una app para mujeres y quien leyó una app de citas. Además es un nombre propio femenino común en la región. La conclusión no es cambiar el nombre. Es que el nombre **no puede cargar el significado solo**: hay que contarlo en el About, en el onboarding y al menos una vez en la campaña de lanzamiento."),
  gap(60),
  table(["Elemento", "Cómo queda"], [
    ["Descriptor fijo", "Kenia · Entrenamiento de running. Va en la bio de Instagram, el header de la web, la firma de campaña, la portada del brandbook y la ficha de tienda. Describe la actividad, no la identidad: no obliga a nadie a llamarse corredor."],
    ["Ficha de tienda — nombre", "Kenia: Entrenamiento Running (28 de 30 caracteres)"],
    ["Ficha de tienda — subtítulo", "Tu plan se adapta a tu semana (29 de 30 caracteres)"],
    ["Escritura de la marca", "Kenia, con mayúscula inicial y el resto en minúscula. Ni KENIA ni kenia."],
    ["Escritura de los planes", "Kenia Starter, Kenia Coach y Kenia Elite. Hoy están en la app en mayúsculas: hay que unificarlo antes de subir a tienda."],
  ], [2800, 6848]),
  brk(),
);

/* ---------- 04 ---------- */
add(
  h1("Círculo dorado"),
  p("Por qué existe la marca, cómo trabaja y qué vende. En ese orden, porque en ese orden se comunica."),

  h2("Por qué — la creencia"),
  quote("Nadie deja de correr por falta de voluntad. Deja de correr porque sigue un plan que no sabe nada de su vida, y cuando ese plan se rompe, lo hace sentir a él como el que falló."),
  gap(180),
  p("Kenia existe porque correr no es de un tipo de cuerpo ni de un tipo de persona. Es de cualquiera que tenga un plan que se adapte a él, en vez de exigirle que él se adapte al plan."),

  h2("Cómo — las cinco formas de trabajar"),
  table(["Forma", "Qué significa en la práctica"], [
    ["Pregunta antes de mandar", "Sueño, molestias, estado de las piernas, estrés, tiempo real disponible. Y entonces decide: mantener, bajar intensidad, modificar o descansar."],
    ["Explica el porqué de cada sesión", "No «hoy 6 × 3 min», sino «hoy 6 × 3 min porque tu objetivo es el 10K, y como dormiste mal quitamos una repetición»."],
    ["Aprende de ti", "Cómo respondes a la carga, cuánto tardas en recuperar, cuándo te sobrecargas, qué días te saltas."],
    ["Optimiza que sigas, no que el plan sea perfecto", "El mejor plan es inútil si abandonas a las cuatro semanas."],
    ["Método de autor, vigilancia de equipo, ejecución de máquina", "La fórmula la diseñas tú. La IA la aplica persona a persona. El equipo vigila el conjunto desde el dashboard. Nadie de Kenia habla con un usuario sobre su plan, pero nadie deja los planes sin mirar."],
  ], [3000, 6648]),
  gap(200),
  panel("Advertencia de caducidad", [
    "«Monitoreamos todos los planes» es cierto con 40 suscriptores y deja de serlo con 1.200. Sirve como prueba de rigor hoy. No se construye la marca encima de esa frase.",
  ], NARANJA),

  h2("Qué — el producto"),
  p("App de suscripción que lleva tus marcas y te entrena con ellas. Registra historial, ritmos, distancias y carga, y sobre eso arma un plan que se recalcula a diario. Más comunidad, ligas, ranking y rutas compartidas."),
  gap(60),
  table(["Plan", "Mes", "Año (−15 %)", "Qué transformación vende"], [
    ["Kenia Starter", "9,99 $", "101,90 $", "Te da dirección. Plan personalizado, entrenamientos diarios, seguimiento, explicación básica, comunidad."],
    ["Kenia Coach ★", "19,99 $", "203,90 $", "Te acompaña. Adaptación diaria, perfil adaptativo, coach conversacional, explicación avanzada, prevención."],
    ["Kenia Elite", "29,99 $", "305,90 $", "Te conoce y anticipa. Cuerpo digital, planificación predictiva, estrategia de competición, recuperación post-carrera."],
  ], [1700, 1000, 1350, 5598]),
  gap(200),
  p("**Kenia Coach es el plan estrella** y el que hay que vender a la mayoría: es el que encarna la promesa de la marca. Starter baja la barrera de entrada. Elite ancla el precio y, desde esta versión, ya tiene comprador definido, como se ve en la sección 8."),
  gap(60),
  panel("Regla que no se rompe", [
    "Ningún plan, Elite incluido, puede prometer supervisión ni revisión humana. Kenia es un coach virtual y no interviene ninguna persona en el plan de un usuario. Cualquier pieza que insinúe lo contrario rompe el cuarto valor de la marca y termina en reembolsos y reseñas de una estrella.",
  ], NARANJA),
  brk(),
);

/* ---------- 05 ---------- */
add(
  h1("Misión y visión"),
  p("La misión es lo que hacemos hoy. La visión es el mundo que queremos dejar."),
  h2("Misión"),
  quote("Ayudamos a cualquier persona que quiera correr a sostener el entrenamiento dentro de su vida real, con un plan que decide con ella cada día, para que llegue a su meta sin abandonar en el intento."),
  gap(180),
  p("La prueba que le hicimos: si tapas la palabra Kenia, ¿serviría para cualquier app? No. «Un plan que decide con ella cada día» es específico de este producto y de ningún otro."),
  p("Versión corta para uso interno: que entrenar para correr sea sostenible en la vida real."),
  h2("Visión"),
  quote("Que en Latinoamérica correr deje de ser cosa de corredores."),
  gap(180),
  p("Que cualquiera, sin importar su cuerpo, su horario o su punto de partida, pueda decir «yo entreno» y sostenerlo, porque tuvo un entrenador que se adaptaba a su vida y no al revés."),
  gap(60),
  panel("La métrica que prueba la visión", [
    "No son descargas ni facturación. Es **cuánta gente sigue entrenando a los seis meses**. Es la métrica de «no abandonó», y es la única que demuestra que la marca cumple lo que dice.",
    "Falta ponerle horizonte largo: una meta de negocio a tres años. El punto de partida ya lo tenemos —40 suscriptores a los 90 días— pero hace falta que decidas a dónde va eso.",
  ]),
  brk(),
);

/* ---------- 06 ---------- */
add(
  h1("Valores"),
  p("Cuatro. Cada uno con lo que cuesta, porque un valor que no cuesta nada no es un valor."),

  h3("1 · Nunca con culpa"),
  p("Kenia no usa la vergüenza como motor. Ni en una notificación, ni en un post, ni en una campaña."),
  p("**Lo que cuesta:** renunciamos a la palanca emocional más barata y más eficaz del fitness. La frase «donde termina tu excusa, empieza tu carrera» se retira aunque funcione en engagement."),
  p("**El momento crítico de este valor no es el día de descanso: es el primer día después del fallo.** Ahí es donde las entrevistas encontraron la culpa. «Tardé bastante en retomar la rutina.» «Me siento muy incómoda cuando no hago lo que está pautado ese día.»"),

  h3("2 · Explicamos siempre el porqué"),
  p("Ninguna instrucción sin su razón."),
  p("**Lo que cuesta:** es más caro de construir y más lento de comunicar que un simple «haz esto». En marketing significa contenido que enseña, no que arenga."),

  h3("3 · El plan se adapta a ti, tú no al plan"),
  p("Descansar es parte del plan, no una falla del plan."),
  p("**Lo que cuesta:** aceptamos que alguien entrene menos si su cuerpo lo pide, aunque «menos uso» se vea peor en las métricas de engagement."),
  gap(40),
  panel("Adaptarse no es esquivar", [
    "Tu planteamiento del 8 de septiembre entra aquí y no contradice el valor, porque son dos cosas distintas:",
    "**La condición climática se entrena**, a propósito y con su porqué, porque el día de la carrera no se elige el clima.",
    "**La vida real se adapta.** Fiebre, un hijo enfermo, cuatro horas de sueño, una tormenta con rayos: ahí el plan se mueve.",
    "**El clima se entrena, el riesgo se esquiva, y a la persona no se la juzga nunca.**",
  ]),

  h3("4 · Prometemos solo lo que sostenemos"),
  p("No vendemos un entrenador humano si lo que hay es un coach virtual. No inflamos lo que hace la IA."),
  p("**Lo que cuesta:** Kenia Elite se vende con menos épica de la que podríamos. Este valor salió del propio equipo técnico, que avisó que no prometería un entrenador humano personal por treinta dólares salvo que se pueda sostener. Ese aviso se convirtió en valor de marca."),

  h2("Política de cuerpo"),
  p("Consecuencia directa del primer valor. Son reglas duras, no recomendaciones:"),
  bullet("No se pide el peso como métrica principal ni se grafica como objetivo por defecto."),
  bullet("No hay fotos de antes y después."),
  bullet("Las calorías no se comunican como objetivo. Si aparecen, son dato secundario del entrenamiento."),
  bullet("Los cuerpos en la comunicación son diversos por norma, no atléticos por norma."),
  bullet("Si un usuario declara un objetivo que sugiere un trastorno alimentario, la app no lo optimiza: lo deriva. Falta definir cómo, contigo y con Kurt."),
  brk(),
);

/* ---------- 07 ---------- */
add(
  h1("Posicionamiento"),
  p("Contra qué compite Kenia de verdad, y con qué frase se defiende."),

  h2("Los dos enemigos"),
  table(["", "Enemigo", "Dónde se usa"], [
    ["A", "El plan genérico que te exige adaptarte a él", "Contenido de producto, comparativas, demos, pauta de conversión. Ataca el método de la categoría."],
    ["B", "La creencia de que tú no eres corredor", "Narrativa de marca, comunidad, contenido de captación. Ataca la cabeza del usuario, y está confirmado: once de doce entrevistados no se llaman corredores."],
  ], [600, 3100, 5948]),
  gap(200),
  p("La bisagra que une a los dos, y que puede funcionar como declaración de marca: **el plan se adapta a ti, no tú al plan.**"),

  h2("El mapa competitivo real"),
  table(["Frente", "Quién", "Cómo se le responde"], [
    ["Rival de percepción", "Strava", "No se le ataca. Hacen cosas distintas: Strava es el muro donde publicas lo que hiciste, Kenia es la voz que decide contigo lo que vas a hacer."],
    ["Rival funcional real", "El plan que te pasó un conocido por WhatsApp", "Ese plan no sabe que hoy llovió, que dormiste mal ni que el martes tienes reunión. Y tu amigo no te va a contestar a las 5:30 de la mañana."],
    ["Rival emergente", "IA generalista: Gemini, ChatGPT", "Kenia no compite contra la idea de un plan hecho por IA. Compite contra una IA que no te conoce."],
    ["Referente de producto", "Runna, RunMotion Coach", "Son lo que Kenia quiere ser de mayor. Nadie los nombró en las entrevistas: no sirven como argumento de venta aquí."],
  ], [2000, 2400, 5248]),

  h2("Declaración de posicionamiento"),
  p("Esta frase no es para publicar. Es la que resuelve discusiones internas cuando haya que decidir si una pieza sirve o no."),
  quote("Para el corredor venezolano que ya entrena y no tiene quien decida con él, Kenia es un entrenador de running que se recalcula cada día. A diferencia del plan genérico, del que te pasó un conocido o de una IA que no te conoce, Kenia pregunta cómo estás antes de mandarte a entrenar, te explica por qué, y ajusta cuando tu semana cambia."),
  gap(200),
  panel("La regla comercial que ordena todo", [
    "Dar un plan es fácil. Lo dijiste tú: «es más como una fórmula». Garmin Coach lo regala, Nike Run Club lo regala, cualquier PDF de internet lo tiene.",
    "**Entonces el plan no es el producto.** Lo que vende Kenia empieza el martes que dormiste cuatro horas, el jueves que te duele el gemelo, la semana que el trabajo se te comió. Vendemos la adaptación, no la planificación.",
    "Cualquier pieza que ponga «plan personalizado» como titular principal está vendiendo la parte que los demás regalan.",
  ]),
  brk(),
);

/* ---------- 08 ---------- */
add(
  h1("Público objetivo y buyer personas"),
  p("Cinco perfiles, ordenados por lo que aporta cada uno al negocio."),
  quote("Kenia es para quien ya corre y no tiene quien le diga qué hacer mañana."),
  gap(200),
  p("De 25 a 45 años, en Caracas y las principales ciudades del país, con ingresos en dólares. Ya tienen el hábito y están en el proceso: preparando una carrera, intentando bajar una marca, volviendo de una lesión, o queriendo dejar de correr siempre lo mismo. Ya pagan por entrenar —box, gimnasio, carreras, reloj— y valoran no lesionarse tanto como mejorar."),
  gap(60),
  panel("El criterio que ordena el público", [
    "**El público no se define por nivel. Se define por quién decide tu semana.** Eso permite que quepan en la misma marca una persona que corre 5 km tres veces por semana y otra que hace 85 kilómetros semanales, sin que el mensaje se rompa. Lo que tienen en común no es el ritmo: es que nadie decide con ellos.",
    "Queda fuera quien no corre y no tiene el hábito, y quien ya tiene un entrenador humano que lo ve correr y está contento con él.",
  ]),
  gap(160),
  panel("La restricción que atraviesa a las cinco", [
    "Kenia cobra en dólares. El público no es «runners de Venezuela»: es **el runner venezolano con capacidad de pago en divisas** —ingreso remoto o freelance, profesional independiente, negocio propio o remesa estable—. Decirlo explícito no reduce el mercado: evita que la pauta le hable a quien no puede pagar y nos reviente el costo por adquisición.",
  ], NARANJA),

  h2("Las cinco, en orden de prioridad"),
  table(["#", "Persona", "Qué aporta al negocio", "Plan"], [
    ["1", "Daniel, 29 — el que ya se inscribió", "Volumen y velocidad de conversión. Siete de doce entrevistados están en esta situación.", "Elite"],
    ["2", "Patricia, 42 — la que vuelve", "Mayor valor a largo plazo y la menos sensible al precio.", "Coach"],
    ["3", "Rebeca, 34 — la que lleva años igual", "El segmento más grande y el peor atendido.", "Coach"],
    ["4", "Andrés, 32 — el que corre y levanta", "La comunidad semilla y las primeras reseñas de tienda.", "Coach"],
    ["5", "Gabriel, 36 — el que va por la marca", "Prueba y reputación, no volumen. Es quien sostiene Elite todo el año.", "Elite"],
  ], [500, 2900, 5148, 1100]),
  gap(200),
  p("**Las dos que lideran la campaña de lanzamiento son Daniel y Patricia.** Las otras tres se atienden con contenido orgánico y comunidad, no con pauta fría."),
  brk(),
);

/* --- fichas --- */
add(
  persona("01", "Daniel, 29", "el que ya se inscribió",
    "Tiene una carrera con fecha, un tiempo en la cabeza y miedo de llegar roto. Siete de doce entrevistados están en esta situación: es el perfil más numeroso y el único con urgencia propia.",
    [
      ["Edad y demografía", "27 a 34 años. Caracas, este de la ciudad. Vive solo o en pareja, sin hijos."],
      ["Ocupación", "Profesional independiente o empleado con nómina en dólares. Horario de oficina, entrena antes o después."],
      ["Hábitos", "Corre 3 o 4 veces por semana. Ya tiene reloj y mira ritmo, zonas y progresión después de cada salida."],
      ["Hobbies", "Fútbol o pádel con amigos, videojuegos, viajes cortos. Le gustan las cosas que se miden."],
      ["Medios que consume", "Instagram y Strava a diario. YouTube para análisis de carreras y reviews de zapatillas. Grupos de WhatsApp de running de la ciudad."],
      ["Hábitos de compra", "Compra por especificación, no por marca: compara antes. Ya invirtió en reloj y zapatillas buenas. Dijo que pagaría entre 30 y 50 dólares al mes."],
      ["Objetivos", "Terminar el 21K en un tiempo concreto, no caminar, y no lesionarse en las últimas cuatro semanas, que es su miedo real."],
      ["Personalidad", "Analítico, competitivo consigo mismo, impaciente. Quiere entender antes de obedecer: si no le explicas el porqué, cambia el plan por su cuenta."],
      ["Su objeción", "«Esto lo tengo gratis en Garmin Coach.»"],
      ["Qué lo hace irse", "«Ya corrí la carrera.» Es el churn más grave de los cinco: hay que engancharlo al siguiente objetivo antes de que cruce la meta del primero."],
    ],
    "La periodización hasta el día de la carrera, con la fecha metida dentro del plan. El aviso de sobrecarga, que es exactamente lo que lo rompe. El porqué de cada sesión, que es lo que necesita para no improvisar. Y la estrategia de competición de Kenia Elite."),
  brk(),

  persona("02", "Patricia, 42", "la que vuelve",
    "Ya se lesionó una vez y no piensa repetirlo. Es la menos sensible al precio de las cinco y la que más tiempo se queda.",
    [
      ["Edad y demografía", "38 a 48 años. Caracas. Casada o divorciada, con hijos en edad escolar. Clase media consolidada."],
      ["Ocupación", "Profesional con carrera hecha: gerente, abogada, médica, dueña de un negocio pequeño. Agenda apretada y poco margen para improvisar."],
      ["Hábitos", "Corre temprano porque después ya no puede. Combinó running con gimnasio y ahí fue donde se lesionó. Hace o hizo fisioterapia."],
      ["Hobbies", "Yoga o pilates, cocina, lectura, planes con la familia el fin de semana."],
      ["Medios que consume", "Instagram de bienestar, salud femenina y fisioterapia. WhatsApp como canal principal. Le hace más caso a su fisio y a su médico que a cualquier app."],
      ["Hábitos de compra", "Paga por tranquilidad y por que se lo resuelvan. Ya paga fisio, gimnasio y consultas. Decide despacio, pero cuando decide se queda."],
      ["Objetivos", "Volver a correr sin volver a lesionarse y sostener el hábito. Quizá un 10K, pero terminarlo entera importa más que el tiempo."],
      ["Personalidad", "Prudente, responsable, realista. No busca destacar y desconfía de las promesas grandes."],
      ["Su objeción", "«¿Una app va a saber cómo está mi rodilla?»"],
      ["Qué la hace irse", "«Me volví a lesionar.» Es el peor churn posible para esta marca, porque falla en la promesa más literal."],
    ],
    "Progresión conservadora de carga, con techo. El aviso de sobrecarga antes de que duela, que es literalmente lo que le pasó. Que le pregunte por la rodilla y por el gimnasio antes de asignarle la sesión. Y que el descanso esté dentro del plan y no sea una derrota.",
    ["Riesgo de marca con esta persona",
     "Con Patricia el copy de prevención tiene que ser conservador: **«reducimos el riesgo», nunca «no te vas a lesionar»**. Es el cuarto valor aplicado a la letra."]),
  brk(),

  persona("03", "Rebeca, 34", "la que lleva años igual",
    "Corre desde hace años, no falla nunca, y sus tiempos son los mismos que hace dos años. Es el segmento más grande y el peor atendido de los cinco.",
    [
      ["Edad y demografía", "30 a 40 años. Caracas, Valencia, Maracaibo o Barquisimeto. En pareja, con o sin hijos pequeños."],
      ["Ocupación", "Profesional empleada, freelance en divisas o negocio propio pequeño. Trabajo estable y rutina fija."],
      ["Hábitos", "Tres veces por semana desde hace años. La misma ruta, la misma distancia, el mismo ritmo. Nunca ha seguido un plan y nadie le ha explicado por qué no progresa."],
      ["Hobbies", "El grupo de corredoras del domingo, brunch, series, viajes. Lo social pesa tanto como el deporte."],
      ["Medios que consume", "Instagram es su canal. Cuentas de corredoras como ella, no de élite. TikTok de rutinas y consejos. Confía más en otra corredora que en un experto."],
      ["Hábitos de compra", "Ya invierte en zapatillas e inscripciones. Pagaría alrededor de 30 dólares. Compra por recomendación, no por anuncio: necesita que alguien de su grupo lo esté usando."],
      ["Objetivos", "Su primer 21K, o bajar una marca que lleva dos años clavada. Quiere sentir que avanza."],
      ["Personalidad", "Constante, disciplinada, sin ego competitivo. Se aburre con lo complicado. Su frase es «yo solo quiero correr»."],
      ["Su objeción", "«¿No es demasiado complicado? Yo solo quiero correr.»"],
      ["Qué la hace irse", "«Es mucha estructura para mí.» El onboarding tiene que demostrarle valor en la primera semana, no en la sexta."],
    ],
    "Explicarle por qué correr siempre igual no mejora nada. Meterle variedad con sentido: series, ritmo, largos. Y mostrarle progresión donde antes solo veía kilómetros: es la única de las cinco para quien llevar las marcas sí es argumento de venta, porque es la primera vez que sus datos significan algo."),
  brk(),

  persona("04", "Andrés, 32", "el que corre y levanta",
    "Es un animal en el box y se ahoga en los 5K, y está subiendo carga por los dos lados a la vez. Es la comunidad semilla y son los primeros usuarios reales de la app.",
    [
      ["Edad y demografía", "27 a 38 años. Caracas. Soltero o en pareja, sin hijos. Vida social organizada alrededor del box."],
      ["Ocupación", "Desarrollador remoto, diseñador, freelance o empleado de startup. Cobra en dólares, horario flexible."],
      ["Hábitos", "Box 4 o 5 días por semana, corre 1 o 2 sin plan. Sube carga en las dos disciplinas al mismo tiempo, que es el patrón que lesionó a 3 de los 12 entrevistados."],
      ["Hobbies", "El box es el hobby, la tribu y la agenda. Además videojuegos, nutrición y suplementación, tecnología."],
      ["Medios que consume", "El grupo de WhatsApp del box y el box mismo. Instagram de CrossFit y hyrox, YouTube y TikTok de técnica. Quien lo influye es su coach y sus compañeros, no una marca."],
      ["Hábitos de compra", "Ya paga mensualidad de box, así que la comparación de precio le es favorable y no discute suscripciones. Compra rápido y por impulso social: si el box lo usa, lo usa."],
      ["Objetivos", "No ser el que se ahoga en la carrera del WOD. Hacer un 10K decente. Que correr deje de ser su punto débil."],
      ["Personalidad", "Competitivo, orgulloso, con disciplina probada. Le cuesta ser principiante en algo. Aprende rápido y por eso mismo cree que ya no te necesita."],
      ["Su objeción", "«¿Para qué pago si mi coach del box me puede armar algo?»"],
      ["Qué lo hace irse", "«Ya entendí cómo armar la rutina, la sigo por mi cuenta.» Su retención depende de que el plan siga cambiando con él."],
    ],
    "Ser la única app que le pregunta por el gimnasio antes de mandarlo a correr, y repartirle la semana entre box y running sin que se reviente. Es también la persona que encarna tu cruzada: el que hace CrossFit y sí corre."),
  brk(),

  persona("05", "Gabriel, 36", "el que va por la marca",
    "Hace el volumen, hace el trabajo, y el reloj no baja. Quiere bajar de tres horas en maratón y está haciendo la periodización solo. Sale del caso que trajiste tú de la reunión.",
    [
      ["Edad y demografía", "32 a 45 años. Caracas. Casado, con o sin hijos. Vida muy estructurada, se acuesta temprano."],
      ["Ocupación", "Ingeniero, gerente, consultor o dueño de negocio. Alguien con control sobre su agenda, porque sin ese control no se sostienen seis sesiones semanales."],
      ["Hábitos", "5 o 6 días, entre 70 y 90 kilómetros por semana. Series en pista, tirada larga el domingo, fuerza dos veces como complemento. Duerme y come en función de entrenar."],
      ["Hobbies", "El running es el hobby. Además ciclismo o natación de complemento, café de especialidad, seguir las majors y analizar sus propios datos."],
      ["Medios que consume", "Strava a diario y con club. YouTube de análisis de entrenamiento y de carreras, podcasts, foros. Usa vocabulario técnico correcto: umbral, VO2, zonas. Instagram lo usa poco para informarse."],
      ["Hábitos de compra", "El de mayor disposición a pagar de los cinco. Zapatillas de placa de carbono, reloj de gama alta, viajes a carreras. Compara Kenia con un entrenador de 80 a 100 dólares, no con una app gratis."],
      ["Objetivos", "Un número. Sub-3 en maratón, o sub-40 en 10K. Tiene fecha, tiene carrera elegida y tiene un histórico de intentos fallidos."],
      ["Personalidad", "Metódico y exigente, paciente con el proceso e impaciente con lo que no tiene fundamento. No necesita motivación, necesita criterio. Detecta al instante una sesión mal puesta, y si la detecta pierde la confianza en el sistema entero."],
      ["Su objeción", "«Una app no me va a dar lo que me da un entrenador que me ve correr.»"],
      ["Qué lo hace irse", "«El plan no me llevó a la marca», o mucho antes, «me puso una sesión que yo sabía que estaba mal». Se va por falta de criterio, no por falta de constancia."],
    ],
    "Periodización con intención hacia una fecha: bloques, picos y tapering, no un plan de doce semanas igual para todos. Control de carga en volumen alto, que es justo donde él se lesiona. El cuerpo digital de Kenia Elite y la estrategia de competición. Y algo que a los otros cuatro les da igual: que sus marcas sean la materia prima de la decisión de mañana.",
    ["Esta persona resuelve un hueco, y hay que tratarla con cuidado",
     "**Kenia Elite estaba definida y con precio, pero no tenía comprador sostenido.** Daniel también compra Elite, pero solo las doce semanas que dura su preparación. Gabriel es el único que lo sostiene todo el año, porque su temporada no termina: termina un objetivo y empieza el siguiente.",
     "**Pero es un caso, no un segmento.** Los otros cuatro perfiles salen de doce entrevistas. Este sale de una conversación sobre una sola persona.",
     "**Lo que proponemos:** que tu conocido use Kenia Elite gratis o a precio de fundador durante su bloque de maratón, a cambio de decir sin filtro qué sesiones no le cuadran. Es la prueba de estrés más barata que va a tener el producto, y si sale bien es la mejor historia de lanzamiento que puede tener la marca.",
     "**Lo que no haríamos:** poner pauta detrás de este perfil en octubre. Son pocos, caros de alcanzar y los más difíciles de convencer. Su valor es de prueba y reputación, no de volumen."]),
  brk(),
);

/* ---------- 09 ---------- */
add(
  h1("Beneficios"),
  p("Los racionales justifican la compra. Los emocionales la provocan. Hacen falta los dos, y saber cuál va primero en cada canal."),

  h2("Racionales"),
  table(["", "Beneficio", "Prueba"], [
    ["R1", "El plan se recalcula a diario según sueño, fatiga, dolor, estrés y tiempo disponible", "Ni Strava, ni Nike Run Club, ni Garmin Coach lo hacen. Runna y RunMotion adaptan, pero por resultado, no por estado previo."],
    ["R2", "Entiendes por qué haces cada sesión y por qué cambió", "Explicación en cada entrenamiento."],
    ["R3", "Aviso temprano de sobrecarga, con alternativas automáticas", "«Tu carga subió 32 % en dos semanas y tu rendimiento empeora.»"],
    ["R4", "Le hablas al coach y responde", "Coach conversacional desde Kenia Coach."],
    ["R5", "Cuesta una fracción de un entrenador", "Dato propio: un entrenador en Caracas cobra de 30 a 80 dólares al mes. Kenia Coach cuesta 19,99."],
    ["R6", "En español y para tu realidad", "RunMotion está pensado para el público francés; Runna para el anglosajón."],
    ["R9", "Te prepara para correr en cualquier condición y te protege cuando la condición es un riesgo", "El clima es la causa nº 1 de sesión perdida en Caracas. Ninguna carrera se suspende por lluvia ni por calor, y ninguna app global modela las condiciones de aquí."],
  ], [600, 3400, 5648]),
  gap(200),
  p("**Los tres que hacen el trabajo pesado son R1, R2 y R3.** El resto acompaña. R9 todavía no está construido: es la recomendación de producto con mejor relación impacto-esfuerzo que salió de toda la investigación."),

  h2("Emocionales"),
  table(["", "Beneficio", "De qué libera"], [
    ["E8", "Volver sin castigo — el central", "Perder un día, o tres semanas, no significa empezar de cero. El plan se recalcula y sigue. Es donde vive la culpa real."],
    ["E3", "Legitimidad", "Poder decir «soy corredor» sin sentirse impostor. Verificado: solo 1 de 12 lo dice sin matizar."],
    ["E1", "Alivio", "De la fatiga de decidir cada día si entrenar, cuánto y cómo."],
    ["E4", "Seguridad", "Del miedo a repetir la lesión. Es todo lo que le importa a Patricia."],
    ["E5", "Ser entendido", "De la soledad del que entrena solo."],
    ["E6", "Pertenencia", "Del ranking que te deja siempre abajo."],
    ["E7", "Orgullo silencioso", "De necesitar demostrarlo. No el de la foto en la meta: el de haber sostenido doce semanas."],
  ], [600, 2700, 6348]),
  gap(200),
  panel("La excepción: Gabriel", [
    "La persona 5 no se compra con emocionales. Ninguno de los ocho le mueve la aguja: no le falta constancia, no le falta legitimidad y no necesita permiso para descansar.",
    "A él lo convencen R1, R2 y R3, y lo que Elite añade. **Es el único perfil que se vende con argumento técnico puro**, y también el único que detecta al instante si el argumento es hueco.",
  ]),

  h2("Qué beneficio va primero en cada canal"),
  table(["Canal", "Primero", "Después", "Por qué"], [
    ["Pauta en Meta", "R1 · se recalcula a diario", "E1 · alivio", "Público frío: hay que decir qué es en tres segundos."],
    ["Orgánico y contenido", "E3 · legitimidad", "R2 · el porqué", "Es donde vive el enemigo B y donde se construye relato."],
    ["Ficha de tienda", "R1 · R2 · R3", "E8", "Quien llega ahí ya sabe que quiere entrenar. Quiere saber qué hace la app."],
    ["Comunidad y ligas", "E6 · pertenencia", "E7 · orgullo", "El que ya está adentro no necesita que le vendan."],
    ["Reactivar inactivos", "E8 · volver sin castigo", "R1", "Lo que frena el regreso no es la falta de permiso: es lo que cuesta retomar."],
    ["Página de precios", "R5 · fracción de un entrenador", "R1", "Enseñar la comparación, no insinuarla: ocho de doce no sabían el precio de un entrenador."],
  ], [2000, 2500, 1600, 3548]),
  brk(),
);

/* ---------- 10 ---------- */
add(
  h1("Arquetipo y personalidad"),
  p("Quién es Kenia dentro de la historia del usuario."),
  quote("El usuario es el héroe. Kenia es quien lo entrena.", "Sabio dominante · Cuidador de apoyo"),
  gap(200),
  p("Esto resuelve una discusión que tuvimos y que valía la pena tener. El instinto del equipo era el arquetipo del Héroe, porque la marca le habla a quien ya corre y se prepara para maratones. Y es correcto en el territorio, pero hay que separar dos cosas:"),
  gap(60),
  table(["", "Quién es el héroe", "Qué implica"], [
    ["El universo de la marca", "Héroe", "Carreras, kilómetros, meta, madrugada, línea de salida. La estética oscura, intensa y de alto rendimiento es correcta y se conserva."],
    ["El rol de la marca", "Sabio · mentor", "El que prepara al héroe. La voz explica, adapta y protege. No arenga ni avergüenza."],
  ], [2500, 1800, 5348]),
  gap(200),
  p("En el esquema clásico, el mentor del héroe nunca es otro héroe. Y hay una razón práctica que decide el asunto: **el Héroe nunca te dice que pares.** Su mensaje raíz es «empuja más». Y el comportamiento más distintivo de Kenia, el que ningún competidor tiene, es decirte «hoy no entrenas» cuando tu cuerpo lo pide."),
  gap(60),
  panel("Lo que esto conserva del sistema visual actual", [
    "El mundo visual —negro, naranja, cuerpos en esfuerzo, asfalto, velocidad— **es el correcto**, porque es el mundo del héroe al que la marca acompaña. Lo que hay que corregir es mucho más quirúrgico: el copy. Las cuatro frases del manual no fallan por intensas, fallan porque hablan como el héroe en lugar de hablarle al héroe.",
  ]),

  h2("Si Kenia fuera una persona"),
  p("Tiene entre 35 y 40 años y ya no tiene nada que demostrar. Corrió lo suficiente como para saber que la disciplina no es gritar más fuerte. Viste funcional y sin logos gritones. Se levanta temprano y no lo publica. Escucha más de lo que habla. Cuando te ve mal, no te grita: te pregunta."),
  gap(60),
  table(["Cinco adjetivos que sí", "Cinco que definitivamente no"], [
    ["Atento", "Agresivo"],
    ["Claro", "Motivacional"],
    ["Firme", "Culposo"],
    ["Honesto", "Épico"],
    ["Cercano", "Condescendiente"],
  ], [4824, 4824]),
  gap(200),
  p("**Motivacional** está en la lista negra a propósito. Al corredor que ya está en el proceso no le falta motivación: le falta criterio. Kenia no motiva. Organiza, explica y protege. La motivación es consecuencia, no producto."),

  h2("La frase de posicionamiento frente a Strava"),
  quote("Si Strava es el muro donde publicas lo que hiciste, Kenia es la voz que decide contigo lo que vas a hacer."),
  brk(),
);

/* ---------- 11 ---------- */
add(
  h1("Tono de voz"),
  p("Medido en los cuatro ejes del Nielsen Norman Group, que es el estándar que usamos para que esto no dependa del gusto de quien escriba."),
  gap(60),
  table(["Eje", "Posición", "Por qué ahí"], [
    ["Divertido ↔ Serio", "4 — más serio", "Kenia habla de tu cuerpo, tu descanso y tu riesgo de lesión. El humor ahí resta autoridad. No es solemne: es seria como un buen fisioterapeuta, que se ríe contigo pero no de lo que te duele."],
    ["Casual ↔ Formal", "2 — casual", "Tuteo venezolano. Habla como el compañero de entreno que además sabe. Nada de «estimado usuario». Nunca «vos» ni «usted»."],
    ["Irreverente ↔ Respetuoso", "4 — respetuosa", "No se burla del usuario, ni de la competencia, ni de quien va lento. No puedes legitimar a alguien y burlarte de otro igual."],
    ["Entusiasta ↔ Neutro", "3 — equilibrio", "Entusiasta con el proceso, neutra con los datos. Nunca eufórica: la euforia envejece mal cuando el usuario falla."],
  ], [2200, 1600, 5848]),

  h2("El tono cambia según el momento"),
  table(["Momento", "Tono", "Cómo suena"], [
    ["Onboarding", "Cálido y concreto", "«Antes de armar tu plan necesito saber tres cosas. Toma dos minutos.»"],
    ["Antes de entrenar", "Pregunta, no ordena", "«¿Cómo amaneciste hoy?»"],
    ["Sesión asignada", "Directo, con el porqué", "«Hoy: 6 × 3 min fuerte. Trabajamos tu capacidad de sostener ritmo para el 10K.»"],
    ["Descanso recomendado", "Firme y explicativo", "«Hoy no entrenas. Dormiste cuatro horas y llevas dos semanas subiendo carga. Descansar también entrena.»"],
    ["Sesión bajo lluvia o calor", "Con propósito, nunca desafiante", "«Hoy toca rodaje bajo lluvia. Tu carrera de noviembre puede ser así, y quiero que no sea la primera vez.»"],
    ["Condición de riesgo", "Directa y sin drama", "«Hoy no. Está cayendo con rayos. Movemos la sesión al jueves.»"],
    ["Racha rota o dos semanas sin aparecer — el momento más importante de la marca", "Sin reproche, con opción", "«Tu plan te esperó. ¿Lo retomamos donde quedó o lo armamos de nuevo con tu semana de ahora?»"],
    ["Fallo de pago", "Neutro y práctico", "«No pudimos procesar tu pago. Tu plan sigue guardado. Actualiza el método cuando puedas.»"],
    ["Cancelación", "Agradecido, sin retención agresiva", "«Listo, cancelado. Tu historial queda guardado por si vuelves.»"],
  ], [2900, 2200, 4548]),

  h2("Palabras que no se usan"),
  table(["Prohibido", "Por qué"], [
    ["«sin excusas», «excusas»", "Es el enemigo declarado, dicho con la voz de la marca."],
    ["«fallaste», «perdiste tu racha», «otra vez no entrenaste»", "Culpa explícita."],
    ["«quemar grasa», «cuerpo perfecto», «bikini», «verano»", "Política de cuerpo."],
    ["«bestia», «modo bestia», «guerrero», «máquina»", "No por intensas: por vacías. Son clichés que cualquier box firma, y ponen a la marca a arengar en vez de a entrenar."],
    ["«deberías»", "Juicio moral disfrazado de consejo."],
    ["«solo tú puedes»", "Le pone toda la carga a quien ya se siente culpable."],
  ], [3400, 6248]),

  h2("Cómo exigir sin culpabilizar"),
  p("Es la pregunta que rompe a casi todas las marcas de fitness. La respuesta de Kenia cabe en una línea:"),
  quote("La exigencia se le pone al plan, nunca a la persona."),
  gap(200),
  p("El plan es duro cuando toca serlo: series exigentes, semanas de carga, objetivos que cuestan. Lo que nunca se juzga es a quien lo sigue. Tres reglas operativas:"),
  numbered("El sujeto de la frase dura es el entrenamiento, no el usuario. «Esta semana es exigente» sí; «tienes que esforzarte más» no."),
  numbered("Cada exigencia viene con su razón. Sin porqué, es una orden."),
  numbered("El descanso se comunica con la misma firmeza que el esfuerzo. Si «descansa» suena a disculpa, el usuario no descansa."),
  gap(120),
  panel("Las cuatro frases del manual actual se retiran", [
    "«El asfalto no miente» · «Tu único rival te mira en el espejo» · «La meta es solo el principio» · «Donde termina tu excusa, empieza tu carrera».",
    "No se retiran por intensas: la intensidad es correcta para este público. Se retiran porque **hablan como el héroe en lugar de hablarle al héroe**, y la cuarta además lo avergüenza usando la palabra que la marca declaró enemiga.",
    "Las nuevas las escribimos nosotros. Un solo ejemplo para fijar la altura: donde antes decía «donde termina tu excusa, empieza tu carrera», ahora diría **«tu semana cambia, tu plan también»**. La segunda dice lo que el producto hace. La primera decía lo contrario.",
  ]),
  brk(),
);

/* ---------- 12 ---------- */
add(
  h1("Identidad visual"),
  p("La identidad la hizo el diseñador y nuestro encargo es integrarla, no rehacerla. El diagnóstico es de desalineación con la estrategia, no de calidad."),

  h2("Lo que funciona y no se toca"),
  bullet("**El monograma-flecha.** Tiene carácter, dirección y velocidad. Es el activo más fuerte del paquete."),
  bullet("**El naranja sobre negro.** Potente y visible."),
  bullet("**El oficio de las aplicaciones.** Los pósters, el mural, la interfaz de reloj y los mapas de ruta están bien resueltos y muestran un sistema con recorrido."),

  h2("La jerarquía de color, ya decidida: manda el azul"),
  p("Ustedes movieron el peso de la paleta hacia el azul porque el naranja acercaba demasiado a Kenia al territorio de Strava. La decisión es buena y además encaja con el arquetipo: el mentor es azul —método, medición, calma, criterio— y el naranja pasa a ser el color del esfuerzo del héroe, el momento de correr. Que aparezca poco es lo que le da fuerza."),
  gap(60),
  table(["Color", "Código", "Función", "Dónde se usa"], [
    ["Cian", "#079CB5", "Primario. La marca y el método.", "Superficies de marca, fondos de datos, gráficos, interfaz"],
    ["Petróleo", "#00677E", "Profundidad.", "Fondos de sección, degradados con el cian, capas"],
    ["Negro", "#1B1B1B", "El suelo. La madrugada, el asfalto.", "Base del sistema. La marca sigue siendo oscura por defecto"],
    ["Naranja", "#FF5A00", "Acento de acción. El momento de correr.", "Botones, el kilómetro en curso, el estado activo. Poco y potente"],
    ["Blanco", "#F4F4F4", "Texto y aire.", "Tipografía sobre oscuro"],
  ], [1300, 1400, 3200, 3748]),
  gap(200),
  quote("El azul es la marca. El naranja es el momento."),
  gap(200),
  p("**Proporción:** naranja por debajo del 10 % de la superficie de cualquier pieza. Si una pieza es mayoritariamente naranja, o es el logo aislado o está mal."),
  gap(60),
  panel("Tres cosas que conviene saber antes de ejecutarlo", [
    "**No es un ajuste, es un cambio de identidad.** El logotipo hoy es naranja y todo el manual está construido sobre él. Si el azul pasa a primario, cambian el lockup, las aplicaciones y el ícono de tienda. La decisión es tuya y del diseñador, pero conviene saber el tamaño de lo que se aprueba.",
    "**Alejarse de Strava por el naranja no sirve si se cae en otro cliché.** Azul y naranja es de las combinaciones más repetidas en apps de deporte. Lo que diferencia no es el tono: es la proporción y el tratamiento. Antes de cerrarlo, poner las fichas de Strava, Runna, RunMotion, Nike Run Club y Garmin una al lado de otra y comprobar en qué se distingue Kenia.",
    "**Donde más se gana es en el ícono de la tienda.** En una pantalla llena de apps de fitness naranjas y rojas, un ícono azul con la K se ve antes. Es el argumento más fuerte a favor del cambio y vale la pena probarlo a 60 píxeles antes de decidir.",
  ], NARANJA),

  h2("Lo que le falta al manual"),
  p("Esto es feedback para el diseñador, no trabajo nuestro. Son las piezas que hoy no existen y que hacen falta antes de que la app salga:"),
  gap(60),
  table(["Pieza que falta", "Por qué importa"], [
    ["Tipografía", "Ni familia, ni pesos, ni jerarquía, ni tamaños. Es la ausencia más urgente para una app que sale en octubre."],
    ["Ícono de app", "Versión cuadrada, monograma aislado y prueba a 60 y a 1024 píxeles. La K es una flecha muy abstracta y a tamaño de ícono hay riesgo de que se lea otra cosa."],
    ["Área de resguardo y usos incorrectos", "Tamaño mínimo del logo y qué no se puede hacer con él."],
    ["Reglas de contraste", "Accesibilidad del texto sobre cada color de la paleta, sobre todo dentro de la interfaz."],
  ], [3200, 6448]),
  gap(200),
  panel("Dos cosas de fotografía a corregir antes de publicar", [
    "El póster de la maratón de Melbourne lleva el logo de un tercero: como lámina de portafolio se defiende, publicado sería un problema. Y hay fotografía de banco con ropa de marcas ajenas en la portada. Ninguna de las dos se usa en canales.",
  ]),
  brk(),
);

/* ---------- 13 ---------- */
add(
  h1("Brief de moodboard"),
  p("El moodboard lo armamos nosotros. Esto es el criterio, para poder decir «esto sí, esto no» sin discutir de gustos."),
  p("El principio es el mismo del arquetipo: el moodboard muestra **el mundo del héroe con la mirada del mentor**. Concentración en vez de furia, método en vez de arenga, cuerpo que trabaja en vez de cuerpo que posa. La prueba rápida para cada imagen: si esta foto la firmaría Nike sin dudar, probablemente es la voz del héroe y no la nuestra."),

  h2("Los siete territorios"),
  table(["#", "Territorio", "Qué entra"], [
    ["01", "El esfuerzo real", "Cara de trabajo y concentración: mandíbula suelta, mirada al frente, respiración. El corredor concentrado en el kilómetro 15, no el atleta gritando en cámara lenta."],
    ["02", "La hora azul", "5:30 de la mañana, luz baja, ciudad vacía, faroles encendidos. Es el momento real de casi todo corredor con trabajo y tiene menos ruido que el atardecer dorado."],
    ["03", "Caracas reconocible", "El Ávila, la Cota Mil, el Parque del Este. Runna se ve como Londres, RunMotion como París, Nike Run Club como ningún lugar. Kenia puede ser la única que se ve como aquí, y eso no lo copia nadie desde afuera."],
    ["04", "El dato como material visual", "Trazos de ruta sobre mapa, splits, zonas de ritmo, curvas de carga. Estética de instrumento de medición, no de infografía de redes."],
    ["05", "La recuperación y el regreso", "El estiramiento, las piernas en alto, el día libre. Y sobre todo el primer día de vuelta: atarse los cordones otra vez después de una semana perdida. La categoría fotografía el esfuerzo, casi nunca esto."],
    ["06", "La comunidad sin ranking", "Grupo corriendo junto, gente esperando en la salida, alguien acompañando en el último kilómetro. Nada de podio ni medallero."],
    ["07", "Las condiciones", "Correr bajo la lluvia y bajo el calor de Caracas. Camiseta empapada, vapor del asfalto, sol de mediodía. Entra por tu decisión de que Kenia prepare para cualquier circunstancia."],
  ], [500, 2500, 6648]),
  gap(200),
  panel("El riesgo del territorio 7", [
    "Es el que más fácil se desliza hacia «sin excusas». La categoría fotografía la adversidad como épica: rictus, grito, desafío al cielo, «yo salí y tú no». Kenia la fotografía como preparación: alguien concentrado que ya sabe cómo se siente esto.",
    "**La prueba para cada imagen:** ¿celebra al que salió, o muestra a alguien preparado? Si celebra, está juzgando al que no salió.",
  ], NARANJA),

  h2("Qué evitar en todas las piezas"),
  bullet("Cara de grito, puño al aire y «modo bestia»: es voz de héroe, no de mentor."),
  bullet("Ropa con logos de terceros visibles."),
  bullet("Amanecer genérico de banco de imágenes sin lugar: podría ser cualquier marca del mundo."),
  bullet("Antes y después, y solo cuerpos atléticos y delgados."),
  bullet("Azul frío saturado con naranja tipo videojuego: es el cliché visual de las apps deportivas."),

  h2("Criterio de tipografía para el diseñador"),
  numbered("**Números excelentes.** Kenia vive de cifras: 5:53/km, 21,097 km, 6 × 3 min. Una familia con números mal resueltos hunde toda la interfaz. Hay que pedir cifras tabulares."),
  numbered("**Diálogo con el logo.** La «enia» es geométrica redondeada; la familia de texto tiene que convivir con eso sin imitarlo."),
  numbered("**Legible a tamaño pequeño, en movimiento, con sol y con sudor.** Es el contexto real de uso."),
  numbered("**Al menos cuatro pesos.** Una app con planes, datos y explicaciones necesita jerarquía."),
  brk(),
);

/* ---------- 14 ---------- */
add(
  h1("Lo que proponemos reforzar"),
  p("La plataforma está completa. Estas son ocho cosas que hoy no existen y que creemos que suman más de lo que cuestan. Cinco las podemos entregar nosotros; tres dependen de una decisión tuya."),
  gap(60),
  table(["#", "Propuesta", "De quién", "Qué es"], [
    ["1", "Pirámide de mensajes", "Nosotros", "Un mensaje maestro, tres pilares y las pruebas de cada uno. Es lo que nos falta para escribir campaña y contenido sin volver a discutir qué se dice primero. Los pilares serían decide contigo, te explica por qué y te cuida la carga."],
    ["2", "Prueba social desde ya", "Tú", "La app lanza sin una sola reseña, sin testimonios y sin casos, con un producto que promete adaptarse y prevenir lesiones. Es el activo que más falta y hay una fuente gratis: los doce entrevistados y la gente del box. Proponemos grabar testimonios el 13 y tener diez reseñas el día uno en tienda."],
    ["3", "Tu cara", "Tú", "El método es de autor y el autor no aparece por ningún lado. Con un mes de contenido sin app por delante, la cara del que diseñó el método es lo más creíble que tenemos. No hace falta que salgas en el producto: basta con una serie corta explicando el método."],
    ["4", "Cerrar «Kenia Sport»", "Nosotros", "El manual se llama «Marca Oficial Kenia Sport» y el descriptor decidido es «Kenia · Entrenamiento de running». Nadie decidió formalmente que «Kenia Sport» desaparece. Hay que decidirlo antes de que aparezca en un handle, un dominio o la ficha de tienda."],
    ["5", "Una métrica para la marca", "Nosotros", "Tenemos métricas de negocio y una de impacto, pero nada mide si la marca está funcionando. Dos preguntas a los 30 días lo resuelven: ¿te sentiste juzgado alguna vez por la app? y ¿cómo se lo explicarías a un amigo? La segunda dice, con sus palabras, si el posicionamiento llegó."],
    ["6", "Descargo médico", "Tú", "Estamos dando planes de entrenamiento a desconocidos y prometiendo prevención de lesiones, en un público donde tres de doce ya se lesionaron. Hace falta un descargo visible y una regla de qué hace la app ante una molestia declarada. Es riesgo real, no formalidad."],
    ["7", "Salir del box antes de escalar pauta", "Nosotros", "Las doce entrevistas son de tu red. No dicen nada sobre alguien que no te conoce, que es exactamente a quien va a llegar la pauta. Ocho conversaciones con desconocidos antes de subir presupuesto."],
    ["8", "Cerrar los datos de las personas", "Nosotros", "Los hábitos, objetivos y precios de las cinco personas están verificados. La edad, la ocupación y sobre todo qué cuentas siguen son hipótesis nuestras. Cuatro preguntas en el mismo grupo de WhatsApp lo cierran, y convierten la segmentación de pauta en dato propio en vez de intereses genéricos de Meta."],
  ], [500, 2400, 1200, 5548]),
  gap(200),
  panel("Si solo se puede hacer una", [
    "La número 2. Una marca nueva, en una categoría donde nadie te conoce, con un precio en dólares y una promesa de salud, se juega el primer mes en si alguien más lo dice por ti. Y el 13 de septiembre vamos a tener a la gente del box delante de una cámara.",
  ]),
  brk(),
);

/* ---------- 15 ---------- */
add(
  h1("Decisiones pendientes"),
  p("Todo lo demás está cerrado. Esto es lo que bloquea, ordenado por quién lo desbloquea."),

  h2("Ninro"),
  table(["Qué", "Por qué importa"], [
    ["Registro de marca, dominio y handles", "«Kenia» es un topónimo y conviene verificarlo con un abogado antes de invertir en identidad. Y hay que ver qué handles están libres."],
    ["Meta de negocio a tres años", "Sin horizonte largo, la visión no es creíble. Tenemos el punto de partida: 40 suscriptores a los 90 días. Falta a dónde va eso."],
    ["El piloto de Gabriel", "Hablar con tu conocido el del maratón y decidir si entra a probar Elite durante su bloque."],
    ["Prueba social y tu cara", "Las propuestas 2 y 3 de la sección anterior. Las dos se ejecutan el 13 o no se ejecutan."],
    ["Protocolo ante trastornos alimentarios", "Está en la política de cuerpo y falta definir el cómo, con Kurt."],
  ], [3200, 6448]),

  h2("Kurt"),
  table(["Qué", "Por qué importa"], [
    ["Qué de los tres planes entra en la v1", "Es la más importante de todas. Si vendemos cuerpo digital en octubre y llega en marzo, la marca se rompe en el primer mes. Y ahora hay un comprador de Elite que lo va a notar en la primera semana."],
    ["Cómo paga exactamente un venezolano", "Puede tumbar el negocio entero. Si alguien no puede pagar en tres toques, las 1.200 descargas comprometidas no valen nada."],
    ["Si la app emite eventos de analítica", "Hay comprometida conversión del 3 %, retención a 30 días del 25 % y costo por adquisición por debajo de 25 dólares. Sin medición dentro de la app, nada de eso se puede reportar."],
    ["Qué hace el plan el día después de un fallo", "Es la promesa central de la marca. Si la app no recalcula sola, hay que cambiar la promesa o construirlo."],
    ["Si el plan puede programar clima, no solo esquivarlo", "Es tu planteamiento del 8, convertido en pregunta técnica."],
    ["A qué volumen semanal aguanta el generador", "Si se diseñó para 30 o 40 kilómetros semanales, un corredor de 85 ve el techo enseguida."],
    ["Unificar los nombres de plan", "De KENIA STARTER a Kenia Starter, dentro de la app y en la ficha de tienda."],
  ], [3200, 6448]),

  h2("Diseñador"),
  bullet("Tipografía, área de resguardo, usos incorrectos e ícono de app, con prueba a 60 píxeles."),
  bullet("Aplicar la jerarquía de color azul-primero al logotipo y al ícono."),
  bullet("Sustituir la fotografía que lleva marcas de terceros."),

  h2("Organic Club"),
  bullet("Las cuatro frases de campaña que reemplazan a las retiradas."),
  bullet("El tagline. «Master your pace» se retira: está en inglés, es evocativo y no explica, y cuando el nombre no explica el tagline tiene que hacerlo."),
  bullet("El nombre de la comunidad. El criterio ya está fijado: tiene que ser algo a lo que te unes, no un título que haya que ganarse."),
  bullet("El moodboard, con los siete territorios de la sección 13."),
  brk(),
);

/* ---------- 16 ---------- */
add(
  h1("Calendario"),
  table(["Cuándo", "Qué"], [
    ["Antes del 13 de septiembre", "Entrega de branding, análisis de marketing y estrategia de contenido."],
    ["13 de septiembre", "Primera grabación con Ninro. Es también la única ventana para grabar testimonios de la gente del box."],
    ["15 al 20 de septiembre", "Viaje de Ninro a Buenos Aires."],
    ["4 de octubre", "Carrera Caracas Rock. Cae antes del lanzamiento."],
    ["Mediados de octubre", "Lanzamiento de la app en tiendas."],
    ["Noviembre", "Media Maratón de Hipereventos. Es el ancla de la campaña de lanzamiento."],
  ], [2600, 7048]),
  gap(240),
  panel("Dos cosas del calendario que hay que diseñar, no descubrir", [
    "**El mes sin app.** Entre la grabación del 13 y el lanzamiento de mediados de octubre hay un mes en el que se puede publicar contenido pero no se puede descargar nada. Ese mes tiene una función: construir lista de espera y llegar al día uno con gente que descargue, que es lo que decide el arranque en tiendas. El contenido del 13 hay que guionizarlo para las dos fases.",
    "**Caracas Rock es el 4 de octubre y la app sale después.** Cuatro de los doce entrevistados corren esa carrera y se la pierden. Por eso la campaña se ancla a la media maratón de noviembre, que sí encaja. Caracas Rock igual sirve: es el evento donde está el público, cuatro semanas antes. Es el sitio para grabar, captar contactos y construir la lista.",
  ], NARANJA),
  gap(500),
  new Paragraph({
    spacing: { before: 300, after: 80 },
    border: { top: { style: BorderStyle.SINGLE, size: 18, color: NEGRO, space: 10 } },
    children: [new TextRun({ text: "Organic Club", font: SERIF, size: 24, bold: true, color: NEGRO })],
  }),
  new Paragraph({
    children: [new TextRun({ text: "Liam Libre · Alonzo García   |   Plataforma de marca Kenia · versión 1.0 · 8 de septiembre de 2026", font: SANS, size: 19, color: GRIS })],
  }),
);

/* ================= BUILD ================= */
const doc = new Document({
  creator: "Organic Club",
  title: "Kenia — Plataforma de marca",
  description: "Plataforma de marca completa de Kenia. Preparada por Organic Club para Ninro Libre.",
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
      children: [new TextRun({ text: "Organic Club  ·  Plataforma de marca Kenia  ·  v1.0  ·  8 de septiembre de 2026",
                               font: SANS, size: 15, color: GRIS_CLARO })],
    })] }) },
    children: C,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(process.argv[2], buf);
  console.log("escrito:", process.argv[2], (buf.length / 1024).toFixed(0) + " KB");
});
