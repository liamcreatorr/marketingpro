/* Genera el contrato de prestación de servicios Organic Club × Ninro Libre (Kenia).
   Lee los datos de las partes de datos-partes.json si existe (no va al repositorio);
   si no, usa datos-partes.ejemplo.json y deja líneas en blanco para completar a mano.
   Uso: node contrato/generar-contrato.js ["salida.docx"]                          */

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Footer,
  WidthType, ShadingType, BorderStyle, AlignmentType, PageBreak, PageNumber,
} = require('docx');

/* ---------------- datos ---------------- */
const dir = __dirname;
const real = path.join(dir, 'datos-partes.json');
const usaReales = fs.existsSync(real);
const D = JSON.parse(fs.readFileSync(usaReales ? real : path.join(dir, 'datos-partes.ejemplo.json'), 'utf8'));
const salida = process.argv[2] || path.join(dir,
  usaReales ? 'Contrato Organic Club - Kenia-FIRMA.docx' : 'Contrato Organic Club - Kenia (plantilla).docx');

const RAYA = '____________________';
const v = (x, raya = RAYA) => (x === undefined || x === null || String(x).trim() === '') ? raya : String(x);
const cli = D.cliente, ag = D.agencia, K = D.condiciones;
const rl = cli.representante_legal, kt = cli.representante_autorizado;
const [liam, alonzo] = ag.socios;

const NUM = ['', 'un (1)', 'dos (2)', 'tres (3)', 'cuatro (4)', 'cinco (5)', 'seis (6)', 'siete (7)',
  'ocho (8)', 'nueve (9)', 'diez (10)', 'once (11)', 'doce (12)', 'trece (13)', 'catorce (14)', 'quince (15)'];
const n = x => NUM[x] || String(x);
const dias = x => x === 1 ? 'un (1) día' : n(x) + ' días';

/* ---------------- estilo (casa Organic Club, acento Kenia) ---------------- */
const CIAN = "079CB5", PETROLEO = "00677E", NEGRO = "1B1B1B", GRIS = "3C3C3C",
      GRIS_CLARO = "6E7F82", FONDO = "EAF2F4", FONDO_2 = "F2F6F7", LINEA = "C7D9DD";
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
  alignment: o.align ?? AlignmentType.JUSTIFIED,
  spacing: { after: o.after ?? 140, line: 300 },
  indent: o.indent ? { left: o.indent } : undefined,
  children: runs(t, { font: SANS, size: o.size ?? 21, color: o.color || NEGRO }),
});

/* 6.1, 6.2… con sangría francesa */
const item = (num, t) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after: 110, line: 300 },
  indent: { left: 600, hanging: 600 },
  children: [new TextRun({ text: num + "\t", font: SANS, size: 21, bold: true, color: PETROLEO }),
             ...runs(t, { font: SANS, size: 21, color: NEGRO })],
  tabStops: [{ type: "left", position: 600 }],
});

const sub = t => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after: 80, line: 290 },
  indent: { left: 960, hanging: 300 },
  children: [new TextRun({ text: "–\t", font: SANS, size: 21, color: CIAN }),
             ...runs(t, { font: SANS, size: 21, color: NEGRO })],
  tabStops: [{ type: "left", position: 960 }],
});

const ORD = ['PRIMERA', 'SEGUNDA', 'TERCERA', 'CUARTA', 'QUINTA', 'SEXTA', 'SÉPTIMA', 'OCTAVA',
  'NOVENA', 'DÉCIMA', 'DÉCIMA PRIMERA', 'DÉCIMA SEGUNDA', 'DÉCIMA TERCERA', 'DÉCIMA CUARTA',
  'DÉCIMA QUINTA', 'DÉCIMA SEXTA', 'DÉCIMA SÉPTIMA', 'DÉCIMA OCTAVA', 'DÉCIMA NOVENA',
  'VIGÉSIMA', 'VIGÉSIMA PRIMERA', 'VIGÉSIMA SEGUNDA', 'VIGÉSIMA TERCERA', 'VIGÉSIMA CUARTA',
  'VIGÉSIMA QUINTA'];
let cN = 0;
function clausula(titulo) {
  const ord = ORD[cN]; cN += 1;
  return new Paragraph({
    keepNext: true,
    spacing: { before: 300, after: 140 },
    children: [
      new TextRun({ text: "CLÁUSULA " + ord, font: SANS, size: 18, bold: true, color: CIAN, characterSpacing: 24 }),
      new TextRun({ text: "   " + titulo, font: SERIF, size: 23, bold: true, color: NEGRO }),
    ],
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINEA, space: 4 } },
  });
}

function titulo(t, sub) {
  const out = [new Paragraph({
    spacing: { before: 0, after: sub ? 60 : 220 },
    children: [new TextRun({ text: t, font: SERIF, size: 30, bold: true, color: NEGRO })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: CIAN, space: 8 } },
  })];
  if (sub) out.push(new Paragraph({ spacing: { after: 220 }, children: runs(sub, { font: SANS, size: 19, color: GRIS_CLARO }) }));
  return out;
}

function table(headers, rows, widths, opt = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  if (total !== W) throw new Error("anchos suman " + total + ", deben sumar " + W);
  const head = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: PETROLEO, color: "auto" },
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({
        text: String(h).toUpperCase(), font: SANS, size: 16, bold: true, color: "FFFFFF", characterSpacing: 20,
      })] })],
    })),
  });
  const body = rows.map((r, ri) => new TableRow({
    cantSplit: true,
    children: r.map((c, i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: (opt.total && ri === rows.length - 1) ? FONDO : (ri % 2 ? FONDO_2 : "FFFFFF"), color: "auto" },
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      children: [new Paragraph({ spacing: { after: 0, line: 264 }, children: runs(c, { font: SANS, size: 19, color: NEGRO }) })],
    })),
  }));
  const L = { style: BorderStyle.SINGLE, size: 4, color: LINEA };
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: widths,
    borders: { top: L, bottom: L, left: L, right: L,
               insideHorizontal: { ...L, size: 2 }, insideVertical: { ...L, size: 2 } },
    rows: [head, ...body],
  });
}

function recuadro(kids) {
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    borders: { top: NONE, bottom: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE,
               left: { style: BorderStyle.SINGLE, size: 18, color: CIAN } },
    rows: [new TableRow({ children: [new TableCell({
      width: { size: W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: FONDO, color: "auto" },
      margins: { top: 160, bottom: 160, left: 260, right: 220 },
      children: kids,
    })] })],
  });
}

const gap = (h = 200) => new Paragraph({ spacing: { after: h }, children: [] });
const brk = () => new Paragraph({ children: [new PageBreak()] });

/* ---------------- comparecencia ---------------- */
function persona(s) {
  return `**${s.nombre.toUpperCase()}**, ${s.nacionalidad ? 'de nacionalidad ' + s.nacionalidad + ', ' : ''}mayor de edad, ` +
    `titular de la cédula de identidad N.º ${v(s.cedula)}, ` +
    `con domicilio en ${v(s.domicilio, '________________________________')}` +
    (s.telefono ? `, teléfono ${s.telefono}` : '') +
    (s.correo ? `, correo electrónico ${s.correo}` : '');
}

const empresa = `**${cli.razon_social.toUpperCase()}**, sociedad de responsabilidad limitada constituida conforme a las leyes del ${v(cli.jurisdiccion)}, ` +
  `identificada con el ${cli.tipo_id_fiscal} N.º ${v(cli.id_fiscal)}, con domicilio en ${v(cli.domicilio, '________________________________')}` +
  `${cli.correo ? ' y correo electrónico ' + cli.correo : ''}, titular de la aplicación móvil denominada **«Kenia»**`;

/* ---------------- calendario de pagos ---------------- */
const dl = K.dia_limite_pago;
const pagos = [
  ["1", "Septiembre 2026", "Mes de fundador", "US$ 1.000,00", `${dl} de septiembre de 2026`, v(K.estado_pago_septiembre, '☐ Pagado   ☐ Pendiente')],
  ["2", "Octubre 2026", "Mensualidad ordinaria", "US$ 1.200,00", `${dl} de octubre de 2026`, "☐ Pagado   ☐ Pendiente"],
  ["3", "Noviembre 2026", "Mensualidad ordinaria", "US$ 1.200,00", `${dl} de noviembre de 2026`, "☐ Pagado   ☐ Pendiente"],
  ["4", "Diciembre 2026", "Mensualidad ordinaria", "US$ 1.200,00", `${dl} de diciembre de 2026`, "☐ Pagado   ☐ Pendiente"],
  ["", "**Total del plazo inicial**", "", "**US$ 4.600,00**", "", ""],
];

/* ================= CONTENIDO ================= */
const C = [];
const add = (...xs) => xs.forEach(x => Array.isArray(x) ? C.push(...x) : C.push(x));

/* ---------- PORTADA ---------- */
add(
  gap(900),
  new Paragraph({ spacing: { after: 90 }, children: [new TextRun({ text: "ORGANIC CLUB  ·  DOCUMENTO CONTRACTUAL", font: SANS, size: 16, bold: true, color: CIAN, characterSpacing: 30 })] }),
  new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Contrato de prestación", font: SERIF, size: 56, bold: true, color: NEGRO })] }),
  new Paragraph({ spacing: { after: 240 }, children: [new TextRun({ text: "de servicios profesionales", font: SERIF, size: 56, bold: true, color: NEGRO })] }),
  new Paragraph({ spacing: { after: 420 }, children: [new TextRun({ text: "Estrategia de marca, contenido y lanzamiento de la aplicación Kenia", font: SERIF, size: 28, color: GRIS })] }),
  table(["", ""], [
    ["**El cliente**", `${cli.razon_social} — titular de la aplicación Kenia`],
    ["**Representante legal**", `${rl.nombre} — fundador y CEO de Kenia`],
    ["**Representante autorizado**", `${kt.nombre} — desarrollador y contacto técnico`],
    ["**La agencia**", `${ag.nombre_comercial} — ${liam.nombre} y ${alonzo.nombre}`],
    ["**Inicio de los servicios**", K.inicio_servicios],
    ["**Inicio de la publicación**", K.inicio_publicacion],
    ["**Plazo inicial**", `Hasta el ${K.fin_plazo_inicial}, renovable mes a mes`],
    ["**Honorarios**", "US$ 1.000 el primer mes · US$ 1.200 mensuales desde el segundo"],
    ["**Lugar y fecha de firma**", `${K.ciudad_firma}, ${K.fecha_firma}`],
  ], [3000, 6648]),
  gap(360),
  p("Este contrato formaliza la relación de trabajo que comenzó con la **Propuesta de Estrategia y Gestión de Contenido del 30 de agosto de 2026**, recoge el cronograma real del proyecto y fija el alcance, los entregables, los honorarios y las condiciones de pago que regirán a partir de su firma.", { color: GRIS }),
  brk(),
);

/* ---------- COMPARECENCIA ---------- */
add(
  ...titulo("Contrato de prestación de servicios profesionales"),
  p(`Entre, por una parte, ${empresa}, representada en este acto por su representante legal, ${persona(rl)}, y por su representante autorizado, ${persona(kt)}; sociedad que en lo sucesivo y a los efectos de este contrato se denominará **EL CLIENTE**;`),
  p(`y por la otra, ${persona(liam)}; y ${persona(alonzo)}; quienes actúan conjuntamente bajo la denominación comercial **${ag.nombre_comercial.toUpperCase()}**${ag.rif ? ', RIF N.º ' + ag.rif : ''}, y que en lo sucesivo se denominarán **LA AGENCIA**;`),
  p("ambas partes, reconociéndose mutuamente la capacidad legal necesaria para contratar y obligarse, han convenido celebrar el presente **Contrato de Prestación de Servicios Profesionales**, que se regirá por las cláusulas siguientes:"),
);

/* PRIMERA */
add(clausula("Objeto"),
  p("LA AGENCIA se obliga a prestar a EL CLIENTE servicios profesionales de **estrategia de marca, análisis de marketing, estrategia y gestión de contenido para redes sociales y campaña de lanzamiento** de la aplicación móvil de entrenamiento de running denominada **«Kenia»**, en los términos, plazos y condiciones establecidos en este contrato y en sus anexos. EL CLIENTE se obliga a pagar los honorarios convenidos y a cumplir las obligaciones que este contrato le asigna."),
);

/* SEGUNDA */
add(clausula("Antecedentes y documentos que integran el contrato"),
  item("2.1", "Las partes reconocen que LA AGENCIA viene prestando los servicios objeto de este contrato desde el **" + K.inicio_servicios + "**, sobre la base de la propuesta comercial del 30 de agosto de 2026, y que el presente documento formaliza esa relación con efectos desde dicha fecha."),
  item("2.2", "Forman parte integrante de este contrato: el **Anexo A** (entregables), el **Anexo B** (calendario de pagos), el **Anexo C** (metas de referencia) y el **Anexo D** (directorio del proyecto)."),
  item("2.3", "En caso de contradicción entre la propuesta del 30 de agosto de 2026 y este contrato, **prevalece este contrato**. En particular, quedan sin efecto el cronograma de la propuesta que situaba el lanzamiento el 14 de septiembre y la inclusión del ASO en su Fase 1, por las razones recogidas en las cláusulas cuarta y quinta."),
);

/* TERCERA */
add(clausula("Alcance de los servicios"),
  p("Los servicios comprenden:"),
  item("3.1", "**Plataforma de marca:** propósito, misión, visión, valores, público objetivo, buyer personas, beneficios, arquetipo, personalidad, tono de voz, brief de moodboard, significado del nombre y descriptor fijo de la marca."),
  item("3.2", "**Análisis de marketing:** PESTEL, tendencias, análisis del servicio, socios estratégicos, pricing, benchmark de competencia, matriz FODA con estrategias cruzadas y objetivos SMART."),
  item("3.3", "**Estrategia de contenido:** embudo, pilares, formatos, plan de canales, calendario editorial, relato de marca, guiones de grabación, tagline, frases de campaña y nombre de comunidad."),
  item("3.4", "**Campaña de lanzamiento:** plan de pre-lanzamiento y lista de espera, video de lanzamiento, piezas de la semana de lanzamiento y las acciones de captación que las partes acuerden."),
  item("3.5", `**Gestión mensual de contenido:** planificación, guionización, dirección de grabación, edición y publicación en los canales definidos en la estrategia de contenido, con un volumen de **${K.volumen_mensual}**. El volumen podrá redistribuirse entre formatos de mutuo acuerdo sin alterar los honorarios.`),
  item("3.6", "**Gestión de campañas pagadas en Meta** (Instagram y Facebook): configuración, segmentación, creatividades y optimización, con el presupuesto de pauta que EL CLIENTE paga directamente a la plataforma."),
  item("3.7", "**Informe mensual** de resultados y una **reunión mensual** de seguimiento con EL CLIENTE."),
);

/* CUARTA */
add(clausula("Servicios excluidos"),
  p("No forman parte de los honorarios de este contrato, y se cotizarán por separado si EL CLIENTE los solicita:"),
  item("4.1", "**Optimización para tiendas de aplicaciones (ASO)**: título, subtítulo, descripción, capturas, video de vista previa e investigación de palabras clave en App Store y Google Play. Las partes dejan constancia de que la propuesta del 30 de agosto de 2026 lo incluía en su Fase 1 y de que, por acuerdo posterior entre ellas, **queda excluido** del precio convenido. El descriptor fijo de la marca (cláusula 3.1) sí está incluido, aunque pueda emplearse en la ficha de tienda."),
  item("4.2", "El copy interno de la aplicación y el guion de *onboarding*."),
  item("4.3", "El rediseño de la identidad visual, el logotipo, la tipografía y el ícono de la aplicación, elaborados por un tercero. LA AGENCIA integra esa identidad y aporta observaciones."),
  item("4.4", "El desarrollo, mantenimiento e instrumentación técnica de la aplicación, incluida la medición de eventos de analítica dentro de ella."),
  item("4.5", "El presupuesto de pauta publicitaria, que EL CLIENTE paga directamente a Meta u otras plataformas."),
  item("4.6", "La atención al cliente y el soporte técnico a los usuarios de la aplicación."),
  item("4.7", "Cualquier servicio no descrito expresamente en la cláusula tercera."),
);

/* QUINTA */
add(clausula("Vigencia y cronograma"),
  item("5.1", `Este contrato tiene un **plazo inicial desde el ${K.inicio_servicios} hasta el ${K.fin_plazo_inicial}**. Vencido ese plazo, se renovará automáticamente por períodos sucesivos de un (1) mes, salvo que cualquiera de las partes notifique su voluntad de no renovarlo con al menos ${n(K.dias_preaviso_terminacion)} días de anticipación.`),
  item("5.2", "Las partes reconocen que la fecha de lanzamiento de la aplicación en tiendas es **mediados de octubre de 2026**, y que la fija EL CLIENTE en función del desarrollo técnico. La campaña de lanzamiento se ejecutará sobre la fecha que EL CLIENTE confirme por escrito."),
  item("5.3", `La **publicación de contenido** en los canales de la marca comienza el **${K.inicio_publicacion}**. Desde esa fecha se presta la gestión mensual de contenido descrita en la cláusula 3.5.`),
  item("5.4", "Los entregables y su periodicidad se detallan en el **Anexo A**. Los plazos que dependen de insumos de EL CLIENTE se rigen además por la cláusula décima tercera."),
);

/* SEXTA */
add(clausula("Honorarios"),
  item("6.1", "EL CLIENTE pagará a LA AGENCIA por los servicios de este contrato los siguientes honorarios mensuales:"),
  sub("**Primer mes (septiembre de 2026):** un mil dólares de los Estados Unidos de América (**US$ 1.000,00**), en condición de *mes de fundador*, en lugar de los US$ 1.200,00 de la propuesta original."),
  sub("**Desde el segundo mes (octubre de 2026) en adelante:** un mil doscientos dólares de los Estados Unidos de América (**US$ 1.200,00**) mensuales."),
  item("6.2", "El monto total correspondiente al plazo inicial es de **US$ 4.600,00**, conforme al calendario del **Anexo B**."),
  item("6.3", "Los honorarios no incluyen el presupuesto de pauta ni los gastos descritos en la cláusula novena."),
  item("6.4", "Los honorarios podrán revisarse a partir de la primera renovación posterior al plazo inicial, previo acuerdo escrito de ambas partes."),
);

/* SÉPTIMA */
add(clausula("Forma, fecha y lugar de pago"),
  item("7.1", `Los honorarios se pagan **por mes adelantado**, a más tardar el **día ${dl} de cada mes** que se va a trabajar. Si ese día no es hábil, el pago vence el primer día hábil siguiente.`),
  item("7.2", `El pago se hará mediante **${v(K.metodo_pago, '________________________________')}**${K.datos_cuenta_pago ? ', a la cuenta: ' + K.datos_cuenta_pago : ', a la cuenta que LA AGENCIA indique por escrito'}. Cualquier cambio de método o de cuenta solo será válido si LA AGENCIA lo comunica por escrito.`),
  item("7.3", "La moneda de cuenta del contrato es el **dólar de los Estados Unidos de América**. Si las partes acuerdan un pago en bolívares, se calculará al tipo de cambio oficial publicado por el Banco Central de Venezuela vigente el día del pago efectivo."),
  item("7.4", "Las comisiones bancarias o de plataforma que genere el pago corren por cuenta de EL CLIENTE, de modo que LA AGENCIA reciba el monto íntegro."),
  item("7.5", "EL CLIENTE enviará el comprobante de cada pago al canal de comunicación del proyecto. LA AGENCIA emitirá el recibo correspondiente dentro de los tres (3) días hábiles siguientes."),
);

/* OCTAVA */
add(clausula("Retraso en el pago"),
  item("8.1", `Si un pago no se recibe en la fecha de la cláusula 7.1, LA AGENCIA lo notificará por escrito. Transcurridos **${n(K.dias_gracia_suspension)} días continuos** desde el vencimiento sin que el pago se haya efectuado, LA AGENCIA podrá **suspender la prestación de los servicios** hasta que se regularice, sin que ello constituya incumplimiento de su parte.`),
  item("8.2", "Durante la suspensión, los plazos de los entregables de LA AGENCIA se desplazan por el mismo número de días que dure el retraso. La suspensión no extingue la obligación de pago."),
  item("8.3", "Si el retraso supera los treinta (30) días continuos, LA AGENCIA podrá dar por terminado el contrato conforme a la cláusula vigésima primera."),
);

/* NOVENA */
add(clausula("Gastos no incluidos"),
  item("9.1", "Correrán por cuenta de EL CLIENTE, **siempre que los haya aprobado por escrito antes de incurrirse**: alquiler de locaciones o equipos especiales, traslados fuera de la ciudad de Caracas, honorarios de talento o terceros, música o material con licencia de pago, inscripciones a eventos y herramientas de pago que se usen exclusivamente para el proyecto."),
  item("9.2", "LA AGENCIA presentará el soporte de cada gasto aprobado, y EL CLIENTE lo reembolsará dentro de los cinco (5) días hábiles siguientes."),
);

/* DÉCIMA */
add(clausula("Obligaciones de LA AGENCIA"),
  item("10.1", "Prestar los servicios con diligencia profesional, conforme a la plataforma de marca y a la estrategia de contenido aprobadas."),
  item("10.2", "Cumplir los plazos del Anexo A, salvo los supuestos de la cláusula décima tercera."),
  item("10.3", "Someter a aprobación de EL CLIENTE las piezas antes de su publicación, conforme a la cláusula décima segunda."),
  item("10.4", "Informar de manera oportuna cualquier riesgo que pueda afectar los plazos o las metas de referencia."),
  item("10.5", "Usar los accesos y la información de EL CLIENTE solo para los fines de este contrato."),
);

/* UNDÉCIMA */
add(clausula("Obligaciones de EL CLIENTE"),
  item("11.1", "Pagar los honorarios en la forma y fechas convenidas."),
  item("11.2", "Entregar oportunamente la información, los materiales, las aprobaciones y los accesos necesarios para la prestación de los servicios, incluidos los accesos de administración a las cuentas de redes sociales y al administrador comercial de Meta."),
  item("11.3", "Estar disponible para las grabaciones acordadas y facilitar el acceso a las personas y lugares que participen en ellas."),
  item("11.4", `Garantizar que su representante autorizado, **${kt.nombre}**, responda las consultas técnicas que condicionen el contenido —entre ellas, la fecha exacta de lanzamiento, la existencia y duración de la prueba gratuita, los métodos de pago disponibles para el usuario venezolano y el funcionamiento del plan tras una sesión no realizada— en un plazo máximo de **cinco (5) días hábiles** desde que LA AGENCIA las solicite por escrito.`),
  item("11.5", "Pagar directamente a Meta u otras plataformas el presupuesto de pauta acordado."),
  item("11.6", "Responder por la **veracidad de las características, funciones y beneficios de la aplicación** que comunique a LA AGENCIA para su difusión, en especial los relacionados con salud, adaptación del plan y prevención de lesiones."),
);

/* DUODÉCIMA */
add(clausula("Aprobaciones y rondas de revisión"),
  item("12.1", `EL CLIENTE aprobará o hará observaciones a cada pieza dentro de los **${dias(K.dias_habiles_aprobacion)} hábiles** siguientes a su envío. Si no responde en ese plazo, LA AGENCIA podrá reiterar el envío; si transcurre un (1) día hábil más sin respuesta, la pieza se tendrá por aprobada.`),
  item("12.2", `Cada pieza incluye hasta **${n(K.rondas_revision)} rondas de revisión**. Los cambios adicionales, o los que se pidan sobre una pieza ya aprobada o publicada, se considerarán trabajo adicional y podrán cotizarse por separado.`),
  item("12.3", "Se considera aprobación válida la que se dé por escrito por el canal de comunicación del proyecto, incluido el mensaje de WhatsApp o el correo electrónico."),
);

/* DÉCIMA TERCERA */
add(clausula("Dependencias y desplazamiento de plazos"),
  item("13.1", "Cuando un entregable de LA AGENCIA dependa de un insumo, una aprobación o una respuesta de EL CLIENTE o de su contacto técnico, y estos lleguen fuera del plazo de la cláusula 11.4 o del que las partes hayan acordado, **el plazo del entregable se desplazará por el mismo número de días del retraso**, sin que ello constituya incumplimiento de LA AGENCIA."),
  item("13.2", "Los retrasos atribuibles a EL CLIENTE no suspenden ni reducen la obligación de pago de los honorarios."),
  item("13.3", "LA AGENCIA notificará por escrito cada dependencia vencida en cuanto se produzca, indicando el entregable afectado."),
);

/* DÉCIMA CUARTA */
add(clausula("Metas de referencia y naturaleza de la obligación"),
  item("14.1", "Las metas del **Anexo C** son **objetivos de referencia** para orientar y evaluar el trabajo. La obligación de LA AGENCIA es **de medios y no de resultado**: se compromete a ejecutar con diligencia profesional la estrategia acordada, no a garantizar un número determinado de descargas, suscriptores o ventas."),
  item("14.2", "Las partes reconocen que esas metas dependen también de factores ajenos a LA AGENCIA: el funcionamiento y la fecha de publicación de la aplicación, los métodos de pago disponibles para el usuario, el precio, la existencia de prueba gratuita, el presupuesto de pauta, las políticas de Meta, Apple y Google, y la instrumentación de analítica dentro de la aplicación."),
  item("14.3", "Las metas a 90 días se cuentan **desde la fecha efectiva de publicación de la aplicación en tiendas**. Las partes podrán revisarlas de mutuo acuerdo cuando cambie alguna de las condiciones de la cláusula 14.2."),
);

/* DÉCIMA QUINTA */
add(clausula("Propiedad intelectual"),
  item("15.1", "Una vez pagados íntegramente los honorarios del período correspondiente, los derechos patrimoniales sobre los entregables finales y aprobados de ese período **pasan a ser de EL CLIENTE**, quien podrá usarlos, modificarlos y reproducirlos sin limitación."),
  item("15.2", "LA AGENCIA conserva la titularidad de sus metodologías, plantillas, herramientas y conocimientos generales, así como de las propuestas y piezas no aprobadas ni pagadas."),
  item("15.3", "El material de grabación en bruto se entregará a EL CLIENTE si lo solicita por escrito al término del contrato."),
  item("15.4", "LA AGENCIA podrá mencionar el proyecto e incluir piezas publicadas en su portafolio, salvo que EL CLIENTE se oponga por escrito."),
);

/* DÉCIMA SEXTA */
add(clausula("Derechos de imagen y testimonios"),
  item("16.1", "EL CLIENTE autoriza, y su representante legal consiente a título personal, el uso del nombre, la imagen y la voz de este último en las piezas de comunicación de Kenia producidas bajo este contrato, durante su vigencia y mientras esas piezas permanezcan publicadas en los canales de la marca."),
  item("16.2", "Para cada tercero que aparezca en las piezas —miembros del box, entrevistados, asistentes a eventos— LA AGENCIA recabará una **autorización escrita de uso de imagen** antes de la publicación. EL CLIENTE colaborará en obtenerlas."),
  item("16.3", "Las reseñas en tiendas de aplicaciones se solicitarán **sin pago, sin contraprestación y sin dictar su contenido**, conforme a las políticas de Apple y Google."),
);

/* DÉCIMA SÉPTIMA */
add(clausula("Confidencialidad y datos personales"),
  item("17.1", "Cada parte mantendrá en reserva la información técnica, comercial y financiera de la otra a la que tenga acceso con ocasión de este contrato, y no la divulgará a terceros sin autorización escrita. Esta obligación subsiste durante dos (2) años tras la terminación del contrato."),
  item("17.2", "Los datos personales de los entrevistados, de los miembros de la lista de espera y de cualquier persona que participe en el proyecto se usarán solo para los fines del proyecto, y no se cederán a terceros."),
  item("17.3", "No se considera confidencial la información que ya sea pública o que deba revelarse por mandato legal o judicial."),
);

/* DÉCIMA OCTAVA */
add(clausula("Cuentas, accesos y activos digitales"),
  item("18.1", "Las cuentas de redes sociales, el administrador comercial de Meta, el canal de WhatsApp de la lista de espera, los dominios y los demás activos digitales de Kenia son **propiedad de EL CLIENTE** y deben estar registrados a su nombre o al de su representante legal."),
  item("18.2", "Al terminar el contrato, LA AGENCIA entregará la lista de accesos que haya administrado y retirará los suyos dentro de los cinco (5) días hábiles siguientes, previa confirmación de que no hay pagos pendientes."),
);

/* DÉCIMA NOVENA */
add(clausula("Comunicación, interlocutores e informes"),
  item("19.1", `Los interlocutores del proyecto son: por EL CLIENTE, su representante legal, **${rl.nombre}**, y su representante autorizado y contacto técnico, **${kt.nombre}**; y por LA AGENCIA, **${liam.nombre}** y **${alonzo.nombre}**, indistintamente. Sus datos de contacto constan en el **Anexo D**.`),
  item("19.2", "El canal ordinario de comunicación es el grupo de WhatsApp del proyecto; los asuntos contractuales —pagos, cambios de alcance, terminación— se confirmarán además por correo electrónico."),
  item("19.3", "LA AGENCIA entregará el **informe mensual** dentro de los cinco (5) primeros días hábiles del mes siguiente, y las partes celebrarán una reunión mensual de seguimiento en la fecha que acuerden."),
);

/* VIGÉSIMA */
add(clausula("Modificaciones del alcance"),
  item("20.1", "Cualquier ampliación o modificación de los servicios, de los entregables o de los honorarios deberá constar por escrito y aceptarse por ambas partes, mediante adenda firmada o mediante correo electrónico confirmado por las dos."),
  item("20.2", "Las solicitudes que excedan el alcance de la cláusula tercera serán cotizadas por LA AGENCIA antes de iniciarse; ningún trabajo adicional se ejecutará sin aprobación previa de su precio."),
);

/* VIGÉSIMA PRIMERA */
add(clausula("Terminación anticipada"),
  item("21.1", `Cualquiera de las partes podrá terminar el contrato sin expresión de causa, notificándolo por escrito con al menos **${n(K.dias_preaviso_terminacion)} días continuos** de anticipación. Durante el preaviso, los servicios y los pagos continúan con normalidad.`),
  item("21.2", "Cualquiera de las partes podrá terminar el contrato por incumplimiento grave de la otra, si esta no lo subsana dentro de los diez (10) días continuos siguientes a su notificación escrita."),
  item("21.3", "En todo caso de terminación, EL CLIENTE pagará los honorarios devengados hasta la fecha efectiva de terminación y los gastos aprobados pendientes, y LA AGENCIA entregará los trabajos terminados y pagados."),
);

/* VIGÉSIMA SEGUNDA */
add(clausula("Limitación de responsabilidad"),
  item("22.1", "LA AGENCIA no será responsable por suspensiones, bloqueos, cambios de algoritmo o de políticas de Meta, Apple, Google u otras plataformas; por fallas, retrasos o rechazos de la aplicación en tiendas; ni por la exactitud de la información del producto suministrada por EL CLIENTE."),
  item("22.2", "En cualquier caso, la responsabilidad total de LA AGENCIA frente a EL CLIENTE por causa de este contrato no excederá del monto de los honorarios efectivamente pagados durante los dos (2) últimos meses."),
);

/* VIGÉSIMA TERCERA */
add(clausula("Independencia de las partes"),
  p("Este contrato es de naturaleza civil y mercantil. No crea relación laboral, sociedad ni mandato entre las partes. LA AGENCIA presta sus servicios con autonomía técnica y organizativa, con sus propios medios y personal, y es la única responsable de las obligaciones con quienes contrate para ejecutarlos."),
);

/* VIGÉSIMA CUARTA */
add(clausula("Notificaciones"),
  p("Las notificaciones relacionadas con este contrato se harán por escrito a las direcciones de correo electrónico indicadas en el encabezado y en el Anexo D, y surtirán efecto el día hábil siguiente a su envío. Cada parte deberá informar por escrito cualquier cambio de sus datos de contacto."),
);

/* VIGÉSIMA QUINTA */
add(clausula("Ley aplicable, controversias y domicilio"),
  item("25.1", "Este contrato se rige por las leyes de la República Bolivariana de Venezuela, lugar donde se prestan los servicios y donde tienen su domicilio los representantes de ambas partes."),
  item("25.2", "Las partes procurarán resolver de buena fe y por vía amistosa cualquier controversia, en una reunión que se celebrará dentro de los diez (10) días siguientes a la solicitud de cualquiera de ellas."),
  item("25.3", `De no alcanzarse un acuerdo, las partes eligen como domicilio especial la ciudad de **${K.ciudad_firma}**, a cuya jurisdicción declaran someterse.`),
);

/* ---------- FIRMAS ---------- */
function firma(rol, nombre, cedula) {
  return new TableCell({
    width: { size: W / 2, type: WidthType.DXA },
    borders: { top: NONE, bottom: NONE, left: NONE, right: NONE },
    margins: { top: 0, bottom: 0, left: 200, right: 200 },
    children: [
      gap(900),
      new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 6, color: NEGRO, space: 6 } }, spacing: { after: 40 },
        children: [new TextRun({ text: nombre, font: SANS, size: 21, bold: true, color: NEGRO })] }),
      new Paragraph({ spacing: { after: 30 }, children: [new TextRun({ text: "C.I. N.º " + v(cedula), font: SANS, size: 19, color: GRIS })] }),
      new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: rol, font: SANS, size: 16, bold: true, color: CIAN, characterSpacing: 20 })] }),
    ],
  });
}
add(
  gap(240),
  p(`Se hacen dos (2) ejemplares de un mismo tenor y a un solo efecto, uno para cada parte, en la ciudad de **${K.ciudad_firma}**, el **${K.fecha_firma}**.`),
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W / 2, W / 2],
    borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE },
    rows: [
      new TableRow({ cantSplit: true, children: [
        firma("EL CLIENTE · REPRESENTANTE LEGAL", rl.nombre, rl.cedula),
        firma("LA AGENCIA · " + ag.nombre_comercial.toUpperCase(), liam.nombre, liam.cedula)] }),
      new TableRow({ cantSplit: true, children: [
        firma("EL CLIENTE · REPRESENTANTE AUTORIZADO", kt.nombre, kt.cedula),
        firma("LA AGENCIA · " + ag.nombre_comercial.toUpperCase(), alonzo.nombre, alonzo.cedula)] }),
    ],
  }),
  gap(120),
  p(`${rl.nombre} y ${kt.nombre} suscriben en nombre y representación de **${cli.razon_social}**.`, { color: GRIS, size: 19 }),
  brk(),
);

/* ---------- ANEXO A ---------- */
const subA = t => new Paragraph({ spacing: { before: 60, after: 120 }, children: [new TextRun({ text: t, font: SERIF, size: 24, bold: true, color: PETROLEO })] });
add(
  ...titulo("Anexo A · Entregables", "Qué entrega LA AGENCIA, en qué forma y con qué periodicidad. Los plazos que dependen de insumos de EL CLIENTE se rigen por la cláusula décima tercera."),
  subA("A.1  Fase estratégica — entregada"),
  p("EL CLIENTE declara haber recibido los siguientes entregables, que forman parte de los servicios pagados con la mensualidad de septiembre de 2026:"),
  table(["Entregable", "Contenido", "Fecha"], [
    ["Plataforma de marca", "Núcleo estratégico, público objetivo, cinco buyer personas, beneficios, arquetipo, personalidad y tono, brief de moodboard, naming y descriptor", "8 sep 2026"],
    ["Análisis de entorno (PESTEL)", "Factores políticos, económicos, sociales, tecnológicos, ecológicos y legales, con fuentes", "8 sep 2026"],
    ["Estrategia de contenido", "Embudo, pilares, formatos, canales, calendario, plan de lanzamiento, propuestas de tagline y frases, y relato de marca", "10 sep 2026"],
    ["Guion de grabación", "Guion y lista de tomas de la grabación del 13 de septiembre de 2026", "10 sep 2026"],
  ], [2700, 5248, 1700]),
  gap(240),
  subA("A.2  Fase estratégica — por completar"),
  table(["Entregable", "Plazo"], [
    ["Análisis de marketing complementario: tendencias, análisis del servicio, socios estratégicos, pricing, benchmark, matriz FODA y objetivos SMART", `Hasta el ${K.fecha_analisis_complementario}`],
    ["Moodboard de la marca", `Hasta el ${K.fecha_analisis_complementario}`],
  ], [6648, 3000]),
  gap(240),
  subA("A.3  Gestión mensual — desde el " + K.inicio_publicacion),
  table(["Entregable", "Periodicidad o plazo"], [
    ["Calendario editorial del mes, para aprobación de EL CLIENTE", "Antes del día 25 del mes anterior"],
    [`Producción y publicación de contenido: ${K.volumen_mensual}`, "Mensual, según el calendario aprobado"],
    ["Gestión de campañas pagadas en Meta", "Continua, con el presupuesto de EL CLIENTE"],
    ["Campaña de lanzamiento de la aplicación", "En la fecha de lanzamiento que EL CLIENTE confirme por escrito"],
    ["Informe mensual de resultados", "Cinco primeros días hábiles del mes siguiente"],
    ["Reunión mensual de seguimiento", "Una al mes, en la fecha que acuerden las partes"],
  ], [6648, 3000]),
  brk(),
);

/* ---------- ANEXO B ---------- */
add(
  ...titulo("Anexo B · Calendario de pagos", `Pago por mes adelantado, a más tardar el día ${dl} de cada mes (cláusula séptima).`),
  table(["Cuota", "Período", "Concepto", "Monto", "Vence", "Estado"], pagos, [700, 1650, 1900, 1550, 2048, 1800], { total: true }),
  gap(240),
  recuadro([
    new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: "DATOS DE PAGO", font: SANS, size: 16, bold: true, color: CIAN, characterSpacing: 30 })] }),
    p(`**Método:** ${v(K.metodo_pago, '________________________________')}`, { after: 60, align: AlignmentType.LEFT }),
    p(`**Cuenta o identificador:** ${v(K.datos_cuenta_pago, '________________________________')}`, { after: 60, align: AlignmentType.LEFT }),
    p("**Moneda:** dólares de los Estados Unidos de América. Comisiones a cargo de EL CLIENTE.", { after: 60, align: AlignmentType.LEFT }),
    p("**Comprobante:** al grupo del proyecto el mismo día del pago. Recibo de LA AGENCIA en tres días hábiles.", { after: 0, align: AlignmentType.LEFT }),
  ]),
  gap(200),
  p("A partir de enero de 2027, si el contrato se renueva, cada mensualidad de US$ 1.200,00 vencerá el mismo día de cada mes, salvo revisión acordada conforme a la cláusula 6.4.", { color: GRIS }),
  p("El presupuesto de pauta (referencia: US$ 300 mensuales) **no figura en este calendario**: EL CLIENTE lo paga directamente a Meta.", { color: GRIS }),
  brk(),
);

/* ---------- ANEXO C ---------- */
add(
  ...titulo("Anexo C · Metas de referencia", "Objetivos de referencia, no garantías de resultado (cláusula décima cuarta). Horizonte: 90 días desde la publicación de la aplicación en tiendas, con US$ 300 mensuales de pauta."),
  table(["Indicador", "Meta a 90 días", "Depende principalmente de"], [
    ["Crecimiento de audiencia", "2.500 seguidores", "Contenido y pauta"],
    ["Descargas", "1.200", "Contenido, pauta y ficha de tienda"],
    ["Suscriptores de pago", "40", "Producto, precio, prueba gratuita y métodos de pago"],
    ["Conversión descarga → pago", "3 %", "Producto, precio y métodos de pago"],
    ["Costo de adquisición por suscriptor (solo pauta)", "Menos de US$ 25", "Pauta, conversión y métodos de pago"],
    ["Retención a 30 días", "25 %", "Producto"],
    ["Recuperación de la inversión en pauta", "Menos de 2 meses", "Precio, conversión y retención"],
  ], [3600, 2300, 3748]),
  gap(220),
  p("La medición de conversión, suscriptores y retención requiere que la aplicación emita eventos de analítica o que EL CLIENTE facilite los datos de las consolas de App Store y Google Play. Sin esos datos, LA AGENCIA informará solo de los indicadores que pueda medir en redes y en pauta.", { color: GRIS }),
  brk(),
);

/* ---------- ANEXO D ---------- */
add(
  ...titulo("Anexo D · Directorio del proyecto", "Datos de contacto para comunicaciones y notificaciones (cláusulas décima novena y vigésima cuarta)."),
  table(["Nombre", "Rol", "C.I.", "Teléfono", "Correo"], [
    [cli.razon_social, "Cliente · titular de Kenia", cli.tipo_id_fiscal + " " + v(cli.id_fiscal, '__________'), "—", v(cli.correo, '__________')],
    [rl.nombre, "Cliente · representante legal", v(rl.cedula, '__________'), v(rl.telefono, '__________'), v(rl.correo, '__________')],
    [kt.nombre, "Cliente · representante autorizado y contacto técnico", v(kt.cedula, '__________'), v(kt.telefono, '__________'), v(kt.correo, '__________')],
    [liam.nombre, "Agencia · Organic Club", v(liam.cedula, '__________'), v(liam.telefono, '__________'), v(liam.correo, '__________')],
    [alonzo.nombre, "Agencia · Organic Club", v(alonzo.cedula, '__________'), v(alonzo.telefono, '__________'), v(alonzo.correo, '__________')],
  ], [2048, 2300, 1500, 1500, 2300]),
  gap(400),
  p("Las partes firman en señal de conformidad con los cuatro anexos.", { color: GRIS }),
  gap(300),
  table(["", "Iniciales"], [
    [rl.nombre + " — EL CLIENTE", ""],
    [kt.nombre + " — EL CLIENTE", ""],
    [liam.nombre + " — LA AGENCIA", ""],
    [alonzo.nombre + " — LA AGENCIA", ""],
  ], [6648, 3000]),
);

/* ================= BUILD ================= */
const doc = new Document({
  creator: "Organic Club",
  title: "Contrato de prestación de servicios — Organic Club × Kenia",
  description: "Contrato de prestación de servicios profesionales entre Organic Club y Ninro Libre para la aplicación Kenia.",
  styles: { default: { document: { run: { font: SANS, size: 21, color: NEGRO } } } },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 },
                          margin: { top: 1300, bottom: 1200, left: 1296, right: 1296 } } },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { before: 120 },
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: LINEA, space: 8 } },
      children: [
        new TextRun({ text: "Contrato Organic Club × Kenia  ·  Iniciales: _____  _____  _____  _____  ·  Página ", font: SANS, size: 15, color: GRIS_CLARO }),
        new TextRun({ children: [PageNumber.CURRENT], font: SANS, size: 15, color: GRIS_CLARO }),
        new TextRun({ text: " de ", font: SANS, size: 15, color: GRIS_CLARO }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: SANS, size: 15, color: GRIS_CLARO }),
      ],
    })] }) },
    children: C,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(salida, buf);
  console.log((usaReales ? "con datos reales" : "plantilla en blanco") + " →", salida, (buf.length / 1024).toFixed(0) + " KB");
});
