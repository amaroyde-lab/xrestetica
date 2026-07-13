const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, HeadingLevel,
  LevelFormat, VerticalAlign
} = require('docx');
const fs = require('fs');

const CL_BLUE = "1F497D";
const HEADER_BG = "2E5F9E";
const SUBTOTAL_BG = "D6E4F0";
const ALT_ROW = "F2F7FB";
const WHITE = "FFFFFF";
const LIGHT_GRAY = "F5F5F5";

const border = (color = "CCCCCC") => ({
  top: { style: BorderStyle.SINGLE, size: 4, color },
  bottom: { style: BorderStyle.SINGLE, size: 4, color },
  left: { style: BorderStyle.SINGLE, size: 4, color },
  right: { style: BorderStyle.SINGLE, size: 4, color },
});
const noBorder = {
  top: { style: BorderStyle.NONE },
  bottom: { style: BorderStyle.NONE },
  left: { style: BorderStyle.NONE },
  right: { style: BorderStyle.NONE },
};

function headerCell(text, widthPct) {
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    shading: { fill: HEADER_BG, type: ShadingType.CLEAR },
    borders: border("2E5F9E"),
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, color: WHITE, size: 18, font: "Arial" })]
    })]
  });
}

function dataCell(children, widthPct, fill = WHITE, align = AlignmentType.LEFT) {
  const paras = Array.isArray(children) ? children : [children];
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    shading: { fill, type: ShadingType.CLEAR },
    borders: border("BBBBBB"),
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    verticalAlign: VerticalAlign.TOP,
    children: paras.map(p => {
      if (typeof p === 'string') {
        return new Paragraph({
          alignment: align,
          children: [new TextRun({ text: p, size: 18, font: "Arial" })]
        });
      }
      return p;
    })
  });
}

function subtotalCell(text, widthPct, align = AlignmentType.LEFT) {
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    shading: { fill: SUBTOTAL_BG, type: ShadingType.CLEAR },
    borders: border("2E5F9E"),
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, bold: true, size: 18, font: "Arial", color: "1F3D6E" })]
    })]
  });
}

function totalCell(text, widthPct, align = AlignmentType.LEFT, fontSize = 18) {
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    shading: { fill: "1F497D", type: ShadingType.CLEAR },
    borders: border("1F497D"),
    margins: { top: 100, bottom: 100, left: 160, right: 160 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, bold: true, size: fontSize, font: "Arial", color: WHITE })]
    })]
  });
}

function sectionTitle(text) {
  return new Paragraph({
    spacing: { before: 280, after: 100 },
    children: [new TextRun({
      text,
      bold: true,
      size: 22,
      font: "Arial",
      color: "1F497D"
    })]
  });
}

function bodyPara(text, { bold = false, size = 19, spaceBefore = 60, spaceAfter = 60 } = {}) {
  return new Paragraph({
    spacing: { before: spaceBefore, after: spaceAfter },
    children: [new TextRun({ text, bold, size, font: "Arial", color: "2C2C2C" })]
  });
}

const PAGE_W = 9026;

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 19 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 }
      }
    },
    children: [

      // ── TÍTULO DEL DOCUMENTO ──────────────────────────────────────────
      new Paragraph({
        spacing: { before: 0, after: 100 },
        children: [new TextRun({
          text: "VALORIZACIÓN ECONÓMICA DE MERCADO",
          bold: true, size: 28, font: "Arial", color: "1F497D"
        })]
      }),
      new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "2E5F9E", space: 1 } },
        spacing: { before: 0, after: 200 },
        children: [new TextRun({
          text: "Proyecto de Plataforma Web XR Estética · Sección 2.4 del Acuerdo de Compromiso A+S",
          size: 19, font: "Arial", color: "555555"
        })]
      }),

      // ── DATOS GENERALES ───────────────────────────────────────────────
      sectionTitle("Datos generales del proyecto"),

      new Table({
        width: { size: PAGE_W, type: WidthType.DXA },
        columnWidths: [2256, 2256, 2256, 2258],
        rows: [
          new TableRow({ children: [
            headerCell("Asignatura", 25),
            headerCell("Docente / PM", 25),
            headerCell("Socio comunitario", 25),
            headerCell("Fecha estimación", 25),
          ]}),
          new TableRow({ children: [
            dataCell("Ingeniería de Software INFB8072", 25, ALT_ROW),
            dataCell("Carolina Martínez Valenzuela", 25, ALT_ROW),
            dataCell("Xaviera Ramírez · XR Estética", 25, ALT_ROW),
            dataCell("Junio 2026", 25, ALT_ROW, AlignmentType.CENTER),
          ]}),
        ]
      }),

      bodyPara(""),

      // ── SECCIÓN A ─────────────────────────────────────────────────────
      sectionTitle("A)  Costo de mano de obra (estimación de horas)"),
      bodyPara("Base de cálculo: 5 desarrolladores × 8 semanas × 32 hrs/semana (dedicación académica efectiva) = 1.280 hrs brutas. Se aplican 900 hrs netas estimadas de desarrollo real, considerando reuniones, revisiones y capacitación. Tarifas de mercado vigentes en Chile 2026 para células de desarrollo de este perfil.", { size: 18 }),

      new Paragraph({ spacing: { before: 120, after: 60 }, children: [] }),

      new Table({
        width: { size: PAGE_W, type: WidthType.DXA },
        columnWidths: [3160, 630, 1350, 1450, 2436],
        rows: [
          new TableRow({ children: [
            headerCell("Rol / Perfil de mercado", 35),
            headerCell("Hrs", 7),
            headerCell("Tarifa/hr CLP", 15),
            headerCell("Subtotal CLP", 16),
            headerCell("Módulo principal", 27),
          ]}),
          new TableRow({ children: [
            dataCell([
              new Paragraph({ children: [new TextRun({ text: "5 × Desarrollador Fullstack Jr.–Mid", bold: true, size: 18, font: "Arial" })] }),
              new Paragraph({ spacing: { before: 20 }, children: [new TextRun({ text: "WordPress, WooCommerce, PHP 8.2, JS ES2022, Tailwind CSS", size: 16, font: "Arial", color: "666666" })] }),
            ], 35, WHITE),
            dataCell("640", 7, WHITE, AlignmentType.CENTER),
            dataCell("$28.000", 15, WHITE, AlignmentType.RIGHT),
            dataCell("$17.920.000", 16, WHITE, AlignmentType.RIGHT),
            dataCell("E-commerce, Agendamiento, B2B, UX", 27, WHITE),
          ]}),
          new TableRow({ children: [
            dataCell([
              new Paragraph({ children: [new TextRun({ text: "Diseño UI/UX (rol incluido en equipo)", bold: true, size: 18, font: "Arial" })] }),
              new Paragraph({ spacing: { before: 20 }, children: [new TextRun({ text: "Wireframes, prototipos, responsivo, portafolio antes/después", size: 16, font: "Arial", color: "666666" })] }),
            ], 35, ALT_ROW),
            dataCell("80", 7, ALT_ROW, AlignmentType.CENTER),
            dataCell("$34.000", 15, ALT_ROW, AlignmentType.RIGHT),
            dataCell("$2.720.000", 16, ALT_ROW, AlignmentType.RIGHT),
            dataCell("Contenido y UX", 27, ALT_ROW),
          ]}),
          new TableRow({ children: [
            dataCell([
              new Paragraph({ children: [new TextRun({ text: "QA & Testing (rol incluido en equipo)", bold: true, size: 18, font: "Arial" })] }),
              new Paragraph({ spacing: { before: 20 }, children: [new TextRun({ text: "Pruebas funcionales, regresión, compatibilidad móvil/desktop", size: 16, font: "Arial", color: "666666" })] }),
            ], 35, WHITE),
            dataCell("60", 7, WHITE, AlignmentType.CENTER),
            dataCell("$25.000", 15, WHITE, AlignmentType.RIGHT),
            dataCell("$1.500.000", 16, WHITE, AlignmentType.RIGHT),
            dataCell("Todos los módulos", 27, WHITE),
          ]}),
          new TableRow({ children: [
            dataCell([
              new Paragraph({ children: [new TextRun({ text: "1 × PM / Consultor Senior (docente supervisora)", bold: true, size: 18, font: "Arial" })] }),
              new Paragraph({ spacing: { before: 20 }, children: [new TextRun({ text: "Supervisión técnica, gestión de alcance, revisión de entregables, reuniones con socio comunitario", size: 16, font: "Arial", color: "666666" })] }),
            ], 35, ALT_ROW),
            dataCell("48", 7, ALT_ROW, AlignmentType.CENTER),
            dataCell("$75.000", 15, ALT_ROW, AlignmentType.RIGHT),
            dataCell("$3.600.000", 16, ALT_ROW, AlignmentType.RIGHT),
            dataCell("Gestión y traspaso", 27, ALT_ROW),
          ]}),
          new TableRow({ children: [
            dataCell([
              new Paragraph({ children: [new TextRun({ text: "Integración pasarela Webpay / MercadoPago SDK v3", bold: true, size: 18, font: "Arial" })] }),
              new Paragraph({ spacing: { before: 20 }, children: [new TextRun({ text: "Configuración, pruebas sandbox, certificación en entorno real", size: 16, font: "Arial", color: "666666" })] }),
            ], 35, WHITE),
            dataCell("24", 7, WHITE, AlignmentType.CENTER),
            dataCell("$40.000", 15, WHITE, AlignmentType.RIGHT),
            dataCell("$960.000", 16, WHITE, AlignmentType.RIGHT),
            dataCell("E-commerce", 27, WHITE),
          ]}),
          new TableRow({ children: [
            dataCell([
              new Paragraph({ children: [new TextRun({ text: "Configuración seguridad y rendimiento", bold: true, size: 18, font: "Arial" })] }),
              new Paragraph({ spacing: { before: 20 }, children: [new TextRun({ text: "Wordfence, WP Rocket, UpdraftPlus, hardening WordPress", size: 16, font: "Arial", color: "666666" })] }),
            ], 35, ALT_ROW),
            dataCell("16", 7, ALT_ROW, AlignmentType.CENTER),
            dataCell("$35.000", 15, ALT_ROW, AlignmentType.RIGHT),
            dataCell("$560.000", 16, ALT_ROW, AlignmentType.RIGHT),
            dataCell("Infraestructura", 27, ALT_ROW),
          ]}),
          new TableRow({ children: [
            dataCell([
              new Paragraph({ children: [new TextRun({ text: "Instagram Graph API + CRM básico WordPress", bold: true, size: 18, font: "Arial" })] }),
              new Paragraph({ spacing: { before: 20 }, children: [new TextRun({ text: "Feed nativo, formulario cotización B2B, panel unificado de administración", size: 16, font: "Arial", color: "666666" })] }),
            ], 35, WHITE),
            dataCell("20", 7, WHITE, AlignmentType.CENTER),
            dataCell("$38.000", 15, WHITE, AlignmentType.RIGHT),
            dataCell("$760.000", 16, WHITE, AlignmentType.RIGHT),
            dataCell("B2B y Contenido", 27, WHITE),
          ]}),
          new TableRow({ children: [
            dataCell([
              new Paragraph({ children: [new TextRun({ text: "Capacitación y manual de traspaso de activos digitales", bold: true, size: 18, font: "Arial" })] }),
              new Paragraph({ spacing: { before: 20 }, children: [new TextRun({ text: "Sesiones con propietaria, documentación de credenciales y CMS", size: 16, font: "Arial", color: "666666" })] }),
            ], 35, ALT_ROW),
            dataCell("12", 7, ALT_ROW, AlignmentType.CENTER),
            dataCell("$50.000", 15, ALT_ROW, AlignmentType.RIGHT),
            dataCell("$600.000", 16, ALT_ROW, AlignmentType.RIGHT),
            dataCell("Gestión y traspaso", 27, ALT_ROW),
          ]}),
          new TableRow({ children: [
            subtotalCell("SUBTOTAL MANO DE OBRA", 35),
            subtotalCell("900 hrs", 7, AlignmentType.CENTER),
            subtotalCell("—", 15, AlignmentType.CENTER),
            subtotalCell("$28.620.000", 16, AlignmentType.RIGHT),
            subtotalCell("—", 27, AlignmentType.CENTER),
          ]}),
        ]
      }),

      bodyPara(""),

      // ── SECCIÓN B ─────────────────────────────────────────────────────
      sectionTitle("B)  Costos de infraestructura y licencias (inversión inicial, año 1)"),

      new Paragraph({ spacing: { before: 80, after: 60 }, children: [] }),

      new Table({
        width: { size: PAGE_W, type: WidthType.DXA },
        columnWidths: [3612, 1502, 1506, 2406],
        rows: [
          new TableRow({ children: [
            headerCell("Ítem", 40),
            headerCell("Detalle", 17),
            headerCell("Valor referencia", 17),
            headerCell("Costo estimado CLP", 26),
          ]}),
          new TableRow({ children: [
            dataCell([
              new Paragraph({ children: [new TextRun({ text: "Licencia Amelia Pro — plan Standard (anual)", bold: true, size: 18, font: "Arial" })] }),
              new Paragraph({ spacing: { before: 20 }, children: [new TextRun({ text: "1 dominio, pasarelas múltiples (WooCommerce, Stripe), pago anticipado, REST API, recordatorios automáticos", size: 16, font: "Arial", color: "666666" })] }),
            ], 40, WHITE),
            dataCell("wpamelia.com", 17, WHITE, AlignmentType.CENTER),
            dataCell("USD $89/año", 17, WHITE, AlignmentType.CENTER),
            dataCell("$80.280", 26, WHITE, AlignmentType.RIGHT),
          ]}),
          new TableRow({ children: [
            dataCell([
              new Paragraph({ children: [new TextRun({ text: "Hosting profesional WordPress anual", bold: true, size: 18, font: "Arial" })] }),
              new Paragraph({ spacing: { before: 20 }, children: [new TextRun({ text: "PHP 8.2+, MySQL 8.0, NVMe SSD, SSL Let's Encrypt incluido, cPanel, soporte 24/7 (ej: DCH.cl / HostChile)", size: 16, font: "Arial", color: "666666" })] }),
            ], 40, ALT_ROW),
            dataCell("Plan Premium WP", 17, ALT_ROW, AlignmentType.CENTER),
            dataCell("Mercado CL", 17, ALT_ROW, AlignmentType.CENTER),
            dataCell("$79.990", 26, ALT_ROW, AlignmentType.RIGHT),
          ]}),
          new TableRow({ children: [
            dataCell([
              new Paragraph({ children: [new TextRun({ text: "Dominio .cl anual", bold: true, size: 18, font: "Arial" })] }),
              new Paragraph({ spacing: { before: 20 }, children: [new TextRun({ text: "Registro NIC Chile, titularidad exclusiva del socio comunitario", size: 16, font: "Arial", color: "666666" })] }),
            ], 40, WHITE),
            dataCell("1 año", 17, WHITE, AlignmentType.CENTER),
            dataCell("Mercado CL", 17, WHITE, AlignmentType.CENTER),
            dataCell("$9.950", 26, WHITE, AlignmentType.RIGHT),
          ]}),
          new TableRow({ children: [
            dataCell([
              new Paragraph({ children: [new TextRun({ text: "WP Rocket — caché y rendimiento (anual)", bold: true, size: 18, font: "Arial" })] }),
              new Paragraph({ spacing: { before: 20 }, children: [new TextRun({ text: "Licencia single-site, compresión, lazy load, CDN-ready", size: 16, font: "Arial", color: "666666" })] }),
            ], 40, ALT_ROW),
            dataCell("wp-rocket.me", 17, ALT_ROW, AlignmentType.CENTER),
            dataCell("USD $59/año", 17, ALT_ROW, AlignmentType.CENTER),
            dataCell("$53.220", 26, ALT_ROW, AlignmentType.RIGHT),
          ]}),
          new TableRow({ children: [
            dataCell([
              new Paragraph({ children: [new TextRun({ text: "UpdraftPlus Premium — backups (anual)", bold: true, size: 18, font: "Arial" })] }),
              new Paragraph({ spacing: { before: 20 }, children: [new TextRun({ text: "Personal plan, 1 sitio, restauración con un clic, almacenamiento externo", size: 16, font: "Arial", color: "666666" })] }),
            ], 40, WHITE),
            dataCell("updraftplus.com", 17, WHITE, AlignmentType.CENTER),
            dataCell("USD $42/año", 17, WHITE, AlignmentType.CENTER),
            dataCell("$37.880", 26, WHITE, AlignmentType.RIGHT),
          ]}),
          new TableRow({ children: [
            dataCell([
              new Paragraph({ children: [new TextRun({ text: "Wordfence Security (firewall y escaneo)", bold: true, size: 18, font: "Arial" })] }),
              new Paragraph({ spacing: { before: 20 }, children: [new TextRun({ text: "Versión gratuita suficiente para este alcance (licencia premium: USD $99/año si se requiere)", size: 16, font: "Arial", color: "666666" })] }),
            ], 40, ALT_ROW),
            dataCell("wordfence.com", 17, ALT_ROW, AlignmentType.CENTER),
            dataCell("USD $0", 17, ALT_ROW, AlignmentType.CENTER),
            dataCell("$0", 26, ALT_ROW, AlignmentType.RIGHT),
          ]}),
          new TableRow({ children: [
            subtotalCell("SUBTOTAL INFRAESTRUCTURA Y LICENCIAS", 40),
            subtotalCell("—", 17, AlignmentType.CENTER),
            subtotalCell("—", 17, AlignmentType.CENTER),
            subtotalCell("$261.320", 26, AlignmentType.RIGHT),
          ]}),
        ]
      }),

      bodyPara(""),

      // ── SECCIÓN C ─────────────────────────────────────────────────────
      sectionTitle("C)  Resumen ejecutivo — Costo comercial de mercado"),

      new Table({
        width: { size: PAGE_W, type: WidthType.DXA },
        columnWidths: [6769, 2257],
        rows: [
          new TableRow({ children: [
            headerCell("Concepto", 75),
            headerCell("Monto CLP", 25),
          ]}),
          new TableRow({ children: [
            dataCell("Mano de obra total (900 horas estimadas)", 75, ALT_ROW),
            dataCell("$28.620.000", 25, ALT_ROW, AlignmentType.RIGHT),
          ]}),
          new TableRow({ children: [
            dataCell("Infraestructura y licencias (año 1)", 75, WHITE),
            dataCell("$261.320", 25, WHITE, AlignmentType.RIGHT),
          ]}),
          new TableRow({ children: [
            dataCell("Margen de gestión de proyecto y overhead de agencia (15%)", 75, ALT_ROW),
            dataCell("$4.332.198", 25, ALT_ROW, AlignmentType.RIGHT),
          ]}),
          new TableRow({ children: [
            totalCell("COSTO COMERCIAL TOTAL ESTIMADO DE MERCADO  —  Año 2026", 75, AlignmentType.LEFT, 20),
            totalCell("$33.213.518 CLP", 25, AlignmentType.RIGHT, 20),
          ]}),
        ]
      }),

      new Paragraph({ spacing: { before: 200, after: 80 }, children: [] }),

      // ── PÁRRAFO 2.4 ───────────────────────────────────────────────────
      new Paragraph({
        border: {
          top: { style: BorderStyle.SINGLE, size: 8, color: "2E5F9E", space: 1 },
          bottom: { style: BorderStyle.SINGLE, size: 8, color: "2E5F9E", space: 1 },
          left: { style: BorderStyle.THICK, size: 24, color: "2E5F9E", space: 4 },
          right: { style: BorderStyle.NONE },
        },
        shading: { fill: "EEF4FB", type: ShadingType.CLEAR },
        spacing: { before: 100, after: 60, line: 276 },
        indent: { left: 240, right: 120 },
        children: [
          new TextRun({ text: "Texto para la sección 2.4 del Acuerdo de Compromiso A+S  (copiar y pegar directamente)", bold: true, size: 18, font: "Arial", color: "1F497D" }),
          new TextRun({ text: "\n", break: 1 }),
        ]
      }),
      new Paragraph({
        shading: { fill: "EEF4FB", type: ShadingType.CLEAR },
        border: {
          bottom: { style: BorderStyle.SINGLE, size: 8, color: "2E5F9E", space: 1 },
          left: { style: BorderStyle.THICK, size: 24, color: "2E5F9E", space: 4 },
          right: { style: BorderStyle.NONE },
          top: { style: BorderStyle.NONE },
        },
        spacing: { before: 60, after: 100, line: 300 },
        indent: { left: 240, right: 120 },
        children: [
          new TextRun({
            text: "El presente servicio, desarrollado en el marco de la asignatura Ingeniería de Software (INFB8072) de la Universidad Tecnológica Metropolitana bajo la metodología de Aprendizaje más Servicio (A+S), tiene un valor de referencia comercial estimado en $33.213.518 (treinta y tres millones doscientos trece mil quinientos dieciocho pesos) en el mercado chileno de desarrollo web durante el año 2026. Dicha estimación corresponde a la contratación equivalente de un equipo de cinco (5) desarrolladores fullstack, un (1) consultor/PM senior y los costos de infraestructura y licencias asociados (hosting profesional, plugin de reservas Amelia Pro, dominio .cl, plugins de caché y respaldo), calculados sobre un total aproximado de 900 horas de trabajo y tarifas vigentes en el mercado nacional para perfiles de este tipo ($25.000–$75.000 CLP/hora según nivel y rol). Este valor es únicamente de referencia y tiene por objeto transparentar el alcance técnico y económico del proyecto ante el socio comunitario. El costo real para la organización, institución, empresa y/o persona que recibe el servicio es de $0 (cero pesos), ya que la iniciativa se enmarca en una actividad académica sin vínculo comercial ni contractual entre la universidad y el socio comunitario, conforme a lo establecido en el punto 4.1 del presente acuerdo.",
            size: 19,
            font: "Arial",
            color: "1A1A1A",
          })
        ]
      }),

      new Paragraph({ spacing: { before: 200, after: 60 }, children: [] }),

      // ── NOTAS METODOLÓGICAS ───────────────────────────────────────────
      new Paragraph({
        spacing: { before: 80, after: 60 },
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 1 } },
        children: [
          new TextRun({ text: "Notas metodológicas y fuentes", bold: true, size: 17, font: "Arial", color: "555555" })
        ]
      }),
      new Paragraph({
        spacing: { before: 40, after: 40, line: 260 },
        children: [
          new TextRun({
            text: "Tarifas de mano de obra basadas en rangos de mercado chileno 2026: desarrollador fullstack junior–mid $25.000–$35.000 CLP/hr; UI/UX $30.000–$40.000 CLP/hr; PM/consultor senior $60.000–$90.000 CLP/hr (fuentes: LX3.ai, AltaHosting.cl, Cronoshare.cl, Vita Wallet, Freelancermap.com — consultadas junio 2026). Horas estimadas para 5 desarrolladores en 8 semanas a dedicación académica efectiva (~32 hrs/sem/persona para el proyecto). Tipo de cambio utilizado: USD 1 = CLP $902 (13 de junio de 2026, fuente: DolarOnline.cl / Wise). Licencia Amelia Pro Standard: USD $89/año (wpamelia.com, verificado mayo 2026). Margen de overhead del 15% refleja costos de coordinación, comunicaciones y gestión típicos de una agencia digital PYME. WooCommerce, WordPress, Wordfence y SSL Let's Encrypt son software de código abierto sin costo de licencia. Este presupuesto es una valorización referencial académica y no constituye cotización comercial formal.",
            size: 16,
            font: "Arial",
            color: "666666",
          })
        ]
      }),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Valorizacion_Economica_XR_Estetica_2026.docx", buffer);
  console.log("Documento generado correctamente.");
});