const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  PageNumber, NumberFormat, Footer, Header, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, PageBreak, Tab, TabStopType,
  convertInchesToTwip, LevelFormat, UnderlineType
} = require("docx");
const fs = require("fs");

// Load chapters
const perilepsi = require("./chapters/perilepsi");
const ch1 = require("./chapters/ch1_eisagogi");
const ch2 = require("./chapters/ch2_theoritiko");
const ch3 = require("./chapters/ch3_pyrosvestiko");
const ch4 = require("./chapters/ch4_ekpaideutiko");
const ch5 = require("./chapters/ch5_epidrasi");
const ch6 = require("./chapters/ch6_casestudies");
const ch7 = require("./chapters/ch7_technologia");
const ch8 = require("./chapters/ch8_symperasmata");

const chapters = [ch1, ch2, ch3, ch4, ch5, ch6, ch7, ch8];

// === BIBLIOGRAPHY DATA ===
const bibliographyEntries = [
  `Akhloufi, M. A., Couturier, A., & Castro, N. A. (2021). Unmanned aerial vehicles for wildfire monitoring and management: State of the art and future trends. Drones, 5(1), Article 26.`,
  `Burke, C. S., Stagl, K. C., Salas, E., Pierce, L., & Kendall, D. (2006). Understanding team adaptation: A conceptual analysis and model. Journal of Applied Psychology, 91(6), 1189-1207.`,
  `Chuvieco, E., & Congalton, R. G. (1989). Application of remote sensing and geographic information systems to forest fire hazard mapping. Remote Sensing of Environment, 29(2), 147-159.`,
  `Ericsson, K. A. (2006). The influence of experience and deliberate practice on the development of superior expert performance. In K. A. Ericsson et al. (Eds.), The Cambridge handbook of expertise and expert performance (pp. 683-703). Cambridge University Press.`,
  `Ericsson, K. A., Krampe, R. T., & Tesch-Römer, C. (1993). The role of deliberate practice in the acquisition of expert performance. Psychological Review, 100(3), 363-406.`,
  `European Commission. (2021). Union Civil Protection Mechanism. DG ECHO.`,
  `European Commission. (2020). Overview of natural and man-made disaster risks the European Union may face: 2020 edition. Publications Office of the EU.`,
  `Flin, R., O'Connor, P., & Crichton, M. (2008). Safety at the sharp end: A guide to non-technical skills. Ashgate Publishing.`,
  `Grossman, D., & Christensen, L. W. (2008). On combat: The psychology and physiology of deadly conflict in war and in peace (3rd ed.). Warrior Science Publications.`,
  `INSARAG. (2020). INSARAG guidelines (Vols. I-III). UN OCHA.`,
  `Jahnke, S. A., Poston, W. S. C., Haddock, C. K., & Murphy, B. (2016). Firefighting and mental health: Experiences of repeated exposure to trauma. Work, 53(4), 737-744.`,
  `Kleim, B., & Westphal, M. (2011). Mental health in first responders: A review and recommendation for prevention and intervention strategies. European Journal of Psychotraumatology, 2(1), Article 7585.`,
  `Kolb, D. A. (1984). Experiential learning: Experience as the source of learning and development. Prentice-Hall.`,
  `Kolb, D. A., Boyatzis, R. E., & Mainemelis, C. (2001). Experiential learning theory: Previous research and new directions. In R. J. Sternberg & L. Zhang (Eds.), Perspectives on thinking, learning, and cognitive styles (pp. 227-247). Lawrence Erlbaum Associates.`,
  `Lagouvardos, K., Kotroni, V., Giannaros, T. M., & Dafis, S. (2019). Meteorological conditions conducive to the rapid spread of the deadly wildfire in eastern Attica, Greece. Bulletin of the American Meteorological Society, 100(11), 2137-2145.`,
  `Meichenbaum, D. (2007). Stress inoculation training: A preventative and treatment approach. In P. M. Lehrer et al. (Eds.), Principles and practice of stress management (3rd ed., pp. 497-516). Guilford Press.`,
  `Salas, E., Burke, C. S., & Stagl, K. C. (2004). Developing teams and team leaders. In D. V. Day et al. (Eds.), Leader development for transforming organizations (pp. 325-355). Lawrence Erlbaum Associates.`,
  `Salas, E., DiazGranados, D., Klein, C., Burke, C. S., et al. (2008). Does team training improve team performance? A meta-analysis. Human Factors, 50(6), 903-933.`,
  `Salas, E., Tannenbaum, S. I., Kraiger, K., & Smith-Jentsch, K. A. (2012). The science of training and development in organizations: What matters in practice. Psychological Science in the Public Interest, 13(2), 74-101.`,
  `Smith, D. L., Haller, J. M., Korre, M., et al. (2019). The relation of emergency duties to cardiac death among US firefighters. American Journal of Cardiology, 123(5), 736-741.`,
  `Soteriades, E. S., Smith, D. L., Tsismenakis, A. J., et al. (2011). Cardiovascular disease in US firefighters: A systematic review. Cardiology in Review, 19(4), 202-215.`,
  `Williams-Bell, F. M., Kapralos, B., Hogue, A., Murphy, B. M., & Weckman, E. J. (2015). Using serious games and virtual simulation for training in the fire service: A review. Fire Technology, 51(3), 553-584.`,
  `Γενική Γραμματεία Πολιτικής Προστασίας. (2021). Γενικό Σχέδιο Αντιμετώπισης Εκτάκτων Αναγκών εξαιτίας Δασικών Πυρκαγιών (ΙΟΛΑΟΣ 2). ΥΚΚΠΠ.`,
  `Πυροσβεστικό Σώμα Ελλάδας. (n.d.). Εκπαιδευτικές δραστηριότητες και επιχειρησιακή ετοιμότητα. https://www.fireservice.gr`,
  `Υπουργείο Κλιματικής Κρίσης και Πολιτικής Προστασίας. (2022). Έκθεση αντιπυρικής περιόδου 2021. Αθήνα.`,
  `EFFIS. (n.d.). Statistics portal. European Commission, JRC. https://effis.jrc.ec.europa.eu/`
];

// === HELPER FUNCTIONS ===

function makeTextRun(text, opts = {}) {
  return new TextRun({
    text,
    font: "Times New Roman",
    size: opts.size || 24, // 12pt
    bold: opts.bold || false,
    italics: opts.italics || false,
    underline: opts.underline ? { type: UnderlineType.SINGLE } : undefined,
  });
}

function makeParagraph(text, opts = {}) {
  const runs = [];
  if (typeof text === "string") {
    runs.push(makeTextRun(text, opts));
  } else if (Array.isArray(text)) {
    text.forEach(t => runs.push(t));
  }
  return new Paragraph({
    children: runs,
    spacing: { line: 360, after: 120 }, // 1.5 line spacing
    alignment: opts.alignment || AlignmentType.JUSTIFIED,
    indent: opts.indent ? { firstLine: 720 } : undefined,
    heading: opts.heading || undefined,
    pageBreakBefore: opts.pageBreak || false,
  });
}

function makeHeading1(text, pageBreak = true) {
  return new Paragraph({
    children: [makeTextRun(text, { bold: true, size: 28 })], // 14pt
    spacing: { before: 480, after: 240, line: 360 },
    alignment: AlignmentType.CENTER,
    pageBreakBefore: pageBreak,
  });
}

function makeHeading2(text) {
  return new Paragraph({
    children: [makeTextRun(text, { bold: true, size: 26 })], // 13pt
    spacing: { before: 360, after: 200, line: 360 },
    alignment: AlignmentType.LEFT,
  });
}

function makeHeading3(text) {
  return new Paragraph({
    children: [makeTextRun(text, { bold: true, size: 24 })], // 12pt
    spacing: { before: 240, after: 120, line: 360 },
    alignment: AlignmentType.LEFT,
  });
}

function makeBodyParagraph(text) {
  return new Paragraph({
    children: [makeTextRun(text)],
    spacing: { line: 360, after: 120 },
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: 720 },
  });
}

function makeTable(tableData) {
  const { caption, headers, rows } = tableData;
  const elements = [];

  // Caption
  elements.push(new Paragraph({
    children: [makeTextRun(caption, { bold: true, size: 22, italics: true })],
    spacing: { before: 240, after: 120, line: 360 },
    alignment: AlignmentType.CENTER,
  }));

  // Header row
  const headerRow = new TableRow({
    children: headers.map(h => new TableCell({
      children: [new Paragraph({
        children: [makeTextRun(h, { bold: true, size: 20 })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 60 },
      })],
      shading: { type: ShadingType.SOLID, color: "D9E2F3" },
      width: { size: Math.floor(100 / headers.length), type: WidthType.PERCENTAGE },
    })),
  });

  // Data rows
  const dataRows = rows.map(row => new TableRow({
    children: row.map(cell => new TableCell({
      children: [new Paragraph({
        children: [makeTextRun(cell, { size: 20 })],
        alignment: AlignmentType.LEFT,
        spacing: { before: 40, after: 40 },
      })],
      width: { size: Math.floor(100 / headers.length), type: WidthType.PERCENTAGE },
    })),
  }));

  elements.push(new Table({
    rows: [headerRow, ...dataRows],
    width: { size: 100, type: WidthType.PERCENTAGE },
  }));

  elements.push(new Paragraph({ spacing: { after: 200 } }));

  return elements;
}

function buildChapterElements(chapter) {
  const elements = [];

  // Chapter title
  elements.push(makeHeading1(`ΚΕΦΑΛΑΙΟ ${chapter.number}: ${chapter.title}`));

  for (const section of chapter.sections) {
    // Determine heading level based on numbering pattern
    const parts = section.heading.split(" ")[0].split(".");
    if (parts.length <= 2 && !parts[1]) {
      elements.push(makeHeading2(section.heading));
    } else if (parts.length === 2) {
      elements.push(makeHeading2(section.heading));
    } else {
      elements.push(makeHeading3(section.heading));
    }

    // Paragraphs
    for (const para of section.paragraphs) {
      if (para.trim()) {
        elements.push(makeBodyParagraph(para));
      }
    }

    // Table if present
    if (section.table) {
      elements.push(...makeTable(section.table));
    }
  }

  return elements;
}

// === BUILD DOCUMENT ===

async function generateDocument() {
  const allElements = [];

  // --- COVER PAGE ---
  allElements.push(new Paragraph({ spacing: { before: 2400 } }));
  allElements.push(makeParagraph("[ΠΑΝΕΠΙΣΤΗΜΙΟ]", { alignment: AlignmentType.CENTER, bold: true, size: 28 }));
  allElements.push(makeParagraph("[ΤΜΗΜΑ]", { alignment: AlignmentType.CENTER, bold: true, size: 26 }));
  allElements.push(new Paragraph({ spacing: { before: 1200 } }));
  allElements.push(new Paragraph({
    children: [makeTextRun("ΔΙΠΛΩΜΑΤΙΚΗ ΕΡΓΑΣΙΑ", { bold: true, size: 32 })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 600, after: 600 },
  }));
  allElements.push(new Paragraph({ spacing: { before: 400 } }));
  allElements.push(new Paragraph({
    children: [makeTextRun("Ο ΡΟΛΟΣ ΤΗΣ ΕΚΠΑΙΔΕΥΣΗΣ ΜΕΣΩ ΑΣΚΗΣΕΩΝ", { bold: true, size: 30 })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 100 },
  }));
  allElements.push(new Paragraph({
    children: [makeTextRun("ΣΤΗ ΒΕΛΤΙΩΣΗ ΤΗΣ ΕΠΙΧΕΙΡΗΣΙΑΚΗΣ ΕΤΟΙΜΟΤΗΤΑΣ", { bold: true, size: 30 })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 100 },
  }));
  allElements.push(new Paragraph({
    children: [makeTextRun("ΤΟΥ ΠΥΡΟΣΒΕΣΤΙΚΟΥ ΣΩΜΑΤΟΣ", { bold: true, size: 30 })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 800 },
  }));
  allElements.push(new Paragraph({ spacing: { before: 1200 } }));
  allElements.push(new Paragraph({
    children: [makeTextRun("Φοιτητής: Άγγελος Καλαφατάκης", { size: 26 })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
  }));
  allElements.push(new Paragraph({
    children: [makeTextRun("Επιβλέπων Καθηγητής: [ΕΠΙΒΛΕΠΩΝ ΚΑΘΗΓΗΤΗΣ]", { size: 26 })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 600 },
  }));
  allElements.push(new Paragraph({
    children: [makeTextRun("[ΠΟΛΗ], [ΕΤΟΣ]", { size: 24 })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
  }));

  // --- PERILEPSI ---
  allElements.push(makeHeading1("ΠΕΡΙΛΗΨΗ", true));
  allElements.push(makeBodyParagraph(perilepsi.perilepsi));
  allElements.push(new Paragraph({
    children: [
      makeTextRun("Λέξεις-κλειδιά: ", { bold: true }),
      makeTextRun(perilepsi.lekseis_kleidi, { italics: true }),
    ],
    spacing: { before: 240, line: 360 },
    alignment: AlignmentType.JUSTIFIED,
  }));

  // --- ABSTRACT ---
  allElements.push(makeHeading1("ABSTRACT", true));
  allElements.push(makeBodyParagraph(perilepsi.abstract_en));
  allElements.push(new Paragraph({
    children: [
      makeTextRun("Keywords: ", { bold: true }),
      makeTextRun(perilepsi.keywords_en, { italics: true }),
    ],
    spacing: { before: 240, line: 360 },
    alignment: AlignmentType.JUSTIFIED,
  }));

  // --- TABLE OF CONTENTS (manual) ---
  allElements.push(makeHeading1("ΠΙΝΑΚΑΣ ΠΕΡΙΕΧΟΜΕΝΩΝ", true));
  const tocEntries = [
    "ΠΕΡΙΛΗΨΗ",
    "ABSTRACT",
    "ΚΕΦΑΛΑΙΟ 1: ΕΙΣΑΓΩΓΗ",
    "  1.1 Αντικείμενο και σκοπός της εργασίας",
    "  1.2 Μεθοδολογική προσέγγιση",
    "  1.3 Δομή της εργασίας",
    "ΚΕΦΑΛΑΙΟ 2: ΘΕΩΡΗΤΙΚΟ ΠΛΑΙΣΙΟ",
    "  2.1 Η έννοια της επιχειρησιακής ετοιμότητας",
    "  2.2 Θεωρίες μάθησης και εκπαίδευσης ενηλίκων",
    "  2.3 Η εκπαίδευση σε επαγγέλματα υψηλού κινδύνου",
    "ΚΕΦΑΛΑΙΟ 3: ΤΟ ΠΥΡΟΣΒΕΣΤΙΚΟ ΣΩΜΑ",
    "  3.1 Η αποστολή και ο ρόλος του ΠΣ στην Ελλάδα",
    "  3.2 Σωματικές, τεχνικές και ψυχολογικές απαιτήσεις",
    "  3.3 Εξέλιξη των κινδύνων: κλιματική αλλαγή",
    "ΚΕΦΑΛΑΙΟ 4: ΤΟ ΕΚΠΑΙΔΕΥΤΙΚΟ ΣΥΣΤΗΜΑ",
    "  4.1 Βασική εκπαίδευση",
    "  4.2 Συνεχιζόμενη εκπαίδευση",
    "  4.3 Τύποι ασκήσεων",
    "  4.4 Σχεδιασμός και αξιολόγηση ασκήσεων",
    "ΚΕΦΑΛΑΙΟ 5: Η ΕΠΙΔΡΑΣΗ ΤΗΣ ΕΚΠΑΙΔΕΥΣΗΣ",
    "  5.1 Επίδραση στη σωματική ετοιμότητα",
    "  5.2 Επίδραση στην τεχνική επάρκεια",
    "  5.3 Επίδραση στην ψυχολογική ανθεκτικότητα",
    "  5.4 Επίδραση στην ομαδική συνεργασία",
    "  5.5 Σχέση συχνότητας ασκήσεων και ετοιμότητας",
    "ΚΕΦΑΛΑΙΟ 6: ΔΙΕΘΝΗΣ ΣΥΝΕΡΓΑΣΙΑ & CASE STUDIES",
    "  6.1 Ο Ευρωπαϊκός Μηχανισμός UCPM",
    "  6.2 Case Study 1: Πυρκαγιά Μάτι (2018)",
    "  6.3 Case Study 2: Ασκήσεις MODEX — ΕΜΑΚ",
    "  6.4 Case Study 3: Πυρκαγιές 2021",
    "  6.5 Οφέλη διακρατικών ασκήσεων",
    "ΚΕΦΑΛΑΙΟ 7: Ο ΡΟΛΟΣ ΤΗΣ ΤΕΧΝΟΛΟΓΙΑΣ",
    "  7.1 Προσομοιωτές εκπαίδευσης",
    "  7.2 Εικονική και επαυξημένη πραγματικότητα",
    "  7.3 Μη επανδρωμένα αεροσκάφη (drones)",
    "  7.4 Γεωγραφικά Συστήματα Πληροφοριών (GIS)",
    "  7.5 Ηλεκτρονική μάθηση",
    "  7.6 Αποτίμηση τεχνολογικών εφαρμογών",
    "ΚΕΦΑΛΑΙΟ 8: ΣΥΜΠΕΡΑΣΜΑΤΑ ΚΑΙ ΠΡΟΤΑΣΕΙΣ",
    "  8.1 Σύνοψη ευρημάτων",
    "  8.2 Προτάσεις βελτίωσης",
    "  8.3 Περιορισμοί της εργασίας",
    "  8.4 Προτάσεις για μελλοντική έρευνα",
    "ΒΙΒΛΙΟΓΡΑΦΙΑ",
  ];
  for (const entry of tocEntries) {
    const isSubEntry = entry.startsWith("  ");
    allElements.push(new Paragraph({
      children: [makeTextRun(entry.trim(), { size: isSubEntry ? 22 : 24, bold: !isSubEntry })],
      spacing: { after: isSubEntry ? 40 : 80, line: 300 },
      indent: isSubEntry ? { left: 720 } : undefined,
    }));
  }

  // --- CHAPTERS ---
  for (const chapter of chapters) {
    allElements.push(...buildChapterElements(chapter));
  }

  // --- BIBLIOGRAPHY ---
  allElements.push(makeHeading1("ΒΙΒΛΙΟΓΡΑΦΙΑ", true));
  for (const entry of bibliographyEntries) {
    allElements.push(new Paragraph({
      children: [makeTextRun(entry, { size: 22 })],
      spacing: { after: 120, line: 300 },
      alignment: AlignmentType.JUSTIFIED,
      indent: { left: 720, hanging: 720 }, // Hanging indent for APA
    }));
  }

  // --- CREATE DOCUMENT ---
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: {
            top: convertInchesToTwip(0.79), // ~2cm
            bottom: convertInchesToTwip(0.79),
            left: convertInchesToTwip(0.98), // ~2.5cm
            right: convertInchesToTwip(0.79),
          },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            children: [makeTextRun("", { size: 18 })],
            alignment: AlignmentType.RIGHT,
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: [
              new TextRun({
                children: [PageNumber.CURRENT],
                font: "Times New Roman",
                size: 20,
              }),
            ],
            alignment: AlignmentType.RIGHT,
          })],
        }),
      },
      children: allElements,
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync("./output/diplomatiki.docx", buffer);
  console.log("✓ Δημιουργήθηκε: output/diplomatiki.docx");

  // Count approximate pages (rough: ~3000 chars per page)
  let totalChars = 0;
  for (const ch of chapters) {
    for (const s of ch.sections) {
      for (const p of s.paragraphs) {
        totalChars += p.length;
      }
    }
  }
  totalChars += perilepsi.perilepsi.length + perilepsi.abstract_en.length;
  console.log(`  Εκτίμηση: ~${Math.round(totalChars / 2500)} σελίδες κυρίου κειμένου`);
  console.log(`  Πηγές βιβλιογραφίας: ${bibliographyEntries.length}`);
  console.log(`  Κεφάλαια: ${chapters.length}`);
  console.log(`  Πίνακες: 6`);
}

generateDocument().catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
