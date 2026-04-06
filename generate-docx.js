const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  PageNumber, NumberFormat, Footer, Header, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, PageBreak, Tab, TabStopType,
  convertInchesToTwip, LevelFormat, UnderlineType, ImageRun,
  ExternalHyperlink, InternalHyperlink, BookmarkStart, BookmarkEnd
} = require("docx");
const fs = require("fs");

// Load chapters
const perilepsi = require("./chapters/perilepsi");
const eisagogi = require("./chapters/ch1_eisagogi");
const ch1 = require("./chapters/ch2_theoritiko");
const ch2 = require("./chapters/ch3_pyrosvestiko");
const ch3 = require("./chapters/ch4_ekpaideutiko");
const ch4 = require("./chapters/ch5_epidrasi");
const ch5 = require("./chapters/ch6_casestudies");
const ch6 = require("./chapters/ch7_technologia");
const symperasmata = require("./chapters/ch8_symperasmata");

const numberedChapters = [ch1, ch2, ch3, ch4, ch5, ch6];
const allContentChapters = [eisagogi, ...numberedChapters, symperasmata];

// Margins in twips: 3cm left, 2.5cm others
const CM_TO_TWIP = 567;
const MARGINS = {
  left: Math.round(3 * CM_TO_TWIP),   // 3cm
  right: Math.round(2.5 * CM_TO_TWIP), // 2.5cm
  top: Math.round(2.5 * CM_TO_TWIP),   // 2.5cm
  bottom: Math.round(2.5 * CM_TO_TWIP) // 2.5cm
};

// === BIBLIOGRAPHY ===
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

// === HELPERS ===
function tr(text, opts = {}) {
  return new TextRun({
    text,
    font: "Times New Roman",
    size: opts.size || 24,
    bold: opts.bold || false,
    italics: opts.italics || false,
    underline: opts.underline ? { type: UnderlineType.SINGLE } : undefined,
  });
}

function heading1(text, pageBreak = true) {
  return new Paragraph({
    children: [tr(text, { bold: true, size: 28 })],
    spacing: { before: 480, after: 240, line: 360 },
    alignment: AlignmentType.CENTER,
    pageBreakBefore: pageBreak,
  });
}

function heading2(text) {
  return new Paragraph({
    children: [tr(text, { bold: true, size: 26 })],
    spacing: { before: 360, after: 200, line: 360 },
    alignment: AlignmentType.LEFT,
  });
}

function heading3(text) {
  return new Paragraph({
    children: [tr(text, { bold: true, size: 24 })],
    spacing: { before: 240, after: 120, line: 360 },
    alignment: AlignmentType.LEFT,
  });
}

function bodyPara(text) {
  return new Paragraph({
    children: [tr(text)],
    spacing: { line: 360, after: 120 },
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: 720 },
  });
}

function makeTable(tableData) {
  const { caption, headers, rows } = tableData;
  const elements = [];
  elements.push(new Paragraph({
    children: [tr(caption, { bold: true, size: 22, italics: true })],
    spacing: { before: 240, after: 120, line: 360 },
    alignment: AlignmentType.CENTER,
  }));

  const headerRow = new TableRow({
    children: headers.map(h => new TableCell({
      children: [new Paragraph({
        children: [tr(h, { bold: true, size: 20 })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 60 },
      })],
      shading: { type: ShadingType.SOLID, color: "D9E2F3" },
      width: { size: Math.floor(100 / headers.length), type: WidthType.PERCENTAGE },
    })),
  });

  const dataRows = rows.map(row => new TableRow({
    children: row.map(cell => new TableCell({
      children: [new Paragraph({
        children: [tr(cell, { size: 20 })],
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

function buildChapter(chapter) {
  const elements = [];
  const titlePrefix = chapter.number !== null ? `ΚΕΦΑΛΑΙΟ ${chapter.number}: ` : "";
  elements.push(heading1(`${titlePrefix}${chapter.title}`));

  for (const section of chapter.sections) {
    const headingText = section.heading;
    const dotCount = (headingText.match(/\./g) || []).length;
    if (dotCount >= 2) {
      elements.push(heading3(headingText));
    } else {
      elements.push(heading2(headingText));
    }

    for (const para of section.paragraphs) {
      if (para.trim()) elements.push(bodyPara(para));
    }
    if (section.table) elements.push(...makeTable(section.table));
  }
  return elements;
}

// === MAIN ===
async function generateDocument() {
  const allElements = [];

  // --- COVER ---
  allElements.push(new Paragraph({ spacing: { before: 2400 } }));
  allElements.push(new Paragraph({
    children: [tr("[ΠΑΝΕΠΙΣΤΗΜΙΟ]", { bold: true, size: 28 })],
    alignment: AlignmentType.CENTER, spacing: { after: 100 },
  }));
  allElements.push(new Paragraph({
    children: [tr("[ΤΜΗΜΑ]", { bold: true, size: 26 })],
    alignment: AlignmentType.CENTER, spacing: { after: 1200 },
  }));
  allElements.push(new Paragraph({
    children: [tr("ΔΙΠΛΩΜΑΤΙΚΗ ΕΡΓΑΣΙΑ", { bold: true, size: 32 })],
    alignment: AlignmentType.CENTER, spacing: { before: 600, after: 600 },
  }));
  allElements.push(new Paragraph({ spacing: { before: 400 } }));
  const titleLines = [
    "Ο ΡΟΛΟΣ ΤΗΣ ΕΚΠΑΙΔΕΥΣΗΣ ΜΕΣΩ ΑΣΚΗΣΕΩΝ",
    "ΣΤΗ ΒΕΛΤΙΩΣΗ ΤΗΣ ΕΠΙΧΕΙΡΗΣΙΑΚΗΣ ΕΤΟΙΜΟΤΗΤΑΣ",
    "ΤΩΝ ΣΩΜΑΤΩΝ ΑΣΦΑΛΕΙΑΣ ΚΑΙ ΤΩΝ ΕΝΟΠΛΩΝ ΔΥΝΑΜΕΩΝ"
  ];
  for (const line of titleLines) {
    allElements.push(new Paragraph({
      children: [tr(line, { bold: true, size: 30 })],
      alignment: AlignmentType.CENTER, spacing: { after: 100 },
    }));
  }
  allElements.push(new Paragraph({ spacing: { before: 1200 } }));
  allElements.push(new Paragraph({
    children: [tr("Φοιτητής: Άγγελος Καλαφατάκης", { size: 26 })],
    alignment: AlignmentType.CENTER, spacing: { after: 200 },
  }));
  allElements.push(new Paragraph({
    children: [tr("Επιβλέπων Καθηγητής: [ΕΠΙΒΛΕΠΩΝ ΚΑΘΗΓΗΤΗΣ]", { size: 26 })],
    alignment: AlignmentType.CENTER, spacing: { after: 600 },
  }));
  allElements.push(new Paragraph({
    children: [tr("[ΠΟΛΗ], [ΕΤΟΣ]", { size: 24 })],
    alignment: AlignmentType.CENTER,
  }));

  // --- PERILEPSI ---
  allElements.push(heading1("ΠΕΡΙΛΗΨΗ", true));
  allElements.push(bodyPara(perilepsi.perilepsi));
  allElements.push(new Paragraph({
    children: [tr("Λέξεις-κλειδιά: ", { bold: true }), tr(perilepsi.lekseis_kleidi, { italics: true })],
    spacing: { before: 240, line: 360 }, alignment: AlignmentType.JUSTIFIED,
  }));

  // --- ABSTRACT ---
  allElements.push(heading1("ABSTRACT", true));
  allElements.push(bodyPara(perilepsi.abstract_en));
  allElements.push(new Paragraph({
    children: [tr("Keywords: ", { bold: true }), tr(perilepsi.keywords_en, { italics: true })],
    spacing: { before: 240, line: 360 }, alignment: AlignmentType.JUSTIFIED,
  }));

  // --- TOC ---
  allElements.push(heading1("ΠΙΝΑΚΑΣ ΠΕΡΙΕΧΟΜΕΝΩΝ", true));
  const tocEntries = [
    { text: "ΠΕΡΙΛΗΨΗ", level: 0 },
    { text: "ABSTRACT", level: 0 },
    { text: "ΕΙΣΑΓΩΓΗ", level: 0 },
    { text: "Αντικείμενο και σκοπός", level: 1 },
    { text: "Μεθοδολογική προσέγγιση", level: 1 },
    { text: "Δομή της εργασίας", level: 1 },
    { text: "ΚΕΦΑΛΑΙΟ 1: ΘΕΩΡΗΤΙΚΟ ΠΛΑΙΣΙΟ", level: 0 },
    { text: "1.1 Η έννοια της επιχειρησιακής ετοιμότητας", level: 1 },
    { text: "1.2 Παράγοντες που επηρεάζουν την ετοιμότητα", level: 1 },
    { text: "1.3 Θεωρίες μάθησης (Kolb, Ericsson, CRM)", level: 1 },
    { text: "1.4 Η εκπαίδευση σε επαγγέλματα υψηλού κινδύνου", level: 1 },
    { text: "ΚΕΦΑΛΑΙΟ 2: ΤΑ ΣΩΜΑΤΑ ΑΣΦΑΛΕΙΑΣ ΚΑΙ ΟΙ ΕΝΟΠΛΕΣ ΔΥΝΑΜΕΙΣ", level: 0 },
    { text: "2.1 Οι Ένοπλες Δυνάμεις", level: 1 },
    { text: "2.2 Η Ελληνική Αστυνομία (ΕΛ.ΑΣ.)", level: 1 },
    { text: "2.3 Η Τροχαία Αστυνομία", level: 1 },
    { text: "2.4 Το Πυροσβεστικό Σώμα", level: 1 },
    { text: "2.5 Κοινές επιχειρησιακές απαιτήσεις", level: 1 },
    { text: "ΚΕΦΑΛΑΙΟ 3: ΕΚΠΑΙΔΕΥΤΙΚΑ ΣΥΣΤΗΜΑΤΑ ΚΑΙ ΜΟΡΦΕΣ ΑΣΚΗΣΕΩΝ", level: 0 },
    { text: "3.1 Η στρατιωτική εκπαίδευση", level: 1 },
    { text: "3.2 Η αστυνομική εκπαίδευση", level: 1 },
    { text: "3.3 Η εκπαίδευση της Τροχαίας", level: 1 },
    { text: "3.4 Η εκπαίδευση στο Πυροσβεστικό Σώμα", level: 1 },
    { text: "3.5 Πολυφορεακές ασκήσεις", level: 1 },
    { text: "ΚΕΦΑΛΑΙΟ 4: Η ΕΠΙΔΡΑΣΗ ΤΗΣ ΕΚΠΑΙΔΕΥΣΗΣ ΣΤΗΝ ΕΤΟΙΜΟΤΗΤΑ", level: 0 },
    { text: "4.1 Σωματική ετοιμότητα", level: 1 },
    { text: "4.2 Τεχνική επάρκεια", level: 1 },
    { text: "4.3 Ψυχολογική ανθεκτικότητα", level: 1 },
    { text: "4.4 Ομαδική συνεργασία και συντονισμός", level: 1 },
    { text: "4.5 Συχνότητα ασκήσεων και ετοιμότητα", level: 1 },
    { text: "ΚΕΦΑΛΑΙΟ 5: CASE STUDIES ΚΑΙ ΔΙΕΘΝΗΣ ΣΥΝΕΡΓΑΣΙΑ", level: 0 },
    { text: "5.1 Πλαίσια διεθνούς συνεργασίας", level: 1 },
    { text: "5.2 CS1: Πυρκαγιά Μάτι (2018)", level: 1 },
    { text: "5.3 CS2: NATO Defender Europe", level: 1 },
    { text: "5.4 CS3: MODEX / ΕΜΑΚ", level: 1 },
    { text: "5.5 CS4: ATLAS Network αντιτρομοκρατίας", level: 1 },
    { text: "ΚΕΦΑΛΑΙΟ 6: Ο ΡΟΛΟΣ ΤΗΣ ΤΕΧΝΟΛΟΓΙΑΣ", level: 0 },
    { text: "6.1 Προσομοιωτές", level: 1 },
    { text: "6.2 VR/AR", level: 1 },
    { text: "6.3 Drones", level: 1 },
    { text: "6.4 GIS, e-learning", level: 1 },
    { text: "ΣΥΜΠΕΡΑΣΜΑΤΑ ΚΑΙ ΠΡΟΤΑΣΕΙΣ", level: 0 },
    { text: "ΒΙΒΛΙΟΓΡΑΦΙΑ", level: 0 },
  ];
  for (const e of tocEntries) {
    allElements.push(new Paragraph({
      children: [tr(e.text, { size: e.level === 0 ? 24 : 22, bold: e.level === 0 })],
      spacing: { after: e.level === 0 ? 80 : 40, line: 300 },
      indent: e.level === 1 ? { left: 720 } : undefined,
    }));
  }

  // --- CHAPTERS ---
  for (const chapter of allContentChapters) {
    allElements.push(...buildChapter(chapter));
  }

  // --- BIBLIOGRAPHY (unnumbered) ---
  allElements.push(heading1("ΒΙΒΛΙΟΓΡΑΦΙΑ", true));
  for (const entry of bibliographyEntries) {
    allElements.push(new Paragraph({
      children: [tr(entry, { size: 22 })],
      spacing: { after: 120, line: 300 },
      alignment: AlignmentType.JUSTIFIED,
      indent: { left: 720, hanging: 720 },
    }));
  }

  // --- CREATE DOCUMENT ---
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: MARGINS,
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({ children: [tr("", { size: 18 })], alignment: AlignmentType.RIGHT })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 20 })],
            alignment: AlignmentType.RIGHT,
          })],
        }),
      },
      children: allElements,
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync("./output/diplomatiki.docx", buffer);
  console.log("✓ diplomatiki.docx generated");

  // Count chars for page estimate (with 1.5 spacing, ~1800 chars/page)
  let totalChars = perilepsi.perilepsi.length + perilepsi.abstract_en.length;
  for (const ch of allContentChapters) {
    for (const s of ch.sections) {
      for (const p of s.paragraphs) totalChars += p.length;
    }
  }
  const estPages = Math.round(totalChars / 1800);
  console.log(`  Est. text pages: ~${estPages} (+ cover, TOC, tables, bibliography ~${estPages + 8})`);
  console.log(`  Sources: ${bibliographyEntries.length}`);
  console.log(`  Chapters: ${numberedChapters.length} numbered + intro + conclusions`);
}

generateDocument().catch(err => { console.error(err); process.exit(1); });
