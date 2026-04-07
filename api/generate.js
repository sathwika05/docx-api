const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, BorderStyle, WidthType, VerticalAlign,
  TabStopType
} = require('docx');

const DARK_BLUE = "01144c";
const BLACK     = "000000";
const CALIBRI   = "Calibri";
const PT = n => n * 2;
const CONTENT_WIDTH_DXA = 10080;

function safeArray(v) { return Array.isArray(v) ? v : []; }

function nameParagraph(name) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: name, font: CALIBRI, size: PT(24), bold: true, color: BLACK })]
  });
}

function contactParagraph(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 40 },
    children: [new TextRun({ text, font: CALIBRI, size: PT(11), color: BLACK })]
  });
}

function sectionTitle(text) {
  return new Paragraph({
    spacing: { before: 120, after: 40 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: DARK_BLUE, space: 1 } },
    children: [new TextRun({ text, font: CALIBRI, size: PT(11), bold: true, color: DARK_BLUE })]
  });
}

function summaryParagraph(text) {
  return new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text, font: CALIBRI, size: PT(11), color: BLACK })]
  });
}

function skillRow(label, value) {
  return new Paragraph({
    spacing: { before: 0, after: 20 },
    children: [
      new TextRun({ text: `${label}: `, font: CALIBRI, size: PT(11), bold: true, color: BLACK }),
      new TextRun({ text: value || '', font: CALIBRI, size: PT(10.5), color: BLACK })
    ]
  });
}

function bulletItem(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 0, after: 40 },
    children: [new TextRun({ text, font: CALIBRI, size: PT(10.5), color: BLACK })]
  });
}

function titleDateRow(title, date) {
  return new Paragraph({
    spacing: { before: 60, after: 0 },
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_WIDTH_DXA }],
    children: [
      new TextRun({ text: title, font: CALIBRI, size: PT(11), bold: true, color: BLACK }),
      new TextRun({ text: `\t${date}`, font: CALIBRI, size: PT(11), color: BLACK })
    ]
  });
}

function projectTitle(text) {
  return new Paragraph({
    spacing: { before: 60, after: 0 },
    children: [new TextRun({ text, font: CALIBRI, size: PT(11), bold: true, color: BLACK })]
  });
}

function educationRow(institution, location, degree, date) {
  const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const borders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
  return new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [7560, 2520],
    borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideH: noBorder, insideV: noBorder },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders,
            width: { size: 7560, type: WidthType.DXA },
            verticalAlign: VerticalAlign.TOP,
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
            children: [new Paragraph({
              spacing: { before: 60, after: 0 },
              children: [
                new TextRun({ text: institution, font: CALIBRI, size: PT(10.5), bold: true, color: BLACK }),
                new TextRun({ text: ` (${location})`, font: CALIBRI, size: PT(10.5), color: BLACK }),
                new TextRun({ text: ` - ${degree}`, font: CALIBRI, size: PT(10.5), color: BLACK }),
              ]
            })]
          }),
          new TableCell({
            borders,
            width: { size: 2520, type: WidthType.DXA },
            verticalAlign: VerticalAlign.TOP,
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
            children: [new Paragraph({
              alignment: AlignmentType.RIGHT,
              spacing: { before: 60, after: 0 },
              children: [new TextRun({ text: date, font: CALIBRI, size: PT(11), color: BLACK })]
            })]
          })
        ]
      })
    ]
  });
}

function buildDoc(d) {
  const children = [];

  // HEADER
  children.push(nameParagraph("Sathwika Parshaboina"));
  children.push(contactParagraph("937.815.4324 | sathwikap25@gmail.com"));
  children.push(contactParagraph("sathwikap.com | linkedin.com/in/sathwikaparshaboina/"));

  // SUMMARY
  children.push(sectionTitle("Summary"));
  children.push(summaryParagraph(d.summary || ''));

  // SKILLS
  children.push(sectionTitle("Skills"));
  children.push(skillRow("Backend",                         d.skills?.backend));
  children.push(skillRow("Frontend",                        d.skills?.frontend));
  children.push(skillRow("AI/LLM",                          d.skills?.ai));
  children.push(skillRow("Testing & Quality",                d.skills?.testing));
  children.push(skillRow("Version Control & Collaboration",  d.skills?.collaboration));
  children.push(skillRow("DevOps & Deployment",              d.skills?.devops));

  // EXPERIENCE
  children.push(sectionTitle("Experience"));
  children.push(titleDateRow("AI Full Stack Engineer - Freelance", "July 2025 - Present"));
  safeArray(d.freelance_bullets).forEach(b => children.push(bulletItem(b)));
  children.push(titleDateRow("Software Engineer, Avis Budget Group – Parsippany, NJ (Hybrid)", "June 2022 – July 2025"));
  safeArray(d.avis_bullets).forEach(b => children.push(bulletItem(b)));

  // PROJECTS
  children.push(sectionTitle("Projects"));
  children.push(projectTitle("AI-Vid"));
  safeArray(d.projects?.ai_vid_bullets).forEach(b => children.push(bulletItem(b)));
  children.push(projectTitle("AI-Appointment-Intake-Workflow"));
  safeArray(d.projects?.ai_appointment_bullets).forEach(b => children.push(bulletItem(b)));
  children.push(projectTitle("AI Task Management System"));
  safeArray(d.projects?.ai_task_bullets).forEach(b => children.push(bulletItem(b)));
  children.push(projectTitle("AI DocCrawler"));
  safeArray(d.projects?.ai_doccrawler_bullets).forEach(b => children.push(bulletItem(b)));

  // EDUCATION
  children.push(sectionTitle("Education"));
  children.push(educationRow("Harrisburg University", "Pennsylvania, USA", "Executive Master of Science in AI for Business", "Aug 2025 - Present"));
  children.push(educationRow("Wright State University", "Ohio, USA", "Master of Science in Computer Science", "Jan 2021 - Jul 2022"));

  return new Document({
    numbering: {
      config: [{
        reference: "bullets",
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 360 } } }
        }]
      }]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 720, right: 1080, bottom: 720, left: 1080 }
        }
      },
      children
    }]
  });
}

// ── Vercel handler ────────────────────────────────────────────────────────────
module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed. Use POST.' });
  }

  try {
    const d = req.body;
    const doc = buildDoc(d);
    const buffer = await Packer.toBuffer(doc);
    const base64 = buffer.toString('base64');

    res.status(200).json({
      base64,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      fileName: 'resume.docx'
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
