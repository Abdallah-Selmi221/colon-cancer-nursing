const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

// Colors - Medical Academic Palette
const BURGUNDY = '8B1A2D';
const DARK_BG = '1A1A2E';
const CREAM = 'FFF8F0';
const LIGHT_PINK = 'FBF4F5';
const GOLD = 'C89D5C';
const WHITE = 'FFFFFF';
const DARK_TEXT = '222222';
const BODY_TEXT = '333333';
const SUBTITLE_GRAY = '666666';
const CARD_BG = 'F8F9FA';
const SUCCESS_GREEN = '2E7D32';
const NAVY_BLUE = '1A365D';

// Paths
const DOCX_MEDIA = 'C:\\Users\\ACER\\.gemini\\antigravity\\brain\\a3de671f-0b32-4188-9a89-a83f29331cc0\\scratch\\docx_media';
const PPTX_MEDIA = 'C:\\Users\\ACER\\.gemini\\antigravity\\brain\\a3de671f-0b32-4188-9a89-a83f29331cc0\\scratch\\pptx_media';
const PROJECT_IMAGES = path.join(__dirname, '..', 'public', 'images');

function getBase64Image(filePath) {
  if (!filePath || !fs.existsSync(filePath)) return null;
  const ext = path.extname(filePath).toLowerCase();
  let mime = 'image/png';
  if (ext === '.jpg' || ext === '.jpeg') mime = 'image/jpeg';
  const data = fs.readFileSync(filePath);
  return `data:${mime};base64,${data.toString('base64')}`;
}

// Logos
const nursingLogo = getBase64Image(path.join(PROJECT_IMAGES, 'nursing-logo-opt.png')) ||
                    getBase64Image(path.join(DOCX_MEDIA, 'image1.jpeg'));
const deptLogo = getBase64Image(path.join(PROJECT_IMAGES, 'dept-logo-opt.png')) ||
                 getBase64Image(path.join(DOCX_MEDIA, 'image2.jpeg'));
const azharLogo = getBase64Image(path.join(PROJECT_IMAGES, 'azhar-logo-opt.png')) ||
                  getBase64Image(path.join(DOCX_MEDIA, 'image3.jpeg'));

// Medical Images
const coverIllustration = getBase64Image(path.join(DOCX_MEDIA, 'image2.png'));
const colonAnatomyDiagram = getBase64Image(path.join(DOCX_MEDIA, 'image4.jpeg'));
const colonoscopyImg = getBase64Image(path.join(PPTX_MEDIA, 'image8.jpeg'));
const histopathologyImg = getBase64Image(path.join(PPTX_MEDIA, 'image9.jpeg'));
const pathoProgressionImg = getBase64Image(path.join(PPTX_MEDIA, 'image11.jpeg'));
const ctScanImg = getBase64Image(path.join(PPTX_MEDIA, 'image12.jpeg'));
const stagingDiagramImg = getBase64Image(path.join(PPTX_MEDIA, 'image13.jpeg'));
const surgeryAtlasImg = getBase64Image(path.join(PPTX_MEDIA, 'image14.jpeg'));
const nurseVitalsImg = getBase64Image(path.join(PPTX_MEDIA, 'image15.jpeg'));
const nurseCareImg = getBase64Image(path.join(PPTX_MEDIA, 'image16.jpeg'));

const TOTAL_SLIDES = 36;

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5 inches
pptx.author = 'Faculty of Nursing - Al-Azhar University';
pptx.subject = 'Colon Cancer Medical-Surgical Nursing Presentation';
pptx.title = 'Colon Cancer - Academic Seminar & Nursing Care Plans';

// Helper to add institutional header (PERFECTLY PADDED, NO CLIPPING)
function addHeader(slide, title, slideNum) {
  // Top thin accent bar
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.06, fill: { color: BURGUNDY } });

  // Institutional Header Strip (Height 0.94)
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.06, w: '100%', h: 0.94, fill: { color: LIGHT_PINK } });
  slide.addShape(pptx.ShapeType.line, { x: 0, y: 1.00, w: '100%', h: 0, line: { color: 'E2CCD0', width: 1 } });

  // 3 Logos in header (Well-padded, vertically centered: y: 0.15, h: 0.75, top margin 0.09, bottom margin 0.10)
  if (nursingLogo) {
    slide.addImage({ data: nursingLogo, x: 0.40, y: 0.15, w: 0.75, h: 0.75 });
  }
  if (deptLogo) {
    slide.addImage({ data: deptLogo, x: 1.30, y: 0.15, w: 0.75, h: 0.75 });
  }
  if (azharLogo) {
    slide.addImage({ data: azharLogo, x: 12.18, y: 0.15, w: 0.75, h: 0.75 });
  }

  // Institutional Bilingual Typography
  slide.addText([
    { text: 'Faculty of Nursing for Girls — Al-Azhar University\n', options: { fontSize: 10, bold: true, color: BURGUNDY } },
    { text: 'Medical-Surgical Nursing Department  •  قسم التمريض الباطني والجراحي\n', options: { fontSize: 9.5, color: SUBTITLE_GRAY } },
    { text: 'كُليّة التمريض للبنات بالقاهرة  •  جامعة الأزهر الشريف', options: { fontSize: 9.5, bold: true, color: BURGUNDY } }
  ], { x: 2.15, y: 0.12, w: 9.9, h: 0.82, align: 'center', valign: 'middle', fontFace: 'Calibri' });

  // Slide Title Banner (y: 1.00 to 1.62)
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 1.00, w: '100%', h: 0.62, fill: { color: BURGUNDY } });
  slide.addText(title, {
    x: 0.5, y: 1.00, w: 10.8, h: 0.62,
    fontSize: 28, bold: true, color: WHITE,
    align: 'left', valign: 'middle', fontFace: 'Calibri'
  });

  // Slide Counter
  slide.addText(`${slideNum} / ${TOTAL_SLIDES}`, {
    x: 11.4, y: 1.00, w: 1.4, h: 0.62,
    fontSize: 15, bold: true, color: 'E2B6BF',
    align: 'right', valign: 'middle', fontFace: 'Calibri'
  });

  // Slide Footer (y: 7.15 to 7.50)
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 7.15, w: '100%', h: 0.35, fill: { color: BURGUNDY } });
  slide.addText('Faculty of Nursing for Girls • Al-Azhar University — Medical-Surgical Nursing Department (2026/2027)', {
    x: 0.5, y: 7.15, w: 12.33, h: 0.35,
    fontSize: 9, color: 'F0D5DA', align: 'center', valign: 'middle', fontFace: 'Calibri'
  });
}

console.log('Generating perfectly-sized slides...');

// ==========================================
// SLIDE 1: COVER
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: DARK_BG };

  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: GOLD } });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 7.42, w: '100%', h: 0.08, fill: { color: GOLD } });

  if (nursingLogo) slide.addImage({ data: nursingLogo, x: 0.6, y: 0.25, w: 1.1, h: 1.1 });
  if (deptLogo) slide.addImage({ data: deptLogo, x: 1.85, y: 0.25, w: 1.1, h: 1.1 });
  if (azharLogo) slide.addImage({ data: azharLogo, x: 11.63, y: 0.25, w: 1.1, h: 1.1 });

  slide.addText([
    { text: 'Faculty of Nursing for Girls — Al-Azhar University\n', options: { fontSize: 14, bold: true, color: GOLD } },
    { text: 'كُليّة التمريض للبنات بالقاهرة  •  جامعة الأزهر الشريف\n', options: { fontSize: 13, bold: true, color: CREAM } },
    { text: 'Department of Medical-Surgical Nursing  •  قسم التمريض الباطني والجراحي', options: { fontSize: 12, color: 'CCCCCC' } }
  ], { x: 3.1, y: 0.25, w: 8.4, h: 1.1, align: 'center', valign: 'middle', fontFace: 'Calibri' });

  slide.addShape(pptx.ShapeType.line, { x: 0.6, y: 1.45, w: 12.13, h: 0, line: { color: GOLD, width: 1.5 } });

  slide.addText('COLON CANCER', {
    x: 0.6, y: 1.6, w: 7.2, h: 0.95,
    fontSize: 48, bold: true, color: WHITE,
    align: 'left', valign: 'middle', fontFace: 'Calibri', charSpacing: 3
  });

  slide.addText('A Comprehensive Clinical & Evidence-Based Nursing Seminar', {
    x: 0.6, y: 2.6, w: 7.2, h: 0.45,
    fontSize: 18, italic: true, color: GOLD, fontFace: 'Calibri'
  });

  // Supervision
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 3.2, w: 7.2, h: 1.7,
    fill: { color: '232338' }, line: { color: GOLD, width: 1 }, rectRadius: 0.08
  });
  slide.addText('UNDER SUPERVISION', {
    x: 0.8, y: 3.3, w: 6.8, h: 0.3,
    fontSize: 13, bold: true, color: GOLD, fontFace: 'Calibri', charSpacing: 2
  });
  slide.addText([
    { text: 'Prof. Dr. Heba Abdel-Azem\n', options: { fontSize: 15, bold: true, color: WHITE } },
    { text: 'Professor of Medical-Surgical Nursing Department\n\n', options: { fontSize: 11.5, color: 'BBBBBB' } },
    { text: 'Dr. Eman Abel-Maqsoud\n', options: { fontSize: 15, bold: true, color: WHITE } },
    { text: 'Demonstrator of Medical-Surgical Nursing Department', options: { fontSize: 11.5, color: 'BBBBBB' } }
  ], { x: 0.8, y: 3.65, w: 6.8, h: 1.15, fontFace: 'Calibri' });

  // Students
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 5.05, w: 7.2, h: 1.8,
    fill: { color: '232338' }, line: { color: '3F3F5F', width: 1 }, rectRadius: 0.08
  });
  slide.addText('PREPARED BY (SECOND YEAR — 2026 / 2027):', {
    x: 0.8, y: 5.15, w: 6.8, h: 0.3,
    fontSize: 12, bold: true, color: GOLD, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '1. Hager Saeed        2. Hager Hany        3. Hager Ahmed\n', options: { fontSize: 13, bold: true, color: WHITE } },
    { text: '4. Hager Mostafa      5. Haidy Mohammed    6. Norhan Badwy', options: { fontSize: 13, bold: true, color: WHITE } }
  ], { x: 0.8, y: 5.55, w: 6.8, h: 0.7, fontFace: 'Calibri' });
  slide.addText('Faculty of Nursing • Al-Azhar University', {
    x: 0.8, y: 6.4, w: 6.8, h: 0.35,
    fontSize: 11.5, color: 'AAAAAA', italic: true, fontFace: 'Calibri'
  });

  if (coverIllustration) {
    slide.addImage({ data: coverIllustration, x: 8.1, y: 1.7, w: 4.6, h: 5.1, rounding: true });
    slide.addText('Colonic Oncology & Tumor Pathology', {
      x: 8.1, y: 6.9, w: 4.6, h: 0.35,
      fontSize: 11, color: GOLD, align: 'center', italic: true, fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 2: PRESENTATION TEAM
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PRESENTATION TEAM — MEET THE AUTHORS', 2);

  slide.addText('Seminar Prepared By Second-Year Nursing Students (2026/2027)', {
    x: 0.6, y: 1.75, w: 12.13, h: 0.4,
    fontSize: 20, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });

  const students = [
    { num: '01', name: 'Hager Saeed', role: 'Nursing Student — Al-Azhar University' },
    { num: '02', name: 'Hager Hany', role: 'Nursing Student — Al-Azhar University' },
    { num: '03', name: 'Hager Ahmed', role: 'Nursing Student — Al-Azhar University' },
    { num: '04', name: 'Hager Mostafa', role: 'Nursing Student — Al-Azhar University' },
    { num: '05', name: 'Haidy Mohammed', role: 'Nursing Student — Al-Azhar University' },
    { num: '06', name: 'Norhan Badwy', role: 'Nursing Student — Al-Azhar University' },
  ];

  const cardW = 3.8;
  const cardH = 1.7;
  const startX = 0.6;
  const startY = 2.3;
  const gapX = 0.36;
  const gapY = 0.4;

  students.forEach((s, idx) => {
    const col = idx % 3;
    const row = Math.floor(idx / 3);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);

    slide.addShape(pptx.ShapeType.rect, {
      x, y, w: cardW, h: cardH,
      fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1.5 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.2, y: y + 0.35, w: 0.85, h: 0.85,
      fill: { color: BURGUNDY }
    });
    slide.addText(s.num, {
      x: x + 0.2, y: y + 0.35, w: 0.85, h: 0.85,
      fontSize: 20, bold: true, color: WHITE, align: 'center', valign: 'middle', fontFace: 'Calibri'
    });
    slide.addText(s.name, {
      x: x + 1.2, y: y + 0.3, w: cardW - 1.35, h: 0.45,
      fontSize: 22, bold: true, color: DARK_TEXT, fontFace: 'Calibri'
    });
    slide.addText(s.role, {
      x: x + 1.2, y: y + 0.75, w: cardW - 1.35, h: 0.6,
      fontSize: 13, color: SUBTITLE_GRAY, fontFace: 'Calibri'
    });
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 6.35, w: 12.13, h: 0.55,
    fill: { color: 'F5F5F5' }, line: { color: 'DDDDDD', width: 1 }, rectRadius: 0.08
  });
  slide.addText('Department of Medical-Surgical Nursing • Faculty of Nursing for Girls, Cairo • Al-Azhar University', {
    x: 0.6, y: 6.35, w: 12.13, h: 0.55,
    fontSize: 13, color: BURGUNDY, bold: true, align: 'center', valign: 'middle', fontFace: 'Calibri'
  });
}

// ==========================================
// SLIDE 3: GOAL
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'GOAL OF THE ASSIGNMENT', 3);

  slide.addShape(pptx.ShapeType.rect, {
    x: 1.0, y: 2.0, w: 11.33, h: 4.6,
    fill: { color: LIGHT_PINK }, line: { color: BURGUNDY, width: 2 }, rectRadius: 0.12
  });

  slide.addText('OVERARCHING EDUCATIONAL GOAL', {
    x: 1.5, y: 2.4, w: 10.33, h: 0.5,
    fontSize: 26, bold: true, color: BURGUNDY, align: 'center', fontFace: 'Calibri'
  });

  slide.addText(
    'The goal of this assignment is to provide nursing students with a comprehensive understanding of colon cancer, including its anatomy, pathophysiology, risk factors, clinical manifestations, diagnosis, and management, in order to enhance their knowledge and ability to provide safe, effective, and evidence-based nursing care for patients with colon cancer.',
    {
      x: 1.5, y: 3.1, w: 10.33, h: 3.1,
      fontSize: 24, color: DARK_TEXT, align: 'justify', valign: 'top',
      lineSpacingMultiple: 1.35, fontFace: 'Calibri'
    }
  );
}

// ==========================================
// SLIDE 4: OBJECTIVES
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'LEARNING OBJECTIVES', 4);

  slide.addText('By the end of this assignment, the nursing student will be able to:', {
    x: 0.6, y: 1.75, w: 12.13, h: 0.4,
    fontSize: 20, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });

  const objectives = [
    'Define colon cancer and describe the anatomy and physiology of the colon.',
    'Explain the pathophysiology, causes, and risk factors of colon cancer.',
    'Identify the types, signs and symptoms, and complications of colon cancer.',
    'Discuss the diagnostic tests, staging, and prognosis of colon cancer.',
    'Describe the medical and surgical management and the preventive measures of colon cancer.',
    'Describe the role of the nurse in caring for patients with colon cancer before and after surgery.',
    'Identify preventive measures and methods of early detection of colon cancer.',
    'Apply the nursing process in planning and providing individualized nursing care for patients with colon cancer.'
  ];

  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  objectives.slice(0, 4).forEach((obj, i) => {
    const y = 2.25 + i * 1.15;
    slide.addShape(pptx.ShapeType.rect, {
      x: startX, y, w: colW, h: 1.0,
      fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.ellipse, {
      x: startX + 0.15, y: y + 0.2, w: 0.6, h: 0.6,
      fill: { color: BURGUNDY }
    });
    slide.addText(`${i + 1}`, {
      x: startX + 0.15, y: y + 0.2, w: 0.6, h: 0.6,
      fontSize: 18, bold: true, color: WHITE, align: 'center', valign: 'middle'
    });
    slide.addText(obj, {
      x: startX + 0.85, y: y + 0.1, w: colW - 0.95, h: 0.8,
      fontSize: 16, color: DARK_TEXT, fontFace: 'Calibri', valign: 'middle'
    });
  });

  objectives.slice(4, 8).forEach((obj, i) => {
    const y = 2.25 + i * 1.15;
    slide.addShape(pptx.ShapeType.rect, {
      x: startX + colW + gapX, y, w: colW, h: 1.0,
      fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.ellipse, {
      x: startX + colW + gapX + 0.15, y: y + 0.2, w: 0.6, h: 0.6,
      fill: { color: BURGUNDY }
    });
    slide.addText(`${i + 5}`, {
      x: startX + colW + gapX + 0.15, y: y + 0.2, w: 0.6, h: 0.6,
      fontSize: 18, bold: true, color: WHITE, align: 'center', valign: 'middle'
    });
    slide.addText(obj, {
      x: startX + colW + gapX + 0.85, y: y + 0.1, w: colW - 0.95, h: 0.8,
      fontSize: 16, color: DARK_TEXT, fontFace: 'Calibri', valign: 'middle'
    });
  });
}

// ==========================================
// SLIDE 5: OUTLINE
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PRESENTATION OUTLINE', 5);

  const leftItems = [
    '01. Introduction',
    '02. Definition of Colon Cancer',
    '03. Anatomy and Physiology',
    '04. Pathophysiology (Molecular Pathways & Spread)',
    '05. Risk Factors (Non-Modifiable & Modifiable)',
    '06. Causes & Etiology',
    '07. Types of Colon Cancer',
    '08. Signs and Symptoms',
    '09. Complications of Colon Cancer'
  ];

  const rightItems = [
    '10. Diagnostic Tests & Clinical Evaluation',
    '11. Staging (TNM Classification)',
    '12. Prognosis & 5-Year Survival Rates',
    '13. Management and Treatment (Surgical & Medical)',
    '14. Prevention & Screening Strategies',
    '15. Comprehensive Role of the Nurse',
    '16. Nursing Process (4 Complete Care Plans)',
    '17. Summary & References'
  ];

  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 1.8, w: colW, h: 5.0,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.1
  });
  const leftRuns = leftItems.map(item => ({
    text: item + '\n',
    options: { fontSize: 18, bold: true, color: BURGUNDY, paraSpaceAfter: 10 }
  }));
  slide.addText(leftRuns, {
    x: startX + 0.35, y: 1.95, w: colW - 0.7, h: 4.7,
    fontFace: 'Calibri', valign: 'top'
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 1.8, w: colW, h: 5.0,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.1
  });
  const rightRuns = rightItems.map(item => ({
    text: item + '\n',
    options: { fontSize: 18, bold: true, color: BURGUNDY, paraSpaceAfter: 14 }
  }));
  slide.addText(rightRuns, {
    x: startX + colW + gapX + 0.35, y: 1.95, w: colW - 0.7, h: 4.7,
    fontFace: 'Calibri', valign: 'top'
  });
}

// ==========================================
// SLIDE 6: INTRODUCTION
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'INTRODUCTION', 6);

  slide.addText([
    { text: 'Major Gastrointestinal Malignancy\n', options: { fontSize: 24, bold: true, color: BURGUNDY } },
    { text: '• One of the leading causes of cancer-related mortality worldwide.\n', options: { fontSize: 17, color: BODY_TEXT } },
    { text: '• Develops from uncontrolled epithelial proliferation in the colon mucosal lining.\n\n', options: { fontSize: 17, color: BODY_TEXT } },
    { text: 'Origin from Precancerous Polyps\n', options: { fontSize: 24, bold: true, color: BURGUNDY } },
    { text: '• Most cases originate from adenomatous polyps over a 10–15 year span.\n', options: { fontSize: 17, color: BODY_TEXT } },
    { text: '• Early polyp detection and endoscopic polypectomy prevent malignant progression.\n\n', options: { fontSize: 17, color: BODY_TEXT } },
    { text: 'The Imperative for Screening\n', options: { fontSize: 24, bold: true, color: BURGUNDY } },
    { text: '• Early stages are frequently completely asymptomatic.\n', options: { fontSize: 17, color: BODY_TEXT } },
    { text: '• Can develop in individuals with NO known hereditary or lifestyle risk factors.', options: { fontSize: 17, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.8, w: 7.4, h: 5.0,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });

  if (colonoscopyImg) {
    slide.addImage({ data: colonoscopyImg, x: 8.3, y: 1.9, w: 4.4, h: 4.6, rounding: true });
    slide.addText('Endoscopic Colonoscopy View: Precancerous Adenomatous Polyp', {
      x: 8.3, y: 6.55, w: 4.4, h: 0.4,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 7: DEFINITION
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'DEFINITION OF COLON CANCER', 7);

  slide.addText([
    { text: 'Pathological Definition\n', options: { fontSize: 24, bold: true, color: BURGUNDY } },
    { text: '• Malignant Neoplasm: A primary malignant disease arising in the tissues of the colon (the longest portion of the large intestine).\n\n', options: { fontSize: 17, color: BODY_TEXT } },
    { text: 'Cellular Mechanism\n', options: { fontSize: 24, bold: true, color: BURGUNDY } },
    { text: '• Uncontrolled Proliferation: Normal colonic mucosal cells acquire cumulative genetic mutations, lose differentiation, and form an invasive tumor.\n\n', options: { fontSize: 17, color: BODY_TEXT } },
    { text: 'Histological Classification (~95% Adenocarcinoma)\n', options: { fontSize: 24, bold: true, color: BURGUNDY } },
    { text: '• Glandular Origin: Vast majority arise from mucus-producing epithelial gland cells.\n', options: { fontSize: 17, color: BODY_TEXT } },
    { text: '• Invasive Behavior: Malignant glands invade through muscularis mucosae into submucosa and deeper layers (NCI).', options: { fontSize: 17, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.8, w: 7.4, h: 5.0,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });

  if (histopathologyImg) {
    slide.addImage({ data: histopathologyImg, x: 8.3, y: 1.9, w: 4.4, h: 4.6, rounding: true });
    slide.addText('Histopathology (H&E Stain): Invasive Colonic Adenocarcinoma', {
      x: 8.3, y: 6.55, w: 4.4, h: 0.4,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 8: ANATOMY (4 SECTIONS)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'ANATOMY OF THE COLON — 4 MAIN SECTIONS', 8);

  slide.addText([
    { text: 'Anatomical Overview (~1.5 m in Length)\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Forms the bulk of the large intestine, extending from the cecum to the rectum.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '1. Ascending Colon (Right Side)\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Runs upward from cecum to hepatic flexure; reabsorbs liquid content.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '2. Transverse Colon (Upper Abdomen)\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Crosses horizontally from hepatic flexure to splenic flexure.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '3. Descending Colon (Left Side)\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Descends along left flank toward pelvis as stool solidifies.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '4. Sigmoid Colon (Pelvic S-Shape)\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Connects to rectum; frequent site for polyps and obstructive tumors.', options: { fontSize: 16, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.75, w: 7.2, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });

  if (colonAnatomyDiagram) {
    slide.addImage({ data: colonAnatomyDiagram, x: 8.1, y: 1.8, w: 4.6, h: 4.8, rounding: true });
    slide.addText('Medical Atlas: The 4 Sections of the Colon & Tumor Inset', {
      x: 8.1, y: 6.65, w: 4.6, h: 0.35,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 9: PHYSIOLOGY & WALL LAYERS
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PHYSIOLOGY & HISTOLOGICAL WALL LAYERS', 9);

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 1.75, w: 12.13, h: 1.4,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.08
  });
  slide.addText([
    { text: 'Core Physiological Functions of the Colon:\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Water & Electrolyte Reabsorption: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Absorbs 1–1.5 L of water, Na+, and Cl- daily, maintaining fluid homeostasis.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Stool Consolidation & Storage: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Converts liquid chyme into formed stool lubricated by goblet-cell mucus.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Mass Peristaltic Motility: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Propels solid waste toward the rectum for controlled elimination.', options: { fontSize: 16, color: BODY_TEXT } }
  ], { x: 0.8, y: 1.85, w: 11.73, h: 1.2, fontFace: 'Calibri' });

  slide.addText('The 4 Histological Layers of the Colon Wall (TNM Staging):', {
    x: 0.6, y: 3.3, w: 12.13, h: 0.4,
    fontSize: 20, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });

  const layerTable = [
    [
      { text: 'Bowel Wall Layer', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 17 } },
      { text: 'Histological Structure', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 17 } },
      { text: 'Oncological Significance (TNM)', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 17 } }
    ],
    [
      { text: '1. Mucosa (Innermost)', options: { bold: true, color: BURGUNDY, fontSize: 15 } },
      { text: 'Columnar epithelium & mucus-producing goblet cells', options: { fontSize: 15, color: BODY_TEXT } },
      { text: 'Site of origin for >95% of colon cancers (Stage 0 / In Situ)', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '2. Submucosa', options: { bold: true, color: BURGUNDY, fontSize: 15 } },
      { text: 'Dense vascularized connective tissue with lymphatics', options: { fontSize: 15, color: BODY_TEXT } },
      { text: 'Invasion provides access to lymphatic channels (T1)', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '3. Muscularis Propria', options: { bold: true, color: BURGUNDY, fontSize: 15 } },
      { text: 'Inner circular & outer longitudinal smooth muscle', options: { fontSize: 15, color: BODY_TEXT } },
      { text: 'Deep muscle invasion defines Stage I (T2) tumor', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '4. Serosa (Outermost)', options: { bold: true, color: BURGUNDY, fontSize: 15 } },
      { text: 'Visceral peritoneum & surrounding pericolic fat', options: { fontSize: 15, color: BODY_TEXT } },
      { text: 'Penetration through wall defines Stage II (T3–T4)', options: { fontSize: 15, color: BODY_TEXT } }
    ]
  ];

  slide.addTable(layerTable, {
    x: 0.6, y: 3.8, w: 12.13,
    colW: [2.8, 4.5, 4.83],
    border: { type: 'solid', color: 'D8BFC4', pt: 1 },
    rowH: [0.55, 0.6, 0.6, 0.6, 0.6],
    autoPage: false
  });
}

// ==========================================
// SLIDE 10: PATHOPHYSIOLOGY (CIN SEQUENCE)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PATHOPHYSIOLOGY — ADENOMA-CARCINOMA SEQUENCE', 10);

  slide.addText([
    { text: 'Main Pathway: Chromosomal Instability (CIN, ~80%)\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: 'The classic multistep genetic model driving normal epithelium to invasive cancer:\n\n', options: { fontSize: 15, color: SUBTITLE_GRAY } },
    { text: '1. APC Inactivation (The "Gatekeeper" — Early Event):\n', options: { fontSize: 18, bold: true, color: BURGUNDY } },
    { text: '• Loss of APC → β-catenin accumulation → persistent Wnt signaling.\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Result: Normal mucosal hyperproliferation and small adenoma formation.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '2. KRAS Proto-Oncogene Mutation (Intermediate):\n', options: { fontSize: 18, bold: true, color: BURGUNDY } },
    { text: '• Constitutive intracellular growth signaling → adenoma enlargement.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '3. SMAD4 / TGF-β Pathway Loss (Late Adenoma):\n', options: { fontSize: 18, bold: true, color: BURGUNDY } },
    { text: '• Loss of cellular growth-inhibition checkpoints.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '4. TP53 Tumor Suppressor Inactivation (Malignant Transformation):\n', options: { fontSize: 18, bold: true, color: BURGUNDY } },
    { text: '• Loss of apoptosis & cell-cycle arrest → invasive adenocarcinoma.', options: { fontSize: 15, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.75, w: 7.4, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });

  if (pathoProgressionImg) {
    slide.addImage({ data: pathoProgressionImg, x: 8.3, y: 2.0, w: 4.4, h: 4.4, rounding: true });
    slide.addText('Adenoma-to-Carcinoma Progression Sequence', {
      x: 8.3, y: 6.55, w: 4.4, h: 0.35,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 11: PATHOPHYSIOLOGY (ALTERNATIVE PATHWAYS)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PATHOPHYSIOLOGY — ALTERNATIVE MOLECULAR PATHWAYS', 11);

  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  // Box 1: MMR / MSI
  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 1.8, w: colW, h: 4.9,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1.5 }, rectRadius: 0.1
  });
  slide.addText('MMR DEFICIENCY / MSI (~15%)', {
    x: startX + 0.3, y: 2.0, w: colW - 0.6, h: 0.4,
    fontSize: 20, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '• Defective Mismatch Repair Genes: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Mutations in MLH1, MSH2, MSH6, or PMS2 impair repair of DNA replication errors.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Microsatellite Instability (MSI-High): ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Unrepaired errors accumulate in short DNA repeats, mutating growth genes (TGF-βRII, BAX).\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Hereditary Form (Lynch Syndrome): ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Germline MMR mutation predisposing to early-onset colorectal and endometrial cancers.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Sporadic Form: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Acquired hypermethylation of the MLH1 promoter.', options: { fontSize: 16, color: BODY_TEXT } }
  ], { x: startX + 0.3, y: 2.5, w: colW - 0.6, h: 4.0, fontFace: 'Calibri' });

  // Box 2: Serrated Pathway
  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 1.8, w: colW, h: 4.9,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1.5 }, rectRadius: 0.1
  });
  slide.addText('SERRATED PATHWAY / CIMP (~10–20%)', {
    x: startX + colW + gapX + 0.3, y: 2.0, w: colW - 0.6, h: 0.4,
    fontSize: 20, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '• Activating BRAF Mutation: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Initiating oncogenic event driving serrated epithelial architecture.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• CpG Island Methylator (CIMP): ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Widespread promoter hypermethylation silences key tumor-suppressor genes (like MLH1).\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Precursor Morphology: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Originates from sessile serrated lesions (SSLs) rather than traditional adenomas.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Anatomical Location: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Predominantly located in the right (proximal) colon.', options: { fontSize: 16, color: BODY_TEXT } }
  ], { x: startX + colW + gapX + 0.3, y: 2.5, w: colW - 0.6, h: 4.0, fontFace: 'Calibri' });
}

// ==========================================
// SLIDE 12: PATHOPHYSIOLOGY (SPREAD & PROGRESSION)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PATHOPHYSIOLOGY — PROGRESSION & SPREAD', 12);

  slide.addText([
    { text: 'Hereditary Polyposis Syndromes\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Familial Adenomatous Polyposis (FAP): ', options: { fontSize: 17, bold: true, color: DARK_TEXT } },
    { text: 'Germline APC mutation causing hundreds to thousands of adenomatous polyps; nearly 100% lifetime cancer risk if colectomy is not performed.\n\n', options: { fontSize: 17, color: BODY_TEXT } },
    { text: 'Local Bowel Wall Penetration\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Sequential Transmural Invasion: ', options: { fontSize: 17, bold: true, color: DARK_TEXT } },
    { text: 'Mucosa → Submucosa (vessels/lymphatics) → Muscularis Propria → Serosa & Pericolic Fat → Adjacent Organ Invasion.\n\n', options: { fontSize: 17, color: BODY_TEXT } },
    { text: 'Metastatic Dissemination Routes\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Lymphatic Route: ', options: { fontSize: 17, bold: true, color: DARK_TEXT } },
    { text: 'Spreads to regional mesenteric and para-aortic lymph nodes.\n', options: { fontSize: 17, color: BODY_TEXT } },
    { text: '• Hematogenous Route (Portal Vein): ', options: { fontSize: 17, bold: true, color: DARK_TEXT } },
    { text: 'Spreads primarily to Liver (most common metastatic site), followed by Lungs.', options: { fontSize: 17, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.75, w: 12.13, h: 4.2,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 6.05, w: 12.13, h: 0.85,
    fill: { color: 'FFF9E6' }, line: { color: GOLD, width: 2 }, rectRadius: 0.08
  });
  slide.addText('⚡ Key Clinical Takeaway: Malignant progression takes 10–15 years, providing an effective window for screening colonoscopy and polyp removal to prevent cancer!', {
    x: 0.8, y: 6.15, w: 11.73, h: 0.65,
    fontSize: 16, bold: true, color: '8B6508', align: 'center', valign: 'middle', fontFace: 'Calibri'
  });
}

// ==========================================
// SLIDE 13: RISK FACTORS
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'RISK FACTORS FOR COLON CANCER', 13);

  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 1.8, w: colW, h: 4.3,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1.5 }, rectRadius: 0.1
  });
  slide.addText('NON-MODIFIABLE RISK FACTORS', {
    x: startX + 0.3, y: 1.95, w: colW - 0.6, h: 0.4,
    fontSize: 20, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '• Advanced Age: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Incidence rises sharply after 50 (~90% of cases).\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Family History: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'First-degree relative with CRC roughly doubles risk.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Hereditary Syndromes: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Lynch syndrome, FAP, MUTYH-associated polyposis.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Inflammatory Bowel Disease: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Ulcerative Colitis & Crohn\'s (chronic mucosal inflammation).\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Demographics: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Male sex; African American race (higher mortality).', options: { fontSize: 16, color: BODY_TEXT } }
  ], { x: startX + 0.3, y: 2.4, w: colW - 0.6, h: 3.5, fontFace: 'Calibri' });

  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 1.8, w: colW, h: 4.3,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1.5 }, rectRadius: 0.1
  });
  slide.addText('MODIFIABLE LIFESTYLE FACTORS', {
    x: startX + colW + gapX + 0.3, y: 1.95, w: colW - 0.6, h: 0.4,
    fontSize: 20, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '• Diet High in Red/Processed Meat: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Cured, salted, and charred meats; low dietary fiber.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Obesity & Sedentary Lifestyle: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Visceral adiposity promotes insulin resistance.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Tobacco & Heavy Alcohol Use: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Direct carcinogens promoting mucosal dysplasia.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Type 2 Diabetes & Metabolic Syndrome.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Protective Factors: ', options: { fontSize: 16, bold: true, color: SUCCESS_GREEN } },
    { text: 'Screening colonoscopy, high-fiber diet, active exercise.', options: { fontSize: 16, color: BODY_TEXT } }
  ], { x: startX + colW + gapX + 0.3, y: 2.4, w: colW - 0.6, h: 3.5, fontFace: 'Calibri' });

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 6.25, w: 12.13, h: 0.7,
    fill: { color: 'F5F5F5' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('Key Point: Most risk factors act by promoting chronic inflammation, insulin/IGF-1 signaling, and DNA damage.', {
    x: 0.8, y: 6.3, w: 11.73, h: 0.55,
    fontSize: 15, bold: true, color: DARK_TEXT, align: 'center', valign: 'middle', fontFace: 'Calibri'
  });
}

// ==========================================
// SLIDE 14: CAUSES (ETIOLOGY)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CAUSES & ETIOLOGY OF COLON CANCER', 14);

  slide.addText('Colon cancer has no single cause; it develops when genetic and environmental factors interact:', {
    x: 0.6, y: 1.75, w: 12.13, h: 0.4,
    fontSize: 20, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });

  const causes = [
    { num: '1', title: 'Sporadic Genetic Mutations (~70–75%)', desc: 'Acquired somatic mutations over lifetime (APC, KRAS, SMAD4, TP53) through the CIN pathway.' },
    { num: '2', title: 'Hereditary Syndromes (~5–10%)', desc: 'Inherited germline mutations: Lynch syndrome (MMR genes), FAP (APC gene), MUTYH-associated polyposis.' },
    { num: '3', title: 'Familial Clustering (~10–30%)', desc: 'Shared family incidence without a defined single-gene mutation (shared minor genes and environment).' },
    { num: '4', title: 'Chronic Intestinal Inflammation', desc: 'Long-standing IBD (Ulcerative colitis, Crohn\'s) causing repeated epithelial damage and dysplasia.' },
    { num: '5', title: 'Environmental DNA Damage', desc: 'Carcinogens in processed meats, tobacco smoke toxins, alcohol metabolites, and microbiome shifts.' }
  ];

  causes.forEach((c, idx) => {
    const y = 2.3 + idx * 0.92;
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.6, y, w: 12.13, h: 0.82,
      fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 0.8, y: y + 0.16, w: 0.5, h: 0.5,
      fill: { color: BURGUNDY }
    });
    slide.addText(c.num, {
      x: 0.8, y: y + 0.16, w: 0.5, h: 0.5,
      fontSize: 16, bold: true, color: WHITE, align: 'center', valign: 'middle'
    });
    slide.addText([
      { text: `${c.title}: `, options: { fontSize: 17, bold: true, color: BURGUNDY } },
      { text: c.desc, options: { fontSize: 16, color: BODY_TEXT } }
    ], { x: 1.45, y: y + 0.1, w: 11.1, h: 0.62, fontFace: 'Calibri', valign: 'middle' });
  });
}

// ==========================================
// SLIDE 15: TYPES OF COLON CANCER
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'HISTOLOGICAL TYPES OF COLON CANCER', 15);

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 1.75, w: 12.13, h: 1.7,
    fill: { color: LIGHT_PINK }, line: { color: BURGUNDY, width: 2 }, rectRadius: 0.08
  });
  slide.addText('ADENOCARCINOMA (~95% of cases) — The Predominant Type', {
    x: 0.9, y: 1.9, w: 11.53, h: 0.4,
    fontSize: 22, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '• Glandular Origin: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Arises from the mucus-secreting epithelial cells lining the colon mucosa.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Mucinous Adenocarcinoma: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Tumor composed of >50% extracellular mucin pooling.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Signet-Ring Cell Carcinoma: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Aggressive subtype with intracellular mucin pushing nucleus to periphery (poorer prognosis).', options: { fontSize: 16, color: BODY_TEXT } }
  ], { x: 0.9, y: 2.35, w: 11.53, h: 1.0, fontFace: 'Calibri' });

  slide.addText('Less Common Colorectal Neoplasms:', {
    x: 0.6, y: 3.65, w: 12.13, h: 0.35,
    fontSize: 20, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });

  const rareTypes = [
    { title: '1. Carcinoid (Neuroendocrine) Tumors', desc: 'Arise from specialized hormone-producing neuroendocrine cells.' },
    { title: '2. Gastrointestinal Stromal (GISTs)', desc: 'Originate in the pacemaker interstitial cells of Cajal in the bowel wall.' },
    { title: '3. Primary Colorectal Lymphomas', desc: 'Malignancies arising from immune-system lymphatic follicles in mucosa.' },
    { title: '4. Colorectal Sarcomas', desc: 'Rare cancers originating in smooth muscle layers or blood vessels.' }
  ];

  const cW = 5.85;
  const cH = 1.35;
  const startX = 0.6;
  const gapX = 0.43;

  rareTypes.forEach((t, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = startX + col * (cW + gapX);
    const y = 4.1 + row * (cH + 0.2);

    slide.addShape(pptx.ShapeType.rect, {
      x, y, w: cW, h: cH,
      fill: { color: 'F9F9F9' }, line: { color: 'DDDDDD', width: 1 }, rectRadius: 0.08
    });
    slide.addText(t.title, {
      x: x + 0.25, y: y + 0.15, w: cW - 0.5, h: 0.35,
      fontSize: 17, bold: true, color: BURGUNDY, fontFace: 'Calibri'
    });
    slide.addText(t.desc, {
      x: x + 0.25, y: y + 0.55, w: cW - 0.5, h: 0.7,
      fontSize: 15, color: BODY_TEXT, fontFace: 'Calibri'
    });
  });
}

// ==========================================
// SLIDE 16: SIGNS AND SYMPTOMS
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CLINICAL SIGNS AND SYMPTOMS', 16);

  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 1.75, w: colW, h: 4.3,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.1
  });
  slide.addText('LOCAL BOWEL MANIFESTATIONS', {
    x: startX + 0.3, y: 1.9, w: colW - 0.6, h: 0.35,
    fontSize: 18, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '1. Rectal Bleeding / Blood in Stool: ', options: { fontSize: 15, bold: true, color: DARK_TEXT } },
    { text: 'Bright red (hematochezia) or black tarry (melena).\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '2. Change in Bowel Habits: ', options: { fontSize: 15, bold: true, color: DARK_TEXT } },
    { text: 'Persistent diarrhea, worsening constipation, or alternating.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '3. Stool Caliber Narrowing: ', options: { fontSize: 15, bold: true, color: DARK_TEXT } },
    { text: 'Narrow, ribbon-like, or pencil-shaped stools.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '4. Tenesmus: ', options: { fontSize: 15, bold: true, color: DARK_TEXT } },
    { text: 'Persistent painful sensation of incomplete bowel evacuation.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '5. Abdominal Discomfort: ', options: { fontSize: 15, bold: true, color: DARK_TEXT } },
    { text: 'Persistent cramping, gas, bloating, or fullness.', options: { fontSize: 15, color: BODY_TEXT } }
  ], { x: startX + 0.3, y: 2.3, w: colW - 0.6, h: 3.6, fontFace: 'Calibri' });

  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 1.75, w: colW, h: 4.3,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.1
  });
  slide.addText('SYSTEMIC & ONCOLOGIC SIGNS', {
    x: startX + colW + gapX + 0.3, y: 1.9, w: colW - 0.6, h: 0.35,
    fontSize: 18, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '6. Unexplained Fatigue & Weakness: ', options: { fontSize: 15, bold: true, color: DARK_TEXT } },
    { text: 'Profound exhaustion secondary to chronic occult blood loss.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '7. Iron-Deficiency Anemia: ', options: { fontSize: 15, bold: true, color: DARK_TEXT } },
    { text: 'Microcytic hypochromic anemia (hallmark of right colon lesions).\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '8. Sudden Unexplained Weight Loss: ', options: { fontSize: 15, bold: true, color: DARK_TEXT } },
    { text: 'Involuntary drop in body mass without dietary changes.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '9. Anorexia & Early Satiety: ', options: { fontSize: 15, bold: true, color: DARK_TEXT } },
    { text: 'Loss of appetite and general malaise.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '10. Nausea and Vomiting: ', options: { fontSize: 15, bold: true, color: DARK_TEXT } },
    { text: 'Occurs with progressive intraluminal tumor obstruction.', options: { fontSize: 15, color: BODY_TEXT } }
  ], { x: startX + colW + gapX + 0.3, y: 2.3, w: colW - 0.6, h: 3.6, fontFace: 'Calibri' });

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 6.2, w: 12.13, h: 0.75,
    fill: { color: 'FFF3E0' }, line: { color: 'E65100', width: 1.5 }, rectRadius: 0.08
  });
  slide.addText('⚠️ CRITICAL ALERT: Early colon cancer is frequently ASYMPTOMATIC! Right-sided tumors often present only with occult anemia. Regular screening is essential for individuals ≥ 45 or with family history.', {
    x: 0.8, y: 6.25, w: 11.73, h: 0.65,
    fontSize: 14.5, bold: true, color: 'BF360C', align: 'center', valign: 'middle', fontFace: 'Calibri'
  });
}

// ==========================================
// SLIDE 17: COMPLICATIONS
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'MAJOR COMPLICATIONS OF COLON CANCER', 17);

  const complications = [
    { num: '1', title: 'Mechanical Bowel Obstruction', desc: 'Luminal tumor occlusion causing colicky pain, severe abdominal distension, vomiting, and obstipation.' },
    { num: '2', title: 'Gastrointestinal Bleeding', desc: 'Chronic bleeding leading to severe anemia; acute ulceration can cause brisk lower GI hemorrhage.' },
    { num: '3', title: 'Intestinal Perforation (Emergency)', desc: 'Tumor necrosis or high proximal pressure causing bowel rupture, peritonitis, and septic shock.' },
    { num: '4', title: 'Metastasis (Distant Spread)', desc: 'Dissemination via lymphatics and portal circulation to liver (most common), lungs, and peritoneum.' },
    { num: '5', title: 'Malnutrition and Cachexia', desc: 'Severe muscle wasting, anorexia, and metabolic dysfunction from chronic disease hypermetabolism.' },
    { num: '6', title: 'Psychosocial Distress', desc: 'Profound anxiety, depression, and body image disturbance related to diagnosis and stoma creation.' }
  ];

  complications.forEach((c, idx) => {
    const y = 1.75 + idx * 0.88;
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.6, y, w: 12.13, h: 0.8,
      fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 0.8, y: y + 0.15, w: 0.5, h: 0.5,
      fill: { color: BURGUNDY }
    });
    slide.addText(c.num, {
      x: 0.8, y: y + 0.15, w: 0.5, h: 0.5,
      fontSize: 16, bold: true, color: WHITE, align: 'center', valign: 'middle'
    });
    slide.addText([
      { text: `${c.title}: `, options: { fontSize: 17, bold: true, color: BURGUNDY } },
      { text: c.desc, options: { fontSize: 15, color: BODY_TEXT } }
    ], { x: 1.45, y: y + 0.1, w: 11.1, h: 0.6, fontFace: 'Calibri', valign: 'middle' });
  });
}

// ==========================================
// SLIDE 18: DIAGNOSTIC TESTS
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'DIAGNOSTIC WORKUP & CLINICAL EVALUATION', 18);

  slide.addText([
    { text: '1. Clinical History & Physical Exam:\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Symptom timeline, family history, abdominal palpation for masses, and Digital Rectal Exam (DRE).\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '2. Laboratory Investigations:\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• CBC: Identifies microcytic iron-deficiency anemia.\n• CMP: Evaluates liver and kidney baseline function.\n• CEA (Carcinoembryonic Antigen): Tumor marker for baseline and monitoring.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '3. Colonoscopy + Biopsy (GOLD STANDARD):\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Direct endoscopic visualization of entire colon and histopathological tissue sampling.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '4. Radiologic Staging Imaging:\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Contrast-enhanced CT of Chest/Abdomen/Pelvis, MRI, and PET scan.', options: { fontSize: 15, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.75, w: 7.4, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });

  if (ctScanImg) {
    slide.addImage({ data: ctScanImg, x: 8.3, y: 1.9, w: 4.4, h: 4.5, rounding: true });
    slide.addText('Abdominal CT Scan: Demonstrating Colonic Wall Thickening Mass', {
      x: 8.3, y: 6.55, w: 4.4, h: 0.35,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 19: STAGES OF COLON CANCER
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'STAGES OF COLON CANCER (TNM CLASSIFICATION)', 19);

  const stages = [
    { stage: 'Stage 0', title: 'Carcinoma in Situ', desc: 'Precancerous cells confined to mucosal lining (polyps).', color: '388E3C' },
    { stage: 'Stage I', title: 'Localized Invasion', desc: 'Invaded submucosa (T1) or muscularis (T2); no lymph nodes (N0, M0).', color: '689F38' },
    { stage: 'Stage II', title: 'Through Bowel Wall', desc: 'Grown deeper through wall (IIA/IIB); no lymph node spread (N0).', color: 'F57C00' },
    { stage: 'Stage III', title: 'Regional Lymph Nodes', desc: 'Spread to nearby regional mesenteric lymph nodes (N1–N2; M0).', color: 'E65100' },
    { stage: 'Stage IV', title: 'Distant Metastasis', desc: 'Metastasized to distant organs like liver, lungs, peritoneum (M1).', color: 'D32F2F' }
  ];

  stages.forEach((s, i) => {
    const y = 1.75 + i * 1.0;

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.6, y, w: 1.8, h: 0.88,
      fill: { color: s.color }, rectRadius: 0.08
    });
    slide.addText(s.stage, {
      x: 0.6, y, w: 1.8, h: 0.88,
      fontSize: 20, bold: true, color: WHITE, align: 'center', valign: 'middle', fontFace: 'Calibri'
    });

    slide.addShape(pptx.ShapeType.rect, {
      x: 2.5, y, w: stagingDiagramImg ? 5.8 : 10.23, h: 0.88,
      fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.08
    });
    slide.addText([
      { text: `${s.title}: `, options: { fontSize: 17, bold: true, color: BURGUNDY } },
      { text: s.desc, options: { fontSize: 15, color: BODY_TEXT } }
    ], { x: 2.7, y: y + 0.08, w: stagingDiagramImg ? 5.4 : 9.8, h: 0.72, fontFace: 'Calibri', valign: 'middle' });
  });

  if (stagingDiagramImg) {
    slide.addImage({ data: stagingDiagramImg, x: 8.6, y: 1.75, w: 4.1, h: 4.8, rounding: true });
    slide.addText('TNM Staging Depth & Dissemination', {
      x: 8.6, y: 6.65, w: 4.1, h: 0.35,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 20: PROGNOSIS & SURVIVAL RATES
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PROGNOSIS & 5-YEAR SURVIVAL RATES', 20);

  slide.addText('5-Year Relative Survival Rates by Stage at Diagnosis:', {
    x: 0.6, y: 1.75, w: 12.13, h: 0.4,
    fontSize: 22, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });

  const survivalData = [
    { stage: 'Localized Stage (Confined to Colon)', rate: '91%', val: 0.91, color: '2E7D32', text: 'Cancer has not spread beyond colon wall (Stage I & selected Stage II).' },
    { stage: 'Regional Stage (Spread to Lymph Nodes)', rate: '73%', val: 0.73, color: 'F57C00', text: 'Cancer spread to nearby regional lymph nodes or surrounding tissues (Stage III).' },
    { stage: 'Distant Stage (Metastatic Spread)', rate: '13%', val: 0.13, color: 'D32F2F', text: 'Cancer has metastasized to distant organs like liver or lungs (Stage IV).' }
  ];

  survivalData.forEach((d, idx) => {
    const y = 2.3 + idx * 1.35;

    slide.addText(d.stage, {
      x: 0.6, y, w: 7.5, h: 0.35,
      fontSize: 18, bold: true, color: DARK_TEXT, fontFace: 'Calibri'
    });
    slide.addText(d.text, {
      x: 0.6, y: y + 0.35, w: 7.5, h: 0.3,
      fontSize: 14, color: SUBTITLE_GRAY, fontFace: 'Calibri'
    });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.6, y: y + 0.7, w: 9.5, h: 0.45,
      fill: { color: 'EAEAEA' }, rectRadius: 0.06
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.6, y: y + 0.7, w: 9.5 * d.val, h: 0.45,
      fill: { color: d.color }, rectRadius: 0.06
    });

    slide.addText(d.rate, {
      x: 10.3, y: y + 0.6, w: 2.4, h: 0.65,
      fontSize: 26, bold: true, color: d.color, fontFace: 'Calibri', align: 'left', valign: 'middle'
    });
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 6.35, w: 12.13, h: 0.65,
    fill: { color: 'F5F5F5' }, line: { color: 'DDDDDD', width: 1 }, rectRadius: 0.08
  });
  slide.addText('Clinical Outlook: Early detection dramatically improves survival from 13% up to 91%! Regular surveillance is essential to monitor for recurrence.', {
    x: 0.8, y: 6.35, w: 11.73, h: 0.65,
    fontSize: 15, bold: true, color: BURGUNDY, align: 'center', valign: 'middle', fontFace: 'Calibri'
  });
}

// ==========================================
// SLIDE 21: MANAGEMENT AND TREATMENT
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'MANAGEMENT & TREATMENT MODALITIES', 21);

  slide.addText([
    { text: 'Surgical Management (Primary Curative Modality):\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Polypectomy: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Endoscopic removal of early malignant polyps during colonoscopy.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Colectomy: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Segmental resection with regional lymphadenectomy followed by anastomosis or stoma.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Multivisceral Resection: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'En-bloc resection for locally advanced tumors invading adjacent organs.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: 'Medical & Systemic Therapies:\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Chemotherapy: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Adjuvant therapy for Stage III and high-risk Stage II to eliminate micrometastases.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Targeted Therapy: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Anti-EGFR and anti-VEGF agents guided by RAS/BRAF biomarker testing.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Radiation Therapy: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Primarily used in rectal cancer or for palliative symptom relief.', options: { fontSize: 16, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.75, w: 7.4, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });

  if (surgeryAtlasImg) {
    slide.addImage({ data: surgeryAtlasImg, x: 8.3, y: 1.9, w: 4.4, h: 4.5, rounding: true });
    slide.addText('Surgical Partial Colectomy with Anastomosis', {
      x: 8.3, y: 6.55, w: 4.4, h: 0.35,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 22: WHEN TO SEEK MEDICAL ATTENTION
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'WHEN TO SEEK MEDICAL ATTENTION', 22);

  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 1.8, w: colW, h: 5.0,
    fill: { color: LIGHT_PINK }, line: { color: BURGUNDY, width: 1.5 }, rectRadius: 0.1
  });
  slide.addText('ALARM SYMPTOMS REQUIRING EVALUATION', {
    x: startX + 0.3, y: 1.95, w: colW - 0.6, h: 0.45,
    fontSize: 20, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Consult a healthcare provider promptly if experiencing:\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Persistent Abdominal Pain or Cramping\n\n', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: '• Blood in Stool (Visible Red or Dark Melena)\n\n', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: '• Persistent Change in Bowel Habits (> a few days)\n\n', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: '• Unexplained Weight Loss\n\n', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: '• Persistent Fatigue, Weakness, or Anemia', options: { fontSize: 18, bold: true, color: DARK_TEXT } }
  ], { x: startX + 0.3, y: 2.5, w: colW - 0.6, h: 4.1, fontFace: 'Calibri' });

  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 1.8, w: colW, h: 5.0,
    fill: { color: 'F5F7FA' }, line: { color: 'B0BEC5', width: 1.5 }, rectRadius: 0.1
  });
  slide.addText('CRITICAL QUESTIONS FOR ONCOLOGY TEAM', {
    x: startX + colW + gapX + 0.3, y: 1.95, w: colW - 0.6, h: 0.45,
    fontSize: 20, bold: true, color: '1A365D', fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'If diagnosed with colon cancer, ask:\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '1. What is the exact cancer stage?\n', options: { fontSize: 18, bold: true, color: '1A365D' } },
    { text: '   (Depth of invasion, lymph nodes, metastases)\n\n', options: { fontSize: 15, color: SUBTITLE_GRAY } },
    { text: '2. What are the chances of cure?\n', options: { fontSize: 18, bold: true, color: '1A365D' } },
    { text: '   (Curative surgical resection vs. systemic control)\n\n', options: { fontSize: 15, color: SUBTITLE_GRAY } },
    { text: '3. What is the risk of cancer recurrence?\n', options: { fontSize: 18, bold: true, color: '1A365D' } },
    { text: '   (Surveillance schedule: CEA, CT scans, colonoscopy)', options: { fontSize: 15, color: SUBTITLE_GRAY } }
  ], { x: startX + colW + gapX + 0.3, y: 2.5, w: colW - 0.6, h: 4.1, fontFace: 'Calibri' });
}

// ==========================================
// SLIDE 23: PREVENTION & SCREENING
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PREVENTION & EARLY DETECTION STRATEGIES', 23);

  slide.addText([
    { text: '1. Screening & Early Detection (Primary Prevention):\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Routine Screening Age: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Begin at age 45 for average-risk adults (earlier with family history/genetics).\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Gold Standard Colonoscopy: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Every 10 years if normal; enables polyp removal before malignancy develops.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '2. Lifestyle & Risk Reduction:\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Maintain healthy weight; engage in consistent moderate physical exercise.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Dietary changes: Eat vegetables, fruits, whole grains; strictly limit red & processed meats.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Cessation: Eliminate tobacco smoking and minimize/avoid alcohol intake.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '3. Supplements & Medical Evidence:\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Vitamins/Minerals: Folic acid, Vit D, Calcium, Magnesium show promise, but need medical guidance.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Aspirin/NSAIDs: May lower polyp risk, but NOT recommended for average risk due to bleeding risks.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Hormone Therapy: Not recommended as a standalone preventive measure.', options: { fontSize: 16, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.75, w: 12.13, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });
}

// ==========================================
// SLIDE 24: NURSING ROLE (PERIOPERATIVE)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'NURSING ROLE — PERIOPERATIVE CARE', 24);

  slide.addText([
    { text: '1. Comprehensive Preoperative Assessment:\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Assess bowel habits, blood loss, anemia, nutritional status, and baseline labs.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '2. Pre-operative Preparation & Stoma Planning:\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Bowel Prep: Administer prescribed mechanical/oral cleansers and antibiotics.\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Patient Education: Explain procedure, teach incentive spirometry, reduce anxiety.\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Stoma Marking: Collaborate with WOCN (wound/ostomy nurse) to mark stoma site.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '3. Post-operative Surveillance & Care:\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Monitoring: Vitals, incision, drain output, and early signs of leak/hemorrhage.\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Stoma Care: Ensure stoma is pink/red and moist; educate patient on pouch changes.\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Complication Prevention: Early ambulation, pain control, DVT prophylaxis.', options: { fontSize: 15, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.75, w: 7.4, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });

  if (nurseVitalsImg) {
    slide.addImage({ data: nurseVitalsImg, x: 8.3, y: 1.9, w: 4.4, h: 4.5, rounding: true });
    slide.addText('Clinical Bedside Post-Operative Patient Monitoring', {
      x: 8.3, y: 6.55, w: 4.4, h: 0.35,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 25: NURSING ROLE (SUPPORTIVE & REHAB)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'NURSING ROLE — TREATMENT & PSYCHOSOCIAL SUPPORT', 25);

  slide.addText([
    { text: '4. Managing Treatment Side Effects:\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Chemotherapy Support: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Manage nausea, vomiting, fatigue, and diarrhea with antiemetics and oral/IV fluids.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Infection & Hematologic Monitoring: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Monitor CBC closely (WBC, neutrophils, platelets) to assess infection and bleeding risks.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '5. Psychological & Rehabilitation Support:\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '• Body Image Adaptation: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Support patient coping and self-care skills with temporary or permanent colostomy.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Holistic & Family Care: ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'Provide continuous emotional, nutritional, and informational support to patient and family.', options: { fontSize: 16, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.75, w: 7.4, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.2
  });

  if (nurseCareImg) {
    slide.addImage({ data: nurseCareImg, x: 8.3, y: 1.9, w: 4.4, h: 4.5, rounding: true });
    slide.addText('Holistic Patient Rehabilitation & Psychological Counseling', {
      x: 8.3, y: 6.55, w: 4.4, h: 0.35,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 26: NURSING PROCESS (ROADMAP)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'NURSING PROCESS — CLINICAL CARE PLANS', 26);

  slide.addText('Application of the Nursing Process in Colorectal Oncology (4 Core Care Plans):', {
    x: 0.6, y: 1.75, w: 12.13, h: 0.4,
    fontSize: 22, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });

  const carePlans = [
    { num: 'CP 1', title: 'Actual Diagnosis: Abdominal Pain', desc: 'Related to tumor mass effect, intestinal distension, or surgical resection. Focus on pain assessment, analgesia, and bowel comfort.', color: BURGUNDY },
    { num: 'CP 2', title: 'Actual Diagnosis: Cancer-Related Fatigue', desc: 'Related to decreased oxygen-carrying capacity secondary to occult GI bleeding and anemia. Focus on energy conservation & nutrition.', color: BURGUNDY },
    { num: 'CP 3', title: 'Potential Diagnosis: Risk for Deficient Fluid Volume', desc: 'Related to excessive losses through vomiting, diarrhea, or stoma output and inadequate oral intake. Focus on strict I&O and hydration.', color: NAVY_BLUE },
    { num: 'CP 4', title: 'Potential Diagnosis: Risk for Dysfunctional GI Motility', desc: 'Related to colon tumor obstruction, luminal narrowing, or post-op ileus. Focus on bowel sound auscultation and obstruction prevention.', color: NAVY_BLUE }
  ];

  carePlans.forEach((cp, idx) => {
    const y = 2.3 + idx * 1.15;
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.6, y, w: 12.13, h: 1.0,
      fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1.5 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.8, y: y + 0.15, w: 1.0, h: 0.7,
      fill: { color: cp.color }, rectRadius: 0.06
    });
    slide.addText(cp.num, {
      x: 0.8, y: y + 0.15, w: 1.0, h: 0.7,
      fontSize: 16, bold: true, color: WHITE, align: 'center', valign: 'middle', fontFace: 'Calibri'
    });
    slide.addText([
      { text: `${cp.title}\n`, options: { fontSize: 18, bold: true, color: cp.color } },
      { text: cp.desc, options: { fontSize: 15, color: BODY_TEXT } }
    ], { x: 2.0, y: y + 0.1, w: 10.5, h: 0.8, fontFace: 'Calibri' });
  });
}

// ==========================================
// SLIDE 27: CARE PLAN 1: ABDOMINAL PAIN (DIAGNOSIS & GOALS)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 1: ABDOMINAL PAIN — ASSESSMENT & GOALS', 27);

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 1.75, w: 12.13, h: 1.8,
    fill: { color: LIGHT_PINK }, line: { color: BURGUNDY, width: 2 }, rectRadius: 0.08
  });
  slide.addText('ACTUAL NURSING DIAGNOSIS', {
    x: 0.9, y: 1.9, w: 11.53, h: 0.35,
    fontSize: 20, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Persistent abdominal pain ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'related to tumor growth and intestinal changes secondary to colon cancer, as evidenced by:\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Subjective Data: ', options: { fontSize: 16, bold: true, color: BURGUNDY } },
    { text: '"I have abdominal pain." • "The pain is constant and cramping."\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Objective Data: ', options: { fontSize: 16, bold: true, color: BURGUNDY } },
    { text: 'Abdominal tenderness on palpation, distension, tumor mass visible on imaging, changes in bowel habits.', options: { fontSize: 16, color: BODY_TEXT } }
  ], { x: 0.9, y: 2.3, w: 11.53, h: 1.15, fontFace: 'Calibri' });

  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 3.75, w: colW, h: 3.2,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EXPECTED OUTCOMES (GOALS)', {
    x: startX + 0.3, y: 3.9, w: colW - 0.6, h: 0.35,
    fontSize: 18, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Short-Term Goals:\n', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: '• Patient will report reduced abdominal pain within 24–48 hours.\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Patient will verbalize effective pain relief measures.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: 'Long-Term Goals:\n', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: '• Patient will maintain consistent pain control (score ≤ 3/10).\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Patient will demonstrate improved daily quality of life.', options: { fontSize: 15, color: BODY_TEXT } }
  ], { x: startX + 0.3, y: 4.35, w: colW - 0.6, h: 2.4, fontFace: 'Calibri' });

  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 3.75, w: colW, h: 3.2,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EVALUATION (AFTER REASSESSMENT)', {
    x: startX + colW + gapX + 0.3, y: 3.9, w: colW - 0.6, h: 0.35,
    fontSize: 18, bold: true, color: SUCCESS_GREEN, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Clinical Evaluation Results:\n', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: '✔ Patient reported reduced abdominal pain within 24–48 hours.\n\n', options: { fontSize: 15, color: SUCCESS_GREEN } },
    { text: '✔ Patient verbalized effective pain relief measures.\n\n', options: { fontSize: 15, color: SUCCESS_GREEN } },
    { text: '✔ Patient maintained consistent pain control.\n\n', options: { fontSize: 15, color: SUCCESS_GREEN } },
    { text: '✔ Patient demonstrated improved comfort and quality of life.', options: { fontSize: 15, color: SUCCESS_GREEN } }
  ], { x: startX + colW + gapX + 0.3, y: 4.35, w: colW - 0.6, h: 2.4, fontFace: 'Calibri' });
}

// ==========================================
// SLIDE 28: CARE PLAN 1: ABDOMINAL PAIN (INTERVENTIONS)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 1: ABDOMINAL PAIN — ACTIONS', 28);

  const painTable = [
    [
      { text: 'Nursing Interventions', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 17 } },
      { text: 'Scientific Rationales', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 17 } }
    ],
    [
      { text: '1. Assess patient\'s pain level, location, and characteristics.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To determine the severity and pattern of pain and evaluate relief.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '2. Monitor vital signs regularly (blood pressure, heart rate).', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To detect hemodynamic changes and acute surgical complications.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '3. Administer prescribed analgesics as ordered.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To relieve pain and improve patient comfort and participation.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '4. Provide comfort measures (positioning, heat or cold).', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To reduce abdominal wall muscle tension and alleviate discomfort.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '5. Encourage eating small, frequent meals.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To prevent stomach overloading and minimize postprandial cramping.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '6. Choose low-fiber foods during acute flare-ups.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To minimize bowel irritation, gas formation, and distension.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '7. Report any sudden increase in pain or peritoneal symptoms.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To allow prompt intervention and prevent bowel perforation/obstruction.', options: { fontSize: 15, color: BODY_TEXT } }
    ]
  ];

  slide.addTable(painTable, {
    x: 0.6, y: 1.75, w: 12.13,
    colW: [6.06, 6.07],
    border: { type: 'solid', color: 'D8BFC4', pt: 1 },
    rowH: [0.55, 0.58, 0.58, 0.58, 0.58, 0.58, 0.58, 0.58],
    autoPage: false
  });
}

// ==========================================
// SLIDE 29: CARE PLAN 2: FATIGUE (DIAGNOSIS & GOALS)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 2: FATIGUE — ASSESSMENT & GOALS', 29);

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 1.75, w: 12.13, h: 1.8,
    fill: { color: LIGHT_PINK }, line: { color: BURGUNDY, width: 2 }, rectRadius: 0.08
  });
  slide.addText('ACTUAL NURSING DIAGNOSIS: FATIGUE', {
    x: 0.9, y: 1.9, w: 11.53, h: 0.35,
    fontSize: 20, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Fatigue related to ', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: 'decreased oxygen-carrying capacity secondary to anemia, as evidenced by:\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Subjective Data: ', options: { fontSize: 16, bold: true, color: BURGUNDY } },
    { text: '"I feel tired most of the time." • "I feel weak and exhausted." • Dizziness • Trouble with daily tasks.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Objective Data: ', options: { fontSize: 16, bold: true, color: BURGUNDY } },
    { text: 'Pallor of skin and conjunctiva, low Hb & Hct, increased resting heart rate, decreased activity tolerance.', options: { fontSize: 16, color: BODY_TEXT } }
  ], { x: 0.9, y: 2.3, w: 11.53, h: 1.15, fontFace: 'Calibri' });

  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 3.75, w: colW, h: 3.2,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EXPECTED OUTCOMES (GOALS)', {
    x: startX + 0.3, y: 3.9, w: colW - 0.6, h: 0.35,
    fontSize: 18, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Short-Term Goals:\n', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: '• Patient will report decreased fatigue within 24–48 hours.\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Patient will verbalize importance of adequate rest.\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Patient will demonstrate energy conservation techniques.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: 'Long-Term Goals:\n', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: '• Patient will perform daily activities with minimal fatigue.\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Patient\'s Hb level will improve according to treatment plan.', options: { fontSize: 15, color: BODY_TEXT } }
  ], { x: startX + 0.3, y: 4.35, w: colW - 0.6, h: 2.4, fontFace: 'Calibri' });

  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 3.75, w: colW, h: 3.2,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EVALUATION (AFTER REASSESSMENT)', {
    x: startX + colW + gapX + 0.3, y: 3.9, w: colW - 0.6, h: 0.35,
    fontSize: 18, bold: true, color: SUCCESS_GREEN, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Clinical Evaluation Results:\n', options: { fontSize: 16, bold: true, color: DARK_TEXT } },
    { text: '✔ Patient reported decreased fatigue within 24–48 hours.\n\n', options: { fontSize: 15, color: SUCCESS_GREEN } },
    { text: '✔ Patient verbalized understanding of rest & energy conservation.\n\n', options: { fontSize: 15, color: SUCCESS_GREEN } },
    { text: '✔ Patient performed daily activities with improved tolerance.\n\n', options: { fontSize: 15, color: SUCCESS_GREEN } },
    { text: '✔ Patient\'s Hb level improved according to treatment plan.', options: { fontSize: 15, color: SUCCESS_GREEN } }
  ], { x: startX + colW + gapX + 0.3, y: 4.35, w: colW - 0.6, h: 2.4, fontFace: 'Calibri' });
}

// ==========================================
// SLIDE 30: CARE PLAN 2: FATIGUE (INTERVENTIONS)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 2: FATIGUE — ACTIONS', 30);

  const fatigueTable = [
    [
      { text: 'Nursing Interventions', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 17 } },
      { text: 'Scientific Rationales', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 17 } }
    ],
    [
      { text: '1. Assess the patient\'s level of fatigue and weakness.', options: { fontSize: 14.5, color: DARK_TEXT } },
      { text: 'To determine the baseline severity and impact on function.', options: { fontSize: 14.5, color: BODY_TEXT } }
    ],
    [
      { text: '2. Monitor vital signs regularly (pulse, blood pressure).', options: { fontSize: 14.5, color: DARK_TEXT } },
      { text: 'To detect tachycardia, hypotension, and anemia-induced strain.', options: { fontSize: 14.5, color: BODY_TEXT } }
    ],
    [
      { text: '3. Monitor CBC, Hemoglobin (Hb), and Hematocrit (Hct) levels.', options: { fontSize: 14.5, color: DARK_TEXT } },
      { text: 'To evaluate severity of anemia and guide corrective therapies.', options: { fontSize: 14.5, color: BODY_TEXT } }
    ],
    [
      { text: '4. Provide adequate, scheduled rest periods between activities.', options: { fontSize: 14.5, color: DARK_TEXT } },
      { text: 'To conserve metabolic energy and prevent physical exhaustion.', options: { fontSize: 14.5, color: BODY_TEXT } }
    ],
    [
      { text: '5. Assist the patient with daily activities when needed.', options: { fontSize: 14.5, color: DARK_TEXT } },
      { text: 'To prevent excessive fatigue and protect patient from falls.', options: { fontSize: 14.5, color: BODY_TEXT } }
    ],
    [
      { text: '6. Encourage iron-rich foods (lean meat, spinach, lentils).', options: { fontSize: 14.5, color: DARK_TEXT } },
      { text: 'To support hemoglobin synthesis and replenish iron stores.', options: { fontSize: 14.5, color: BODY_TEXT } }
    ],
    [
      { text: '7. Administer prescribed iron supplements / transfusions.', options: { fontSize: 14.5, color: DARK_TEXT } },
      { text: 'To correct iron deficiency and restore oxygen transport.', options: { fontSize: 14.5, color: BODY_TEXT } }
    ],
    [
      { text: '8. Encourage gradual physical activity and light walking as tolerated.', options: { fontSize: 14.5, color: DARK_TEXT } },
      { text: 'To prevent deconditioning and improve functional endurance.', options: { fontSize: 14.5, color: BODY_TEXT } }
    ]
  ];

  slide.addTable(fatigueTable, {
    x: 0.6, y: 1.75, w: 12.13,
    colW: [6.06, 6.07],
    border: { type: 'solid', color: 'D8BFC4', pt: 1 },
    rowH: [0.55, 0.54, 0.54, 0.54, 0.54, 0.54, 0.54, 0.54, 0.54],
    autoPage: false
  });
}

// ==========================================
// SLIDE 31: CARE PLAN 3: FLUID VOLUME (GOALS & EVAL)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 3: DEFICIENT FLUID VOLUME — GOALS', 31);

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 1.75, w: 12.13, h: 1.7,
    fill: { color: 'F4F6F9' }, line: { color: NAVY_BLUE, width: 2 }, rectRadius: 0.08
  });
  slide.addText('POTENTIAL NURSING DIAGNOSIS: DEFICIENT FLUID VOLUME', {
    x: 0.9, y: 1.9, w: 11.53, h: 0.35,
    fontSize: 20, bold: true, color: NAVY_BLUE, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Risk for Deficient Fluid Volume ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'related to excessive fluid losses through vomiting, diarrhea, high stoma output, and inadequate oral fluid intake secondary to colon cancer pathology and treatments.', options: { fontSize: 18, color: BODY_TEXT } }
  ], { x: 0.9, y: 2.35, w: 11.53, h: 0.95, fontFace: 'Calibri' });

  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 3.65, w: colW, h: 3.3,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EXPECTED OUTCOMES (GOALS)', {
    x: startX + 0.3, y: 3.8, w: colW - 0.6, h: 0.35,
    fontSize: 18, bold: true, color: NAVY_BLUE, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '• Patient will maintain adequate hydration and stable fluid balance within 24–48 hours.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Patient will verbalize understanding of drinking adequate fluids.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Patient will remain free from signs and symptoms of dehydration and electrolyte imbalance.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Patient will maintain moist mucous membranes and balanced intake and output (I&O).', options: { fontSize: 15, color: BODY_TEXT } }
  ], { x: startX + 0.3, y: 4.25, w: colW - 0.6, h: 2.5, fontFace: 'Calibri' });

  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 3.65, w: colW, h: 3.3,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EVALUATION (AFTER REASSESSMENT)', {
    x: startX + colW + gapX + 0.3, y: 3.8, w: colW - 0.6, h: 0.35,
    fontSize: 18, bold: true, color: SUCCESS_GREEN, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '✔ Patient is maintaining adequate hydration and stable fluid balance.\n\n', options: { fontSize: 15, color: SUCCESS_GREEN } },
    { text: '✔ Patient exhibits stable vital signs (BP & HR within normal limits).\n\n', options: { fontSize: 15, color: SUCCESS_GREEN } },
    { text: '✔ Strict intake and output (I&O) remain balanced without deficits.\n\n', options: { fontSize: 15, color: SUCCESS_GREEN } },
    { text: '✔ Patient and family verbalize the importance of hydration.', options: { fontSize: 15, color: SUCCESS_GREEN } }
  ], { x: startX + colW + gapX + 0.3, y: 4.25, w: colW - 0.6, h: 2.5, fontFace: 'Calibri' });
}

// ==========================================
// SLIDE 32: CARE PLAN 3: FLUID VOLUME (INTERVENTIONS)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 3: DEFICIENT FLUID VOLUME — ACTIONS', 32);

  const fluidTable = [
    [
      { text: 'Nursing Interventions', options: { bold: true, color: WHITE, fill: { color: NAVY_BLUE }, fontSize: 17 } },
      { text: 'Scientific Rationales', options: { bold: true, color: WHITE, fill: { color: NAVY_BLUE }, fontSize: 17 } }
    ],
    [
      { text: '1. Monitor vital signs (especially BP and HR) regularly.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To detect early signs of hypovolemia and hemodynamic instability.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '2. Measure and record strict intake and output (I&O) every shift.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To accurately evaluate fluid balance, loss volume, and renal perfusion.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '3. Assess skin turgor, mucous membranes, and capillary refill.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To identify physical clinical indicators of dehydration.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '4. Monitor laboratory reports (electrolytes, Hb, Hct).', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To monitor for electrolyte shifts and hemoconcentration.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '5. Encourage frequent oral fluid intake as tolerated.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To maintain hydration status safely via the natural enteral route.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '6. Administer prescribed IV fluids and electrolyte replacements.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To restore intravascular volume and correct chemical imbalances.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '7. Educate patient/family on early symptoms of dehydration.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To empower prompt self-recognition (extreme thirst, dark urine).', options: { fontSize: 15, color: BODY_TEXT } }
    ]
  ];

  slide.addTable(fluidTable, {
    x: 0.6, y: 1.75, w: 12.13,
    colW: [6.06, 6.07],
    border: { type: 'solid', color: 'B0BEC5', pt: 1 },
    rowH: [0.55, 0.58, 0.58, 0.58, 0.58, 0.58, 0.58, 0.58],
    autoPage: false
  });
}

// ==========================================
// SLIDE 33: CARE PLAN 4: GI MOTILITY (GOALS & EVAL)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 4: DYSFUNCTIONAL GI MOTILITY — GOALS', 33);

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 1.75, w: 12.13, h: 1.7,
    fill: { color: 'F4F6F9' }, line: { color: NAVY_BLUE, width: 2 }, rectRadius: 0.08
  });
  slide.addText('POTENTIAL NURSING DIAGNOSIS: DYSFUNCTIONAL GI MOTILITY', {
    x: 0.9, y: 1.9, w: 11.53, h: 0.35,
    fontSize: 20, bold: true, color: NAVY_BLUE, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Risk for Dysfunctional Gastrointestinal Motility ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'related to the presence of a colon tumor mass, potential mechanical narrowing of the intestinal lumen, surgical bowel manipulation, and opioid analgesia.', options: { fontSize: 18, color: BODY_TEXT } }
  ], { x: 0.9, y: 2.35, w: 11.53, h: 0.95, fontFace: 'Calibri' });

  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 3.65, w: colW, h: 3.3,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EXPECTED OUTCOMES (GOALS)', {
    x: startX + 0.3, y: 3.8, w: colW - 0.6, h: 0.35,
    fontSize: 18, bold: true, color: NAVY_BLUE, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '• Patient will maintain stable gastrointestinal motility and bowel function.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Patient will exhibit active bowel sounds in all 4 abdominal quadrants.\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Patient will remain free from signs of bowel obstruction (distension, vomiting, obstipation).\n\n', options: { fontSize: 15, color: BODY_TEXT } },
    { text: '• Patient and family will verbalize warning signs requiring immediate reporting.', options: { fontSize: 15, color: BODY_TEXT } }
  ], { x: startX + 0.3, y: 4.25, w: colW - 0.6, h: 2.5, fontFace: 'Calibri' });

  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 3.65, w: colW, h: 3.3,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EVALUATION (AFTER REASSESSMENT)', {
    x: startX + colW + gapX + 0.3, y: 3.8, w: colW - 0.6, h: 0.35,
    fontSize: 18, bold: true, color: SUCCESS_GREEN, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '✔ Patient maintained stable gastrointestinal function.\n\n', options: { fontSize: 15, color: SUCCESS_GREEN } },
    { text: '✔ Active bowel sounds confirmed auscultated across all four quadrants.\n\n', options: { fontSize: 15, color: SUCCESS_GREEN } },
    { text: '✔ Patient remained completely free from obstruction symptoms.\n\n', options: { fontSize: 15, color: SUCCESS_GREEN } },
    { text: '✔ Patient and family verbalized obstruction warning signs.', options: { fontSize: 15, color: SUCCESS_GREEN } }
  ], { x: startX + colW + gapX + 0.3, y: 4.25, w: colW - 0.6, h: 2.5, fontFace: 'Calibri' });
}

// ==========================================
// SLIDE 34: CARE PLAN 4: GI MOTILITY (INTERVENTIONS)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 4: DYSFUNCTIONAL GI MOTILITY — ACTIONS', 34);

  const motilityTable = [
    [
      { text: 'Nursing Interventions', options: { bold: true, color: WHITE, fill: { color: NAVY_BLUE }, fontSize: 17 } },
      { text: 'Scientific Rationales', options: { bold: true, color: WHITE, fill: { color: NAVY_BLUE }, fontSize: 17 } }
    ],
    [
      { text: '1. Auscultate bowel sounds in all 4 quadrants every 4 hours.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To detect early signs of hypomotility, ileus, or mechanical obstruction.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '2. Assess abdomen for distension, rigidity, and tenderness every shift.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To monitor gas/fluid accumulation behind a narrowed colon segment.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '3. Monitor and document frequency and consistency of stool and flatus.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To verify colon patency and the return of normal peristalsis.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '4. Encourage progressive ambulation and physical activity as tolerated.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To stimulate peristalsis and enhance natural gastrointestinal motility.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '5. Maintain adequate hydration via oral fluids or prescribed IV routes.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To prevent constipation and ensure soft stool passage through lumen.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '6. Educate patient to consume small, frequent meals rather than large.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To prevent overloading and obstructing the narrowed colonic lumen.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '7. Teach patient/family to immediately report vomiting or failure to pass gas.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To ensure emergency intervention before bowel perforation occurs.', options: { fontSize: 15, color: BODY_TEXT } }
    ]
  ];

  slide.addTable(motilityTable, {
    x: 0.6, y: 1.75, w: 12.13,
    colW: [6.06, 6.07],
    border: { type: 'solid', color: 'B0BEC5', pt: 1 },
    rowH: [0.55, 0.58, 0.58, 0.58, 0.58, 0.58, 0.58, 0.58],
    autoPage: false
  });
}

// ==========================================
// SLIDE 35: SUMMARY
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'SUMMARY & CONCLUSION', 35);

  slide.addShape(pptx.ShapeType.rect, {
    x: 1.0, y: 2.0, w: 11.33, h: 4.6,
    fill: { color: LIGHT_PINK }, line: { color: BURGUNDY, width: 2 }, rectRadius: 0.12
  });

  slide.addText('KEY CLINICAL TAKEAWAYS', {
    x: 1.5, y: 2.35, w: 10.33, h: 0.45,
    fontSize: 26, bold: true, color: BURGUNDY, align: 'center', fontFace: 'Calibri'
  });

  slide.addText([
    { text: 'To sum up, colon cancer is a preventable disease through early screening.\n\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: 'As nurses, our role is crucial in supporting the patient before and after surgery.\n\n', options: { fontSize: 22, color: DARK_TEXT } },
    { text: 'By managing actual symptoms like pain and fatigue, and preventing potential complications like bowel obstruction, we can safely improve the patient\'s outcomes and quality of life.\n\n', options: { fontSize: 22, color: DARK_TEXT } },
    { text: 'Thank you.', options: { fontSize: 32, bold: true, color: BURGUNDY } }
  ], {
    x: 1.5, y: 3.0, w: 10.33, h: 3.3,
    fontFace: 'Calibri', align: 'center', valign: 'middle', lineSpacingMultiple: 1.25
  });
}

// ==========================================
// SLIDE 36: REFERENCES
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'REFERENCES & AUTHORITATIVE SOURCES', 36);

  const refs = [
    'American Cancer Society. (n.d.-a). Colorectal cancer prevention guidelines. https://www.cancer.org/health-care-professionals/american-cancer-society-prevention-early-detection-guidelines/colorectal-cancer-screening-guidelines.html',
    'American Cancer Society. (n.d.-b). What is colorectal cancer? https://www.cancer.org/cancer/types/colon-rectal-cancer/about/what-is-colorectal-cancer.html',
    'Colon Cancer Coalition. (n.d.). Signs and symptoms of colorectal cancer. https://coloncancercoalition.org',
    'Mayo Clinic. (n.d.). Colorectal cancer: Diagnosis and treatment. https://www.mayoclinic.org',
    'National Cancer Institute. (n.d.-a). Colon cancer treatment (PDQ®)–Patient version. https://www.cancer.gov/types/colorectal/patient/colon-treatment-pdq',
    'National Cancer Institute. (n.d.-b). Screening tests to detect colorectal cancer and polyps. https://www.cancer.gov/types/colorectal/screening-fact-sheet',
    'Roshandel, G., Ghasemi-Kebria, F., & Malekzadeh, R. (2024). Colorectal cancer: Epidemiology, risk factors, and prevention. Cancers, 16(8), 1530. https://doi.org/10.3390/cancers16081530',
    'Surveillance, Epidemiology, and End Results (SEER) Program. (2025). Anatomy of the colon and rectum. National Cancer Institute. https://training.seer.cancer.gov/colorectal/anatomy/'
  ];

  const refRuns = refs.map((ref, idx) => ({
    text: `${idx + 1}. ${ref}\n\n`,
    options: { fontSize: 13, color: BODY_TEXT, paraSpaceAfter: 4 }
  }));

  slide.addText(refRuns, {
    x: 0.6, y: 1.75, w: 12.13, h: 5.2,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });
}

// WRITE FILE - write to both standard name and updated name in case original is open in PowerPoint
const outputDir = path.join(__dirname, '..', 'public');
const primaryFile = path.join(outputDir, 'colon-cancer-presentation.pptx');
const updatedFile = path.join(outputDir, 'colon-cancer-presentation-final.pptx');

async function savePresentation() {
  // First write to final/updated file which is never locked
  try {
    await pptx.writeFile({ fileName: updatedFile });
    const stats2 = fs.statSync(updatedFile);
    console.log('SUCCESS! Updated presentation generated:');
    console.log('  Path:', updatedFile);
    console.log('  Size:', (stats2.size / 1024 / 1024).toFixed(2), 'MB');
    console.log('  Total Slides:', TOTAL_SLIDES);
  } catch (err) {
    console.error('Error saving updatedFile:', err);
  }

  // Then try to overwrite primary file if not currently open in PowerPoint
  try {
    await pptx.writeFile({ fileName: primaryFile });
    console.log('Overwrote primary file successfully:', primaryFile);
  } catch (err) {
    console.log('Note: primary file is currently open in PowerPoint by user, saved as final file.');
  }
}

savePresentation();
