const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

// Colors
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
const ACCENT_RED = 'A62B3E';
const SUCCESS_GREEN = '2E7D32';

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

// Extract Logos
const nursingLogo = getBase64Image(path.join(PROJECT_IMAGES, 'nursing-logo-opt.png')) ||
                    getBase64Image(path.join(DOCX_MEDIA, 'image1.jpeg'));
const deptLogo = getBase64Image(path.join(PROJECT_IMAGES, 'dept-logo-opt.png')) ||
                 getBase64Image(path.join(DOCX_MEDIA, 'image2.jpeg'));
const azharLogo = getBase64Image(path.join(PROJECT_IMAGES, 'azhar-logo-opt.png')) ||
                  getBase64Image(path.join(DOCX_MEDIA, 'image3.jpeg'));

// Medical Images
const coverIllustration = getBase64Image(path.join(DOCX_MEDIA, 'image2.png')); // Colon Cancer with tumor closeup
const colonAnatomyDiagram = getBase64Image(path.join(DOCX_MEDIA, 'image4.jpeg')); // Anatomical parts of colon
const colonoscopyImg = getBase64Image(path.join(PPTX_MEDIA, 'image8.jpeg')); // Endoscopy view of adenomatous polyp
const histopathologyImg = getBase64Image(path.join(PPTX_MEDIA, 'image9.jpeg')); // Adenocarcinoma H&E
const pathoProgressionImg = getBase64Image(path.join(PPTX_MEDIA, 'image11.jpeg')); // Adenoma-carcinoma sequence
const ctScanImg = getBase64Image(path.join(PPTX_MEDIA, 'image12.jpeg')); // Abdominal CT scan
const stagingDiagramImg = getBase64Image(path.join(PPTX_MEDIA, 'image13.jpeg')); // Staging 0-IV
const surgeryAtlasImg = getBase64Image(path.join(PPTX_MEDIA, 'image14.jpeg')); // Partial colectomy
const nurseVitalsImg = getBase64Image(path.join(PPTX_MEDIA, 'image15.jpeg')); // Nurse vital signs monitoring
const nurseCareImg = getBase64Image(path.join(PPTX_MEDIA, 'image16.jpeg')); // Bedside nursing care

const TOTAL_SLIDES = 35;

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5 inches
pptx.author = 'Faculty of Nursing - Al-Azhar University';
pptx.subject = 'Colon Cancer Comprehensive Nursing Seminar';
pptx.title = 'Colon Cancer - Clinical Presentation & Nursing Care Plans';

// Helper to add institutional header
function addHeader(slide, title, slideNum) {
  // Top thin accent bar
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: BURGUNDY } });

  // Institutional Header Strip
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.08, w: '100%', h: 0.90, fill: { color: LIGHT_PINK } });
  slide.addShape(pptx.ShapeType.line, { x: 0, y: 0.98, w: '100%', h: 0, line: { color: 'E5D4D7', width: 1 } });

  // 3 Logos in header
  if (nursingLogo) {
    slide.addImage({ data: nursingLogo, x: 0.35, y: 0.13, w: 0.8, h: 0.8 });
  }
  if (deptLogo) {
    slide.addImage({ data: deptLogo, x: 1.25, y: 0.13, w: 0.8, h: 0.8 });
  }
  if (azharLogo) {
    slide.addImage({ data: azharLogo, x: 12.18, y: 0.13, w: 0.8, h: 0.8 });
  }

  // Institutional Bilingual Typography
  slide.addText([
    { text: 'Faculty of Nursing for Girls — Al-Azhar University\n', options: { fontSize: 10, bold: true, color: BURGUNDY } },
    { text: 'Medical-Surgical Nursing Department  •  قسم التمريض الباطني والجراحي\n', options: { fontSize: 9.5, color: SUBTITLE_GRAY } },
    { text: 'كُليّة التمريض للبنات بالقاهرة  •  جامعة الأزهر الشريف', options: { fontSize: 9.5, bold: true, color: BURGUNDY } }
  ], { x: 2.15, y: 0.12, w: 9.9, h: 0.82, align: 'center', valign: 'middle', fontFace: 'Calibri' });

  // Slide Title Banner
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 1.02, w: '100%', h: 0.68, fill: { color: BURGUNDY } });
  slide.addText(title, {
    x: 0.5, y: 1.02, w: 10.8, h: 0.68,
    fontSize: 32, bold: true, color: WHITE,
    align: 'left', valign: 'middle', fontFace: 'Calibri'
  });

  // Slide Counter
  slide.addText(`${slideNum} / ${TOTAL_SLIDES}`, {
    x: 11.4, y: 1.02, w: 1.4, h: 0.68,
    fontSize: 16, bold: true, color: 'E2B6BF',
    align: 'right', valign: 'middle', fontFace: 'Calibri'
  });

  // Slide Footer
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 7.25, w: '100%', h: 0.25, fill: { color: BURGUNDY } });
  slide.addText('Faculty of Nursing for Girls • Al-Azhar University — Medical-Surgical Nursing Department (2026/2027)', {
    x: 0.5, y: 7.25, w: 12.33, h: 0.25,
    fontSize: 9, color: 'F0D5DA', align: 'center', valign: 'middle', fontFace: 'Calibri'
  });
}

console.log('Building presentation slides...');

// ==========================================
// SLIDE 1: COVER SLIDE
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: DARK_BG };

  // Top and bottom gold accent bars
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: GOLD } });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 7.42, w: '100%', h: 0.08, fill: { color: GOLD } });

  // Logos at top: Nursing (left), Dept (center-left), Azhar (right)
  if (nursingLogo) slide.addImage({ data: nursingLogo, x: 0.6, y: 0.25, w: 1.1, h: 1.1 });
  if (deptLogo) slide.addImage({ data: deptLogo, x: 1.85, y: 0.25, w: 1.1, h: 1.1 });
  if (azharLogo) slide.addImage({ data: azharLogo, x: 11.63, y: 0.25, w: 1.1, h: 1.1 });

  // Institutional titles
  slide.addText([
    { text: 'Faculty of Nursing for Girls — Al-Azhar University\n', options: { fontSize: 14, bold: true, color: GOLD } },
    { text: 'كُليّة التمريض للبنات بالقاهرة  •  جامعة الأزهر الشريف\n', options: { fontSize: 13, bold: true, color: CREAM } },
    { text: 'Department of Medical-Surgical Nursing  •  قسم التمريض الباطني والجراحي', options: { fontSize: 12, color: 'CCCCCC' } }
  ], { x: 3.1, y: 0.25, w: 8.4, h: 1.1, align: 'center', valign: 'middle', fontFace: 'Calibri' });

  slide.addShape(pptx.ShapeType.line, { x: 0.6, y: 1.45, w: 12.13, h: 0, line: { color: GOLD, width: 1.5 } });

  // Left side: Title and Academic Information
  slide.addText('COLON CANCER', {
    x: 0.6, y: 1.6, w: 7.2, h: 1.0,
    fontSize: 50, bold: true, color: WHITE,
    align: 'left', valign: 'middle', fontFace: 'Calibri', charSpacing: 3
  });

  slide.addText('A Comprehensive Clinical & Evidence-Based Nursing Seminar', {
    x: 0.6, y: 2.65, w: 7.2, h: 0.5,
    fontSize: 18, italic: true, color: GOLD, fontFace: 'Calibri'
  });

  // Supervision Box
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 3.25, w: 7.2, h: 1.7,
    fill: { color: '232338' }, line: { color: GOLD, width: 1 }, rectRadius: 0.08
  });

  slide.addText('UNDER SUPERVISION', {
    x: 0.8, y: 3.35, w: 6.8, h: 0.35,
    fontSize: 13, bold: true, color: GOLD, fontFace: 'Calibri', charSpacing: 2
  });

  slide.addText([
    { text: 'Prof. Dr. Heba Abdel-Azem\n', options: { fontSize: 15, bold: true, color: WHITE } },
    { text: 'Professor of Medical-Surgical Nursing Department\n\n', options: { fontSize: 11.5, color: 'BBBBBB' } },
    { text: 'Dr. Eman Abel-Maqsoud\n', options: { fontSize: 15, bold: true, color: WHITE } },
    { text: 'Demonstrator of Medical-Surgical Nursing Department', options: { fontSize: 11.5, color: 'BBBBBB' } }
  ], { x: 0.8, y: 3.75, w: 6.8, h: 1.15, fontFace: 'Calibri' });

  // Student Authors Box
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 5.1, w: 7.2, h: 1.8,
    fill: { color: '232338' }, line: { color: '3F3F5F', width: 1 }, rectRadius: 0.08
  });

  slide.addText('PREPARED BY (SECOND YEAR — 2026 / 2027):', {
    x: 0.8, y: 5.2, w: 6.8, h: 0.35,
    fontSize: 12, bold: true, color: GOLD, fontFace: 'Calibri'
  });

  slide.addText([
    { text: '1. Hager Saeed        2. Hager Hany        3. Hager Ahmed\n', options: { fontSize: 13, bold: true, color: WHITE } },
    { text: '4. Hager Mostafa      5. Haidy Mohammed    6. Norhan Badwy', options: { fontSize: 13, bold: true, color: WHITE } }
  ], { x: 0.8, y: 5.6, w: 6.8, h: 0.7, fontFace: 'Calibri' });

  slide.addText('Faculty of Nursing • Al-Azhar University', {
    x: 0.8, y: 6.45, w: 6.8, h: 0.35,
    fontSize: 11.5, color: 'AAAAAA', italic: true, fontFace: 'Calibri'
  });

  // Right side: Medical Illustration from Word doc
  if (coverIllustration) {
    slide.addImage({ data: coverIllustration, x: 8.1, y: 1.7, w: 4.6, h: 5.2, rounding: true });
    slide.addText('Colonic Oncology & Tumor Pathology', {
      x: 8.1, y: 6.95, w: 4.6, h: 0.35,
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
    x: 0.6, y: 1.85, w: 12.13, h: 0.45,
    fontSize: 22, bold: true, color: BURGUNDY, fontFace: 'Calibri'
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
  const cardH = 1.8;
  const startX = 0.6;
  const startY = 2.45;
  const gapX = 0.36;
  const gapY = 0.45;

  students.forEach((s, idx) => {
    const col = idx % 3;
    const row = Math.floor(idx / 3);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);

    // Card Box
    slide.addShape(pptx.ShapeType.rect, {
      x, y, w: cardW, h: cardH,
      fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1.5 }, rectRadius: 0.1
    });

    // Badge
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.25, y: y + 0.35, w: 0.9, h: 0.9,
      fill: { color: BURGUNDY }
    });
    slide.addText(s.num, {
      x: x + 0.25, y: y + 0.35, w: 0.9, h: 0.9,
      fontSize: 22, bold: true, color: WHITE, align: 'center', valign: 'middle', fontFace: 'Calibri'
    });

    // Name & Department
    slide.addText(s.name, {
      x: x + 1.3, y: y + 0.35, w: cardW - 1.45, h: 0.5,
      fontSize: 24, bold: true, color: DARK_TEXT, fontFace: 'Calibri'
    });
    slide.addText(s.role, {
      x: x + 1.3, y: y + 0.85, w: cardW - 1.45, h: 0.65,
      fontSize: 14, color: SUBTITLE_GRAY, fontFace: 'Calibri'
    });
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 6.45, w: 12.13, h: 0.6,
    fill: { color: 'F5F5F5' }, line: { color: 'DDDDDD', width: 1 }, rectRadius: 0.08
  });
  slide.addText('Department of Medical-Surgical Nursing • Faculty of Nursing for Girls, Cairo • Al-Azhar University', {
    x: 0.6, y: 6.45, w: 12.13, h: 0.6,
    fontSize: 14, color: BURGUNDY, bold: true, align: 'center', valign: 'middle', fontFace: 'Calibri'
  });
}

// ==========================================
// SLIDE 3: GOAL
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'GOAL OF THE ASSIGNMENT', 3);

  slide.addShape(pptx.ShapeType.rect, {
    x: 1.0, y: 2.1, w: 11.33, h: 4.6,
    fill: { color: LIGHT_PINK }, line: { color: BURGUNDY, width: 2 }, rectRadius: 0.15
  });

  slide.addText('OVERARCHING EDUCATIONAL GOAL', {
    x: 1.5, y: 2.4, w: 10.33, h: 0.6,
    fontSize: 26, bold: true, color: BURGUNDY, align: 'center', fontFace: 'Calibri'
  });

  slide.addText(
    'The goal of this assignment is to provide nursing students with a comprehensive understanding of colon cancer, including its anatomy, pathophysiology, risk factors, clinical manifestations, diagnosis, and management, in order to enhance their knowledge and ability to provide safe, effective, and evidence-based nursing care for patients with colon cancer.',
    {
      x: 1.5, y: 3.1, w: 10.33, h: 3.2,
      fontSize: 28, color: DARK_TEXT, align: 'justify', valign: 'top',
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
    x: 0.6, y: 1.8, w: 12.13, h: 0.45,
    fontSize: 24, bold: true, color: BURGUNDY, fontFace: 'Calibri'
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

  // Left Column (1-4)
  objectives.slice(0, 4).forEach((obj, i) => {
    const y = 2.4 + i * 1.15;
    slide.addShape(pptx.ShapeType.rect, {
      x: startX, y, w: colW, h: 1.0,
      fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.08
    });
    slide.addText(`${i + 1}`, {
      x: startX + 0.15, y: y + 0.2, w: 0.6, h: 0.6,
      fontSize: 20, bold: true, color: WHITE, fill: { color: BURGUNDY }, align: 'center', valign: 'middle'
    });
    slide.addText(obj, {
      x: startX + 0.9, y: y + 0.1, w: colW - 1.05, h: 0.8,
      fontSize: 18, color: DARK_TEXT, fontFace: 'Calibri', valign: 'middle'
    });
  });

  // Right Column (5-8)
  objectives.slice(4, 8).forEach((obj, i) => {
    const y = 2.4 + i * 1.15;
    slide.addShape(pptx.ShapeType.rect, {
      x: startX + colW + gapX, y, w: colW, h: 1.0,
      fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.08
    });
    slide.addText(`${i + 5}`, {
      x: startX + colW + gapX + 0.15, y: y + 0.2, w: 0.6, h: 0.6,
      fontSize: 20, bold: true, color: WHITE, fill: { color: BURGUNDY }, align: 'center', valign: 'middle'
    });
    slide.addText(obj, {
      x: startX + colW + gapX + 0.9, y: y + 0.1, w: colW - 1.05, h: 0.8,
      fontSize: 18, color: DARK_TEXT, fontFace: 'Calibri', valign: 'middle'
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
    '04. Pathophysiology',
    '05. Risk Factors',
    '06. Causes (Etiology)',
    '07. Types of Colon Cancer',
    '08. Signs and Symptoms',
    '09. Complications'
  ];

  const rightItems = [
    '10. Diagnostic Tests',
    '11. Prognosis & Survival Rates',
    '12. Management and Treatment',
    '13. Prevention & Screening',
    '14. Nursing Role (Perioperative & Supportive)',
    '15. Nursing Process (4 Care Plans)',
    '16. Summary',
    '17. References'
  ];

  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 1.85, w: colW, h: 5.15,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.1
  });

  const leftRuns = leftItems.map(item => ({
    text: item + '\n',
    options: { fontSize: 20, bold: true, color: BURGUNDY, paraSpaceAfter: 12 }
  }));
  slide.addText(leftRuns, {
    x: startX + 0.4, y: 2.05, w: colW - 0.8, h: 4.75,
    fontFace: 'Calibri', valign: 'top'
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 1.85, w: colW, h: 5.15,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.1
  });

  const rightRuns = rightItems.map(item => ({
    text: item + '\n',
    options: { fontSize: 20, bold: true, color: BURGUNDY, paraSpaceAfter: 16 }
  }));
  slide.addText(rightRuns, {
    x: startX + colW + gapX + 0.4, y: 2.05, w: colW - 0.8, h: 4.75,
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
    { text: '• Major Digestive Malignancy:\n', options: { fontSize: 26, bold: true, color: BURGUNDY } },
    { text: 'Colon cancer is one of the major types of cancer affecting the digestive system. It develops when cells in the lining of the colon undergo abnormal changes that lead to uncontrolled growth.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '• Precancerous Polyps Origin:\n', options: { fontSize: 26, bold: true, color: BURGUNDY } },
    { text: 'In many cases, colon cancer begins as small growths called polyps, some of which may gradually develop into cancer over several years.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '• Critical Role of Early Detection:\n', options: { fontSize: 26, bold: true, color: BURGUNDY } },
    { text: 'Early detection plays an important role in improving treatment outcomes. Screening tests can identify precancerous polyps and detect cancer at an early stage.', options: { fontSize: 22, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.85, w: 7.6, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });

  if (colonoscopyImg) {
    slide.addImage({ data: colonoscopyImg, x: 8.5, y: 2.0, w: 4.2, h: 4.5, rounding: true });
    slide.addText('Endoscopic Colonoscopy View: Precancerous Adenomatous Polyp', {
      x: 8.5, y: 6.55, w: 4.2, h: 0.45,
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
    { text: 'Pathological Definition:\n', options: { fontSize: 28, bold: true, color: BURGUNDY } },
    { text: 'Colon cancer is a malignant disease that develops in the tissues of the colon, which is the longest part of the large intestine.\n\n', options: { fontSize: 24, color: BODY_TEXT } },
    { text: 'Uncontrolled Cellular Growth:\n', options: { fontSize: 28, bold: true, color: BURGUNDY } },
    { text: 'It occurs when abnormal cells grow and divide uncontrollably, forming a malignant tumor.\n\n', options: { fontSize: 24, color: BODY_TEXT } },
    { text: 'Predominant Histological Type:\n', options: { fontSize: 28, bold: true, color: BURGUNDY } },
    { text: 'Most colon cancers are adenocarcinomas, which develop from glandular cells lining the inner surface of the colon (NCI).', options: { fontSize: 24, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.85, w: 7.6, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.2
  });

  if (histopathologyImg) {
    slide.addImage({ data: histopathologyImg, x: 8.5, y: 2.0, w: 4.2, h: 4.5, rounding: true });
    slide.addText('Histopathology (H&E Stain): Invasive Colonic Adenocarcinoma', {
      x: 8.5, y: 6.55, w: 4.2, h: 0.45,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 8: ANATOMY OF THE COLON (4 SECTIONS)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'ANATOMY OF THE COLON — 4 MAIN SECTIONS', 8);

  slide.addText([
    { text: 'Overview & Dimensions:\n', options: { fontSize: 24, bold: true, color: BURGUNDY } },
    { text: 'The colon is the main part of the large intestine (~1.5 m in adults). It absorbs water and electrolytes and forms solid stool.\n\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '1. Ascending Colon: ', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: 'Located on the right side of the abdomen, carries intestinal contents upward from the cecum.\n\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '2. Transverse Colon: ', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: 'Extends across the upper abdomen from right to left.\n\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '3. Descending Colon: ', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: 'Runs downward along the left side of the abdomen.\n\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '4. Sigmoid Colon: ', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: 'An S-shaped section that connects descending colon to rectum.', options: { fontSize: 20, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.85, w: 7.2, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });

  if (colonAnatomyDiagram) {
    slide.addImage({ data: colonAnatomyDiagram, x: 8.1, y: 1.85, w: 4.6, h: 4.8, rounding: true });
    slide.addText('Medical Atlas: The 4 Sections of the Colon & Tumor Inset', {
      x: 8.1, y: 6.7, w: 4.6, h: 0.4,
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

  // Top Section: Core Physiology
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 1.85, w: 12.13, h: 1.5,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.08
  });

  slide.addText([
    { text: 'Core Physiological Functions:\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '1. Water & Electrolyte Absorption: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Reabsorbs 1–1.5 L of water and essential minerals daily.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '2. Stool Formation: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Consolidates digestive residue into solid, lubricated feces.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '3. Storage & Motility: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Stores waste and propels it toward rectum via peristalsis.', options: { fontSize: 18, color: BODY_TEXT } }
  ], { x: 0.8, y: 1.95, w: 11.73, h: 1.3, fontFace: 'Calibri' });

  // Bottom Section: The 4 Layers Table
  slide.addText('The 4 Histological Layers of the Colon Wall (TNM Staging Relevance):', {
    x: 0.6, y: 3.55, w: 12.13, h: 0.45,
    fontSize: 22, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });

  const layerTable = [
    [
      { text: 'Bowel Wall Layer', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 18 } },
      { text: 'Histological Structure', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 18 } },
      { text: 'Oncological & Staging Significance', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 18 } }
    ],
    [
      { text: '1. Mucosa (Innermost)', options: { bold: true, color: BURGUNDY, fontSize: 16 } },
      { text: 'Columnar epithelium & mucus-secreting goblet cells', options: { fontSize: 16, color: BODY_TEXT } },
      { text: 'Site of origin for >95% of colon cancers (Stage 0 / In Situ)', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '2. Submucosa', options: { bold: true, color: BURGUNDY, fontSize: 16 } },
      { text: 'Dense connective tissue rich in blood vessels & lymphatics', options: { fontSize: 16, color: BODY_TEXT } },
      { text: 'Invasion here gives access to lymphatic spread (Stage I / T1)', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '3. Muscularis Propria', options: { bold: true, color: BURGUNDY, fontSize: 16 } },
      { text: 'Inner circular and outer longitudinal smooth muscle', options: { fontSize: 16, color: BODY_TEXT } },
      { text: 'Deep muscle invasion defines Stage I (T2) progression', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '4. Serosa (Outermost)', options: { bold: true, color: BURGUNDY, fontSize: 16 } },
      { text: 'Visceral peritoneum & pericolic adipose tissue', options: { fontSize: 16, color: BODY_TEXT } },
      { text: 'Penetration through the wall defines Stage II (T3–T4)', options: { fontSize: 16, color: BODY_TEXT } }
    ]
  ];

  slide.addTable(layerTable, {
    x: 0.6, y: 4.1, w: 12.13,
    colW: [2.8, 4.5, 4.83],
    border: { type: 'solid', color: 'D8BFC4', pt: 1 },
    rowH: [0.55, 0.6, 0.6, 0.6, 0.6],
    autoPage: false
  });
}

// ==========================================
// SLIDE 10: PATHOPHYSIOLOGY (PART 1: PATHWAYS)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PATHOPHYSIOLOGY — MOLECULAR PATHWAYS', 10);

  slide.addText([
    { text: '1. Adenoma–Carcinoma Sequence (~80%) — CIN Pathway:\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• APC mutation (early, gatekeeper): ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Loss of APC → β-catenin accumulates → Wnt signaling stays on → normal mucosa to small adenoma.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• KRAS mutation: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Continuous growth signaling → intermediate adenoma.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• SMAD4 / TGF-β loss: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Loss of growth inhibition → late adenoma.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• p53 mutation (TP53): ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Loss of apoptosis and cell-cycle control → invasive carcinoma.\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '2. MMR Deficiency / Microsatellite Instability (~15%):\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Defect in MMR genes (MLH1, MSH2, MSH6, PMS2) → Lynch syndrome or sporadic MLH1 hypermethylation.\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '3. Serrated Pathway / CIMP (~10–20%):\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• BRAF mutation → sessile serrated polyp; typically right-sided colon.', options: { fontSize: 18, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.85, w: 7.4, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });

  if (pathoProgressionImg) {
    slide.addImage({ data: pathoProgressionImg, x: 8.3, y: 2.1, w: 4.4, h: 4.3, rounding: true });
    slide.addText('Adenoma-to-Carcinoma Progression Sequence', {
      x: 8.3, y: 6.5, w: 4.4, h: 0.4,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 11: PATHOPHYSIOLOGY (PART 2: PROGRESSION)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PATHOPHYSIOLOGY — PROGRESSION & SPREAD', 11);

  slide.addText([
    { text: '4. Hereditary Syndromes:\n', options: { fontSize: 24, bold: true, color: BURGUNDY } },
    { text: '• FAP (Familial Adenomatous Polyposis): ', options: { fontSize: 20, bold: true, color: DARK_TEXT } },
    { text: 'Germline APC mutation → hundreds to thousands of polyps → nearly 100% cancer risk if untreated.\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '• Lynch Syndrome (HNPCC): ', options: { fontSize: 20, bold: true, color: DARK_TEXT } },
    { text: 'Inherited germline MMR mutation with high risk of colon and extracolonic cancers.\n\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '5. Risk Factors Driving These Cellular Changes:\n', options: { fontSize: 24, bold: true, color: BURGUNDY } },
    { text: '• Age, family history, and chronic IBD inflammation (causing early p53 loss).\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '• High red/processed meat, low fiber, obesity, smoking, alcohol, and sedentary lifestyle.\n\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '6. Progression & Metastatic Dissemination:\n', options: { fontSize: 24, bold: true, color: BURGUNDY } },
    { text: '• Local Invasion: Mucosa → Submucosa → Muscularis Propria → Serosa.\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '• Spread via lymphatics to regional mesenteric nodes.\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '• Spread via portal vein blood: Liver (most common site), followed by Lungs.', options: { fontSize: 20, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.85, w: 12.13, h: 4.4,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });

  // Highlight Box
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 6.15, w: 12.13, h: 0.9,
    fill: { color: 'FFF9E6' }, line: { color: GOLD, width: 2 }, rectRadius: 0.1
  });
  slide.addText('⚡ Key Clinical Takeaway: The malignant transformation takes roughly 10–15 years, which provides a critical window of opportunity for screening colonoscopy and early polyp removal to prevent cancer!', {
    x: 0.8, y: 6.25, w: 11.73, h: 0.7,
    fontSize: 18, bold: true, color: '8B6508', align: 'center', valign: 'middle', fontFace: 'Calibri'
  });
}

// ==========================================
// SLIDE 12: RISK FACTORS
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'RISK FACTORS FOR COLON CANCER', 12);

  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  // Non-modifiable Box
  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 1.85, w: colW, h: 4.3,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1.5 }, rectRadius: 0.1
  });
  slide.addText('NON-MODIFIABLE RISK FACTORS', {
    x: startX + 0.3, y: 2.0, w: colW - 0.6, h: 0.45,
    fontSize: 22, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '• Age: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Risk rises sharply after 50 (~90% of cases).\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Family History: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'First-degree relative with CRC roughly doubles risk.\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Hereditary Syndromes: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Lynch syndrome, FAP, MUTYH, Peutz-Jeghers.\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Personal Medical History: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Previous adenomas, CRC, or IBD (Ulcerative Colitis > Crohn\'s).\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Demographics: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Male sex; African American ethnicity (higher mortality).', options: { fontSize: 18, color: BODY_TEXT } }
  ], { x: startX + 0.3, y: 2.5, w: colW - 0.6, h: 3.5, fontFace: 'Calibri' });

  // Modifiable Box
  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 1.85, w: colW, h: 4.3,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1.5 }, rectRadius: 0.1
  });
  slide.addText('MODIFIABLE LIFESTYLE FACTORS', {
    x: startX + colW + gapX + 0.3, y: 2.0, w: colW - 0.6, h: 0.45,
    fontSize: 22, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '• Unhealthy Diet: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'High in red and processed meat, low in dietary fiber.\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Obesity & Sedentary Lifestyle: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Central abdominal fat and lack of physical exercise.\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Tobacco & Alcohol: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Cigarette smoking and heavy alcohol consumption.\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Type 2 Diabetes & Insulin Resistance.\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Protective Factors: ', options: { fontSize: 18, bold: true, color: SUCCESS_GREEN } },
    { text: 'Screening colonoscopy, high-fiber diet, active lifestyle.', options: { fontSize: 18, color: BODY_TEXT } }
  ], { x: startX + colW + gapX + 0.3, y: 2.5, w: colW - 0.6, h: 3.5, fontFace: 'Calibri' });

  // Key point footer
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 6.25, w: 12.13, h: 0.8,
    fill: { color: 'F5F5F5' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('Key Point: Most risk factors promote chronic inflammation, insulin/IGF-1 signaling, and DNA damage feeding carcinogenesis.', {
    x: 0.8, y: 6.35, w: 11.73, h: 0.6,
    fontSize: 16, bold: true, color: DARK_TEXT, align: 'center', valign: 'middle', fontFace: 'Calibri'
  });
}

// ==========================================
// SLIDE 13: CAUSES (ETIOLOGY)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CAUSES & ETIOLOGY OF COLON CANCER', 13);

  slide.addText('Colon cancer has no single cause; it develops when genetic and epigenetic changes accumulate over time:', {
    x: 0.6, y: 1.85, w: 12.13, h: 0.45,
    fontSize: 22, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });

  const causes = [
    { num: '1', title: 'Sporadic Cases (~70–75%)', desc: 'Acquired somatic mutations over lifetime via adenoma-carcinoma sequence (APC, KRAS, SMAD4, TP53).' },
    { num: '2', title: 'Hereditary Syndromes (~5–10%)', desc: 'Lynch syndrome (germline MMR mutation), FAP (germline APC mutation), MUTYH-associated polyposis.' },
    { num: '3', title: 'Familial Clustering (~10–30%)', desc: 'Clustering in families without a defined single-gene syndrome (shared genetic traits and environmental exposures).' },
    { num: '4', title: 'Chronic Intestinal Inflammation', desc: 'IBD (Ulcerative colitis, Crohn\'s) causing repeated mucosal injury, rapid epithelial regeneration, and dysplasia.' },
    { num: '5', title: 'Environmental DNA Damage', desc: 'Carcinogens in processed meat (heterocyclic amines, N-nitroso), tobacco toxins, and microbiome shifts (Fusobacterium).' }
  ];

  causes.forEach((c, idx) => {
    const y = 2.45 + idx * 0.92;
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
      { text: `${c.title}: `, options: { fontSize: 18, bold: true, color: BURGUNDY } },
      { text: c.desc, options: { fontSize: 18, color: BODY_TEXT } }
    ], { x: 1.45, y: y + 0.1, w: 11.1, h: 0.62, fontFace: 'Calibri', valign: 'middle' });
  });
}

// ==========================================
// SLIDE 14: TYPES OF COLON CANCER
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'HISTOLOGICAL TYPES OF COLON CANCER', 14);

  // Main Card: Adenocarcinoma (~95%)
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 1.85, w: 12.13, h: 1.8,
    fill: { color: LIGHT_PINK }, line: { color: BURGUNDY, width: 2 }, rectRadius: 0.1
  });
  slide.addText('ADENOCARCINOMA (~95% of all cases) — The Predominant Type', {
    x: 0.9, y: 2.0, w: 11.53, h: 0.45,
    fontSize: 24, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '• Cell of Origin: ', options: { fontSize: 19, bold: true, color: DARK_TEXT } },
    { text: 'Develops in the mucus-secreting glandular epithelial cells lining the colon and rectum.\n', options: { fontSize: 19, color: BODY_TEXT } },
    { text: '• Mucinous Adenocarcinoma: ', options: { fontSize: 19, bold: true, color: DARK_TEXT } },
    { text: 'Characterized by abundant extracellular mucin (>50% of tumor mass).\n', options: { fontSize: 19, color: BODY_TEXT } },
    { text: '• Signet Ring Cell Carcinoma: ', options: { fontSize: 19, bold: true, color: DARK_TEXT } },
    { text: 'Rare, aggressive subtype with intracellular mucin pushing nucleus to periphery (poorer prognosis).', options: { fontSize: 19, color: BODY_TEXT } }
  ], { x: 0.9, y: 2.5, w: 11.53, h: 1.05, fontFace: 'Calibri' });

  // 4 Rare Types Cards
  slide.addText('Less Common Types of Colorectal Tumors:', {
    x: 0.6, y: 3.8, w: 12.13, h: 0.4,
    fontSize: 22, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });

  const rareTypes = [
    { title: '1. Carcinoid Tumors', desc: 'Develop from specialized neuroendocrine hormone-producing cells in the digestive tract.' },
    { title: '2. Gastrointestinal Stromal (GISTs)', desc: 'Originate in the interstitial cells of Cajal in bowel wall (rare in colon, more common in stomach).' },
    { title: '3. Primary Colorectal Lymphomas', desc: 'Malignancies arising from immune-system lymphatic cells within the bowel wall.' },
    { title: '4. Colorectal Sarcomas', desc: 'Very rare cancers arising from blood vessels, smooth muscle layers, or connective tissues.' }
  ];

  const cW = 5.85;
  const cH = 1.35;
  const startX = 0.6;
  const gapX = 0.43;

  rareTypes.forEach((t, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = startX + col * (cW + gapX);
    const y = 4.3 + row * (cH + 0.2);

    slide.addShape(pptx.ShapeType.rect, {
      x, y, w: cW, h: cH,
      fill: { color: 'F9F9F9' }, line: { color: 'DDDDDD', width: 1 }, rectRadius: 0.08
    });
    slide.addText(t.title, {
      x: x + 0.25, y: y + 0.15, w: cW - 0.5, h: 0.35,
      fontSize: 18, bold: true, color: BURGUNDY, fontFace: 'Calibri'
    });
    slide.addText(t.desc, {
      x: x + 0.25, y: y + 0.55, w: cW - 0.5, h: 0.7,
      fontSize: 15, color: BODY_TEXT, fontFace: 'Calibri'
    });
  });
}

// ==========================================
// SLIDE 15: SIGNS AND SYMPTOMS
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CLINICAL SIGNS AND SYMPTOMS', 15);

  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  // Left Column (1-6)
  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 1.85, w: colW, h: 4.2,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.1
  });
  slide.addText([
    { text: '1. Blood in Stool / Rectal Bleeding\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '   Bright red bleeding or dark, tarry stool (melena).\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '2. Change in Characteristics of Stool\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '   Darkened or black stool signaling internal bleeding.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '3. Persistent Fatigue or Weakness\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '   Due to chronic occult blood loss and iron-deficiency.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '4. Sudden Unexplained Weight Loss\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '   Significant drop in body mass without dietary changes.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '5. Persistent Feeling of Fullness / Bloating\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '   Constant urge to evacuate without relief.', options: { fontSize: 16, color: BODY_TEXT } }
  ], { x: startX + 0.3, y: 2.0, w: colW - 0.6, h: 3.9, fontFace: 'Calibri' });

  // Right Column (6-11)
  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 1.85, w: colW, h: 4.2,
    fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.1
  });
  slide.addText([
    { text: '6. Change in Bowel Habits\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '   Persistent diarrhea, worsening constipation, or alternating.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '7. Persistent Abdominal Pain or Cramping\n\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '8. Narrow or "Pencil-Shaped" Stools\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '   Caused by progressive tumor narrowing the lumen.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '9. Tenesmus (Persistent Urge to Defecate)\n\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '10. Iron-Deficiency Anemia & Pallor\n\n', options: { fontSize: 20, bold: true, color: BURGUNDY } },
    { text: '11. Anorexia & Loss of Appetite', options: { fontSize: 20, bold: true, color: BURGUNDY } }
  ], { x: startX + colW + gapX + 0.3, y: 2.0, w: colW - 0.6, h: 3.9, fontFace: 'Calibri' });

  // Alert Box
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 6.18, w: 12.13, h: 0.85,
    fill: { color: 'FFF3E0' }, line: { color: 'E65100', width: 1.5 }, rectRadius: 0.08
  });
  slide.addText('⚠️ CRITICAL CLINICAL ALERT: Many patients experience NO symptoms at all in early stages! Right-sided tumors often present only with anemia. Regular screening is vital for anyone aged ≥ 45 or with family history.', {
    x: 0.8, y: 6.25, w: 11.73, h: 0.7,
    fontSize: 16, bold: true, color: 'BF360C', align: 'center', valign: 'middle', fontFace: 'Calibri'
  });
}

// ==========================================
// SLIDE 16: COMPLICATIONS
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'MAJOR COMPLICATIONS OF COLON CANCER', 16);

  const complications = [
    { num: '1', title: 'Mechanical Bowel Obstruction', desc: 'Growing tumor blocks the intestinal lumen, causing severe abdominal pain, distension, nausea, vomiting, and inability to pass stool or flatus.' },
    { num: '2', title: 'Gastrointestinal Bleeding', desc: 'Internal ulceration leads to severe iron-deficiency anemia, profound fatigue, hematochezia/melena, and possible hypovolemic shock.' },
    { num: '3', title: 'Intestinal Perforation (Surgical Emergency)', desc: 'Tumor necrosis or high intraluminal pressure causes bowel rupture, spilling bacteria into peritoneum and causing fatal peritonitis/sepsis.' },
    { num: '4', title: 'Metastasis (Cancer Dissemination)', desc: 'Malignant cells spread via lymphatics and portal bloodstream to regional lymph nodes, peritoneum, and distant organs (liver, lungs).' },
    { num: '5', title: 'Severe Malnutrition & Cachexia', desc: 'Impaired nutrient absorption, anorexia, and systemic tumor hypermetabolism cause progressive muscle wasting and severe weight loss.' },
    { num: '6', title: 'Emotional & Psychological Impact', desc: 'Fear of cancer diagnosis, rigorous treatments, and adjustment to potential permanent colostomy lead to anxiety and depression.' }
  ];

  complications.forEach((c, idx) => {
    const y = 1.85 + idx * 0.88;
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
      { text: `${c.title}: `, options: { fontSize: 18, bold: true, color: BURGUNDY } },
      { text: c.desc, options: { fontSize: 16, color: BODY_TEXT } }
    ], { x: 1.45, y: y + 0.1, w: 11.1, h: 0.6, fontFace: 'Calibri', valign: 'middle' });
  });
}

// ==========================================
// SLIDE 17: DIAGNOSTIC TESTS
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'DIAGNOSTIC WORKUP & CLINICAL EVALUATION', 17);

  slide.addText([
    { text: '1. Clinical History & Physical Exam:\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Symptom timeline, family history, abdominal palpation for masses, and Digital Rectal Examination (DRE).\n\n', options: { fontSize: 17, color: BODY_TEXT } },
    { text: '2. Laboratory Investigations:\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Complete Blood Count (CBC): Identifies microcytic anemia.\n• Comprehensive Metabolic Panel (CMP): Evaluates liver & renal functions.\n• Carcinoembryonic Antigen (CEA): Tumor marker for baseline & recurrence.\n\n', options: { fontSize: 17, color: BODY_TEXT } },
    { text: '3. Colonoscopy with Biopsy (GOLD STANDARD):\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Direct visualization of entire colon and histopathological tissue sampling.\n\n', options: { fontSize: 17, color: BODY_TEXT } },
    { text: '4. Radiologic Staging Imaging:\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Contrast-enhanced CT scan of Chest/Abdomen/Pelvis, MRI, & PET scan.', options: { fontSize: 17, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.85, w: 7.4, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });

  if (ctScanImg) {
    slide.addImage({ data: ctScanImg, x: 8.3, y: 2.1, w: 4.4, h: 4.3, rounding: true });
    slide.addText('Abdominal CT Scan: Demonstrating Colonic Wall Thickening Mass', {
      x: 8.3, y: 6.5, w: 4.4, h: 0.4,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 18: STAGES OF COLON CANCER
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'STAGES OF COLON CANCER (TNM CLASSIFICATION)', 18);

  const stages = [
    { stage: 'Stage 0', title: 'Carcinoma in Situ', desc: 'Abnormal/precancerous cells confined strictly to inner mucosal layer (polyps).', color: '388E3C' },
    { stage: 'Stage I', title: 'Localized Invasion', desc: 'Cancer invaded submucosa (T1) or muscularis propria (T2); no lymph node spread (N0, M0).', color: '689F38' },
    { stage: 'Stage II', title: 'Through Bowel Wall', desc: 'Grown deeper through wall (IIA: through muscle; IIB: penetrates serosa); no lymph node spread.', color: 'F57C00' },
    { stage: 'Stage III', title: 'Regional Lymph Nodes', desc: 'Cancer cells have spread to nearby regional mesenteric lymph nodes (N1–N2); no distant metastasis.', color: 'E65100' },
    { stage: 'Stage IV', title: 'Distant Metastasis', desc: 'Cancer spread to distant organs such as liver, lungs, peritoneum, or ovaries (M1).', color: 'D32F2F' }
  ];

  stages.forEach((s, i) => {
    const y = 1.85 + i * 1.02;

    // Stage Badge
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.6, y, w: 1.8, h: 0.9,
      fill: { color: s.color }, rectRadius: 0.08
    });
    slide.addText(s.stage, {
      x: 0.6, y, w: 1.8, h: 0.9,
      fontSize: 22, bold: true, color: WHITE, align: 'center', valign: 'middle', fontFace: 'Calibri'
    });

    // Description Box
    slide.addShape(pptx.ShapeType.rect, {
      x: 2.5, y, w: stagingDiagramImg ? 5.8 : 10.23, h: 0.9,
      fill: { color: LIGHT_PINK }, line: { color: 'E2CCD0', width: 1 }, rectRadius: 0.08
    });
    slide.addText([
      { text: `${s.title}: `, options: { fontSize: 18, bold: true, color: BURGUNDY } },
      { text: s.desc, options: { fontSize: 16, color: BODY_TEXT } }
    ], { x: 2.7, y: y + 0.1, w: stagingDiagramImg ? 5.4 : 9.8, h: 0.7, fontFace: 'Calibri', valign: 'middle' });
  });

  if (stagingDiagramImg) {
    slide.addImage({ data: stagingDiagramImg, x: 8.6, y: 1.85, w: 4.1, h: 4.8, rounding: true });
    slide.addText('TNM Staging Depth & Dissemination', {
      x: 8.6, y: 6.75, w: 4.1, h: 0.35,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 19: PROGNOSIS & SURVIVAL RATES
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PROGNOSIS & 5-YEAR SURVIVAL RATES', 19);

  slide.addText('5-Year Relative Survival Rates by Stage at Diagnosis:', {
    x: 0.6, y: 1.85, w: 12.13, h: 0.45,
    fontSize: 24, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });

  const survivalData = [
    { stage: 'Localized Stage (Confined to Colon)', rate: '91%', val: 0.91, color: '2E7D32', text: 'Cancer has not spread beyond colon wall (Stage I & selected Stage II).' },
    { stage: 'Regional Stage (Spread to Lymph Nodes)', rate: '73%', val: 0.73, color: 'F57C00', text: 'Cancer spread to nearby regional lymph nodes or surrounding tissues (Stage III).' },
    { stage: 'Distant Stage (Metastatic Spread)', rate: '13%', val: 0.13, color: 'D32F2F', text: 'Cancer has metastasized to distant organs like liver or lungs (Stage IV).' }
  ];

  survivalData.forEach((d, idx) => {
    const y = 2.45 + idx * 1.35;

    slide.addText(d.stage, {
      x: 0.6, y, w: 7.5, h: 0.35,
      fontSize: 19, bold: true, color: DARK_TEXT, fontFace: 'Calibri'
    });
    slide.addText(d.text, {
      x: 0.6, y: y + 0.35, w: 7.5, h: 0.3,
      fontSize: 14, color: SUBTITLE_GRAY, fontFace: 'Calibri'
    });

    // Background track
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.6, y: y + 0.7, w: 9.5, h: 0.45,
      fill: { color: 'EAEAEA' }, rectRadius: 0.06
    });
    // Colored Progress bar
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.6, y: y + 0.7, w: 9.5 * d.val, h: 0.45,
      fill: { color: d.color }, rectRadius: 0.06
    });

    // Percentage Label
    slide.addText(d.rate, {
      x: 10.3, y: y + 0.6, w: 2.4, h: 0.65,
      fontSize: 28, bold: true, color: d.color, fontFace: 'Calibri', align: 'left', valign: 'middle'
    });
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 6.45, w: 12.13, h: 0.65,
    fill: { color: 'F5F5F5' }, line: { color: 'DDDDDD', width: 1 }, rectRadius: 0.08
  });
  slide.addText('Clinical Outlook: Early detection dramatically improves survival from 13% up to 91%! Regular surveillance is essential to monitor for recurrence.', {
    x: 0.8, y: 6.45, w: 11.73, h: 0.65,
    fontSize: 15, bold: true, color: BURGUNDY, align: 'center', valign: 'middle', fontFace: 'Calibri'
  });
}

// ==========================================
// SLIDE 20: MANAGEMENT AND TREATMENT
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'MANAGEMENT & TREATMENT MODALITIES', 20);

  slide.addText([
    { text: 'Surgical Management (Primary Curative Modality):\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Polypectomy: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Endoscopic removal of precancerous or early malignant polyps during colonoscopy.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Colectomy: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Surgical resection of tumor-bearing colon segment + regional lymph nodes followed by anastomosis or stoma.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Multivisceral Resection: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'En-bloc excision for locally advanced tumors invading adjacent structures.\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: 'Systemic & Targeted Therapies:\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Chemotherapy: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Adjuvant therapy for Stage III and high-risk Stage II to eliminate micrometastases.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Targeted Therapy: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Anti-EGFR and anti-VEGF agents guided by RAS/BRAF biomarker testing.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Radiation Therapy: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Primarily used in rectal cancer or for palliation of unresectable lesions.', options: { fontSize: 18, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.85, w: 7.4, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });

  if (surgeryAtlasImg) {
    slide.addImage({ data: surgeryAtlasImg, x: 8.3, y: 2.1, w: 4.4, h: 4.3, rounding: true });
    slide.addText('Surgical Partial Colectomy with Anastomosis', {
      x: 8.3, y: 6.5, w: 4.4, h: 0.4,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 21: WHEN TO SEEK MEDICAL ATTENTION
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'WHEN TO SEEK MEDICAL ATTENTION', 21);

  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  // Warning signs
  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 1.85, w: colW, h: 5.15,
    fill: { color: LIGHT_PINK }, line: { color: BURGUNDY, width: 1.5 }, rectRadius: 0.1
  });
  slide.addText('ALARM SYMPTOMS REQUIRING EVALUATION', {
    x: startX + 0.3, y: 2.05, w: colW - 0.6, h: 0.5,
    fontSize: 22, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Consult a healthcare provider promptly if experiencing:\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Persistent Abdominal Pain or Cramping\n\n', options: { fontSize: 20, bold: true, color: DARK_TEXT } },
    { text: '• Blood in Stool (Visible Red or Dark Melena)\n\n', options: { fontSize: 20, bold: true, color: DARK_TEXT } },
    { text: '• Persistent Change in Bowel Habits (> a few days)\n\n', options: { fontSize: 20, bold: true, color: DARK_TEXT } },
    { text: '• Unexplained Weight Loss\n\n', options: { fontSize: 20, bold: true, color: DARK_TEXT } },
    { text: '• Unexplained Persistent Fatigue, Weakness, or Anemia', options: { fontSize: 20, bold: true, color: DARK_TEXT } }
  ], { x: startX + 0.3, y: 2.65, w: colW - 0.6, h: 4.1, fontFace: 'Calibri' });

  // Questions to ask
  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 1.85, w: colW, h: 5.15,
    fill: { color: 'F5F7FA' }, line: { color: 'B0BEC5', width: 1.5 }, rectRadius: 0.1
  });
  slide.addText('CRITICAL QUESTIONS FOR ONCOLOGY TEAM', {
    x: startX + colW + gapX + 0.3, y: 2.05, w: colW - 0.6, h: 0.5,
    fontSize: 22, bold: true, color: '1A365D', fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'If diagnosed with colon cancer, patients and nurses should ask:\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '1. What is the exact stage of the cancer?\n', options: { fontSize: 20, bold: true, color: '1A365D' } },
    { text: '   (Depth of invasion, lymph node involvement, metastasis)\n\n', options: { fontSize: 16, color: SUBTITLE_GRAY } },
    { text: '2. What are the chances of cure?\n', options: { fontSize: 20, bold: true, color: '1A365D' } },
    { text: '   (Expected treatment outcomes, surgery vs. systemic therapy)\n\n', options: { fontSize: 16, color: SUBTITLE_GRAY } },
    { text: '3. What is the risk of cancer recurrence?\n', options: { fontSize: 20, bold: true, color: '1A365D' } },
    { text: '   (Post-treatment surveillance schedule, CEA, CT, colonoscopy)', options: { fontSize: 16, color: SUBTITLE_GRAY } }
  ], { x: startX + colW + gapX + 0.3, y: 2.65, w: colW - 0.6, h: 4.1, fontFace: 'Calibri' });
}

// ==========================================
// SLIDE 22: PREVENTION & SCREENING
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PREVENTION & EARLY DETECTION STRATEGIES', 22);

  slide.addText([
    { text: '1. Screening and Early Detection (Primary Strategy):\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Screening Age: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Average-risk individuals should begin screening at age 45 (earlier with family history).\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Gold Standard Interval: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Screening colonoscopy every 10 years if normal; allows finding and removing polyps before malignancy.\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '2. Lifestyle & Risk Reduction:\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Maintain healthy weight and engage in regular moderate-to-vigorous physical activity.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Dietary modifications: Consume vegetables, fruits, whole grains; strictly limit red & processed meats.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Avoid tobacco smoking and eliminate or minimize alcohol intake.\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '3. Supplements & Medical Considerations:\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Vitamins/Minerals: Folic acid, Vitamin D, Calcium, Magnesium show promise, but require medical advice.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Aspirin/NSAIDs: May reduce polyp risk, but NOT recommended for average risk due to GI bleeding harm.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Hormone Therapy: Carries cardiovascular/other risks; not a standalone preventive modality.', options: { fontSize: 18, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.85, w: 12.13, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });
}

// ==========================================
// SLIDE 23: NURSING ROLE: PERIOPERATIVE CARE
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'NURSING ROLE — PERIOPERATIVE CARE', 23);

  slide.addText([
    { text: '1. Comprehensive Nursing Assessment:\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Assess bowel habits, blood in stool, pain, nutritional status, and baseline labs.\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '2. Pre-operative Nursing Care:\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Bowel Preparation: Administer prescribed laxatives, oral cleansers, and enemas.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Patient Education: Explain surgical procedure, manage anxiety, teach incentive spirometry.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Stoma Site Marking: Collaborate with WOCN (wound, ostomy, continence nurse) prior to surgery.\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '3. Post-operative Nursing Care:\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Vital Signs & Wound Surveillance: Monitor surgical incision, drain output, and signs of bleeding/leak.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Stoma / Colostomy Care: Ensure stoma is pink/red and moist; educate patient on pouch management.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Complication Prevention: Administer analgesics, encourage early ambulation to prevent DVT & ileus.', options: { fontSize: 18, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.85, w: 7.4, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });

  if (nurseVitalsImg) {
    slide.addImage({ data: nurseVitalsImg, x: 8.3, y: 2.1, w: 4.4, h: 4.3, rounding: true });
    slide.addText('Clinical Bedside Post-Operative Patient Monitoring', {
      x: 8.3, y: 6.5, w: 4.4, h: 0.4,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 24: NURSING ROLE: TREATMENT & SUPPORT
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'NURSING ROLE — TREATMENT & PSYCHOSOCIAL SUPPORT', 24);

  slide.addText([
    { text: '4. Managing Treatment Side Effects:\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Chemotherapy & Radiation Support: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Manage nausea, vomiting, fatigue, and diarrhea with prescribed antiemetics and adequate IV/oral fluids.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Infection & Hematologic Monitoring: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Monitor CBC closely (WBC, absolute neutrophil count, platelets) for infection and bleeding risks.\n\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '5. Psychological and Rehabilitation Support:\n', options: { fontSize: 22, bold: true, color: BURGUNDY } },
    { text: '• Body Image & Coping Adaptation: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Help patient adjust to altered body image and lifestyle changes when living with temporary or permanent colostomy.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Holistic & Family-Centered Care: ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'Provide continuous emotional, psychological, and informational support to the patient and family throughout rehabilitation.', options: { fontSize: 18, color: BODY_TEXT } }
  ], {
    x: 0.6, y: 1.85, w: 7.4, h: 5.15,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.2
  });

  if (nurseCareImg) {
    slide.addImage({ data: nurseCareImg, x: 8.3, y: 2.1, w: 4.4, h: 4.3, rounding: true });
    slide.addText('Holistic Patient Rehabilitation & Psychological Counseling', {
      x: 8.3, y: 6.5, w: 4.4, h: 0.4,
      fontSize: 11, italic: true, color: SUBTITLE_GRAY, align: 'center', fontFace: 'Calibri'
    });
  }
}

// ==========================================
// SLIDE 25: NURSING PROCESS — ROADMAP
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'NURSING PROCESS — CLINICAL CARE PLANS', 25);

  slide.addText('Systematic Application of the Nursing Process in Colorectal Oncology:', {
    x: 0.6, y: 1.85, w: 12.13, h: 0.45,
    fontSize: 24, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });

  const carePlans = [
    { num: 'CP 1', title: 'Actual Diagnosis: Abdominal Pain', desc: 'Related to tumor mass effect, intestinal distension, or surgical resection. Focus on pain assessment, analgesia, and bowel comfort.', color: BURGUNDY },
    { num: 'CP 2', title: 'Actual Diagnosis: Cancer-Related Fatigue', desc: 'Related to decreased oxygen-carrying capacity secondary to occult GI bleeding and anemia. Focus on energy conservation & nutrition.', color: BURGUNDY },
    { num: 'CP 3', title: 'Potential Diagnosis: Risk for Deficient Fluid Volume', desc: 'Related to excessive losses through vomiting, diarrhea, or stoma output and inadequate oral intake. Focus on strict I&O and hydration.', color: '1A365D' },
    { num: 'CP 4', title: 'Potential Diagnosis: Risk for Dysfunctional GI Motility', desc: 'Related to colon tumor obstruction, luminal narrowing, or post-op ileus. Focus on bowel sound auscultation and obstruction prevention.', color: '1A365D' }
  ];

  carePlans.forEach((cp, idx) => {
    const y = 2.45 + idx * 1.15;
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
      { text: `${cp.title}\n`, options: { fontSize: 20, bold: true, color: cp.color } },
      { text: cp.desc, options: { fontSize: 16, color: BODY_TEXT } }
    ], { x: 2.0, y: y + 0.1, w: 10.5, h: 0.8, fontFace: 'Calibri' });
  });
}

// ==========================================
// SLIDE 26: CARE PLAN 1: ABDOMINAL PAIN (DIAGNOSIS & GOALS)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 1: ABDOMINAL PAIN — ASSESSMENT & GOALS', 26);

  // Diagnosis Box
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 1.85, w: 12.13, h: 2.0,
    fill: { color: LIGHT_PINK }, line: { color: BURGUNDY, width: 2 }, rectRadius: 0.1
  });
  slide.addText('ACTUAL NURSING DIAGNOSIS', {
    x: 0.9, y: 2.0, w: 11.53, h: 0.4,
    fontSize: 22, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Persistent abdominal pain ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'related to tumor growth and intestinal changes secondary to colon cancer, as evidenced by:\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Subjective Data: ', options: { fontSize: 18, bold: true, color: BURGUNDY } },
    { text: '"I have abdominal pain." • "The pain is constant and cramping."\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Objective Data: ', options: { fontSize: 18, bold: true, color: BURGUNDY } },
    { text: 'Abdominal tenderness on palpation, distension, tumor mass visible on imaging, changes in bowel habits.', options: { fontSize: 18, color: BODY_TEXT } }
  ], { x: 0.9, y: 2.45, w: 11.53, h: 1.3, fontFace: 'Calibri' });

  // Goals & Evaluation Side by Side
  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  // Expected Outcomes
  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 4.05, w: colW, h: 3.0,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EXPECTED OUTCOMES (GOALS)', {
    x: startX + 0.3, y: 4.2, w: colW - 0.6, h: 0.4,
    fontSize: 20, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Short-Term Goals:\n', options: { fontSize: 17, bold: true, color: DARK_TEXT } },
    { text: '• Patient will report reduced abdominal pain within 24–48 hours.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Patient will verbalize effective pain relief measures.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: 'Long-Term Goals:\n', options: { fontSize: 17, bold: true, color: DARK_TEXT } },
    { text: '• Patient will maintain consistent pain control (score ≤ 3/10).\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Patient will demonstrate improved daily quality of life.', options: { fontSize: 16, color: BODY_TEXT } }
  ], { x: startX + 0.3, y: 4.65, w: colW - 0.6, h: 2.3, fontFace: 'Calibri' });

  // Evaluation
  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 4.05, w: colW, h: 3.0,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EVALUATION (AFTER REASSESSMENT)', {
    x: startX + colW + gapX + 0.3, y: 4.2, w: colW - 0.6, h: 0.4,
    fontSize: 20, bold: true, color: SUCCESS_GREEN, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Clinical Evaluation Results:\n', options: { fontSize: 17, bold: true, color: DARK_TEXT } },
    { text: '✔ Patient reported reduced abdominal pain within 24–48 hours.\n\n', options: { fontSize: 16, color: SUCCESS_GREEN } },
    { text: '✔ Patient verbalized effective pain relief measures.\n\n', options: { fontSize: 16, color: SUCCESS_GREEN } },
    { text: '✔ Patient maintained consistent pain control.\n\n', options: { fontSize: 16, color: SUCCESS_GREEN } },
    { text: '✔ Patient demonstrated improved comfort and quality of life.', options: { fontSize: 16, color: SUCCESS_GREEN } }
  ], { x: startX + colW + gapX + 0.3, y: 4.65, w: colW - 0.6, h: 2.3, fontFace: 'Calibri' });
}

// ==========================================
// SLIDE 27: CARE PLAN 1: ABDOMINAL PAIN (INTERVENTIONS & RATIONALES)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 1: ABDOMINAL PAIN — INTERVENTIONS', 27);

  const painTable = [
    [
      { text: 'Nursing Interventions', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 18 } },
      { text: 'Scientific Rationales', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 18 } }
    ],
    [
      { text: '1. Assess patient\'s pain level, location, and characteristics.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To determine the severity and pattern of pain and evaluate relief.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '2. Monitor vital signs regularly (blood pressure, heart rate).', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To detect hemodynamic changes and acute surgical complications.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '3. Administer prescribed analgesics as ordered.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To relieve pain and improve patient comfort and participation.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '4. Provide comfort measures (positioning, heat or cold).', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To reduce abdominal wall muscle tension and alleviate discomfort.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '5. Encourage eating small, frequent meals.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To prevent stomach overloading and minimize postprandial cramping.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '6. Choose low-fiber foods during acute flare-ups.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To minimize bowel irritation, gas formation, and distension.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '7. Report any sudden increase in pain or peritoneal symptoms.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To allow prompt intervention and prevent bowel perforation/obstruction.', options: { fontSize: 16, color: BODY_TEXT } }
    ]
  ];

  slide.addTable(painTable, {
    x: 0.6, y: 1.85, w: 12.13,
    colW: [6.06, 6.07],
    border: { type: 'solid', color: 'D8BFC4', pt: 1 },
    rowH: [0.55, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6],
    autoPage: false
  });
}

// ==========================================
// SLIDE 28: CARE PLAN 2: FATIGUE (DIAGNOSIS & GOALS)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 2: FATIGUE — ASSESSMENT & GOALS', 28);

  // Diagnosis Box
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 1.85, w: 12.13, h: 2.0,
    fill: { color: LIGHT_PINK }, line: { color: BURGUNDY, width: 2 }, rectRadius: 0.1
  });
  slide.addText('ACTUAL NURSING DIAGNOSIS: FATIGUE', {
    x: 0.9, y: 2.0, w: 11.53, h: 0.4,
    fontSize: 22, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Fatigue related to ', options: { fontSize: 18, bold: true, color: DARK_TEXT } },
    { text: 'decreased oxygen-carrying capacity secondary to anemia, as evidenced by:\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Subjective Data: ', options: { fontSize: 18, bold: true, color: BURGUNDY } },
    { text: '"I feel tired most of the time." • "I feel weak and exhausted." • Dizziness • Trouble with daily tasks.\n', options: { fontSize: 18, color: BODY_TEXT } },
    { text: '• Objective Data: ', options: { fontSize: 18, bold: true, color: BURGUNDY } },
    { text: 'Pallor of skin and conjunctiva, low Hb & Hct, increased resting heart rate, decreased activity tolerance.', options: { fontSize: 18, color: BODY_TEXT } }
  ], { x: 0.9, y: 2.45, w: 11.53, h: 1.3, fontFace: 'Calibri' });

  // Goals & Evaluation Side by Side
  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 4.05, w: colW, h: 3.0,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EXPECTED OUTCOMES (GOALS)', {
    x: startX + 0.3, y: 4.2, w: colW - 0.6, h: 0.4,
    fontSize: 20, bold: true, color: BURGUNDY, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Short-Term Goals:\n', options: { fontSize: 17, bold: true, color: DARK_TEXT } },
    { text: '• Patient will report decreased fatigue within 24–48 hours.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Patient will verbalize importance of adequate rest.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Patient will demonstrate energy conservation techniques.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: 'Long-Term Goals:\n', options: { fontSize: 17, bold: true, color: DARK_TEXT } },
    { text: '• Patient will perform daily activities with minimal fatigue.\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Patient\'s Hb level will improve according to treatment plan.', options: { fontSize: 16, color: BODY_TEXT } }
  ], { x: startX + 0.3, y: 4.65, w: colW - 0.6, h: 2.3, fontFace: 'Calibri' });

  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 4.05, w: colW, h: 3.0,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EVALUATION (AFTER REASSESSMENT)', {
    x: startX + colW + gapX + 0.3, y: 4.2, w: colW - 0.6, h: 0.4,
    fontSize: 20, bold: true, color: SUCCESS_GREEN, fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Clinical Evaluation Results:\n', options: { fontSize: 17, bold: true, color: DARK_TEXT } },
    { text: '✔ Patient reported decreased fatigue within 24–48 hours.\n\n', options: { fontSize: 16, color: SUCCESS_GREEN } },
    { text: '✔ Patient verbalized understanding of rest & energy conservation.\n\n', options: { fontSize: 16, color: SUCCESS_GREEN } },
    { text: '✔ Patient performed daily activities with improved tolerance.\n\n', options: { fontSize: 16, color: SUCCESS_GREEN } },
    { text: '✔ Patient\'s Hb level improved according to treatment plan.', options: { fontSize: 16, color: SUCCESS_GREEN } }
  ], { x: startX + colW + gapX + 0.3, y: 4.65, w: colW - 0.6, h: 2.3, fontFace: 'Calibri' });
}

// ==========================================
// SLIDE 29: CARE PLAN 2: FATIGUE (INTERVENTIONS & RATIONALES)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 2: FATIGUE — INTERVENTIONS', 29);

  const fatigueTable = [
    [
      { text: 'Nursing Interventions', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 18 } },
      { text: 'Scientific Rationales', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 18 } }
    ],
    [
      { text: '1. Assess the patient\'s level of fatigue and weakness.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To determine the baseline severity and impact on function.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '2. Monitor vital signs regularly (pulse, blood pressure).', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To detect tachycardia, hypotension, and anemia-induced strain.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '3. Monitor CBC, Hemoglobin (Hb), and Hematocrit (Hct) levels.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To evaluate severity of anemia and guide corrective therapies.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '4. Provide adequate, scheduled rest periods between activities.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To conserve metabolic energy and prevent physical exhaustion.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '5. Assist the patient with daily activities when needed.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To prevent excessive fatigue and protect patient from falls.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '6. Encourage iron-rich foods (lean meat, spinach, lentils).', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To support hemoglobin synthesis and replenish iron stores.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '7. Administer prescribed iron supplements / transfusions.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To correct iron deficiency and restore oxygen transport.', options: { fontSize: 15, color: BODY_TEXT } }
    ],
    [
      { text: '8. Encourage gradual physical activity and light walking as tolerated.', options: { fontSize: 15, color: DARK_TEXT } },
      { text: 'To prevent deconditioning and improve functional endurance.', options: { fontSize: 15, color: BODY_TEXT } }
    ]
  ];

  slide.addTable(fatigueTable, {
    x: 0.6, y: 1.85, w: 12.13,
    colW: [6.06, 6.07],
    border: { type: 'solid', color: 'D8BFC4', pt: 1 },
    rowH: [0.55, 0.55, 0.55, 0.55, 0.55, 0.55, 0.55, 0.55, 0.55],
    autoPage: false
  });
}

// ==========================================
// SLIDE 30: CARE PLAN 3: DEFICIENT FLUID VOLUME (GOALS & EVAL)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 3: DEFICIENT FLUID VOLUME — GOALS', 30);

  // Diagnosis Box
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 1.85, w: 12.13, h: 1.8,
    fill: { color: 'F4F6F9' }, line: { color: '1A365D', width: 2 }, rectRadius: 0.1
  });
  slide.addText('POTENTIAL NURSING DIAGNOSIS: DEFICIENT FLUID VOLUME', {
    x: 0.9, y: 2.0, w: 11.53, h: 0.4,
    fontSize: 22, bold: true, color: '1A365D', fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Risk for Deficient Fluid Volume ', options: { fontSize: 20, bold: true, color: DARK_TEXT } },
    { text: 'related to excessive fluid losses through vomiting, diarrhea, high stoma output, and inadequate oral fluid intake secondary to colon cancer pathology and systemic treatments.', options: { fontSize: 20, color: BODY_TEXT } }
  ], { x: 0.9, y: 2.5, w: 11.53, h: 1.0, fontFace: 'Calibri' });

  // Goals & Evaluation Side by Side
  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 3.85, w: colW, h: 3.2,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EXPECTED OUTCOMES (GOALS)', {
    x: startX + 0.3, y: 4.0, w: colW - 0.6, h: 0.4,
    fontSize: 20, bold: true, color: '1A365D', fontFace: 'Calibri'
  });
  slide.addText([
    { text: '• Patient will maintain adequate hydration and stable fluid balance within 24–48 hours.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Patient will verbalize understanding of drinking adequate fluids.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Patient will remain free from signs and symptoms of dehydration and electrolyte imbalance.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Patient will maintain moist mucous membranes and balanced intake and output (I&O).', options: { fontSize: 16, color: BODY_TEXT } }
  ], { x: startX + 0.3, y: 4.45, w: colW - 0.6, h: 2.5, fontFace: 'Calibri' });

  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 3.85, w: colW, h: 3.2,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EVALUATION (AFTER REASSESSMENT)', {
    x: startX + colW + gapX + 0.3, y: 4.0, w: colW - 0.6, h: 0.4,
    fontSize: 20, bold: true, color: SUCCESS_GREEN, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '✔ Patient is maintaining adequate hydration and stable fluid balance.\n\n', options: { fontSize: 16, color: SUCCESS_GREEN } },
    { text: '✔ Patient exhibits stable vital signs (BP & HR within normal limits).\n\n', options: { fontSize: 16, color: SUCCESS_GREEN } },
    { text: '✔ Strict intake and output (I&O) remain balanced without deficits.\n\n', options: { fontSize: 16, color: SUCCESS_GREEN } },
    { text: '✔ Patient and family verbalize the importance of hydration.', options: { fontSize: 16, color: SUCCESS_GREEN } }
  ], { x: startX + colW + gapX + 0.3, y: 4.45, w: colW - 0.6, h: 2.5, fontFace: 'Calibri' });
}

// ==========================================
// SLIDE 31: CARE PLAN 3: DEFICIENT FLUID VOLUME (INTERVENTIONS)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 3: DEFICIENT FLUID VOLUME — ACTIONS', 31);

  const fluidTable = [
    [
      { text: 'Nursing Interventions', options: { bold: true, color: WHITE, fill: { color: '1A365D' }, fontSize: 18 } },
      { text: 'Scientific Rationales', options: { bold: true, color: WHITE, fill: { color: '1A365D' }, fontSize: 18 } }
    ],
    [
      { text: '1. Monitor vital signs (especially BP and HR) regularly.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To detect early signs of hypovolemia and hemodynamic instability.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '2. Measure and record strict intake and output (I&O) every shift.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To accurately evaluate fluid balance, loss volume, and renal perfusion.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '3. Assess skin turgor, mucous membranes, and capillary refill.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To identify physical clinical indicators of dehydration.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '4. Monitor laboratory reports (electrolytes, Hb, Hct).', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To monitor for electrolyte shifts and hemoconcentration.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '5. Encourage frequent oral fluid intake as tolerated.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To maintain hydration status safely via the natural enteral route.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '6. Administer prescribed IV fluids and electrolyte replacements.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To restore intravascular volume and correct chemical imbalances.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '7. Educate patient/family on early symptoms of dehydration.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To empower prompt self-recognition (extreme thirst, dark urine).', options: { fontSize: 16, color: BODY_TEXT } }
    ]
  ];

  slide.addTable(fluidTable, {
    x: 0.6, y: 1.85, w: 12.13,
    colW: [6.06, 6.07],
    border: { type: 'solid', color: 'B0BEC5', pt: 1 },
    rowH: [0.55, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6],
    autoPage: false
  });
}

// ==========================================
// SLIDE 32: CARE PLAN 4: DYSFUNCTIONAL GI MOTILITY (GOALS & EVAL)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 4: DYSFUNCTIONAL GI MOTILITY — GOALS', 32);

  // Diagnosis Box
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 1.85, w: 12.13, h: 1.8,
    fill: { color: 'F4F6F9' }, line: { color: '1A365D', width: 2 }, rectRadius: 0.1
  });
  slide.addText('POTENTIAL NURSING DIAGNOSIS: DYSFUNCTIONAL GI MOTILITY', {
    x: 0.9, y: 2.0, w: 11.53, h: 0.4,
    fontSize: 22, bold: true, color: '1A365D', fontFace: 'Calibri'
  });
  slide.addText([
    { text: 'Risk for Dysfunctional Gastrointestinal Motility ', options: { fontSize: 20, bold: true, color: DARK_TEXT } },
    { text: 'related to the presence of a colon tumor mass, potential mechanical narrowing of the intestinal lumen, surgical bowel manipulation, and opioid analgesia.', options: { fontSize: 20, color: BODY_TEXT } }
  ], { x: 0.9, y: 2.5, w: 11.53, h: 1.0, fontFace: 'Calibri' });

  // Goals & Evaluation Side by Side
  const colW = 5.85;
  const startX = 0.6;
  const gapX = 0.43;

  slide.addShape(pptx.ShapeType.rect, {
    x: startX, y: 3.85, w: colW, h: 3.2,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EXPECTED OUTCOMES (GOALS)', {
    x: startX + 0.3, y: 4.0, w: colW - 0.6, h: 0.4,
    fontSize: 20, bold: true, color: '1A365D', fontFace: 'Calibri'
  });
  slide.addText([
    { text: '• Patient will maintain stable gastrointestinal motility and bowel function.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Patient will exhibit active bowel sounds in all 4 abdominal quadrants.\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Patient will remain free from signs of bowel obstruction (distension, vomiting, obstipation).\n\n', options: { fontSize: 16, color: BODY_TEXT } },
    { text: '• Patient and family will verbalize warning signs requiring immediate reporting.', options: { fontSize: 16, color: BODY_TEXT } }
  ], { x: startX + 0.3, y: 4.45, w: colW - 0.6, h: 2.5, fontFace: 'Calibri' });

  slide.addShape(pptx.ShapeType.rect, {
    x: startX + colW + gapX, y: 3.85, w: colW, h: 3.2,
    fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', width: 1 }, rectRadius: 0.08
  });
  slide.addText('EVALUATION (AFTER REASSESSMENT)', {
    x: startX + colW + gapX + 0.3, y: 4.0, w: colW - 0.6, h: 0.4,
    fontSize: 20, bold: true, color: SUCCESS_GREEN, fontFace: 'Calibri'
  });
  slide.addText([
    { text: '✔ Patient maintained stable gastrointestinal function.\n\n', options: { fontSize: 16, color: SUCCESS_GREEN } },
    { text: '✔ Active bowel sounds confirmed auscultated across all four quadrants.\n\n', options: { fontSize: 16, color: SUCCESS_GREEN } },
    { text: '✔ Patient remained completely free from obstruction symptoms.\n\n', options: { fontSize: 16, color: SUCCESS_GREEN } },
    { text: '✔ Patient and family verbalized obstruction warning signs.', options: { fontSize: 16, color: SUCCESS_GREEN } }
  ], { x: startX + colW + gapX + 0.3, y: 4.45, w: colW - 0.6, h: 2.5, fontFace: 'Calibri' });
}

// ==========================================
// SLIDE 33: CARE PLAN 4: DYSFUNCTIONAL GI MOTILITY (INTERVENTIONS)
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CARE PLAN 4: DYSFUNCTIONAL GI MOTILITY — ACTIONS', 33);

  const motilityTable = [
    [
      { text: 'Nursing Interventions', options: { bold: true, color: WHITE, fill: { color: '1A365D' }, fontSize: 18 } },
      { text: 'Scientific Rationales', options: { bold: true, color: WHITE, fill: { color: '1A365D' }, fontSize: 18 } }
    ],
    [
      { text: '1. Auscultate bowel sounds in all 4 quadrants every 4 hours.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To detect early signs of hypomotility, ileus, or mechanical obstruction.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '2. Assess abdomen for distension, rigidity, and tenderness every shift.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To monitor gas/fluid accumulation behind a narrowed colon segment.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '3. Monitor and document frequency and consistency of stool and flatus.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To verify colon patency and the return of normal peristalsis.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '4. Encourage progressive ambulation and physical activity as tolerated.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To stimulate peristalsis and enhance natural gastrointestinal motility.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '5. Maintain adequate hydration via oral fluids or prescribed IV routes.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To prevent constipation and ensure soft stool passage through lumen.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '6. Educate patient to consume small, frequent meals rather than large.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To prevent overloading and obstructing the narrowed colonic lumen.', options: { fontSize: 16, color: BODY_TEXT } }
    ],
    [
      { text: '7. Teach patient/family to immediately report vomiting or failure to pass gas.', options: { fontSize: 16, color: DARK_TEXT } },
      { text: 'To ensure emergency intervention before bowel perforation occurs.', options: { fontSize: 16, color: BODY_TEXT } }
    ]
  ];

  slide.addTable(motilityTable, {
    x: 0.6, y: 1.85, w: 12.13,
    colW: [6.06, 6.07],
    border: { type: 'solid', color: 'B0BEC5', pt: 1 },
    rowH: [0.55, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6],
    autoPage: false
  });
}

// ==========================================
// SLIDE 34: SUMMARY
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'SUMMARY & CONCLUSION', 34);

  slide.addShape(pptx.ShapeType.rect, {
    x: 1.0, y: 2.1, w: 11.33, h: 4.6,
    fill: { color: LIGHT_PINK }, line: { color: BURGUNDY, width: 2 }, rectRadius: 0.15
  });

  slide.addText('KEY CLINICAL TAKEAWAYS', {
    x: 1.5, y: 2.4, w: 10.33, h: 0.5,
    fontSize: 26, bold: true, color: BURGUNDY, align: 'center', fontFace: 'Calibri'
  });

  slide.addText([
    { text: 'To sum up, colon cancer is a preventable disease through early screening.\n\n', options: { fontSize: 24, bold: true, color: BURGUNDY } },
    { text: 'As nurses, our role is crucial in supporting the patient before and after surgery.\n\n', options: { fontSize: 24, color: DARK_TEXT } },
    { text: 'By managing actual symptoms like pain and fatigue, and preventing potential complications like bowel obstruction, we can safely improve the patient\'s outcomes and quality of life.\n\n', options: { fontSize: 24, color: DARK_TEXT } },
    { text: 'Thank you.', options: { fontSize: 32, bold: true, color: BURGUNDY } }
  ], {
    x: 1.5, y: 3.1, w: 10.33, h: 3.2,
    fontFace: 'Calibri', align: 'center', valign: 'middle', lineSpacingMultiple: 1.25
  });
}

// ==========================================
// SLIDE 35: REFERENCES
// ==========================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'REFERENCES & AUTHORITATIVE SOURCES', 35);

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
    x: 0.6, y: 1.85, w: 12.13, h: 5.2,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });
}

// ==========================================
// SAVE FILE
// ==========================================
const outputPath = path.join(__dirname, '..', 'public', 'colon-cancer-presentation.pptx');

pptx.writeFile({ fileName: outputPath })
  .then(() => {
    const stats = fs.statSync(outputPath);
    console.log('SUCCESS! Professional PPTX generated successfully:');
    console.log('  Path:', outputPath);
    console.log('  Size:', (stats.size / 1024 / 1024).toFixed(2), 'MB');
    console.log('  Total Slides:', TOTAL_SLIDES);
  })
  .catch(err => {
    console.error('ERROR Generating PPTX:', err);
    process.exit(1);
  });
