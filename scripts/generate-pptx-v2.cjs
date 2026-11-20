// Professional PPTX Generator - Based on Word Document Content
// Font: Headings 32pt, Body 28pt
// Real medical images from the original PPTX

const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

// Colors - professional medical academic
const BURGUNDY = '8B1A2D';
const DARK_BG = '1a1a2e';
const CREAM = 'FFF8F0';
const LIGHT_PINK = 'F8E8E8';
const GOLD = 'D4A574';
const WHITE = 'FFFFFF';
const DARK_TEXT = '2D2D2D';
const BODY_TEXT = '333333';
const SUBTITLE_GRAY = '666666';

// Media paths from original PPTX
const MEDIA_DIR = 'C:\\Users\\ACER\\.gemini\\antigravity\\brain\\a3de671f-0b32-4188-9a89-a83f29331cc0\\scratch\\pptx_media';
const IMG_DIR = path.join(__dirname, '..', 'public', 'images');
const SLIDE_NUM_COLOR = 'CCCCCC'; // replacing FFFFFF99 which is invalid

function getMediaPath(filename) {
  const p = path.join(MEDIA_DIR, filename);
  if (fs.existsSync(p)) return p;
  return null;
}

function getImgPath(filename) {
  const p = path.join(IMG_DIR, filename);
  if (fs.existsSync(p)) return p;
  return null;
}

function imageToBase64(filepath) {
  if (!filepath || !fs.existsSync(filepath)) return null;
  const ext = path.extname(filepath).toLowerCase();
  let mime = 'image/png';
  if (ext === '.jpg' || ext === '.jpeg') mime = 'image/jpeg';
  const data = fs.readFileSync(filepath);
  return `data:${mime};base64,${data.toString('base64')}`;
}

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5
pptx.author = 'Faculty of Nursing - Al-Azhar University';
pptx.subject = 'Colon Cancer Nursing Seminar';
pptx.title = 'Colon Cancer - Comprehensive Clinical Seminar';

// Define slide masters
pptx.defineSlideMaster({
  title: 'CONTENT_SLIDE',
  background: { color: WHITE },
  objects: [
    // Top burgundy bar
    { rect: { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: BURGUNDY } } },
    // Bottom burgundy bar
    { rect: { x: 0, y: 7.3, w: '100%', h: 0.2, fill: { color: BURGUNDY } } },
  ]
});

// Logo images
const nursingLogo = imageToBase64(getImgPath('nursing-logo-opt.png') || getMediaPath('image2.png'));
const azharLogo = imageToBase64(getImgPath('azhar-logo-opt.png') || getMediaPath('image3.png'));
const deptLogo = imageToBase64(getImgPath('dept-logo-opt.png') || getMediaPath('image1.png'));

// Helper: add header to content slide
function addHeader(slide, title, slideNum, totalSlides) {
  // Top bar
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: BURGUNDY } });
  
  // Header background
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.08, w: '100%', h: 0.9, fill: { color: LIGHT_PINK } });
  
  // Logos
  if (nursingLogo) {
    slide.addImage({ data: nursingLogo, x: 0.3, y: 0.13, w: 0.7, h: 0.7 });
  }
  if (azharLogo) {
    slide.addImage({ data: azharLogo, x: 12.33, y: 0.13, w: 0.7, h: 0.7 });
  }
  
  // Institution names
  slide.addText([
    { text: 'Faculty of Nursing for Girls — Al-Azhar University', options: { fontSize: 10, color: BURGUNDY, bold: true } },
    { text: '\n', options: { fontSize: 4 } },
    { text: 'Medical-Surgical Nursing Department', options: { fontSize: 9, color: SUBTITLE_GRAY } },
    { text: '\n', options: { fontSize: 4 } },
    { text: 'كُليّة التمريض للبنات بالقاهرة  •  قسم التمريض الباطني والجراحي  •  جامعة الأزهر الشريف', options: { fontSize: 9, color: BURGUNDY } },
  ], { x: 1.2, y: 0.15, w: 10.9, h: 0.75, align: 'center', valign: 'middle' });
  
  // Title bar
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 1.0, w: '100%', h: 0.65, fill: { color: BURGUNDY } });
  slide.addText(title, { 
    x: 0.5, y: 1.0, w: 12.33, h: 0.65, 
    fontSize: 32, color: WHITE, bold: true, 
    align: 'left', valign: 'middle',
    fontFace: 'Calibri'
  });
  
  // Slide number
  slide.addText(`${slideNum} / ${totalSlides}`, {
    x: 11.5, y: 1.0, w: 1.5, h: 0.65,
    fontSize: 14, color: 'CCCCCC', align: 'right', valign: 'middle',
    fontFace: 'Calibri'
  });
  
  // Bottom bar
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 7.3, w: '100%', h: 0.2, fill: { color: BURGUNDY } });
  slide.addText('Colon Cancer — Nursing Seminar • Al-Azhar University', {
    x: 0.5, y: 7.3, w: 12.33, h: 0.2,
    fontSize: 8, color: 'CCCCCC', align: 'center', valign: 'middle',
    fontFace: 'Calibri'
  });
}

// Helper: add body text
function addBody(slide, texts, startY, opts) {
  const defaults = { x: 0.7, w: 11.9, fontSize: 28, color: BODY_TEXT, fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.3, ...opts };
  let y = startY || 1.85;
  for (const t of texts) {
    if (typeof t === 'string') {
      slide.addText(t, { ...defaults, y, h: 0.5 });
      y += 0.55;
    } else {
      slide.addText(t.text, { ...defaults, ...t, y, h: t.h || 0.5 });
      y += (t.h || 0.5) + 0.05;
    }
  }
  return y;
}

// Helper: add bullet items
function addBullets(slide, items, startY, opts) {
  const defaults = { x: 0.7, w: 11.9, fontSize: 28, color: BODY_TEXT, fontFace: 'Calibri', valign: 'top', bullet: true, lineSpacingMultiple: 1.2, ...opts };
  const textRuns = items.map(item => {
    if (typeof item === 'string') {
      return { text: item, options: { fontSize: defaults.fontSize, color: defaults.color, bullet: true, breakType: 'break' } };
    }
    return item;
  });
  
  const h = Math.min(5.2, items.length * 0.6);
  slide.addText(textRuns, { ...defaults, y: startY, h });
  return startY + h;
}

const TOTAL = 28;

// ============================================================
// SLIDE 1: COVER
// ============================================================
{
  const slide = pptx.addSlide();
  slide.background = { color: DARK_BG };
  
  // Decorative top line
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.06, fill: { color: GOLD } });
  
  // Logos on cover
  if (nursingLogo) slide.addImage({ data: nursingLogo, x: 0.4, y: 0.3, w: 1.0, h: 1.0 });
  if (azharLogo) slide.addImage({ data: azharLogo, x: 11.93, y: 0.3, w: 1.0, h: 1.0 });
  
  // Institution names 
  slide.addText([
    { text: 'Faculty of Nursing for Girls, Al-Azhar University', options: { fontSize: 14, color: GOLD, bold: true } },
    { text: '\nكُليّة التمريض للبنات بالقاهرة', options: { fontSize: 13, color: CREAM } },
    { text: '\nMedical-Surgical Nursing Department', options: { fontSize: 12, color: 'CCCCCC' } },
    { text: '\nقسم التمريض الباطني والجراحي', options: { fontSize: 12, color: CREAM } },
    { text: '\nجامعة الأزهر الشريف', options: { fontSize: 13, color: GOLD } },
  ], { x: 1.6, y: 0.3, w: 10.1, h: 1.3, align: 'center', valign: 'top' });
  
  // Divider line
  slide.addShape(pptx.ShapeType.rect, { x: 2, y: 1.8, w: 9.33, h: 0.02, fill: { color: GOLD } });
  
  // Title
  slide.addText('COLON CANCER', {
    x: 0.5, y: 2.1, w: 12.33, h: 1.4,
    fontSize: 54, color: WHITE, bold: true,
    align: 'center', valign: 'middle',
    fontFace: 'Calibri', charSpacing: 5
  });
  
  slide.addText('A Comprehensive Clinical Seminar on Colorectal Oncology,\nSurgical Interventions & Advanced Nursing Care Plans', {
    x: 1.5, y: 3.4, w: 10.33, h: 0.9,
    fontSize: 16, color: 'CCCCCC',
    align: 'center', valign: 'middle',
    fontFace: 'Calibri', italic: true
  });
  
  // Supervision block
  slide.addShape(pptx.ShapeType.rect, { x: 2.5, y: 4.6, w: 8.33, h: 0.03, fill: { color: '444466' } });
  
  slide.addText('UNDER SUPERVISION', {
    x: 2.5, y: 4.8, w: 8.33, h: 0.4,
    fontSize: 14, color: GOLD, bold: true,
    align: 'center', fontFace: 'Calibri', charSpacing: 3
  });
  
  slide.addText([
    { text: 'Prof. Dr. Heba Abdel-Azem', options: { fontSize: 16, color: WHITE, bold: true } },
    { text: '\nProfessor of Medical Surgical Nursing Department', options: { fontSize: 11, color: 'AAAAAA' } },
    { text: '\n\nDr. Eman Abel-Maqsoud', options: { fontSize: 16, color: WHITE, bold: true } },
    { text: '\nDemonstrator of Medical Surgical Nursing Department', options: { fontSize: 11, color: 'AAAAAA' } },
  ], { x: 2.5, y: 5.2, w: 8.33, h: 1.4, align: 'center', valign: 'top' });
  
  slide.addText('Academic Year 2026 / 2027', {
    x: 0, y: 6.9, w: 13.33, h: 0.4,
    fontSize: 14, color: GOLD, align: 'center', fontFace: 'Calibri'
  });
  
  // Bottom gold bar
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 7.4, w: '100%', h: 0.1, fill: { color: GOLD } });
}

// ============================================================
// SLIDE 2: TEAM
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PRESENTATION TEAM', 2, TOTAL);
  
  slide.addText('Prepared by:', {
    x: 0.7, y: 1.85, w: 11.9, h: 0.5,
    fontSize: 28, color: BURGUNDY, bold: true, fontFace: 'Calibri'
  });
  
  slide.addText('إعداد طالبات كلية التمريض — الفرقة الدراسية الثانية', {
    x: 0.7, y: 2.3, w: 11.9, h: 0.4,
    fontSize: 18, color: SUBTITLE_GRAY, fontFace: 'Calibri', align: 'right'
  });
  
  const names = [
    'Hager Saeed', 'Hager Hany', 'Hager Ahmed', 
    'Hager Mostafa', 'Haidy Mohammed', 'Norhan Badwy'
  ];
  
  // Cards layout - 3 columns x 2 rows
  const cols = 3;
  const cardW = 3.5;
  const cardH = 1.4;
  const startX = 0.9;
  const startY = 3.0;
  const gapX = 0.3;
  const gapY = 0.3;
  
  names.forEach((name, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);
    
    // Card background
    slide.addShape(pptx.ShapeType.rect, {
      x, y, w: cardW, h: cardH,
      fill: { color: LIGHT_PINK },
      rectRadius: 0.1,
      shadow: { type: 'outer', blur: 4, offset: 2, color: '888888' }
    });
    
    // Number circle
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.15, y: y + 0.3, w: 0.6, h: 0.6,
      fill: { color: BURGUNDY }
    });
    slide.addText(`${i + 1}`, {
      x: x + 0.15, y: y + 0.3, w: 0.6, h: 0.6,
      fontSize: 20, color: WHITE, bold: true, align: 'center', valign: 'middle'
    });
    
    // Name
    slide.addText(name, {
      x: x + 0.9, y: y + 0.25, w: cardW - 1.1, h: 0.5,
      fontSize: 20, color: DARK_TEXT, bold: true, fontFace: 'Calibri'
    });
    slide.addText('Nursing Student — Al-Azhar University', {
      x: x + 0.9, y: y + 0.7, w: cardW - 1.1, h: 0.4,
      fontSize: 12, color: SUBTITLE_GRAY, fontFace: 'Calibri'
    });
  });
}

// ============================================================
// SLIDE 3: GOAL
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'GOAL', 3, TOTAL);
  
  slide.addText('The goal of this assignment is to provide nursing students with a comprehensive understanding of colon cancer, including its anatomy, pathophysiology, risk factors, clinical manifestations, diagnosis, and management, in order to enhance their knowledge and ability to provide safe, effective, and evidence-based nursing care for patients with colon cancer.', {
    x: 0.7, y: 2.0, w: 11.9, h: 3.5,
    fontSize: 28, color: BODY_TEXT, fontFace: 'Calibri',
    valign: 'top', lineSpacingMultiple: 1.4,
    paraSpaceAfter: 10
  });
}

// ============================================================
// SLIDE 4: OBJECTIVES
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'OBJECTIVES', 4, TOTAL);
  
  slide.addText('By the end of this assignment, the student will be able to:', {
    x: 0.7, y: 1.85, w: 11.9, h: 0.6,
    fontSize: 28, color: BURGUNDY, bold: true, fontFace: 'Calibri'
  });
  
  const objectives = [
    'Define colon cancer and describe the anatomy and physiology of the colon.',
    'Explain the pathophysiology, causes, and risk factors of colon cancer.',
    'Identify the types, signs and symptoms, and complications of colon cancer.',
    'Discuss the diagnostic tests, staging, and prognosis of colon cancer.',
    'Describe the medical and surgical management and the preventive measures of colon cancer.',
    'Describe the role of the nurse in caring for patients with colon cancer before and after surgery.',
    'Identify preventive measures and methods of early detection of colon.',
    'Apply the nursing process in planning and providing individualized nursing care for patients with colon cancer.'
  ];
  
  const textRuns = objectives.map((obj, i) => ({
    text: `${i + 1}. ${obj}\n`,
    options: { fontSize: 22, color: BODY_TEXT, breakType: 'break', paraSpaceAfter: 4 }
  }));
  
  slide.addText(textRuns, {
    x: 0.7, y: 2.5, w: 11.9, h: 4.6,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });
}

// ============================================================
// SLIDE 5: OUTLINE
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'OUTLINE', 5, TOTAL);
  
  const outlineLeft = [
    '1. Introduction', '2. Definition', '3. Anatomy and Physiology',
    '4. Pathophysiology', '5. Risk Factors', '6. Causes',
    '7. Types', '8. Signs and Symptoms', '9. Complications'
  ];
  
  const outlineRight = [
    '10. Diagnostic Tests', '11. Prognosis',
    '12. Management and Treatment', '13. Prevention',
    '14. Nursing Role', '15. Nursing Process',
    '16. Summary', '17. References'
  ];
  
  const leftRuns = outlineLeft.map(item => ({
    text: item + '\n', options: { fontSize: 24, color: BODY_TEXT, paraSpaceAfter: 6 }
  }));
  
  const rightRuns = outlineRight.map(item => ({
    text: item + '\n', options: { fontSize: 24, color: BODY_TEXT, paraSpaceAfter: 6 }
  }));
  
  slide.addText(leftRuns, {
    x: 0.7, y: 1.85, w: 5.5, h: 5.0,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.2
  });
  
  slide.addText(rightRuns, {
    x: 6.8, y: 1.85, w: 5.5, h: 5.0,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.2
  });
}

// ============================================================
// SLIDE 6: INTRODUCTION
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'INTRODUCTION', 6, TOTAL);
  
  // Image (colonoscopy view - image8.jpeg)
  const introImg = imageToBase64(getMediaPath('image8.jpeg'));
  
  const textW = introImg ? 8.0 : 11.9;
  
  slide.addText([
    { text: 'Colon cancer is one of the major types of cancer affecting the digestive system. It develops when cells in the lining of the colon undergo abnormal changes that lead to uncontrolled growth.\n\n', options: { fontSize: 24, color: BODY_TEXT } },
    { text: 'In many cases, colon cancer begins as small growths called polyps, some of which may gradually develop into cancer over several years.\n\n', options: { fontSize: 24, color: BODY_TEXT } },
    { text: 'Early detection plays an important role in improving treatment outcomes. Screening tests can help identify precancerous polyps and detect cancer at an early stage.', options: { fontSize: 24, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 1.85, w: textW, h: 5.0,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.25
  });
  
  if (introImg) {
    slide.addImage({ data: introImg, x: 9.0, y: 2.2, w: 3.8, h: 4.2, rounding: true });
    slide.addText('Endoscopic Colonoscopy View', {
      x: 9.0, y: 6.5, w: 3.8, h: 0.4,
      fontSize: 10, color: SUBTITLE_GRAY, align: 'center', italic: true, fontFace: 'Calibri'
    });
  }
}

// ============================================================
// SLIDE 7: DEFINITION
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'DEFINITION', 7, TOTAL);
  
  const defImg = imageToBase64(getMediaPath('image9.jpeg'));
  const textW = defImg ? 8.0 : 11.9;
  
  slide.addText([
    { text: 'Colon cancer ', options: { fontSize: 28, color: BODY_TEXT, bold: true } },
    { text: 'is a malignant disease that develops in the tissues of the colon, which is the longest part of the large intestine.\n\n', options: { fontSize: 26, color: BODY_TEXT } },
    { text: 'It occurs when abnormal cells grow and divide uncontrollably, forming a malignant tumor.\n\n', options: { fontSize: 26, color: BODY_TEXT } },
    { text: 'Most colon cancers are ', options: { fontSize: 26, color: BODY_TEXT } },
    { text: 'adenocarcinomas', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: ', which develop from glandular cells lining the inner surface of the colon.', options: { fontSize: 26, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 1.85, w: textW, h: 5.0,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.25
  });
  
  if (defImg) {
    slide.addImage({ data: defImg, x: 9.0, y: 2.2, w: 3.8, h: 3.8, rounding: true });
    slide.addText('Histopathology: Colonic Adenocarcinoma (H&E Stain)', {
      x: 9.0, y: 6.1, w: 3.8, h: 0.5,
      fontSize: 9, color: SUBTITLE_GRAY, align: 'center', italic: true, fontFace: 'Calibri'
    });
  }
}

// ============================================================
// SLIDE 8: ANATOMY AND PHYSIOLOGY
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'ANATOMY AND PHYSIOLOGY', 8, TOTAL);
  
  const anatImg = imageToBase64(getMediaPath('image10.jpeg'));
  const textW = anatImg ? 7.5 : 11.9;
  
  slide.addText([
    { text: 'The colon is the main part of the large intestine (~1.5 m long). It begins at the cecum and continues through four main sections:\n\n', options: { fontSize: 24, color: BODY_TEXT } },
    { text: '1. Ascending colon: ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Located on the right side of the abdomen, carries intestinal contents upward from the cecum.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '2. Transverse colon: ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Extends across the upper abdomen from right to left.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '3. Descending colon: ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Runs downward along the left side of the abdomen.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '4. Sigmoid colon: ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'An S-shaped section connecting the descending colon to the rectum.', options: { fontSize: 22, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 1.85, w: textW, h: 5.2,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });
  
  if (anatImg) {
    slide.addImage({ data: anatImg, x: 8.5, y: 1.85, w: 4.5, h: 5.0 });
    slide.addText('Anatomical Divisions of the Colon', {
      x: 8.5, y: 6.9, w: 4.5, h: 0.3,
      fontSize: 10, color: SUBTITLE_GRAY, align: 'center', italic: true, fontFace: 'Calibri'
    });
  }
}

// ============================================================
// SLIDE 9: PATHOPHYSIOLOGY (Part 1 - Main Pathway)
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PATHOPHYSIOLOGY', 9, TOTAL);
  
  const pathoImg = imageToBase64(getMediaPath('image11.jpeg'));
  const textW = pathoImg ? 8.0 : 11.9;
  
  slide.addText([
    { text: '1. Main Pathway: Adenoma–Carcinoma Sequence (~80%)\n\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• APC mutation (early): ', options: { fontSize: 22, color: BURGUNDY, bold: true } },
    { text: 'loss of APC → β-catenin accumulates → Wnt signaling stays on → normal mucosa → small adenoma.\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '• KRAS mutation: ', options: { fontSize: 22, color: BURGUNDY, bold: true } },
    { text: 'continuous growth signaling → intermediate adenoma.\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '• SMAD4 / TGF-β pathway loss: ', options: { fontSize: 22, color: BURGUNDY, bold: true } },
    { text: 'loss of growth inhibition → late adenoma.\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '• p53 mutation (TP53): ', options: { fontSize: 22, color: BURGUNDY, bold: true } },
    { text: 'loss of apoptosis → adenoma → invasive carcinoma.\n\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '2. MMR Deficiency / Microsatellite Instability (~15%)\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: '• Defect in MLH1, MSH2, MSH6, PMS2 genes.\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '• Lynch syndrome: inherited germline MMR mutation.\n\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '3. Serrated Pathway / CIMP (~10–20%)\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: '• BRAF mutation → sessile serrated polyp. Typically right-sided colon.', options: { fontSize: 21, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 1.85, w: textW, h: 5.2,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.1
  });
  
  if (pathoImg) {
    slide.addImage({ data: pathoImg, x: 9.0, y: 2.0, w: 3.8, h: 4.5 });
    slide.addText('Adenoma–Carcinoma Progression', {
      x: 9.0, y: 6.6, w: 3.8, h: 0.3,
      fontSize: 10, color: SUBTITLE_GRAY, align: 'center', italic: true, fontFace: 'Calibri'
    });
  }
}

// ============================================================
// SLIDE 10: PATHOPHYSIOLOGY (Part 2 - Progression)
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PATHOPHYSIOLOGY (Continued)', 10, TOTAL);
  
  slide.addText([
    { text: '4. Hereditary Syndromes\n\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• FAP: ', options: { fontSize: 24, bold: true, color: BURGUNDY } },
    { text: 'germline APC mutation → hundreds to thousands of polyps → almost 100% cancer risk if untreated.\n', options: { fontSize: 23, color: BODY_TEXT } },
    { text: '• Lynch syndrome: inherited MMR gene mutation.\n\n', options: { fontSize: 23, color: BODY_TEXT } },
    { text: '5. Risk Factors That Drive These Changes\n\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• Age, family history, IBD (chronic inflammation → dysplasia).\n', options: { fontSize: 23, color: BODY_TEXT } },
    { text: '• Diet high in red/processed meat and low in fiber, obesity, smoking, alcohol.\n\n', options: { fontSize: 23, color: BODY_TEXT } },
    { text: '6. Progression\n\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• Local invasion: mucosa → submucosa → muscularis propria → serosa.\n', options: { fontSize: 23, color: BODY_TEXT } },
    { text: '• Spread via lymphatics to regional nodes, via blood (portal vein) to liver, then lung.\n\n', options: { fontSize: 23, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 1.85, w: 11.9, h: 4.5,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });
  
  // Key point box
  slide.addShape(pptx.ShapeType.rect, {
    x: 1.0, y: 6.2, w: 11.3, h: 0.8,
    fill: { color: 'FFF3E0' }, rectRadius: 0.08,
    line: { color: GOLD, width: 1.5 }
  });
  slide.addText('⚡ Key Point: The process takes roughly 10–15 years, which is why screening (colonoscopy) and polyp removal are effective at prevention.', {
    x: 1.2, y: 6.25, w: 10.9, h: 0.7,
    fontSize: 18, color: '8B6914', fontFace: 'Calibri', valign: 'middle', bold: true
  });
}

// ============================================================
// SLIDE 11: RISK FACTORS
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'RISK FACTORS', 11, TOTAL);
  
  // Two columns
  // Left: Non-modifiable
  slide.addText('Non-Modifiable', {
    x: 0.5, y: 1.85, w: 6.0, h: 0.5,
    fontSize: 28, color: BURGUNDY, bold: true, fontFace: 'Calibri'
  });
  
  slide.addText([
    { text: '• Age: risk rises after 50 (about 90% of cases).\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '• Family history: first-degree relative with CRC roughly doubles risk.\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '• Hereditary syndromes: Lynch, FAP, MUTYH-associated polyposis.\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '• Personal history: previous adenomas, CRC, or IBD.\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '• Other: male sex; African American ethnicity.', options: { fontSize: 20, color: BODY_TEXT } },
  ], {
    x: 0.5, y: 2.4, w: 6.0, h: 3.0,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.3
  });
  
  // Right: Modifiable
  slide.addText('Modifiable (Lifestyle)', {
    x: 6.8, y: 1.85, w: 6.0, h: 0.5,
    fontSize: 28, color: BURGUNDY, bold: true, fontFace: 'Calibri'
  });
  
  slide.addText([
    { text: '• Diet high in red and processed meat, low in fiber.\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '• Obesity (especially central/abdominal) and sedentary lifestyle.\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '• Smoking and heavy alcohol use.\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '• Type 2 diabetes and insulin resistance.\n\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: 'Protective Factors:\n', options: { fontSize: 22, color: BURGUNDY, bold: true } },
    { text: '• Regular screening with polyp removal.\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '• High-fiber diet, physical activity.\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: '• Aspirin/NSAIDs (in selected people).', options: { fontSize: 20, color: BODY_TEXT } },
  ], {
    x: 6.8, y: 2.4, w: 6.0, h: 4.0,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.3
  });
}

// ============================================================
// SLIDE 12: CAUSES
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'CAUSES', 12, TOTAL);
  
  slide.addText([
    { text: 'Colon cancer has no single cause. It develops when genetic and epigenetic changes accumulate in colonic epithelial cells.\n\n', options: { fontSize: 24, color: BODY_TEXT } },
    { text: '1. Sporadic (~70-75%): ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Acquired mutations (APC, KRAS, SMAD4, TP53) via adenoma-carcinoma sequence.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '2. Hereditary (~5-10%): ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Lynch syndrome (MMR genes), FAP (APC mutation), MUTYH-associated polyposis.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '3. Familial (~10-30%): ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Clustering in families without a defined syndrome.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '4. Chronic Inflammation: ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'IBD (UC, Crohn\'s) → dysplasia → carcinoma.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '5. Environmental Triggers: ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Red/processed meat carcinogens, low fiber, alcohol, smoking, obesity.', options: { fontSize: 22, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 1.85, w: 11.9, h: 5.2,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.1
  });
}

// ============================================================
// SLIDE 13: TYPES
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'TYPES', 13, TOTAL);
  
  // Main type highlight
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.5, y: 1.85, w: 12.33, h: 1.4,
    fill: { color: LIGHT_PINK }, rectRadius: 0.1
  });
  
  slide.addText([
    { text: 'Adenocarcinoma (~95%) — Predominant Type\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: 'Arises from mucus-producing glandular epithelial cells. Subtypes include mucinous adenocarcinoma and signet-ring cell carcinoma (worse prognosis).', options: { fontSize: 20, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 1.9, w: 11.9, h: 1.3,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.2
  });
  
  // Other types
  slide.addText([
    { text: 'Other Types:\n\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '1. Carcinoid Tumors: ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Develop from specialized hormone-producing cells in the intestine.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '2. GISTs (Gastrointestinal Stromal Tumors): ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Develop from specialized cells in the GI wall. Uncommon in colon.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '3. Lymphomas: ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Cancers from immune system cells. Usually in lymph nodes but can be in colon.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '4. Sarcomas: ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Rare cancers from blood vessels, muscle, or connective tissues in the colon wall.', options: { fontSize: 22, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 3.5, w: 11.9, h: 3.5,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.1
  });
}

// ============================================================
// SLIDE 14: SIGNS AND SYMPTOMS
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'SIGNS AND SYMPTOMS', 14, TOTAL);
  
  slide.addText([
    { text: '1. Blood in Stool or Bleeding from the Rectum\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: '2. Change in characteristics of the Stool — darkened or black stool\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: '3. Fatigue or Weakness — chronic internal bleeding → anemia\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: '4. Sudden Weight Loss — unexplained drop in energy\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: '5. Bloating or Urge to Have a Bowel Movement\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: '6. Change in bowel habits (diarrhea or constipation)\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: '7. Persistent abdominal pain or cramps\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: '8. Narrow or pencil-shaped stools\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: '9. Tenesmus (persistent feeling of needing to pass stool)\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: '10. Iron-deficiency anemia\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: '11. Loss of appetite\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
  ], {
    x: 0.7, y: 1.85, w: 11.9, h: 4.5,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.1
  });
  
  // Alert box
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.7, y: 6.2, w: 11.9, h: 0.8,
    fill: { color: 'FFF3E0' }, rectRadius: 0.08,
    line: { color: GOLD, width: 1.5 }
  });
  slide.addText('⚠️ Sometimes, NO Symptoms at All — Regular screening is crucial, especially if over 45 or have a family history.', {
    x: 0.9, y: 6.25, w: 11.5, h: 0.7,
    fontSize: 18, color: '8B6914', fontFace: 'Calibri', valign: 'middle', bold: true
  });
}

// ============================================================
// SLIDE 15: COMPLICATIONS
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'COMPLICATIONS', 15, TOTAL);
  
  slide.addText([
    { text: '1. Bowel Obstruction\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: 'A growing tumor can block the intestine, causing severe abdominal pain, swelling, nausea, and vomiting.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '2. Gastrointestinal Bleeding\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: 'Internal bleeding → iron-deficiency anemia, fatigue, bloody or black tarry stools.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '3. Perforation\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: 'Hole in intestinal wall → dangerous peritonitis.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '4. Metastasis (Cancer Spread)\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: 'Cancer spreads to lymph nodes, peritoneum, liver, and lungs.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '5. Malnutrition and Cachexia\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: 'Significant weight loss and nutritional deficiencies.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '6. Emotional and Psychological Impact\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: 'Anxiety and depression from diagnosis and treatment.', options: { fontSize: 22, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 1.85, w: 11.9, h: 5.2,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.05
  });
}

// ============================================================
// SLIDE 16: DIAGNOSTIC TESTS
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'DIAGNOSTIC TESTS', 16, TOTAL);
  
  const ctImg = imageToBase64(getMediaPath('image12.jpeg'));
  const textW = ctImg ? 8.0 : 11.9;
  
  slide.addText([
    { text: '1. Clinical Assessment\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• Ask about symptoms; examine abdomen for masses; perform DRE.\n\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '2. Blood Tests\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• CBC (anemia), CMP (organ function), CEA (tumor marker).\n\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '3. Imaging Tests\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• CT scan, MRI, PET scan, Abdominal ultrasound.\n\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '4. Diagnostic Colonoscopy (Gold Standard)\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• Flexible tube with camera examines colon. Biopsy confirms cancer.\n\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '5. Cancer Staging\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• Determines extent, depth of invasion, lymph node and distant spread.', options: { fontSize: 21, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 1.85, w: textW, h: 5.2,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.1
  });
  
  if (ctImg) {
    slide.addImage({ data: ctImg, x: 8.8, y: 2.2, w: 4.0, h: 4.0, rounding: true });
    slide.addText('Abdominal CT Scan', {
      x: 8.8, y: 6.3, w: 4.0, h: 0.3,
      fontSize: 10, color: SUBTITLE_GRAY, align: 'center', italic: true, fontFace: 'Calibri'
    });
  }
}

// ============================================================
// SLIDE 17: STAGING
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'STAGES OF COLON CANCER', 17, TOTAL);
  
  const stagingImg = imageToBase64(getMediaPath('image13.jpeg'));
  
  // Stage cards
  const stages = [
    { stage: 'Stage 0', desc: 'Abnormal or precancerous cells present in colon polyps.', color: '4CAF50' },
    { stage: 'Stage I', desc: 'Cancer has invaded the colon wall but has not spread to lymph nodes or distant organs.', color: '8BC34A' },
    { stage: 'Stage II', desc: 'Cancer has grown deeper through the colon wall (IIA: not reaching outer layer; IIB: reaches/grows through).', color: 'FFC107' },
    { stage: 'Stage III', desc: 'Cancer has spread to regional lymph nodes.', color: 'FF9800' },
    { stage: 'Stage IV', desc: 'Cancer has spread to distant organs (liver, lungs, or ovaries).', color: 'F44336' },
  ];
  
  const startY = 1.85;
  const cardH = 0.85;
  const gap = 0.15;
  
  stages.forEach((s, i) => {
    const y = startY + i * (cardH + gap);
    
    // Stage label
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y, w: 1.8, h: cardH,
      fill: { color: s.color }, rectRadius: 0.08
    });
    slide.addText(s.stage, {
      x: 0.5, y, w: 1.8, h: cardH,
      fontSize: 20, color: WHITE, bold: true, align: 'center', valign: 'middle', fontFace: 'Calibri'
    });
    
    // Description
    slide.addShape(pptx.ShapeType.rect, {
      x: 2.4, y, w: stagingImg ? 6.0 : 10.4, h: cardH,
      fill: { color: 'F5F5F5' }, rectRadius: 0.08
    });
    slide.addText(s.desc, {
      x: 2.6, y, w: stagingImg ? 5.6 : 10.0, h: cardH,
      fontSize: 18, color: BODY_TEXT, valign: 'middle', fontFace: 'Calibri'
    });
  });
  
  if (stagingImg) {
    slide.addImage({ data: stagingImg, x: 8.8, y: 2.0, w: 4.0, h: 4.5 });
    slide.addText('Colon Cancer Staging Progression', {
      x: 8.8, y: 6.6, w: 4.0, h: 0.3,
      fontSize: 10, color: SUBTITLE_GRAY, align: 'center', italic: true, fontFace: 'Calibri'
    });
  }
}

// ============================================================
// SLIDE 18: PROGNOSIS
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PROGNOSIS', 18, TOTAL);
  
  slide.addText('Survival Rates (5-Year)', {
    x: 0.7, y: 1.85, w: 11.9, h: 0.6,
    fontSize: 28, color: BURGUNDY, bold: true, fontFace: 'Calibri'
  });
  
  // Survival rate bars
  const rates = [
    { label: 'Localized cancer', rate: '91%', pct: 91, color: '4CAF50', desc: 'Cancer affects only the colon.' },
    { label: 'Regional cancer', rate: '73%', pct: 73, color: 'FFC107', desc: 'Spread to nearby organs or lymph nodes.' },
    { label: 'Distant cancer', rate: '13%', pct: 13, color: 'F44336', desc: 'Spread to distant organs (liver, lungs).' },
  ];
  
  rates.forEach((r, i) => {
    const y = 2.7 + i * 1.3;
    
    slide.addText(r.label, {
      x: 0.7, y, w: 3.5, h: 0.4,
      fontSize: 22, color: DARK_TEXT, bold: true, fontFace: 'Calibri'
    });
    
    // Background bar
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.7, y: y + 0.45, w: 8.5, h: 0.4,
      fill: { color: 'E0E0E0' }, rectRadius: 0.05
    });
    
    // Fill bar
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.7, y: y + 0.45, w: 8.5 * (r.pct / 100), h: 0.4,
      fill: { color: r.color }, rectRadius: 0.05
    });
    
    // Rate text
    slide.addText(r.rate, {
      x: 9.5, y: y + 0.35, w: 1.5, h: 0.6,
      fontSize: 28, color: r.color, bold: true, align: 'center', fontFace: 'Calibri'
    });
    
    slide.addText(r.desc, {
      x: 4.5, y, w: 5.0, h: 0.4,
      fontSize: 16, color: SUBTITLE_GRAY, fontFace: 'Calibri'
    });
  });
  
  slide.addText([
    { text: 'Outlook: ', options: { fontSize: 22, color: BURGUNDY, bold: true } },
    { text: 'The outlook depends on the stage, patient\'s general health, and response to treatment. Early diagnosis and treatment improve recovery chances.\n\n', options: { fontSize: 20, color: BODY_TEXT } },
    { text: 'Life Expectancy: ', options: { fontSize: 22, color: BURGUNDY, bold: true } },
    { text: 'Varies depending on cancer stage, treatment response, age, and overall health. Regular follow-up is important to detect recurrence.', options: { fontSize: 20, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 5.8, w: 11.9, h: 1.3,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });
}

// ============================================================
// SLIDE 19: MANAGEMENT AND TREATMENT (Surgical)
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'MANAGEMENT AND TREATMENT', 19, TOTAL);
  
  const surgImg = imageToBase64(getMediaPath('image14.jpeg'));
  const textW = surgImg ? 8.0 : 11.9;
  
  slide.addText([
    { text: 'Surgery is the most common treatment for localized colon cancer.\n\n', options: { fontSize: 26, color: BODY_TEXT, bold: true } },
    { text: '1. Polypectomy: ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Removal of cancerous or precancerous polyps.\n\n', options: { fontSize: 23, color: BODY_TEXT } },
    { text: '2. Colectomy: ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Removal of the affected part of the colon.\n\n', options: { fontSize: 23, color: BODY_TEXT } },
    { text: '3. Multivisceral resection: ', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Removal of the tumor and affected nearby organs.\n\n', options: { fontSize: 23, color: BODY_TEXT } },
    { text: 'Other treatments:\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: '• Chemotherapy, Radiation therapy, and Targeted therapy may be used before or after surgery.', options: { fontSize: 23, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 1.85, w: textW, h: 5.0,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });
  
  if (surgImg) {
    slide.addImage({ data: surgImg, x: 8.8, y: 2.2, w: 4.0, h: 4.0, rounding: true });
    slide.addText('Surgical Colectomy Procedure', {
      x: 8.8, y: 6.3, w: 4.0, h: 0.3,
      fontSize: 10, color: SUBTITLE_GRAY, align: 'center', italic: true, fontFace: 'Calibri'
    });
  }
}

// ============================================================
// SLIDE 20: WHEN TO SEEK MEDICAL ATTENTION
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'WHEN TO SEEK MEDICAL ATTENTION', 20, TOTAL);
  
  slide.addText([
    { text: 'Consult a healthcare provider if you have:\n\n', options: { fontSize: 26, color: BODY_TEXT } },
    { text: '• Persistent abdominal pain\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• Blood in your stool\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• Persistent changes in bowel habits\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• Unexplained weight loss\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• Persistent fatigue or anemia\n\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: 'If diagnosed with colon cancer, ask about:\n\n', options: { fontSize: 26, color: BODY_TEXT } },
    { text: '• Cancer stage\n', options: { fontSize: 24, color: BODY_TEXT } },
    { text: '• Chances of cure\n', options: { fontSize: 24, color: BODY_TEXT } },
    { text: '• Risk of recurrence', options: { fontSize: 24, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 1.85, w: 11.9, h: 5.2,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });
}

// ============================================================
// SLIDE 21: PREVENTION (Part 1 - Screening)
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PREVENTION', 21, TOTAL);
  
  slide.addText([
    { text: '1. Screening and Early Detection\n\n', options: { fontSize: 28, color: BURGUNDY, bold: true } },
    { text: '• Regular screening is one of the most powerful tools for preventing colorectal cancer.\n\n', options: { fontSize: 24, color: BODY_TEXT } },
    { text: '• It usually takes 10–15 years for polyps to turn into cancer; screening allows finding and removing them before they develop.\n\n', options: { fontSize: 24, color: BODY_TEXT } },
    { text: '• Screening should start at age 45 (or earlier with strong family history).\n\n', options: { fontSize: 24, color: BODY_TEXT } },
    { text: '2. Lifestyle and Risk Reduction\n\n', options: { fontSize: 28, color: BURGUNDY, bold: true } },
    { text: '• Maintain a healthy body weight\n', options: { fontSize: 24, color: BODY_TEXT } },
    { text: '• Regular moderate to vigorous physical activity\n', options: { fontSize: 24, color: BODY_TEXT } },
    { text: '• Eat more vegetables, fruits, and whole grains\n', options: { fontSize: 24, color: BODY_TEXT } },
    { text: '• Limit red and processed meats\n', options: { fontSize: 24, color: BODY_TEXT } },
    { text: '• Avoid alcohol and smoking', options: { fontSize: 24, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 1.85, w: 11.9, h: 5.2,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.1
  });
}

// ============================================================
// SLIDE 22: PREVENTION (Part 2 - Supplements)
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'PREVENTION (Continued)', 22, TOTAL);
  
  slide.addText([
    { text: '3. Supplements and Medical Factors\n\n', options: { fontSize: 28, color: BURGUNDY, bold: true } },
    { text: 'Vitamins and Minerals:\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Folic acid, Vitamin D, Calcium, and Magnesium may help lower risk, but research is ongoing and medical advice is recommended before supplementing.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: 'NSAIDs (e.g., Aspirin):\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'May lower risk, but taking them solely for cancer prevention is not recommended for average-risk individuals due to side effects like stomach bleeding.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: 'Hormone Therapy:\n', options: { fontSize: 24, color: BURGUNDY, bold: true } },
    { text: 'Menopausal hormone therapy may reduce risk in women, but carries other health risks and should be discussed with a doctor.', options: { fontSize: 22, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 1.85, w: 11.9, h: 5.2,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.2
  });
}

// ============================================================
// SLIDE 23: NURSING ROLE (Part 1)
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'NURSING ROLE', 23, TOTAL);
  
  const nurseImg = imageToBase64(getMediaPath('image15.jpeg'));
  const textW = nurseImg ? 8.0 : 11.9;
  
  slide.addText([
    { text: '1. Comprehensive Nursing Assessment\n\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• Symptom assessment: bowel habits, blood in stool, anemia, discomfort.\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '• Health history and nutrition: nutritional status, family history, baseline pain.\n\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '2. Pre-operative Nursing Care\n\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• Bowel preparation: laxatives, enemas, oral cleansers.\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '• Patient education and psychological support.\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '• Stoma site marking: collaborate with WOCN.\n\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '3. Post-operative Nursing Care\n\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• Monitoring recovery and vital signs.\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '• Stoma and colostomy care (pink/red, moist).\n', options: { fontSize: 21, color: BODY_TEXT } },
    { text: '• Complication prevention: pain meds, early ambulation, DVT prevention.', options: { fontSize: 21, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 1.85, w: textW, h: 5.2,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.05
  });
  
  if (nurseImg) {
    slide.addImage({ data: nurseImg, x: 9.0, y: 2.2, w: 3.8, h: 4.0, rounding: true });
    slide.addText('Clinical Nursing Assessment', {
      x: 9.0, y: 6.3, w: 3.8, h: 0.3,
      fontSize: 10, color: SUBTITLE_GRAY, align: 'center', italic: true, fontFace: 'Calibri'
    });
  }
}

// ============================================================
// SLIDE 24: NURSING ROLE (Part 2 - Side Effects & Psychosocial)
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'NURSING ROLE (Continued)', 24, TOTAL);
  
  const nurseImg2 = imageToBase64(getMediaPath('image16.jpeg'));
  const textW = nurseImg2 ? 8.5 : 11.9;
  
  slide.addText([
    { text: '4. Managing Treatment Side Effects\n\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• Chemotherapy support: managing nausea, vomiting, fatigue, diarrhea.\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '• Adequate fluid intake to prevent dehydration.\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '• Infection and lab monitoring: WBC, platelets for infection/bleeding risk.\n\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '5. Psychological and Rehabilitation Support\n\n', options: { fontSize: 26, color: BURGUNDY, bold: true } },
    { text: '• Body image and coping: help patient adapt to changes (temporary/permanent colostomy).\n', options: { fontSize: 22, color: BODY_TEXT } },
    { text: '• Holistic and family support: continuous emotional support throughout treatment and rehabilitation.', options: { fontSize: 22, color: BODY_TEXT } },
  ], {
    x: 0.7, y: 1.85, w: textW, h: 5.0,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.15
  });
  
  if (nurseImg2) {
    slide.addImage({ data: nurseImg2, x: 9.3, y: 2.5, w: 3.6, h: 3.6, rounding: true });
  }
}

// ============================================================
// SLIDE 25: NURSING PROCESS (Care Plan 1 - Fatigue)
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'NURSING PROCESS — Care Plan 1: Fatigue', 25, TOTAL);
  
  // Table
  const tableData = [
    [
      { text: 'Nursing Interventions', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 16, fontFace: 'Calibri' } },
      { text: 'Scientific Rationales', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 16, fontFace: 'Calibri' } }
    ],
    [
      { text: '1. Assess fatigue severity using 0–10 scale.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } },
      { text: 'Establishes baseline, identifies peak energy periods and triggers.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } }
    ],
    [
      { text: '2. Review CBC results (Hb, Hct); report abnormal findings.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } },
      { text: 'Identifies anemia contributing to tissue hypoxia and fatigue.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } }
    ],
    [
      { text: '3. Cluster nursing care activities; schedule rest periods.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } },
      { text: 'Minimizes energy expenditure, reduces sleep fragmentation.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } }
    ],
    [
      { text: '4. Help patient prioritize tasks; balance activity with rest.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } },
      { text: 'Maintains muscle strength, prevents deconditioning.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } }
    ],
    [
      { text: '5. Support optimal nutrition/hydration (protein, iron-rich foods).', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } },
      { text: 'Addresses nutritional contributors to weakness.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } }
    ]
  ];
  
  slide.addTable(tableData, {
    x: 0.5, y: 1.85, w: 12.33,
    colW: [6.17, 6.17],
    border: { type: 'solid', color: 'CCCCCC', pt: 1 },
    rowH: [0.5, 0.75, 0.75, 0.75, 0.75, 0.75],
    autoPage: false
  });
}

// ============================================================
// SLIDE 26: NURSING PROCESS (Care Plan 2 - Abdominal Pain)
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'NURSING PROCESS — Care Plan 2: Abdominal Pain', 26, TOTAL);
  
  const tableData = [
    [
      { text: 'Nursing Interventions', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 16, fontFace: 'Calibri' } },
      { text: 'Scientific Rationales', options: { bold: true, color: WHITE, fill: { color: BURGUNDY }, fontSize: 16, fontFace: 'Calibri' } }
    ],
    [
      { text: '1. Assess pain using consistent rating scale before and after every intervention.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } },
      { text: 'Provides reliable baseline; evaluates analgesic effectiveness.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } }
    ],
    [
      { text: '2. Observe for acute surgical red flags (severe pain, rigidity, vomiting, distension).', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } },
      { text: 'May indicate bowel obstruction, perforation, or anastomotic leak.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } }
    ],
    [
      { text: '3. Administer prescribed analgesia; monitor therapeutic effect and adverse reactions.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } },
      { text: 'Targeted pain relief while monitoring for respiratory depression or opioid-induced ileus.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } }
    ],
    [
      { text: '4. Implement non-pharmacological comfort measures (semi-Fowler\'s, relaxation, breathing).', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } },
      { text: 'Reduces muscular tension and anxiety, complementing pharmacological analgesia.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } }
    ],
    [
      { text: '5. Offer diet/fluids per prescribed clinical status (NPO vs. ordered diet).', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } },
      { text: 'Diet advancement depends on obstruction risk and peristaltic recovery.', options: { fontSize: 15, fontFace: 'Calibri', color: BODY_TEXT } }
    ]
  ];
  
  slide.addTable(tableData, {
    x: 0.5, y: 1.85, w: 12.33,
    colW: [6.17, 6.17],
    border: { type: 'solid', color: 'CCCCCC', pt: 1 },
    rowH: [0.5, 0.8, 0.8, 0.8, 0.8, 0.8],
    autoPage: false
  });
}

// ============================================================
// SLIDE 27: SUMMARY
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'SUMMARY', 27, TOTAL);
  
  slide.addText([
    { text: 'To sum up, colon cancer is a preventable disease through early screening.\n\n', options: { fontSize: 26, color: BODY_TEXT } },
    { text: 'As nurses, our role is crucial in supporting the patient before and after surgery.\n\n', options: { fontSize: 26, color: BODY_TEXT } },
    { text: 'By managing actual symptoms like pain and fatigue, and preventing potential complications like bowel obstruction, we can safely improve the patient\'s outcomes and quality of life.\n\n', options: { fontSize: 26, color: BODY_TEXT } },
    { text: 'Thank you.', options: { fontSize: 32, color: BURGUNDY, bold: true } },
  ], {
    x: 0.7, y: 2.0, w: 11.9, h: 4.5,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.3
  });
}

// ============================================================
// SLIDE 28: REFERENCES
// ============================================================
{
  const slide = pptx.addSlide();
  addHeader(slide, 'REFERENCES', 28, TOTAL);
  
  const refs = [
    'American Cancer Society. (n.d.-a). Colorectal cancer prevention guidelines.\nhttps://www.cancer.org/health-care-professionals/american-cancer-society-prevention-early-detection-guidelines/colorectal-cancer-screening-guidelines.html',
    'American Cancer Society. (n.d.-b). What is colorectal cancer?\nhttps://www.cancer.org/cancer/types/colon-rectal-cancer/about/what-is-colorectal-cancer.html',
    'Colon Cancer Coalition. (n.d.). Signs and symptoms of colorectal cancer.\nhttps://coloncancercoalition.org',
    'Mayo Clinic. (n.d.). Mayo Clinic. https://www.mayoclinic.org',
    'National Cancer Institute. (n.d.-a). Colon cancer treatment (PDQ®) – Patient version.\nhttps://www.cancer.gov/types/colorectal/patient/colon-treatment-pdq',
    'National Cancer Institute. (n.d.-b). Screening tests to detect colorectal cancer and polyps.\nhttps://www.cancer.gov/types/colorectal/screening-fact-sheet',
    'Roshandel, G., Ghasemi-Kebria, F., & Malekzadeh, R. (2024). Colorectal cancer: Epidemiology, risk factors, and prevention. Cancers, 16(8), 1530.',
    'SEER Program. (2025). Anatomy. National Cancer Institute.\nhttps://training.seer.cancer.gov/colorectal/anatomy/'
  ];
  
  const textRuns = refs.map((ref, i) => ({
    text: `${i + 1}. ${ref}\n\n`,
    options: { fontSize: 14, color: BODY_TEXT }
  }));
  
  slide.addText(textRuns, {
    x: 0.7, y: 1.85, w: 11.9, h: 5.2,
    fontFace: 'Calibri', valign: 'top', lineSpacingMultiple: 1.1
  });
}

// ============================================================
// GENERATE
// ============================================================
const outputPath = path.join(__dirname, '..', 'public', 'colon-cancer-presentation.pptx');

pptx.writeFile({ fileName: outputPath })
  .then(() => {
    const stats = fs.statSync(outputPath);
    console.log('SUCCESS! PPTX generated:');
    console.log('  Path:', outputPath);
    console.log('  Size:', (stats.size / 1024 / 1024).toFixed(2), 'MB');
    console.log('  Slides:', TOTAL);
  })
  .catch(err => {
    console.error('ERROR:', err);
    process.exit(1);
  });
