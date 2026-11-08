import pptxgen from 'pptxgenjs';
import { slides, SECTIONS, ANATOMY_SECTIONS, PATHO_STAGES, DIAGNOSTIC_TESTS } from '../content/slides';
import { CARE_PLANS } from '../content/carePlans';

// Institutional Colors
const BURGUNDY = '7A1F3D';
const MAUVE = '9E4767';
const INK = '2B1F24';
const GRAY = '6E5E65';
const LIGHT_BG = 'FDF6F7';
const CARD_BG = 'FFFFFF';
const BORDER_COLOR = 'F0D5DC';
const ACCENT_PILL = 'F8E3E8';
const GOLD = 'C89D42';
const DARK_BURGUNDY = '5A122A';

// Typography
const FONT_SERIF = 'Georgia';
const FONT_SANS = 'Calibri';

// Asset base paths
const LOGO_NURSING = 'public/images/nursing-logo-opt.png';
const LOGO_DEPT = 'public/images/dept-logo-opt.png';
const LOGO_AZHAR = 'public/images/azhar-logo-opt.png';

const INFOGRAPHIC_DIAG = 'public/images/diag-stages-infographic.jpg';
const INFOGRAPHIC_PROG = 'public/images/prognosis-infographic.jpg';
const INFOGRAPHIC_TX = 'public/images/treatment-infographic.jpg';

const SCENE_DIGESTIVE = 'public/images/scenes/3d-digestive.png';
const SCENE_ANATOMY = 'public/images/scenes/3d-colon-anatomy.png';
const SCENE_SECTIONS = 'public/images/scenes/3d-colon-sections.png';
const SCENE_PATHO = 'public/images/scenes/3d-patho-lattice.png';
const SCENE_TUMOR = 'public/images/scenes/3d-tumor-model.png';

export function createPresentation(): pptxgen {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9'; // 10.0 x 5.625 inches
  pptx.author = 'Faculty of Nursing for Girls, Al-Azhar University';
  pptx.company = 'Medical-Surgical Nursing Department';
  pptx.title = 'Colon Cancer - Clinical Nursing Seminar';
  pptx.subject = 'Colorectal Oncology & Evidence-Based Nursing Care Plans';

  // Helper to resolve image path for browser vs node
  const getImg = (relPath: string) => {
    if (typeof window !== 'undefined') {
      const base = (import.meta as any).env?.BASE_URL || '/';
      return `${base}${relPath.replace(/^public\//, '')}`;
    }
    return relPath;
  };

  // Helper: Adds standard academic institutional header (Slides 2-46)
  const addInstitutionalHeader = (slide: pptxgen.Slide, sectionTitle?: string) => {
    // Header background bar
    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: 10,
      h: 0.64,
      fill: { color: 'FFFFFF' },
      line: { color: BORDER_COLOR, width: 0.8 },
    });

    // Left: Faculty of Nursing Logo & Names
    slide.addImage({
      path: getImg(LOGO_NURSING),
      x: 0.25,
      y: 0.08,
      w: 0.48,
      h: 0.48,
    });
    slide.addText(
      [
        { text: 'Faculty of Nursing for Girls, Al-Azhar University\n', options: { fontSize: 7.5, bold: true, color: BURGUNDY } },
        { text: 'كُليّة التمريض للبنات بالقاهرة', options: { fontSize: 6.5, bold: false, color: MAUVE } },
      ],
      { x: 0.8, y: 0.1, w: 2.8, h: 0.45, fontFace: FONT_SANS, margin: 0, align: 'left' }
    );

    // Center: Medical-Surgical Nursing Dept Logo & Names
    slide.addImage({
      path: getImg(LOGO_DEPT),
      x: 3.85,
      y: 0.08,
      w: 0.48,
      h: 0.48,
    });
    slide.addText(
      [
        { text: 'Medical-Surgical Nursing Department\n', options: { fontSize: 7.5, bold: true, color: BURGUNDY } },
        { text: 'قسم التمريض الباطني والجراحي', options: { fontSize: 6.5, bold: false, color: MAUVE } },
      ],
      { x: 4.38, y: 0.1, w: 2.7, h: 0.45, fontFace: FONT_SANS, margin: 0, align: 'left' }
    );

    // Right: Al-Azhar University Logo & Names
    slide.addText(
      [
        { text: 'Al-Azhar University\n', options: { fontSize: 7.5, bold: true, color: BURGUNDY } },
        { text: 'جامعة الأزهر الشريف', options: { fontSize: 6.5, bold: false, color: MAUVE } },
      ],
      { x: 7.25, y: 0.1, w: 1.95, h: 0.45, fontFace: FONT_SANS, margin: 0, align: 'right' }
    );
    slide.addImage({
      path: getImg(LOGO_AZHAR),
      x: 9.25,
      y: 0.08,
      w: 0.48,
      h: 0.48,
    });

    // Optional top topic pill
    if (sectionTitle) {
      slide.addText(sectionTitle.toUpperCase(), {
        x: 0.6,
        y: 0.74,
        w: 4.5,
        h: 0.25,
        fontSize: 8,
        bold: true,
        color: MAUVE,
        fontFace: FONT_SANS,
        charSpacing: 1.5,
      });
    }
  };

  // Helper: Slide Title
  const addSlideTitle = (slide: pptxgen.Slide, title: string, yPos: number = 0.95) => {
    slide.addText(title, {
      x: 0.6,
      y: yPos,
      w: 8.8,
      h: 0.5,
      fontSize: 20,
      bold: true,
      color: BURGUNDY,
      fontFace: FONT_SERIF,
    });
  };

  // -------------------------------------------------------------
  // SLIDE 1: COVER
  // -------------------------------------------------------------
  const slide1 = pptx.addSlide();
  slide1.background = { color: LIGHT_BG };

  // Institutional Header on Cover
  slide1.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 10,
    h: 0.75,
    fill: { color: 'FFFFFF' },
    line: { color: BORDER_COLOR, width: 1 },
  });
  slide1.addImage({ path: getImg(LOGO_NURSING), x: 0.3, y: 0.1, w: 0.55, h: 0.55 });
  slide1.addText(
    [
      { text: 'Faculty of Nursing for Girls, Al-Azhar University\n', options: { fontSize: 8.5, bold: true, color: BURGUNDY } },
      { text: 'كُليّة التمريض للبنات بالقاهرة', options: { fontSize: 7.5, color: MAUVE } },
    ],
    { x: 0.95, y: 0.12, w: 3.2, h: 0.5, fontFace: FONT_SANS, margin: 0 }
  );

  slide1.addImage({ path: getImg(LOGO_DEPT), x: 4.35, y: 0.1, w: 0.55, h: 0.55 });
  slide1.addText(
    [
      { text: 'Medical-Surgical Nursing Department\n', options: { fontSize: 8.5, bold: true, color: BURGUNDY } },
      { text: 'قسم التمريض الباطني والجراحي', options: { fontSize: 7.5, color: MAUVE } },
    ],
    { x: 4.95, y: 0.12, w: 2.7, h: 0.5, fontFace: FONT_SANS, margin: 0 }
  );

  slide1.addText(
    [
      { text: 'Al-Azhar University\n', options: { fontSize: 8.5, bold: true, color: BURGUNDY } },
      { text: 'جامعة الأزهر الشريف', options: { fontSize: 7.5, color: MAUVE } },
    ],
    { x: 7.5, y: 0.12, w: 1.65, h: 0.5, fontFace: FONT_SANS, margin: 0, align: 'right' }
  );
  slide1.addImage({ path: getImg(LOGO_AZHAR), x: 9.2, y: 0.1, w: 0.55, h: 0.55 });

  // Main Cover Title & Subtitles
  slide1.addText('COLON CANCER', {
    x: 0.6,
    y: 1.15,
    w: 5.5,
    h: 0.9,
    fontSize: 40,
    bold: true,
    color: BURGUNDY,
    fontFace: FONT_SERIF,
  });

  slide1.addText('Nursing College Presentation', {
    x: 0.6,
    y: 2.05,
    w: 5.5,
    h: 0.4,
    fontSize: 16,
    bold: true,
    color: MAUVE,
    fontFace: FONT_SERIF,
  });

  slide1.addText(
    'A Comprehensive Clinical Seminar on Colorectal Oncology, Surgical Interventions & Advanced Nursing Care Plans.',
    {
      x: 0.6,
      y: 2.45,
      w: 5.3,
      h: 0.7,
      fontSize: 11,
      color: GRAY,
      fontFace: FONT_SANS,
      lineSpacing: 14,
    }
  );

  // Under Supervision Card
  slide1.addShape(pptx.ShapeType.roundRect, {
    x: 0.6,
    y: 3.25,
    w: 5.3,
    h: 1.9,
    rectRadius: 0.1,
    fill: { color: 'FFFFFF' },
    line: { color: BORDER_COLOR, width: 1 },
  });

  slide1.addText('UNDER SUPERVISION', {
    x: 0.8,
    y: 3.35,
    w: 4.9,
    h: 0.25,
    fontSize: 8,
    bold: true,
    color: BURGUNDY,
    fontFace: FONT_SANS,
    charSpacing: 1.5,
  });

  slide1.addText(
    [
      { text: 'Prof. Dr. Heba Abdel-Azem\n', options: { fontSize: 11, bold: true, color: BURGUNDY } },
      { text: 'Professor of Medical Surgical Nursing Department\n\n', options: { fontSize: 9, color: MAUVE } },
      { text: 'Dr. Eman Abel-Maqsoud\n', options: { fontSize: 11, bold: true, color: BURGUNDY } },
      { text: 'Demonstrator of Medical Surgical Nursing Department', options: { fontSize: 9, color: MAUVE } },
    ],
    { x: 0.8, y: 3.65, w: 4.9, h: 1.1, fontFace: FONT_SANS, margin: 0 }
  );

  slide1.addText('Academic Year 2026 / 2027', {
    x: 3.7,
    y: 4.8,
    w: 2.0,
    h: 0.25,
    fontSize: 8,
    bold: true,
    color: DARK_BURGUNDY,
    fontFace: FONT_SANS,
    align: 'right',
  });

  // Right Side: 3D Digestive Oncology Model Visual
  slide1.addShape(pptx.ShapeType.roundRect, {
    x: 6.2,
    y: 1.15,
    w: 3.4,
    h: 4.0,
    rectRadius: 0.12,
    fill: { color: 'FFFFFF' },
    line: { color: BORDER_COLOR, width: 1 },
  });

  slide1.addImage({
    path: getImg(SCENE_DIGESTIVE),
    x: 6.35,
    y: 1.3,
    w: 3.1,
    h: 3.4,
  });

  slide1.addText('3D Digestive Oncology Model', {
    x: 6.2,
    y: 4.85,
    w: 3.4,
    h: 0.25,
    fontSize: 8,
    color: GRAY,
    fontFace: FONT_SANS,
    align: 'center',
    italic: true,
  });

  // -------------------------------------------------------------
  // SLIDE 2: PRESENTATION TEAM
  // -------------------------------------------------------------
  const slide2 = pptx.addSlide();
  slide2.background = { color: LIGHT_BG };
  addInstitutionalHeader(slide2, 'Presentation Team');
  addSlideTitle(slide2, 'Presentation Team');

  slide2.addText('إعداد طالبات كلية التمريض • الفرقة الدراسية', {
    x: 0.6,
    y: 1.45,
    w: 8.8,
    h: 0.3,
    fontSize: 10,
    bold: true,
    color: BURGUNDY,
    fontFace: FONT_SANS,
    align: 'center',
  });

  slide2.addText('Faculty of Nursing for Girls • Al-Azhar University • Medical-Surgical Nursing Department', {
    x: 0.6,
    y: 1.75,
    w: 8.8,
    h: 0.25,
    fontSize: 9,
    color: GRAY,
    fontFace: FONT_SANS,
    align: 'center',
  });

  const students = [
    { num: 1, name: 'Hager Hany' },
    { num: 2, name: 'Hager Saeed' },
    { num: 3, name: 'Hager Mostafa' },
    { num: 4, name: 'Haidy Mahmoud' },
    { num: 5, name: 'Norhan Badawy' },
    { num: 6, name: 'Hager Ahmed' },
  ];

  const colW = 2.8;
  const cardH = 0.95;
  const gapX = 0.2;
  const startX = 0.6;

  students.forEach((s, idx) => {
    const row = Math.floor(idx / 3);
    const col = idx % 3;
    const x = startX + col * (colW + gapX);
    const y = 2.15 + row * (cardH + 0.2);

    slide2.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w: colW,
      h: cardH,
      rectRadius: 0.08,
      fill: { color: 'FFFFFF' },
      line: { color: BORDER_COLOR, width: 1 },
    });

    // Badge circle
    slide2.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.15,
      y: y + 0.22,
      w: 0.5,
      h: 0.5,
      fill: { color: ACCENT_PILL },
      line: { color: BORDER_COLOR, width: 0.8 },
    });
    slide2.addText(String(s.num), {
      x: x + 0.15,
      y: y + 0.22,
      w: 0.5,
      h: 0.5,
      fontSize: 12,
      bold: true,
      color: BURGUNDY,
      fontFace: FONT_SANS,
      align: 'center',
      valign: 'middle',
    });

    slide2.addText(
      [
        { text: `${s.name}\n`, options: { fontSize: 13, bold: true, color: BURGUNDY } },
        { text: 'Nursing Student • Al-Azhar University', options: { fontSize: 8.5, color: GRAY } },
      ],
      { x: x + 0.75, y: y + 0.18, w: colW - 0.85, h: 0.6, fontFace: FONT_SANS, margin: 0 }
    );
  });

  slide2.addShape(pptx.ShapeType.roundRect, {
    x: 0.6,
    y: 4.6,
    w: 8.8,
    h: 0.55,
    rectRadius: 0.08,
    fill: { color: 'FFFFFF' },
    line: { color: BORDER_COLOR, width: 0.8 },
  });
  slide2.addText(
    'Clinical seminar prepared under the academic curriculum of the Medical-Surgical Nursing Department, Faculty of Nursing for Girls, Al-Azhar University.',
    {
      x: 0.8,
      y: 4.65,
      w: 8.4,
      h: 0.45,
      fontSize: 9.5,
      color: BURGUNDY,
      fontFace: FONT_SANS,
      align: 'center',
    }
  );

  // -------------------------------------------------------------
  // SLIDE 3: GOAL
  // -------------------------------------------------------------
  const slide3 = pptx.addSlide();
  slide3.background = { color: LIGHT_BG };
  addInstitutionalHeader(slide3, 'Educational Framework');
  addSlideTitle(slide3, 'GOAL');

  // Main Goal Card
  slide3.addShape(pptx.ShapeType.roundRect, {
    x: 0.6,
    y: 1.5,
    w: 5.6,
    h: 2.2,
    rectRadius: 0.1,
    fill: { color: 'FFFFFF' },
    line: { color: BORDER_COLOR, width: 1 },
  });
  slide3.addText(
    'The goal of this assignment is to provide nursing students with a comprehensive understanding of colon cancer, including its anatomy, pathophysiology, risk factors, clinical manifestations, diagnosis, and management, in order to enhance their ability to deliver safe, evidence-based nursing care to patients with colon cancer.',
    {
      x: 0.85,
      y: 1.65,
      w: 5.1,
      h: 1.9,
      fontSize: 12.5,
      color: INK,
      fontFace: FONT_SERIF,
      lineSpacing: 18,
      italic: true,
    }
  );

  // 4 Competency Pillars
  const pillars = [
    'Anatomical & Pathophysiological Mastery',
    'Evidence-Based Diagnostic Assessment',
    'Multimodal Oncologic Management',
    'Patient-Centered Nursing Process & Care Plans',
  ];
  pillars.forEach((p, idx) => {
    const y = 3.85 + idx * 0.35;
    slide3.addShape(pptx.ShapeType.roundRect, {
      x: 0.6,
      y,
      w: 5.6,
      h: 0.28,
      rectRadius: 0.05,
      fill: { color: ACCENT_PILL },
      line: { color: BORDER_COLOR, width: 0.6 },
    });
    slide3.addText(`✓  ${p}`, {
      x: 0.75,
      y,
      w: 5.3,
      h: 0.28,
      fontSize: 9,
      bold: true,
      color: BURGUNDY,
      fontFace: FONT_SANS,
      valign: 'middle',
    });
  });

  // Right Card: Institutional Clinical Focus Box
  slide3.addShape(pptx.ShapeType.roundRect, {
    x: 6.4,
    y: 1.5,
    w: 3.0,
    h: 3.65,
    rectRadius: 0.1,
    fill: { color: 'FFFFFF' },
    line: { color: BORDER_COLOR, width: 1 },
  });
  slide3.addImage({ path: getImg(LOGO_DEPT), x: 7.15, y: 1.8, w: 1.5, h: 1.5 });
  slide3.addText(
    [
      { text: 'Medical-Surgical Nursing\n', options: { fontSize: 12, bold: true, color: BURGUNDY } },
      { text: 'Clinical Competency Standards\n\n', options: { fontSize: 9.5, color: MAUVE } },
      { text: 'Delivering holistic, evidence-based care tailored to oncology patients undergoing colorectal surgical and medical interventions.', options: { fontSize: 8.5, color: GRAY } },
    ],
    { x: 6.6, y: 3.45, w: 2.6, h: 1.5, fontFace: FONT_SANS, align: 'center' }
  );

  // -------------------------------------------------------------
  // SLIDE 4: OBJECTIVES
  // -------------------------------------------------------------
  const slide4 = pptx.addSlide();
  slide4.background = { color: LIGHT_BG };
  addInstitutionalHeader(slide4, 'Educational Framework');
  addSlideTitle(slide4, 'OBJECTIVES');

  slide4.addText('By the end of this assignment, the student will be able to:', {
    x: 0.6,
    y: 1.42,
    w: 8.8,
    h: 0.25,
    fontSize: 11,
    bold: true,
    color: MAUVE,
    fontFace: FONT_SANS,
  });

  const objectives = [
    'Define colon cancer and describe the anatomy and physiology of the colon.',
    'Explain the pathophysiology, causes, and risk factors of colon cancer.',
    'Identify the types, signs and symptoms, and complications of colon cancer.',
    'Discuss the diagnostic tests, staging, and prognosis of colon cancer.',
    'Outline the medical and surgical management and the preventive measures of colon cancer.',
    'Describe the role of the nurse in caring for patients with colon cancer before and after surgery.',
    'Apply the nursing process to a patient with colon cancer.',
  ];

  // 2-column layout (items 0-2 left, items 3-5 right, item 6 full width bottom)
  const objW = 4.3;
  objectives.forEach((obj, idx) => {
    if (idx < 6) {
      const col = idx % 2;
      const row = Math.floor(idx / 2);
      const x = 0.6 + col * (objW + 0.2);
      const y = 1.75 + row * 0.95;

      slide4.addShape(pptx.ShapeType.roundRect, {
        x,
        y,
        w: objW,
        h: 0.85,
        rectRadius: 0.08,
        fill: { color: 'FFFFFF' },
        line: { color: BORDER_COLOR, width: 0.8 },
      });

      slide4.addShape(pptx.ShapeType.ellipse, {
        x: x + 0.15,
        y: y + 0.18,
        w: 0.48,
        h: 0.48,
        fill: { color: ACCENT_PILL },
        line: { color: BORDER_COLOR, width: 0.6 },
      });
      slide4.addText(String(idx + 1), {
        x: x + 0.15,
        y: y + 0.18,
        w: 0.48,
        h: 0.48,
        fontSize: 11,
        bold: true,
        color: BURGUNDY,
        fontFace: FONT_SANS,
        align: 'center',
        valign: 'middle',
      });

      slide4.addText(obj, {
        x: x + 0.72,
        y: y + 0.12,
        w: objW - 0.85,
        h: 0.65,
        fontSize: 9.5,
        color: INK,
        fontFace: FONT_SANS,
        valign: 'middle',
      });
    } else {
      // 7th objective spans full width
      const y = 4.65;
      slide4.addShape(pptx.ShapeType.roundRect, {
        x: 0.6,
        y,
        w: 8.8,
        h: 0.6,
        rectRadius: 0.08,
        fill: { color: 'FFFFFF' },
        line: { color: BURGUNDY, width: 1 },
      });
      slide4.addShape(pptx.ShapeType.ellipse, {
        x: 0.75,
        y: y + 0.1,
        w: 0.4,
        h: 0.4,
        fill: { color: BURGUNDY },
      });
      slide4.addText('7', {
        x: 0.75,
        y: y + 0.1,
        w: 0.4,
        h: 0.4,
        fontSize: 11,
        bold: true,
        color: 'FFFFFF',
        fontFace: FONT_SANS,
        align: 'center',
        valign: 'middle',
      });
      slide4.addText(obj, {
        x: 1.25,
        y: y + 0.08,
        w: 8.0,
        h: 0.44,
        fontSize: 10,
        bold: true,
        color: BURGUNDY,
        fontFace: FONT_SANS,
        valign: 'middle',
      });
    }
  });

  // -------------------------------------------------------------
  // SLIDE 5: OUTLINE (21 SECTIONS)
  // -------------------------------------------------------------
  const slide5 = pptx.addSlide();
  slide5.background = { color: LIGHT_BG };
  addInstitutionalHeader(slide5, 'Agenda');
  addSlideTitle(slide5, 'Presentation Outlines');

  const outlineColW = 2.8;
  const outlineRowH = 0.45;
  SECTIONS.forEach((sec, idx) => {
    const col = Math.floor(idx / 7);
    const row = idx % 7;
    const x = 0.6 + col * (outlineColW + 0.2);
    const y = 1.5 + row * (outlineRowH + 0.08);

    slide5.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w: outlineColW,
      h: outlineRowH,
      rectRadius: 0.05,
      fill: { color: 'FFFFFF' },
      line: { color: BORDER_COLOR, width: 0.6 },
    });

    slide5.addText(
      [
        { text: `${String(idx + 1).padStart(2, '0')}  `, options: { fontSize: 8.5, bold: true, color: MAUVE } },
        { text: sec.label, options: { fontSize: 9.5, bold: true, color: BURGUNDY } },
      ],
      { x: x + 0.15, y: y + 0.05, w: outlineColW - 0.25, h: outlineRowH - 0.1, fontFace: FONT_SANS, valign: 'middle' }
    );
  });

  // -------------------------------------------------------------
  // SLIDES 6 TO 46: SYSTEMATIC CLINICAL SLIDES
  // -------------------------------------------------------------
  for (let i = 5; i < slides.length; i++) {
    const s = slides[i];
    const pptxSlide = pptx.addSlide();
    pptxSlide.background = { color: LIGHT_BG };

    const sectionObj = SECTIONS.find((sec) => sec.id === s.section);
    addInstitutionalHeader(pptxSlide, sectionObj?.label || 'Clinical Seminar');
    addSlideTitle(pptxSlide, s.title);

    // Render based on Slide Kind
    switch (s.kind) {
      case 'intro': {
        if (s.points) {
          // Left: bullet cards
          s.points.forEach((pt, pIdx) => {
            const y = 1.6 + pIdx * 1.0;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y,
              w: 5.5,
              h: 0.85,
              rectRadius: 0.08,
              fill: { color: 'FFFFFF' },
              line: { color: BORDER_COLOR, width: 0.8 },
            });
            pptxSlide.addText(pt, {
              x: 0.85,
              y: y + 0.1,
              w: 5.0,
              h: 0.65,
              fontSize: 12,
              color: INK,
              fontFace: FONT_SANS,
              bullet: true,
              valign: 'middle',
            });
          });

          // Right: Digestive model visual
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 6.4,
            y: 1.6,
            w: 3.0,
            h: 3.2,
            rectRadius: 0.1,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });
          pptxSlide.addImage({ path: getImg(SCENE_DIGESTIVE), x: 6.55, y: 1.75, w: 2.7, h: 2.7 });
          pptxSlide.addText('Digestive System & Colonic Lining', {
            x: 6.4,
            y: 4.45,
            w: 3.0,
            h: 0.25,
            fontSize: 8.5,
            color: GRAY,
            align: 'center',
            italic: true,
          });
        } else if (s.chips) {
          // Intro: Risk factors chips
          pptxSlide.addText(s.chipsLabel || 'Risk Factors', {
            x: 0.6,
            y: 1.5,
            w: 8.8,
            h: 0.35,
            fontSize: 14,
            bold: true,
            color: BURGUNDY,
            fontFace: FONT_SERIF,
          });

          // Chips in 2x3 grid
          s.chips.forEach((chip, cIdx) => {
            const col = cIdx % 3;
            const row = Math.floor(cIdx / 3);
            const x = 0.6 + col * 2.95;
            const y = 2.0 + row * 0.85;

            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x,
              y,
              w: 2.8,
              h: 0.7,
              rectRadius: 0.08,
              fill: { color: 'FFFFFF' },
              line: { color: BORDER_COLOR, width: 0.8 },
            });
            pptxSlide.addText(`•  ${chip}`, {
              x: x + 0.15,
              y,
              w: 2.5,
              h: 0.7,
              fontSize: 11,
              bold: true,
              color: BURGUNDY,
              fontFace: FONT_SANS,
              valign: 'middle',
            });
          });

          if (s.statement) {
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y: 3.9,
              w: 8.8,
              h: 0.9,
              rectRadius: 0.1,
              fill: { color: ACCENT_PILL },
              line: { color: BURGUNDY, width: 1 },
            });
            pptxSlide.addText(s.statement, {
              x: 0.9,
              y: 4.0,
              w: 8.2,
              h: 0.7,
              fontSize: 14,
              bold: true,
              color: BURGUNDY,
              fontFace: FONT_SERIF,
              align: 'center',
              valign: 'middle',
            });
          }
        }
        break;
      }

      case 'definition': {
        s.points.forEach((pt, pIdx) => {
          const y = 1.6 + pIdx * 1.1;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 8.8,
            h: 0.95,
            rectRadius: 0.08,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });

          pptxSlide.addShape(pptx.ShapeType.ellipse, {
            x: 0.85,
            y: y + 0.22,
            w: 0.5,
            h: 0.5,
            fill: { color: ACCENT_PILL },
            line: { color: BORDER_COLOR, width: 0.6 },
          });
          pptxSlide.addText(String(pIdx + 1), {
            x: 0.85,
            y: y + 0.22,
            w: 0.5,
            h: 0.5,
            fontSize: 12,
            bold: true,
            color: BURGUNDY,
            fontFace: FONT_SANS,
            align: 'center',
            valign: 'middle',
          });

          pptxSlide.addText(pt, {
            x: 1.55,
            y: y + 0.12,
            w: 7.6,
            h: 0.7,
            fontSize: 13,
            color: INK,
            fontFace: FONT_SANS,
            valign: 'middle',
          });
        });
        break;
      }

      case 'anatomy': {
        if (s.mode === 'facts') {
          // Anatomy Facts + 3D Colon Anatomy Model
          s.facts?.forEach((fact, fIdx) => {
            const y = 1.6 + fIdx * 0.9;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y,
              w: 5.2,
              h: 0.78,
              rectRadius: 0.08,
              fill: { color: 'FFFFFF' },
              line: { color: BORDER_COLOR, width: 0.8 },
            });
            pptxSlide.addText(fact, {
              x: 0.8,
              y: y + 0.08,
              w: 4.8,
              h: 0.62,
              fontSize: 11,
              color: INK,
              fontFace: FONT_SANS,
              bullet: true,
              valign: 'middle',
            });
          });

          if (s.statement) {
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y: 4.4,
              w: 5.2,
              h: 0.6,
              rectRadius: 0.08,
              fill: { color: ACCENT_PILL },
              line: { color: BURGUNDY, width: 0.8 },
            });
            pptxSlide.addText(s.statement, {
              x: 0.8,
              y: 4.4,
              w: 4.8,
              h: 0.6,
              fontSize: 12,
              bold: true,
              color: BURGUNDY,
              fontFace: FONT_SERIF,
              valign: 'middle',
            });
          }

          // 3D Colon Model Visual
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 6.1,
            y: 1.6,
            w: 3.3,
            h: 3.4,
            rectRadius: 0.1,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });
          pptxSlide.addImage({ path: getImg(SCENE_ANATOMY), x: 6.25, y: 1.75, w: 3.0, h: 2.8 });
          pptxSlide.addText('3D Colon Anatomy Model', {
            x: 6.1,
            y: 4.65,
            w: 3.3,
            h: 0.25,
            fontSize: 8.5,
            color: GRAY,
            align: 'center',
            italic: true,
          });
        } else {
          // 4 Sections of Colon + 3D Sections visual
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y: 1.5,
            w: 4.3,
            h: 3.6,
            rectRadius: 0.1,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });
          pptxSlide.addImage({ path: getImg(SCENE_SECTIONS), x: 0.75, y: 1.65, w: 4.0, h: 3.1 });
          pptxSlide.addText('Interactive 3D Anatomical Segmentation', {
            x: 0.6,
            y: 4.8,
            w: 4.3,
            h: 0.25,
            fontSize: 8.5,
            color: GRAY,
            align: 'center',
            italic: true,
          });

          // 4 Sections Cards
          ANATOMY_SECTIONS.forEach((sec, idx) => {
            const y = 1.5 + idx * 0.88;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 5.2,
              y,
              w: 4.2,
              h: 0.78,
              rectRadius: 0.08,
              fill: { color: 'FFFFFF' },
              line: { color: BORDER_COLOR, width: 0.8 },
            });
            pptxSlide.addText(
              [
                { text: `${sec.label}\n`, options: { fontSize: 11, bold: true, color: BURGUNDY } },
                { text: sec.text, options: { fontSize: 9, color: GRAY } },
              ],
              { x: 5.4, y: y + 0.1, w: 3.8, h: 0.58, fontFace: FONT_SANS, margin: 0 }
            );
          });
        }
        break;
      }

      case 'patho': {
        if (s.mode === 'genes') {
          // Pathophysiology: Genetic mutations
          slide1.background = { color: LIGHT_BG };
          pptxSlide.addText('Key Genetic Drivers:', {
            x: 0.6,
            y: 1.5,
            w: 4.5,
            h: 0.3,
            fontSize: 12,
            bold: true,
            color: BURGUNDY,
            fontFace: FONT_SERIF,
          });

          s.genes?.forEach((g, gIdx) => {
            const x = 0.6 + gIdx * 1.6;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x,
              y: 1.85,
              w: 1.45,
              h: 0.45,
              rectRadius: 0.05,
              fill: { color: ACCENT_PILL },
              line: { color: BURGUNDY, width: 1 },
            });
            pptxSlide.addText(g, {
              x,
              y: 1.85,
              w: 1.45,
              h: 0.45,
              fontSize: 11,
              bold: true,
              color: BURGUNDY,
              fontFace: FONT_SANS,
              align: 'center',
              valign: 'middle',
            });
          });

          s.points?.forEach((pt, pIdx) => {
            const y = 2.5 + pIdx * 0.85;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y,
              w: 8.8,
              h: 0.72,
              rectRadius: 0.08,
              fill: { color: 'FFFFFF' },
              line: { color: BORDER_COLOR, width: 0.8 },
            });
            pptxSlide.addText(pt, {
              x: 0.85,
              y: y + 0.08,
              w: 8.3,
              h: 0.56,
              fontSize: 12,
              color: INK,
              fontFace: FONT_SANS,
              bullet: true,
              valign: 'middle',
            });
          });
        } else {
          // Progression steps + 3D cell lattice
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y: 1.5,
            w: 4.4,
            h: 3.6,
            rectRadius: 0.1,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });
          pptxSlide.addImage({ path: getImg(SCENE_PATHO), x: 0.75, y: 1.65, w: 4.1, h: 3.1 });
          pptxSlide.addText('3D Cellular Lattice Progression', {
            x: 0.6,
            y: 4.8,
            w: 4.4,
            h: 0.25,
            fontSize: 8.5,
            color: GRAY,
            align: 'center',
            italic: true,
          });

          PATHO_STAGES.forEach((stage, idx) => {
            const y = 1.5 + idx * 0.88;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 5.3,
              y,
              w: 4.1,
              h: 0.75,
              rectRadius: 0.08,
              fill: { color: idx === 3 ? BURGUNDY : 'FFFFFF' },
              line: { color: BURGUNDY, width: 1 },
            });
            pptxSlide.addText(
              [
                { text: `Step ${idx + 1}:  `, options: { fontSize: 10, bold: true, color: idx === 3 ? 'FFFFFF' : MAUVE } },
                { text: stage, options: { fontSize: 11, bold: true, color: idx === 3 ? 'FFFFFF' : BURGUNDY } },
              ],
              { x: 5.5, y: y + 0.1, w: 3.7, h: 0.55, fontFace: FONT_SANS, valign: 'middle' }
            );
          });
        }
        break;
      }

      case 'cards': {
        const items = s.items;
        if (s.layout === 'rows') {
          items.forEach((it, idx) => {
            const y = 1.5 + idx * 0.68;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y,
              w: 8.8,
              h: 0.58,
              rectRadius: 0.06,
              fill: { color: 'FFFFFF' },
              line: { color: BORDER_COLOR, width: 0.8 },
            });
            pptxSlide.addText(`•  ${it.title}`, {
              x: 0.9,
              y,
              w: 8.2,
              h: 0.58,
              fontSize: 12,
              bold: true,
              color: BURGUNDY,
              fontFace: FONT_SANS,
              valign: 'middle',
            });
          });
        } else {
          // Grid layout (2, 3 or 4 columns)
          const cols = s.columns || (items.length <= 4 ? 2 : 3);
          const colW = (8.8 - (cols - 1) * 0.25) / cols;
          const rows = Math.ceil(items.length / cols);
          const rowH = (3.4 - (rows - 1) * 0.25) / rows;

          items.forEach((it, idx) => {
            const col = idx % cols;
            const row = Math.floor(idx / cols);
            const x = 0.6 + col * (colW + 0.25);
            const y = 1.5 + row * (rowH + 0.25);

            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x,
              y,
              w: colW,
              h: rowH,
              rectRadius: 0.08,
              fill: { color: 'FFFFFF' },
              line: { color: BORDER_COLOR, width: 0.8 },
            });

            if (it.body) {
              pptxSlide.addText(
                [
                  { text: `${it.title}\n`, options: { fontSize: 11, bold: true, color: BURGUNDY } },
                  { text: it.body, options: { fontSize: 9, color: GRAY } },
                ],
                { x: x + 0.15, y: y + 0.1, w: colW - 0.3, h: rowH - 0.2, fontFace: FONT_SANS, margin: 0 }
              );
            } else {
              pptxSlide.addText(it.title, {
                x: x + 0.15,
                y,
                w: colW - 0.3,
                h: rowH,
                fontSize: 11,
                bold: true,
                color: BURGUNDY,
                fontFace: FONT_SANS,
                valign: 'middle',
                align: 'center',
              });
            }
          });
        }

        if (s.flag) {
          pptxSlide.addText(`Note: ${s.flag}`, {
            x: 0.6,
            y: 5.05,
            w: 8.8,
            h: 0.25,
            fontSize: 7.5,
            color: MAUVE,
            fontFace: FONT_SANS,
            italic: true,
          });
        }
        break;
      }

      case 'types': {
        // Adenocarcinoma 95%
        pptxSlide.addShape(pptx.ShapeType.roundRect, {
          x: 0.6,
          y: 1.6,
          w: 8.8,
          h: 2.5,
          rectRadius: 0.12,
          fill: { color: 'FFFFFF' },
          line: { color: BORDER_COLOR, width: 1 },
        });

        pptxSlide.addShape(pptx.ShapeType.ellipse, {
          x: 1.0,
          y: 1.95,
          w: 1.8,
          h: 1.8,
          fill: { color: BURGUNDY },
        });
        pptxSlide.addText('95%', {
          x: 1.0,
          y: 1.95,
          w: 1.8,
          h: 1.8,
          fontSize: 32,
          bold: true,
          color: 'FFFFFF',
          fontFace: FONT_SERIF,
          align: 'center',
          valign: 'middle',
        });

        pptxSlide.addText(
          [
            { text: `${s.item.title}\n\n`, options: { fontSize: 22, bold: true, color: BURGUNDY, fontFace: FONT_SERIF } },
            { text: s.item.body || '', options: { fontSize: 13, color: INK, fontFace: FONT_SANS, lineSpacing: 18 } },
          ],
          { x: 3.2, y: 1.95, w: 5.8, h: 1.8, fontFace: FONT_SANS, margin: 0 }
        );
        break;
      }

      case 'symptoms': {
        s.items.forEach((it, idx) => {
          const y = 1.6 + idx * 1.1;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 8.8,
            h: 0.95,
            rectRadius: 0.08,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });

          // Spot indicator
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.85,
            y: y + 0.25,
            w: 1.1,
            h: 0.45,
            rectRadius: 0.05,
            fill: { color: ACCENT_PILL },
            line: { color: BURGUNDY, width: 0.6 },
          });
          pptxSlide.addText(it.spot.toUpperCase(), {
            x: 0.85,
            y: y + 0.25,
            w: 1.1,
            h: 0.45,
            fontSize: 8.5,
            bold: true,
            color: BURGUNDY,
            fontFace: FONT_SANS,
            align: 'center',
            valign: 'middle',
          });

          pptxSlide.addText(
            [
              { text: `${it.title}\n`, options: { fontSize: 13, bold: true, color: BURGUNDY } },
              { text: it.body, options: { fontSize: 10, color: GRAY } },
            ],
            { x: 2.15, y: y + 0.15, w: 7.0, h: 0.65, fontFace: FONT_SANS, margin: 0 }
          );
        });
        break;
      }

      case 'complications': {
        s.items.forEach((it, idx) => {
          const x = 0.6 + idx * 4.55;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x,
            y: 1.6,
            w: 4.25,
            h: 3.2,
            rectRadius: 0.1,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 1 },
          });

          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: x + 0.25,
            y: 1.85,
            w: 3.75,
            h: 0.5,
            rectRadius: 0.06,
            fill: { color: ACCENT_PILL },
          });
          pptxSlide.addText(it.title, {
            x: x + 0.25,
            y: 1.85,
            w: 3.75,
            h: 0.5,
            fontSize: 13,
            bold: true,
            color: BURGUNDY,
            fontFace: FONT_SERIF,
            align: 'center',
            valign: 'middle',
          });

          pptxSlide.addText(it.body || '', {
            x: x + 0.35,
            y: 2.6,
            w: 3.55,
            h: 2.0,
            fontSize: 12,
            color: INK,
            fontFace: FONT_SANS,
            lineSpacing: 18,
            align: 'center',
          });
        });
        break;
      }

      case 'exam': {
        s.steps.forEach((step, idx) => {
          const y = 1.8 + idx * 1.3;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 8.8,
            h: 1.1,
            rectRadius: 0.08,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });
          pptxSlide.addShape(pptx.ShapeType.ellipse, {
            x: 0.9,
            y: y + 0.3,
            w: 0.5,
            h: 0.5,
            fill: { color: BURGUNDY },
          });
          pptxSlide.addText(String(idx + 1), {
            x: 0.9,
            y: y + 0.3,
            w: 0.5,
            h: 0.5,
            fontSize: 12,
            bold: true,
            color: 'FFFFFF',
            fontFace: FONT_SANS,
            align: 'center',
            valign: 'middle',
          });
          pptxSlide.addText(step, {
            x: 1.6,
            y: y + 0.15,
            w: 7.5,
            h: 0.8,
            fontSize: 14,
            color: INK,
            fontFace: FONT_SANS,
            valign: 'middle',
          });
        });
        break;
      }

      case 'tests': {
        DIAGNOSTIC_TESTS.forEach((t, idx) => {
          const y = 1.5 + idx * 1.15;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 5.4,
            h: 1.0,
            rectRadius: 0.08,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });

          pptxSlide.addText(
            [
              { text: `${t.title}\n`, options: { fontSize: 13, bold: true, color: BURGUNDY } },
              { text: t.text, options: { fontSize: 10, color: GRAY } },
            ],
            { x: 0.85, y: y + 0.15, w: 4.9, h: 0.7, fontFace: FONT_SANS, margin: 0 }
          );
        });

        // Right side: Medical Infographic Reference Card
        pptxSlide.addShape(pptx.ShapeType.roundRect, {
          x: 6.3,
          y: 1.5,
          w: 3.1,
          h: 3.45,
          rectRadius: 0.1,
          fill: { color: 'FFFFFF' },
          line: { color: BORDER_COLOR, width: 1 },
        });
        pptxSlide.addImage({ path: getImg(INFOGRAPHIC_DIAG), x: 6.45, y: 1.65, w: 2.8, h: 2.8 });
        pptxSlide.addText('Clinical Diagnostic & Staging Infographic', {
          x: 6.3,
          y: 4.6,
          w: 3.1,
          h: 0.25,
          fontSize: 8,
          color: BURGUNDY,
          bold: true,
          align: 'center',
        });
        break;
      }

      case 'whyTests': {
        pptxSlide.addShape(pptx.ShapeType.roundRect, {
          x: 1.2,
          y: 2.0,
          w: 7.6,
          h: 1.8,
          rectRadius: 0.12,
          fill: { color: 'FFFFFF' },
          line: { color: BURGUNDY, width: 1.5 },
        });
        pptxSlide.addText(s.text, {
          x: 1.5,
          y: 2.2,
          w: 7.0,
          h: 1.4,
          fontSize: 18,
          bold: true,
          color: BURGUNDY,
          fontFace: FONT_SERIF,
          align: 'center',
          valign: 'middle',
        });
        break;
      }

      case 'stages': {
        s.items.forEach((it, idx) => {
          const y = 1.6 + idx * 1.05;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 5.4,
            h: 0.9,
            rectRadius: 0.08,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });

          pptxSlide.addShape(pptx.ShapeType.ellipse, {
            x: 0.85,
            y: y + 0.2,
            w: 0.5,
            h: 0.5,
            fill: { color: BURGUNDY },
          });
          pptxSlide.addText(it.stage, {
            x: 0.85,
            y: y + 0.2,
            w: 0.5,
            h: 0.5,
            fontSize: 11,
            bold: true,
            color: 'FFFFFF',
            fontFace: FONT_SERIF,
            align: 'center',
            valign: 'middle',
          });

          pptxSlide.addText(
            [
              { text: `Stage ${it.stage}\n`, options: { fontSize: 12, bold: true, color: BURGUNDY } },
              { text: it.text, options: { fontSize: 10, color: GRAY } },
            ],
            { x: 1.55, y: y + 0.12, w: 4.2, h: 0.66, fontFace: FONT_SANS, margin: 0 }
          );
        });

        // Infographic reference on right
        pptxSlide.addShape(pptx.ShapeType.roundRect, {
          x: 6.3,
          y: 1.6,
          w: 3.1,
          h: 3.2,
          rectRadius: 0.1,
          fill: { color: 'FFFFFF' },
          line: { color: BORDER_COLOR, width: 1 },
        });
        pptxSlide.addImage({ path: getImg(INFOGRAPHIC_DIAG), x: 6.45, y: 1.75, w: 2.8, h: 2.65 });
        pptxSlide.addText('TNM Staging System Reference', {
          x: 6.3,
          y: 4.5,
          w: 3.1,
          h: 0.25,
          fontSize: 8,
          color: BURGUNDY,
          bold: true,
          align: 'center',
        });
        break;
      }

      case 'prognosis': {
        if (s.mode === 'outlook') {
          // Outlook points + 3D tumor colon model
          s.points.forEach((pt, pIdx) => {
            const y = 1.6 + pIdx * 1.0;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y,
              w: 5.4,
              h: 0.88,
              rectRadius: 0.08,
              fill: { color: 'FFFFFF' },
              line: { color: BORDER_COLOR, width: 0.8 },
            });
            pptxSlide.addText(pt, {
              x: 0.85,
              y: y + 0.1,
              w: 4.9,
              h: 0.68,
              fontSize: 11,
              color: INK,
              fontFace: FONT_SANS,
              bullet: true,
              valign: 'middle',
            });
          });

          // 3D Tumor Visual
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 6.3,
            y: 1.6,
            w: 3.1,
            h: 3.2,
            rectRadius: 0.1,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });
          pptxSlide.addImage({ path: getImg(SCENE_TUMOR), x: 6.45, y: 1.8, w: 2.8, h: 2.4 });
          pptxSlide.addText('3D Colon Tumor Visualization', {
            x: 6.3,
            y: 4.45,
            w: 3.1,
            h: 0.25,
            fontSize: 8.5,
            color: GRAY,
            align: 'center',
            italic: true,
          });
        } else {
          // Prognosis points + Trajectory Gauge + Takeaway
          s.points.forEach((pt, pIdx) => {
            const y = 1.5 + pIdx * 0.75;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y,
              w: 5.4,
              h: 0.65,
              rectRadius: 0.06,
              fill: { color: 'FFFFFF' },
              line: { color: BORDER_COLOR, width: 0.8 },
            });
            pptxSlide.addText(pt, {
              x: 0.8,
              y: y + 0.05,
              w: 5.0,
              h: 0.55,
              fontSize: 10,
              color: INK,
              fontFace: FONT_SANS,
              bullet: true,
              valign: 'middle',
            });
          });

          // Right Infographic Reference
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 6.3,
            y: 1.5,
            w: 3.1,
            h: 2.2,
            rectRadius: 0.08,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });
          pptxSlide.addImage({ path: getImg(INFOGRAPHIC_PROG), x: 6.45, y: 1.6, w: 2.8, h: 1.8 });
          pptxSlide.addText('Prognosis & Survival Infographic', {
            x: 6.3,
            y: 3.45,
            w: 3.1,
            h: 0.2,
            fontSize: 7.5,
            color: BURGUNDY,
            bold: true,
            align: 'center',
          });

          if (s.takeaway) {
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y: 3.85,
              w: 8.8,
              h: 1.1,
              rectRadius: 0.08,
              fill: { color: ACCENT_PILL },
              line: { color: BURGUNDY, width: 1 },
            });
            pptxSlide.addText(
              [
                { text: `${s.takeawayLabel || 'Key Takeaway'}:  `, options: { fontSize: 11, bold: true, color: BURGUNDY } },
                { text: s.takeaway, options: { fontSize: 10, color: INK } },
              ],
              { x: 0.85, y: 3.95, w: 8.3, h: 0.9, fontFace: FONT_SANS, margin: 0, valign: 'middle' }
            );
          }
        }
        break;
      }

      case 'surgery': {
        pptxSlide.addText(s.lead, {
          x: 0.6,
          y: 1.45,
          w: 8.8,
          h: 0.3,
          fontSize: 11,
          bold: true,
          color: MAUVE,
          fontFace: FONT_SANS,
        });

        s.items.forEach((it, idx) => {
          const y = 1.85 + idx * 0.98;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 5.4,
            h: 0.88,
            rectRadius: 0.08,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });
          pptxSlide.addText(
            [
              { text: `${it.title}\n`, options: { fontSize: 12, bold: true, color: BURGUNDY } },
              { text: it.body || '', options: { fontSize: 10, color: GRAY } },
            ],
            { x: 0.85, y: y + 0.12, w: 4.9, h: 0.64, fontFace: FONT_SANS, margin: 0 }
          );
        });

        // Treatment Infographic
        pptxSlide.addShape(pptx.ShapeType.roundRect, {
          x: 6.3,
          y: 1.85,
          w: 3.1,
          h: 3.0,
          rectRadius: 0.1,
          fill: { color: 'FFFFFF' },
          line: { color: BORDER_COLOR, width: 1 },
        });
        pptxSlide.addImage({ path: getImg(INFOGRAPHIC_TX), x: 6.45, y: 1.95, w: 2.8, h: 2.5 });
        pptxSlide.addText('Surgical Interventions Overview', {
          x: 6.3,
          y: 4.55,
          w: 3.1,
          h: 0.25,
          fontSize: 8,
          color: BURGUNDY,
          bold: true,
          align: 'center',
        });
        break;
      }

      case 'additional': {
        pptxSlide.addText(s.lead, {
          x: 0.6,
          y: 1.45,
          w: 8.8,
          h: 0.3,
          fontSize: 11,
          bold: true,
          color: MAUVE,
          fontFace: FONT_SANS,
        });

        s.items.forEach((it, idx) => {
          const x = 0.6 + idx * 3.0;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x,
            y: 1.85,
            w: 2.8,
            h: 0.9,
            rectRadius: 0.08,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });
          pptxSlide.addText(it, {
            x,
            y: 1.85,
            w: 2.8,
            h: 0.9,
            fontSize: 13,
            bold: true,
            color: BURGUNDY,
            fontFace: FONT_SERIF,
            align: 'center',
            valign: 'middle',
          });
        });

        // Seek Medical Care Card
        pptxSlide.addShape(pptx.ShapeType.roundRect, {
          x: 0.6,
          y: 3.0,
          w: 8.8,
          h: 1.8,
          rectRadius: 0.1,
          fill: { color: ACCENT_PILL },
          line: { color: BURGUNDY, width: 1 },
        });
        pptxSlide.addText(s.seekTitle, {
          x: 0.9,
          y: 3.15,
          w: 8.2,
          h: 0.35,
          fontSize: 14,
          bold: true,
          color: BURGUNDY,
          fontFace: FONT_SERIF,
        });
        pptxSlide.addText(s.seekText, {
          x: 0.9,
          y: 3.55,
          w: 8.2,
          h: 1.1,
          fontSize: 12,
          color: INK,
          fontFace: FONT_SANS,
          lineSpacing: 18,
        });
        break;
      }

      case 'questions': {
        s.items.forEach((q, idx) => {
          const y = 1.7 + idx * 1.1;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 8.8,
            h: 0.9,
            rectRadius: 0.08,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });
          pptxSlide.addShape(pptx.ShapeType.ellipse, {
            x: 0.9,
            y: y + 0.2,
            w: 0.5,
            h: 0.5,
            fill: { color: ACCENT_PILL },
            line: { color: BURGUNDY, width: 0.8 },
          });
          pptxSlide.addText('?', {
            x: 0.9,
            y: y + 0.2,
            w: 0.5,
            h: 0.5,
            fontSize: 14,
            bold: true,
            color: BURGUNDY,
            fontFace: FONT_SERIF,
            align: 'center',
            valign: 'middle',
          });
          pptxSlide.addText(q, {
            x: 1.6,
            y: y + 0.15,
            w: 7.5,
            h: 0.6,
            fontSize: 14,
            bold: true,
            color: BURGUNDY,
            fontFace: FONT_SERIF,
            valign: 'middle',
          });
        });
        break;
      }

      case 'screening': {
        s.points.forEach((pt, pIdx) => {
          const y = 1.6 + pIdx * 1.1;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 8.8,
            h: 0.95,
            rectRadius: 0.08,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });
          pptxSlide.addShape(pptx.ShapeType.ellipse, {
            x: 0.9,
            y: y + 0.22,
            w: 0.5,
            h: 0.5,
            fill: { color: ACCENT_PILL },
          });
          pptxSlide.addText('✓', {
            x: 0.9,
            y: y + 0.22,
            w: 0.5,
            h: 0.5,
            fontSize: 14,
            bold: true,
            color: BURGUNDY,
            align: 'center',
            valign: 'middle',
          });
          pptxSlide.addText(pt, {
            x: 1.6,
            y: y + 0.12,
            w: 7.5,
            h: 0.7,
            fontSize: 12.5,
            color: INK,
            fontFace: FONT_SANS,
            valign: 'middle',
          });
        });
        break;
      }

      case 'steps': {
        pptxSlide.addText(`Phase: ${s.phase}`, {
          x: 0.6,
          y: 1.45,
          w: 8.8,
          h: 0.28,
          fontSize: 11,
          bold: true,
          color: MAUVE,
          fontFace: FONT_SANS,
        });

        s.items.forEach((it, idx) => {
          const y = 1.8 + idx * 0.78;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 8.8,
            h: 0.7,
            rectRadius: 0.06,
            fill: { color: 'FFFFFF' },
            line: { color: BORDER_COLOR, width: 0.8 },
          });
          pptxSlide.addText(
            [
              { text: `${it.title}:  `, options: { fontSize: 11, bold: true, color: BURGUNDY } },
              { text: it.body || '', options: { fontSize: 10, color: INK } },
            ],
            { x: 0.85, y: y + 0.08, w: 8.3, h: 0.54, fontFace: FONT_SANS, margin: 0, valign: 'middle' }
          );
        });
        break;
      }

      case 'careplan': {
        const plan = CARE_PLANS[s.plan];
        if (s.part === 'dx') {
          // Diagnosis & Goals Card
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y: 1.5,
            w: 8.8,
            h: 1.4,
            rectRadius: 0.08,
            fill: { color: 'FFFFFF' },
            line: { color: BURGUNDY, width: 1.2 },
          });
          pptxSlide.addText(plan.badge.toUpperCase(), {
            x: 0.85,
            y: 1.6,
            w: 8.3,
            h: 0.25,
            fontSize: 8.5,
            bold: true,
            color: MAUVE,
            charSpacing: 1.5,
          });
          pptxSlide.addText(plan.diagnosis, {
            x: 0.85,
            y: 1.9,
            w: 8.3,
            h: 0.9,
            fontSize: 12.5,
            bold: true,
            color: BURGUNDY,
            fontFace: FONT_SERIF,
          });

          // Subjective / Objective Data
          if (plan.subjective || plan.objective) {
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y: 3.05,
              w: 4.3,
              h: 0.9,
              rectRadius: 0.06,
              fill: { color: 'FFFFFF' },
              line: { color: BORDER_COLOR, width: 0.8 },
            });
            pptxSlide.addText(
              [
                { text: 'Subjective Data:\n', options: { fontSize: 9, bold: true, color: MAUVE } },
                { text: plan.subjective || 'N/A', options: { fontSize: 9.5, color: INK } },
              ],
              { x: 0.75, y: 3.12, w: 4.0, h: 0.75, fontFace: FONT_SANS, margin: 0 }
            );

            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 5.1,
              y: 3.05,
              w: 4.3,
              h: 0.9,
              rectRadius: 0.06,
              fill: { color: 'FFFFFF' },
              line: { color: BORDER_COLOR, width: 0.8 },
            });
            pptxSlide.addText(
              [
                { text: 'Objective Data:\n', options: { fontSize: 9, bold: true, color: MAUVE } },
                { text: plan.objective || 'N/A', options: { fontSize: 9.5, color: INK } },
              ],
              { x: 5.25, y: 3.12, w: 4.0, h: 0.75, fontFace: FONT_SANS, margin: 0 }
            );
          }

          // Expected Outcomes
          plan.outcomes.forEach((out, idx) => {
            const x = 0.6 + idx * 4.5;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x,
              y: 4.1,
              w: 4.3,
              h: 0.9,
              rectRadius: 0.06,
              fill: { color: ACCENT_PILL },
              line: { color: BORDER_COLOR, width: 0.8 },
            });
            pptxSlide.addText(
              [
                { text: `${out.label}\n`, options: { fontSize: 9, bold: true, color: BURGUNDY } },
                { text: out.text, options: { fontSize: 9.5, color: INK } },
              ],
              { x: x + 0.15, y: 4.18, w: 4.0, h: 0.75, fontFace: FONT_SANS, margin: 0 }
            );
          });
        } else {
          // Interventions & Rationales Table
          const tableRows: pptxgen.TableRow[] = [
            [
              { text: 'Nursing Interventions', options: { bold: true, color: 'FFFFFF', fill: { color: BURGUNDY }, fontSize: 10, align: 'center' } },
              { text: 'Scientific & Evidence-Based Rationales', options: { bold: true, color: 'FFFFFF', fill: { color: BURGUNDY }, fontSize: 10, align: 'center' } },
            ],
          ];

          plan.interventions.forEach((inv) => {
            tableRows.push([
              { text: inv.action, options: { fontSize: 9.5, color: INK, fill: { color: 'FFFFFF' }, valign: 'middle' } },
              { text: inv.rationale, options: { fontSize: 9.5, color: GRAY, fill: { color: 'FFFFFF' }, valign: 'middle' } },
            ]);
          });

          pptxSlide.addTable(tableRows, {
            x: 0.6,
            y: 1.5,
            w: 8.8,
            colW: [4.4, 4.4],
            border: { color: BORDER_COLOR, pt: 1 },
            autoPage: false,
          });
        }
        break;
      }

      case 'pending': {
        pptxSlide.addShape(pptx.ShapeType.roundRect, {
          x: 1.5,
          y: 1.8,
          w: 7.0,
          h: 2.4,
          rectRadius: 0.15,
          fill: { color: 'FFFFFF' },
          line: { color: BORDER_COLOR, width: 1.5, dashType: 'dash' },
        });

        pptxSlide.addText(s.title.toUpperCase(), {
          x: 1.8,
          y: 2.1,
          w: 6.4,
          h: 0.4,
          fontSize: 16,
          bold: true,
          color: BURGUNDY,
          fontFace: FONT_SERIF,
          align: 'center',
        });

        pptxSlide.addText(
          'Academic section scheduled in course curriculum.\nContent preserved pending upcoming clinical symposium update.',
          {
            x: 1.8,
            y: 2.65,
            w: 6.4,
            h: 0.9,
            fontSize: 12,
            color: GRAY,
            fontFace: FONT_SANS,
            align: 'center',
            lineSpacing: 18,
          }
        );
        break;
      }
    }
  }

  return pptx;
}
