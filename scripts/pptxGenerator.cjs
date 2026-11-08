"use strict";
var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

// src/export/pptxGenerator.ts
var pptxGenerator_exports = {};
__export(pptxGenerator_exports, {
  createPresentation: () => createPresentation
});
module.exports = __toCommonJS(pptxGenerator_exports);
var import_pptxgenjs = __toESM(require("pptxgenjs"), 1);

// src/content/slides.ts
var SECTIONS = [
  { id: "cover", label: "Cover" },
  { id: "goals", label: "Goals" },
  { id: "objectives", label: "Objectives" },
  { id: "outline", label: "Outlines" },
  { id: "introduction", label: "Introduction" },
  { id: "definition", label: "Definition" },
  { id: "anatomy", label: "Anatomy and Physiology" },
  { id: "pathophysiology", label: "Pathophysiology" },
  { id: "risk", label: "Risk Factors" },
  { id: "causes", label: "Causes" },
  { id: "types", label: "Types" },
  { id: "symptoms", label: "Signs and Symptoms" },
  { id: "complications", label: "Complications" },
  { id: "diagnostics", label: "Diagnostic Tests" },
  { id: "prognosis", label: "Prognosis" },
  { id: "treatment", label: "Management and Treatment" },
  { id: "prevention", label: "Prevention" },
  { id: "nursingRole", label: "Nursing Role" },
  { id: "nursingProcess", label: "Nursing Process" },
  { id: "summary", label: "Summary" },
  { id: "references", label: "References" }
];
var ANATOMY_SECTIONS = [
  { key: "ascending", label: "Ascending Colon", text: "Located on the right side; carries contents upward." },
  { key: "transverse", label: "Transverse Colon", text: "Extends across the upper abdomen from right to left." },
  { key: "descending", label: "Descending Colon", text: "Runs downward along the left side." },
  { key: "sigmoid", label: "Sigmoid Colon", text: "S-shaped section connecting the descending colon to the rectum." }
];
var PATHO_STAGES = [
  "Normal colon cells",
  "Abnormal cellular changes",
  "Polyp formation",
  "Cancer development"
];
var PENDING = "Academic section scheduled in course curriculum.";
var slides = [
  { id: "cover", section: "cover", kind: "cover", title: "Colon Cancer", subtitle: "Nursing College Presentation" },
  { id: "team", section: "cover", kind: "team", title: "Presentation Team" },
  {
    id: "goals",
    section: "goals",
    kind: "goal",
    title: "GOAL",
    text: "The goal of this assignment is to provide nursing students with a comprehensive understanding of colon cancer, including its anatomy, pathophysiology, risk factors, clinical manifestations, diagnosis, and management, in order to enhance their ability to deliver safe, evidence-based nursing care to patients with colon cancer."
  },
  {
    id: "objectives",
    section: "objectives",
    kind: "objectives",
    title: "OBJECTIVES",
    lead: "By the end of this assignment, the student will be able to:",
    items: [
      "Define colon cancer and describe the anatomy and physiology of the colon.",
      "Explain the pathophysiology, causes, and risk factors of colon cancer.",
      "Identify the types, signs and symptoms, and complications of colon cancer.",
      "Discuss the diagnostic tests, staging, and prognosis of colon cancer.",
      "Outline the medical and surgical management and the preventive measures of colon cancer.",
      "Describe the role of the nurse in caring for patients with colon cancer before and after surgery.",
      "Apply the nursing process to a patient with colon cancer."
    ]
  },
  { id: "outline", section: "outline", kind: "outline", title: "Outline" },
  {
    id: "intro-1",
    section: "introduction",
    kind: "intro",
    title: "Introduction",
    visual: true,
    points: [
      "Colon cancer is a major cancer of the digestive system.",
      "It starts when colon lining cells grow abnormally and uncontrollably.",
      "Often begins as small polyps that may turn cancerous over several years."
    ]
  },
  {
    id: "intro-2",
    section: "introduction",
    kind: "intro",
    title: "Introduction: Risk Factors",
    chipsLabel: "Risk factors",
    chips: ["Older age", "Family history", "Inherited genetic conditions", "Inflammatory bowel disease", "Smoking", "Lifestyle"],
    statement: "It can also occur without any known risk factors."
  },
  {
    id: "definition",
    section: "definition",
    kind: "definition",
    title: "Definition",
    points: [
      "A malignant disease of the colon, the longest part of the large intestine.",
      "Abnormal cells divide uncontrollably and form a malignant tumor.",
      "Most cases are adenocarcinomas, arising from glandular cells lining the colon."
    ]
  },
  {
    id: "anatomy-1",
    section: "anatomy",
    kind: "anatomy",
    mode: "facts",
    title: "Anatomy of the Colon",
    facts: [
      "The colon is the main part of the large intestine.",
      "It is approximately 1.5 meters long.",
      "Its main functions are absorbing water and electrolytes and forming solid stool."
    ],
    statement: "The colon has four main sections."
  },
  { id: "anatomy-2", section: "anatomy", kind: "anatomy", mode: "explore", title: "The Four Sections of the Colon" },
  {
    id: "patho-1",
    section: "pathophysiology",
    kind: "patho",
    mode: "genes",
    title: "Pathophysiology",
    genes: ["APC", "KRAS", "p53"],
    points: [
      "Genetic mutations (APC, KRAS, p53).",
      "Abnormal cells develop into polyps then cancer.",
      "DNA repair defects cause uncontrolled cell growth."
    ]
  },
  { id: "patho-2", section: "pathophysiology", kind: "patho", mode: "progression", title: "Pathophysiology: From Normal Cells to Cancer" },
  {
    id: "risk",
    section: "risk",
    kind: "cards",
    layout: "grid",
    title: "Risk Factors",
    items: [
      { title: "Age over 50" },
      { title: "Family history" },
      { title: "Obesity" },
      { title: "Smoking and alcohol" },
      { title: "Red and processed meat" },
      { title: "Physical inactivity" },
      { title: "Diabetes" }
    ],
    flag: "Age over 50 is listed here, while the Prevention slides note screening begins at age 45."
  },
  {
    id: "causes",
    section: "causes",
    kind: "cards",
    layout: "rows",
    title: "Causes",
    items: [
      { title: "Genetic mutations" },
      { title: "Hereditary syndromes" },
      { title: "Family history" },
      { title: "Chronic intestinal inflammation" },
      { title: "Unhealthy lifestyle" }
    ]
  },
  {
    id: "types-1",
    section: "types",
    kind: "types",
    title: "Types of Cancer in the Colon and Rectum",
    percent: 95,
    item: { title: "Adenocarcinoma", body: "The most common type (about 95%), develops in mucus-producing cells lining the colon and rectum." }
  },
  {
    id: "types-2",
    section: "types",
    kind: "cards",
    layout: "grid",
    columns: 2,
    title: "Types: Other Cancers of the Colon and Rectum",
    items: [
      { title: "Carcinoid Tumors", body: "Develop from hormone-producing cells in the intestine." },
      { title: "Gastrointestinal Stromal Tumors (GISTs)", body: "Develop from specialized cells in the digestive tract wall." },
      { title: "Lymphomas", body: "Cancers that develop from immune system cells and may affect the colon and rectum." },
      { title: "Sarcomas", body: "Rare cancers that develop in muscles, blood vessels, or connective tissues of the intestinal wall." }
    ]
  },
  {
    id: "symptoms-1",
    section: "symptoms",
    kind: "symptoms",
    title: "Signs and Symptoms of Colon Cancer",
    items: [
      { title: "Blood in Stool", body: "Blood may appear in the stool or as rectal bleeding.", spot: "rectum" },
      { title: "Change in Stool Appearance", body: "Changes in stool color, including dark or black stools.", spot: "lower" },
      { title: "Fatigue and Weakness", body: "May result from chronic bleeding and iron-deficiency anemia.", spot: "body" }
    ]
  },
  {
    id: "symptoms-2",
    section: "symptoms",
    kind: "symptoms",
    title: "Signs and Symptoms: Continued",
    items: [
      { title: "Unexplained Weight Loss", body: "Sudden weight loss without an obvious reason.", spot: "body" },
      { title: "Abdominal Bloating and Bowel Changes", body: "Persistent bloating or feeling that the bowel has not completely emptied.", spot: "abdomen" },
      { title: "No Symptoms", body: "Some patients may have no symptoms, especially in early stages.", spot: "none" }
    ]
  },
  {
    id: "compl-1",
    section: "complications",
    kind: "complications",
    title: "Complications of Colon Cancer",
    items: [
      { title: "Bowel Obstruction", art: "obstruction", body: "Tumor growth blocks the intestine, causing abdominal pain, swelling, nausea, and vomiting." },
      { title: "Gastrointestinal Bleeding", art: "bleeding", body: "Tumor bleeding may cause anemia, fatigue, and bloody or black stools." }
    ]
  },
  {
    id: "compl-2",
    section: "complications",
    kind: "complications",
    title: "Complications: Continued",
    items: [
      { title: "Perforation", art: "perforation", body: "A hole in the intestinal wall can lead to peritonitis and serious infection." },
      { title: "Metastasis", art: "metastasis", body: "Cancer spreads to nearby lymph nodes or distant organs, especially the liver and lungs." }
    ]
  },
  {
    id: "compl-3",
    section: "complications",
    kind: "complications",
    title: "Complications: Continued",
    items: [
      { title: "Malnutrition and Weight Loss", art: "malnutrition", body: "Reduced food intake and treatment side effects may cause weight loss and nutritional deficiencies." },
      { title: "Emotional and Psychological Impact", art: "emotional", body: "Cancer diagnosis and treatment may cause anxiety, stress, and depression." }
    ]
  },
  {
    id: "dx-exam",
    section: "diagnostics",
    kind: "exam",
    title: "Clinical Examination",
    steps: ["The doctor asks about symptoms and medical history.", "Examines the abdomen and may perform a digital rectal examination."]
  },
  { id: "dx-tests", section: "diagnostics", kind: "tests", title: "Diagnostic Tests" },
  { id: "dx-why", section: "diagnostics", kind: "whyTests", title: "Why These Tests?", text: "These tests help doctors diagnose cancer and determine its stage." },
  {
    id: "stages-1",
    section: "diagnostics",
    kind: "stages",
    title: "Stages of Colon Cancer",
    items: [
      { stage: "0", text: "Abnormal or precancerous cells." },
      { stage: "I", text: "Cancer is in the colon wall." },
      { stage: "II", text: "Cancer spreads deeper through the colon wall." }
    ]
  },
  {
    id: "stages-2",
    section: "diagnostics",
    kind: "stages",
    title: "Stages of Colon Cancer: Continued",
    items: [
      { stage: "III", text: "Cancer spreads to nearby lymph nodes." },
      { stage: "IV", text: "Cancer spreads to distant organs such as the liver or lungs." }
    ]
  },
  {
    id: "prog-1",
    section: "prognosis",
    kind: "prognosis",
    mode: "outlook",
    title: "Outlook of Colon Cancer",
    points: [
      "The outlook depends on the cancer stage, patient's general health, and response to treatment.",
      "Early diagnosis and treatment improve the chances of recovery.",
      "Treatment may include surgery, chemotherapy, and radiation therapy."
    ]
  },
  {
    id: "prog-2",
    section: "prognosis",
    kind: "prognosis",
    mode: "prognosis",
    title: "Prognosis of Colon Cancer",
    points: [
      "Prognosis depends on how early cancer is detected and whether it has spread.",
      "Early-stage cancer has a better prognosis, while advanced cancer has a poorer prognosis.",
      "Regular follow-up after treatment helps detect recurrence early."
    ],
    takeawayLabel: "Key takeaway",
    takeaway: "Regular follow-up after treatment helps detect cancer recurrence early. The chances of recovery also depend on the patient's overall health and response to treatment."
  },
  {
    id: "tx-1",
    section: "treatment",
    kind: "surgery",
    title: "Management and Treatment",
    lead: "Surgery is the main treatment for colon cancer that hasn't spread.",
    items: [
      { title: "Polypectomy", art: "polypectomy", body: "Removal of cancerous or precancerous polyps." },
      { title: "Colectomy", art: "colectomy", body: "Removal of the affected part of the colon." },
      { title: "Multivisceral Resection", art: "multivisceral", body: "Removal of the tumor and nearby affected organs." }
    ]
  },
  {
    id: "tx-2",
    section: "treatment",
    kind: "additional",
    title: "Additional Treatment and When to Seek Medical Care",
    lead: "Chemotherapy, radiation therapy and targeted therapy may be used before or after surgery.",
    items: ["Chemotherapy", "Radiation therapy", "Targeted therapy"],
    seekTitle: "When to Seek Medical Care",
    seekText: "Consult a doctor if you have persistent abdominal pain or blood in your stool, as these symptoms may have different causes and require medical evaluation."
  },
  {
    id: "tx-3",
    section: "treatment",
    kind: "questions",
    title: "Important Questions",
    items: ["What is the cancer stage?", "Can treatment eliminate the cancer?", "Can the cancer return?"]
  },
  {
    id: "prev-1",
    section: "prevention",
    kind: "screening",
    title: "Prevention: Screening & Early Detection",
    points: [
      "Regular screening is one of the most powerful tools for preventing colorectal cancer",
      "It takes 10 to 15 years for polyps to turn into cancer, screening allows removal before they develop",
      "Start screening at age 45 (or earlier if strong family history/genetic risk)"
    ]
  },
  {
    id: "prev-2",
    section: "prevention",
    kind: "cards",
    layout: "grid",
    columns: 2,
    title: "Lifestyle & Risk Reduction",
    items: [
      { title: "Healthy Weight", body: "Maintaining healthy weight lowers risk" },
      { title: "Physical Activity", body: "Regular moderate to vigorous activity lowers risk of polyps and cancer" },
      { title: "Diet", body: "Eat more vegetables, fruits, and whole grains. Limit red meats and processed meats." },
      { title: "Avoid Alcohol & Smoking", body: "Quitting reduces colorectal cancer risk" }
    ]
  },
  {
    id: "prev-3",
    section: "prevention",
    kind: "cards",
    layout: "grid",
    columns: 3,
    title: "Supplements & Medical Factors",
    items: [
      { title: "Vitamins & Minerals", body: "Folic acid, Vitamin D, Calcium, and Magnesium may help lower risk" },
      { title: "NSAIDs (Aspirin)", body: "May lower risk, but not recommended for prevention due to stomach bleeding" },
      { title: "Hormone Therapy", body: "May reduce risk in women, but carries other health risks" }
    ]
  },
  {
    id: "nr-1",
    section: "nursingRole",
    kind: "steps",
    phase: "Assessment and pre-operative care",
    art: "site",
    title: "Nursing Role: Assessment & Pre-Operative Care",
    items: [
      { title: "Comprehensive Assessment", body: "Changes in bowel habits, blood in stool, unexplained anemia, abdominal discomfort, nutritional status, family history" },
      { title: "Bowel Preparation", body: "Laxatives, enemas, or oral bowel cleansers" },
      { title: "Patient Education & Psychological Support", body: "Explain procedure, address fears and anxiety" },
      { title: "Stoma Site Marking", body: "Mark optimal site with ostomy specialist (WOCN)" }
    ]
  },
  {
    id: "nr-2",
    section: "nursingRole",
    kind: "steps",
    phase: "Post-operative care",
    art: "stoma",
    title: "Post-Operative Nursing Care",
    items: [
      { title: "Monitoring", body: "Vital signs, abdominal pain, bowel sounds, return of intestinal function" },
      { title: "Stoma Care", body: "Stoma should be pink/red and moist, teach pouch change and skin protection" },
      { title: "Prevent Complications", body: "Pain medication, early ambulation, deep breathing to prevent DVT" }
    ]
  },
  {
    id: "nr-3",
    section: "nursingRole",
    kind: "steps",
    phase: "Treatment side effects",
    title: "Managing Treatment Side Effects",
    items: [
      { title: "Chemotherapy & Radiation Support", body: "Manage nausea, vomiting, fatigue, diarrhea with medications and fluids to prevent dehydration" },
      { title: "Infection & Lab Monitoring", body: "Monitor WBCs and platelets for infection and bleeding risks" }
    ]
  },
  {
    id: "nr-4",
    section: "nursingRole",
    kind: "steps",
    phase: "Psychological and rehabilitation support",
    title: "Psychological & Rehabilitation Support",
    items: [
      { title: "Body Image & Coping", body: "Help patient adapt to bodily changes with colostomy" },
      { title: "Holistic Support", body: "Continuous emotional support for patient and family" }
    ]
  },
  { id: "cp-f1", section: "nursingProcess", kind: "careplan", plan: "fatigue", part: "dx", title: "Nursing Care Plan: Fatigue" },
  { id: "cp-f2", section: "nursingProcess", kind: "careplan", plan: "fatigue", part: "rx", title: "Fatigue: Interventions & Rationales" },
  { id: "cp-p1", section: "nursingProcess", kind: "careplan", plan: "pain", part: "dx", title: "Nursing Care Plan: Abdominal Pain" },
  { id: "cp-p2", section: "nursingProcess", kind: "careplan", plan: "pain", part: "rx", title: "Abdominal Pain: Interventions & Rationales" },
  { id: "cp-m1", section: "nursingProcess", kind: "careplan", plan: "motility", part: "dx", title: "Nursing Care Plan: Risk for Dysfunctional Gastrointestinal Motility" },
  { id: "cp-m2", section: "nursingProcess", kind: "careplan", plan: "motility", part: "rx", title: "Gastrointestinal Motility: Interventions & Rationales" },
  { id: "summary", section: "summary", kind: "pending", title: "Summary", note: PENDING },
  { id: "references", section: "references", kind: "pending", title: "References", note: PENDING }
];
var DIAGNOSTIC_TESTS = [
  { id: "blood", title: "Blood Tests", text: "CBC, CMP and CEA.", chips: ["CBC", "CMP", "CEA"] },
  { id: "imaging", title: "Imaging Tests", text: "CT scan, MRI, PET scan and abdominal ultrasound.", chips: ["CT scan", "MRI", "PET scan", "Abdominal ultrasound"] },
  { id: "colonoscopy", title: "Colonoscopy", text: "Allows the doctor to examine the colon and take a biopsy for laboratory examination.", chips: [] }
];

// src/content/carePlans.ts
var CARE_PLANS = {
  fatigue: {
    id: "fatigue",
    title: "Fatigue",
    badge: "Actual Nursing Care Plan",
    diagnosis: "Fatigue related to decreased oxygen carrying capacity secondary to anemia as evidenced by",
    evidencedBy: true,
    subjective: "Patient states feeling tired, weak, and dizzy.",
    objective: "Pallor of skin, low Hb/Hct levels, and decreased activity tolerance.",
    outcomes: [
      { label: "Short-Term Goal", text: "Patient will report decreased fatigue within 24\u201348 hours." },
      {
        label: "Long-Term Goal",
        text: "Patient will perform daily activities with minimal fatigue and demonstrate improved activity tolerance."
      }
    ],
    interventions: [
      {
        action: "Assess fatigue level, monitor vital signs regularly.",
        rationale: "To determine fatigue severity, evaluate anemia, and detect abnormalities early."
      },
      {
        action: "Provide adequate rest periods and assist with daily activities as needed.",
        rationale: "To conserve energy and prevent excessive exhaustion."
      },
      {
        action: "Encourage iron-rich foods, administer prescribed iron supplements.",
        rationale: "To support hemoglobin production, correct iron deficiency."
      }
    ]
  },
  pain: {
    id: "pain",
    title: "Abdominal Pain",
    badge: "Actual Nursing Care Plan",
    diagnosis: "Persistent abdominal pain related to tumor growth and intestinal changes secondary to colon cancer.",
    subjective: "\u201CI have abdominal pain.\u201D",
    objective: "Abdominal tenderness on palpation.",
    outcomes: [
      { label: "Short-Term Goal", text: "Patient will report reduced abdominal pain within 24\u201348 hours." },
      { label: "Long-Term Goal", text: "Patient will maintain pain control." }
    ],
    interventions: [
      {
        action: "Assess the patient's pain level, location, characteristics and vital signs.",
        rationale: "To determine the severity and pattern of pain, and to detect changes."
      },
      {
        action: "Administer prescribed analgesics.",
        rationale: "To relieve pain, reduce muscle tension, and improve patient comfort."
      },
      {
        action: "Educate the patient to eat small, frequent, low-fiber meals.",
        rationale: "To prevent stomach overloading, minimize bowel irritation, and allow early medical intervention."
      }
    ]
  },
  motility: {
    id: "motility",
    title: "Risk for Dysfunctional Gastrointestinal Motility",
    badge: "Potential Nursing Care Plan",
    diagnosis: "Risk for Dysfunctional Gastrointestinal Motility related to the presence of a colon tumor mass and potential narrowing of the intestinal lumen.",
    outcomes: [{ label: "Goal", text: "Patient will maintain stable gastrointestinal function." }],
    interventions: [
      {
        action: "Auscultate bowel sounds in all 4 quadrants every 4 hours.",
        rationale: "To detect early signs of bowel obstruction."
      },
      {
        action: "Encourage progressive ambulation and physical activity as tolerated.",
        rationale: "To stimulate peristalsis and enhance natural GI motility."
      },
      {
        action: "Maintain adequate hydration via oral fluids or prescribed IV routes.",
        rationale: "To prevent constipation and ease stool passage."
      }
    ]
  }
};

// src/export/pptxGenerator.ts
var import_meta = {};
var BURGUNDY = "7A1F3D";
var MAUVE = "9E4767";
var INK = "2B1F24";
var GRAY = "6E5E65";
var LIGHT_BG = "FDF6F7";
var BORDER_COLOR = "F0D5DC";
var ACCENT_PILL = "F8E3E8";
var DARK_BURGUNDY = "5A122A";
var FONT_SERIF = "Georgia";
var FONT_SANS = "Calibri";
var LOGO_NURSING = "public/images/nursing-logo-opt.png";
var LOGO_DEPT = "public/images/dept-logo-opt.png";
var LOGO_AZHAR = "public/images/azhar-logo-opt.png";
var INFOGRAPHIC_DIAG = "public/images/diag-stages-infographic.jpg";
var INFOGRAPHIC_PROG = "public/images/prognosis-infographic.jpg";
var INFOGRAPHIC_TX = "public/images/treatment-infographic.jpg";
var SCENE_DIGESTIVE = "public/images/scenes/3d-digestive.png";
var SCENE_ANATOMY = "public/images/scenes/3d-colon-anatomy.png";
var SCENE_SECTIONS = "public/images/scenes/3d-colon-sections.png";
var SCENE_PATHO = "public/images/scenes/3d-patho-lattice.png";
var SCENE_TUMOR = "public/images/scenes/3d-tumor-model.png";
function createPresentation() {
  const pptx = new import_pptxgenjs.default();
  pptx.layout = "LAYOUT_16x9";
  pptx.author = "Faculty of Nursing for Girls, Al-Azhar University";
  pptx.company = "Medical-Surgical Nursing Department";
  pptx.title = "Colon Cancer - Clinical Nursing Seminar";
  pptx.subject = "Colorectal Oncology & Evidence-Based Nursing Care Plans";
  const getImg = (relPath) => {
    if (typeof window !== "undefined") {
      const base = import_meta.env?.BASE_URL || "/";
      return `${base}${relPath.replace(/^public\//, "")}`;
    }
    return relPath;
  };
  const addInstitutionalHeader = (slide, sectionTitle) => {
    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: 10,
      h: 0.64,
      fill: { color: "FFFFFF" },
      line: { color: BORDER_COLOR, width: 0.8 }
    });
    slide.addImage({
      path: getImg(LOGO_NURSING),
      x: 0.25,
      y: 0.08,
      w: 0.48,
      h: 0.48
    });
    slide.addText(
      [
        { text: "Faculty of Nursing for Girls, Al-Azhar University\n", options: { fontSize: 7.5, bold: true, color: BURGUNDY } },
        { text: "\u0643\u064F\u0644\u064A\u0651\u0629 \u0627\u0644\u062A\u0645\u0631\u064A\u0636 \u0644\u0644\u0628\u0646\u0627\u062A \u0628\u0627\u0644\u0642\u0627\u0647\u0631\u0629", options: { fontSize: 6.5, bold: false, color: MAUVE } }
      ],
      { x: 0.8, y: 0.1, w: 2.8, h: 0.45, fontFace: FONT_SANS, margin: 0, align: "left" }
    );
    slide.addImage({
      path: getImg(LOGO_DEPT),
      x: 3.85,
      y: 0.08,
      w: 0.48,
      h: 0.48
    });
    slide.addText(
      [
        { text: "Medical-Surgical Nursing Department\n", options: { fontSize: 7.5, bold: true, color: BURGUNDY } },
        { text: "\u0642\u0633\u0645 \u0627\u0644\u062A\u0645\u0631\u064A\u0636 \u0627\u0644\u0628\u0627\u0637\u0646\u064A \u0648\u0627\u0644\u062C\u0631\u0627\u062D\u064A", options: { fontSize: 6.5, bold: false, color: MAUVE } }
      ],
      { x: 4.38, y: 0.1, w: 2.7, h: 0.45, fontFace: FONT_SANS, margin: 0, align: "left" }
    );
    slide.addText(
      [
        { text: "Al-Azhar University\n", options: { fontSize: 7.5, bold: true, color: BURGUNDY } },
        { text: "\u062C\u0627\u0645\u0639\u0629 \u0627\u0644\u0623\u0632\u0647\u0631 \u0627\u0644\u0634\u0631\u064A\u0641", options: { fontSize: 6.5, bold: false, color: MAUVE } }
      ],
      { x: 7.25, y: 0.1, w: 1.95, h: 0.45, fontFace: FONT_SANS, margin: 0, align: "right" }
    );
    slide.addImage({
      path: getImg(LOGO_AZHAR),
      x: 9.25,
      y: 0.08,
      w: 0.48,
      h: 0.48
    });
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
        charSpacing: 1.5
      });
    }
  };
  const addSlideTitle = (slide, title, yPos = 0.95) => {
    slide.addText(title, {
      x: 0.6,
      y: yPos,
      w: 8.8,
      h: 0.5,
      fontSize: 20,
      bold: true,
      color: BURGUNDY,
      fontFace: FONT_SERIF
    });
  };
  const slide1 = pptx.addSlide();
  slide1.background = { color: LIGHT_BG };
  slide1.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 10,
    h: 0.75,
    fill: { color: "FFFFFF" },
    line: { color: BORDER_COLOR, width: 1 }
  });
  slide1.addImage({ path: getImg(LOGO_NURSING), x: 0.3, y: 0.1, w: 0.55, h: 0.55 });
  slide1.addText(
    [
      { text: "Faculty of Nursing for Girls, Al-Azhar University\n", options: { fontSize: 8.5, bold: true, color: BURGUNDY } },
      { text: "\u0643\u064F\u0644\u064A\u0651\u0629 \u0627\u0644\u062A\u0645\u0631\u064A\u0636 \u0644\u0644\u0628\u0646\u0627\u062A \u0628\u0627\u0644\u0642\u0627\u0647\u0631\u0629", options: { fontSize: 7.5, color: MAUVE } }
    ],
    { x: 0.95, y: 0.12, w: 3.2, h: 0.5, fontFace: FONT_SANS, margin: 0 }
  );
  slide1.addImage({ path: getImg(LOGO_DEPT), x: 4.35, y: 0.1, w: 0.55, h: 0.55 });
  slide1.addText(
    [
      { text: "Medical-Surgical Nursing Department\n", options: { fontSize: 8.5, bold: true, color: BURGUNDY } },
      { text: "\u0642\u0633\u0645 \u0627\u0644\u062A\u0645\u0631\u064A\u0636 \u0627\u0644\u0628\u0627\u0637\u0646\u064A \u0648\u0627\u0644\u062C\u0631\u0627\u062D\u064A", options: { fontSize: 7.5, color: MAUVE } }
    ],
    { x: 4.95, y: 0.12, w: 2.7, h: 0.5, fontFace: FONT_SANS, margin: 0 }
  );
  slide1.addText(
    [
      { text: "Al-Azhar University\n", options: { fontSize: 8.5, bold: true, color: BURGUNDY } },
      { text: "\u062C\u0627\u0645\u0639\u0629 \u0627\u0644\u0623\u0632\u0647\u0631 \u0627\u0644\u0634\u0631\u064A\u0641", options: { fontSize: 7.5, color: MAUVE } }
    ],
    { x: 7.5, y: 0.12, w: 1.65, h: 0.5, fontFace: FONT_SANS, margin: 0, align: "right" }
  );
  slide1.addImage({ path: getImg(LOGO_AZHAR), x: 9.2, y: 0.1, w: 0.55, h: 0.55 });
  slide1.addText("COLON CANCER", {
    x: 0.6,
    y: 1.15,
    w: 5.5,
    h: 0.9,
    fontSize: 40,
    bold: true,
    color: BURGUNDY,
    fontFace: FONT_SERIF
  });
  slide1.addText("Nursing College Presentation", {
    x: 0.6,
    y: 2.05,
    w: 5.5,
    h: 0.4,
    fontSize: 16,
    bold: true,
    color: MAUVE,
    fontFace: FONT_SERIF
  });
  slide1.addText(
    "A Comprehensive Clinical Seminar on Colorectal Oncology, Surgical Interventions & Advanced Nursing Care Plans.",
    {
      x: 0.6,
      y: 2.45,
      w: 5.3,
      h: 0.7,
      fontSize: 11,
      color: GRAY,
      fontFace: FONT_SANS,
      lineSpacing: 14
    }
  );
  slide1.addShape(pptx.ShapeType.roundRect, {
    x: 0.6,
    y: 3.25,
    w: 5.3,
    h: 1.9,
    rectRadius: 0.1,
    fill: { color: "FFFFFF" },
    line: { color: BORDER_COLOR, width: 1 }
  });
  slide1.addText("UNDER SUPERVISION", {
    x: 0.8,
    y: 3.35,
    w: 4.9,
    h: 0.25,
    fontSize: 8,
    bold: true,
    color: BURGUNDY,
    fontFace: FONT_SANS,
    charSpacing: 1.5
  });
  slide1.addText(
    [
      { text: "Prof. Dr. Heba Abdel-Azem\n", options: { fontSize: 11, bold: true, color: BURGUNDY } },
      { text: "Professor of Medical Surgical Nursing Department\n\n", options: { fontSize: 9, color: MAUVE } },
      { text: "Dr. Eman Abel-Maqsoud\n", options: { fontSize: 11, bold: true, color: BURGUNDY } },
      { text: "Demonstrator of Medical Surgical Nursing Department", options: { fontSize: 9, color: MAUVE } }
    ],
    { x: 0.8, y: 3.65, w: 4.9, h: 1.1, fontFace: FONT_SANS, margin: 0 }
  );
  slide1.addText("Academic Year 2026 / 2027", {
    x: 3.7,
    y: 4.8,
    w: 2,
    h: 0.25,
    fontSize: 8,
    bold: true,
    color: DARK_BURGUNDY,
    fontFace: FONT_SANS,
    align: "right"
  });
  slide1.addShape(pptx.ShapeType.roundRect, {
    x: 6.2,
    y: 1.15,
    w: 3.4,
    h: 4,
    rectRadius: 0.12,
    fill: { color: "FFFFFF" },
    line: { color: BORDER_COLOR, width: 1 }
  });
  slide1.addImage({
    path: getImg(SCENE_DIGESTIVE),
    x: 6.35,
    y: 1.3,
    w: 3.1,
    h: 3.4
  });
  slide1.addText("3D Digestive Oncology Model", {
    x: 6.2,
    y: 4.85,
    w: 3.4,
    h: 0.25,
    fontSize: 8,
    color: GRAY,
    fontFace: FONT_SANS,
    align: "center",
    italic: true
  });
  const slide2 = pptx.addSlide();
  slide2.background = { color: LIGHT_BG };
  addInstitutionalHeader(slide2, "Presentation Team");
  addSlideTitle(slide2, "Presentation Team");
  slide2.addText("\u0625\u0639\u062F\u0627\u062F \u0637\u0627\u0644\u0628\u0627\u062A \u0643\u0644\u064A\u0629 \u0627\u0644\u062A\u0645\u0631\u064A\u0636 \u2022 \u0627\u0644\u0641\u0631\u0642\u0629 \u0627\u0644\u062F\u0631\u0627\u0633\u064A\u0629", {
    x: 0.6,
    y: 1.45,
    w: 8.8,
    h: 0.3,
    fontSize: 10,
    bold: true,
    color: BURGUNDY,
    fontFace: FONT_SANS,
    align: "center"
  });
  slide2.addText("Faculty of Nursing for Girls \u2022 Al-Azhar University \u2022 Medical-Surgical Nursing Department", {
    x: 0.6,
    y: 1.75,
    w: 8.8,
    h: 0.25,
    fontSize: 9,
    color: GRAY,
    fontFace: FONT_SANS,
    align: "center"
  });
  const students = [
    { num: 1, name: "Hager Hany" },
    { num: 2, name: "Hager Saeed" },
    { num: 3, name: "Hager Mostafa" },
    { num: 4, name: "Haidy Mahmoud" },
    { num: 5, name: "Norhan Badawy" },
    { num: 6, name: "Hager Ahmed" }
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
      fill: { color: "FFFFFF" },
      line: { color: BORDER_COLOR, width: 1 }
    });
    slide2.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.15,
      y: y + 0.22,
      w: 0.5,
      h: 0.5,
      fill: { color: ACCENT_PILL },
      line: { color: BORDER_COLOR, width: 0.8 }
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
      align: "center",
      valign: "middle"
    });
    slide2.addText(
      [
        { text: `${s.name}
`, options: { fontSize: 13, bold: true, color: BURGUNDY } },
        { text: "Nursing Student \u2022 Al-Azhar University", options: { fontSize: 8.5, color: GRAY } }
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
    fill: { color: "FFFFFF" },
    line: { color: BORDER_COLOR, width: 0.8 }
  });
  slide2.addText(
    "Clinical seminar prepared under the academic curriculum of the Medical-Surgical Nursing Department, Faculty of Nursing for Girls, Al-Azhar University.",
    {
      x: 0.8,
      y: 4.65,
      w: 8.4,
      h: 0.45,
      fontSize: 9.5,
      color: BURGUNDY,
      fontFace: FONT_SANS,
      align: "center"
    }
  );
  const slide3 = pptx.addSlide();
  slide3.background = { color: LIGHT_BG };
  addInstitutionalHeader(slide3, "Educational Framework");
  addSlideTitle(slide3, "GOAL");
  slide3.addShape(pptx.ShapeType.roundRect, {
    x: 0.6,
    y: 1.5,
    w: 5.6,
    h: 2.2,
    rectRadius: 0.1,
    fill: { color: "FFFFFF" },
    line: { color: BORDER_COLOR, width: 1 }
  });
  slide3.addText(
    "The goal of this assignment is to provide nursing students with a comprehensive understanding of colon cancer, including its anatomy, pathophysiology, risk factors, clinical manifestations, diagnosis, and management, in order to enhance their ability to deliver safe, evidence-based nursing care to patients with colon cancer.",
    {
      x: 0.85,
      y: 1.65,
      w: 5.1,
      h: 1.9,
      fontSize: 12.5,
      color: INK,
      fontFace: FONT_SERIF,
      lineSpacing: 18,
      italic: true
    }
  );
  const pillars = [
    "Anatomical & Pathophysiological Mastery",
    "Evidence-Based Diagnostic Assessment",
    "Multimodal Oncologic Management",
    "Patient-Centered Nursing Process & Care Plans"
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
      line: { color: BORDER_COLOR, width: 0.6 }
    });
    slide3.addText(`\u2713  ${p}`, {
      x: 0.75,
      y,
      w: 5.3,
      h: 0.28,
      fontSize: 9,
      bold: true,
      color: BURGUNDY,
      fontFace: FONT_SANS,
      valign: "middle"
    });
  });
  slide3.addShape(pptx.ShapeType.roundRect, {
    x: 6.4,
    y: 1.5,
    w: 3,
    h: 3.65,
    rectRadius: 0.1,
    fill: { color: "FFFFFF" },
    line: { color: BORDER_COLOR, width: 1 }
  });
  slide3.addImage({ path: getImg(LOGO_DEPT), x: 7.15, y: 1.8, w: 1.5, h: 1.5 });
  slide3.addText(
    [
      { text: "Medical-Surgical Nursing\n", options: { fontSize: 12, bold: true, color: BURGUNDY } },
      { text: "Clinical Competency Standards\n\n", options: { fontSize: 9.5, color: MAUVE } },
      { text: "Delivering holistic, evidence-based care tailored to oncology patients undergoing colorectal surgical and medical interventions.", options: { fontSize: 8.5, color: GRAY } }
    ],
    { x: 6.6, y: 3.45, w: 2.6, h: 1.5, fontFace: FONT_SANS, align: "center" }
  );
  const slide4 = pptx.addSlide();
  slide4.background = { color: LIGHT_BG };
  addInstitutionalHeader(slide4, "Educational Framework");
  addSlideTitle(slide4, "OBJECTIVES");
  slide4.addText("By the end of this assignment, the student will be able to:", {
    x: 0.6,
    y: 1.42,
    w: 8.8,
    h: 0.25,
    fontSize: 11,
    bold: true,
    color: MAUVE,
    fontFace: FONT_SANS
  });
  const objectives = [
    "Define colon cancer and describe the anatomy and physiology of the colon.",
    "Explain the pathophysiology, causes, and risk factors of colon cancer.",
    "Identify the types, signs and symptoms, and complications of colon cancer.",
    "Discuss the diagnostic tests, staging, and prognosis of colon cancer.",
    "Outline the medical and surgical management and the preventive measures of colon cancer.",
    "Describe the role of the nurse in caring for patients with colon cancer before and after surgery.",
    "Apply the nursing process to a patient with colon cancer."
  ];
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
        fill: { color: "FFFFFF" },
        line: { color: BORDER_COLOR, width: 0.8 }
      });
      slide4.addShape(pptx.ShapeType.ellipse, {
        x: x + 0.15,
        y: y + 0.18,
        w: 0.48,
        h: 0.48,
        fill: { color: ACCENT_PILL },
        line: { color: BORDER_COLOR, width: 0.6 }
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
        align: "center",
        valign: "middle"
      });
      slide4.addText(obj, {
        x: x + 0.72,
        y: y + 0.12,
        w: objW - 0.85,
        h: 0.65,
        fontSize: 9.5,
        color: INK,
        fontFace: FONT_SANS,
        valign: "middle"
      });
    } else {
      const y = 4.65;
      slide4.addShape(pptx.ShapeType.roundRect, {
        x: 0.6,
        y,
        w: 8.8,
        h: 0.6,
        rectRadius: 0.08,
        fill: { color: "FFFFFF" },
        line: { color: BURGUNDY, width: 1 }
      });
      slide4.addShape(pptx.ShapeType.ellipse, {
        x: 0.75,
        y: y + 0.1,
        w: 0.4,
        h: 0.4,
        fill: { color: BURGUNDY }
      });
      slide4.addText("7", {
        x: 0.75,
        y: y + 0.1,
        w: 0.4,
        h: 0.4,
        fontSize: 11,
        bold: true,
        color: "FFFFFF",
        fontFace: FONT_SANS,
        align: "center",
        valign: "middle"
      });
      slide4.addText(obj, {
        x: 1.25,
        y: y + 0.08,
        w: 8,
        h: 0.44,
        fontSize: 10,
        bold: true,
        color: BURGUNDY,
        fontFace: FONT_SANS,
        valign: "middle"
      });
    }
  });
  const slide5 = pptx.addSlide();
  slide5.background = { color: LIGHT_BG };
  addInstitutionalHeader(slide5, "Agenda");
  addSlideTitle(slide5, "Presentation Outlines");
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
      fill: { color: "FFFFFF" },
      line: { color: BORDER_COLOR, width: 0.6 }
    });
    slide5.addText(
      [
        { text: `${String(idx + 1).padStart(2, "0")}  `, options: { fontSize: 8.5, bold: true, color: MAUVE } },
        { text: sec.label, options: { fontSize: 9.5, bold: true, color: BURGUNDY } }
      ],
      { x: x + 0.15, y: y + 0.05, w: outlineColW - 0.25, h: outlineRowH - 0.1, fontFace: FONT_SANS, valign: "middle" }
    );
  });
  for (let i = 5; i < slides.length; i++) {
    const s = slides[i];
    const pptxSlide = pptx.addSlide();
    pptxSlide.background = { color: LIGHT_BG };
    const sectionObj = SECTIONS.find((sec) => sec.id === s.section);
    addInstitutionalHeader(pptxSlide, sectionObj?.label || "Clinical Seminar");
    addSlideTitle(pptxSlide, s.title);
    switch (s.kind) {
      case "intro": {
        if (s.points) {
          s.points.forEach((pt, pIdx) => {
            const y = 1.6 + pIdx * 1;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y,
              w: 5.5,
              h: 0.85,
              rectRadius: 0.08,
              fill: { color: "FFFFFF" },
              line: { color: BORDER_COLOR, width: 0.8 }
            });
            pptxSlide.addText(pt, {
              x: 0.85,
              y: y + 0.1,
              w: 5,
              h: 0.65,
              fontSize: 12,
              color: INK,
              fontFace: FONT_SANS,
              bullet: true,
              valign: "middle"
            });
          });
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 6.4,
            y: 1.6,
            w: 3,
            h: 3.2,
            rectRadius: 0.1,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
          });
          pptxSlide.addImage({ path: getImg(SCENE_DIGESTIVE), x: 6.55, y: 1.75, w: 2.7, h: 2.7 });
          pptxSlide.addText("Digestive System & Colonic Lining", {
            x: 6.4,
            y: 4.45,
            w: 3,
            h: 0.25,
            fontSize: 8.5,
            color: GRAY,
            align: "center",
            italic: true
          });
        } else if (s.chips) {
          pptxSlide.addText(s.chipsLabel || "Risk Factors", {
            x: 0.6,
            y: 1.5,
            w: 8.8,
            h: 0.35,
            fontSize: 14,
            bold: true,
            color: BURGUNDY,
            fontFace: FONT_SERIF
          });
          s.chips.forEach((chip, cIdx) => {
            const col = cIdx % 3;
            const row = Math.floor(cIdx / 3);
            const x = 0.6 + col * 2.95;
            const y = 2 + row * 0.85;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x,
              y,
              w: 2.8,
              h: 0.7,
              rectRadius: 0.08,
              fill: { color: "FFFFFF" },
              line: { color: BORDER_COLOR, width: 0.8 }
            });
            pptxSlide.addText(`\u2022  ${chip}`, {
              x: x + 0.15,
              y,
              w: 2.5,
              h: 0.7,
              fontSize: 11,
              bold: true,
              color: BURGUNDY,
              fontFace: FONT_SANS,
              valign: "middle"
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
              line: { color: BURGUNDY, width: 1 }
            });
            pptxSlide.addText(s.statement, {
              x: 0.9,
              y: 4,
              w: 8.2,
              h: 0.7,
              fontSize: 14,
              bold: true,
              color: BURGUNDY,
              fontFace: FONT_SERIF,
              align: "center",
              valign: "middle"
            });
          }
        }
        break;
      }
      case "definition": {
        s.points.forEach((pt, pIdx) => {
          const y = 1.6 + pIdx * 1.1;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 8.8,
            h: 0.95,
            rectRadius: 0.08,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
          });
          pptxSlide.addShape(pptx.ShapeType.ellipse, {
            x: 0.85,
            y: y + 0.22,
            w: 0.5,
            h: 0.5,
            fill: { color: ACCENT_PILL },
            line: { color: BORDER_COLOR, width: 0.6 }
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
            align: "center",
            valign: "middle"
          });
          pptxSlide.addText(pt, {
            x: 1.55,
            y: y + 0.12,
            w: 7.6,
            h: 0.7,
            fontSize: 13,
            color: INK,
            fontFace: FONT_SANS,
            valign: "middle"
          });
        });
        break;
      }
      case "anatomy": {
        if (s.mode === "facts") {
          s.facts?.forEach((fact, fIdx) => {
            const y = 1.6 + fIdx * 0.9;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y,
              w: 5.2,
              h: 0.78,
              rectRadius: 0.08,
              fill: { color: "FFFFFF" },
              line: { color: BORDER_COLOR, width: 0.8 }
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
              valign: "middle"
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
              line: { color: BURGUNDY, width: 0.8 }
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
              valign: "middle"
            });
          }
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 6.1,
            y: 1.6,
            w: 3.3,
            h: 3.4,
            rectRadius: 0.1,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
          });
          pptxSlide.addImage({ path: getImg(SCENE_ANATOMY), x: 6.25, y: 1.75, w: 3, h: 2.8 });
          pptxSlide.addText("3D Colon Anatomy Model", {
            x: 6.1,
            y: 4.65,
            w: 3.3,
            h: 0.25,
            fontSize: 8.5,
            color: GRAY,
            align: "center",
            italic: true
          });
        } else {
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y: 1.5,
            w: 4.3,
            h: 3.6,
            rectRadius: 0.1,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
          });
          pptxSlide.addImage({ path: getImg(SCENE_SECTIONS), x: 0.75, y: 1.65, w: 4, h: 3.1 });
          pptxSlide.addText("Interactive 3D Anatomical Segmentation", {
            x: 0.6,
            y: 4.8,
            w: 4.3,
            h: 0.25,
            fontSize: 8.5,
            color: GRAY,
            align: "center",
            italic: true
          });
          ANATOMY_SECTIONS.forEach((sec, idx) => {
            const y = 1.5 + idx * 0.88;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 5.2,
              y,
              w: 4.2,
              h: 0.78,
              rectRadius: 0.08,
              fill: { color: "FFFFFF" },
              line: { color: BORDER_COLOR, width: 0.8 }
            });
            pptxSlide.addText(
              [
                { text: `${sec.label}
`, options: { fontSize: 11, bold: true, color: BURGUNDY } },
                { text: sec.text, options: { fontSize: 9, color: GRAY } }
              ],
              { x: 5.4, y: y + 0.1, w: 3.8, h: 0.58, fontFace: FONT_SANS, margin: 0 }
            );
          });
        }
        break;
      }
      case "patho": {
        if (s.mode === "genes") {
          slide1.background = { color: LIGHT_BG };
          pptxSlide.addText("Key Genetic Drivers:", {
            x: 0.6,
            y: 1.5,
            w: 4.5,
            h: 0.3,
            fontSize: 12,
            bold: true,
            color: BURGUNDY,
            fontFace: FONT_SERIF
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
              line: { color: BURGUNDY, width: 1 }
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
              align: "center",
              valign: "middle"
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
              fill: { color: "FFFFFF" },
              line: { color: BORDER_COLOR, width: 0.8 }
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
              valign: "middle"
            });
          });
        } else {
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y: 1.5,
            w: 4.4,
            h: 3.6,
            rectRadius: 0.1,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
          });
          pptxSlide.addImage({ path: getImg(SCENE_PATHO), x: 0.75, y: 1.65, w: 4.1, h: 3.1 });
          pptxSlide.addText("3D Cellular Lattice Progression", {
            x: 0.6,
            y: 4.8,
            w: 4.4,
            h: 0.25,
            fontSize: 8.5,
            color: GRAY,
            align: "center",
            italic: true
          });
          PATHO_STAGES.forEach((stage, idx) => {
            const y = 1.5 + idx * 0.88;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 5.3,
              y,
              w: 4.1,
              h: 0.75,
              rectRadius: 0.08,
              fill: { color: idx === 3 ? BURGUNDY : "FFFFFF" },
              line: { color: BURGUNDY, width: 1 }
            });
            pptxSlide.addText(
              [
                { text: `Step ${idx + 1}:  `, options: { fontSize: 10, bold: true, color: idx === 3 ? "FFFFFF" : MAUVE } },
                { text: stage, options: { fontSize: 11, bold: true, color: idx === 3 ? "FFFFFF" : BURGUNDY } }
              ],
              { x: 5.5, y: y + 0.1, w: 3.7, h: 0.55, fontFace: FONT_SANS, valign: "middle" }
            );
          });
        }
        break;
      }
      case "cards": {
        const items = s.items;
        if (s.layout === "rows") {
          items.forEach((it, idx) => {
            const y = 1.5 + idx * 0.68;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y,
              w: 8.8,
              h: 0.58,
              rectRadius: 0.06,
              fill: { color: "FFFFFF" },
              line: { color: BORDER_COLOR, width: 0.8 }
            });
            pptxSlide.addText(`\u2022  ${it.title}`, {
              x: 0.9,
              y,
              w: 8.2,
              h: 0.58,
              fontSize: 12,
              bold: true,
              color: BURGUNDY,
              fontFace: FONT_SANS,
              valign: "middle"
            });
          });
        } else {
          const cols = s.columns || (items.length <= 4 ? 2 : 3);
          const colW2 = (8.8 - (cols - 1) * 0.25) / cols;
          const rows = Math.ceil(items.length / cols);
          const rowH = (3.4 - (rows - 1) * 0.25) / rows;
          items.forEach((it, idx) => {
            const col = idx % cols;
            const row = Math.floor(idx / cols);
            const x = 0.6 + col * (colW2 + 0.25);
            const y = 1.5 + row * (rowH + 0.25);
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x,
              y,
              w: colW2,
              h: rowH,
              rectRadius: 0.08,
              fill: { color: "FFFFFF" },
              line: { color: BORDER_COLOR, width: 0.8 }
            });
            if (it.body) {
              pptxSlide.addText(
                [
                  { text: `${it.title}
`, options: { fontSize: 11, bold: true, color: BURGUNDY } },
                  { text: it.body, options: { fontSize: 9, color: GRAY } }
                ],
                { x: x + 0.15, y: y + 0.1, w: colW2 - 0.3, h: rowH - 0.2, fontFace: FONT_SANS, margin: 0 }
              );
            } else {
              pptxSlide.addText(it.title, {
                x: x + 0.15,
                y,
                w: colW2 - 0.3,
                h: rowH,
                fontSize: 11,
                bold: true,
                color: BURGUNDY,
                fontFace: FONT_SANS,
                valign: "middle",
                align: "center"
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
            italic: true
          });
        }
        break;
      }
      case "types": {
        pptxSlide.addShape(pptx.ShapeType.roundRect, {
          x: 0.6,
          y: 1.6,
          w: 8.8,
          h: 2.5,
          rectRadius: 0.12,
          fill: { color: "FFFFFF" },
          line: { color: BORDER_COLOR, width: 1 }
        });
        pptxSlide.addShape(pptx.ShapeType.ellipse, {
          x: 1,
          y: 1.95,
          w: 1.8,
          h: 1.8,
          fill: { color: BURGUNDY }
        });
        pptxSlide.addText("95%", {
          x: 1,
          y: 1.95,
          w: 1.8,
          h: 1.8,
          fontSize: 32,
          bold: true,
          color: "FFFFFF",
          fontFace: FONT_SERIF,
          align: "center",
          valign: "middle"
        });
        pptxSlide.addText(
          [
            { text: `${s.item.title}

`, options: { fontSize: 22, bold: true, color: BURGUNDY, fontFace: FONT_SERIF } },
            { text: s.item.body || "", options: { fontSize: 13, color: INK, fontFace: FONT_SANS, lineSpacing: 18 } }
          ],
          { x: 3.2, y: 1.95, w: 5.8, h: 1.8, fontFace: FONT_SANS, margin: 0 }
        );
        break;
      }
      case "symptoms": {
        s.items.forEach((it, idx) => {
          const y = 1.6 + idx * 1.1;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 8.8,
            h: 0.95,
            rectRadius: 0.08,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
          });
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.85,
            y: y + 0.25,
            w: 1.1,
            h: 0.45,
            rectRadius: 0.05,
            fill: { color: ACCENT_PILL },
            line: { color: BURGUNDY, width: 0.6 }
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
            align: "center",
            valign: "middle"
          });
          pptxSlide.addText(
            [
              { text: `${it.title}
`, options: { fontSize: 13, bold: true, color: BURGUNDY } },
              { text: it.body, options: { fontSize: 10, color: GRAY } }
            ],
            { x: 2.15, y: y + 0.15, w: 7, h: 0.65, fontFace: FONT_SANS, margin: 0 }
          );
        });
        break;
      }
      case "complications": {
        s.items.forEach((it, idx) => {
          const x = 0.6 + idx * 4.55;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x,
            y: 1.6,
            w: 4.25,
            h: 3.2,
            rectRadius: 0.1,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 1 }
          });
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: x + 0.25,
            y: 1.85,
            w: 3.75,
            h: 0.5,
            rectRadius: 0.06,
            fill: { color: ACCENT_PILL }
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
            align: "center",
            valign: "middle"
          });
          pptxSlide.addText(it.body || "", {
            x: x + 0.35,
            y: 2.6,
            w: 3.55,
            h: 2,
            fontSize: 12,
            color: INK,
            fontFace: FONT_SANS,
            lineSpacing: 18,
            align: "center"
          });
        });
        break;
      }
      case "exam": {
        s.steps.forEach((step, idx) => {
          const y = 1.8 + idx * 1.3;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 8.8,
            h: 1.1,
            rectRadius: 0.08,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
          });
          pptxSlide.addShape(pptx.ShapeType.ellipse, {
            x: 0.9,
            y: y + 0.3,
            w: 0.5,
            h: 0.5,
            fill: { color: BURGUNDY }
          });
          pptxSlide.addText(String(idx + 1), {
            x: 0.9,
            y: y + 0.3,
            w: 0.5,
            h: 0.5,
            fontSize: 12,
            bold: true,
            color: "FFFFFF",
            fontFace: FONT_SANS,
            align: "center",
            valign: "middle"
          });
          pptxSlide.addText(step, {
            x: 1.6,
            y: y + 0.15,
            w: 7.5,
            h: 0.8,
            fontSize: 14,
            color: INK,
            fontFace: FONT_SANS,
            valign: "middle"
          });
        });
        break;
      }
      case "tests": {
        DIAGNOSTIC_TESTS.forEach((t, idx) => {
          const y = 1.5 + idx * 1.15;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 5.4,
            h: 1,
            rectRadius: 0.08,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
          });
          pptxSlide.addText(
            [
              { text: `${t.title}
`, options: { fontSize: 13, bold: true, color: BURGUNDY } },
              { text: t.text, options: { fontSize: 10, color: GRAY } }
            ],
            { x: 0.85, y: y + 0.15, w: 4.9, h: 0.7, fontFace: FONT_SANS, margin: 0 }
          );
        });
        pptxSlide.addShape(pptx.ShapeType.roundRect, {
          x: 6.3,
          y: 1.5,
          w: 3.1,
          h: 3.45,
          rectRadius: 0.1,
          fill: { color: "FFFFFF" },
          line: { color: BORDER_COLOR, width: 1 }
        });
        pptxSlide.addImage({ path: getImg(INFOGRAPHIC_DIAG), x: 6.45, y: 1.65, w: 2.8, h: 2.8 });
        pptxSlide.addText("Clinical Diagnostic & Staging Infographic", {
          x: 6.3,
          y: 4.6,
          w: 3.1,
          h: 0.25,
          fontSize: 8,
          color: BURGUNDY,
          bold: true,
          align: "center"
        });
        break;
      }
      case "whyTests": {
        pptxSlide.addShape(pptx.ShapeType.roundRect, {
          x: 1.2,
          y: 2,
          w: 7.6,
          h: 1.8,
          rectRadius: 0.12,
          fill: { color: "FFFFFF" },
          line: { color: BURGUNDY, width: 1.5 }
        });
        pptxSlide.addText(s.text, {
          x: 1.5,
          y: 2.2,
          w: 7,
          h: 1.4,
          fontSize: 18,
          bold: true,
          color: BURGUNDY,
          fontFace: FONT_SERIF,
          align: "center",
          valign: "middle"
        });
        break;
      }
      case "stages": {
        s.items.forEach((it, idx) => {
          const y = 1.6 + idx * 1.05;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 5.4,
            h: 0.9,
            rectRadius: 0.08,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
          });
          pptxSlide.addShape(pptx.ShapeType.ellipse, {
            x: 0.85,
            y: y + 0.2,
            w: 0.5,
            h: 0.5,
            fill: { color: BURGUNDY }
          });
          pptxSlide.addText(it.stage, {
            x: 0.85,
            y: y + 0.2,
            w: 0.5,
            h: 0.5,
            fontSize: 11,
            bold: true,
            color: "FFFFFF",
            fontFace: FONT_SERIF,
            align: "center",
            valign: "middle"
          });
          pptxSlide.addText(
            [
              { text: `Stage ${it.stage}
`, options: { fontSize: 12, bold: true, color: BURGUNDY } },
              { text: it.text, options: { fontSize: 10, color: GRAY } }
            ],
            { x: 1.55, y: y + 0.12, w: 4.2, h: 0.66, fontFace: FONT_SANS, margin: 0 }
          );
        });
        pptxSlide.addShape(pptx.ShapeType.roundRect, {
          x: 6.3,
          y: 1.6,
          w: 3.1,
          h: 3.2,
          rectRadius: 0.1,
          fill: { color: "FFFFFF" },
          line: { color: BORDER_COLOR, width: 1 }
        });
        pptxSlide.addImage({ path: getImg(INFOGRAPHIC_DIAG), x: 6.45, y: 1.75, w: 2.8, h: 2.65 });
        pptxSlide.addText("TNM Staging System Reference", {
          x: 6.3,
          y: 4.5,
          w: 3.1,
          h: 0.25,
          fontSize: 8,
          color: BURGUNDY,
          bold: true,
          align: "center"
        });
        break;
      }
      case "prognosis": {
        if (s.mode === "outlook") {
          s.points.forEach((pt, pIdx) => {
            const y = 1.6 + pIdx * 1;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y,
              w: 5.4,
              h: 0.88,
              rectRadius: 0.08,
              fill: { color: "FFFFFF" },
              line: { color: BORDER_COLOR, width: 0.8 }
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
              valign: "middle"
            });
          });
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 6.3,
            y: 1.6,
            w: 3.1,
            h: 3.2,
            rectRadius: 0.1,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
          });
          pptxSlide.addImage({ path: getImg(SCENE_TUMOR), x: 6.45, y: 1.8, w: 2.8, h: 2.4 });
          pptxSlide.addText("3D Colon Tumor Visualization", {
            x: 6.3,
            y: 4.45,
            w: 3.1,
            h: 0.25,
            fontSize: 8.5,
            color: GRAY,
            align: "center",
            italic: true
          });
        } else {
          s.points.forEach((pt, pIdx) => {
            const y = 1.5 + pIdx * 0.75;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y,
              w: 5.4,
              h: 0.65,
              rectRadius: 0.06,
              fill: { color: "FFFFFF" },
              line: { color: BORDER_COLOR, width: 0.8 }
            });
            pptxSlide.addText(pt, {
              x: 0.8,
              y: y + 0.05,
              w: 5,
              h: 0.55,
              fontSize: 10,
              color: INK,
              fontFace: FONT_SANS,
              bullet: true,
              valign: "middle"
            });
          });
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 6.3,
            y: 1.5,
            w: 3.1,
            h: 2.2,
            rectRadius: 0.08,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
          });
          pptxSlide.addImage({ path: getImg(INFOGRAPHIC_PROG), x: 6.45, y: 1.6, w: 2.8, h: 1.8 });
          pptxSlide.addText("Prognosis & Survival Infographic", {
            x: 6.3,
            y: 3.45,
            w: 3.1,
            h: 0.2,
            fontSize: 7.5,
            color: BURGUNDY,
            bold: true,
            align: "center"
          });
          if (s.takeaway) {
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y: 3.85,
              w: 8.8,
              h: 1.1,
              rectRadius: 0.08,
              fill: { color: ACCENT_PILL },
              line: { color: BURGUNDY, width: 1 }
            });
            pptxSlide.addText(
              [
                { text: `${s.takeawayLabel || "Key Takeaway"}:  `, options: { fontSize: 11, bold: true, color: BURGUNDY } },
                { text: s.takeaway, options: { fontSize: 10, color: INK } }
              ],
              { x: 0.85, y: 3.95, w: 8.3, h: 0.9, fontFace: FONT_SANS, margin: 0, valign: "middle" }
            );
          }
        }
        break;
      }
      case "surgery": {
        pptxSlide.addText(s.lead, {
          x: 0.6,
          y: 1.45,
          w: 8.8,
          h: 0.3,
          fontSize: 11,
          bold: true,
          color: MAUVE,
          fontFace: FONT_SANS
        });
        s.items.forEach((it, idx) => {
          const y = 1.85 + idx * 0.98;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 5.4,
            h: 0.88,
            rectRadius: 0.08,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
          });
          pptxSlide.addText(
            [
              { text: `${it.title}
`, options: { fontSize: 12, bold: true, color: BURGUNDY } },
              { text: it.body || "", options: { fontSize: 10, color: GRAY } }
            ],
            { x: 0.85, y: y + 0.12, w: 4.9, h: 0.64, fontFace: FONT_SANS, margin: 0 }
          );
        });
        pptxSlide.addShape(pptx.ShapeType.roundRect, {
          x: 6.3,
          y: 1.85,
          w: 3.1,
          h: 3,
          rectRadius: 0.1,
          fill: { color: "FFFFFF" },
          line: { color: BORDER_COLOR, width: 1 }
        });
        pptxSlide.addImage({ path: getImg(INFOGRAPHIC_TX), x: 6.45, y: 1.95, w: 2.8, h: 2.5 });
        pptxSlide.addText("Surgical Interventions Overview", {
          x: 6.3,
          y: 4.55,
          w: 3.1,
          h: 0.25,
          fontSize: 8,
          color: BURGUNDY,
          bold: true,
          align: "center"
        });
        break;
      }
      case "additional": {
        pptxSlide.addText(s.lead, {
          x: 0.6,
          y: 1.45,
          w: 8.8,
          h: 0.3,
          fontSize: 11,
          bold: true,
          color: MAUVE,
          fontFace: FONT_SANS
        });
        s.items.forEach((it, idx) => {
          const x = 0.6 + idx * 3;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x,
            y: 1.85,
            w: 2.8,
            h: 0.9,
            rectRadius: 0.08,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
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
            align: "center",
            valign: "middle"
          });
        });
        pptxSlide.addShape(pptx.ShapeType.roundRect, {
          x: 0.6,
          y: 3,
          w: 8.8,
          h: 1.8,
          rectRadius: 0.1,
          fill: { color: ACCENT_PILL },
          line: { color: BURGUNDY, width: 1 }
        });
        pptxSlide.addText(s.seekTitle, {
          x: 0.9,
          y: 3.15,
          w: 8.2,
          h: 0.35,
          fontSize: 14,
          bold: true,
          color: BURGUNDY,
          fontFace: FONT_SERIF
        });
        pptxSlide.addText(s.seekText, {
          x: 0.9,
          y: 3.55,
          w: 8.2,
          h: 1.1,
          fontSize: 12,
          color: INK,
          fontFace: FONT_SANS,
          lineSpacing: 18
        });
        break;
      }
      case "questions": {
        s.items.forEach((q, idx) => {
          const y = 1.7 + idx * 1.1;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 8.8,
            h: 0.9,
            rectRadius: 0.08,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
          });
          pptxSlide.addShape(pptx.ShapeType.ellipse, {
            x: 0.9,
            y: y + 0.2,
            w: 0.5,
            h: 0.5,
            fill: { color: ACCENT_PILL },
            line: { color: BURGUNDY, width: 0.8 }
          });
          pptxSlide.addText("?", {
            x: 0.9,
            y: y + 0.2,
            w: 0.5,
            h: 0.5,
            fontSize: 14,
            bold: true,
            color: BURGUNDY,
            fontFace: FONT_SERIF,
            align: "center",
            valign: "middle"
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
            valign: "middle"
          });
        });
        break;
      }
      case "screening": {
        s.points.forEach((pt, pIdx) => {
          const y = 1.6 + pIdx * 1.1;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 8.8,
            h: 0.95,
            rectRadius: 0.08,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
          });
          pptxSlide.addShape(pptx.ShapeType.ellipse, {
            x: 0.9,
            y: y + 0.22,
            w: 0.5,
            h: 0.5,
            fill: { color: ACCENT_PILL }
          });
          pptxSlide.addText("\u2713", {
            x: 0.9,
            y: y + 0.22,
            w: 0.5,
            h: 0.5,
            fontSize: 14,
            bold: true,
            color: BURGUNDY,
            align: "center",
            valign: "middle"
          });
          pptxSlide.addText(pt, {
            x: 1.6,
            y: y + 0.12,
            w: 7.5,
            h: 0.7,
            fontSize: 12.5,
            color: INK,
            fontFace: FONT_SANS,
            valign: "middle"
          });
        });
        break;
      }
      case "steps": {
        pptxSlide.addText(`Phase: ${s.phase}`, {
          x: 0.6,
          y: 1.45,
          w: 8.8,
          h: 0.28,
          fontSize: 11,
          bold: true,
          color: MAUVE,
          fontFace: FONT_SANS
        });
        s.items.forEach((it, idx) => {
          const y = 1.8 + idx * 0.78;
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y,
            w: 8.8,
            h: 0.7,
            rectRadius: 0.06,
            fill: { color: "FFFFFF" },
            line: { color: BORDER_COLOR, width: 0.8 }
          });
          pptxSlide.addText(
            [
              { text: `${it.title}:  `, options: { fontSize: 11, bold: true, color: BURGUNDY } },
              { text: it.body || "", options: { fontSize: 10, color: INK } }
            ],
            { x: 0.85, y: y + 0.08, w: 8.3, h: 0.54, fontFace: FONT_SANS, margin: 0, valign: "middle" }
          );
        });
        break;
      }
      case "careplan": {
        const plan = CARE_PLANS[s.plan];
        if (s.part === "dx") {
          pptxSlide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6,
            y: 1.5,
            w: 8.8,
            h: 1.4,
            rectRadius: 0.08,
            fill: { color: "FFFFFF" },
            line: { color: BURGUNDY, width: 1.2 }
          });
          pptxSlide.addText(plan.badge.toUpperCase(), {
            x: 0.85,
            y: 1.6,
            w: 8.3,
            h: 0.25,
            fontSize: 8.5,
            bold: true,
            color: MAUVE,
            charSpacing: 1.5
          });
          pptxSlide.addText(plan.diagnosis, {
            x: 0.85,
            y: 1.9,
            w: 8.3,
            h: 0.9,
            fontSize: 12.5,
            bold: true,
            color: BURGUNDY,
            fontFace: FONT_SERIF
          });
          if (plan.subjective || plan.objective) {
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 0.6,
              y: 3.05,
              w: 4.3,
              h: 0.9,
              rectRadius: 0.06,
              fill: { color: "FFFFFF" },
              line: { color: BORDER_COLOR, width: 0.8 }
            });
            pptxSlide.addText(
              [
                { text: "Subjective Data:\n", options: { fontSize: 9, bold: true, color: MAUVE } },
                { text: plan.subjective || "N/A", options: { fontSize: 9.5, color: INK } }
              ],
              { x: 0.75, y: 3.12, w: 4, h: 0.75, fontFace: FONT_SANS, margin: 0 }
            );
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x: 5.1,
              y: 3.05,
              w: 4.3,
              h: 0.9,
              rectRadius: 0.06,
              fill: { color: "FFFFFF" },
              line: { color: BORDER_COLOR, width: 0.8 }
            });
            pptxSlide.addText(
              [
                { text: "Objective Data:\n", options: { fontSize: 9, bold: true, color: MAUVE } },
                { text: plan.objective || "N/A", options: { fontSize: 9.5, color: INK } }
              ],
              { x: 5.25, y: 3.12, w: 4, h: 0.75, fontFace: FONT_SANS, margin: 0 }
            );
          }
          plan.outcomes.forEach((out, idx) => {
            const x = 0.6 + idx * 4.5;
            pptxSlide.addShape(pptx.ShapeType.roundRect, {
              x,
              y: 4.1,
              w: 4.3,
              h: 0.9,
              rectRadius: 0.06,
              fill: { color: ACCENT_PILL },
              line: { color: BORDER_COLOR, width: 0.8 }
            });
            pptxSlide.addText(
              [
                { text: `${out.label}
`, options: { fontSize: 9, bold: true, color: BURGUNDY } },
                { text: out.text, options: { fontSize: 9.5, color: INK } }
              ],
              { x: x + 0.15, y: 4.18, w: 4, h: 0.75, fontFace: FONT_SANS, margin: 0 }
            );
          });
        } else {
          const tableRows = [
            [
              { text: "Nursing Interventions", options: { bold: true, color: "FFFFFF", fill: { color: BURGUNDY }, fontSize: 10, align: "center" } },
              { text: "Scientific & Evidence-Based Rationales", options: { bold: true, color: "FFFFFF", fill: { color: BURGUNDY }, fontSize: 10, align: "center" } }
            ]
          ];
          plan.interventions.forEach((inv) => {
            tableRows.push([
              { text: inv.action, options: { fontSize: 9.5, color: INK, fill: { color: "FFFFFF" }, valign: "middle" } },
              { text: inv.rationale, options: { fontSize: 9.5, color: GRAY, fill: { color: "FFFFFF" }, valign: "middle" } }
            ]);
          });
          pptxSlide.addTable(tableRows, {
            x: 0.6,
            y: 1.5,
            w: 8.8,
            colW: [4.4, 4.4],
            border: { color: BORDER_COLOR, pt: 1 },
            autoPage: false
          });
        }
        break;
      }
      case "pending": {
        pptxSlide.addShape(pptx.ShapeType.roundRect, {
          x: 1.5,
          y: 1.8,
          w: 7,
          h: 2.4,
          rectRadius: 0.15,
          fill: { color: "FFFFFF" },
          line: { color: BORDER_COLOR, width: 1.5, dashType: "dash" }
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
          align: "center"
        });
        pptxSlide.addText(
          "Academic section scheduled in course curriculum.\nContent preserved pending upcoming clinical symposium update.",
          {
            x: 1.8,
            y: 2.65,
            w: 6.4,
            h: 0.9,
            fontSize: 12,
            color: GRAY,
            fontFace: FONT_SANS,
            align: "center",
            lineSpacing: 18
          }
        );
        break;
      }
    }
  }
  return pptx;
}
// Annotate the CommonJS export names for ESM import in node:
0 && (module.exports = {
  createPresentation
});
