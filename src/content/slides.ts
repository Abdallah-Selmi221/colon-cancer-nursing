import type { CarePlan } from './carePlans';

export const SECTIONS = [
  { id: 'cover', label: 'Cover' },
  { id: 'goals', label: 'Goals' },
  { id: 'objectives', label: 'Objectives' },
  { id: 'outline', label: 'Outlines' },
  { id: 'introduction', label: 'Introduction' },
  { id: 'definition', label: 'Definition' },
  { id: 'anatomy', label: 'Anatomy and Physiology' },
  { id: 'pathophysiology', label: 'Pathophysiology' },
  { id: 'risk', label: 'Risk Factors' },
  { id: 'causes', label: 'Causes' },
  { id: 'types', label: 'Types' },
  { id: 'symptoms', label: 'Signs and Symptoms' },
  { id: 'complications', label: 'Complications' },
  { id: 'diagnostics', label: 'Diagnostic Tests' },
  { id: 'prognosis', label: 'Prognosis' },
  { id: 'treatment', label: 'Management and Treatment' },
  { id: 'prevention', label: 'Prevention' },
  { id: 'nursingRole', label: 'Nursing Role' },
  { id: 'nursingProcess', label: 'Nursing Process' },
  { id: 'summary', label: 'Summary' },
  { id: 'references', label: 'References' },
] as const;

export type SectionId = (typeof SECTIONS)[number]['id'];

export type SpotId = 'rectum' | 'lower' | 'body' | 'abdomen' | 'none';
export type ComplicationArtId =
  | 'obstruction' | 'bleeding' | 'perforation' | 'metastasis' | 'malnutrition' | 'emotional';
export type SurgeryArtId = 'polypectomy' | 'colectomy' | 'multivisceral';
export type SectionKey = 'ascending' | 'transverse' | 'descending' | 'sigmoid';
export interface Item { title: string; body?: string }

export const ANATOMY_SECTIONS: { key: SectionKey; label: string; text: string }[] = [
  { key: 'ascending', label: 'Ascending Colon', text: 'Located on the right side; carries contents upward.' },
  { key: 'transverse', label: 'Transverse Colon', text: 'Extends across the upper abdomen from right to left.' },
  { key: 'descending', label: 'Descending Colon', text: 'Runs downward along the left side.' },
  { key: 'sigmoid', label: 'Sigmoid Colon', text: 'S-shaped section connecting the descending colon to the rectum.' },
];

export const PATHO_STAGES = [
  'Normal colon cells',
  'Abnormal cellular changes',
  'Polyp formation',
  'Cancer development',
];

interface Base { id: string; section: SectionId; title: string; flag?: string }
export type Slide = Base &
  (
    | { kind: 'cover'; subtitle: string }
    | { kind: 'team' }
    | { kind: 'goal'; text: string }
    | { kind: 'objectives'; lead: string; items: string[] }
    | { kind: 'pending'; note: string }
    | { kind: 'outline' }
    | { kind: 'intro'; points?: string[]; chipsLabel?: string; chips?: string[]; statement?: string; visual?: boolean }
    | { kind: 'definition'; points: string[] }
    | { kind: 'anatomy'; mode: 'facts' | 'explore'; facts?: string[]; statement?: string }
    | { kind: 'patho'; mode: 'genes' | 'progression'; points?: string[]; genes?: string[] }
    | { kind: 'cards'; layout: 'rows' | 'grid'; columns?: number; items: Item[] }
    | { kind: 'types'; item: Item; percent: number }
    | { kind: 'symptoms'; items: (Item & { spot: SpotId })[] }
    | { kind: 'complications'; items: (Item & { art: ComplicationArtId })[] }
    | { kind: 'exam'; steps: string[] }
    | { kind: 'tests' }
    | { kind: 'whyTests'; text: string }
    | { kind: 'stages'; items: { stage: string; text: string }[] }
    | { kind: 'prognosis'; mode: 'outlook' | 'prognosis'; points: string[]; takeaway?: string; takeawayLabel?: string }
    | { kind: 'surgery'; lead: string; items: (Item & { art: SurgeryArtId })[] }
    | { kind: 'additional'; lead: string; items: string[]; seekTitle: string; seekText: string }
    | { kind: 'questions'; items: string[] }
    | { kind: 'screening'; points: string[] }
    | { kind: 'steps'; phase: string; items: Item[]; art?: 'site' | 'stoma' }
    | { kind: 'careplan'; plan: CarePlan['id']; part: 'dx' | 'rx' }
  );
export type SlideKind = Slide['kind'];

const PENDING = 'Academic section scheduled in course curriculum.';

export const slides: Slide[] = [
  { id: 'cover', section: 'cover', kind: 'cover', title: 'Colon Cancer', subtitle: 'Nursing College Presentation' },
  { id: 'team', section: 'cover', kind: 'team', title: 'Presentation Team' },
  {
    id: 'goals',
    section: 'goals',
    kind: 'goal',
    title: 'GOAL',
    text: 'The goal of this assignment is to provide nursing students with a comprehensive understanding of colon cancer, including its anatomy, pathophysiology, risk factors, clinical manifestations, diagnosis, and management, in order to enhance their ability to deliver safe, evidence-based nursing care to patients with colon cancer.',
  },
  {
    id: 'objectives',
    section: 'objectives',
    kind: 'objectives',
    title: 'OBJECTIVES',
    lead: 'By the end of this assignment, the student will be able to:',
    items: [
      'Define colon cancer and describe the anatomy and physiology of the colon.',
      'Explain the pathophysiology, causes, and risk factors of colon cancer.',
      'Identify the types, signs and symptoms, and complications of colon cancer.',
      'Discuss the diagnostic tests, staging, and prognosis of colon cancer.',
      'Outline the medical and surgical management and the preventive measures of colon cancer.',
      'Describe the role of the nurse in caring for patients with colon cancer before and after surgery.',
      'Apply the nursing process to a patient with colon cancer.',
    ],
  },
  { id: 'outline', section: 'outline', kind: 'outline', title: 'Outline' },
  {
    id: 'intro-1', section: 'introduction', kind: 'intro', title: 'Introduction', visual: true,
    points: [
      'Colon cancer is a major cancer of the digestive system.',
      'It starts when colon lining cells grow abnormally and uncontrollably.',
      'Often begins as small polyps that may turn cancerous over several years.',
    ],
  },
  {
    id: 'intro-2', section: 'introduction', kind: 'intro', title: 'Introduction: Risk Factors',
    chipsLabel: 'Risk factors',
    chips: ['Older age', 'Family history', 'Inherited genetic conditions', 'Inflammatory bowel disease', 'Smoking', 'Lifestyle'],
    statement: 'It can also occur without any known risk factors.',
  },
  {
    id: 'definition', section: 'definition', kind: 'definition', title: 'Definition',
    points: [
      'A malignant disease of the colon, the longest part of the large intestine.',
      'Abnormal cells divide uncontrollably and form a malignant tumor.',
      'Most cases are adenocarcinomas, arising from glandular cells lining the colon.',
    ],
  },
  {
    id: 'anatomy-1', section: 'anatomy', kind: 'anatomy', mode: 'facts', title: 'Anatomy of the Colon',
    facts: [
      'The colon is the main part of the large intestine.',
      'It is approximately 1.5 meters long.',
      'Its main functions are absorbing water and electrolytes and forming solid stool.',
    ],
    statement: 'The colon has four main sections.',
  },
  { id: 'anatomy-2', section: 'anatomy', kind: 'anatomy', mode: 'explore', title: 'The Four Sections of the Colon' },
  {
    id: 'patho-1', section: 'pathophysiology', kind: 'patho', mode: 'genes', title: 'Pathophysiology',
    genes: ['APC', 'KRAS', 'p53'],
    points: [
      'Genetic mutations (APC, KRAS, p53).',
      'Abnormal cells develop into polyps then cancer.',
      'DNA repair defects cause uncontrolled cell growth.',
    ],
  },
  { id: 'patho-2', section: 'pathophysiology', kind: 'patho', mode: 'progression', title: 'Pathophysiology: From Normal Cells to Cancer' },
  {
    id: 'risk', section: 'risk', kind: 'cards', layout: 'grid', title: 'Risk Factors',
    items: [
      { title: 'Age over 50' }, { title: 'Family history' }, { title: 'Obesity' },
      { title: 'Smoking and alcohol' }, { title: 'Red and processed meat' },
      { title: 'Physical inactivity' }, { title: 'Diabetes' },
    ],
    flag: 'Age over 50 is listed here, while the Prevention slides note screening begins at age 45.',
  },
  {
    id: 'causes', section: 'causes', kind: 'cards', layout: 'rows', title: 'Causes',
    items: [
      { title: 'Genetic mutations' }, { title: 'Hereditary syndromes' }, { title: 'Family history' },
      { title: 'Chronic intestinal inflammation' }, { title: 'Unhealthy lifestyle' },
    ],
  },
  {
    id: 'types-1', section: 'types', kind: 'types', title: 'Types of Cancer in the Colon and Rectum', percent: 95,
    item: { title: 'Adenocarcinoma', body: 'The most common type (about 95%), develops in mucus-producing cells lining the colon and rectum.' },
  },
  {
    id: 'types-2', section: 'types', kind: 'cards', layout: 'grid', columns: 2, title: 'Types: Other Cancers of the Colon and Rectum',
    items: [
      { title: 'Carcinoid Tumors', body: 'Develop from hormone-producing cells in the intestine.' },
      { title: 'Gastrointestinal Stromal Tumors (GISTs)', body: 'Develop from specialized cells in the digestive tract wall.' },
      { title: 'Lymphomas', body: 'Cancers that develop from immune system cells and may affect the colon and rectum.' },
      { title: 'Sarcomas', body: 'Rare cancers that develop in muscles, blood vessels, or connective tissues of the intestinal wall.' },
    ],
  },
  {
    id: 'symptoms-1', section: 'symptoms', kind: 'symptoms', title: 'Signs and Symptoms of Colon Cancer',
    items: [
      { title: 'Blood in Stool', body: 'Blood may appear in the stool or as rectal bleeding.', spot: 'rectum' },
      { title: 'Change in Stool Appearance', body: 'Changes in stool color, including dark or black stools.', spot: 'lower' },
      { title: 'Fatigue and Weakness', body: 'May result from chronic bleeding and iron-deficiency anemia.', spot: 'body' },
    ],
  },
  {
    id: 'symptoms-2', section: 'symptoms', kind: 'symptoms', title: 'Signs and Symptoms: Continued',
    items: [
      { title: 'Unexplained Weight Loss', body: 'Sudden weight loss without an obvious reason.', spot: 'body' },
      { title: 'Abdominal Bloating and Bowel Changes', body: 'Persistent bloating or feeling that the bowel has not completely emptied.', spot: 'abdomen' },
      { title: 'No Symptoms', body: 'Some patients may have no symptoms, especially in early stages.', spot: 'none' },
    ],
  },
  {
    id: 'compl-1', section: 'complications', kind: 'complications', title: 'Complications of Colon Cancer',
    items: [
      { title: 'Bowel Obstruction', art: 'obstruction', body: 'Tumor growth blocks the intestine, causing abdominal pain, swelling, nausea, and vomiting.' },
      { title: 'Gastrointestinal Bleeding', art: 'bleeding', body: 'Tumor bleeding may cause anemia, fatigue, and bloody or black stools.' },
    ],
  },
  {
    id: 'compl-2', section: 'complications', kind: 'complications', title: 'Complications: Continued',
    items: [
      { title: 'Perforation', art: 'perforation', body: 'A hole in the intestinal wall can lead to peritonitis and serious infection.' },
      { title: 'Metastasis', art: 'metastasis', body: 'Cancer spreads to nearby lymph nodes or distant organs, especially the liver and lungs.' },
    ],
  },
  {
    id: 'compl-3', section: 'complications', kind: 'complications', title: 'Complications: Continued',
    items: [
      { title: 'Malnutrition and Weight Loss', art: 'malnutrition', body: 'Reduced food intake and treatment side effects may cause weight loss and nutritional deficiencies.' },
      { title: 'Emotional and Psychological Impact', art: 'emotional', body: 'Cancer diagnosis and treatment may cause anxiety, stress, and depression.' },
    ],
  },
  {
    id: 'dx-exam', section: 'diagnostics', kind: 'exam', title: 'Clinical Examination',
    steps: ['The doctor asks about symptoms and medical history.', 'Examines the abdomen and may perform a digital rectal examination.'],
  },
  { id: 'dx-tests', section: 'diagnostics', kind: 'tests', title: 'Diagnostic Tests' },
  { id: 'dx-why', section: 'diagnostics', kind: 'whyTests', title: 'Why These Tests?', text: 'These tests help doctors diagnose cancer and determine its stage.' },
  {
    id: 'stages-1', section: 'diagnostics', kind: 'stages', title: 'Stages of Colon Cancer',
    items: [
      { stage: '0', text: 'Abnormal or precancerous cells.' },
      { stage: 'I', text: 'Cancer is in the colon wall.' },
      { stage: 'II', text: 'Cancer spreads deeper through the colon wall.' },
    ],
  },
  {
    id: 'stages-2', section: 'diagnostics', kind: 'stages', title: 'Stages of Colon Cancer: Continued',
    items: [
      { stage: 'III', text: 'Cancer spreads to nearby lymph nodes.' },
      { stage: 'IV', text: 'Cancer spreads to distant organs such as the liver or lungs.' },
    ],
  },
  {
    id: 'prog-1', section: 'prognosis', kind: 'prognosis', mode: 'outlook', title: 'Outlook of Colon Cancer',
    points: [
      "The outlook depends on the cancer stage, patient's general health, and response to treatment.",
      'Early diagnosis and treatment improve the chances of recovery.',
      'Treatment may include surgery, chemotherapy, and radiation therapy.',
    ],
  },
  {
    id: 'prog-2', section: 'prognosis', kind: 'prognosis', mode: 'prognosis', title: 'Prognosis of Colon Cancer',
    points: [
      'Prognosis depends on how early cancer is detected and whether it has spread.',
      'Early-stage cancer has a better prognosis, while advanced cancer has a poorer prognosis.',
      'Regular follow-up after treatment helps detect recurrence early.',
    ],
    takeawayLabel: 'Key takeaway',
    takeaway: "Regular follow-up after treatment helps detect cancer recurrence early. The chances of recovery also depend on the patient's overall health and response to treatment.",
  },
  {
    id: 'tx-1', section: 'treatment', kind: 'surgery', title: 'Management and Treatment',
    lead: "Surgery is the main treatment for colon cancer that hasn't spread.",
    items: [
      { title: 'Polypectomy', art: 'polypectomy', body: 'Removal of cancerous or precancerous polyps.' },
      { title: 'Colectomy', art: 'colectomy', body: 'Removal of the affected part of the colon.' },
      { title: 'Multivisceral Resection', art: 'multivisceral', body: 'Removal of the tumor and nearby affected organs.' },
    ],
  },
  {
    id: 'tx-2', section: 'treatment', kind: 'additional', title: 'Additional Treatment and When to Seek Medical Care',
    lead: 'Chemotherapy, radiation therapy and targeted therapy may be used before or after surgery.',
    items: ['Chemotherapy', 'Radiation therapy', 'Targeted therapy'],
    seekTitle: 'When to Seek Medical Care',
    seekText: 'Consult a doctor if you have persistent abdominal pain or blood in your stool, as these symptoms may have different causes and require medical evaluation.',
  },
  {
    id: 'tx-3', section: 'treatment', kind: 'questions', title: 'Important Questions',
    items: ['What is the cancer stage?', 'Can treatment eliminate the cancer?', 'Can the cancer return?'],
  },
  {
    id: 'prev-1', section: 'prevention', kind: 'screening', title: 'Prevention: Screening & Early Detection',
    points: [
      'Regular screening is one of the most powerful tools for preventing colorectal cancer',
      'It takes 10 to 15 years for polyps to turn into cancer, screening allows removal before they develop',
      'Start screening at age 45 (or earlier if strong family history/genetic risk)',
    ],
  },
  {
    id: 'prev-2', section: 'prevention', kind: 'cards', layout: 'grid', columns: 2, title: 'Lifestyle & Risk Reduction',
    items: [
      { title: 'Healthy Weight', body: 'Maintaining healthy weight lowers risk' },
      { title: 'Physical Activity', body: 'Regular moderate to vigorous activity lowers risk of polyps and cancer' },
      { title: 'Diet', body: 'Eat more vegetables, fruits, and whole grains. Limit red meats and processed meats.' },
      { title: 'Avoid Alcohol & Smoking', body: 'Quitting reduces colorectal cancer risk' },
    ],
  },
  {
    id: 'prev-3', section: 'prevention', kind: 'cards', layout: 'grid', columns: 3, title: 'Supplements & Medical Factors',
    items: [
      { title: 'Vitamins & Minerals', body: 'Folic acid, Vitamin D, Calcium, and Magnesium may help lower risk' },
      { title: 'NSAIDs (Aspirin)', body: 'May lower risk, but not recommended for prevention due to stomach bleeding' },
      { title: 'Hormone Therapy', body: 'May reduce risk in women, but carries other health risks' },
    ],
  },
  {
    id: 'nr-1', section: 'nursingRole', kind: 'steps', phase: 'Assessment and pre-operative care', art: 'site', title: 'Nursing Role: Assessment & Pre-Operative Care',
    items: [
      { title: 'Comprehensive Assessment', body: 'Changes in bowel habits, blood in stool, unexplained anemia, abdominal discomfort, nutritional status, family history' },
      { title: 'Bowel Preparation', body: 'Laxatives, enemas, or oral bowel cleansers' },
      { title: 'Patient Education & Psychological Support', body: 'Explain procedure, address fears and anxiety' },
      { title: 'Stoma Site Marking', body: 'Mark optimal site with ostomy specialist (WOCN)' },
    ],
  },
  {
    id: 'nr-2', section: 'nursingRole', kind: 'steps', phase: 'Post-operative care', art: 'stoma', title: 'Post-Operative Nursing Care',
    items: [
      { title: 'Monitoring', body: 'Vital signs, abdominal pain, bowel sounds, return of intestinal function' },
      { title: 'Stoma Care', body: 'Stoma should be pink/red and moist, teach pouch change and skin protection' },
      { title: 'Prevent Complications', body: 'Pain medication, early ambulation, deep breathing to prevent DVT' },
    ],
  },
  {
    id: 'nr-3', section: 'nursingRole', kind: 'steps', phase: 'Treatment side effects', title: 'Managing Treatment Side Effects',
    items: [
      { title: 'Chemotherapy & Radiation Support', body: 'Manage nausea, vomiting, fatigue, diarrhea with medications and fluids to prevent dehydration' },
      { title: 'Infection & Lab Monitoring', body: 'Monitor WBCs and platelets for infection and bleeding risks' },
    ],
  },
  {
    id: 'nr-4', section: 'nursingRole', kind: 'steps', phase: 'Psychological and rehabilitation support', title: 'Psychological & Rehabilitation Support',
    items: [
      { title: 'Body Image & Coping', body: 'Help patient adapt to bodily changes with colostomy' },
      { title: 'Holistic Support', body: 'Continuous emotional support for patient and family' },
    ],
  },
  { id: 'cp-f1', section: 'nursingProcess', kind: 'careplan', plan: 'fatigue', part: 'dx', title: 'Nursing Care Plan: Fatigue' },
  { id: 'cp-f2', section: 'nursingProcess', kind: 'careplan', plan: 'fatigue', part: 'rx', title: 'Fatigue: Interventions & Rationales' },
  { id: 'cp-p1', section: 'nursingProcess', kind: 'careplan', plan: 'pain', part: 'dx', title: 'Nursing Care Plan: Abdominal Pain' },
  { id: 'cp-p2', section: 'nursingProcess', kind: 'careplan', plan: 'pain', part: 'rx', title: 'Abdominal Pain: Interventions & Rationales' },
  { id: 'cp-m1', section: 'nursingProcess', kind: 'careplan', plan: 'motility', part: 'dx', title: 'Nursing Care Plan: Risk for Dysfunctional Gastrointestinal Motility' },
  { id: 'cp-m2', section: 'nursingProcess', kind: 'careplan', plan: 'motility', part: 'rx', title: 'Gastrointestinal Motility: Interventions & Rationales' },
  { id: 'summary', section: 'summary', kind: 'pending', title: 'Summary', note: PENDING },
  { id: 'references', section: 'references', kind: 'pending', title: 'References', note: PENDING },
];

export const firstIndexOf = (section: SectionId) => slides.findIndex((s) => s.section === section);

export const DIAGNOSTIC_TESTS = [
  { id: 'blood', title: 'Blood Tests', text: 'CBC, CMP and CEA.', chips: ['CBC', 'CMP', 'CEA'] },
  { id: 'imaging', title: 'Imaging Tests', text: 'CT scan, MRI, PET scan and abdominal ultrasound.', chips: ['CT scan', 'MRI', 'PET scan', 'Abdominal ultrasound'] },
  { id: 'colonoscopy', title: 'Colonoscopy', text: 'Allows the doctor to examine the colon and take a biopsy for laboratory examination.', chips: [] as string[] },
] as const;

export const STAGE_RAIL = ['0', 'I', 'II', 'III', 'IV'];
