export interface CarePlan {
  id: 'fatigue' | 'pain' | 'motility';
  title: string;
  badge: string;
  diagnosis: string;
  evidencedBy?: boolean;
  subjective?: string;
  objective?: string;
  outcomes: { label: string; text: string }[];
  interventions: { action: string; rationale: string }[];
}

export const CARE_PLANS: Record<CarePlan['id'], CarePlan> = {
  fatigue: {
    id: 'fatigue',
    title: 'Fatigue',
    badge: 'Actual Nursing Care Plan',
    diagnosis:
      'Fatigue related to decreased oxygen carrying capacity secondary to anemia as evidenced by',
    evidencedBy: true,
    subjective: 'Patient states feeling tired, weak, and dizzy.',
    objective: 'Pallor of skin, low Hb/Hct levels, and decreased activity tolerance.',
    outcomes: [
      { label: 'Short-Term Goal', text: 'Patient will report decreased fatigue within 24–48 hours.' },
      {
        label: 'Long-Term Goal',
        text: 'Patient will perform daily activities with minimal fatigue and demonstrate improved activity tolerance.',
      },
    ],
    interventions: [
      {
        action: 'Assess fatigue level, monitor vital signs regularly.',
        rationale: 'To determine fatigue severity, evaluate anemia, and detect abnormalities early.',
      },
      {
        action: 'Provide adequate rest periods and assist with daily activities as needed.',
        rationale: 'To conserve energy and prevent excessive exhaustion.',
      },
      {
        action: 'Encourage iron-rich foods, administer prescribed iron supplements.',
        rationale: 'To support hemoglobin production, correct iron deficiency.',
      },
    ],
  },
  pain: {
    id: 'pain',
    title: 'Abdominal Pain',
    badge: 'Actual Nursing Care Plan',
    diagnosis:
      'Persistent abdominal pain related to tumor growth and intestinal changes secondary to colon cancer.',
    subjective: '“I have abdominal pain.”',
    objective: 'Abdominal tenderness on palpation.',
    outcomes: [
      { label: 'Short-Term Goal', text: 'Patient will report reduced abdominal pain within 24–48 hours.' },
      { label: 'Long-Term Goal', text: 'Patient will maintain pain control.' },
    ],
    interventions: [
      {
        action: "Assess the patient's pain level, location, characteristics and vital signs.",
        rationale: 'To determine the severity and pattern of pain, and to detect changes.',
      },
      {
        action: 'Administer prescribed analgesics.',
        rationale: 'To relieve pain, reduce muscle tension, and improve patient comfort.',
      },
      {
        action: 'Educate the patient to eat small, frequent, low-fiber meals.',
        rationale:
          'To prevent stomach overloading, minimize bowel irritation, and allow early medical intervention.',
      },
    ],
  },
  motility: {
    id: 'motility',
    title: 'Risk for Dysfunctional Gastrointestinal Motility',
    badge: 'Potential Nursing Care Plan',
    diagnosis:
      'Risk for Dysfunctional Gastrointestinal Motility related to the presence of a colon tumor mass and potential narrowing of the intestinal lumen.',
    outcomes: [{ label: 'Goal', text: 'Patient will maintain stable gastrointestinal function.' }],
    interventions: [
      {
        action: 'Auscultate bowel sounds in all 4 quadrants every 4 hours.',
        rationale: 'To detect early signs of bowel obstruction.',
      },
      {
        action: 'Encourage progressive ambulation and physical activity as tolerated.',
        rationale: 'To stimulate peristalsis and enhance natural GI motility.',
      },
      {
        action: 'Maintain adequate hydration via oral fluids or prescribed IV routes.',
        rationale: 'To prevent constipation and ease stool passage.',
      },
    ],
  },
};
