import type { ComponentType } from 'react';
import type { Slide, SlideKind } from '../content/slides';
import CoverSlide from '../slides/CoverSlide';
import TeamSlide from '../slides/TeamSlide';
import GoalSlide from '../slides/GoalSlide';
import ObjectivesSlide from '../slides/ObjectivesSlide';
import PendingSlide from '../slides/PendingSlide';
import OutlineSlide from '../slides/OutlineSlide';
import IntroSlide from '../slides/IntroSlide';
import DefinitionSlide from '../slides/DefinitionSlide';
import AnatomySlide from '../slides/AnatomySlide';
import PathoSlide from '../slides/PathoSlide';
import CardsSlide from '../slides/CardsSlide';
import TypesSlide from '../slides/TypesSlide';
import SymptomsSlide from '../slides/SymptomsSlide';
import ComplicationsSlide from '../slides/ComplicationsSlide';
import { ExamSlide, TestsSlide, WhyTestsSlide } from '../slides/DiagnosticSlides';
import StagesSlide from '../slides/StagesSlide';
import PrognosisSlide from '../slides/PrognosisSlide';
import { SurgerySlide, AdditionalSlide, QuestionsSlide } from '../slides/TreatmentSlides';
import ScreeningSlide from '../slides/ScreeningSlide';
import StepsSlide from '../slides/StepsSlide';
import CarePlanSlide from '../slides/CarePlanSlide';

type C = ComponentType<{ slide: Slide }>;
export const REGISTRY = {
  cover: CoverSlide,
  team: TeamSlide,
  goal: GoalSlide,
  objectives: ObjectivesSlide,
  pending: PendingSlide,
  outline: OutlineSlide,
  intro: IntroSlide,
  definition: DefinitionSlide,
  anatomy: AnatomySlide,
  patho: PathoSlide,
  cards: CardsSlide,
  types: TypesSlide,
  symptoms: SymptomsSlide,
  complications: ComplicationsSlide,
  exam: ExamSlide,
  tests: TestsSlide,
  whyTests: WhyTestsSlide,
  stages: StagesSlide,
  prognosis: PrognosisSlide,
  surgery: SurgerySlide,
  additional: AdditionalSlide,
  questions: QuestionsSlide,
  screening: ScreeningSlide,
  steps: StepsSlide,
  careplan: CarePlanSlide,
} as unknown as Record<SlideKind, C>;
