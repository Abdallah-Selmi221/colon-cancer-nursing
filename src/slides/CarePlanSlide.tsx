import type { Slide } from '../content/slides';
import { CARE_PLANS } from '../content/carePlans';

export default function CarePlanSlide({ slide }: { slide: Extract<Slide, { kind: 'careplan' }> }) {
  const plan = CARE_PLANS[slide.plan];
  if (slide.part === 'rx') {
    return (
      <div className="slide-body">
        <p className="label" data-reveal="item">Nursing Interventions &amp; Rationales</p>
        <ol className="rx">
          {plan.interventions.map((r, i) => (
            <li key={r.action} className="rx-row" data-reveal="item">
              <div className="rx-act"><span className="rx-n">{i + 1}</span><div><h3>Intervention</h3><p>{r.action}</p></div></div>
              <svg viewBox="0 0 60 24" className="rx-arrow" aria-hidden><path d="M2 12 H52 M42 4 L54 12 L42 20" fill="none" stroke="#a56f80" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" /></svg>
              <div className="rx-why"><h3>Rationale</h3><p>{r.rationale}</p></div>
            </li>
          ))}
        </ol>
      </div>
    );
  }
  const hasData = plan.subjective || plan.objective;
  return (
    <div className="slide-body cp">
      <p className="badge" data-reveal="item">{plan.badge}</p>
      <div className={`cp-top ${hasData ? '' : 'solo'}`}>
        <section className="dx" data-reveal="item">
          <h3>Nursing Diagnosis</h3>
          <p>{plan.diagnosis}</p>
        </section>
        {hasData && (
          <div className="data">
            {plan.evidencedBy && <p className="evid" data-reveal="item">as evidenced by</p>}
            {plan.subjective && <section className="data-card" data-reveal="item"><h3>Subjective Data</h3><p>{plan.subjective}</p></section>}
            {plan.objective && <section className="data-card" data-reveal="item"><h3>Objective Data</h3><p>{plan.objective}</p></section>}
          </div>
        )}
      </div>
      <section className="outcomes" data-reveal="item">
        <h3>Expected Outcomes</h3>
        <div className="outcome-row">
          {plan.outcomes.map((o) => (
            <div key={o.label} className="outcome"><h4>{o.label}</h4><p>{o.text}</p></div>
          ))}
        </div>
      </section>
    </div>
  );
}
