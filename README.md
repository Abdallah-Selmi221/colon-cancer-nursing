# Colon Cancer: Nursing College Presentation

React + Vite + TypeScript, React Three Fiber, Drei, GSAP, Framer Motion. 45 slides, one at a time.

## Run

```bash
npm install
npm run dev        # http://localhost:5173
npm run typecheck  # TypeScript check
npm run build      # production build into dist/
```

Node 18 or newer. Fonts are bundled (no internet needed at presentation time).

## Controls

| Key | Action |
|---|---|
| → / ↓ / Space / PageDown | Next slide |
| ← / ↑ / PageUp | Previous slide |
| Home / End | First / last slide |
| O | Slide overview (click any slide) |
| F | Full screen (Esc exits) |
| R | Show review flags on slides that have one |

The URL hash follows the slide (`#/12`), so a refresh keeps your place.

## Where things are

```
src/content/slides.ts      ALL slide text, order, and section map (edit text here)
src/content/carePlans.ts   The three nursing care plans
src/engine/                Presentation engine: state, keyboard, stage + GSAP transitions, controls, overview
src/animations/reveal.ts   The single entrance timeline every slide uses (data-reveal attributes)
src/slides/                One component per slide layout
src/three/                 3D scenes: digestive system, colon (selectable sections), cross-section, cell progression
src/illustrations/         SVG medical illustrations (complications, tests, stages, treatment, body)
src/styles/                Tokens/shell (global.css) and slide layouts (slides.css)
```

To fill Goals, Objectives, Summary or References: change that slide in `slides.ts` from `kind: 'pending'` to a layout such as `intro` (points) or `cards`.

## Slide order (45)

1 Cover · 2 Goals (pending) · 3 Objectives (pending) · 4 Outline · 5–6 Introduction · 7 Definition · 8–9 Anatomy and Physiology · 10–11 Pathophysiology · 12 Risk Factors · 13 Causes · 14–15 Types · 16–17 Signs and Symptoms · 18–20 Complications · 21–23 Diagnostic Tests · 24–25 Stages · 26–27 Prognosis · 28–30 Management and Treatment · 31–33 Prevention · 34–37 Nursing Role · 38–43 Nursing Process (Fatigue 38–39, Abdominal Pain 40–41, GI Motility 42–43) · 44 Summary (pending) · 45 References (pending)

Risk Factors comes before Causes, following the academic outline you specified.

## Content notes for your review (press R in the presentation to see them on the slides)

- Causes: source ended with "Unhealthy life"; shown as "Unhealthy lifestyle".
- Risk Factors says "Age over 50"; Prevention says screening starts at 45. Confirm wording.
- Post-operative care: "deep breathing to prevent DVT" kept as supplied; deep breathing is usually linked to lung complications, early ambulation to DVT.
- Prevention: Hormone Therapy statement kept as supplied; confirm against your source.
- Abdominal Pain rationale 3 is kept as supplied; the link between diet advice and "early medical intervention" is unclear.
- Stage III uses your text (lymph nodes). Your infographic also mentions "outer layer".
- The GI Motility plan has no subjective or objective data in the source, so none is shown.
- Slide 11 stage names (Normal colon cells, Abnormal cellular changes, Polyp formation, Cancer development) follow your brief. The 3D cell model is a schematic, not a histology image.
- The Key takeaway on slide 27 comes from your infographic image, not the text file.

## Testing checklist (run locally)

- [ ] `npm install` finishes without errors; `npm run typecheck` and `npm run build` pass
- [ ] `npm run dev` opens; slide 1 shows the 3D digestive system with a slow camera push-in and the title letters rising in
- [ ] Browser console shows no errors on any slide (open DevTools, step through all 45)
- [ ] Right/left arrows, Space, Home, End all work; the counter reads n / 45; "Next:" shows the following slide's title
- [ ] O opens the overview; clicking a card jumps there; Esc closes it
- [ ] F enters full screen; Esc leaves it; controls fade after a pause and return on mouse move
- [ ] Outline (slide 4): every item jumps to the first slide of that section
- [ ] Slide 9: click each of the four sections on the 3D model and in the list; highlight and text stay in sync; drag rotates the model
- [ ] Slide 11: progression plays automatically, the four buttons jump to each stage, "Play progression" restarts it
- [ ] Slides 16–17: hovering or clicking a symptom highlights the right body area
- [ ] Slide 22: tube fills with blood, colonoscopy view advances once
- [ ] Slides 38–43: each care plan reads Diagnosis → data → outcomes, then Intervention → Rationale rows
- [ ] Resize to a laptop (1366×768) and tablet (1024×768) width: no horizontal scroll, text readable
- [ ] Turn on "Reduce motion" in your OS: slides appear without movement
- [ ] Disable hardware acceleration in the browser: flat colon drawings replace the 3D scenes, nothing breaks
