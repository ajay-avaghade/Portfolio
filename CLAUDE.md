# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

## What This Repo Is

A static PM portfolio site for **Ajay Avaghade**, deployed at `https://ajay-avaghade.github.io/Portfolio/`. It also doubles as a job application engine — every company gets a bespoke prototype, CEO pitch deck, resume, and hiring message built directly from this repo.

Application output files (resumes, hiring messages) live in `~/Job Apply AI/` — outside this repo, not committed.

---

## Commands

```bash
# Serve locally
npm start          # http-server on port 3000

# Build a CEO deck (run from repo root)
node assets/build-[company]-deck.js

# Push to GitHub Pages
git add [specific files]
git commit -m "..."
git push origin main
```

No build step, no bundler. The portfolio is plain HTML/CSS/JS — open `index.html` directly in a browser or serve with `npm start`.

---

## Architecture

### Portfolio Site (`index.html` + `styles.css` + `script.js`)

Single-page portfolio. All content is in `index.html` (~1843 lines). `styles.css` (~1377 lines) contains a glassmorphism dark-theme design system. `script.js` (~452 lines) handles:
- Scroll-triggered `.animate-on-scroll` → `.is-visible` transitions via `IntersectionObserver`
- Typewriter effect on hero subtitle
- Mouse-tracking spotlight glow on `.metric-card`, `.case-card`, `.invoice-card`
- Case study modal (`#case-modal`) — content injected from the `caseStudies` JS object in `script.js`
- Sticky cart bar visibility based on scroll position
- Promo code unlock logic
- Hash-based deep-link router at bottom of `index.html` (not in `script.js`) — `#real-world/company-slug` scrolls to `id="rw-[company-slug]"`

**Design tokens** (defined in `styles.css`):
- Font: Inter (body), Outfit (headings)
- Accent: `#7c3aed` (purple), glow: `rgba(124,58,237,0.3)`
- Background: `#0f0f1a`, Surface: `rgba(255,255,255,0.03)`
- Glass panel: `background rgba(255,255,255,0.04)`, `border 1px solid rgba(255,255,255,0.08)`, `backdrop-filter blur(12px)`

### Real World Section Pattern

Each company lives in `index.html` as a `.rw-company-block` with `id="rw-[slug]"`. Add new companies before the `<!-- More coming -->` comment near the end of the Real World section. The deep-link URL format is `/#real-world/[slug]`.

### Prototypes (`[company]-[feature]-prototype.html`)

~60+ standalone HTML files. Each is fully self-contained (no external JS except Google Fonts). Two layout generations exist:

**Current layout** (used in all `/apply` builds): 4 phones side-by-side, 375px each, no annotation panel. Navigation via prev/next arrows + numbered dot tabs below phones. Confetti on screen 3. See `.claude/commands/apply.md` Phase 2 for the full spec.

**Older layout** (legacy prototypes): single 375px phone frame + 300px annotation panel to the right. Still valid for one-off prototypes not generated via `/apply`.

Navigation pattern used in every prototype:
```js
const LABELS = { 1: '01 — SCREEN NAME', ... };
const ANNO   = { 1: [{ n:'INSIGHT 01', t:'Title', b:'Body', metric:{v:'stat', l:'label'} }], ... };

function go(n) {
  [1,2,3,4].forEach(i => { /* hide s{i}, deactivate d{i} nav tab, deactivate n{i} dot */ });
  document.getElementById('s'+n).classList.add('active');
  document.getElementById('screen-label').textContent = LABELS[n];
  renderAnno(n);
  // optionally: if(n===3) startConfetti();
}
function renderAnno(n) { /* injects ANNO[n] cards into #anno-cards */ }
// init
renderAnno(1);
```

Annotation panel renders 3 cards per screen: `INSIGHT 01`, `INSIGHT 02`, `PM RATIONALE` (or `DESIGN DECISION`), each optionally with a `.anno-metric` callout.

**CSS variables** every prototype defines:
```css
:root {
  --dark, --hero, --card, --lift   /* background layers */
  --green/--red/--gold/--purple    /* accent (company-specific) */
  --muted, --border, --white       /* text/border utilities */
}
```

**Known bug**: Never use `.ns` as a classname for content — it collides with the notch sensor dot. Use `.ndot-n` or `.ndot` for the notch dot instead.

### CEO Decks (`assets/build-[company]-deck.js` → `assets/[company]-deck.pptx`)

Built with **PptxGenJS** (`npm install pptxgenjs` — already in `node_modules`). All decks use `LAYOUT_16x9` (10" × 7.5").

Shared conventions across all build scripts:
```js
const C = { dark, hero, accent, gold, white, lgray, muted };  // hex without '#'
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.14 });
// Shadow MUST be a factory function — never reuse the same object across shapes
```

Standard 8-slide structure: Cover (dark) → Problem (dark) → Insight (light) → Mechanic (dark) → Product (light) → Impact & ROI (dark) → Proof of Work (light/split) → Rollout Plan (dark).

**Critical rules**: No `#` prefix on hex colors in PptxGenJS. Apostrophes inside single-quoted JS strings cause `SyntaxError` — rephrase or escape. The `makeShadow()` must be a factory (called fresh each time), not a reused object.

**PptxGenJS async — mandatory pattern** (file won't be written without this):
```js
(async () => {
  const prs = new PptxGenJS();
  // ... build slides ...
  await prs.writeFile({ fileName: 'assets/name.pptx' });
  console.log('Done');
})();
```
Never use `prs.save()` (not a function) or `fs.writeFileSync` on the result of `prs.write()` (returns a Blob, not a Buffer). Always `await prs.writeFile({ fileName: '...' })` inside an async wrapper.

**PptxGenJS shape types** — use string literals only:
```js
slide.addShape('rect', {...})    // ✅
slide.addShape('ellipse', {...}) // ✅
slide.addShape('line', {...})    // ✅
slide.addShape(prs.ShapeType.rect, {...}) // ❌ throws TypeError
```

---

## The `/apply` Workflow

The main Claude Code skill lives at `.claude/commands/apply.md`. When given a JD, it runs 6 phases: JD analysis → prototype → CEO deck → portfolio update → resume → outreach. The skill file is self-evolving — update its Evolution Log and Existing Case Studies table after every application run.

**Source of truth for Ajay's background**: The `## WHO IS AJAY` section at the top of `.claude/commands/apply.md`. Never invent metrics — use only what's listed there.

**Output destinations**:
- Prototypes: `Portfolio/[company]-[feature]-prototype.html`
- Decks: `Portfolio/assets/[company]-[feature]-deck.pptx`
- Build scripts: `Portfolio/assets/build-[company]-[feature]-deck.js`
- Portfolio block: added to `index.html` before `<!-- More coming -->`
- Resumes: `~/Job Apply AI/Ajay_Avaghade_Resume_[Company].html`
- Hiring messages: `~/Job Apply AI/Ajay_Avaghade_[Company]_HiringMessage.html`

**Resume rules** (enforced strictly):
- Single A4/Letter page — `@page { margin:.20in .26in; }`, `font-size: 8.3pt`, `padding: 12px 22px`
- Section order: Name → Summary → Experience → Education → Skills → Case Studies → Projects → Awards
- Every bullet: `text-align:justify; text-align-last:left` to prevent orphan word stretching
- Bold metric phrases must appear mid-sentence, not at the end of a bullet
- Every bullet's last line must have ≥ 6 words
- Case studies section: name + 1 sentence + prototype/deck links only — no metrics in the summary line
- Experience bullets must never reference the case study built for this application
- UAE jobs only: include relocation line; India/other jobs: omit entirely
