const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0D1117',
  hero:   '161B22',
  card:   '1C2333',
  bright: '34D399',
  accent: '10B981',
  purple: '818CF8',
  orange: 'F97316',
  red:    'EF4444',
  white:  'F0F6FC',
  lgray:  '8B949E',
  border: '30363D',
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 4, offset: 2, angle: 45, opacity: 0.15 });
const W = 10, H = 5.625;

// ── SLIDE 1 — COVER ──────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  for (let i = 0; i < 5; i++) {
    const lx = -0.5 + i * 2.2;
    const lw = Math.min(3.5, W - lx - 0.15);
    if (lx > W || lw <= 0) continue;
    s.addShape(prs.ShapeType.line, { x: lx, y: 0, w: lw, h: H, line: { color: C.bright, width: 0.4 }, rotate: 30, transparency: 88 });
  }
  s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 0.06, h: H, fill: { color: C.bright } });

  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 0.35, w: 1.8, h: 0.28, fill: { color: C.hero }, line: { color: C.bright, width: 0.8 }, rectRadius: 0.06 });
  s.addText('DEEL · EQUITY PRODUCT', { x: 0.5, y: 0.35, w: 1.8, h: 0.28, color: C.bright, fontSize: 7, bold: true, align: 'center', valign: 'middle', charSpacing: 1.5 });

  s.addText('EquityOnboard', { x: 0.5, y: 0.8, w: 6, h: 0.9, color: C.white, fontSize: 44, bold: true, fontFace: 'Calibri' });
  s.addText('AI-Guided Equity Expansion to New Geographies', { x: 0.5, y: 1.65, w: 6.5, h: 0.42, color: C.bright, fontSize: 18, bold: true });
  s.addText('Compliance gap analysis + country sub-plan scaffold in hours, not weeks', { x: 0.5, y: 2.1, w: 6.5, h: 0.35, color: C.lgray, fontSize: 12 });
  s.addText('Presented by Ajay Avaghade · Senior PM, Equity', { x: 0.5, y: 2.6, w: 5, h: 0.28, color: C.lgray, fontSize: 10 });

  s.addShape(prs.ShapeType.roundRect, { x: 7.2, y: 1.6, w: 2.4, h: 1.6, fill: { color: C.hero }, line: { color: C.bright, width: 1 }, rectRadius: 0.1, shadow: makeShadow() });
  s.addText('3 wks → 4 hrs', { x: 7.2, y: 1.75, w: 2.4, h: 0.55, color: C.bright, fontSize: 20, bold: true, align: 'center', fontFace: 'Courier New' });
  s.addText('new country\nequity expansion', { x: 7.2, y: 2.3, w: 2.4, h: 0.55, color: C.lgray, fontSize: 9, align: 'center', valign: 'middle' });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · EquityOnboard · Confidential', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
  s.addText('ajay-avaghade.github.io/Portfolio', { x: 5.5, y: 5.38, w: 4.1, h: 0.245, color: C.bright, fontSize: 8, align: 'right', valign: 'middle' });
}

// ── SLIDE 2 — THE PROBLEM ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE PROBLEM', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('Every New Country Is a 3-Week Compliance Research Project — Blocking Equity Expansion', { x: 0.4, y: 0.44, w: 9.2, h: 0.65, color: C.white, fontSize: 21, bold: true });

  const cols = [
    { icon: '🌍', stat: '150+', label: 'Countries where Deel operates', sub: 'Equity offered in a fraction — each new country requires securities law, tax treatment, and sub-plan research' },
    { icon: '⏳', stat: '3 weeks', label: 'Manual compliance research per country', sub: 'Compliance analyst reads securities law, maps to equity plan, drafts sub-plan, routes for legal review' },
    { icon: '❌', stat: '60%', label: 'Expansion requests blocked or delayed', sub: 'Companies drop equity expansion plans because ops complexity is too high — costing Deel upsell opportunities' },
  ];
  cols.forEach((col, i) => {
    const x = 0.4 + i * 3.1;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.2, w: 2.9, h: 2.4, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1, shadow: makeShadow() });
    s.addText(col.icon, { x, y: 1.32, w: 2.9, h: 0.4, fontSize: 22, align: 'center' });
    s.addText(col.stat, { x, y: 1.7, w: 2.9, h: 0.55, color: C.bright, fontSize: 28, bold: true, align: 'center', fontFace: 'Courier New' });
    s.addText(col.label, { x: x + 0.1, y: 2.25, w: 2.7, h: 0.3, color: C.white, fontSize: 10, bold: true, align: 'center' });
    s.addText(col.sub, { x: x + 0.1, y: 2.55, w: 2.7, h: 0.7, color: C.lgray, fontSize: 8, align: 'center', valign: 'top' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 3.75, w: 9.2, h: 0.65, fill: { color: C.hero }, line: { color: C.bright, width: 0.8 }, rectRadius: 0.08 });
  s.addText('Root cause: Expanding equity to a new country requires reading securities law, tax treatment, sub-plan drafting, and legal review — all done manually, from scratch, every time. The compliance knowledge exists inside Deel\'s team but is not productised.', {
    x: 0.6, y: 3.78, w: 8.8, h: 0.6, color: C.white, fontSize: 9, valign: 'middle'
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · EquityOnboard', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
}

// ── SLIDE 3 — THE INSIGHT ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: 'FFFFFF' };

  s.addText('THE INSIGHT', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: '777777', fontSize: 8, bold: true, charSpacing: 2 });
  s.addText("Deel Already Has the Compliance Knowledge. EquityOnboard Productises It.", { x: 0.4, y: 0.44, w: 9.2, h: 0.55, color: '111111', fontSize: 21, bold: true });

  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 1.1, w: 4.1, h: 3.4, fill: { color: 'FFF5F5' }, line: { color: 'FECACA', width: 0.8 }, rectRadius: 0.1 });
  s.addText('❌  Manual Process (Today)', { x: 0.6, y: 1.22, w: 3.7, h: 0.3, color: 'DC2626', fontSize: 11, bold: true });
  const leftPts = [
    'Compliance analyst manually researches securities law for the target country from scratch',
    'Tax treatment mapped by hand — Income Tax Act, withholding rules, perquisite definitions',
    'Sub-plan drafted in Word doc, routed via email for legal review (2-3 rounds typical)',
    'Companies wait 3+ weeks to learn if equity expansion is even feasible',
    'Same research happens again for every subsequent expansion — no reuse',
  ];
  leftPts.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 0.6, y: 1.6 + i * 0.44, w: 3.7, h: 0.38, color: '333333', fontSize: 9, valign: 'top' });
  });

  s.addShape(prs.ShapeType.ellipse, { x: 4.55, y: 2.4, w: 0.6, h: 0.45, fill: { color: C.dark }, line: { color: C.dark, width: 0 } });
  s.addText('VS', { x: 4.55, y: 2.4, w: 0.6, h: 0.45, color: C.white, fontSize: 9, bold: true, align: 'center', valign: 'middle' });

  s.addShape(prs.ShapeType.roundRect, { x: 5.2, y: 1.1, w: 4.4, h: 3.4, fill: { color: 'F0FDF9' }, line: { color: '6EE7B7', width: 0.8 }, rectRadius: 0.1 });
  s.addText('✅  EquityOnboard (Proposed)', { x: 5.4, y: 1.22, w: 4.0, h: 0.3, color: '059669', fontSize: 11, bold: true });
  const rightPts = [
    'AI engine queries Deel\'s codified country compliance ruleset (built once, updated continuously)',
    'Gap analysis in minutes: AI maps company\'s equity plan against country requirements, flags exact gaps',
    'Sub-plan scaffold auto-generated: compliant sections pre-drafted, action items for legal clearly marked',
    'Legal team only reviews flagged gaps — not full document from scratch',
    'Knowledge compounds: each new country enriches the ruleset for all future expansions',
  ];
  rightPts.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 5.4, y: 1.6 + i * 0.44, w: 4.0, h: 0.38, color: '333333', fontSize: 9, valign: 'top' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 4.62, w: 9.2, h: 0.65, fill: { color: C.dark }, line: { color: C.dark, width: 0 }, rectRadius: 0.08 });
  s.addText("Key insight: Deel's compliance team already knows every country's equity rules. EquityOnboard encodes that knowledge into a queryable ruleset — so it runs in 4 hours instead of 3 weeks, and gets better with every new country added.", {
    x: 0.6, y: 4.65, w: 8.8, h: 0.6, color: C.white, fontSize: 9, valign: 'middle'
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: 'F3F4F6' } });
  s.addText('Deel · EquityOnboard', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: '777777', fontSize: 8, valign: 'middle' });
}

// ── SLIDE 4 — THE MECHANIC ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE MECHANIC', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('4 Hours From Decision to Legal-Ready Sub-Plan', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: C.white, fontSize: 24, bold: true });

  const steps = [
    { n: '1', title: 'Country Selection', body: 'HR admin selects target country from expansion picker. AI surfaces complexity rating (Low/Medium/High), estimated gap count, and time-to-launch estimate.' },
    { n: '2', title: 'AI Gap Analysis', body: 'Engine maps company\'s equity plan type (RSU / ESOP / SAR) against Deel country compliance ruleset. Outputs: met requirements, gap count, blocker vs. advisory flags.' },
    { n: '3', title: 'Sub-Plan Scaffold', body: 'AI generates draft sub-plan sections: grant eligibility, vesting rules, tax withholding mechanics, exercise/sale procedures. Compliant sections pre-drafted. Gaps marked for legal.' },
    { n: '4', title: 'Legal Routing', body: 'Draft sub-plan sent to Deel\'s compliance team with flagged action items pre-highlighted. Legal reviews gaps only (not full document). Average review: 2–4 hours.' },
    { n: '5', title: 'Live + Reusable', body: 'Approved sub-plan stored in Deel\'s country ruleset for future expansions. Next company expanding to same country gets 80% pre-filled. Knowledge compounds over time.' },
  ];

  steps.forEach((step, i) => {
    const x = 0.38 + i * 1.88;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.05, w: 1.72, h: 3.4, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1, shadow: makeShadow() });
    s.addShape(prs.ShapeType.ellipse, { x: x + 0.62, y: 1.12, w: 0.46, h: 0.38, fill: { color: C.bright }, line: { color: C.bright, width: 0 } });
    s.addText(step.n, { x: x + 0.62, y: 1.12, w: 0.46, h: 0.38, color: C.dark, fontSize: 12, bold: true, align: 'center', valign: 'middle' });
    s.addText(step.title, { x: x + 0.06, y: 1.58, w: 1.6, h: 0.38, color: C.white, fontSize: 10, bold: true, align: 'center' });
    s.addText(step.body, { x: x + 0.08, y: 2.0, w: 1.56, h: 1.9, color: C.lgray, fontSize: 8, align: 'center', valign: 'top' });
    if (i < 4) s.addShape(prs.ShapeType.line, { x: x + 1.72, y: 2.2, w: 0.16, h: 0, line: { color: C.bright, width: 1.2 } });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.38, y: 4.56, w: 9.24, h: 0.64, fill: { color: C.hero }, line: { color: C.bright, width: 0.8 }, rectRadius: 0.08 });
  s.addText('PhonePe proof: Built Kotak Cherry WealthTech (Founding PM) — zero-to-one product launch with MVP scoping, UX research, Compliance + Legal co-ownership, and GTM execution → 100K+ downloads in launch window. Direct parallel: regulated 0→1 product delivery.', {
    x: 0.56, y: 4.58, w: 9.0, h: 0.6, color: C.white, fontSize: 8.5, valign: 'middle'
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · EquityOnboard', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
}

// ── SLIDE 5 — THE PRODUCT ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: 'FFFFFF' };

  s.addText('THE PRODUCT', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: '777777', fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('4 Key Screens — Select, Analyse, Generate, Approve', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: '111111', fontSize: 22, bold: true });

  const cards = [
    { n: '01', title: 'Country Picker', body: 'Country cards with complexity badges (Low/Medium/High), estimated gap count, and time-to-launch. Recommended next markets surfaced based on demand signals from Deel sales data.' },
    { n: '02', title: 'AI Gap Analysis', body: 'Progress bar (6/9 requirements met). Met requirements shown in green. Gaps colour-coded: Blocker (red) vs Advisory (orange). One-tap to generate sub-plan from gaps found.' },
    { n: '03', title: 'Sub-Plan Ready', body: 'Draft scaffold sections listed: Grant Eligibility, Vesting Rules, Tax Mechanics — all compliant sections pre-drafted. 2 action items flagged for legal in orange. Download PDF or send to legal queue.' },
    { n: '04', title: 'Approval Queue', body: 'Compliance team\'s work queue: India (In Review, assigned), Brazil (Queued), Nigeria (Queued). Stats: 3 awaiting, 5 approved this month, 4hr avg review time.' },
  ];

  cards.forEach((card, i) => {
    const x = 0.38 + i * 2.38;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.08, w: 2.2, h: 3.55, fill: { color: 'F9FAFB' }, line: { color: 'E5E7EB', width: 0.6 }, rectRadius: 0.1, shadow: makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y: 1.08, w: 2.2, h: 0.06, fill: { color: C.bright }, line: { color: C.bright, width: 0 } });
    s.addText(card.n, { x, y: 1.16, w: 2.2, h: 0.35, color: C.bright, fontSize: 20, bold: true, align: 'center', fontFace: 'Courier New' });
    s.addText(card.title, { x: x + 0.1, y: 1.5, w: 2.0, h: 0.34, color: '111111', fontSize: 11, bold: true, align: 'center' });
    s.addText(card.body, { x: x + 0.1, y: 1.88, w: 2.0, h: 2.4, color: '555555', fontSize: 8.5, align: 'left', valign: 'top' });
  });

  s.addText('Interactive prototype: deel-equityonboard-prototype.html', { x: 0.4, y: 4.74, w: 9.2, h: 0.3, color: '777777', fontSize: 9, align: 'center' });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: 'F3F4F6' } });
  s.addText('Deel · EquityOnboard', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: '777777', fontSize: 8, valign: 'middle' });
}

// ── SLIDE 6 — IMPACT & ROI ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('IMPACT & ROI', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('Every New Country Unlocked Is a New Revenue Line for Deel', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: C.white, fontSize: 22, bold: true });

  s.addShape(prs.ShapeType.roundRect, { x: 0.35, y: 1.05, w: 4.55, h: 3.55, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1 });
  s.addText('SPEED & EFFICIENCY', { x: 0.5, y: 1.15, w: 4.2, h: 0.22, color: C.bright, fontSize: 8, bold: true, charSpacing: 1.5 });
  const leftMetrics = [
    { val: '3 wks → 4 hrs', lbl: 'New country equity expansion time' },
    { val: '80%', lbl: 'Pre-filled sub-plan scaffold for legal review' },
    { val: '0 rework', lbl: 'Compliance knowledge reused across all customers' },
    { val: '5× faster', lbl: 'Legal review time (gaps only, not full document)' },
  ];
  leftMetrics.forEach((m, i) => {
    const y = 1.45 + i * 0.76;
    s.addShape(prs.ShapeType.roundRect, { x: 0.5, y, w: 4.2, h: 0.6, fill: { color: C.card }, line: { color: C.border, width: 0.4 }, rectRadius: 0.07 });
    s.addText(m.val, { x: 0.6, y: y + 0.04, w: 4.0, h: 0.3, color: C.bright, fontSize: 16, bold: true, fontFace: 'Courier New' });
    s.addText(m.lbl, { x: 0.6, y: y + 0.34, w: 4.0, h: 0.22, color: C.lgray, fontSize: 8.5 });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.05, w: 4.55, h: 3.55, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1 });
  s.addText('REVENUE IMPACT', { x: 5.25, y: 1.15, w: 4.2, h: 0.22, color: C.orange, fontSize: 8, bold: true, charSpacing: 1.5 });
  const rightMetrics = [
    { val: '+30 countries', lbl: 'Equity expansion addressable per year (vs. 6 manual)' },
    { val: '$250K+', lbl: 'Average Deel ACV per enterprise equity customer' },
    { val: '$7.5M+', lbl: 'Incremental ARR potential (30 new countries × ACV)' },
    { val: '0 blocked', lbl: 'Expansion requests due to compliance complexity' },
  ];
  rightMetrics.forEach((m, i) => {
    const y = 1.45 + i * 0.76;
    s.addShape(prs.ShapeType.roundRect, { x: 5.25, y, w: 4.2, h: 0.6, fill: { color: C.card }, line: { color: C.border, width: 0.4 }, rectRadius: 0.07 });
    s.addText(m.val, { x: 5.35, y: y + 0.04, w: 4.0, h: 0.3, color: C.orange, fontSize: 16, bold: true, fontFace: 'Courier New' });
    s.addText(m.lbl, { x: 5.35, y: y + 0.34, w: 4.0, h: 0.22, color: C.lgray, fontSize: 8.5 });
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · EquityOnboard', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
}

// ── SLIDE 7 — PROOF OF WORK ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: 'FFFFFF' };

  s.addText('PROOF OF WORK', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: '777777', fontSize: 8, bold: true, charSpacing: 2 });
  s.addText("I've Launched Regulated Financial Products from Zero — Twice.", { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: '111111', fontSize: 22, bold: true });

  s.addShape(prs.ShapeType.roundRect, { x: 0.35, y: 1.05, w: 4.55, h: 3.5, fill: { color: C.dark }, line: { color: C.border, width: 0 }, rectRadius: 0.1 });
  s.addText('What I Built at PhonePe & Kotak', { x: 0.5, y: 1.15, w: 4.2, h: 0.28, color: C.bright, fontSize: 10, bold: true });
  const leftPoints = [
    'Kotak Cherry WealthTech (Founding PM, 0→1): scoped MVP, led UX research, co-owned GTM with Compliance + Legal, shipped regulated investment product → 100K+ downloads in launch window',
    'Self-serve merchant onboarding at PhonePe: built PG integration platform with T&C, consent, and regulatory flows — regulated financial product at scale (5,000+ B2B merchants)',
    'Built zero-to-one platforms twice in complex, compliance-heavy domains (fintech, payments)',
    'Cross-functional delivery: Compliance, Legal, Finance, Ops as core stakeholders on every project',
    'KPI instrumentation and success metrics defined before build began — not after',
  ];
  leftPoints.forEach((pt, i) => {
    s.addText(`${i + 1}. ${pt}`, { x: 0.5, y: 1.5 + i * 0.58, w: 4.2, h: 0.52, color: C.white, fontSize: 8.5, valign: 'top' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.05, w: 4.55, h: 3.5, fill: { color: 'F0FDF9' }, line: { color: '6EE7B7', width: 0.8 }, rectRadius: 0.1 });
  s.addText('How It Maps to This Role', { x: 5.25, y: 1.15, w: 4.2, h: 0.28, color: '059669', fontSize: 10, bold: true });
  const rightPoints = [
    'Kotak Cherry 0→1 launch → EquityOnboard 0→1 build: same problem — launch a regulated financial product from scratch with Compliance and Legal as co-owners',
    'Self-serve merchant onboarding → EquityOnboard wizard: same UX pattern — guided setup with compliance guardrails built in',
    'Regulated domain experience → Equity product requires deep Compliance + Legal integration',
    'Compliance + Legal as core stakeholders → EquityOnboard routes directly to Deel compliance team',
    'KPI-first product → Expansion velocity, legal review TAT, countries live per quarter',
  ];
  rightPoints.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 5.25, y: 1.5 + i * 0.58, w: 4.2, h: 0.52, color: '333333', fontSize: 8.5, valign: 'top' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.35, y: 4.65, w: 9.3, h: 0.5, fill: { color: C.dark }, line: { color: C.dark, width: 0 }, rectRadius: 0.07 });
  s.addText('"I\'ve launched two regulated financial products from zero. EquityOnboard is the third — and I know exactly how to get it through Compliance."', { x: 0.5, y: 4.67, w: 9.0, h: 0.46, color: C.bright, fontSize: 9.5, italic: true, valign: 'middle', align: 'center' });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: 'F3F4F6' } });
  s.addText('Deel · EquityOnboard', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: '777777', fontSize: 8, valign: 'middle' });
}

// ── SLIDE 8 — ROLLOUT PLAN ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('ROLLOUT PLAN', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('Start with Top 10 Demand Countries, Scale to 150+', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: C.white, fontSize: 22, bold: true });

  const phases = [
    { ph: 'Phase 1', period: 'Months 1–2', title: 'Top 10 Countries + Core Engine', body: 'Codify compliance ruleset for top 10 demand countries (India, Brazil, UAE, SG, UK, DE, AU, CA, FR, NL). Build gap analysis engine and sub-plan scaffold generator. Instrument expansion velocity KPI.' },
    { ph: 'Phase 2', period: 'Months 3–4', title: 'Self-Serve Wizard + Legal Routing', body: 'Launch country picker UX, AI gap analysis wizard, and sub-plan generation. Build legal review queue with flagged action items. A/B test 4-hour TAT vs. baseline.' },
    { ph: 'Phase 3', period: 'Months 5–6', title: 'Scale to 50+ Countries + Knowledge Loop', body: 'Expand ruleset to 50+ countries. Build knowledge compounding: each approved sub-plan enriches ruleset for future. Integrate with EquityPulse for end-to-end country launch.' },
  ];

  phases.forEach((phase, i) => {
    const x = 0.35 + i * 3.2;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.05, w: 3.0, h: 2.85, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1, shadow: makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y: 1.05, w: 3.0, h: 0.05, fill: { color: C.bright }, line: { color: C.bright, width: 0 } });
    s.addText(phase.ph, { x: x + 0.12, y: 1.12, w: 2.76, h: 0.26, color: C.bright, fontSize: 9, bold: true });
    s.addText(phase.period, { x: x + 0.12, y: 1.37, w: 2.76, h: 0.22, color: C.lgray, fontSize: 8 });
    s.addText(phase.title, { x: x + 0.12, y: 1.6, w: 2.76, h: 0.32, color: C.white, fontSize: 11, bold: true });
    s.addText(phase.body, { x: x + 0.12, y: 1.96, w: 2.76, h: 1.6, color: C.lgray, fontSize: 8.5, valign: 'top' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.35, y: 4.05, w: 9.3, h: 0.72, fill: { color: C.hero }, line: { color: C.bright, width: 0.8 }, rectRadius: 0.08 });
  s.addText('What I need to build this:', { x: 0.55, y: 4.1, w: 2.5, h: 0.24, color: C.bright, fontSize: 9, bold: true });
  s.addText('Access to Deel\'s existing country compliance knowledge base  ·  Compliance team input on gap classification framework  ·  Legal team alignment on sub-plan review workflow  ·  1 backend engineer (ruleset engine + generation API)', {
    x: 0.55, y: 4.34, w: 9.0, h: 0.4, color: C.white, fontSize: 8.5
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439', { x: 0.4, y: 5.38, w: 9.2, h: 0.245, color: C.bright, fontSize: 8.5, align: 'center', valign: 'middle' });
}

prs.writeFile({ fileName: 'assets/deel-equityonboard-deck.pptx' })
  .then(() => console.log('✅ deel-equityonboard-deck.pptx saved'))
  .catch(e => console.error(e));
