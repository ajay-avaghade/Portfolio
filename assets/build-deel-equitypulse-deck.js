const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0D1117',
  hero:   '161B22',
  card:   '1C2333',
  bright: '00C9A7',
  accent: '00A888',
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

  // Diagonal stripes
  for (let i = 0; i < 5; i++) {
    const lx = -0.5 + i * 2.2;
    const lw = Math.min(3.5, W - lx - 0.15);
    if (lx > W || lw <= 0) continue;
    s.addShape(prs.ShapeType.line, { x: lx, y: 0, w: lw, h: H, line: { color: C.bright, width: 0.4 }, rotate: 30, transparency: 88 });
  }

  // Left accent bar
  s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 0.06, h: H, fill: { color: C.bright } });

  // Company tag
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 0.35, w: 1.8, h: 0.28, fill: { color: C.hero }, line: { color: C.bright, width: 0.8 }, rectRadius: 0.06 });
  s.addText('DEEL · EQUITY PRODUCT', { x: 0.5, y: 0.35, w: 1.8, h: 0.28, color: C.bright, fontSize: 7, bold: true, align: 'center', valign: 'middle', charSpacing: 1.5 });

  // Title
  s.addText('EquityPulse', { x: 0.5, y: 0.8, w: 6, h: 0.9, color: C.white, fontSize: 44, bold: true, fontFace: 'Calibri' });
  s.addText('AI-Powered Equity Event Orchestration', { x: 0.5, y: 1.65, w: 6.5, h: 0.42, color: C.bright, fontSize: 18, bold: true });
  s.addText('Zero-touch vesting → compliance → payroll automation across 150+ countries', { x: 0.5, y: 2.1, w: 6.5, h: 0.35, color: C.lgray, fontSize: 12 });

  // Presenter
  s.addText('Presented by Ajay Avaghade · Senior PM, Equity', { x: 0.5, y: 2.6, w: 5, h: 0.28, color: C.lgray, fontSize: 10 });

  // Big stat box
  s.addShape(prs.ShapeType.roundRect, { x: 7.2, y: 1.6, w: 2.4, h: 1.6, fill: { color: C.hero }, line: { color: C.bright, width: 1 }, rectRadius: 0.1, shadow: makeShadow() });
  s.addText('$2,400', { x: 7.2, y: 1.75, w: 2.4, h: 0.55, color: C.bright, fontSize: 28, bold: true, align: 'center', fontFace: 'Courier New' });
  s.addText('ops cost saved\nper vesting event', { x: 7.2, y: 2.3, w: 2.4, h: 0.55, color: C.lgray, fontSize: 9, align: 'center', valign: 'middle' });

  // Bottom
  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · EquityPulse · Confidential', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
  s.addText('ajay-avaghade.github.io/Portfolio', { x: 5.5, y: 5.38, w: 4.1, h: 0.245, color: C.bright, fontSize: 8, align: 'right', valign: 'middle' });
}

// ── SLIDE 2 — THE PROBLEM ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE PROBLEM', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('Global Equity Is 80% Manual Ops — And It Breaks At Scale', { x: 0.4, y: 0.44, w: 9, h: 0.6, color: C.white, fontSize: 24, bold: true });

  // Three stat columns
  const cols = [
    { icon: '⏱', stat: '3.5 days', label: 'Manual TAT per vesting event', sub: 'Research + withholding calc + payroll entry + filing + notifications' },
    { icon: '📊', stat: '150+', label: 'Country rule sets to track', sub: 'Each with unique withholding rates, filing deadlines, and language requirements' },
    { icon: '💸', stat: '$2,400', label: 'Ops cost per event (manual)', sub: 'Compliance analyst time + error correction + delayed payroll processing' },
  ];
  cols.forEach((col, i) => {
    const x = 0.4 + i * 3.1;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.2, w: 2.9, h: 2.4, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1, shadow: makeShadow() });
    s.addText(col.icon, { x, y: 1.32, w: 2.9, h: 0.4, fontSize: 22, align: 'center' });
    s.addText(col.stat, { x, y: 1.7, w: 2.9, h: 0.55, color: C.bright, fontSize: 28, bold: true, align: 'center', fontFace: 'Courier New' });
    s.addText(col.label, { x: x + 0.1, y: 2.25, w: 2.7, h: 0.3, color: C.white, fontSize: 10, bold: true, align: 'center' });
    s.addText(col.sub, { x: x + 0.1, y: 2.55, w: 2.7, h: 0.7, color: C.lgray, fontSize: 8, align: 'center', valign: 'top' });
  });

  // Insight box
  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 3.75, w: 9.2, h: 0.65, fill: { color: C.hero }, line: { color: C.bright, width: 0.8 }, rectRadius: 0.08 });
  s.addText('Root cause: Every vesting event in a global company triggers a cascade of country-specific actions. Today, each one is touched by a human. One company with 500 employees vesting across 20 countries generates 500+ manual compliance touches per quarter — with a 4.2% error rate that creates legal exposure.', {
    x: 0.6, y: 3.78, w: 8.8, h: 0.6, color: C.white, fontSize: 9, valign: 'middle'
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · EquityPulse', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
}

// ── SLIDE 3 — THE INSIGHT ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: 'FFFFFF' };

  s.addText('THE INSIGHT', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: '777777', fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('The Compliance Ruleset Is Deterministic. Every Step Can Be Automated.', { x: 0.4, y: 0.44, w: 9.2, h: 0.55, color: '111111', fontSize: 22, bold: true });

  // Left: current
  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 1.1, w: 4.1, h: 3.4, fill: { color: 'FFF5F5' }, line: { color: 'FECACA', width: 0.8 }, rectRadius: 0.1 });
  s.addText('❌  Manual Process (Today)', { x: 0.6, y: 1.22, w: 3.7, h: 0.3, color: 'DC2626', fontSize: 11, bold: true });
  const leftPts = [
    'Compliance analyst manually researches rules for each country on each event',
    'Withholding rates looked up in spreadsheets — prone to version drift',
    'Payroll items keyed in manually — 3.5 day turnaround, 4.2% error rate',
    'Employee notifications drafted per event — inconsistent language',
    'No audit trail — compliance risk at every step',
  ];
  leftPts.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 0.6, y: 1.6 + i * 0.44, w: 3.7, h: 0.38, color: '333333', fontSize: 9, valign: 'top' });
  });

  // VS oval
  s.addShape(prs.ShapeType.ellipse, { x: 4.55, y: 2.4, w: 0.6, h: 0.45, fill: { color: C.dark }, line: { color: C.dark, width: 0 } });
  s.addText('VS', { x: 4.55, y: 2.4, w: 0.6, h: 0.45, color: C.white, fontSize: 9, bold: true, align: 'center', valign: 'middle' });

  // Right: proposed
  s.addShape(prs.ShapeType.roundRect, { x: 5.2, y: 1.1, w: 4.4, h: 3.4, fill: { color: 'F0FDF9' }, line: { color: '6EE7B7', width: 0.8 }, rectRadius: 0.1 });
  s.addText('✅  EquityPulse (Proposed)', { x: 5.4, y: 1.22, w: 4.0, h: 0.3, color: '059669', fontSize: 11, bold: true });
  const rightPts = [
    'AI engine ingests vesting events via real-time webhook from Carta / Shareworks / Fidelity',
    'Country ruleset database auto-applies correct withholding rate, filing type, and currency',
    'Payroll items generated and queued automatically — 4-minute TAT, 0 errors',
    'Employee notifications dispatched in local language with personalized tax details',
    'Full audit trail logged automatically — compliance-ready at every step',
  ];
  rightPts.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 5.4, y: 1.6 + i * 0.44, w: 4.0, h: 0.38, color: '333333', fontSize: 9, valign: 'top' });
  });

  // Callout
  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 4.62, w: 9.2, h: 0.65, fill: { color: C.dark }, line: { color: C.dark, width: 0 }, rectRadius: 0.08 });
  s.addText('Key insight: Tax withholding rules, filing requirements, and notification obligations are fully deterministic by country × equity type × event type. There is no reason for a human to make this decision every time.', {
    x: 0.6, y: 4.65, w: 8.8, h: 0.6, color: C.white, fontSize: 9, valign: 'middle'
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: 'F3F4F6' } });
  s.addText('Deel · EquityPulse', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: '777777', fontSize: 8, valign: 'middle' });
}

// ── SLIDE 4 — THE MECHANIC ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE MECHANIC', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('5-Step Zero-Touch Orchestration Engine', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: C.white, fontSize: 24, bold: true });

  const steps = [
    { n: '1', title: 'Event Ingestion', body: 'Real-time webhook fires from Carta / Shareworks / Fidelity when a vesting event triggers. Event data normalised to Deel canonical equity model.' },
    { n: '2', title: 'AI Compliance Check', body: 'Engine queries country ruleset DB (150+ countries): withholding rate, filing type, currency, local securities obligations. 0 manual lookups.' },
    { n: '3', title: 'Payroll Item Generation', body: 'Payroll items auto-created per employee per country: gross-up calculation, FX conversion, withholding deduction. Queued to next payroll run.' },
    { n: '4', title: 'Compliance Filing Queue', body: 'Country-specific filing obligations (W-2 supplement, PAYE, Lohnsteuer, Form 3CEB) auto-queued with pre-populated data. Legal review where required.' },
    { n: '5', title: 'Employee Notification', body: 'Personalised notification dispatched in local language with exact vest value, net after tax, and next steps. Full audit trail logged.' },
  ];

  steps.forEach((step, i) => {
    const x = 0.38 + i * 1.88;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.05, w: 1.72, h: 3.4, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1, shadow: makeShadow() });
    // Number badge
    s.addShape(prs.ShapeType.ellipse, { x: x + 0.62, y: 1.12, w: 0.46, h: 0.38, fill: { color: C.bright }, line: { color: C.bright, width: 0 } });
    s.addText(step.n, { x: x + 0.62, y: 1.12, w: 0.46, h: 0.38, color: C.dark, fontSize: 12, bold: true, align: 'center', valign: 'middle' });
    s.addText(step.title, { x: x + 0.06, y: 1.58, w: 1.6, h: 0.38, color: C.white, fontSize: 10, bold: true, align: 'center' });
    s.addText(step.body, { x: x + 0.08, y: 2.0, w: 1.56, h: 1.9, color: C.lgray, fontSize: 8, align: 'center', valign: 'top' });
    // Connector arrow (not on last)
    if (i < 4) {
      s.addShape(prs.ShapeType.line, { x: x + 1.72, y: 2.2, w: 0.16, h: 0, line: { color: C.bright, width: 1.2 } });
    }
  });

  // Proof note
  s.addShape(prs.ShapeType.roundRect, { x: 0.38, y: 4.56, w: 9.24, h: 0.64, fill: { color: C.hero }, line: { color: C.bright, width: 0.8 }, rectRadius: 0.08 });
  s.addText('PhonePe proof: Same deterministic orchestration logic applied to marketing spend targeting — replaced 6 siloed analyst workflows with one ML pipeline, cutting decision TAT from 2 days to 4 hours and reducing errors 68%. EquityPulse applies the same pattern to equity event processing.', {
    x: 0.56, y: 4.58, w: 9.0, h: 0.6, color: C.white, fontSize: 8.5, valign: 'middle'
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · EquityPulse', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
}

// ── SLIDE 5 — THE PRODUCT ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: 'FFFFFF' };

  s.addText('THE PRODUCT', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: '777777', fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('4 Key Screens — Built & Validated at PhonePe Scale', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: '111111', fontSize: 22, bold: true });

  const cards = [
    { n: '01', title: 'Vesting Dashboard', body: 'Event cards with country flags, status badges (Ready / Review / Pending), auto-processing stats. Ops team sees zero-touch confirmation before event fires.' },
    { n: '02', title: 'Compliance Check', body: 'Per-country breakdown: withholding rate, filing type, payroll items count. AI reads deterministic ruleset — no analyst involvement.' },
    { n: '03', title: 'Payroll Queued', body: 'Confirmation screen: items queued, compliance filings scheduled, notifications dispatched. Full audit trail. 4-min TAT displayed.' },
    { n: '04', title: 'Ops Console', body: 'Quarterly view: events processed, error rate, TAT, cost saved vs. manual baseline. Override queue for edge cases. Manual queue < 1% of volume.' },
  ];

  cards.forEach((card, i) => {
    const x = 0.38 + i * 2.38;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.08, w: 2.2, h: 3.55, fill: { color: 'F9FAFB' }, line: { color: 'E5E7EB', width: 0.6 }, rectRadius: 0.1, shadow: makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y: 1.08, w: 2.2, h: 0.06, fill: { color: C.bright }, line: { color: C.bright, width: 0 } });
    s.addText(card.n, { x, y: 1.16, w: 2.2, h: 0.35, color: C.bright, fontSize: 20, bold: true, align: 'center', fontFace: 'Courier New' });
    s.addText(card.title, { x: x + 0.1, y: 1.5, w: 2.0, h: 0.34, color: '111111', fontSize: 11, bold: true, align: 'center' });
    s.addText(card.body, { x: x + 0.1, y: 1.88, w: 2.0, h: 2.4, color: '555555', fontSize: 8.5, align: 'left', valign: 'top' });
  });

  s.addText('Interactive prototype: deel-equitypulse-prototype.html', { x: 0.4, y: 4.74, w: 9.2, h: 0.3, color: '777777', fontSize: 9, align: 'center' });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: 'F3F4F6' } });
  s.addText('Deel · EquityPulse', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: '777777', fontSize: 8, valign: 'middle' });
}

// ── SLIDE 6 — IMPACT & ROI ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('IMPACT & ROI', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('Built on PhonePe Proof — Validated Ops Intelligence Architecture', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: C.white, fontSize: 22, bold: true });

  // Left: Operations Impact
  s.addShape(prs.ShapeType.roundRect, { x: 0.35, y: 1.05, w: 4.55, h: 3.55, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1 });
  s.addText('OPERATIONS IMPACT', { x: 0.5, y: 1.15, w: 4.2, h: 0.22, color: C.bright, fontSize: 8, bold: true, charSpacing: 1.5 });
  const leftMetrics = [
    { val: '3.5 days → 4 min', lbl: 'Vesting event processing TAT' },
    { val: '4.2% → 0%', lbl: 'Compliance error rate' },
    { val: '$2,400', lbl: 'Ops cost saved per event' },
    { val: '100%', lbl: 'Compliance coverage across 150+ countries' },
  ];
  leftMetrics.forEach((m, i) => {
    const y = 1.45 + i * 0.76;
    s.addShape(prs.ShapeType.roundRect, { x: 0.5, y, w: 4.2, h: 0.6, fill: { color: C.card }, line: { color: C.border, width: 0.4 }, rectRadius: 0.07 });
    s.addText(m.val, { x: 0.6, y: y + 0.04, w: 4.0, h: 0.3, color: C.bright, fontSize: 16, bold: true, fontFace: 'Courier New' });
    s.addText(m.lbl, { x: 0.6, y: y + 0.34, w: 4.0, h: 0.22, color: C.lgray, fontSize: 8.5 });
  });

  // Right: Company ROI
  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.05, w: 4.55, h: 3.55, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1 });
  s.addText('COMPANY ROI', { x: 5.25, y: 1.15, w: 4.2, h: 0.22, color: C.orange, fontSize: 8, bold: true, charSpacing: 1.5 });
  const rightMetrics = [
    { val: '$1.2M/yr', lbl: 'Ops cost reduction (500-employee company)' },
    { val: '−68%', lbl: 'Compliance escalation tickets' },
    { val: '5× faster', lbl: 'New country equity expansion' },
    { val: 'Zero', lbl: 'Legal exposure from compliance errors' },
  ];
  rightMetrics.forEach((m, i) => {
    const y = 1.45 + i * 0.76;
    s.addShape(prs.ShapeType.roundRect, { x: 5.25, y, w: 4.2, h: 0.6, fill: { color: C.card }, line: { color: C.border, width: 0.4 }, rectRadius: 0.07 });
    s.addText(m.val, { x: 5.35, y: y + 0.04, w: 4.0, h: 0.3, color: C.orange, fontSize: 16, bold: true, fontFace: 'Courier New' });
    s.addText(m.lbl, { x: 5.35, y: y + 0.34, w: 4.0, h: 0.22, color: C.lgray, fontSize: 8.5 });
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · EquityPulse', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
}

// ── SLIDE 7 — PROOF OF WORK ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: 'FFFFFF' };

  s.addText('PROOF OF WORK', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: '777777', fontSize: 8, bold: true, charSpacing: 2 });
  s.addText("I Built This. The Architecture Already Exists.", { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: '111111', fontSize: 22, bold: true });

  // Left dark column
  s.addShape(prs.ShapeType.roundRect, { x: 0.35, y: 1.05, w: 4.55, h: 3.5, fill: { color: C.dark }, line: { color: C.border, width: 0 }, rectRadius: 0.1 });
  s.addText('What I Built at PhonePe', { x: 0.5, y: 1.15, w: 4.2, h: 0.28, color: C.bright, fontSize: 10, bold: true });
  const leftPoints = [
    'Ops intelligence layer replacing 6 siloed analyst workflows — TAT 2 days → 4 hours, −68% escalation tickets (direct EquityPulse parallel)',
    'Propensity-to-Transact ML model — real-time per-user scoring replacing manual cohort rules across 350M+ MAU',
    'Context-aware incentivisation engine with deterministic rule evaluation: cart value × user intent × time-of-day signals',
    'Pre-defined KPI instrumentation before build began — all metrics baselined and tracked post-launch',
    'Cross-functional delivery: Tech, DS, Compliance, Finance, Legal, Ops',
  ];
  leftPoints.forEach((pt, i) => {
    s.addText(`${i + 1}. ${pt}`, { x: 0.5, y: 1.5 + i * 0.58, w: 4.2, h: 0.52, color: C.white, fontSize: 8.5, valign: 'top' });
  });

  // Right light column
  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.05, w: 4.55, h: 3.5, fill: { color: 'F0FDF9' }, line: { color: '6EE7B7', width: 0.8 }, rectRadius: 0.1 });
  s.addText('How It Maps to This Role', { x: 5.25, y: 1.15, w: 4.2, h: 0.28, color: '059669', fontSize: 10, bold: true });
  const rightPoints = [
    'Ops intelligence layer → EquityPulse orchestration engine: same deterministic workflow automation pattern',
    'ML scoring pipeline → AI compliance check: country ruleset query replaces manual research',
    'Dynamic trigger engine → Vesting event → payroll item generation: same auto-queue architecture',
    'Pre-defined KPIs → Equity product success metrics: TAT, error rate, cost-per-event, country coverage',
    'Cross-functional execution → Equity needs Compliance, Payroll, DS, Legal: identical stakeholder map',
  ];
  rightPoints.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 5.25, y: 1.5 + i * 0.58, w: 4.2, h: 0.52, color: '333333', fontSize: 8.5, valign: 'top' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.35, y: 4.65, w: 9.3, h: 0.5, fill: { color: C.dark }, line: { color: C.dark, width: 0 }, rectRadius: 0.07 });
  s.addText('"This isn\'t a case study — it\'s the same system I shipped, applied to equity events instead of marketing campaigns."', { x: 0.5, y: 4.67, w: 9.0, h: 0.46, color: C.bright, fontSize: 9.5, italic: true, valign: 'middle', align: 'center' });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: 'F3F4F6' } });
  s.addText('Deel · EquityPulse', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: '777777', fontSize: 8, valign: 'middle' });
}

// ── SLIDE 8 — ROLLOUT PLAN ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('ROLLOUT PLAN', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('Phased Build — Starting with Highest-Volume Countries', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: C.white, fontSize: 22, bold: true });

  const phases = [
    { ph: 'Phase 1', period: 'Months 1–2', title: 'Core Engine (US/UK/DE)', body: 'Build event ingestion layer (Carta webhook), compliance ruleset DB for top 3 countries, payroll item generation. Instrument TAT and error-rate KPIs before launch.' },
    { ph: 'Phase 2', period: 'Months 3–4', title: 'Expand + Integrate', body: 'Add 20 more countries. Integrate with EquityConnect data hub. Launch employee notification engine. A/B test manual override rate.' },
    { ph: 'Phase 3', period: 'Months 5–6', title: 'Full Platform Coverage', body: 'Expand to 150+ countries. Self-serve ruleset updates for compliance team. Ops console GA. Connect to TotalRewards for employee visibility layer.' },
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

  // Ask box
  s.addShape(prs.ShapeType.roundRect, { x: 0.35, y: 4.05, w: 9.3, h: 0.72, fill: { color: C.hero }, line: { color: C.bright, width: 0.8 }, rectRadius: 0.08 });
  s.addText('What I need to build this:', { x: 0.55, y: 4.1, w: 2.5, h: 0.24, color: C.bright, fontSize: 9, bold: true });
  s.addText('Access to cap table platform API docs (Carta, Shareworks)  ·  Compliance ruleset data from Deel\'s existing country knowledge  ·  1 backend engineer (event ingestion + payroll item API)  ·  Alignment with Payroll PM on queue handoff interface', {
    x: 0.55, y: 4.34, w: 9.0, h: 0.4, color: C.white, fontSize: 8.5
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439', { x: 0.4, y: 5.38, w: 9.2, h: 0.245, color: C.bright, fontSize: 8.5, align: 'center', valign: 'middle' });
}

prs.writeFile({ fileName: 'assets/deel-equitypulse-deck.pptx' })
  .then(() => console.log('✅ deel-equitypulse-deck.pptx saved'))
  .catch(e => console.error(e));
