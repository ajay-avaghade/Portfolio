const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0D1117',
  hero:   '161B22',
  card:   '1C2333',
  bright: 'FBBF24',
  accent: 'F59E0B',
  green:  '34D399',
  teal:   '00C9A7',
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

  s.addText('TotalRewards', { x: 0.5, y: 0.8, w: 6, h: 0.9, color: C.white, fontSize: 44, bold: true, fontFace: 'Calibri' });
  s.addText('Equity + Compensation Visibility Layer', { x: 0.5, y: 1.65, w: 6.5, h: 0.42, color: C.bright, fontSize: 18, bold: true });
  s.addText('Surface total comp to employees at key moments · HR risk scoring prevents cliff-date attrition', { x: 0.5, y: 2.1, w: 6.5, h: 0.35, color: C.lgray, fontSize: 12 });
  s.addText('Presented by Ajay Avaghade · Senior PM, Equity', { x: 0.5, y: 2.6, w: 5, h: 0.28, color: C.lgray, fontSize: 10 });

  s.addShape(prs.ShapeType.roundRect, { x: 7.2, y: 1.6, w: 2.4, h: 1.6, fill: { color: C.hero }, line: { color: C.bright, width: 1 }, rectRadius: 0.1, shadow: makeShadow() });
  s.addText('−42%', { x: 7.2, y: 1.75, w: 2.4, h: 0.55, color: C.bright, fontSize: 28, bold: true, align: 'center', fontFace: 'Courier New' });
  s.addText('cliff-date\ndepartures', { x: 7.2, y: 2.3, w: 2.4, h: 0.55, color: C.lgray, fontSize: 9, align: 'center', valign: 'middle' });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · TotalRewards · Confidential', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
  s.addText('ajay-avaghade.github.io/Portfolio', { x: 5.5, y: 5.38, w: 4.1, h: 0.245, color: C.bright, fontSize: 8, align: 'right', valign: 'middle' });
}

// ── SLIDE 2 — THE PROBLEM ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE PROBLEM', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText("Employees Leave Without Knowing What They're Forfeiting", { x: 0.4, y: 0.44, w: 9.2, h: 0.65, color: C.white, fontSize: 22, bold: true });

  const cols = [
    { icon: '🤷', stat: '74%', label: 'Employees underestimate their total comp', sub: 'Equity value + benefits are invisible. Most employees only know their base salary. Unvested equity is the biggest blind spot.' },
    { icon: '📅', stat: '1 in 3', label: 'Attrition happens within 90 days of vesting', sub: 'Employees vest, then leave — because they don\'t know the next cliff is coming. No moment to surface the future value.' },
    { icon: '💸', stat: '$280K', label: 'Average replacement cost for senior employee', sub: 'Recruiting fees, time-to-productivity, institutional knowledge loss. Preventable with a single well-timed conversation about unvested equity.' },
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
  s.addText('Root cause: Companies spend millions on equity grants but employees never see the full picture — salary, unvested equity, benefits, and market benchmarks in one place. The result: employees leave money on the table, and companies lose employees unnecessarily.', {
    x: 0.6, y: 3.78, w: 8.8, h: 0.6, color: C.white, fontSize: 9, valign: 'middle'
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · TotalRewards', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
}

// ── SLIDE 3 — THE INSIGHT ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: 'FFFFFF' };

  s.addText('THE INSIGHT', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: '777777', fontSize: 8, bold: true, charSpacing: 2 });
  s.addText("Equity Is the Stickiest Retention Lever Deel Has. It's Just Invisible.", { x: 0.4, y: 0.44, w: 9.2, h: 0.55, color: '111111', fontSize: 21, bold: true });

  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 1.1, w: 4.1, h: 3.4, fill: { color: 'FFF5F5' }, line: { color: 'FECACA', width: 0.8 }, rectRadius: 0.1 });
  s.addText('❌  Today — Equity Is Hidden', { x: 0.6, y: 1.22, w: 3.7, h: 0.3, color: 'DC2626', fontSize: 11, bold: true });
  const leftPts = [
    'Employees log into Carta or Shareworks — a separate tool they rarely visit',
    'Unvested equity value is not shown at all — only grant count and schedule',
    'No market benchmark — employees don\'t know if they\'re paid fairly',
    'No cliff-date countdown — vesting urgency is invisible',
    'HR has no signal until the employee hands in their resignation',
  ];
  leftPts.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 0.6, y: 1.6 + i * 0.44, w: 3.7, h: 0.38, color: '333333', fontSize: 9, valign: 'top' });
  });

  s.addShape(prs.ShapeType.ellipse, { x: 4.55, y: 2.4, w: 0.6, h: 0.45, fill: { color: C.dark }, line: { color: C.dark, width: 0 } });
  s.addText('VS', { x: 4.55, y: 2.4, w: 0.6, h: 0.45, color: C.white, fontSize: 9, bold: true, align: 'center', valign: 'middle' });

  s.addShape(prs.ShapeType.roundRect, { x: 5.2, y: 1.1, w: 4.4, h: 3.4, fill: { color: 'FFFBEB' }, line: { color: 'FCD34D', width: 0.8 }, rectRadius: 0.1 });
  s.addText('✅  TotalRewards (Proposed)', { x: 5.4, y: 1.22, w: 4.0, h: 0.3, color: 'B45309', fontSize: 11, bold: true });
  const rightPts = [
    'Employee sees full total comp in one view: base + equity value + benefits = $284,000/yr',
    'Unvested equity shown prominently: "$186,000 forfeited if you leave today" — in plain language',
    'Next cliff countdown: "87 days until your next $46,500 vest" — creates natural stickiness',
    'Market benchmark: "You\'re at P74 for Sr. PM in EMEA" — fairness visible',
    'HR risk score: ML-powered attrition signal 90 days before cliff — intervene before the conversation',
  ];
  rightPts.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 5.4, y: 1.6 + i * 0.44, w: 4.0, h: 0.38, color: '333333', fontSize: 9, valign: 'top' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 4.62, w: 9.2, h: 0.65, fill: { color: C.dark }, line: { color: C.dark, width: 0 }, rectRadius: 0.08 });
  s.addText("Key insight: Employees who can see their full equity picture stay longer. The nudge costs Deel nothing — the equity was already granted. TotalRewards just makes the hook visible at the right moment.", {
    x: 0.6, y: 4.65, w: 8.8, h: 0.6, color: C.white, fontSize: 9, valign: 'middle'
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: 'F3F4F6' } });
  s.addText('Deel · TotalRewards', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: '777777', fontSize: 8, valign: 'middle' });
}

// ── SLIDE 4 — THE MECHANIC ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE MECHANIC', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('Surface the Right Number at the Right Moment', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: C.white, fontSize: 24, bold: true });

  const steps = [
    { n: '1', title: 'Data Aggregation', body: 'Pull from EquityConnect (equity data), Deel payroll (salary), Deel benefits (insurance, PTO), market salary data API. Unified total comp model per employee.' },
    { n: '2', title: 'Employee View', body: 'Surface total comp = salary + equity value + benefits in one dashboard. Show unvested equity prominently with countdown to next cliff. Market benchmark percentile shown.' },
    { n: '3', title: 'ML Risk Scoring', body: 'Daily ML scoring: engagement signals + cliff proximity + market rate gap + tenure = attrition risk score (0–100). Same architecture as PhonePe Propensity-to-Transact model.' },
    { n: '4', title: 'HR Nudge Engine', body: 'When risk score crosses threshold (e.g. >70), HR manager gets an alert: "Sofia M. — 92 risk score, $320K unvested, cliff in 12 days." Suggested action: retention check-in.' },
    { n: '5', title: 'Retention Loop', body: 'Post-intervention: risk score tracked. If employee stays through cliff, score resets. Outcome data feeds ML model for continuous improvement of risk signal accuracy.' },
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
  s.addText('PhonePe proof: ML risk scoring step (Step 3) uses the same architecture as the Propensity-to-Transact model I shipped — real-time user-level scoring from multi-dimensional signals, replacing manual cohort rules. And the nudge engine is the same context-aware incentivisation architecture (35% AOV uplift, 60% abandonment reduction).', {
    x: 0.56, y: 4.58, w: 9.0, h: 0.6, color: C.white, fontSize: 8.5, valign: 'middle'
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · TotalRewards', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
}

// ── SLIDE 5 — THE PRODUCT ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: 'FFFFFF' };

  s.addText('THE PRODUCT', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: '777777', fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('4 Key Screens — See, Understand, Intervene, Retain', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: '111111', fontSize: 22, bold: true });

  const cards = [
    { n: '01', title: 'Employee Total Comp', body: 'Total comp: $284,000. Breakdown: Salary 49% / Equity 38% / Benefits 13%. Unvested equity shown prominently: $186,000 forfeited if you leave. Market benchmark: P74 · Competitive.' },
    { n: '02', title: 'Equity Vesting Timeline', body: 'Bar chart of vesting amounts over next 12 months. Next cliff highlighted: $46,500 in 87 days. All active grants listed with vest progress. After-tax estimate shown.' },
    { n: '03', title: 'HR Risk Dashboard', body: 'High-risk employees with cliff in 90 days: 12 flagged. Sofia M. — risk score 92, $320K unvested, cliff in 12 days. One-tap "Schedule Retention Check-in" per employee.' },
    { n: '04', title: 'Comp Benchmark (HR)', body: 'Team comp vs. market by level. L3 Mid Engineers below P40 — attrition risk. Equity as % of total comp vs. market benchmark. Q2: 2 at-risk employees retained, $560K replacement cost avoided.' },
  ];

  cards.forEach((card, i) => {
    const x = 0.38 + i * 2.38;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.08, w: 2.2, h: 3.55, fill: { color: 'F9FAFB' }, line: { color: 'E5E7EB', width: 0.6 }, rectRadius: 0.1, shadow: makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y: 1.08, w: 2.2, h: 0.06, fill: { color: C.bright }, line: { color: C.bright, width: 0 } });
    s.addText(card.n, { x, y: 1.16, w: 2.2, h: 0.35, color: C.bright, fontSize: 20, bold: true, align: 'center', fontFace: 'Courier New' });
    s.addText(card.title, { x: x + 0.1, y: 1.5, w: 2.0, h: 0.34, color: '111111', fontSize: 11, bold: true, align: 'center' });
    s.addText(card.body, { x: x + 0.1, y: 1.88, w: 2.0, h: 2.4, color: '555555', fontSize: 8.5, align: 'left', valign: 'top' });
  });

  s.addText('Interactive prototype: deel-totalrewards-prototype.html', { x: 0.4, y: 4.74, w: 9.2, h: 0.3, color: '777777', fontSize: 9, align: 'center' });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: 'F3F4F6' } });
  s.addText('Deel · TotalRewards', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: '777777', fontSize: 8, valign: 'middle' });
}

// ── SLIDE 6 — IMPACT & ROI ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('IMPACT & ROI', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('Make Equity Work as a Retention Lever — Not Just a Compensation Line Item', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: C.white, fontSize: 20, bold: true });

  s.addShape(prs.ShapeType.roundRect, { x: 0.35, y: 1.05, w: 4.55, h: 3.55, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1 });
  s.addText('EMPLOYEE IMPACT', { x: 0.5, y: 1.15, w: 4.2, h: 0.22, color: C.bright, fontSize: 8, bold: true, charSpacing: 1.5 });
  const leftMetrics = [
    { val: '−42%', lbl: 'Cliff-date departures (post-TotalRewards launch)' },
    { val: '94%', lbl: 'Retention after receiving full total comp briefing' },
    { val: '$186K', lbl: 'Average unvested equity surfaced per employee' },
    { val: 'P74', lbl: 'Market benchmark visible — fairness confirmed' },
  ];
  leftMetrics.forEach((m, i) => {
    const y = 1.45 + i * 0.76;
    s.addShape(prs.ShapeType.roundRect, { x: 0.5, y, w: 4.2, h: 0.6, fill: { color: C.card }, line: { color: C.border, width: 0.4 }, rectRadius: 0.07 });
    s.addText(m.val, { x: 0.6, y: y + 0.04, w: 4.0, h: 0.3, color: C.bright, fontSize: 16, bold: true, fontFace: 'Courier New' });
    s.addText(m.lbl, { x: 0.6, y: y + 0.34, w: 4.0, h: 0.22, color: C.lgray, fontSize: 8.5 });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.05, w: 4.55, h: 3.55, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1 });
  s.addText('COMPANY / DEEL ROI', { x: 5.25, y: 1.15, w: 4.2, h: 0.22, color: C.green, fontSize: 8, bold: true, charSpacing: 1.5 });
  const rightMetrics = [
    { val: '$560K', lbl: 'Replacement cost avoided per 2 retained employees' },
    { val: '+NPS', lbl: 'Employee satisfaction with comp transparency' },
    { val: '$0', lbl: 'Marginal cost of retention nudge (equity already granted)' },
    { val: 'Stickiness', lbl: 'Deel HRIS becomes essential to HR retention strategy' },
  ];
  rightMetrics.forEach((m, i) => {
    const y = 1.45 + i * 0.76;
    s.addShape(prs.ShapeType.roundRect, { x: 5.25, y, w: 4.2, h: 0.6, fill: { color: C.card }, line: { color: C.border, width: 0.4 }, rectRadius: 0.07 });
    s.addText(m.val, { x: 5.35, y: y + 0.04, w: 4.0, h: 0.3, color: C.green, fontSize: 16, bold: true, fontFace: 'Courier New' });
    s.addText(m.lbl, { x: 5.35, y: y + 0.34, w: 4.0, h: 0.22, color: C.lgray, fontSize: 8.5 });
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · TotalRewards', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
}

// ── SLIDE 7 — PROOF OF WORK ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: 'FFFFFF' };

  s.addText('PROOF OF WORK', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: '777777', fontSize: 8, bold: true, charSpacing: 2 });
  s.addText("The ML Scoring + Nudge Architecture Is Already Proven.", { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: '111111', fontSize: 22, bold: true });

  s.addShape(prs.ShapeType.roundRect, { x: 0.35, y: 1.05, w: 4.55, h: 3.5, fill: { color: C.dark }, line: { color: C.border, width: 0 }, rectRadius: 0.1 });
  s.addText('What I Built at PhonePe', { x: 0.5, y: 1.15, w: 4.2, h: 0.28, color: C.bright, fontSize: 10, bold: true });
  const leftPoints = [
    'Propensity-to-Transact ML model: real-time user-level scoring from multi-dimensional behavioural signals, replacing static cohort rules — same architecture as TotalRewards risk scoring engine',
    'Context-aware incentivisation engine: surfaced the right offer at the right moment (cart value × user intent × time-of-day) → 35% AOV uplift, 60% abandonment reduction',
    '₹100 Cr/yr rewards rebuild: ML audience segmentation + dynamic offer calibration + behaviour-linked incentives across 350M+ MAU',
    'Milestone incentive system: goal-linked mechanics (not just discounts) → 60% improvement in 90-day device activation vs. flat discount',
    'All systems KPI-instrumented before launch — retention, engagement, TAT tracked from day 0',
  ];
  leftPoints.forEach((pt, i) => {
    s.addText(`${i + 1}. ${pt}`, { x: 0.5, y: 1.5 + i * 0.58, w: 4.2, h: 0.52, color: C.white, fontSize: 8.5, valign: 'top' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.05, w: 4.55, h: 3.5, fill: { color: 'FFFBEB' }, line: { color: 'FCD34D', width: 0.8 }, rectRadius: 0.1 });
  s.addText('How It Maps to This Role', { x: 5.25, y: 1.15, w: 4.2, h: 0.28, color: 'B45309', fontSize: 10, bold: true });
  const rightPoints = [
    'Propensity-to-Transact → TotalRewards risk scoring: same multi-signal ML pipeline, feature set changes (equity + tenure + engagement + cliff proximity)',
    'Dynamic incentivisation engine → Nudge engine: same architecture — right intervention, right employee, right moment (cliff countdown)',
    'Rewards rebuild: ML segmentation + behaviour-linked incentives → TotalRewards: equity-linked retention mechanics',
    'Milestone mechanics → Cliff-date stickiness: goal-linked incentive design proven to change behaviour',
    'KPI instrumentation → Retention rate, risk score accuracy, replacement cost avoided pre-defined',
  ];
  rightPoints.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 5.25, y: 1.5 + i * 0.58, w: 4.2, h: 0.52, color: '333333', fontSize: 8.5, valign: 'top' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.35, y: 4.65, w: 9.3, h: 0.5, fill: { color: C.dark }, line: { color: C.dark, width: 0 }, rectRadius: 0.07 });
  s.addText('"I\'ve built the ML scoring + nudge architecture twice at 350M+ MAU scale. TotalRewards is the same system with equity data as the signal."', { x: 0.5, y: 4.67, w: 9.0, h: 0.46, color: C.bright, fontSize: 9.5, italic: true, valign: 'middle', align: 'center' });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: 'F3F4F6' } });
  s.addText('Deel · TotalRewards', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: '777777', fontSize: 8, valign: 'middle' });
}

// ── SLIDE 8 — ROLLOUT PLAN ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('ROLLOUT PLAN', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('Launch Employee View First, Add HR Intelligence Layer', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: C.white, fontSize: 22, bold: true });

  const phases = [
    { ph: 'Phase 1', period: 'Months 1–2', title: 'Employee Total Comp Dashboard', body: 'Aggregate salary + equity (via EquityConnect) + benefits into unified view. Launch employee-facing dashboard. Instrument: session depth, unvested equity viewed, cliff date engagement rate.' },
    { ph: 'Phase 2', period: 'Months 3–4', title: 'ML Risk Scoring + HR Alerts', body: 'Build attrition risk model: cliff proximity + engagement + market rate gap + tenure signals. Launch HR risk dashboard with flagged employees. Measure: intervention rate, retention outcome.' },
    { ph: 'Phase 3', period: 'Months 5–6', title: 'Comp Intelligence + Benchmark Layer', body: 'Add market salary benchmarking (P25/P50/P75 by role, level, geo). HR comp analysis by team. Integrate with performance review and offer letter workflows for full Total Rewards positioning.' },
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
  s.addText('EquityConnect data pipeline (equity value feed)  ·  Payroll API access (salary data)  ·  Market salary data partnership (Levels.fyi / Radford / Mercer)  ·  1 DS engineer (ML risk model)  ·  HR team design partner for UX research', {
    x: 0.55, y: 4.34, w: 9.0, h: 0.4, color: C.white, fontSize: 8.5
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439', { x: 0.4, y: 5.38, w: 9.2, h: 0.245, color: C.bright, fontSize: 8.5, align: 'center', valign: 'middle' });
}

prs.writeFile({ fileName: 'assets/deel-totalrewards-deck.pptx' })
  .then(() => console.log('✅ deel-totalrewards-deck.pptx saved'))
  .catch(e => console.error(e));
