// build-eyewa-rxops-deck.js
// eyewa RxOps — AI-Powered Prescription Order Intelligence
// Run: node assets/build-eyewa-rxops-deck.js

const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '071525',
  hero:   '0A1E38',
  card:   '0D1B2A',
  teal:   '00BFA5',
  amber:  'F59E0B',
  red:    'EF4444',
  green:  '22C55E',
  white:  'FFFFFF',
  muted:  '94A3B8',
  lgray:  'E2E8F0',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.18 });

// ─── SLIDE 1: COVER ───────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  // diagonal accent lines
  for (let i = 0; i < 8; i++) {
    s.addShape(prs.ShapeType.line, { x: -0.5 + i * 1.5, y: 0, w: 4, h: 7.5, line: { color: C.teal, width: 0.4, transparency: 88 }, rotate: 30 });
  }

  // teal top bar
  s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal } });

  // eyebrow
  s.addText('EYEWA · INTERNAL OPERATIONS', { x: 0.6, y: 0.55, w: 6, h: 0.25, fontSize: 9, color: C.teal, bold: true, charSpacing: 3 });

  // product name
  s.addText('RxOps', { x: 0.6, y: 0.9, w: 9, h: 1.4, fontSize: 72, color: C.white, bold: true });

  // subtitle
  s.addText('AI-Powered Prescription Order Intelligence', { x: 0.6, y: 2.15, w: 7, h: 0.5, fontSize: 22, color: C.muted });

  // description
  s.addText('Eliminating SLA breaches and manual errors across the Rx order pipeline through real-time AI validation, intelligent lab routing, and predictive delay detection.', {
    x: 0.6, y: 2.8, w: 6.2, h: 0.9, fontSize: 13, color: C.muted, lineSpacingMultiple: 1.4,
  });

  // presenter line
  s.addText('Presented by Ajay Avaghade · Growth & Operations PM', { x: 0.6, y: 3.85, w: 6, h: 0.28, fontSize: 11, color: '475569' });

  // big stat bottom-right
  s.addShape(prs.ShapeType.rect, { x: 7.4, y: 1.8, w: 2.2, h: 2.2, fill: { color: C.teal, transparency: 90 }, line: { color: C.teal, width: 1 }, shadow: makeShadow() });
  s.addText('30%', { x: 7.4, y: 2.0, w: 2.2, h: 0.9, fontSize: 48, color: C.teal, bold: true, align: 'center' });
  s.addText('of Rx orders\nhave manual\nerrors today', { x: 7.4, y: 2.85, w: 2.2, h: 0.8, fontSize: 10, color: C.muted, align: 'center', lineSpacingMultiple: 1.35 });
}

// ─── SLIDE 2: THE PROBLEM ─────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.hero };

  s.addText('THE PROBLEM', { x: 0.55, y: 0.35, w: 9, h: 0.22, fontSize: 9, color: C.teal, bold: true, charSpacing: 3 });
  s.addText('Every Rx Order Touches 7 Manual Steps — Each One a Point of Failure', {
    x: 0.55, y: 0.62, w: 9, h: 0.75, fontSize: 26, color: C.white, bold: true, lineSpacingMultiple: 1.2,
  });

  // 3 stat columns
  const stats = [
    { icon: '📦', num: '30%', lbl: 'Rx orders delayed or remade', sub: 'Due to prescription entry errors, lab overload, or missed SLA handoffs' },
    { icon: '⏱', num: '8.2hrs', lbl: 'Avg manual Rx validation time', sub: 'Per order — even a simple single-vision script goes through 3 manual checks' },
    { icon: '💸', num: 'AED 250', lbl: 'Cost of one lens remake', sub: 'Wrong prescription or incorrect axis cut = full redo at lab expense' },
  ];

  stats.forEach((st, i) => {
    const x = 0.55 + i * 3.15;
    s.addShape(prs.ShapeType.rect, { x, y: 1.65, w: 2.9, h: 2.8, fill: { color: C.card }, line: { color: '1a2e42', width: 0.8 }, shadow: makeShadow() });
    s.addText(st.icon, { x, y: 1.8, w: 2.9, h: 0.45, fontSize: 22, align: 'center' });
    s.addText(st.num, { x, y: 2.25, w: 2.9, h: 0.75, fontSize: 38, color: C.red, bold: true, align: 'center' });
    s.addText(st.lbl, { x, y: 2.95, w: 2.9, h: 0.4, fontSize: 11, color: C.white, bold: true, align: 'center' });
    s.addText(st.sub, { x: x + 0.1, y: 3.38, w: 2.7, h: 0.85, fontSize: 9.5, color: C.muted, align: 'center', lineSpacingMultiple: 1.4 });
  });

  // Insight box
  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 4.62, w: 9.0, h: 0.88, fill: { color: '0f0a00' }, line: { color: C.amber, width: 0.8 } });
  s.addText('Root cause: Rx operations at eyewa run on human judgement and siloed lab portals. No system predicts which order will miss SLA until it already has. Every mistake costs time, money, and a customer relationship.', {
    x: 0.75, y: 4.72, w: 8.6, h: 0.68, fontSize: 11, color: C.amber, lineSpacingMultiple: 1.4,
  });
}

// ─── SLIDE 3: THE INSIGHT ─────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.white };

  s.addText('THE INSIGHT', { x: 0.55, y: 0.35, w: 9, h: 0.22, fontSize: 9, color: C.teal, bold: true, charSpacing: 3 });
  s.addText('Reactive Ops Firefighting vs. Predictive AI Intelligence', {
    x: 0.55, y: 0.62, w: 9, h: 0.6, fontSize: 26, color: '0f172a', bold: true,
  });

  // Left: current state
  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 1.42, w: 4.2, h: 3.4, fill: { color: 'fef2f2' }, line: { color: 'fca5a5', width: 0.8 } });
  s.addText('CURRENT STATE', { x: 0.75, y: 1.58, w: 3.8, h: 0.25, fontSize: 9, color: C.red, bold: true, charSpacing: 2 });
  s.addText('Reactive & Siloed', { x: 0.75, y: 1.88, w: 3.8, h: 0.35, fontSize: 15, color: '1e293b', bold: true });
  const current = [
    'Prescription errors found AFTER lab cuts lens',
    'Lab routing is manual — no capacity awareness',
    'SLA breach discovered when customer complains',
    'Ops team firefights daily instead of managing',
    'Each lab portal is a separate system to check',
  ];
  current.forEach((t, i) => {
    s.addText('✗  ' + t, { x: 0.75, y: 2.32 + i * 0.42, w: 3.8, h: 0.38, fontSize: 10.5, color: '64748b', lineSpacingMultiple: 1.3 });
  });

  // VS divider
  s.addShape(prs.ShapeType.ellipse, { x: 4.52, y: 2.55, w: 0.96, h: 0.96, fill: { color: C.teal }, shadow: makeShadow() });
  s.addText('VS', { x: 4.52, y: 2.55, w: 0.96, h: 0.96, fontSize: 13, color: C.white, bold: true, align: 'center', valign: 'middle' });

  // Right: proposed
  s.addShape(prs.ShapeType.rect, { x: 5.25, y: 1.42, w: 4.2, h: 3.4, fill: { color: 'f0fdf9' }, line: { color: '6ee7d4', width: 0.8 } });
  s.addText('RXOPS', { x: 5.45, y: 1.58, w: 3.8, h: 0.25, fontSize: 9, color: C.teal, bold: true, charSpacing: 2 });
  s.addText('Predictive & Unified', { x: 5.45, y: 1.88, w: 3.8, h: 0.35, fontSize: 15, color: '1e293b', bold: true });
  const proposed = [
    'AI flags Rx errors before order reaches the lab',
    'ML routes each order to the optimal lab in real-time',
    'SLA risk predicted 24 hrs in advance — intervene early',
    'Ops dashboard shows entire pipeline in one view',
    'One system pulls from all lab portals automatically',
  ];
  proposed.forEach((t, i) => {
    s.addText('✓  ' + t, { x: 5.45, y: 2.32 + i * 0.42, w: 3.8, h: 0.38, fontSize: 10.5, color: '166534', lineSpacingMultiple: 1.3 });
  });

  // callout box
  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 5.0, w: 8.9, h: 0.72, fill: { color: 'f0fdf9' }, line: { color: C.teal, width: 0.8 } });
  s.addText('Key insight: The cost of detecting an Rx error at the lab is 40x higher than detecting it before the order is accepted. RxOps moves the detection point from post-cut to pre-routing — turning a cost centre into a quality advantage.', {
    x: 0.75, y: 5.1, w: 8.5, h: 0.52, fontSize: 10.5, color: '166534', lineSpacingMultiple: 1.4,
  });
}

// ─── SLIDE 4: THE MECHANIC ────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE MECHANIC', { x: 0.55, y: 0.35, w: 9, h: 0.22, fontSize: 9, color: C.teal, bold: true, charSpacing: 3 });
  s.addText('Five Layers of AI Intelligence Across the Rx Lifecycle', {
    x: 0.55, y: 0.62, w: 9, h: 0.55, fontSize: 26, color: C.white, bold: true,
  });

  const steps = [
    { n: '01', title: 'Rx Ingestion', body: 'Customer uploads prescription. OCR extracts SPH, CYL, AXIS, ADD, PD values automatically. PDF and image formats supported.' },
    { n: '02', title: 'AI Validator', body: 'ML model checks for out-of-range values, boundary conditions (AXIS 0/180), ADD vs patient age mismatches, missing PD. 87% avg confidence score.' },
    { n: '03', title: 'Smart Lab Router', body: 'Per-order ML scoring across all lab partners on capacity, TAT, specialization match, geographic proximity, and SLA buffer. One-click confirm or override.' },
    { n: '04', title: 'At-Risk Monitor', body: 'Every open order is scored every 30 minutes. Orders predicted to miss SLA are surfaced 24 hrs early. Ops team sees intervention options directly on the dashboard.' },
    { n: '05', title: 'QC & Dispatch', body: 'Lab confirms completion. System auto-triggers dispatch and closes the loop. All metrics — TAT, accuracy, SLA hit rate — feed back into the routing ML model.' },
  ];

  steps.forEach((st, i) => {
    const x = 0.3 + i * 1.88;
    s.addShape(prs.ShapeType.rect, { x, y: 1.45, w: 1.68, h: 3.55, fill: { color: C.card }, line: { color: '1a2e42', width: 0.8 }, shadow: makeShadow() });
    s.addShape(prs.ShapeType.ellipse, { x: x + 0.44, y: 1.52, w: 0.8, h: 0.8, fill: { color: C.teal, transparency: 85 }, line: { color: C.teal, width: 0.8 } });
    s.addText(st.n, { x, y: 1.52, w: 1.68, h: 0.8, fontSize: 16, color: C.teal, bold: true, align: 'center', valign: 'middle' });
    s.addText(st.title, { x: x + 0.08, y: 2.42, w: 1.52, h: 0.44, fontSize: 11, color: C.white, bold: true, align: 'center', lineSpacingMultiple: 1.2 });
    s.addText(st.body, { x: x + 0.1, y: 2.9, w: 1.48, h: 1.85, fontSize: 8.8, color: C.muted, align: 'center', lineSpacingMultiple: 1.38 });
    // connector arrow (not last)
    if (i < 4) {
      s.addShape(prs.ShapeType.line, { x: x + 1.68, y: 2.22, w: 0.2, h: 0, line: { color: C.teal, width: 1, dashType: 'dash' } });
    }
  });

  // proof note
  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 5.12, w: 9.0, h: 0.62, fill: { color: '071525' }, line: { color: C.teal, width: 0.6 } });
  s.addText('PhonePe proof point: Same propensity-to-transact ML architecture used for the Rx risk model — real-time per-order scoring replacing manual rule-based decisions. Deployed at 350M+ MAU scale.', {
    x: 0.75, y: 5.22, w: 8.6, h: 0.42, fontSize: 10, color: C.teal, lineSpacingMultiple: 1.3,
  });
}

// ─── SLIDE 5: THE PRODUCT ─────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.white };

  s.addText('THE PRODUCT', { x: 0.55, y: 0.35, w: 9, h: 0.22, fontSize: 9, color: C.teal, bold: true, charSpacing: 3 });
  s.addText('4 Console Views — One Unified Operations Layer', {
    x: 0.55, y: 0.62, w: 9, h: 0.5, fontSize: 26, color: '0f172a', bold: true,
  });

  const screens = [
    { n: '01', title: 'Ops Dashboard', desc: 'Live order pipeline with AI risk scores. KPI cards: at-risk count, avg processing time, lab utilization. One-click intervention on delayed orders.', mock: ['Orders: 1,247', 'At-Risk: 34 🔴', 'Gulf Optics: 71%', 'Vision Labs: 89% ⚠', '─────────────', 'EW-284628 HIGH', 'EW-284673 HIGH'] },
    { n: '02', title: 'Rx Validator', desc: 'AI scans every uploaded prescription. Flags out-of-range values, boundary conditions, and age-ADD mismatches. Approve, request clarification, or override.', mock: ['OD: SPH -2.50', 'CYL -0.75 AXIS 180°⚠', 'ADD +2.00 OS: SPH -2.25', '────────────', '2 Flags Found', 'Confidence: 87%', '[ Approve ] [ Flag ]'] },
    { n: '03', title: 'Smart Lab Router', desc: 'Per-order ML recommendation across all lab partners. Shows capacity, TAT, SLA buffer, and specialization match score. One click to confirm routing.', mock: ['Order EW-284711', 'Complexity: Medium', '────────────', 'Gulf Optics 94% ✓', 'Vision Labs  71%', 'Al Noor     43%', '[ Confirm Route ]'] },
    { n: '04', title: 'Admin Config', desc: 'Routing rules engine with toggle controls. Alert thresholds for SLA, capacity, Rx confidence. Lab registry with live capacity feeds from all partners.', mock: ['Routing Rules', '────────────', 'Nearest lab ✓ ON', 'Cap cap 90% ✓ ON', 'Progressive cert ON', '────────────', 'Alert: 24hr buffer'] },
  ];

  screens.forEach((sc, i) => {
    const x = 0.35 + i * 2.38;
    s.addShape(prs.ShapeType.rect, { x, y: 1.3, w: 2.2, h: 4.4, fill: { color: 'f8fafc' }, line: { color: 'd1d5db', width: 0.6 }, shadow: makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y: 1.3, w: 2.2, h: 0.32, fill: { color: C.teal } });
    s.addText(sc.n + ' — ' + sc.title, { x: x + 0.08, y: 1.3, w: 2.04, h: 0.32, fontSize: 9, color: C.white, bold: true, valign: 'middle' });
    s.addText(sc.desc, { x: x + 0.1, y: 1.68, w: 2.0, h: 1.1, fontSize: 8.8, color: '475569', lineSpacingMultiple: 1.35 });
    // ASCII mock
    s.addShape(prs.ShapeType.rect, { x: x + 0.1, y: 2.82, w: 2.0, h: 1.75, fill: { color: '071525' }, line: { color: '1a2e42', width: 0.5 } });
    s.addText(sc.mock.join('\n'), { x: x + 0.14, y: 2.88, w: 1.92, h: 1.65, fontSize: 7.8, color: '94a3b8', fontFace: 'Courier New', lineSpacingMultiple: 1.45 });
  });
}

// ─── SLIDE 6: IMPACT & ROI ────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.hero };

  s.addText('IMPACT & ROI', { x: 0.55, y: 0.35, w: 9, h: 0.22, fontSize: 9, color: C.teal, bold: true, charSpacing: 3 });
  s.addText('Projected Impact — Built on PhonePe Proof Points at Scale', {
    x: 0.55, y: 0.62, w: 9, h: 0.55, fontSize: 26, color: C.white, bold: true,
  });

  const leftMetrics = [
    { val: '↑ 60%', lbl: 'SLA on-time rate', sub: 'From current ~70% baseline' },
    { val: '↓ 80%', lbl: 'Rx validation errors', sub: 'AI catch rate at 87% confidence' },
    { val: '−1.2 days', lbl: 'Avg processing time', sub: 'Validation step: 8.2h → 3.1h' },
    { val: '↑ 4×', lbl: 'Orders per ops FTE', sub: 'Less firefighting, more orchestration' },
  ];

  const rightMetrics = [
    { val: 'AED 250', lbl: 'Saved per avoided remake', sub: 'Target: 30% reduction in remakes' },
    { val: '↓ 65%', lbl: 'Customer SLA complaints', sub: 'Proactive comms before breach' },
    { val: 'AED 1.2M+', lbl: 'Annual rework cost saved', sub: 'At current order volume (est.)' },
    { val: '< 8 wks', lbl: 'Time to first measurable ROI', sub: 'Pilot with 200 orders/day' },
  ];

  // Left column
  s.addText('OPS EFFICIENCY', { x: 0.55, y: 1.4, w: 4.2, h: 0.22, fontSize: 8.5, color: C.teal, bold: true, charSpacing: 2 });
  leftMetrics.forEach((m, i) => {
    s.addShape(prs.ShapeType.rect, { x: 0.55, y: 1.68 + i * 0.88, w: 4.2, h: 0.78, fill: { color: C.card }, line: { color: '1a2e42', width: 0.6 } });
    s.addText(m.val, { x: 0.7, y: 1.73 + i * 0.88, w: 1.6, h: 0.68, fontSize: 24, color: C.teal, bold: true, valign: 'middle' });
    s.addText(m.lbl + '\n' + m.sub, { x: 2.35, y: 1.73 + i * 0.88, w: 2.25, h: 0.68, fontSize: 9.5, color: C.muted, lineSpacingMultiple: 1.35, valign: 'middle' });
  });

  // Right column
  s.addText('BUSINESS ROI', { x: 5.25, y: 1.4, w: 4.2, h: 0.22, fontSize: 8.5, color: C.amber, bold: true, charSpacing: 2 });
  rightMetrics.forEach((m, i) => {
    s.addShape(prs.ShapeType.rect, { x: 5.25, y: 1.68 + i * 0.88, w: 4.2, h: 0.78, fill: { color: C.card }, line: { color: '1a2e42', width: 0.6 } });
    s.addText(m.val, { x: 5.4, y: 1.73 + i * 0.88, w: 1.8, h: 0.68, fontSize: 24, color: C.amber, bold: true, valign: 'middle' });
    s.addText(m.lbl + '\n' + m.sub, { x: 7.25, y: 1.73 + i * 0.88, w: 2.05, h: 0.68, fontSize: 9.5, color: C.muted, lineSpacingMultiple: 1.35, valign: 'middle' });
  });

  // insight box
  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 5.28, w: 8.9, h: 0.72, fill: { color: '050f1a' }, line: { color: C.teal, width: 0.8 } });
  s.addText('At PhonePe: rebuilding the offer ops workflow reduced manual TAT from 2 days to 30 minutes. The same process re-engineering approach applied to Rx validation is the core of RxOps.', {
    x: 0.75, y: 5.38, w: 8.5, h: 0.52, fontSize: 10.5, color: C.teal, lineSpacingMultiple: 1.4,
  });
}

// ─── SLIDE 7: PROOF OF WORK ───────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.white };

  s.addText('PROOF OF WORK', { x: 0.55, y: 0.35, w: 9, h: 0.22, fontSize: 9, color: C.teal, bold: true, charSpacing: 3 });
  s.addText('I Built This Architecture. Here Is the Evidence.', {
    x: 0.55, y: 0.62, w: 9, h: 0.55, fontSize: 26, color: '0f172a', bold: true,
  });

  // Left: what I built at PhonePe
  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 1.38, w: 4.3, h: 3.9, fill: { color: C.dark }, line: { color: '1a2e42', width: 0.8 } });
  s.addText('WHAT I BUILT AT PHONEPE', { x: 0.75, y: 1.55, w: 3.9, h: 0.22, fontSize: 8.5, color: C.teal, bold: true, charSpacing: 2 });
  const built = [
    ['Offer ops workflow rebuild', 'Manual 2-day TAT → 30-min automated pipeline. Identical to RxOps validation step redesign.'],
    ['Propensity-to-Transact ML model', 'Real-time per-user scoring replacing rule-based decisions. Same architecture as Rx risk scoring.'],
    ['Multi-vendor lab-equivalent platform', '500+ brand partners on a self-serve system with segmentation and performance scorecards.'],
    ['B2B merchant routing & onboarding', '5,000+ merchants, 30-min onboarding TAT. Same routing-decision logic as smart lab router.'],
  ];
  built.forEach(([title, desc], i) => {
    s.addText('■  ' + title, { x: 0.75, y: 1.88 + i * 0.84, w: 3.9, h: 0.24, fontSize: 10.5, color: C.white, bold: true });
    s.addText(desc, { x: 0.95, y: 2.14 + i * 0.84, w: 3.7, h: 0.52, fontSize: 9.5, color: C.muted, lineSpacingMultiple: 1.35 });
  });

  // Right: how it maps to eyewa JD
  s.addShape(prs.ShapeType.rect, { x: 5.15, y: 1.38, w: 4.3, h: 3.9, fill: { color: 'f0fdf9' }, line: { color: '6ee7d4', width: 0.8 } });
  s.addText('HOW IT MAPS TO EYEWA', { x: 5.35, y: 1.55, w: 3.9, h: 0.22, fontSize: 8.5, color: '166534', bold: true, charSpacing: 2 });
  const maps = [
    ['Roadmap for internal ops tools', 'The exact mandate — Rx ops, lab routing, SLA alerting are internal tooling challenges.'],
    ['AI and automation use cases', 'Prescription validation and lab routing are AI use cases I defined, spec-ed, and shipped.'],
    ['Data-driven product decisions', 'ML model selection, threshold tuning, A/B test design — all owned end-to-end.'],
    ['Cross-functional stakeholder mgmt', 'Aligned engineering, ops, brand partners, and data science across competing priorities.'],
  ];
  maps.forEach(([title, desc], i) => {
    s.addText('✓  ' + title, { x: 5.35, y: 1.88 + i * 0.84, w: 3.9, h: 0.24, fontSize: 10.5, color: '166534', bold: true });
    s.addText(desc, { x: 5.55, y: 2.14 + i * 0.84, w: 3.7, h: 0.52, fontSize: 9.5, color: '475569', lineSpacingMultiple: 1.35 });
  });

  // quote box
  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 5.42, w: 8.9, h: 0.72, fill: { color: 'f0fdf9' }, line: { color: C.teal, width: 0.8 } });
  s.addText('"The challenge at eyewa is not understanding the problem — it is building the intelligence layer to solve it at scale. That is exactly what I did at PhonePe."', {
    x: 0.75, y: 5.52, w: 8.5, h: 0.52, fontSize: 10.5, color: '166534', italic: true, lineSpacingMultiple: 1.35,
  });
}

// ─── SLIDE 8: ROLLOUT PLAN ────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('ROLLOUT PLAN', { x: 0.55, y: 0.35, w: 9, h: 0.22, fontSize: 9, color: C.teal, bold: true, charSpacing: 3 });
  s.addText('Three Phases to Full Deployment — First Value in 6 Weeks', {
    x: 0.55, y: 0.62, w: 9, h: 0.55, fontSize: 26, color: C.white, bold: true,
  });

  const phases = [
    { phase: 'Phase 1', period: 'Weeks 1–6', title: 'Rx Validator Pilot', color: C.teal, items: ['Deploy AI Rx Validator on 100% of new Rx orders', 'Run in shadow mode first — flag without blocking', 'Measure: false positive rate, validation time reduction', 'A/B vs manual review on a held-out order set'] },
    { phase: 'Phase 2', period: 'Weeks 7–14', title: 'Smart Lab Router', color: C.amber, items: ['Integrate capacity feeds from all 4 lab partners', 'Deploy ML routing for single-vision orders first', 'Expand to progressive and hi-index after 2-week bake', 'Launch at-risk SLA dashboard for ops team'] },
    { phase: 'Phase 3', period: 'Weeks 15–20', title: 'Full RxOps Console', color: C.green, items: ['Full production rollout across all order types', 'Ops dashboard live for all team members', 'Predictive SLA alerting 24 hrs in advance', 'Feedback loop: QC data back into routing model'] },
  ];

  phases.forEach((ph, i) => {
    const x = 0.45 + i * 3.18;
    s.addShape(prs.ShapeType.rect, { x, y: 1.42, w: 2.95, h: 3.5, fill: { color: C.card }, line: { color: ph.color, width: 1.2 }, shadow: makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y: 1.42, w: 2.95, h: 0.36, fill: { color: ph.color, transparency: 88 }, line: { color: ph.color, width: 0.6 } });
    s.addText(ph.phase + '  ·  ' + ph.period, { x, y: 1.42, w: 2.95, h: 0.36, fontSize: 9, color: ph.color, bold: true, align: 'center', valign: 'middle' });
    s.addText(ph.title, { x: x + 0.12, y: 1.88, w: 2.7, h: 0.38, fontSize: 13.5, color: C.white, bold: true });
    ph.items.forEach((item, j) => {
      s.addText('→  ' + item, { x: x + 0.12, y: 2.34 + j * 0.56, w: 2.7, h: 0.48, fontSize: 9.8, color: C.muted, lineSpacingMultiple: 1.3 });
    });
  });

  // Ask box
  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 5.08, w: 9.0, h: 0.9, fill: { color: C.card }, line: { color: C.teal, width: 0.8 } });
  s.addText('WHAT I NEED TO BUILD THIS', { x: 0.75, y: 5.18, w: 3.5, h: 0.2, fontSize: 8.5, color: C.teal, bold: true, charSpacing: 2 });
  s.addText('Access to Rx order data  ·  Lab portal API credentials  ·  1 backend engineer  ·  1 data analyst for model training  ·  Ops team as pilot users (weeks 1–6)', {
    x: 0.75, y: 5.4, w: 8.6, h: 0.45, fontSize: 10.5, color: C.muted, lineSpacingMultiple: 1.35,
  });

  // Contact footer
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade', {
    x: 0.55, y: 6.9, w: 9.0, h: 0.25, fontSize: 9, color: '475569', align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/eyewa-rxops-deck.pptx' })
  .then(() => console.log('✅ eyewa-rxops-deck.pptx written'))
  .catch(e => console.error(e));
