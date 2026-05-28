const pptxgen = require('pptxgenjs');

const C = {
  dark:   '0C0408',
  hero:   '2C1020',
  red:    'E8001B',
  gold:   'C8952A',
  white:  'FFFFFF',
  lgray:  'F6F1F3',
  green:  '16A34A',
  orange: 'EA580C',
  ink:    '18100E',
  muted:  '9B8B90',
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 8, offset: 2, angle: 135, opacity: 0.12 });
const makeCardShadow = () => ({ type: 'outer', color: '000000', blur: 5, offset: 1, angle: 135, opacity: 0.10 });

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Ajay Avaghade';
pres.title = 'MAF OpsIQ — CEO Pitch Deck';

// ─────────────────────────────────────────────────────────────
// SLIDE 1 — COVER
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.dark };

  for (let i = 0; i < 9; i++) {
    s.addShape(pres.shapes.RECTANGLE, {
      x: -1.2 + i * 1.5, y: -0.5, w: 0.06, h: 9,
      fill: { color: '200A14' }, line: { color: '200A14' }, rotate: 25,
    });
  }

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.12, h: 5.625,
    fill: { color: C.red }, line: { color: C.red },
  });

  s.addText('MAF RETAIL  ·  MARKETPLACE OPERATIONS  ·  CASE STUDY 03', {
    x: 0.35, y: 0.32, w: 7.5, h: 0.26,
    fontSize: 7.5, fontFace: 'Calibri', color: C.muted, charSpacing: 2, margin: 0,
  });
  s.addText('OpsIQ', {
    x: 0.35, y: 0.9, w: 8, h: 1.4,
    fontSize: 72, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });
  s.addText('Replacing 4-Week Excel Lag with 4-Hour Anomaly Detection', {
    x: 0.35, y: 2.32, w: 7.5, h: 0.56,
    fontSize: 17, fontFace: 'Calibri', color: C.gold, margin: 0,
  });
  s.addText('Ajay Avaghade  ·  Product Manager', {
    x: 0.35, y: 2.96, w: 6, h: 0.36,
    fontSize: 11, fontFace: 'Calibri', color: C.muted, margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.9, y: 3.5, w: 2.8, h: 1.65,
    fill: { color: '180A12' }, line: { color: C.gold },
  });
  s.addText('−82%', {
    x: 6.9, y: 3.62, w: 2.8, h: 0.78,
    fontSize: 42, fontFace: 'Calibri', color: C.gold, bold: true, align: 'center', margin: 0,
  });
  s.addText('ops issue fix cycle', {
    x: 6.9, y: 4.42, w: 2.8, h: 0.5,
    fontSize: 9, fontFace: 'Calibri', color: C.muted, align: 'center', margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.36, w: 10, h: 0.27,
    fill: { color: '120810' }, line: { color: '120810' },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.36, w: 1.8, h: 0.27,
    fill: { color: C.red }, line: { color: C.red },
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.dark };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.12, h: 5.625,
    fill: { color: C.red }, line: { color: C.red },
  });
  s.addText('THE PROBLEM', {
    x: 0.35, y: 0.3, w: 4, h: 0.26,
    fontSize: 8, fontFace: 'Calibri', color: C.red, bold: true, charSpacing: 3, margin: 0,
  });
  s.addText('A Fill Rate Crisis Detected in 4 Days Is a\nCustomer Experience Crisis That Already Happened.', {
    x: 0.35, y: 0.6, w: 9.3, h: 1.1,
    fontSize: 22, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });

  const cols = [
    { icon: '⏰', stat: '4 Days',     label: 'Avg ops issue detection lag',  sub: 'Weekly Excel reports = 4-7 day visibility gap' },
    { icon: '🔥', stat: '6 Siloed',   label: 'Dashboards ops team monitors', sub: 'GMV, fill rate, SLA, returns, seller health, NPS — all separate' },
    { icon: '📉', stat: '−0.8pt',     label: 'NPS impact of a 4-hour drop',  sub: 'By the time it\'s in Excel, 2,000 customers already complained' },
  ];
  const cW = 2.72, cGap = 0.24, cStartX = 0.58;
  cols.forEach((col, i) => {
    const cx = cStartX + i * (cW + cGap);
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: 1.85, w: cW, h: 2.15,
      fill: { color: '180A12' }, line: { color: '3A1A28' }, shadow: makeCardShadow(),
    });
    s.addText(col.icon, { x: cx + 0.15, y: 1.92, w: 0.5, h: 0.44, fontSize: 20, margin: 0 });
    s.addText(col.stat, {
      x: cx + 0.12, y: 2.42, w: cW - 0.2, h: 0.62,
      fontSize: 28, fontFace: 'Calibri', color: C.gold, bold: true, margin: 0,
    });
    s.addText(col.label, {
      x: cx + 0.12, y: 3.07, w: cW - 0.2, h: 0.38,
      fontSize: 9.5, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
    });
    s.addText(col.sub, {
      x: cx + 0.12, y: 3.48, w: cW - 0.2, h: 0.42,
      fontSize: 8, fontFace: 'Calibri', color: C.muted, margin: 0,
    });
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 4.18, w: 9.3, h: 0.9,
    fill: { color: '180A12' }, line: { color: '3A1020' },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 4.18, w: 0.06, h: 0.9,
    fill: { color: C.gold }, line: { color: C.gold },
  });
  s.addText("The Chilled Foods fill rate drop from 94%→61% happened in 4 hours. With today's Excel-based ops model, the VP finds out on Friday. With OpsIQ, the ops team gets an alert in 4 minutes.", {
    x: 0.55, y: 4.23, w: 8.9, h: 0.8,
    fontSize: 9.5, fontFace: 'Calibri', color: 'C0A8B0', margin: 0, valign: 'middle',
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 3 — THE INSIGHT
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.08,
    fill: { color: C.red }, line: { color: C.red },
  });
  s.addText('THE INSIGHT', {
    x: 0.5, y: 0.2, w: 4, h: 0.26,
    fontSize: 8, fontFace: 'Calibri', color: C.red, bold: true, charSpacing: 3, margin: 0,
  });
  s.addText("The Fix Is Not a Better Dashboard.\nIt's a Unified Intelligence Layer with a Brain.", {
    x: 0.5, y: 0.5, w: 9, h: 0.95,
    fontSize: 24, fontFace: 'Calibri', color: C.ink, bold: true, margin: 0,
  });

  const colY = 1.6, colH = 2.62, lX = 0.5, rX = 5.22, cW = 4.28;

  s.addShape(pres.shapes.RECTANGLE, {
    x: lX, y: colY, w: cW, h: colH,
    fill: { color: C.dark }, line: { color: '2A1020' }, shadow: makeCardShadow(),
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: lX, y: colY, w: cW, h: 0.36,
    fill: { color: C.hero }, line: { color: C.hero },
  });
  s.addText('❌  Today', {
    x: lX + 0.15, y: colY + 0.06, w: cW - 0.2, h: 0.28,
    fontSize: 10.5, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });
  s.addText([
    { text: '· 6 separate dashboards, none talking to each other', options: { breakLine: true } },
    { text: '· Weekly Excel export — issues aged 4-7 days before action', options: { breakLine: true } },
    { text: '· Ops team manually correlates signals across systems', options: { breakLine: true } },
    { text: '· Jira tickets created after VP escalation, not before' },
  ], {
    x: lX + 0.15, y: colY + 0.52, w: cW - 0.25, h: colH - 0.68,
    fontSize: 9.5, fontFace: 'Calibri', color: 'C0A8B0', margin: 0, valign: 'top',
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: rX, y: colY, w: cW, h: colH,
    fill: { color: C.lgray }, line: { color: 'DDD0D4' }, shadow: makeCardShadow(),
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: rX, y: colY, w: cW, h: 0.36,
    fill: { color: C.green }, line: { color: C.green },
  });
  s.addText('✅  OpsIQ', {
    x: rX + 0.15, y: colY + 0.06, w: cW - 0.2, h: 0.28,
    fontSize: 10.5, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });
  s.addText([
    { text: '· Single health dashboard — all KPIs unified, anomaly feed live', options: { breakLine: true } },
    { text: '· ML anomaly detection fires in <30 min of signal emergence', options: { breakLine: true } },
    { text: '· AI root-cause analysis: 3 ranked drivers with confidence scores', options: { breakLine: true } },
    { text: '· Auto-generated Jira backlog — tickets ready before ops manager asks' },
  ], {
    x: rX + 0.15, y: colY + 0.52, w: cW - 0.25, h: colH - 0.68,
    fontSize: 9.5, fontFace: 'Calibri', color: C.ink, margin: 0, valign: 'top',
  });

  s.addShape(pres.shapes.OVAL, {
    x: 4.72, y: 2.58, w: 0.56, h: 0.56,
    fill: { color: C.red }, line: { color: C.white }, shadow: makeCardShadow(),
  });
  s.addText('→', {
    x: 4.72, y: 2.58, w: 0.56, h: 0.56,
    fontSize: 18, fontFace: 'Calibri', color: C.white,
    bold: true, align: 'center', valign: 'middle', margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.38, w: 9, h: 0.84,
    fill: { color: '180A12' }, line: { color: '3A1020' },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.38, w: 0.06, h: 0.84,
    fill: { color: C.red }, line: { color: C.red },
  });
  s.addText('I built this exact transition at PhonePe: 6 siloed analyst workflows → unified intelligence layer. TAT 2 days → 4 hours. −68% escalation tickets in 60 days.', {
    x: 0.7, y: 4.43, w: 8.6, h: 0.74,
    fontSize: 10, fontFace: 'Calibri', color: 'C0A8B0', bold: true, margin: 0, valign: 'middle',
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 4 — THE MECHANIC
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.hero };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.12, h: 5.625,
    fill: { color: C.gold }, line: { color: C.gold },
  });
  s.addText('THE MECHANIC', {
    x: 0.35, y: 0.28, w: 5, h: 0.26,
    fontSize: 8, fontFace: 'Calibri', color: C.gold, bold: true, charSpacing: 3, margin: 0,
  });
  s.addText('Detect → Diagnose → Act — All in Under 30 Minutes', {
    x: 0.35, y: 0.56, w: 9.3, h: 0.72,
    fontSize: 26, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });

  const steps = [
    { num: '1', title: 'Data Ingestion',        body: 'OpsIQ ingests real-time signals: order data, inventory feeds, seller SLA logs, NPS verbatims, hub capacity telemetry — unified into single pipeline.' },
    { num: '2', title: 'Anomaly Detection',      body: 'ML model: 3σ deviation from 30-day rolling baseline triggers alert. Category × hub × time-of-day segmented. Severity: Critical / High / Watch.' },
    { num: '3', title: 'NLP Classification',     body: 'Verbatim alerts NLP-tagged across 8 issue types: stock-out, SLA breach, seller inactive, routing error, hub overload, taxonomy error, pricing, quality.' },
    { num: '4', title: 'AI Root Cause',          body: '3 ranked root causes with confidence scores. Evidence cited from data. Generated in <5 seconds. Ops team validates — not investigates.' },
    { num: '5', title: 'Auto-Jira + Action',     body: 'Each root cause maps to an auto-drafted Jira ticket with owner, SLA, and severity. 1-click create. No manual ticket writing. Fix cycle: 4h vs 4 days.' },
  ];
  const sW = 1.66, sGap = 0.16, sX = 0.35, sY = 1.52, sH = 3.0;

  s.addShape(pres.shapes.LINE, {
    x: sX + sW * 0.5, y: 1.84, w: (sW + sGap) * 4, h: 0,
    line: { color: '5A3040', width: 2, dashType: 'dash' },
  });

  const stepColors = [C.red, C.orange, C.gold, C.green, '0070C9'];
  steps.forEach((step, i) => {
    const px = sX + i * (sW + sGap);
    s.addShape(pres.shapes.RECTANGLE, {
      x: px, y: sY, w: sW, h: sH,
      fill: { color: '180A12' }, line: { color: '3A1828' }, shadow: makeCardShadow(),
    });
    s.addShape(pres.shapes.OVAL, {
      x: px + sW / 2 - 0.22, y: sY - 0.3, w: 0.44, h: 0.44,
      fill: { color: stepColors[i] }, line: { color: C.hero },
    });
    s.addText(step.num, {
      x: px + sW / 2 - 0.22, y: sY - 0.3, w: 0.44, h: 0.44,
      fontSize: 11, fontFace: 'Calibri', color: C.white,
      bold: true, align: 'center', valign: 'middle', margin: 0,
    });
    s.addText(step.title, {
      x: px + 0.1, y: sY + 0.14, w: sW - 0.17, h: 0.38,
      fontSize: 9.5, fontFace: 'Calibri', color: C.gold, bold: true, margin: 0,
    });
    s.addShape(pres.shapes.LINE, {
      x: px + 0.1, y: sY + 0.56, w: sW - 0.2, h: 0,
      line: { color: '3A1828', width: 1 },
    });
    s.addText(step.body, {
      x: px + 0.1, y: sY + 0.65, w: sW - 0.17, h: sH - 0.82,
      fontSize: 8.5, fontFace: 'Calibri', color: 'C0A8B0', margin: 0, valign: 'top',
    });
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 4.78, w: 9.3, h: 0.47,
    fill: { color: '120810' }, line: { color: '2A1020' },
  });
  s.addText('Built this at PhonePe: 6 siloed ops workflows → unified layer. TAT 2d→4h. −68% escalation tickets in 60 days.', {
    x: 0.5, y: 4.82, w: 9.0, h: 0.39,
    fontSize: 8.5, fontFace: 'Calibri', color: C.muted, italic: true, margin: 0, valign: 'middle',
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 5 — THE PRODUCT
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.08,
    fill: { color: C.red }, line: { color: C.red },
  });
  s.addText('THE PRODUCT', {
    x: 0.5, y: 0.2, w: 4, h: 0.26,
    fontSize: 8, fontFace: 'Calibri', color: C.red, bold: true, charSpacing: 3, margin: 0,
  });
  s.addText('4 Screen States — From Signal to Resolution', {
    x: 0.5, y: 0.5, w: 9, h: 0.62,
    fontSize: 24, fontFace: 'Calibri', color: C.ink, bold: true, margin: 0,
  });

  const cards = [
    { num: '01', title: 'Health Dashboard',  body: 'Real-time KPI tiles: GMV, fill rate, SLA compliance, return rate. Live anomaly feed with severity badges. 1-click drill-down per anomaly. Red/orange/green status at a glance.' },
    { num: '02', title: 'Anomaly Drill-Down', body: 'Alert detail: time detected, orders at risk, GMV exposure, NLP issue tags. Event timeline showing signal emergence. Impact quantified before ops team acts.' },
    { num: '03', title: 'AI Root Cause',      body: '3 ranked root causes with confidence scores (84%, 71%, 58%). Evidence summary per driver. Recommended action per cause with 1-click approve/escalate/create Jira.' },
    { num: '04', title: 'Escalation Console', body: 'Seller SLA breach log. Filter by severity, category, hub. Auto-generated Jira tickets ready to create. Seller communication auto-drafted. Fix cycle tracking.' },
  ];
  const accentColors = [C.red, C.orange, C.gold, C.green];
  const cW = 2.1, cGap = 0.16, cStartX = 0.42, cY = 1.3, cH = 3.48;

  cards.forEach((card, i) => {
    const cx = cStartX + i * (cW + cGap);
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cY, w: cW, h: cH,
      fill: { color: C.lgray }, line: { color: 'DDD0D4' }, shadow: makeCardShadow(),
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cY, w: cW, h: 0.06,
      fill: { color: accentColors[i] }, line: { color: accentColors[i] },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx + 0.12, y: cY + 0.18, w: 0.38, h: 0.26,
      fill: { color: accentColors[i] }, line: { color: accentColors[i] },
    });
    s.addText(card.num, {
      x: cx + 0.12, y: cY + 0.18, w: 0.38, h: 0.26,
      fontSize: 8.5, fontFace: 'Calibri', color: C.white,
      bold: true, align: 'center', valign: 'middle', margin: 0,
    });
    s.addText(card.title, {
      x: cx + 0.12, y: cY + 0.54, w: cW - 0.2, h: 0.44,
      fontSize: 10.5, fontFace: 'Calibri', color: C.ink, bold: true, margin: 0,
    });
    s.addShape(pres.shapes.LINE, {
      x: cx + 0.12, y: cY + 1.04, w: cW - 0.24, h: 0,
      line: { color: 'DDD0D4', width: 1 },
    });
    s.addText(card.body, {
      x: cx + 0.12, y: cY + 1.18, w: cW - 0.2, h: cH - 1.36,
      fontSize: 9, fontFace: 'Calibri', color: '4A3840', margin: 0, valign: 'top',
    });
  });

  s.addText('Prototype: maf-opsiq-prototype.html  ·  All 4 states interactive', {
    x: 0.5, y: 4.98, w: 9, h: 0.28,
    fontSize: 7.5, fontFace: 'Calibri', color: C.muted, italic: true, align: 'center', margin: 0,
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.dark };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.12, h: 5.625,
    fill: { color: C.red }, line: { color: C.red },
  });
  s.addText('IMPACT & ROI', {
    x: 0.35, y: 0.28, w: 5, h: 0.26,
    fontSize: 8, fontFace: 'Calibri', color: C.red, bold: true, charSpacing: 3, margin: 0,
  });
  s.addText('4 Days → 30 Minutes. The Same Fix. Zero Additional Cost.', {
    x: 0.35, y: 0.56, w: 9.3, h: 0.58,
    fontSize: 22, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });

  const sections = [
    {
      title: 'Ops Team Impact', color: C.gold, x: 0.35,
      metrics: [
        { val: '4 min',  lbl: 'Anomaly detection time', sub: 'vs 4-7 day Excel lag today' },
        { val: '↓ 82%',  lbl: 'Issue fix cycle time', sub: 'from days to hours' },
        { val: '↑ 68%',  lbl: 'Ops capacity freed', sub: 'reactive firefighting → strategy' },
        { val: '3',      lbl: 'Root causes diagnosed', sub: 'automatically per anomaly' },
      ],
    },
    {
      title: 'Business Impact', color: C.green, x: 5.1,
      metrics: [
        { val: '↓ 0.4pt',  lbl: 'NPS impact prevented', sub: 'per critical anomaly caught early' },
        { val: 'AED 68K',  lbl: 'GMV exposure resolved', sub: 'per Chilled Foods incident' },
        { val: '↓ 68%',    lbl: 'VP escalation tickets', sub: 'PhonePe baseline — same architecture' },
        { val: '24h',      lbl: 'Anomaly detection now', sub: 'vs 4-week Excel lag today' },
      ],
    },
  ];

  sections.forEach(sec => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: sec.x, y: 1.3, w: 4.35, h: 0.3,
      fill: { color: '180A12' }, line: { color: '180A12' },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: sec.x, y: 1.3, w: 0.06, h: 0.3,
      fill: { color: sec.color }, line: { color: sec.color },
    });
    s.addText(sec.title, {
      x: sec.x + 0.15, y: 1.32, w: 4.0, h: 0.26,
      fontSize: 9.5, fontFace: 'Calibri', color: sec.color, bold: true, margin: 0,
    });
    const mW = 2.1, mH = 1.1, mGap = 0.15;
    sec.metrics.forEach((m, i) => {
      const row = Math.floor(i / 2), col = i % 2;
      const mx = sec.x + col * (mW + mGap);
      const my = 1.72 + row * (mH + 0.1);
      s.addShape(pres.shapes.RECTANGLE, {
        x: mx, y: my, w: mW, h: mH,
        fill: { color: '160A12' }, line: { color: '2E1422' }, shadow: makeCardShadow(),
      });
      s.addText(m.val, {
        x: mx + 0.1, y: my + 0.08, w: mW - 0.18, h: 0.44,
        fontSize: 20, fontFace: 'Calibri', color: sec.color, bold: true, margin: 0,
      });
      s.addText(m.lbl, {
        x: mx + 0.1, y: my + 0.52, w: mW - 0.18, h: 0.3,
        fontSize: 8.5, fontFace: 'Calibri', color: C.white, margin: 0,
      });
      s.addText(m.sub, {
        x: mx + 0.1, y: my + 0.82, w: mW - 0.18, h: 0.22,
        fontSize: 7.5, fontFace: 'Calibri', color: C.muted, margin: 0,
      });
    });
  });

  s.addShape(pres.shapes.LINE, {
    x: 5.0, y: 1.3, w: 0, h: 3.42,
    line: { color: '3A1828', width: 1, dashType: 'dash' },
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 4.92, w: 9.3, h: 0.4,
    fill: { color: '180A12' }, line: { color: '2A1020' },
  });
  s.addText("I built this transition at PhonePe. The architecture is proven. This is the same system applied to MAF's marketplace ops problem.", {
    x: 0.5, y: 4.95, w: 9.0, h: 0.34,
    fontSize: 8.5, fontFace: 'Calibri', color: C.muted, italic: true, margin: 0, valign: 'middle',
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.08,
    fill: { color: C.red }, line: { color: C.red },
  });
  s.addText('PROOF OF WORK', {
    x: 0.5, y: 0.2, w: 5, h: 0.26,
    fontSize: 8, fontFace: 'Calibri', color: C.red, bold: true, charSpacing: 3, margin: 0,
  });
  s.addText('I Replaced 6 Siloed Workflows with This. At PhonePe.', {
    x: 0.5, y: 0.5, w: 9, h: 0.62,
    fontSize: 24, fontFace: 'Calibri', color: C.ink, bold: true, margin: 0,
  });

  const lX = 0.5, rX = 5.2, pW = 4.3, pY = 1.35, pH = 2.9;

  s.addShape(pres.shapes.RECTANGLE, {
    x: lX, y: pY, w: pW, h: pH,
    fill: { color: C.dark }, line: { color: '2A1020' }, shadow: makeCardShadow(),
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: lX, y: pY, w: pW, h: 0.36,
    fill: { color: C.hero }, line: { color: C.hero },
  });
  s.addText('PhonePe — What I Built', {
    x: lX + 0.15, y: pY + 0.06, w: pW - 0.2, h: 0.28,
    fontSize: 9.5, fontFace: 'Calibri', color: C.gold, bold: true, margin: 0,
  });
  s.addText([
    { text: '· Inherited 6 fragmented analyst workflows across Offers, Merchants, GMV, Escalations, Compliance, and NPS', options: { breakLine: true } },
    { text: '· Defined KPIs, instrumented full pipeline before build began', options: { breakLine: true } },
    { text: '· Replaced all 6 with unified ops intelligence layer', options: { breakLine: true } },
    { text: '· Result: TAT 2 days → 4 hours  ·  −68% escalation tickets  ·  60-day deployment' },
  ], {
    x: lX + 0.15, y: pY + 0.5, w: pW - 0.25, h: pH - 0.64,
    fontSize: 8.5, fontFace: 'Calibri', color: 'C0A8B0', margin: 0, valign: 'top',
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: rX, y: pY, w: pW, h: pH,
    fill: { color: C.lgray }, line: { color: 'DDD0D4' }, shadow: makeCardShadow(),
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: rX, y: pY, w: pW, h: 0.36,
    fill: { color: C.green }, line: { color: C.green },
  });
  s.addText('Direct Map to OpsIQ', {
    x: rX + 0.15, y: pY + 0.06, w: pW - 0.2, h: 0.28,
    fontSize: 9.5, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });
  s.addText([
    { text: '→ OpsIQ unified health dashboard (replaces 6 siloed dashboards)', options: { breakLine: true } },
    { text: '→ OpsIQ KPI instrumentation + anomaly detection engine', options: { breakLine: true } },
    { text: '→ OpsIQ AI root-cause + auto-Jira (replaces manual investigation)', options: { breakLine: true } },
    { text: '→ OpsIQ 30-min detection vs 4-day Excel lag' },
  ], {
    x: rX + 0.15, y: pY + 0.5, w: pW - 0.25, h: pH - 0.64,
    fontSize: 8.5, fontFace: 'Calibri', color: C.ink, margin: 0, valign: 'top',
  });

  s.addShape(pres.shapes.OVAL, {
    x: 4.72, y: 2.58, w: 0.56, h: 0.56,
    fill: { color: C.red }, line: { color: C.white }, shadow: makeCardShadow(),
  });
  s.addText('→', {
    x: 4.72, y: 2.58, w: 0.56, h: 0.56,
    fontSize: 18, fontFace: 'Calibri', color: C.white,
    bold: true, align: 'center', valign: 'middle', margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.42, w: 9, h: 0.82,
    fill: { color: '180A12' }, line: { color: '3A1020' },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.42, w: 0.06, h: 0.82,
    fill: { color: C.red }, line: { color: C.red },
  });
  s.addText("This isn't a concept. I shipped this at 350M+ MAU scale. The MAF version is the same architecture applied to a smaller, more tractable problem.", {
    x: 0.7, y: 4.47, w: 8.6, h: 0.72,
    fontSize: 10, fontFace: 'Calibri', color: 'C0A8B0', bold: true, margin: 0, valign: 'middle',
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.dark };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.12, h: 5.625,
    fill: { color: C.gold }, line: { color: C.gold },
  });
  s.addText('ROLLOUT PLAN', {
    x: 0.35, y: 0.28, w: 5, h: 0.26,
    fontSize: 8, fontFace: 'Calibri', color: C.gold, bold: true, charSpacing: 3, margin: 0,
  });
  s.addText('Start with One Category. Prove the Detection Speed. Scale to All Ops.', {
    x: 0.35, y: 0.56, w: 9.3, h: 0.58,
    fontSize: 20, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });

  const phases = [
    { phase: 'PHASE 1', period: 'Month 1–2', title: 'Pilot: Chilled Foods Category', color: C.red,
      body: 'Deploy OpsIQ for Chilled Foods only. Wire to inventory feeds + seller SLA logs + NPS stream. Measure: detection lag (target <30 min), false positive rate, ops team adoption.' },
    { phase: 'PHASE 2', period: 'Month 3–4', title: 'Anomaly Detection + AI RCA', color: C.gold,
      body: 'Activate ML anomaly detection across all categories. Deploy AI root-cause analysis with NLP issue tagging. Integrate with Jira for auto-ticket creation. Target: −60% escalation tickets.' },
    { phase: 'PHASE 3', period: 'Month 5–6', title: 'Full Marketplace Ops Intelligence', color: C.green,
      body: 'All KPIs unified in single health dashboard. Seller escalation console live for all category managers. Predictive alerting: flag at-risk sellers 48h before breach.' },
  ];
  const phW = 2.88, phGap = 0.18, phX = 0.35, phY = 1.3, phH = 2.68;

  phases.forEach((ph, i) => {
    const px = phX + i * (phW + phGap);
    s.addShape(pres.shapes.RECTANGLE, {
      x: px, y: phY, w: phW, h: phH,
      fill: { color: '180A12' }, line: { color: '3A1828' }, shadow: makeCardShadow(),
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: px, y: phY, w: phW, h: 0.06,
      fill: { color: ph.color }, line: { color: ph.color },
    });
    s.addText(ph.phase, {
      x: px + 0.12, y: phY + 0.16, w: phW - 0.2, h: 0.26,
      fontSize: 7.5, fontFace: 'Calibri', color: ph.color, bold: true, charSpacing: 2, margin: 0,
    });
    s.addText(ph.period, {
      x: px + 0.12, y: phY + 0.44, w: phW - 0.2, h: 0.26,
      fontSize: 9, fontFace: 'Calibri', color: C.muted, margin: 0,
    });
    s.addText(ph.title, {
      x: px + 0.12, y: phY + 0.72, w: phW - 0.2, h: 0.44,
      fontSize: 11, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
    });
    s.addShape(pres.shapes.LINE, {
      x: px + 0.12, y: phY + 1.22, w: phW - 0.24, h: 0,
      line: { color: '3A1828', width: 1 },
    });
    s.addText(ph.body, {
      x: px + 0.12, y: phY + 1.32, w: phW - 0.2, h: phH - 1.48,
      fontSize: 8.5, fontFace: 'Calibri', color: 'C0A8B0', margin: 0, valign: 'top',
    });
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 4.16, w: 9.3, h: 0.68,
    fill: { color: '120810' }, line: { color: '3A1020' },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 4.16, w: 0.06, h: 0.68,
    fill: { color: C.gold }, line: { color: C.gold },
  });
  s.addText('What I need: Real-time access to order, inventory, and SLA data feeds  ·  Jira API integration  ·  1 data engineer + 1 backend engineer  ·  Ops team as co-owners from sprint 1', {
    x: 0.55, y: 4.2, w: 8.9, h: 0.6,
    fontSize: 9, fontFace: 'Calibri', color: C.muted, margin: 0, valign: 'middle',
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.22, w: 10, h: 0.4,
    fill: { color: '120810' }, line: { color: '120810' },
  });
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439', {
    x: 0.35, y: 5.25, w: 9.3, h: 0.32,
    fontSize: 9, fontFace: 'Calibri', color: C.muted, margin: 0, valign: 'middle',
  });
}

pres.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/maf-opsiq-deck.pptx' });
console.log('✅ maf-opsiq-deck.pptx written');
