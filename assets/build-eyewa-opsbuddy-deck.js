const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9';

const C = {
  dark:   '060C14',
  hero:   '0D1A2B',
  bright: '3B82F6',
  cyan:   '06B6D4',
  white:  'FFFFFF',
  lgray:  'F0F4FF',
  muted:  '94A3B8',
  amber:  'F59E0B',
  red:    'EF4444',
  green:  '22C55E',
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 3, offset: 1, angle: 45, opacity: 0.12 });

// ─── SLIDE 1 — COVER ─────────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  // Diagonal line motif
  for (let i = 0; i < 14; i++) {
    s.addShape(pptx.ShapeType.line, {
      x: -1 + i * 0.85, y: 0, w: 0, h: 7.5,
      line: { color: C.bright, width: 0.4, transparency: 88 },
      rotate: 25,
    });
  }

  // Blue accent block top-left
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 0.18, h: 7.5, fill: { color: C.bright },
  });

  // Eyebrow
  s.addText('EYEWA — INTERNAL OPERATIONS PRODUCT', {
    x: 0.38, y: 0.32, w: 7, h: 0.28,
    fontSize: 9, color: C.muted, bold: false, charSpacing: 1.8,
  });

  // Main title
  s.addText('OpsBuddy', {
    x: 0.38, y: 0.72, w: 7, h: 1.4,
    fontSize: 72, color: C.white, bold: true, fontFace: 'Arial Black',
  });

  // Blue title underline accent — replaced with colored subtitle
  s.addText('Your AI Co-Pilot for Operational Intelligence', {
    x: 0.38, y: 2.05, w: 7.4, h: 0.52,
    fontSize: 22, color: C.bright, bold: false, italic: true,
  });

  // Subtitle
  s.addText('Answering every ops question before the manager asks it', {
    x: 0.38, y: 2.65, w: 7.4, h: 0.38,
    fontSize: 15, color: C.muted,
  });

  // Presenter line
  s.addText('Presented by Ajay Avaghade  ·  Growth & Monetization PM', {
    x: 0.38, y: 6.8, w: 6, h: 0.32,
    fontSize: 10, color: C.muted,
  });

  // Big stat box — bottom right
  s.addShape(pptx.ShapeType.rect, {
    x: 8.0, y: 5.6, w: 1.8, h: 1.6,
    fill: { color: C.bright, transparency: 88 },
    line: { color: C.bright, width: 1 },
  });
  s.addText('4+\nhrs/day', {
    x: 8.0, y: 5.7, w: 1.8, h: 0.9,
    fontSize: 28, color: C.bright, bold: true, align: 'center',
  });
  s.addText('ops managers spend\nsearching dashboards', {
    x: 8.0, y: 6.48, w: 1.8, h: 0.52,
    fontSize: 8, color: C.muted, align: 'center',
  });
}

// ─── SLIDE 2 — THE PROBLEM ───────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  s.addText('THE PROBLEM', {
    x: 0.5, y: 0.22, w: 9, h: 0.26,
    fontSize: 9, color: C.bright, bold: true, charSpacing: 2,
  });
  s.addText('Ops Teams Waste 4+ Hours Daily Switching Between Disconnected Dashboards', {
    x: 0.5, y: 0.52, w: 9, h: 0.82,
    fontSize: 26, color: C.white, bold: true,
  });

  // 3 stat columns
  const cols = [
    { icon: '🔀', stat: '7+', label: 'Tools ops teams juggle daily', sub: 'Shopify, ERP, lab portals, 3PL, spreadsheets — none connected' },
    { icon: '⏳', stat: '4.2 hrs', label: 'Avg ops manager time lost per day', sub: 'Spent searching, switching tabs, and manually compiling reports' },
    { icon: '🚨', stat: '68%', label: 'Anomalies detected after impact', sub: 'Supply issues, lab delays, stockouts — found too late to prevent damage' },
  ];
  cols.forEach((col, i) => {
    const x = 0.4 + i * 3.1;
    s.addShape(pptx.ShapeType.rect, {
      x, y: 1.55, w: 2.85, h: 3.4,
      fill: { color: C.hero }, line: { color: C.bright, width: 1, transparency: 60 },
      shadow: makeShadow(),
    });
    s.addText(col.icon, { x, y: 1.72, w: 2.85, h: 0.5, fontSize: 22, align: 'center' });
    s.addText(col.stat, { x, y: 2.22, w: 2.85, h: 0.72, fontSize: 36, color: C.bright, bold: true, align: 'center' });
    s.addText(col.label, { x: x + 0.12, y: 2.96, w: 2.6, h: 0.52, fontSize: 12, color: C.white, bold: true, align: 'center' });
    s.addText(col.sub, { x: x + 0.1, y: 3.5, w: 2.65, h: 0.88, fontSize: 9.5, color: C.muted, align: 'center' });
  });

  // Insight box
  s.addShape(pptx.ShapeType.rect, {
    x: 0.4, y: 5.12, w: 9.2, h: 0.98,
    fill: { color: C.bright, transparency: 90 }, line: { color: C.bright, width: 1 },
  });
  s.addText('Root cause: eyewa operates across Rx orders, lab partners, supply chain, and 3PL — but each domain lives in a silo. There is no unified intelligence layer that connects the dots and surfaces what ops teams need to know right now.', {
    x: 0.6, y: 5.18, w: 8.8, h: 0.86,
    fontSize: 10.5, color: C.white, italic: true,
  });
}

// ─── SLIDE 3 — THE INSIGHT ───────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE INSIGHT', {
    x: 0.5, y: 0.22, w: 9, h: 0.26,
    fontSize: 9, color: C.bright, bold: true, charSpacing: 2,
  });
  s.addText('Ops Teams Don\'t Need More Dashboards. They Need One Answer.', {
    x: 0.5, y: 0.52, w: 9, h: 0.72,
    fontSize: 26, color: C.dark, bold: true,
  });

  // Left column — Current
  s.addShape(pptx.ShapeType.rect, {
    x: 0.4, y: 1.42, w: 4.1, h: 4.2,
    fill: { color: 'FEF2F2' }, line: { color: 'FCA5A5', width: 1 },
  });
  s.addText('❌  Fragmented Dashboards (Today)', {
    x: 0.6, y: 1.58, w: 3.7, h: 0.38, fontSize: 12, color: '991B1B', bold: true,
  });
  const leftPoints = [
    'Tab 1: Shopify orders — can\'t see lab status',
    'Tab 2: Lab portal — no stock visibility',
    'Tab 3: ERP — no correlation to Rx delays',
    'Tab 4: 3PL — delivery status only',
    'Tab 5: Spreadsheet — manually maintained',
    'Question: "Why did refunds spike?" → takes 40 mins to find answer',
  ];
  leftPoints.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 0.65, y: 2.04 + i * 0.5, w: 3.65, h: 0.44, fontSize: 10.5, color: '374151' });
  });

  // VS divider
  s.addShape(pptx.ShapeType.ellipse, {
    x: 4.62, y: 2.9, w: 0.72, h: 0.72,
    fill: { color: C.bright }, line: { color: C.bright, width: 0 },
  });
  s.addText('VS', { x: 4.62, y: 2.96, w: 0.72, h: 0.42, fontSize: 10, color: C.white, bold: true, align: 'center' });

  // Right column — Proposed
  s.addShape(pptx.ShapeType.rect, {
    x: 5.46, y: 1.42, w: 4.1, h: 4.2,
    fill: { color: 'EFF6FF' }, line: { color: '93C5FD', width: 1 },
  });
  s.addText('✅  OpsBuddy (Proposed)', {
    x: 5.66, y: 1.58, w: 3.7, h: 0.38, fontSize: 12, color: '1D4ED8', bold: true,
  });
  const rightPoints = [
    'One chat interface — asks in plain English',
    '"Why did refunds spike?" → answered in 8 seconds',
    'Pulls from all 5 sources simultaneously',
    'Daily AI brief delivered every morning at 8am',
    'Anomaly alerts surface before managers ask',
    'Natural language → SQL → answer. No switching.',
  ];
  rightPoints.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 5.66, y: 2.04 + i * 0.5, w: 3.65, h: 0.44, fontSize: 10.5, color: '374151' });
  });

  // Callout box
  s.addShape(pptx.ShapeType.rect, {
    x: 0.4, y: 5.72, w: 9.2, h: 0.6,
    fill: { color: C.bright, transparency: 88 }, line: { color: C.bright, width: 1 },
  });
  s.addText('Key insight: The same operations data eyewa already has — in Shopify, lab portals, ERP, 3PL — can become a single conversational intelligence layer. No new data required. Just connection and reasoning.', {
    x: 0.6, y: 5.78, w: 8.8, h: 0.52,
    fontSize: 9.5, color: C.dark, italic: true,
  });
}

// ─── SLIDE 4 — THE MECHANIC ──────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  s.addText('THE MECHANIC', {
    x: 0.5, y: 0.22, w: 9, h: 0.26,
    fontSize: 9, color: C.bright, bold: true, charSpacing: 2,
  });
  s.addText('How OpsBuddy Turns Raw Data Into Instant Answers', {
    x: 0.5, y: 0.52, w: 9, h: 0.62,
    fontSize: 26, color: C.white, bold: true,
  });

  // 5 step cards
  const steps = [
    { n: '01', title: 'Data Ingestion', body: 'Shopify, lab portals, ERP, 3PL feeds sync every 15 minutes. Unified ops data lake — no manual entry.' },
    { n: '02', title: 'NL Query Layer', body: 'Ops manager asks in plain English. LLM translates to SQL + context-aware retrieval across all data sources.' },
    { n: '03', title: 'AI Daily Brief', body: 'Every morning at 8am, OpsBuddy sends a proactive summary: yesterday\'s performance, today\'s priorities, anomaly flags.' },
    { n: '04', title: 'Anomaly Detection', body: 'ML monitors 14 key signals in real-time. Refund spikes, lab SLA drifts, stockout risk — alerted before damage compounds.' },
    { n: '05', title: 'Action Routing', body: 'Answers include next steps. "Route this order to Lab B" or "Draft PO for Oakley RB3025" — one click execution.' },
  ];

  steps.forEach((step, i) => {
    const x = 0.38 + i * 1.92;
    s.addShape(pptx.ShapeType.rect, {
      x, y: 1.38, w: 1.72, h: 4.0,
      fill: { color: C.hero }, line: { color: C.bright, width: 1, transparency: 70 },
      shadow: makeShadow(),
    });
    // Step number badge
    s.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.52, y: 1.52, w: 0.68, h: 0.68,
      fill: { color: C.bright }, line: { color: C.bright, width: 0 },
    });
    s.addText(step.n, { x: x + 0.52, y: 1.56, w: 0.68, h: 0.46, fontSize: 13, color: C.white, bold: true, align: 'center' });
    s.addText(step.title, { x: x + 0.08, y: 2.3, w: 1.55, h: 0.52, fontSize: 11, color: C.white, bold: true, align: 'center' });
    s.addText(step.body, { x: x + 0.08, y: 2.88, w: 1.55, h: 2.2, fontSize: 9, color: C.muted, align: 'center' });

    // Connecting arrow (except last)
    if (i < 4) {
      s.addShape(pptx.ShapeType.line, {
        x: x + 1.72, y: 3.38, w: 0.2, h: 0,
        line: { color: C.bright, width: 1.2, dashType: 'dash', transparency: 40 },
      });
    }
  });

  // PhonePe proof note
  s.addShape(pptx.ShapeType.rect, {
    x: 0.38, y: 5.55, w: 9.2, h: 0.72,
    fill: { color: C.bright, transparency: 92 }, line: { color: C.bright, width: 1, transparency: 40 },
  });
  s.addText('PhonePe proof: Built a unified ops intelligence layer for the Offers platform — 350M+ MAU, ₹1000+ Cr/yr budget. Replaced 6 separate analyst queries with one ML-driven dashboard. Ops TAT for offer decisioning dropped from 2 days → 4 hours.', {
    x: 0.58, y: 5.62, w: 8.8, h: 0.6,
    fontSize: 9.5, color: C.white, italic: true,
  });
}

// ─── SLIDE 5 — THE PRODUCT ───────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE PRODUCT', {
    x: 0.5, y: 0.22, w: 9, h: 0.26,
    fontSize: 9, color: C.bright, bold: true, charSpacing: 2,
  });
  s.addText('4 Core Screens — Built & Validated', {
    x: 0.5, y: 0.52, w: 9, h: 0.62,
    fontSize: 26, color: C.dark, bold: true,
  });

  const screens = [
    {
      n: 'S1', title: 'Chat Interface',
      desc: 'Natural language Q&A with embedded data responses. Ask "Why are KSA orders delayed?" and get a structured answer with root cause, affected SKUs, and next step.',
      mock: '┌──────────────────┐\n│ 💬 Chat  │ Context│\n│──────────────────│\n│ Q: Why are KSA   │\n│ orders delayed?  │\n│                  │\n│ A: 68% tied to   │\n│ Vision Labs SLA  │\n│ breach. 14 orders│\n│ > 5 days.        │\n│ ▸ Route to Lab B │\n└──────────────────┘',
    },
    {
      n: 'S2', title: 'Daily Brief',
      desc: 'Morning AI summary at 8am: yesterday performance, today priorities, overnight anomaly flags. Replaces 40 mins of manual dashboard checking.',
      mock: '┌──────────────────┐\n│ 📋 Daily Brief   │\n│ Mon 25 May · 8am │\n│──────────────────│\n│ Yesterday        │\n│ Orders: 847 ↑12% │\n│ Rx errors: 18 ✓  │\n│                  │\n│ Today Priority   │\n│ ⚠ Lab B at cap  │\n│ ⚠ Oakley spike  │\n└──────────────────┘',
    },
    {
      n: 'S3', title: 'Anomaly Feed',
      desc: 'Real-time alert stream with severity scoring. HIGH / MED / LOW tags, root cause snippets, and one-click action routing so ops teams act, not just read.',
      mock: '┌──────────────────┐\n│ 🚨 Anomalies  2H │\n│──────────────────│\n│ [HIGH] Refund    │\n│ spike +340% KSA  │\n│ Oakley RB2140    │\n│ → Review returns │\n│                  │\n│ [MED] Lab B SLA  │\n│ trending 4.8d    │\n│ → Reroute queue  │\n└──────────────────┘',
    },
    {
      n: 'S4', title: 'Admin Config',
      desc: 'Data connector registry, alert threshold settings, data source toggle. Ops lead controls what OpsBuddy monitors — no engineering dependency.',
      mock: '┌──────────────────┐\n│ ⚙ Admin Config  │\n│──────────────────│\n│ Shopify    🟢 ON │\n│ Lab Portal 🟢 ON │\n│ ERP        🟢 ON │\n│ 3PL        🟡 LAG│\n│──────────────────│\n│ Alert Rules  6/6 │\n│ Refresh: 15min   │\n└──────────────────┘',
    },
  ];

  screens.forEach((sc, i) => {
    const x = 0.32 + i * 2.42;
    s.addShape(pptx.ShapeType.rect, {
      x, y: 1.32, w: 2.22, h: 5.0,
      fill: { color: C.white }, line: { color: '93C5FD', width: 1 },
      shadow: makeShadow(),
    });
    s.addText(`${sc.n} — ${sc.title}`, {
      x: x + 0.1, y: 1.44, w: 2.02, h: 0.36,
      fontSize: 10.5, color: C.bright, bold: true,
    });
    // Mock block
    s.addShape(pptx.ShapeType.rect, {
      x: x + 0.1, y: 1.84, w: 2.02, h: 2.42,
      fill: { color: C.dark }, line: { color: C.bright, width: 0.5, transparency: 60 },
    });
    s.addText(sc.mock, {
      x: x + 0.14, y: 1.88, w: 1.94, h: 2.34,
      fontSize: 7, color: 'A5F3FC', fontFace: 'Courier New', valign: 'top',
    });
    s.addText(sc.desc, {
      x: x + 0.1, y: 4.34, w: 2.02, h: 1.82,
      fontSize: 9, color: '374151', valign: 'top',
    });
  });
}

// ─── SLIDE 6 — IMPACT & ROI ──────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  s.addText('IMPACT & ROI', {
    x: 0.5, y: 0.22, w: 9, h: 0.26,
    fontSize: 9, color: C.bright, bold: true, charSpacing: 2,
  });
  s.addText('Projected Impact — Built on PhonePe Proof Points', {
    x: 0.5, y: 0.52, w: 9, h: 0.62,
    fontSize: 26, color: C.white, bold: true,
  });

  // Left section — Ops Efficiency
  s.addShape(pptx.ShapeType.rect, {
    x: 0.38, y: 1.3, w: 4.5, h: 0.38,
    fill: { color: C.bright, transparency: 80 }, line: { color: C.bright, width: 0 },
  });
  s.addText('OPS TEAM EFFICIENCY', {
    x: 0.48, y: 1.35, w: 4.2, h: 0.26,
    fontSize: 9, color: C.white, bold: true, charSpacing: 1.5,
  });

  const leftMetrics = [
    { val: '↓ 75%', label: 'Time spent on manual dashboard switching', sub: '4.2 hrs → 1.1 hrs per ops manager per day' },
    { val: '↑ 8 sec', label: 'Avg query answer time (vs 40 min manual)', sub: 'Natural language → structured answer' },
    { val: '↑ 3×', label: 'Anomaly detection speed', sub: 'Issues flagged in real-time, not at EOD review' },
    { val: '↑ 90%', label: 'Ops team briefing coverage', sub: 'Daily brief replaces 3 separate morning standup reports' },
  ];
  leftMetrics.forEach((m, i) => {
    const y = 1.82 + i * 0.96;
    s.addShape(pptx.ShapeType.rect, {
      x: 0.38, y, w: 4.5, h: 0.84,
      fill: { color: C.hero }, line: { color: C.bright, width: 0.5, transparency: 70 },
    });
    s.addText(m.val, { x: 0.5, y: y + 0.06, w: 1.3, h: 0.44, fontSize: 22, color: C.bright, bold: true });
    s.addText(m.label, { x: 1.88, y: y + 0.06, w: 2.8, h: 0.36, fontSize: 11, color: C.white, bold: true });
    s.addText(m.sub, { x: 1.88, y: y + 0.42, w: 2.8, h: 0.32, fontSize: 9, color: C.muted });
  });

  // Right section — Business ROI
  s.addShape(pptx.ShapeType.rect, {
    x: 5.12, y: 1.3, w: 4.5, h: 0.38,
    fill: { color: C.cyan, transparency: 80 }, line: { color: C.cyan, width: 0 },
  });
  s.addText('BUSINESS & COST ROI', {
    x: 5.22, y: 1.35, w: 4.2, h: 0.26,
    fontSize: 9, color: C.white, bold: true, charSpacing: 1.5,
  });

  const rightMetrics = [
    { val: 'AED 2.4M+', label: 'Annual ops cost recovery', sub: '~15 ops FTEs × 3 hrs saved/day × AED 220/hr' },
    { val: '↓ 55%', label: 'Reactive firefighting incidents', sub: 'Anomalies caught early before they cascade' },
    { val: '↑ 22%', label: 'Cross-border order fill rate', sub: 'KSA/Bahrain expansion: ops now has visibility day-1' },
    { val: '< 4 wks', label: 'Time to deploy new market intelligence', sub: 'New country launch: connect data → OpsBuddy live in days' },
  ];
  rightMetrics.forEach((m, i) => {
    const y = 1.82 + i * 0.96;
    s.addShape(pptx.ShapeType.rect, {
      x: 5.12, y, w: 4.5, h: 0.84,
      fill: { color: C.hero }, line: { color: C.cyan, width: 0.5, transparency: 70 },
    });
    s.addText(m.val, { x: 5.24, y: y + 0.06, w: 1.6, h: 0.44, fontSize: 20, color: C.cyan, bold: true });
    s.addText(m.label, { x: 6.88, y: y + 0.06, w: 2.6, h: 0.36, fontSize: 11, color: C.white, bold: true });
    s.addText(m.sub, { x: 6.88, y: y + 0.42, w: 2.6, h: 0.32, fontSize: 9, color: C.muted });
  });

  // Insight box
  s.addShape(pptx.ShapeType.rect, {
    x: 0.38, y: 5.7, w: 9.24, h: 0.6,
    fill: { color: C.bright, transparency: 90 }, line: { color: C.bright, width: 1 },
  });
  s.addText('OpsBuddy does not replace headcount — it removes the cognitive overhead that prevents smart people from doing their best work. At PhonePe, the equivalent tool reduced analyst escalation requests by 68% within 60 days of launch.', {
    x: 0.58, y: 5.76, w: 8.8, h: 0.52,
    fontSize: 9.5, color: C.white, italic: true,
  });
}

// ─── SLIDE 7 — PROOF OF WORK ─────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.lgray };

  s.addText('PROOF OF WORK', {
    x: 0.5, y: 0.22, w: 9, h: 0.26,
    fontSize: 9, color: C.bright, bold: true, charSpacing: 2,
  });
  s.addText('I Built This. Here\'s the Proof.', {
    x: 0.5, y: 0.52, w: 9, h: 0.62,
    fontSize: 26, color: C.dark, bold: true,
  });

  // Left column — dark, what I built
  s.addShape(pptx.ShapeType.rect, {
    x: 0.38, y: 1.3, w: 4.5, h: 4.72,
    fill: { color: C.dark }, line: { color: C.bright, width: 1, transparency: 50 },
  });
  s.addText('What I Built at PhonePe', {
    x: 0.58, y: 1.48, w: 4.1, h: 0.38,
    fontSize: 13, color: C.bright, bold: true,
  });
  const leftItems = [
    'Owned the Offers Intelligence layer for 350M+ MAU platform',
    'Built ML propensity-to-transact models replacing manual analyst cohorts — ops TAT from 2 days → 4 hours',
    'Designed the ops-facing campaign builder: threshold-based nudges, real-time cohort updates, multi-signal triggers',
    'Replaced 6 separate dashboard queries with one unified signal layer. Analyst escalations ↓68% in 60 days',
    'Ran the A/B framework across 14 concurrent experiments — results surfaced in a single ops digest, not 14 tabs',
    'Shipped merchant anomaly detection: refund spikes, device inactivity, GMV deviation — flagged in <15 min',
  ];
  leftItems.forEach((item, i) => {
    s.addText(`• ${item}`, {
      x: 0.58, y: 1.98 + i * 0.62, w: 4.08, h: 0.56,
      fontSize: 10, color: C.lgray, valign: 'top',
    });
  });

  // Right column — light, mapping to eyewa JD
  s.addShape(pptx.ShapeType.rect, {
    x: 5.12, y: 1.3, w: 4.5, h: 4.72,
    fill: { color: C.white }, line: { color: '93C5FD', width: 1 },
  });
  s.addText('How It Maps to This Role', {
    x: 5.32, y: 1.48, w: 4.1, h: 0.38,
    fontSize: 13, color: '1D4ED8', bold: true,
  });
  const rightItems = [
    'Unified intelligence layer → OpsBuddy chat + daily brief',
    'ML signal layer → Anomaly detection engine (14 ops signals)',
    'Campaign builder → Alert rule configurator (ops-facing)',
    'Single ops digest → Daily Brief replacing 3 manual reports',
    'A/B experiment dashboard → Multi-country ops comparison view',
    'Merchant anomaly detection → eyewa lab + supplier signals',
  ];
  rightItems.forEach((item, i) => {
    s.addText(`✓  ${item}`, {
      x: 5.32, y: 1.98 + i * 0.62, w: 4.08, h: 0.56,
      fontSize: 10, color: '374151', valign: 'top',
    });
  });

  // Closing callout
  s.addShape(pptx.ShapeType.rect, {
    x: 0.38, y: 6.12, w: 9.24, h: 0.72,
    fill: { color: C.bright, transparency: 88 }, line: { color: C.bright, width: 1 },
  });
  s.addText('"The hardest part of building OpsBuddy isn\'t the AI — it\'s knowing which 14 signals matter and why. I\'ve already built that at scale."', {
    x: 0.58, y: 6.2, w: 8.8, h: 0.56,
    fontSize: 11, color: C.dark, italic: true, bold: true,
  });
}

// ─── SLIDE 8 — ROLLOUT PLAN ──────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 0.18, h: 7.5, fill: { color: C.bright },
  });

  s.addText('ROLLOUT PLAN', {
    x: 0.5, y: 0.22, w: 9, h: 0.26,
    fontSize: 9, color: C.bright, bold: true, charSpacing: 2,
  });
  s.addText('Phased Rollout Plan', {
    x: 0.5, y: 0.52, w: 9, h: 0.62,
    fontSize: 26, color: C.white, bold: true,
  });

  const phases = [
    {
      phase: 'Phase 1',
      period: 'Month 1–2: Foundation',
      items: [
        'Connect Shopify + lab portal APIs (read-only)',
        'Build NL query layer with 20 pre-trained ops questions',
        'Deploy daily brief to 3 ops managers (pilot users)',
        'Instrument: query volume, answer accuracy, TAT reduction',
      ],
    },
    {
      phase: 'Phase 2',
      period: 'Month 3–4: Intelligence Layer',
      items: [
        'Add ERP + 3PL connectors',
        'Launch anomaly detection (6 initial signals)',
        'Ship alert configurator for ops lead self-service',
        'Measure: anomaly catch rate vs EOD-review baseline',
      ],
    },
    {
      phase: 'Phase 3',
      period: 'Month 5–6: Scale + Markets',
      items: [
        'Expand anomaly signals to 14 (lab SLA, refund spike, stockout risk)',
        'Multi-market view: UAE + KSA + Bahrain in one dashboard',
        'Roll out to full ops team (target: 100% of daily briefs replaced)',
        'KPI: ops escalation tickets ↓55%, daily manual report time ↓75%',
      ],
    },
  ];

  phases.forEach((ph, i) => {
    const x = 0.38 + i * 3.2;
    s.addShape(pptx.ShapeType.rect, {
      x, y: 1.32, w: 3.0, h: 4.0,
      fill: { color: C.hero }, line: { color: C.bright, width: 1, transparency: 50 },
      shadow: makeShadow(),
    });
    // Phase badge
    s.addShape(pptx.ShapeType.rect, {
      x, y: 1.32, w: 3.0, h: 0.42,
      fill: { color: C.bright, transparency: i * 20 }, line: { color: C.bright, width: 0 },
    });
    s.addText(ph.phase, { x: x + 0.12, y: 1.36, w: 2.76, h: 0.32, fontSize: 11, color: C.white, bold: true });
    s.addText(ph.period, { x: x + 0.12, y: 1.82, w: 2.76, h: 0.36, fontSize: 10.5, color: C.bright, bold: true });
    ph.items.forEach((item, j) => {
      s.addText(`• ${item}`, {
        x: x + 0.12, y: 2.24 + j * 0.62, w: 2.76, h: 0.56,
        fontSize: 9.5, color: C.muted, valign: 'top',
      });
    });
  });

  // What I need box
  s.addShape(pptx.ShapeType.rect, {
    x: 0.38, y: 5.46, w: 9.24, h: 1.08,
    fill: { color: C.bright, transparency: 90 }, line: { color: C.bright, width: 1 },
  });
  s.addText('What I Need to Build This', {
    x: 0.58, y: 5.52, w: 4, h: 0.32,
    fontSize: 10, color: C.bright, bold: true,
  });
  s.addText('Read-only API access to Shopify, lab portals, ERP, 3PL  ·  1 backend engineer for data connectors  ·  1 week of ops manager shadowing  ·  Alignment on data governance / PII scope for ops data', {
    x: 0.58, y: 5.84, w: 8.8, h: 0.48,
    fontSize: 9.5, color: C.white,
  });

  // Contact footer
  s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  linkedin.com/in/ajay-avaghade', {
    x: 0.38, y: 6.92, w: 9.24, h: 0.28,
    fontSize: 9, color: C.muted, align: 'center',
  });
}

// ─── SAVE ────────────────────────────────────────────────────────────────────
pptx.writeFile({ fileName: 'assets/eyewa-opsbuddy-deck.pptx' })
  .then(() => console.log('✅  assets/eyewa-opsbuddy-deck.pptx saved'))
  .catch(err => { console.error(err); process.exit(1); });
