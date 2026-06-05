const pptx = require('pptxgenjs');
const prs = new pptx();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0A0A0A',
  hero:   '141414',
  bright: 'C2A582',
  orange: 'D4B895',
  amber:  'F59E0B',
  white:  'FFFFFF',
  lgray:  'FDFBF7',
  muted:  '8A8A8A',
  ink:    '111111',
  slate:  '404040',
  red:    'EF4444',
  green:  '10B981'
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 3, offset: 1, angle: 45, opacity: 0.12 });

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 1 — COVER (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  // Diagonal stripe accents
  for (let i = 0; i < 6; i++) {
    slide.addShape(prs.ShapeType.rect, {
      x: -0.5 + i * 2.2, y: -0.2, w: 0.18, h: 10,
      fill: { color: C.bright, transparency: 88 },
      line: { type: 'none' },
      rotate: 15,
    });
  }

  // Dezerv gold accent bar top
  slide.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 0.06,
    fill: { color: C.bright }, line: { type: 'none' },
  });

  // Company tag
  slide.addText('DEZERV — INVESTMENT EXECUTION PM', {
    x: 0.6, y: 0.28, w: 6, h: 0.28,
    fontSize: 9, bold: true, color: C.bright,
    charSpacing: 3, fontFace: 'Calibri',
  });

  // Product name
  slide.addText('WealthTrack Pro', {
    x: 0.6, y: 0.72, w: 8, h: 1.1,
    fontSize: 56, bold: true, color: C.white, fontFace: 'Calibri',
  });

  // Subtitle
  slide.addText('Unified high-fidelity client reporting that drives trust and\norganic AUM growth through radical transparency', {
    x: 0.6, y: 1.82, w: 7.4, h: 0.72,
    fontSize: 15, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
  });

  // Presenter
  slide.addText('Presented by Ajay Avaghade  ·  Product Manager — Investment Execution', {
    x: 0.6, y: 2.68, w: 6, h: 0.3,
    fontSize: 11, color: C.muted, fontFace: 'Calibri',
  });

  // Big stat box bottom right
  slide.addShape(prs.ShapeType.rect, {
    x: 7.2, y: 1.6, w: 2.4, h: 1.9,
    fill: { color: C.bright, transparency: 15 },
    line: { color: C.bright, pt: 1.5 },
    rectRadius: 0.1,
  });
  slide.addText('-40%', {
    x: 7.2, y: 1.72, w: 2.4, h: 0.72,
    fontSize: 36, bold: true, color: C.bright, align: 'center', fontFace: 'Calibri',
  });
  slide.addText('reduction in\nclient support\nqueries projected', {
    x: 7.2, y: 2.44, w: 2.4, h: 0.72,
    fontSize: 9.5, color: C.muted, align: 'center', fontFace: 'Calibri', lineSpacingMultiple: 1.3,
  });

  // Bottom bar
  slide.addShape(prs.ShapeType.rect, {
    x: 0, y: 7.08, w: 10, h: 0.42,
    fill: { color: C.hero }, line: { type: 'none' },
  });
  slide.addText('Dezerv — Proprietary & Confidential  ·  Prepared by Ajay Avaghade  ·  2026', {
    x: 0.3, y: 7.1, w: 9.4, h: 0.3,
    fontSize: 8, color: C.muted, fontFace: 'Calibri',
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  slide.addText('THE PROBLEM', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('Fragmented Reporting Creates Friction and Erodes Trust', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.9,
    fontSize: 26, bold: true, color: C.white, fontFace: 'Calibri', lineSpacingMultiple: 1.2,
  });

  // 3 stat columns
  const stats = [
    { icon: '🧩', num: '3+', lbl: 'Asset Classes per Client', sub: 'Affluent clients hold PMS, AIFs, and MFs, each reporting on different schedules and formats.' },
    { icon: '⏳', num: 'T+2', lbl: 'Settlement Anxiety', sub: 'High-value transactions sit in a black box during settlement, driving support calls to Relationship Managers.' },
    { icon: '🧾', num: 'Opaque', lbl: 'Fees & True Tax Yield', sub: 'Gross returns obscure the real net yield after platform fees and capital gains taxes.' },
  ];
  stats.forEach((s, i) => {
    const x = 0.5 + i * 3.1;
    slide.addShape(prs.ShapeType.rect, {
      x, y: 1.72, w: 2.85, h: 3.4,
      fill: { color: C.hero }, line: { color: C.bright, pt: 1, transparency: 60 }, rectRadius: 0.1,
      shadow: makeShadow(),
    });
    slide.addText(s.icon, { x, y: 1.88, w: 2.85, h: 0.44, fontSize: 22, align: 'center' });
    slide.addText(s.num, {
      x, y: 2.36, w: 2.85, h: 0.72,
      fontSize: 38, bold: true, color: C.bright, align: 'center', fontFace: 'Calibri',
    });
    slide.addText(s.lbl, {
      x: x + 0.14, y: 3.12, w: 2.6, h: 0.52,
      fontSize: 11, bold: true, color: C.white, align: 'center', fontFace: 'Calibri', lineSpacingMultiple: 1.2,
    });
    slide.addText(s.sub, {
      x: x + 0.1, y: 3.68, w: 2.68, h: 1.1,
      fontSize: 9, color: C.muted, align: 'center', fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });

  // Insight box
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 5.38, w: 9, h: 1.18,
    fill: { color: C.hero }, line: { color: C.orange, pt: 1.5, transparency: 40 }, rectRadius: 0.1,
  });
  slide.addText('Root cause: Traditional wealth management relies on static, monthly PDF statements that hide execution complexity. For affluent Indians, trust is built through real-time transparency. When they can\'t see exactly where their ₹50L AIF capital call is in the settlement pipeline, or what their true post-tax yield is, they hesitate to allocate more capital.', {
    x: 0.72, y: 5.5, w: 8.6, h: 0.94,
    fontSize: 10.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.4,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 3 — THE INSIGHT (light)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.lgray };

  slide.addText('THE INSIGHT', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.slate, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('Radical Transparency as an AUM Growth Engine', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.82,
    fontSize: 24, bold: true, color: C.ink, fontFace: 'Calibri', lineSpacingMultiple: 1.2,
  });

  // VS divider oval
  slide.addShape(prs.ShapeType.ellipse, {
    x: 4.6, y: 2.3, w: 0.8, h: 0.8,
    fill: { color: C.dark }, line: { type: 'none' },
  });
  slide.addText('VS', {
    x: 4.6, y: 2.36, w: 0.8, h: 0.62,
    fontSize: 11, bold: true, color: C.white, align: 'center', fontFace: 'Calibri',
  });

  // LEFT column — current
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.6, w: 4.0, h: 4.2,
    fill: { color: 'FFF5F5' }, line: { color: 'EF4444', pt: 1.5, transparency: 40 }, rectRadius: 0.1,
  });
  slide.addText('❌  Static Reporting (Today)', {
    x: 0.68, y: 1.76, w: 3.68, h: 0.4,
    fontSize: 13, bold: true, color: C.ink, fontFace: 'Calibri',
  });
  const leftPts = [
    'Offline spreadsheets managed by Relationship Managers to answer basic client questions.',
    'Siloed views: PMS and AIFs tracked separately, requiring manual client mental math.',
    'Gross yield reporting that surprises clients at tax season.',
    'Black-box execution: "Your funds have been debited, wait 2 days for unit allocation."',
  ];
  leftPts.forEach((pt, i) => {
    slide.addText('• ' + pt, {
      x: 0.68, y: 2.22 + i * 0.72, w: 3.6, h: 0.64,
      fontSize: 10.5, color: C.slate, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
    });
  });

  // RIGHT column — proposed
  slide.addShape(prs.ShapeType.rect, {
    x: 5.5, y: 1.6, w: 4.0, h: 4.2,
    fill: { color: 'FAF8F5' }, line: { color: C.bright, pt: 1.5, transparency: 30 }, rectRadius: 0.1,
  });
  slide.addText('✅  WealthTrack Pro (Proposed)', {
    x: 5.68, y: 1.76, w: 3.68, h: 0.4,
    fontSize: 13, bold: true, color: C.ink, fontFace: 'Calibri',
  });
  const rightPts = [
    'Unified API layer aggregating PMS, AIF, and MF ledgers into a single real-time view.',
    'Live settlement steppers (like food delivery tracking) for high-value executions.',
    'True Net Yield calculations automatically deducting Dezerv fees and projected LTCG tax.',
    'Proactive AI growth nudges showing the future compounding impact of an increased SIP.',
  ];
  rightPts.forEach((pt, i) => {
    slide.addText('• ' + pt, {
      x: 5.68, y: 2.22 + i * 0.72, w: 3.6, h: 0.64,
      fontSize: 10.5, color: C.slate, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
    });
  });

  // Callout box
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.0, w: 9, h: 0.96,
    fill: { color: C.dark }, line: { type: 'none' }, rectRadius: 0.1,
  });
  slide.addText('Key insight: Demystifying execution complexity is a feature. By showing clients exactly what is happening under the hood (fees, taxes, settlement states), Dezerv positions itself as a fully aligned partner, turning the reporting dashboard into a proactive growth engine.', {
    x: 0.72, y: 6.1, w: 8.6, h: 0.76,
    fontSize: 10, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.4,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 4 — THE MECHANIC (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  slide.addText('THE MECHANIC', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('How WealthTrack Works: Aggregation, Calculation, Projection', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.8,
    fontSize: 23, bold: true, color: C.white, fontFace: 'Calibri', lineSpacingMultiple: 1.2,
  });

  const steps = [
    { n: '01', title: 'Data Aggregation Layer', desc: 'Internal ledgers, AMC APIs, and AIF reporting feeds are normalized into a unified client portfolio schema, ensuring single-source-of-truth accuracy.' },
    { n: '02', title: 'Live Execution Webhooks', desc: 'Instead of batch updates, PG and AMC settlement webhooks push real-time status changes to the client app, updating the "Settlement Stepper" instantly.' },
    { n: '03', title: 'Net Yield Engine', desc: 'System calculates Gross XIRR, then automatically deducts Dezerv platform fees and estimated capital gains (LTCG/STCG) based on the client\'s holding period.' },
    { n: '04', title: 'AI Projection & Nudging', desc: 'Predictive models forecast 5-year wealth trajectories based on historical performance and current SIPs, triggering highly contextual "Increase SIP" nudges.' },
  ];

  steps.forEach((s, i) => {
    const y = 1.6 + i * 1.2;
    // Step number circle
    slide.addShape(prs.ShapeType.ellipse, {
      x: 0.5, y: y, w: 0.48, h: 0.48,
      fill: { color: C.bright }, line: { type: 'none' },
    });
    slide.addText(s.n, {
      x: 0.5, y: y + 0.02, w: 0.48, h: 0.38,
      fontSize: 9, bold: true, color: C.dark, align: 'center', fontFace: 'Calibri',
    });
    // Connector line (skip last)
    if (i < 3) {
      slide.addShape(prs.ShapeType.rect, {
        x: 0.72, y: y + 0.48, w: 0.04, h: 0.72,
        fill: { color: C.bright, transparency: 60 }, line: { type: 'none' },
      });
    }
    slide.addText(s.title, {
      x: 1.18, y: y, w: 8.2, h: 0.28,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri',
    });
    slide.addText(s.desc, {
      x: 1.18, y: y + 0.28, w: 8.2, h: 0.62,
      fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 5 — THE PRODUCT (light)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.lgray };

  slide.addText('THE PRODUCT', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.slate, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('4 Key Screen States — High-Fidelity Client Transparency', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 22, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  const screens = [
    {
      num: '01', title: 'Unified Portfolio',
      desc: 'Aggregates PMS, AIFs, and MFs into a single view. Shows total AUM and individual asset class XIRR performance transparently.',
      mock: '[ Total Value: ₹2.45 Cr ]\n ↑ 18.4% XIRR\n\n [ Dezerv PMS  | 48% ]\n [ AIFs        | 35% ]\n [ MFs         | 17% ]',
    },
    {
      num: '02', title: 'Settlement Tracking',
      desc: 'Real-time execution steppers. Eliminates anxiety for high-value transactions by showing exactly where capital is in the T+2 flow.',
      mock: '[ Dezerv PMS Addition ]\n ₹1,00,000\n\n ✓ Payment Deducted\n ✓ Sent to AMC\n ⏳ Units Allotted (T+1)',
    },
    {
      num: '03', title: 'Tax & Fee Breakdown',
      desc: 'Builds ultimate trust by proactively showing Gross Yield vs True Net Yield after Dezerv fees and projected capital gains taxes.',
      mock: '[ Gross XIRR: 18.4% ]\n [ Net XIRR:   16.1% ]\n\n Platform Fee: -₹42,000\n Realized Tax: -₹1.15L',
    },
    {
      num: '04', title: 'Wealth Predictor',
      desc: 'AI-driven 5-year projections. Includes contextual nudges proving how small SIP increases compound over time, driving organic AUM.',
      mock: '[ Projected: ₹5.82 Cr ]\n\n 💡 Actionable Insight:\n Increasing SIP by ₹50k\n hits ₹6Cr 14 mos faster.',
    },
  ];

  screens.forEach((s, i) => {
    const x = 0.32 + i * 2.36;
    slide.addShape(prs.ShapeType.rect, {
      x, y: 1.4, w: 2.2, h: 5.3,
      fill: { color: C.white }, line: { color: C.bright, pt: 1.5, transparency: 40 }, rectRadius: 0.1,
      shadow: makeShadow(),
    });
    slide.addShape(prs.ShapeType.rect, {
      x, y: 1.4, w: 2.2, h: 0.36,
      fill: { color: C.dark }, line: { type: 'none' }, rectRadius: 0.1,
    });
    slide.addText(s.num + '  ' + s.title, {
      x: x + 0.1, y: 1.44, w: 2.05, h: 0.28,
      fontSize: 8.5, bold: true, color: C.bright, fontFace: 'Calibri',
    });
    // Mockup block
    slide.addShape(prs.ShapeType.rect, {
      x: x + 0.1, y: 1.86, w: 2.02, h: 2.3,
      fill: { color: 'FAF8F5' }, line: { color: C.bright, pt: 0.75, transparency: 60 }, rectRadius: 0.08,
    });
    slide.addText(s.mock, {
      x: x + 0.14, y: 1.9, w: 1.96, h: 2.2,
      fontSize: 7.5, color: C.slate, fontFace: 'Courier New', lineSpacingMultiple: 1.45,
    });
    slide.addText(s.desc, {
      x: x + 0.1, y: 4.24, w: 2.02, h: 2.2,
      fontSize: 9, color: C.slate, fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });

  slide.addText('Live interactive prototype: dezerv-wealthtrack-prototype.html', {
    x: 0.5, y: 6.88, w: 9, h: 0.28,
    fontSize: 8.5, color: C.slate, italic: true, fontFace: 'Calibri', align: 'center',
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  slide.addText('IMPACT & ROI', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('Projected Impact — Why Transparency Matters', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 24, bold: true, color: C.white, fontFace: 'Calibri',
  });

  // Left section label
  slide.addText('OPERATIONAL EFFICIENCY', {
    x: 0.5, y: 1.38, w: 4.4, h: 0.28,
    fontSize: 8.5, bold: true, color: C.bright, charSpacing: 2, fontFace: 'Calibri',
  });
  const leftMetrics = [
    { val: '−40%', lbl: 'RM Support Queries', sub: 'Clients self-serve settlement status and tax questions' },
    { val: 'T+0', lbl: 'Visibility Delay', sub: 'Execution steps are pushed instantly to the app' },
    { val: '100%', lbl: 'Fee Transparency', sub: 'Eliminates friction and surprises during tax season' },
  ];
  leftMetrics.forEach((m, i) => {
    const y = 1.72 + i * 1.3;
    slide.addShape(prs.ShapeType.rect, {
      x: 0.5, y, w: 4.3, h: 1.15,
      fill: { color: C.hero }, line: { color: C.bright, pt: 0.75, transparency: 50 }, rectRadius: 0.08,
    });
    slide.addText(m.val, {
      x: 0.68, y: y + 0.15, w: 1.6, h: 0.52,
      fontSize: 28, bold: true, color: C.bright, fontFace: 'Calibri',
    });
    slide.addText(m.lbl, {
      x: 2.32, y: y + 0.15, w: 2.3, h: 0.28,
      fontSize: 11, bold: true, color: C.white, fontFace: 'Calibri',
    });
    slide.addText(m.sub, {
      x: 2.32, y: y + 0.47, w: 2.3, h: 0.54,
      fontSize: 8.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
    });
  });

  // Right section label
  slide.addText('BUSINESS & AUM GROWTH', {
    x: 5.1, y: 1.38, w: 4.4, h: 0.28,
    fontSize: 8.5, bold: true, color: C.orange, charSpacing: 2, fontFace: 'Calibri',
  });
  const rightMetrics = [
    { val: '+22%', lbl: 'SIP Upsell Conversion', sub: 'Driven organically by AI wealth projection nudges' },
    { val: '+15%', lbl: 'Share of Wallet', sub: 'Clients consolidate external portfolios into Dezerv due to UX' },
    { val: 'Zero', lbl: 'Churn Due to Confusion', sub: 'Clients never leave because they "don\'t understand the returns"' },
  ];
  rightMetrics.forEach((m, i) => {
    const y = 1.72 + i * 1.3;
    slide.addShape(prs.ShapeType.rect, {
      x: 5.1, y, w: 4.3, h: 1.15,
      fill: { color: C.hero }, line: { color: C.orange, pt: 0.75, transparency: 50 }, rectRadius: 0.08,
    });
    slide.addText(m.val, {
      x: 5.28, y: y + 0.15, w: 1.6, h: 0.52,
      fontSize: 28, bold: true, color: C.orange, fontFace: 'Calibri',
    });
    slide.addText(m.lbl, {
      x: 6.92, y: y + 0.15, w: 2.3, h: 0.28,
      fontSize: 11, bold: true, color: C.white, fontFace: 'Calibri',
    });
    slide.addText(m.sub, {
      x: 6.92, y: y + 0.47, w: 2.3, h: 0.54,
      fontSize: 8.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
    });
  });

  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.56, w: 9, h: 0.6,
    fill: { color: C.hero }, line: { color: C.bright, pt: 1, transparency: 50 }, rectRadius: 0.08,
  });
  slide.addText('WealthTrack Pro transforms the app from a passive dashboard into a primary engagement tool. By treating reporting reliability as a core product feature, Dezerv directly drives organic AUM consolidation.', {
    x: 0.72, y: 6.64, w: 8.6, h: 0.44,
    fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK (light)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.lgray };

  slide.addText('PROOF OF WORK', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.slate, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('I Built This Architecture. At PhonePe. At Scale.', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 23, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  // Left — dark column
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.38, w: 4.3, h: 5.3,
    fill: { color: C.dark }, line: { type: 'none' }, rectRadius: 0.1,
  });
  slide.addText('What I Built at PhonePe', {
    x: 0.7, y: 1.56, w: 3.9, h: 0.34,
    fontSize: 12, bold: true, color: C.bright, fontFace: 'Calibri',
  });
  const leftPts = [
    'Led execution architecture for third-party rewards, rebuilding ₹100 Cr/yr infrastructure into an ML-driven marketplace handling massive transaction volumes seamlessly.',
    'Unified disparate merchant data streams into a single reliable reporting dashboard used by 5,000+ B2B partners.',
    'Optimized complex data reconciliation pipelines, dropping operational TAT from 2 days to 30 minutes by automating ledger updates.',
    'Led cross-functional alignment across Engineering, Data Science, and Finance to ship accurate financial reporting at a 350M+ MAU scale.',
  ];
  leftPts.forEach((pt, i) => {
    slide.addText('• ' + pt, {
      x: 0.7, y: 2.0 + i * 1.08, w: 3.9, h: 0.96,
      fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });

  // Right — light column
  slide.addShape(prs.ShapeType.rect, {
    x: 5.1, y: 1.38, w: 4.3, h: 5.3,
    fill: { color: C.white }, line: { color: C.bright, pt: 1.5, transparency: 40 }, rectRadius: 0.1,
    shadow: makeShadow(),
  });
  slide.addText('How It Maps to the Dezerv JD', {
    x: 5.3, y: 1.56, w: 3.9, h: 0.34,
    fontSize: 12, bold: true, color: C.ink, fontFace: 'Calibri',
  });
  const rightPts = [
    ['Client Reporting Reliability', 'PhonePe B2B partner dashboards → Dezerv High-Net-Worth client reporting'],
    ['Ops TAT Reduction', '2 days to 30 mins at PhonePe → Real-time Dezerv settlement steppers'],
    ['Owning Execution', 'Led marketplace execution and P&L → Dezerv trust & execution accountability'],
    ['Cross-functional Leadership', 'Shipped products across tech, ops, finance → Dezerv PM culture'],
  ];
  rightPts.forEach((pt, i) => {
    const y = 2.0 + i * 1.08;
    slide.addText(pt[0], {
      x: 5.3, y, w: 3.9, h: 0.28,
      fontSize: 10.5, bold: true, color: C.ink, fontFace: 'Calibri',
    });
    slide.addText(pt[1], {
      x: 5.3, y: y + 0.28, w: 3.9, h: 0.68,
      fontSize: 9.5, color: C.slate, fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  slide.addText('ROLLOUT PLAN', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('Phased Deployment — Data First, Nudges Later', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 24, bold: true, color: C.white, fontFace: 'Calibri',
  });

  const phases = [
    {
      phase: 'Phase 1',
      time: 'Month 1–2 · Foundation',
      title: 'Unified Ledgers & Live Settlement',
      pts: [
        'Audit AMC and internal Dezerv data pipelines to build a unified client schema',
        'Deploy the Real-Time Settlement Stepper for PMS & AIF transactions',
        'Beta test with top 5% of RM clients to ensure data accuracy',
        'Measure reduction in RM support queries regarding execution status',
      ],
      color: C.bright,
    },
    {
      phase: 'Phase 2',
      time: 'Month 3–4 · Transparency',
      title: 'True Net Yield & Taxes',
      pts: [
        'Integrate automated tax projection logic (STCG/LTCG calculations)',
        'Roll out the transparent Fee & Tax Breakdown dashboard',
        'Monitor client sentiment and RM feedback on the new gross vs net view',
        'Publish the feature as a core Dezerv trust initiative to the press',
      ],
      color: C.orange,
    },
    {
      phase: 'Phase 3',
      time: 'Month 5–6 · Optimisation',
      title: 'Wealth Predictor & Nudges',
      pts: [
        'Deploy AI projection models based on client risk profiles and history',
        'Implement contextual "Increase SIP" nudges within the reporting flow',
        'A/B test nudge copy and timing for maximum conversion',
        'Track organic AUM growth directly attributed to the WealthTrack dashboard',
      ],
      color: '#10B981',
    },
  ];

  phases.forEach((p, i) => {
    const x = 0.5 + i * 3.1;
    slide.addShape(prs.ShapeType.rect, {
      x, y: 1.42, w: 2.88, h: 4.5,
      fill: { color: C.hero }, line: { color: p.color, pt: 1.5, transparency: 30 }, rectRadius: 0.1,
    });
    slide.addText(p.phase, {
      x: x + 0.14, y: 1.58, w: 2.62, h: 0.28,
      fontSize: 9, bold: true, color: p.color, charSpacing: 2, fontFace: 'Calibri',
    });
    slide.addText(p.time, {
      x: x + 0.14, y: 1.86, w: 2.62, h: 0.28,
      fontSize: 10, color: C.muted, fontFace: 'Calibri', italic: true,
    });
    slide.addText(p.title, {
      x: x + 0.14, y: 2.18, w: 2.62, h: 0.44,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri', lineSpacingMultiple: 1.2,
    });
    p.pts.forEach((pt, j) => {
      slide.addText('• ' + pt, {
        x: x + 0.14, y: 2.68 + j * 0.74, w: 2.62, h: 0.66,
        fontSize: 9, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
      });
    });
  });

  // Ask box
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.1, w: 9, h: 1.02,
    fill: { color: C.hero }, line: { color: C.bright, pt: 1.5, transparency: 30 }, rectRadius: 0.1,
  });
  slide.addText('What I need to build this at Dezerv:', {
    x: 0.72, y: 6.2, w: 3, h: 0.28,
    fontSize: 10, bold: true, color: C.bright, fontFace: 'Calibri',
  });
  slide.addText('Access to internal reporting schemas  ·  Alignment with Engineering on AMC webhook reliability  ·  1 frontend dev for high-fidelity UI execution  ·  RM team bandwidth for beta testing  ·  Clear runway to own reporting UX and analytics metrics', {
    x: 0.72, y: 6.5, w: 8.6, h: 0.54,
    fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
  });

  // Contact footer
  slide.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade', {
    x: 0.5, y: 7.2, w: 9, h: 0.26,
    fontSize: 8.5, color: C.muted, fontFace: 'Calibri', align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/dezerv-wealthtrack-deck.pptx' })
  .then(() => console.log('dezerv-wealthtrack-deck.pptx generated'))
  .catch(e => console.error(e));
