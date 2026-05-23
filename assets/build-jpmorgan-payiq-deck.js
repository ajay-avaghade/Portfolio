const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:  '0A1E3D',
  hero:  '0F2A54',
  blue:  '2563EB',
  lblue: '3B82F6',
  gold:  'C9A227',
  white: 'FFFFFF',
  lgray: 'F1F5F9',
  muted: '64748B',
  teal:  '0D9488',
  green: '16A34A',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });

// ─────────────────────────────────────────
// SLIDE 1 — COVER
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  // Diagonal line motif (bottom-right)
  for (let i = 0; i < 6; i++) {
    sl.addShape(prs.ShapeType.rect, {
      x: 7.0 + i * 0.45, y: -0.2, w: 0.18, h: 9.0,
      fill: { color: C.blue, transparency: 88 - i * 10 },
      line: { type: 'none' },
      rotate: 20,
    });
  }

  // Blue accent left strip
  sl.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: 0.18, h: 7.5,
    fill: { color: C.blue }, line: { type: 'none' },
  });

  // Tag
  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 0.38, w: 3.2, h: 0.28,
    fill: { color: C.blue, transparency: 80 },
    line: { color: C.blue, transparency: 60, pt: 0.5 },
    rectRadius: 0.14,
  });
  sl.addText('J.P. MORGAN · CEEMEA CORE CASH PAYMENTS', {
    x: 0.5, y: 0.38, w: 3.2, h: 0.28,
    fontSize: 7, bold: true, color: C.gold,
    align: 'center', valign: 'middle', fontFace: 'Calibri',
  });

  // Product name
  sl.addText('PayIQ', {
    x: 0.48, y: 0.82, w: 6.5, h: 1.35,
    fontSize: 72, bold: true, color: C.white,
    fontFace: 'Calibri', charSpacing: -2,
  });

  // Subtitle
  sl.addText('CEEMEA Payment Intelligence Platform', {
    x: 0.48, y: 2.0, w: 6.2, h: 0.45,
    fontSize: 20, bold: false, color: C.lblue,
    fontFace: 'Calibri',
  });

  // Description line
  sl.addText('Smart rail routing before · Real-time gpi visibility after · Zero new infrastructure', {
    x: 0.48, y: 2.52, w: 6.5, h: 0.32,
    fontSize: 12, color: 'A0AABA', fontFace: 'Calibri',
  });

  // Presenter
  sl.addText('Presented by Ajay Avaghade · Growth & Monetization PM', {
    x: 0.48, y: 3.0, w: 5.5, h: 0.28,
    fontSize: 10, color: '8899AA', fontFace: 'Calibri', italic: true,
  });

  // Big stat bottom-right
  sl.addShape(prs.ShapeType.rect, {
    x: 7.4, y: 4.8, w: 2.3, h: 1.9,
    fill: { color: C.blue, transparency: 78 },
    line: { color: C.blue, transparency: 55, pt: 1 },
    rectRadius: 0.14,
  });
  sl.addText('USD 18M', {
    x: 7.4, y: 4.82, w: 2.3, h: 0.65,
    fontSize: 30, bold: true, color: C.gold,
    fontFace: 'Calibri', align: 'center',
  });
  sl.addText('annual fee revenue\nuplift projected', {
    x: 7.4, y: 5.44, w: 2.3, h: 0.6,
    fontSize: 10, color: 'A0AABA', fontFace: 'Calibri', align: 'center', valign: 'top',
  });

  // Footer line
  sl.addShape(prs.ShapeType.rect, {
    x: 0, y: 7.2, w: 10, h: 0.02,
    fill: { color: C.blue, transparency: 70 }, line: { type: 'none' },
  });
}

// ─────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.hero };

  // Eyebrow
  sl.addText('THE PROBLEM', {
    x: 0.5, y: 0.38, w: 9, h: 0.24,
    fontSize: 9, bold: true, color: C.gold,
    fontFace: 'Calibri', charSpacing: 1.5,
  });

  // H1
  sl.addText('CEEMEA Corporates Are Losing Millions\nEvery Time They Default to SWIFT', {
    x: 0.5, y: 0.66, w: 9, h: 1.1,
    fontSize: 26, bold: true, color: C.white,
    fontFace: 'Calibri', charSpacing: -0.5,
  });

  // 3 stat columns
  const cols3 = [
    { icon: '💸', stat: '4×', lbl: 'Overpayment vs optimal rail', sub: 'A Dubai→Cairo payment via SWIFT costs 4× more than Egypt Instant Pay with 2-day extra delay' },
    { icon: '📞', stat: '35%', lbl: 'Of RM time chasing payment status', sub: 'Corporates call RMs daily for status updates — J.P. Morgan relationship teams act as manual payment trackers' },
    { icon: '🌐', stat: '50+', lbl: 'CEEMEA payment rails now live', sub: 'SARIE, AANI, KEPSS, Egypt Instant Pay, RTGS systems — all with different cut-offs, costs, and STP rates' },
  ];

  cols3.forEach((c, i) => {
    const x = 0.5 + i * 3.1;
    sl.addShape(prs.ShapeType.rect, {
      x, y: 2.0, w: 2.85, h: 3.4,
      fill: { color: 'FFFFFF', transparency: 93 },
      line: { color: C.blue, transparency: 70, pt: 0.75 },
      rectRadius: 0.12,
    });
    sl.addText(c.icon, {
      x, y: 2.12, w: 2.85, h: 0.45,
      fontSize: 22, align: 'center', fontFace: 'Calibri',
    });
    sl.addText(c.stat, {
      x, y: 2.62, w: 2.85, h: 0.72,
      fontSize: 44, bold: true, color: C.gold,
      fontFace: 'Calibri', align: 'center',
    });
    sl.addText(c.lbl, {
      x: x + 0.12, y: 3.38, w: 2.6, h: 0.48,
      fontSize: 11, bold: true, color: C.white,
      fontFace: 'Calibri', align: 'center',
    });
    sl.addText(c.sub, {
      x: x + 0.12, y: 3.9, w: 2.6, h: 0.85,
      fontSize: 9.5, color: 'A0AABA', fontFace: 'Calibri',
      align: 'center', valign: 'top',
    });
  });

  // Insight box
  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 5.6, w: 9, h: 1.42,
    fill: { color: C.blue, transparency: 85 },
    line: { color: C.blue, transparency: 55, pt: 0.75 },
    rectRadius: 0.12,
  });
  sl.addText('Root cause:  ', {
    x: 0.72, y: 5.72, w: 8.6, h: 0.28,
    fontSize: 11, bold: true, color: C.gold, fontFace: 'Calibri',
  });
  sl.addText('J.P. Morgan has real-time rail connectivity to SARIE, AANI, Egypt Instant Pay, and 40+ CEEMEA local rails — but has not productised that infrastructure into a client-facing intelligence layer. Clients default to SWIFT because they have no visibility into alternatives. Every suboptimal route is simultaneously a client experience failure and a J.P. Morgan revenue leak.', {
    x: 0.72, y: 6.02, w: 8.6, h: 0.85,
    fontSize: 10.5, color: 'C8D8E8', fontFace: 'Calibri', valign: 'top',
  });
}

// ─────────────────────────────────────────
// SLIDE 3 — THE INSIGHT
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE INSIGHT', {
    x: 0.5, y: 0.38, w: 9, h: 0.24,
    fontSize: 9, bold: true, color: C.blue,
    fontFace: 'Calibri', charSpacing: 1.5,
  });
  sl.addText('Rail Intelligence is a Revenue Product,\nNot an Ops Feature', {
    x: 0.5, y: 0.66, w: 9, h: 1.1,
    fontSize: 26, bold: true, color: C.dark,
    fontFace: 'Calibri', charSpacing: -0.5,
  });

  // Left column (current state)
  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.9, w: 4.1, h: 4.7,
    fill: { color: 'FEF2F2', transparency: 0 },
    line: { color: 'FCA5A5', pt: 0.75 },
    rectRadius: 0.12,
    shadow: makeShadow(),
  });
  sl.addText('❌  Current State', {
    x: 0.7, y: 2.04, w: 3.7, h: 0.35,
    fontSize: 13, bold: true, color: 'DC2626', fontFace: 'Calibri',
  });
  const currentPoints = [
    'Client initiates payment → SWIFT by default',
    'No visibility into SARIE, AANI, local RTGS options',
    'Average transaction cost: USD 5,120 via SWIFT',
    'Average delivery time: 1–2 business days',
    'RM receives 3–5 status calls per transaction',
    'ISO 20022 data sits in bank feeds — never surfaces to client',
  ];
  currentPoints.forEach((pt, i) => {
    sl.addText(`• ${pt}`, {
      x: 0.7, y: 2.5 + i * 0.55, w: 3.7, h: 0.48,
      fontSize: 10.5, color: '1A1A2E', fontFace: 'Calibri', valign: 'top',
    });
  });

  // VS divider
  sl.addShape(prs.ShapeType.ellipse, {
    x: 4.62, y: 3.85, w: 0.76, h: 0.76,
    fill: { color: C.dark }, line: { type: 'none' },
  });
  sl.addText('VS', {
    x: 4.62, y: 3.85, w: 0.76, h: 0.76,
    fontSize: 10, bold: true, color: C.white,
    fontFace: 'Calibri', align: 'center', valign: 'middle',
  });

  // Right column (PayIQ)
  sl.addShape(prs.ShapeType.rect, {
    x: 5.4, y: 1.9, w: 4.1, h: 4.7,
    fill: { color: 'EFF6FF', transparency: 0 },
    line: { color: C.blue, pt: 0.75 },
    rectRadius: 0.12,
    shadow: makeShadow(),
  });
  sl.addText('✅  PayIQ', {
    x: 5.6, y: 2.04, w: 3.7, h: 0.35,
    fontSize: 13, bold: true, color: C.blue, fontFace: 'Calibri',
  });
  const payiqPoints = [
    'ML engine evaluates 7 variables → recommends optimal rail',
    'Comparison matrix: cost / speed / certainty / cut-off',
    'Average transaction cost: USD 320–890 via local rails',
    'Average delivery: 2–4 hours via SARIE / AANI',
    'gpi + ISO 20022 status surfaced in real-time client portal',
    'RM query volume drops 65% — zero relationship degradation',
  ];
  payiqPoints.forEach((pt, i) => {
    sl.addText(`• ${pt}`, {
      x: 5.6, y: 2.5 + i * 0.55, w: 3.7, h: 0.48,
      fontSize: 10.5, color: '1A1A2E', fontFace: 'Calibri', valign: 'top',
    });
  });

  // Callout box
  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.78, w: 9, h: 0.52,
    fill: { color: C.gold, transparency: 88 },
    line: { color: C.gold, transparency: 60, pt: 0.75 },
    rectRadius: 0.1,
  });
  sl.addText('Key insight:  PayIQ costs J.P. Morgan nothing net — it runs on infrastructure that already exists. The value is the intelligence layer on top.', {
    x: 0.7, y: 6.8, w: 8.6, h: 0.46,
    fontSize: 10, color: C.dark, fontFace: 'Calibri', valign: 'middle',
    bold: false,
  });
}

// ─────────────────────────────────────────
// SLIDE 4 — THE MECHANIC
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE MECHANIC', {
    x: 0.5, y: 0.38, w: 9, h: 0.24,
    fontSize: 9, bold: true, color: C.gold,
    fontFace: 'Calibri', charSpacing: 1.5,
  });
  sl.addText('How PayIQ Works: Two Moments, One Intelligent Layer', {
    x: 0.5, y: 0.66, w: 9, h: 0.7,
    fontSize: 26, bold: true, color: C.white,
    fontFace: 'Calibri', charSpacing: -0.5,
  });

  const steps = [
    { n: '01', title: 'Client Initiates', body: 'Corporate initiates payment in J.P. Morgan ACCESS portal. Amount, beneficiary, currency, urgency captured.' },
    { n: '02', title: 'PayIQ Scores Rails', body: '7-variable ML model evaluates available CEEMEA rails: amount tier, corridor, cut-off proximity, STP rate, FX spread, certainty, urgency flag.' },
    { n: '03', title: 'Recommendation Served', body: 'Top rail presented with one-click confirm. Comparison matrix available. Client sees cost savings vs SWIFT and ETA.', hl: true },
    { n: '04', title: 'gpi + ISO 20022 Tracking', body: 'Post-submission: real-time SWIFT gpi tracker surfaces in ACCESS. Every corridor hop visible. Exceptions surfaced with self-serve resolution.' },
    { n: '05', title: 'RM Dashboard Updated', body: 'PayIQ logs rail choice, delivery time, cost delta. RM sees client savings, exception rate, and rail utilisation across entire portfolio.' },
  ];

  steps.forEach((st, i) => {
    const x = 0.3 + i * 1.9;
    const highlighted = st.hl;
    sl.addShape(prs.ShapeType.rect, {
      x, y: 1.55, w: 1.7, h: 4.7,
      fill: { color: highlighted ? C.blue : 'FFFFFF', transparency: highlighted ? 0 : 93 },
      line: { color: highlighted ? C.blue : C.blue, transparency: highlighted ? 0 : 75, pt: highlighted ? 0 : 0.5 },
      rectRadius: 0.12,
    });
    sl.addText(st.n, {
      x, y: 1.7, w: 1.7, h: 0.42,
      fontSize: 24, bold: true,
      color: highlighted ? C.gold : C.lblue,
      fontFace: 'Calibri', align: 'center',
    });
    sl.addText(st.title, {
      x: x + 0.1, y: 2.18, w: 1.5, h: 0.55,
      fontSize: 11, bold: true,
      color: highlighted ? C.white : C.white,
      fontFace: 'Calibri', align: 'center',
    });
    sl.addText(st.body, {
      x: x + 0.1, y: 2.78, w: 1.5, h: 2.5,
      fontSize: 9.5,
      color: highlighted ? 'D0E4FF' : 'A0AABA',
      fontFace: 'Calibri', valign: 'top',
    });

    // Connector arrow (except last)
    if (i < steps.length - 1) {
      sl.addText('→', {
        x: x + 1.72, y: 3.8, w: 0.18, h: 0.3,
        fontSize: 11, color: C.blue, fontFace: 'Calibri', align: 'center',
      });
    }
  });

  // Proof note
  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.46, w: 9, h: 0.78,
    fill: { color: C.gold, transparency: 88 },
    line: { color: C.gold, transparency: 60, pt: 0.5 },
    rectRadius: 0.1,
  });
  sl.addText('PhonePe proof:  At PhonePe (350M MAU), I built the exact same context-aware multi-variable trigger engine — cart value × user intent × time-of-day → personalised incentive. PayIQ is the same architecture: payment amount × corridor × urgency × cut-off → optimal rail. The model is proven. The domain is different.', {
    x: 0.7, y: 6.5, w: 8.6, h: 0.72,
    fontSize: 9.5, color: C.gold, fontFace: 'Calibri', valign: 'middle',
  });
}

// ─────────────────────────────────────────
// SLIDE 5 — THE PRODUCT
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE PRODUCT', {
    x: 0.5, y: 0.38, w: 9, h: 0.24,
    fontSize: 9, bold: true, color: C.blue,
    fontFace: 'Calibri', charSpacing: 1.5,
  });
  sl.addText('4 Key Screens — Built & Live as a Working Prototype', {
    x: 0.5, y: 0.66, w: 9, h: 0.7,
    fontSize: 26, bold: true, color: C.dark,
    fontFace: 'Calibri', charSpacing: -0.5,
  });

  const screens = [
    {
      n: '01', title: 'Initiation + Rail Rec', color: C.blue,
      desc: 'Payment form with urgency toggle. PayIQ recommendation card shows optimal rail, ETA, cost vs SWIFT, and ML confidence score. One-click confirm.',
      ascii: '[ New Payment          ]\n[ Beneficiary: Al Rajhi]\n[ Amount: USD 2.4M     ]\n[ Urgency: [Today] ▼   ]\n─────────────────────── \n⚡ PAYIQ: SARIE       \n  ~2 hrs · −USD 4,800  \n  vs SWIFT · 99.1% STP \n───────────────────────\n[  Confirm with SARIE →]',
    },
    {
      n: '02', title: 'Rail Comparison Matrix', color: C.gold,
      desc: 'Full rail matrix: speed, fee, STP rate, cut-off time per rail. SARIE highlighted as best. AANI/RTGS/SWIFT shown for comparison. Cut-off countdown.',
      ascii: 'Rail    Speed  Fee    STP\n─────────────────────── \nSARIE⭐ 2hr   $320  99%\nSWIFT   2d   $5120  97%\nSA-RTGS  1d    $890  98%\nAANI    N/A    —     — \n─────────────────────── \n⏱ SARIE cut-off: 2h14m\n[Confirm: SARIE −$4800→]',
    },
    {
      n: '03', title: 'gpi Payment Delivered', color: C.teal,
      desc: 'Post-payment: SWIFT gpi corridor tracker with 4 steps. ISO 20022 status. Amount confirmed. Rail savings displayed. UETR reference.',
      ascii: '✓  Payment Delivered    \n─────────────────────── \nUSD 2,400,000 → Al Rajhi\nVia SARIE · 2h 24min   \nSaved: USD 4,800 vs SWIFT\n─────────────────────── \ngpi Steps:             \n✓ Initiated (JPM Dubai)\n✓ SARIE Network        \n✓ Saudi Hub            \n✓ Al Rajhi — Credited  ',
    },
    {
      n: '04', title: 'RM / Ops Dashboard', color: C.dark,
      desc: 'Portfolio-level rail utilisation. Exception queue with self-serve flags. Client savings dashboard. SWIFT share drop from 68% to 21% tracked.',
      ascii: 'PayIQ · CEEMEA Ops     \n─────────────────────── \n847 queries resolved   \nUSD 2.1M rail savings  \n1.2% exception rate    \n─────────────────────── \nRail Mix (May 2026)    \nSARIE    ████░ 34%    \nAANI     ███░░ 28%    \nSWIFT    ██░░░ 21%    \nRTGS     █░░░░ 17%   ',
    },
  ];

  screens.forEach((sc, i) => {
    const x = 0.3 + i * 2.42;
    sl.addShape(prs.ShapeType.rect, {
      x, y: 1.55, w: 2.22, h: 5.6,
      fill: { color: C.white }, line: { color: sc.color, pt: 1 },
      rectRadius: 0.12, shadow: makeShadow(),
    });
    // Screen number
    sl.addShape(prs.ShapeType.rect, {
      x, y: 1.55, w: 2.22, h: 0.45,
      fill: { color: sc.color, transparency: sc.color === C.dark ? 0 : 0 },
      line: { type: 'none' }, rectRadius: 0.12,
    });
    sl.addText(`${sc.n} — ${sc.title}`, {
      x: x + 0.08, y: 1.55, w: 2.06, h: 0.45,
      fontSize: 9, bold: true, color: C.white,
      fontFace: 'Calibri', align: 'center', valign: 'middle',
    });
    // ASCII mockup
    sl.addText(sc.ascii, {
      x: x + 0.1, y: 2.08, w: 2.02, h: 2.55,
      fontSize: 7, color: '1E293B',
      fontFace: 'Courier New', valign: 'top',
    });
    // Description
    sl.addText(sc.desc, {
      x: x + 0.1, y: 4.72, w: 2.02, h: 2.1,
      fontSize: 9.5, color: C.muted, fontFace: 'Calibri', valign: 'top',
    });
  });

  sl.addText('Prototype live at: ajay-avaghade.github.io/Portfolio/jpmorgan-payiq-prototype.html', {
    x: 0.5, y: 7.2, w: 9, h: 0.24,
    fontSize: 9, color: C.blue, fontFace: 'Calibri', italic: true, align: 'center',
  });
}

// ─────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.hero };

  sl.addText('IMPACT & ROI', {
    x: 0.5, y: 0.38, w: 9, h: 0.24,
    fontSize: 9, bold: true, color: C.gold,
    fontFace: 'Calibri', charSpacing: 1.5,
  });
  sl.addText('Projected Impact — Built on PhonePe Proof Points', {
    x: 0.5, y: 0.66, w: 9, h: 0.7,
    fontSize: 26, bold: true, color: C.white,
    fontFace: 'Calibri', charSpacing: -0.5,
  });

  // Left section header
  sl.addText('CLIENT IMPACT', {
    x: 0.5, y: 1.55, w: 4.4, h: 0.28,
    fontSize: 9, bold: true, color: C.lblue,
    fontFace: 'Calibri', charSpacing: 1.2,
  });

  const clientMetrics = [
    { val: '−USD 4,800', lbl: 'Average saving per cross-border payment vs SWIFT', color: C.gold },
    { val: '~2 hrs', lbl: 'Average CEEMEA delivery time via optimised rail vs 1–2 days', color: C.teal },
    { val: '−65%', lbl: 'RM payment status queries per client per month', color: C.lblue },
    { val: '99.1%', lbl: 'STP rate via PayIQ-recommended rail vs 87% SWIFT average', color: C.gold },
  ];
  clientMetrics.forEach((m, i) => {
    const row = i % 2; const col = Math.floor(i / 2);
    const x = 0.5 + col * 2.3; const y = 1.92 + row * 1.55;
    sl.addShape(prs.ShapeType.rect, {
      x, y, w: 2.1, h: 1.38,
      fill: { color: 'FFFFFF', transparency: 93 },
      line: { color: C.blue, transparency: 65, pt: 0.5 },
      rectRadius: 0.1,
    });
    sl.addText(m.val, {
      x: x + 0.1, y: y + 0.12, w: 1.9, h: 0.5,
      fontSize: 24, bold: true, color: m.color, fontFace: 'Calibri',
    });
    sl.addText(m.lbl, {
      x: x + 0.1, y: y + 0.64, w: 1.9, h: 0.62,
      fontSize: 9, color: 'A0AABA', fontFace: 'Calibri', valign: 'top',
    });
  });

  // Divider
  sl.addShape(prs.ShapeType.rect, {
    x: 5.05, y: 1.55, w: 0.02, h: 5.0,
    fill: { color: C.blue, transparency: 70 }, line: { type: 'none' },
  });

  // Right section header
  sl.addText('J.P. MORGAN P&L', {
    x: 5.3, y: 1.55, w: 4.2, h: 0.28,
    fontSize: 9, bold: true, color: C.gold,
    fontFace: 'Calibri', charSpacing: 1.2,
  });

  const jpmMetrics = [
    { val: '+USD 18M', lbl: 'Annual fee revenue uplift from rail-optimised transaction mix', color: C.gold },
    { val: '+USD 2.1M', lbl: 'Monthly FX spread revenue via local rail conversions', color: C.lblue },
    { val: '−68%', lbl: 'SWIFT share of CEEMEA transactions (from 100% baseline)', color: C.teal },
    { val: '+34%', lbl: 'Client NPS improvement on payment experience (projected)', color: C.gold },
  ];
  jpmMetrics.forEach((m, i) => {
    const row = i % 2; const col = Math.floor(i / 2);
    const x = 5.3 + col * 2.2; const y = 1.92 + row * 1.55;
    sl.addShape(prs.ShapeType.rect, {
      x, y, w: 2.0, h: 1.38,
      fill: { color: 'FFFFFF', transparency: 93 },
      line: { color: C.gold, transparency: 60, pt: 0.5 },
      rectRadius: 0.1,
    });
    sl.addText(m.val, {
      x: x + 0.1, y: y + 0.12, w: 1.8, h: 0.5,
      fontSize: 22, bold: true, color: m.color, fontFace: 'Calibri',
    });
    sl.addText(m.lbl, {
      x: x + 0.1, y: y + 0.64, w: 1.8, h: 0.62,
      fontSize: 9, color: 'A0AABA', fontFace: 'Calibri', valign: 'top',
    });
  });

  // Insight box
  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 5.2, w: 9, h: 1.0,
    fill: { color: C.gold, transparency: 88 },
    line: { color: C.gold, transparency: 60, pt: 0.75 },
    rectRadius: 0.12,
  });
  sl.addText('The PayIQ investment thesis:  Every transaction PayIQ re-routes from SWIFT to a local rail saves the client USD 2,000–5,000 AND generates an additional FX conversion for J.P. Morgan. The product simultaneously improves client outcomes and J.P. Morgan P&L — the rare case where both vectors point the same direction.', {
    x: 0.7, y: 5.25, w: 8.6, h: 0.9,
    fontSize: 10.5, color: C.gold, fontFace: 'Calibri', valign: 'middle',
  });
}

// ─────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('PROOF OF WORK', {
    x: 0.5, y: 0.38, w: 9, h: 0.24,
    fontSize: 9, bold: true, color: C.blue,
    fontFace: 'Calibri', charSpacing: 1.5,
  });
  sl.addText('I Built This Architecture. Here is the Evidence.', {
    x: 0.5, y: 0.66, w: 9, h: 0.7,
    fontSize: 26, bold: true, color: C.dark,
    fontFace: 'Calibri', charSpacing: -0.5,
  });

  // Left dark column
  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.55, w: 4.3, h: 5.3,
    fill: { color: C.dark }, line: { type: 'none' }, rectRadius: 0.14,
  });
  sl.addText('What I built at PhonePe', {
    x: 0.7, y: 1.7, w: 3.9, h: 0.34,
    fontSize: 12, bold: true, color: C.gold, fontFace: 'Calibri',
  });
  const phonepe = [
    'Propensity-to-Transact ML model: 7-variable scoring engine replacing static rule-based cohort targeting across 350M MAU — same multi-variable inference architecture as PayIQ rail scoring',
    'Context-aware trigger engine: cart value × user intent × time signals → personalised incentive. Identical logic to PayIQ: payment × corridor × urgency × cut-off → optimal rail',
    'Real-time offer eligibility surfacing: aggregated status from multiple data sources into a single clean user-facing layer — same pattern as gpi + ISO 20022 into PayIQ dashboard',
    'Ops dashboard for campaign managers: portfolio-level analytics, exception queue, real-time signal monitoring — maps directly to PayIQ RM dashboard (Slide 5 Screen 4)',
  ];
  phonepe.forEach((pt, i) => {
    sl.addText(`${i + 1}.  ${pt}`, {
      x: 0.7, y: 2.15 + i * 1.1, w: 3.9, h: 0.95,
      fontSize: 9.5, color: 'A0AABA', fontFace: 'Calibri', valign: 'top',
    });
  });

  // Right light column
  sl.addShape(prs.ShapeType.rect, {
    x: 5.1, y: 1.55, w: 4.4, h: 5.3,
    fill: { color: 'EFF6FF' }, line: { color: C.blue, pt: 0.75 }, rectRadius: 0.14,
    shadow: makeShadow(),
  });
  sl.addText('How it maps to this JD', {
    x: 5.3, y: 1.7, w: 4.0, h: 0.34,
    fontSize: 12, bold: true, color: C.blue, fontFace: 'Calibri',
  });
  const mapping = [
    '"Extensive experience in designing and implementing solutions" → ML-driven incentive platform deployed at 350M MAU scale',
    '"Define and delivers product strategy ensuring market-leading capabilities" → PayIQ: first in market to productise CEEMEA rail intelligence into a client-facing layer',
    '"Own P&L including delivery against revenue and expense targets" → 32% marketing burn reduction at PhonePe demonstrates cost-side P&L management alongside GMV growth',
    '"Understand risks inherent in the product" → Built AML-adjacent compliance eligibility engine at PhonePe; ClearPath demonstrates risk product thinking for J.P. Morgan context',
  ];
  mapping.forEach((pt, i) => {
    sl.addText(`• ${pt}`, {
      x: 5.3, y: 2.15 + i * 1.1, w: 3.9, h: 0.95,
      fontSize: 9.5, color: '1A1A2E', fontFace: 'Calibri', valign: 'top',
    });
  });

  // Quote box
  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 7.0, w: 9, h: 0.28,
    fill: { color: C.blue, transparency: 90 },
    line: { color: C.blue, transparency: 60, pt: 0.5 },
    rectRadius: 0.08,
  });
  sl.addText('"PayIQ is not a concept — it is the same context-aware ML engine I deployed at PhonePe applied to CEEMEA payment rail selection."', {
    x: 0.7, y: 7.0, w: 8.6, h: 0.28,
    fontSize: 9, italic: true, color: C.dark, fontFace: 'Calibri', valign: 'middle',
  });
}

// ─────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('ROLLOUT PLAN', {
    x: 0.5, y: 0.38, w: 9, h: 0.24,
    fontSize: 9, bold: true, color: C.gold,
    fontFace: 'Calibri', charSpacing: 1.5,
  });
  sl.addText('Phased Rollout — UAE/KSA Pilot to CEEMEA Scale', {
    x: 0.5, y: 0.66, w: 9, h: 0.7,
    fontSize: 26, bold: true, color: C.white,
    fontFace: 'Calibri', charSpacing: -0.5,
  });

  const phases = [
    {
      ph: 'Phase 1', period: 'Months 1–3', title: 'UAE + KSA Pilot',
      body: 'Deploy PayIQ for AANI (UAE) + SARIE (Saudi) corridors — highest-volume, highest-savings opportunity. Pilot with 5 corporate clients. Instrument rail utilisation, client savings, and RM query volume.',
      color: C.blue,
    },
    {
      ph: 'Phase 2', period: 'Months 4–6', title: 'CEEMEA Rail Expansion',
      body: 'Add Egypt Instant Pay, KEPSS (Kenya), Nigeria NIP, and 10+ additional CEEMEA local rails to PayIQ model. Launch gpi + ISO 20022 status dashboard for all enrolled clients. Begin A/B on opt-in vs default routing.',
      color: C.gold,
    },
    {
      ph: 'Phase 3', period: 'Months 7–12', title: 'Full Platform + P&L Dashboard',
      body: 'Roll PayIQ to all CEEMEA corporate clients as default. Launch RM revenue dashboard showing rail optimisation savings per client. Integrate with J.P. Morgan treasury sales for LiquidSense expansion.',
      color: C.teal,
    },
  ];

  phases.forEach((ph, i) => {
    const x = 0.5 + i * 3.18;
    sl.addShape(prs.ShapeType.rect, {
      x, y: 1.55, w: 2.95, h: 4.3,
      fill: { color: 'FFFFFF', transparency: 93 },
      line: { color: ph.color, transparency: 55, pt: 1 },
      rectRadius: 0.14,
    });
    sl.addShape(prs.ShapeType.rect, {
      x, y: 1.55, w: 2.95, h: 0.42,
      fill: { color: ph.color, transparency: 0 },
      line: { type: 'none' }, rectRadius: 0.14,
    });
    sl.addText(`${ph.ph}  ·  ${ph.period}`, {
      x: x + 0.1, y: 1.55, w: 2.75, h: 0.42,
      fontSize: 9.5, bold: true, color: C.dark,
      fontFace: 'Calibri', align: 'center', valign: 'middle',
    });
    sl.addText(ph.title, {
      x: x + 0.1, y: 2.08, w: 2.75, h: 0.44,
      fontSize: 14, bold: true, color: C.white, fontFace: 'Calibri',
    });
    sl.addText(ph.body, {
      x: x + 0.1, y: 2.6, w: 2.75, h: 2.7,
      fontSize: 10, color: 'A0AABA', fontFace: 'Calibri', valign: 'top',
    });
  });

  // Ask box
  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.02, w: 9, h: 0.88,
    fill: { color: C.blue, transparency: 85 },
    line: { color: C.blue, transparency: 55, pt: 0.75 },
    rectRadius: 0.12,
  });
  sl.addText('What I need to build this', {
    x: 0.7, y: 6.08, w: 8.6, h: 0.26,
    fontSize: 10, bold: true, color: C.gold, fontFace: 'Calibri',
  });
  sl.addText('Access to CEEMEA rail connectivity data  ·  Alignment with ACCESS portal team  ·  1 data scientist for ML rail model  ·  SWIFT gpi + ISO 20022 data feed access  ·  5 pilot corporate client nominations from Sales', {
    x: 0.7, y: 6.36, w: 8.6, h: 0.44,
    fontSize: 9.5, color: 'C8D8E8', fontFace: 'Calibri', valign: 'top',
  });

  // Footer
  sl.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  ajay-avaghade.github.io/Portfolio', {
    x: 0.5, y: 7.12, w: 9, h: 0.24,
    fontSize: 8.5, color: '6B7E9A', fontFace: 'Calibri', align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/jpmorgan-payiq-deck.pptx' }).then(() => {
  console.log('jpmorgan-payiq-deck.pptx done');
}).catch(e => console.error(e));
