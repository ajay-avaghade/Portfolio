// build-navi-growthos-deck.js
// Run: node assets/build-navi-growthos-deck.js
// Output: assets/navi-growthos-deck.pptx

const pptx = require('pptxgenjs');
const prs = new pptx();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '071A12',
  hero:   '0E2A1C',
  bright: '00C853',
  gold:   'FFB800',
  orange: 'FF6B35',
  white:  'FFFFFF',
  lgray:  'F0F7F2',
  muted:  '5C7A69',
  ink:    '0D1F15',
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 4, offset: 2, angle: 45, opacity: 0.14 });

// ─────────────────────────────────────────────
// SLIDE 1 — COVER (dark)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  // Diagonal line motif
  for (let i = 0; i < 8; i++) {
    s.addShape(prs.ShapeType.line, {
      x: -0.5 + i * 1.5, y: 0, w: 3.5, h: 0,
      line: { color: C.bright, width: 0.4, transparency: 88 },
      rotate: 35,
    });
  }

  // Top eyebrow
  s.addText('NAVI TECHNOLOGIES  ·  PRODUCT MANAGER — MARTECH & GROWTH', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 8.5, bold: false, color: C.bright,
    charSpacing: 2, fontFace: 'Calibri',
  });

  // Main title
  s.addText('GrowthOS', {
    x: 0.5, y: 0.68, w: 9, h: 1.2,
    fontSize: 68, bold: true, color: C.white, fontFace: 'Calibri',
    shadow: makeShadow(),
  });

  // Green accent bar
  s.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.82, w: 1.6, h: 0.07,
    fill: { color: C.bright }, line: { color: C.bright },
  });

  // Subtitle
  s.addText('The MarTech Intelligence Layer Powering Navi CLTV', {
    x: 0.5, y: 1.98, w: 7.5, h: 0.38,
    fontSize: 18, bold: false, color: 'C8E6C9', fontFace: 'Calibri',
  });

  // Description block
  s.addText('CDP-unified user profiles · Propensity-ranked product recommendations\nStory variant engine · Precision offer calibration · Campaign analytics', {
    x: 0.5, y: 2.5, w: 7.5, h: 0.6,
    fontSize: 13, bold: false, color: '7AB890', fontFace: 'Calibri',
  });

  // Bottom-right big stat box
  s.addShape(prs.ShapeType.rect, {
    x: 7.1, y: 4.3, w: 2.6, h: 1.9,
    fill: { color: '0A2B18' }, line: { color: C.bright, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('−32%', {
    x: 7.1, y: 4.45, w: 2.6, h: 0.7,
    fontSize: 42, bold: true, color: C.bright, align: 'center', fontFace: 'Calibri',
  });
  s.addText('marketing spend\nper conversion', {
    x: 7.1, y: 5.12, w: 2.6, h: 0.55,
    fontSize: 10.5, bold: false, color: 'A8D5B5', align: 'center', fontFace: 'Calibri',
  });

  // Presenter line
  s.addText('Presented by Ajay Avaghade  ·  Growth & MarTech PM', {
    x: 0.5, y: 6.9, w: 6.5, h: 0.25,
    fontSize: 9.5, bold: false, color: '5C7A69', fontFace: 'Calibri',
  });

  // Big background N watermark
  s.addText('N', {
    x: 6.2, y: 1.5, w: 4, h: 4.5,
    fontSize: 320, bold: true, color: C.bright,
    transparency: 94, fontFace: 'Calibri',
  });
}

// ─────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM (dark)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE PROBLEM', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 9, bold: true, color: C.bright,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('Navi Acquires Users. It Struggles to Keep Them, Cross-Sell to Them, or Spend Efficiently on Them.', {
    x: 0.5, y: 0.58, w: 9, h: 0.85,
    fontSize: 24, bold: true, color: C.white, fontFace: 'Calibri',
  });

  // 3 stat columns
  const cols = [
    { icon: '📊', stat: '1 of 3', label: 'Users transact on only one\nNavi product category', sub: 'UPI, Loans, Insurance — most users\nnever cross the category wall' },
    { icon: '💸', stat: '₹142 Cr', label: 'Wasted on spray-and-pray\noffer campaigns', sub: 'Static discounts given to high-intent\nusers who would convert anyway' },
    { icon: '📉', stat: '68%', label: 'Campaign budget spent\non wrong propensity band', sub: 'HIGH-intent users receive offers;\nLOW-intent users get same blast' },
  ];
  cols.forEach((col, i) => {
    const x = 0.5 + i * 3.1;
    s.addShape(prs.ShapeType.rect, {
      x, y: 1.65, w: 2.9, h: 3.3,
      fill: { color: '0A2B18' }, line: { color: C.bright, width: 0.8 },
      shadow: makeShadow(),
    });
    s.addText(col.icon, { x, y: 1.82, w: 2.9, h: 0.45, fontSize: 22, align: 'center' });
    s.addText(col.stat, {
      x, y: 2.28, w: 2.9, h: 0.7,
      fontSize: 38, bold: true, color: C.gold, align: 'center', fontFace: 'Calibri',
    });
    s.addText(col.label, {
      x: x + 0.12, y: 2.98, w: 2.65, h: 0.65,
      fontSize: 11.5, bold: true, color: C.white, align: 'center', fontFace: 'Calibri',
    });
    s.addText(col.sub, {
      x: x + 0.1, y: 3.65, w: 2.7, h: 0.7,
      fontSize: 9.5, bold: false, color: '7AB890', align: 'center', fontFace: 'Calibri',
    });
  });

  // Insight box
  s.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 5.2, w: 9.2, h: 0.9,
    fill: { color: '0E3520' }, line: { color: C.bright, width: 1 },
  });
  s.addText('Root cause: No unified user profile → no personalization → every campaign is a guess. Navi is burning budget on the wrong people, at the wrong time, with the wrong message — because there is no intelligence layer connecting data to action.', {
    x: 0.7, y: 5.3, w: 8.8, h: 0.7,
    fontSize: 10.5, bold: false, color: 'C8E6C9', fontFace: 'Calibri',
  });
}

// ─────────────────────────────────────────────
// SLIDE 3 — THE INSIGHT (light)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE INSIGHT', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 9, bold: true, color: C.ink,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('Spray-and-Pray Wastes Budget. Precision MarTech Converts the Right User at the Right Cost.', {
    x: 0.5, y: 0.58, w: 9, h: 0.82,
    fontSize: 23, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  // Left column — current state
  s.addShape(prs.ShapeType.rect, {
    x: 0.4, y: 1.6, w: 4.2, h: 3.8,
    fill: { color: 'FFF3F3' }, line: { color: 'E57373', width: 1.2 },
    shadow: makeShadow(),
  });
  s.addText('❌  Current Approach', {
    x: 0.55, y: 1.72, w: 4.0, h: 0.35,
    fontSize: 14, bold: true, color: 'C62828', fontFace: 'Calibri',
  });
  const currentPoints = [
    'One campaign blast to all users regardless of intent',
    'Same offer size for everyone — no calibration',
    'Generic copy — no story variant selection',
    'No suppression of non-convertible users',
    'No feedback loop — campaign learnings lost',
  ];
  currentPoints.forEach((pt, i) => {
    s.addText('• ' + pt, {
      x: 0.6, y: 2.2 + i * 0.55, w: 3.9, h: 0.48,
      fontSize: 12, color: '4A1A1A', fontFace: 'Calibri',
    });
  });

  // VS divider
  s.addShape(prs.ShapeType.ellipse, {
    x: 4.52, y: 2.85, w: 0.7, h: 0.7,
    fill: { color: C.ink }, line: { color: C.ink },
  });
  s.addText('VS', {
    x: 4.52, y: 2.85, w: 0.7, h: 0.7,
    fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle', fontFace: 'Calibri',
  });

  // Right column — proposed
  s.addShape(prs.ShapeType.rect, {
    x: 5.15, y: 1.6, w: 4.3, h: 3.8,
    fill: { color: 'EFFFEE' }, line: { color: C.bright, width: 1.5 },
    shadow: makeShadow(),
  });
  s.addText('✅  GrowthOS (Proposed)', {
    x: 5.3, y: 1.72, w: 4.0, h: 0.35,
    fontSize: 14, bold: true, color: '1B5E20', fontFace: 'Calibri',
  });
  const proposedPoints = [
    'CDP profile scores propensity per product per user',
    'Offer only to MED-band; suppress HIGH & LOW',
    'Minimum calibrated offer — not a flat discount',
    'Story variant engine picks Aspiration / Fear / Social Proof',
    'Analytics loop retrains models every campaign cycle',
  ];
  proposedPoints.forEach((pt, i) => {
    s.addText('• ' + pt, {
      x: 5.3, y: 2.2 + i * 0.55, w: 4.0, h: 0.48,
      fontSize: 12, color: '1A3A1A', fontFace: 'Calibri',
    });
  });

  // Bottom callout
  s.addShape(prs.ShapeType.rect, {
    x: 0.4, y: 5.58, w: 9.1, h: 0.82,
    fill: { color: C.dark }, line: { color: C.bright, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('Key insight: The cost of an offer is zero if the user was going to convert anyway. The cost of suppressing a non-convertible user is also zero. GrowthOS captures both savings — ₹106 Cr back to P&L per campaign cycle.', {
    x: 0.65, y: 5.66, w: 8.8, h: 0.65,
    fontSize: 11, bold: false, color: 'C8E6C9', fontFace: 'Calibri',
  });
}

// ─────────────────────────────────────────────
// SLIDE 4 — THE MECHANIC (dark)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.hero };

  s.addText('THE MECHANIC', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 9, bold: true, color: C.bright,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('5-Layer Intelligence Stack: From Raw Signal to Calibrated Conversion', {
    x: 0.5, y: 0.58, w: 9, h: 0.55,
    fontSize: 22, bold: true, color: C.white, fontFace: 'Calibri',
  });

  const steps = [
    { n: '01', title: 'CDP — Unified Profile', body: 'Ingest signals: transactions, app behavior, support tickets, product usage. Build a real-time scored profile per user with propensity bands (HIGH / MED / LOW) per product.', color: C.bright },
    { n: '02', title: 'Recommendation Engine', body: 'Rank Next Best Product (NBP) per user using propensity scores. Surface ranked recommendation in the right placement: home banner, post-transaction card, push, email.', color: C.bright },
    { n: '03', title: 'Story Variant Engine', body: 'Select narrative frame per user segment: Aspiration (upwardly mobile), Fear (risk-averse), or Social Proof (peer-influenced). One product, three CRM messages.', color: C.gold },
    { n: '04', title: 'Offer Calibration Engine', body: 'HIGH-band users (>0.7): no offer — they convert anyway. MED-band (0.4–0.7): minimum calibrated offer. LOW-band (<0.4): suppressed. Budget saved: ₹106 Cr per cycle.', color: C.gold },
    { n: '05', title: 'Analytics Feedback Loop', body: 'Track CVR, CAC, and CLTV uplift per variant per band. Winner retrains recommendation model. Every campaign improves the next one automatically.', color: C.orange },
  ];

  steps.forEach((st, i) => {
    const y = 1.35 + i * 1.0;
    // Step box
    s.addShape(prs.ShapeType.rect, {
      x: 0.38, y, w: 9.24, h: 0.88,
      fill: { color: '071A12' }, line: { color: st.color, width: 0.7 },
    });
    // Step number
    s.addText(st.n, {
      x: 0.48, y: y + 0.12, w: 0.55, h: 0.62,
      fontSize: 26, bold: true, color: st.color, fontFace: 'Calibri',
    });
    // Title
    s.addText(st.title, {
      x: 1.15, y: y + 0.1, w: 3.0, h: 0.35,
      fontSize: 13, bold: true, color: C.white, fontFace: 'Calibri',
    });
    // Body
    s.addText(st.body, {
      x: 1.15, y: y + 0.44, w: 8.3, h: 0.38,
      fontSize: 10, bold: false, color: '99CCA8', fontFace: 'Calibri',
    });
    // Arrow between steps
    if (i < 4) {
      s.addShape(prs.ShapeType.line, {
        x: 0.62, y: y + 0.92, w: 0, h: 0.08,
        line: { color: st.color, width: 1.5, dashType: 'dash' },
      });
    }
  });

  // Bottom note
  s.addText('PhonePe proof: Propensity-to-Transact ML (same architecture) cut ₹1,000+ Cr/yr marketing burn by 32% at 350M MAU', {
    x: 0.5, y: 6.85, w: 9.2, h: 0.25,
    fontSize: 9, bold: false, color: C.muted, fontFace: 'Calibri', italic: true,
  });
}

// ─────────────────────────────────────────────
// SLIDE 5 — THE PRODUCT (light)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE PRODUCT', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 9, bold: true, color: C.ink,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('4 Key Surfaces — Built & Validated at PhonePe', {
    x: 0.5, y: 0.58, w: 9, h: 0.48,
    fontSize: 22, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  const screens = [
    {
      num: '01',
      title: 'CDP Unified\nProfile',
      desc: 'Per-user propensity ring, 4 product bands scored in real-time, NBO chip with story variant and offer size auto-filled.',
      ascii: '┌──────────┐\n│ ◉ CLTV   │\n│ ₹18.4K→  │\n│ ₹47.2K   │\n│ UPI  91% │\n│ Ins  78% │\n│ NBO: Ins │\n└──────────┘',
    },
    {
      num: '02',
      title: 'CRM Builder\n& Story Variant',
      desc: 'Segment selector, NBP placement row, and 3 story variant cards (Aspiration / Fear / Social Proof) with live consumer preview.',
      ascii: '┌──────────┐\n│ Segment  │\n│ 1.8M usr │\n│ [ASPIRTN]│\n│ [fear   ]│\n│ [social ]│\n│ Preview→ │\n└──────────┘',
    },
    {
      num: '03',
      title: 'Offer Calibration\nEngine',
      desc: 'Propensity band split view: HIGH no-offer, MED minimum cashback, LOW suppressed. Live budget math: ₹142 Cr → ₹36 Cr.',
      ascii: '┌──────────┐\n│ HIGH >0.7│\n│ No offer │\n│ MED band │\n│ ₹200 off │\n│ LOW <0.4 │\n│ Suppress │\n└──────────┘',
    },
    {
      num: '04',
      title: 'Campaign\nAnalytics',
      desc: 'CVR by variant (Aspiration 31% > Fear 28% > Social 24%), CLTV cohort uplift, and budget efficiency waterfall.',
      ascii: '┌──────────┐\n│ CVR 28.4%│\n│ CAC ₹180 │\n│ Asp  31% │\n│ Fear 28% │\n│ Soc  24% │\n│ Saved ₹M │\n└──────────┘',
    },
  ];

  screens.forEach((sc, i) => {
    const x = 0.38 + i * 2.36;
    s.addShape(prs.ShapeType.rect, {
      x, y: 1.22, w: 2.18, h: 5.08,
      fill: { color: C.dark }, line: { color: C.bright, width: 1 },
      shadow: makeShadow(),
    });
    s.addText(sc.num, {
      x: x + 0.12, y: 1.32, w: 0.5, h: 0.3,
      fontSize: 10, bold: true, color: C.bright, fontFace: 'Calibri',
    });
    s.addText(sc.title, {
      x: x + 0.1, y: 1.6, w: 1.98, h: 0.55,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri',
    });
    s.addText(sc.ascii, {
      x: x + 0.08, y: 2.2, w: 2.02, h: 1.95,
      fontSize: 8, bold: false, color: C.bright, fontFace: 'Courier New',
    });
    s.addText(sc.desc, {
      x: x + 0.1, y: 4.2, w: 2.0, h: 1.85,
      fontSize: 9.5, bold: false, color: '99CCA8', fontFace: 'Calibri',
    });
  });
}

// ─────────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI (dark)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('IMPACT & ROI', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 9, bold: true, color: C.bright,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('Projected Impact — Built on PhonePe Proof Points at 350M MAU Scale', {
    x: 0.5, y: 0.58, w: 9, h: 0.52,
    fontSize: 22, bold: true, color: C.white, fontFace: 'Calibri',
  });

  // Left section — User & Product Impact
  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 1.28, w: 4.4, h: 0.38,
    fill: { color: '0A2B18' }, line: { color: C.bright, width: 0.5 },
  });
  s.addText('USER & PRODUCT IMPACT', {
    x: 0.55, y: 1.33, w: 4.1, h: 0.28,
    fontSize: 9, bold: true, color: C.bright, charSpacing: 1.5, fontFace: 'Calibri',
  });

  const leftMetrics = [
    { val: '+18%', lbl: 'Next-product adoption rate', sub: 'vs 12% baseline from undifferentiated campaigns' },
    { val: '+28%', lbl: 'Campaign CVR (GrowthOS cohort)', sub: 'vs 17% baseline with static offer blasts' },
    { val: '1.4×', lbl: 'CLTV uplift for converted users', sub: 'CDP-touched cohort vs untouched control' },
    { val: '−38%', lbl: 'Churn rate in converted cohort', sub: 'Users on 2+ products churn at 62% lower rate' },
  ];
  leftMetrics.forEach((m, i) => {
    const y = 1.82 + i * 0.9;
    s.addShape(prs.ShapeType.rect, {
      x: 0.38, y, w: 4.4, h: 0.82,
      fill: { color: '0A2B18' }, line: { color: '1A4A2A', width: 0.5 },
    });
    s.addText(m.val, {
      x: 0.5, y: y + 0.08, w: 1.3, h: 0.45,
      fontSize: 30, bold: true, color: C.bright, fontFace: 'Calibri',
    });
    s.addText(m.lbl, {
      x: 1.85, y: y + 0.08, w: 2.8, h: 0.28,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri',
    });
    s.addText(m.sub, {
      x: 1.85, y: y + 0.38, w: 2.8, h: 0.32,
      fontSize: 9.5, bold: false, color: '7AB890', fontFace: 'Calibri',
    });
  });

  // Right section — Company & P&L Impact
  s.addShape(prs.ShapeType.rect, {
    x: 5.1, y: 1.28, w: 4.5, h: 0.38,
    fill: { color: '2A1A00' }, line: { color: C.gold, width: 0.5 },
  });
  s.addText('COMPANY / P&L IMPACT', {
    x: 5.28, y: 1.33, w: 4.2, h: 0.28,
    fontSize: 9, bold: true, color: C.gold, charSpacing: 1.5, fontFace: 'Calibri',
  });

  const rightMetrics = [
    { val: '−32%', lbl: 'Marketing spend per conversion', sub: 'Propensity-to-Transact ML shipped at PhonePe at 350M MAU' },
    { val: '₹106 Cr', lbl: 'Budget saved per campaign cycle', sub: 'Offer calibration: ₹142 Cr flat → ₹36 Cr precision spend' },
    { val: '+₹8.2K', lbl: 'Average CLTV uplift per converted user', sub: '₹10.7K control → ₹26.6K GrowthOS cohort CLTV' },
    { val: '500+', lbl: 'Brand partners enableable via CRM', sub: 'Rebuilt rewards marketplace at PhonePe — 11% YoY revenue growth' },
  ];
  rightMetrics.forEach((m, i) => {
    const y = 1.82 + i * 0.9;
    s.addShape(prs.ShapeType.rect, {
      x: 5.1, y, w: 4.5, h: 0.82,
      fill: { color: '1A1000' }, line: { color: '3A2800', width: 0.5 },
    });
    s.addText(m.val, {
      x: 5.22, y: y + 0.08, w: 1.7, h: 0.45,
      fontSize: 26, bold: true, color: C.gold, fontFace: 'Calibri',
    });
    s.addText(m.lbl, {
      x: 6.95, y: y + 0.08, w: 2.55, h: 0.28,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri',
    });
    s.addText(m.sub, {
      x: 6.95, y: y + 0.38, w: 2.55, h: 0.32,
      fontSize: 9.5, bold: false, color: 'B8960A', fontFace: 'Calibri',
    });
  });

  // Insight
  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 5.62, w: 9.22, h: 0.72,
    fill: { color: '0E3520' }, line: { color: C.bright, width: 1 },
  });
  s.addText('PhonePe benchmark: Propensity-to-Transact ML (same architecture) reduced marketing burn by 32% and sustained GMV at 350M MAU — proving that precision MarTech at scale is not a theory, it is a shipped product.', {
    x: 0.58, y: 5.7, w: 9.0, h: 0.55,
    fontSize: 10.5, bold: false, color: 'C8E6C9', fontFace: 'Calibri',
  });
}

// ─────────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK (light)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('PROOF OF WORK', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 9, bold: true, color: C.ink,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('I Built This. At 350M MAU. Here is the Direct Mapping.', {
    x: 0.5, y: 0.58, w: 9, h: 0.5,
    fontSize: 22, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  // Left — PhonePe (dark)
  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 1.22, w: 4.6, h: 4.45,
    fill: { color: C.dark }, line: { color: C.bright, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('What I Built at PhonePe', {
    x: 0.55, y: 1.35, w: 4.2, h: 0.32,
    fontSize: 13, bold: true, color: C.bright, fontFace: 'Calibri',
  });

  const phonepeProof = [
    'Propensity-to-Transact ML: Real-time user-level scoring replacing ₹1,000+ Cr/yr of static targeting → 32% marketing burn reduction',
    'Dynamic Cart Incentivisation Engine: Context-aware triggers (cart value × intent × time) → 60% cart abandonment reduction at Pincode',
    'ML-driven Rewards Marketplace: Audience segmentation + brand self-serve CRM tooling → 500+ partners, 26% engagement uplift',
    'Offer Redemption Redesign: End-to-end eligibility logic + real-time surfacing → 22% checkout conversion lift at 350M MAU',
    'B2B Merchant Onboarding Platform: KYC/KYB automation + parallel review → 5,000+ merchants, 23% lower CAC',
  ];
  phonepeProof.forEach((pt, i) => {
    s.addText('▸  ' + pt, {
      x: 0.52, y: 1.82 + i * 0.74, w: 4.28, h: 0.65,
      fontSize: 10.5, bold: false, color: 'C8E6C9', fontFace: 'Calibri',
    });
  });

  // Right — Navi JD mapping (light)
  s.addShape(prs.ShapeType.rect, {
    x: 5.22, y: 1.22, w: 4.4, h: 4.45,
    fill: { color: 'EFFFEE' }, line: { color: C.bright, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('How It Maps to This Role', {
    x: 5.38, y: 1.35, w: 4.1, h: 0.32,
    fontSize: 13, bold: true, color: '1B5E20', fontFace: 'Calibri',
  });

  const naviMapping = [
    '→  GrowthOS CDP propensity scoring (same ML architecture, real-time, session-aware)',
    '→  GrowthOS Offer Calibration Engine (same context-aware trigger, minimum effective offer)',
    '→  GrowthOS CRM Builder + Story Variant Engine (same audience segmentation + self-serve tooling)',
    '→  GrowthOS recommendation surfacing (same eligibility logic applied to product recommendation)',
    '→  Navi B2B partner / brand onboarding (same KYB flow, same CAC reduction lever)',
  ];
  naviMapping.forEach((pt, i) => {
    s.addText(pt, {
      x: 5.38, y: 1.82 + i * 0.74, w: 4.1, h: 0.65,
      fontSize: 10.5, bold: false, color: '1A3A1A', fontFace: 'Calibri',
    });
  });

  // Closing quote
  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 5.82, w: 9.22, h: 0.62,
    fill: { color: C.dark }, line: { color: C.bright, width: 1 },
  });
  s.addText('Every mechanic in GrowthOS is a direct translation of something I shipped, measured, and iterated at PhonePe — at 10× the user scale of Navi today.', {
    x: 0.6, y: 5.9, w: 8.8, h: 0.5,
    fontSize: 11, bold: false, color: 'C8E6C9', fontFace: 'Calibri', italic: true,
  });
}

// ─────────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN (dark)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('ROLLOUT PLAN', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 9, bold: true, color: C.bright,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('Phased Build: CDP Foundation → Intelligence Engine → Full CLTV Loop', {
    x: 0.5, y: 0.58, w: 9, h: 0.52,
    fontSize: 22, bold: true, color: C.white, fontFace: 'Calibri',
  });

  const phases = [
    {
      phase: 'Phase 1',
      timeline: 'Month 1–2',
      title: 'CDP Foundation & Profile Scoring',
      points: [
        'Ingest: transactions, app events, support, product usage',
        'Build propensity scoring pipeline (UPI / Loans / Insurance)',
        'Tag users into HIGH / MED / LOW bands per product',
        'Success metric: 80% of MAU profiled and banded within 60 days',
      ],
      color: C.bright,
    },
    {
      phase: 'Phase 2',
      timeline: 'Month 3–4',
      title: 'Recommendation + Story Variant Engine',
      points: [
        'Launch NBP recommendation on Home + post-transaction surface',
        'A/B test 3 story variants (Aspiration / Fear / Social Proof)',
        'Integrate CRM builder for ops team self-serve campaigns',
        'Success metric: +15% next-product CTR vs baseline campaign',
      ],
      color: C.gold,
    },
    {
      phase: 'Phase 3',
      timeline: 'Month 5–6',
      title: 'Full Calibration + Analytics Feedback Loop',
      points: [
        'Deploy Offer Calibration Engine — suppress HIGH, calibrate MED',
        'Launch campaign analytics dashboard with variant attribution',
        'Retrain propensity model on first 2 campaign cycles of data',
        'Success metric: ₹106 Cr budget saving confirmed vs flat-offer baseline',
      ],
      color: C.orange,
    },
  ];

  phases.forEach((ph, i) => {
    const x = 0.38 + i * 3.14;
    s.addShape(prs.ShapeType.rect, {
      x, y: 1.28, w: 2.98, h: 4.42,
      fill: { color: '071A12' }, line: { color: ph.color, width: 1.2 },
      shadow: makeShadow(),
    });
    // Phase tag
    s.addShape(prs.ShapeType.rect, {
      x: x + 0.1, y: 1.38, w: 1.1, h: 0.28,
      fill: { color: ph.color }, line: { color: ph.color },
    });
    s.addText(ph.phase, {
      x: x + 0.1, y: 1.38, w: 1.1, h: 0.28,
      fontSize: 9, bold: true, color: C.dark, align: 'center', fontFace: 'Calibri',
    });
    s.addText(ph.timeline, {
      x: x + 1.3, y: 1.4, w: 1.55, h: 0.24,
      fontSize: 9, bold: false, color: ph.color, fontFace: 'Calibri',
    });
    s.addText(ph.title, {
      x: x + 0.12, y: 1.75, w: 2.72, h: 0.55,
      fontSize: 12.5, bold: true, color: C.white, fontFace: 'Calibri',
    });
    ph.points.forEach((pt, j) => {
      s.addText('• ' + pt, {
        x: x + 0.14, y: 2.38 + j * 0.72, w: 2.7, h: 0.66,
        fontSize: 10, bold: false, color: '99CCA8', fontFace: 'Calibri',
      });
    });
  });

  // The ask box
  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 5.85, w: 9.22, h: 0.55,
    fill: { color: '0A2B18' }, line: { color: C.bright, width: 1 },
  });
  s.addText('What I need:  Access to event/transaction data  ·  1 Data Engineer (CDP ingestion)  ·  1 ML Engineer (propensity model)  ·  1 Designer (CRM UI)', {
    x: 0.6, y: 5.93, w: 8.8, h: 0.4,
    fontSize: 10, bold: false, color: 'C8E6C9', fontFace: 'Calibri',
  });

  // Footer
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade', {
    x: 0.38, y: 6.9, w: 9.22, h: 0.22,
    fontSize: 8.5, bold: false, color: C.muted, fontFace: 'Calibri',
  });
}

// ─────────────────────────────────────────────
// WRITE
// ─────────────────────────────────────────────
prs.writeFile({ fileName: 'assets/navi-growthos-deck.pptx' })
  .then(() => console.log('✅  navi-growthos-deck.pptx written'))
  .catch(e => { console.error(e); process.exit(1); });
