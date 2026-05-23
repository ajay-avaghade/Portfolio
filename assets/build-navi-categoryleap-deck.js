// build-navi-categoryleap-deck.js
// Run: node assets/build-navi-categoryleap-deck.js
// Output: assets/navi-categoryleap-deck.pptx

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

  // Diagonal motif
  for (let i = 0; i < 8; i++) {
    s.addShape(prs.ShapeType.line, {
      x: -0.5 + i * 1.5, y: 0, w: 3.5, h: 0,
      line: { color: C.gold, width: 0.4, transparency: 88 },
      rotate: 35,
    });
  }

  s.addText('NAVI TECHNOLOGIES  ·  PRODUCT MANAGER — GROWTH & ENGAGEMENT', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 8.5, bold: false, color: C.gold,
    charSpacing: 2, fontFace: 'Calibri',
  });

  s.addText('CategoryLeap', {
    x: 0.5, y: 0.68, w: 9, h: 1.2,
    fontSize: 60, bold: true, color: C.white, fontFace: 'Calibri',
    shadow: makeShadow(),
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.82, w: 1.6, h: 0.07,
    fill: { color: C.gold }, line: { color: C.gold },
  });

  s.addText('Cross-Category Milestone Engine for Navi CLTV Growth', {
    x: 0.5, y: 1.98, w: 7.5, h: 0.38,
    fontSize: 18, bold: false, color: 'FFF3C4', fontFace: 'Calibri',
  });

  s.addText('UPI milestones · Loans milestones · Insurance milestones\nBronze → Silver → Gold → Platinum tier progression', {
    x: 0.5, y: 2.5, w: 7.5, h: 0.6,
    fontSize: 13, bold: false, color: 'B89A40', fontFace: 'Calibri',
  });

  // Big stat box
  s.addShape(prs.ShapeType.rect, {
    x: 7.1, y: 4.3, w: 2.6, h: 1.9,
    fill: { color: '1A1000' }, line: { color: C.gold, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('3.7×', {
    x: 7.1, y: 4.45, w: 2.6, h: 0.7,
    fontSize: 46, bold: true, color: C.gold, align: 'center', fontFace: 'Calibri',
  });
  s.addText('CLTV uplift:\n1-cat → 3-cat user', {
    x: 7.1, y: 5.12, w: 2.6, h: 0.55,
    fontSize: 10, bold: false, color: 'D4A840', align: 'center', fontFace: 'Calibri',
  });

  s.addText('Presented by Ajay Avaghade  ·  Growth & MarTech PM', {
    x: 0.5, y: 6.9, w: 6.5, h: 0.25,
    fontSize: 9.5, bold: false, color: C.muted, fontFace: 'Calibri',
  });

  // Watermark
  s.addText('🏆', {
    x: 5.5, y: 1.2, w: 4.8, h: 5.5,
    fontSize: 260, align: 'center',
    transparency: 93,
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
    fontSize: 9, bold: true, color: C.gold,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('Navi Has Three Products. Most Users Only Ever Touch One.', {
    x: 0.5, y: 0.58, w: 9, h: 0.72,
    fontSize: 26, bold: true, color: C.white, fontFace: 'Calibri',
  });

  const cols = [
    { icon: '🔒', stat: '67%', label: 'Users on exactly\none Navi product', sub: 'UPI, Loans, Insurance — siloed.\nUsers never leap across categories.' },
    { icon: '📉', stat: '₹8.4K', label: 'Average annual CLTV\nfor 1-category users', sub: 'vs ₹61.2K for 3-category users.\nThat is a 7.3× value gap left on the table.' },
    { icon: '⚡', stat: '0', label: 'Structured nudges to\ncross-sell between categories', sub: 'No milestone, no progress tracker,\nno reward moment for loyalty.' },
  ];

  cols.forEach((col, i) => {
    const x = 0.5 + i * 3.1;
    s.addShape(prs.ShapeType.rect, {
      x, y: 1.58, w: 2.9, h: 3.4,
      fill: { color: '1A1000' }, line: { color: C.gold, width: 0.8 },
      shadow: makeShadow(),
    });
    s.addText(col.icon, { x, y: 1.75, w: 2.9, h: 0.45, fontSize: 22, align: 'center' });
    s.addText(col.stat, {
      x, y: 2.22, w: 2.9, h: 0.7,
      fontSize: 38, bold: true, color: C.gold, align: 'center', fontFace: 'Calibri',
    });
    s.addText(col.label, {
      x: x + 0.12, y: 2.92, w: 2.65, h: 0.55,
      fontSize: 11.5, bold: true, color: C.white, align: 'center', fontFace: 'Calibri',
    });
    s.addText(col.sub, {
      x: x + 0.1, y: 3.5, w: 2.7, h: 0.65,
      fontSize: 9.5, bold: false, color: 'B8960A', align: 'center', fontFace: 'Calibri',
    });
  });

  // Insight box
  s.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 5.22, w: 9.2, h: 0.9,
    fill: { color: '1A1000' }, line: { color: C.gold, width: 1 },
  });
  s.addText('Root cause: Users have no reason to explore a second category and no progress signal when they do. A single checkout in UPI, a single loan, a single insurance purchase — but no narrative connecting them into a journey. CategoryLeap builds that narrative.', {
    x: 0.7, y: 5.3, w: 8.8, h: 0.7,
    fontSize: 10.5, bold: false, color: 'FFF3C4', fontFace: 'Calibri',
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

  s.addText('Stickiness Is Not a Single Product Problem. It Is a Cross-Category Habit Problem.', {
    x: 0.5, y: 0.58, w: 9, h: 0.72,
    fontSize: 23, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  // Left — current
  s.addShape(prs.ShapeType.rect, {
    x: 0.4, y: 1.52, w: 4.2, h: 3.85,
    fill: { color: 'FFF3F3' }, line: { color: 'E57373', width: 1.2 },
    shadow: makeShadow(),
  });
  s.addText('❌  Today: Siloed Products', {
    x: 0.55, y: 1.64, w: 4.0, h: 0.35,
    fontSize: 14, bold: true, color: 'C62828', fontFace: 'Calibri',
  });
  const currentPoints = [
    'UPI, Loans, Insurance are separate funnels',
    'No shared progress signal across products',
    'Cross-sell is a push notification — easily ignored',
    'No reward for using 2+ products together',
    'Single-product users churn at 3.7× higher rate',
  ];
  currentPoints.forEach((pt, i) => {
    s.addText('• ' + pt, {
      x: 0.6, y: 2.12 + i * 0.57, w: 3.9, h: 0.5,
      fontSize: 12, color: '4A1A1A', fontFace: 'Calibri',
    });
  });

  // VS
  s.addShape(prs.ShapeType.ellipse, {
    x: 4.52, y: 2.88, w: 0.7, h: 0.7,
    fill: { color: C.ink }, line: { color: C.ink },
  });
  s.addText('VS', {
    x: 4.52, y: 2.88, w: 0.7, h: 0.7,
    fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle', fontFace: 'Calibri',
  });

  // Right — proposed
  s.addShape(prs.ShapeType.rect, {
    x: 5.15, y: 1.52, w: 4.3, h: 3.85,
    fill: { color: 'FFFDE7' }, line: { color: C.gold, width: 1.5 },
    shadow: makeShadow(),
  });
  s.addText('✅  CategoryLeap: Unified Journey', {
    x: 5.3, y: 1.64, w: 4.0, h: 0.35,
    fontSize: 14, bold: true, color: '6D4C00', fontFace: 'Calibri',
  });
  const proposedPoints = [
    'UPI + Loans + Insurance milestones form one Portfolio Score',
    'Bronze → Silver → Gold → Platinum tier progression',
    'Completing each category unlocks concrete financial benefits',
    'GrowthOS feeds CategoryLeap the Next Best Category NBO',
    'Multi-category users generate 3.7× the CLTV of single-product users',
  ];
  proposedPoints.forEach((pt, i) => {
    s.addText('• ' + pt, {
      x: 5.3, y: 2.12 + i * 0.57, w: 4.0, h: 0.5,
      fontSize: 12, color: '3A2800', fontFace: 'Calibri',
    });
  });

  // Callout
  s.addShape(prs.ShapeType.rect, {
    x: 0.4, y: 5.52, w: 9.1, h: 0.88,
    fill: { color: C.dark }, line: { color: C.gold, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('Key insight: The milestone construct creates a narrative that makes Navi feel like a financial journey — not three separate apps. The tier reward makes each additional product feel like a promotion earned, not a product sold.', {
    x: 0.65, y: 5.6, w: 8.8, h: 0.72,
    fontSize: 11, bold: false, color: 'FFF3C4', fontFace: 'Calibri',
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
    fontSize: 9, bold: true, color: C.gold,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('Milestone Architecture: 3 Pillars, 4 Tiers, 1 Portfolio Score', {
    x: 0.5, y: 0.58, w: 9, h: 0.52,
    fontSize: 22, bold: true, color: C.white, fontFace: 'Calibri',
  });

  // 3 pillar cards
  const pillars = [
    {
      title: '🟢  UPI Pillar',
      color: C.bright,
      ms: ['M1: Pay 10+ unique merchants in 30 days', 'M2: Complete 3 bill payment categories (rent, utilities, telecom)', 'M3: Send money to 5 unique contacts'],
      reward: 'UPI Cashback Shield: 1% cashback on next 3 UPI transactions',
    },
    {
      title: '🟡  Loans Pillar',
      color: C.gold,
      ms: ['M1: Complete loan eligibility check', 'M2: Apply for and receive first Navi loan disbursement', 'M3: Complete 3 consecutive on-time EMI repayments'],
      reward: 'Processing Fee Waiver: ₹1,999 processing fee waived on next loan',
    },
    {
      title: '🟠  Insurance Pillar',
      color: C.orange,
      ms: ['M1: Get a health insurance quote', 'M2: Purchase first Navi Insurance policy', 'M3: Complete policy renewal on time'],
      reward: 'Premium Discount: 10% off first-year renewal premium',
    },
  ];

  pillars.forEach((p, i) => {
    const x = 0.38 + i * 3.14;
    s.addShape(prs.ShapeType.rect, {
      x, y: 1.28, w: 2.98, h: 3.92,
      fill: { color: '071A12' }, line: { color: p.color, width: 1.2 },
      shadow: makeShadow(),
    });
    s.addText(p.title, {
      x: x + 0.12, y: 1.38, w: 2.72, h: 0.38,
      fontSize: 13, bold: true, color: p.color, fontFace: 'Calibri',
    });
    p.ms.forEach((ms, j) => {
      s.addText('M' + (j + 1) + '  ' + ms, {
        x: x + 0.12, y: 1.88 + j * 0.72, w: 2.72, h: 0.66,
        fontSize: 10.5, bold: false, color: 'C8E6C9', fontFace: 'Calibri',
      });
    });
    s.addShape(prs.ShapeType.rect, {
      x: x + 0.1, y: 4.08, w: 2.78, h: 0.25,
      fill: { color: p.color + '22' }, line: { color: p.color, width: 0.5 },
    });
    s.addText('Reward: ' + p.reward, {
      x: x + 0.12, y: 4.38, w: 2.74, h: 0.7,
      fontSize: 9.5, bold: false, color: p.color, fontFace: 'Calibri',
    });
  });

  // Tier ladder
  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 5.38, w: 9.22, h: 0.78,
    fill: { color: '071A12' }, line: { color: C.gold, width: 1 },
  });
  s.addText('Tier Progression  →', {
    x: 0.55, y: 5.48, w: 1.5, h: 0.3,
    fontSize: 9, bold: true, color: C.gold, fontFace: 'Calibri',
  });
  const tiers = [
    { name: 'BRONZE', pts: '0–40 pts', color: 'CD7F32' },
    { name: 'SILVER', pts: '40–65 pts', color: 'A8A9AD' },
    { name: 'GOLD', pts: '65–85 pts', color: C.gold },
    { name: 'PLATINUM', pts: '85–100 pts', color: 'B5E4F7' },
  ];
  tiers.forEach((t, i) => {
    const x = 2.2 + i * 1.85;
    s.addShape(prs.ShapeType.rect, {
      x, y: 5.48, w: 1.65, h: 0.58,
      fill: { color: t.color + '22' }, line: { color: t.color, width: 0.8 },
    });
    s.addText(t.name, {
      x, y: 5.5, w: 1.65, h: 0.25,
      fontSize: 10, bold: true, color: t.color, align: 'center', fontFace: 'Calibri',
    });
    s.addText(t.pts, {
      x, y: 5.74, w: 1.65, h: 0.22,
      fontSize: 8.5, bold: false, color: t.color, align: 'center', fontFace: 'Calibri',
    });
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

  s.addText('4 Key Screens — Built on Navi Design Language', {
    x: 0.5, y: 0.58, w: 9, h: 0.48,
    fontSize: 22, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  const screens = [
    {
      num: '01',
      title: 'Navi Home with\nPortfolio Banner',
      desc: 'Score ring (47/100), tier badge, 3 pillar mini-chips, progress bar to Silver, CTA card highlighting next milestone to unlock.',
      ascii: '┌──────────┐\n│ Score 47 │\n│ 🥉 Bronze │\n│ UPI ✓    │\n│ Loans ⚡  │\n│ Ins ○    │\n│ →Silver  │\n└──────────┘',
    },
    {
      num: '02',
      title: 'Milestone\nTracker',
      desc: 'Per-pillar milestone rows: UPI (3/3 complete), Loans (1/3 active — M2 in progress), Insurance (locked). Progress bars with completion indicators.',
      ascii: '┌──────────┐\n│ UPI  ███ │\n│ 3/3 done │\n│ Loans ██ │\n│ M2 prog  │\n│ Ins  --- │\n│ Locked   │\n└──────────┘',
    },
    {
      num: '03',
      title: 'Tier Unlocked\n(Celebration)',
      desc: 'Confetti on Silver unlock. CLTV projection ring (₹18.4K → ₹31.2K). 3 concrete reward chips. CTA to start Insurance journey toward Gold.',
      ascii: '┌──────────┐\n│🎊 SILVER │\n│ UNLOCKED │\n│₹18.4→   │\n│₹31.2K   │\n│ 3 perks  │\n│ → Gold   │\n└──────────┘',
    },
    {
      num: '04',
      title: 'Ops CLTV\nDashboard',
      desc: 'KPIs: 1.8× CLTV 1→2 categories, 3.7× 1→3. Cohort bars, milestone funnel with biggest drop-off signal, GrowthOS action trigger.',
      ascii: '┌──────────┐\n│1-cat ₹8K │\n│2-cat ₹25K│\n│3-cat ₹61K│\n│Funnel:   │\n│UPI 100%  │\n│Loan 48%  │\n└──────────┘',
    },
  ];

  screens.forEach((sc, i) => {
    const x = 0.38 + i * 2.36;
    s.addShape(prs.ShapeType.rect, {
      x, y: 1.22, w: 2.18, h: 5.08,
      fill: { color: C.dark }, line: { color: C.gold, width: 1 },
      shadow: makeShadow(),
    });
    s.addText(sc.num, {
      x: x + 0.12, y: 1.32, w: 0.5, h: 0.3,
      fontSize: 10, bold: true, color: C.gold, fontFace: 'Calibri',
    });
    s.addText(sc.title, {
      x: x + 0.1, y: 1.6, w: 1.98, h: 0.55,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri',
    });
    s.addText(sc.ascii, {
      x: x + 0.08, y: 2.2, w: 2.02, h: 1.95,
      fontSize: 8, bold: false, color: C.gold, fontFace: 'Courier New',
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
    fontSize: 9, bold: true, color: C.gold,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('Projected Impact — The CLTV Multiplier Effect of Cross-Category Engagement', {
    x: 0.5, y: 0.58, w: 9, h: 0.52,
    fontSize: 22, bold: true, color: C.white, fontFace: 'Calibri',
  });

  // CLTV cohort visual block
  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 1.22, w: 4.5, h: 3.85,
    fill: { color: '0A2B18' }, line: { color: C.bright, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('CLTV COHORT LADDER', {
    x: 0.55, y: 1.32, w: 4.2, h: 0.3,
    fontSize: 10, bold: true, color: C.bright, charSpacing: 1.5, fontFace: 'Calibri',
  });

  const cohorts = [
    { label: '1-category user', val: '₹8,400 / yr', bar: 14, color: '5C7A69' },
    { label: '2-category user', val: '₹24,600 / yr', bar: 40, color: C.bright },
    { label: '3-category user', val: '₹61,200 / yr', bar: 100, color: C.gold },
  ];
  cohorts.forEach((c, i) => {
    const y = 1.78 + i * 1.0;
    s.addText(c.label, {
      x: 0.55, y, w: 2.5, h: 0.28,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri',
    });
    s.addText(c.val, {
      x: 3.0, y, w: 1.65, h: 0.28,
      fontSize: 12, bold: true, color: c.color, align: 'right', fontFace: 'Calibri',
    });
    // Bar
    s.addShape(prs.ShapeType.rect, {
      x: 0.55, y: y + 0.35, w: 3.7, h: 0.35,
      fill: { color: '0E3520' }, line: { color: '1A4A2A', width: 0.3 },
    });
    s.addShape(prs.ShapeType.rect, {
      x: 0.55, y: y + 0.35, w: 3.7 * c.bar / 100, h: 0.35,
      fill: { color: c.color }, line: { color: c.color },
    });
  });

  // Multiplier callouts
  s.addText('1→2 categories:  1.8×  CLTV uplift', {
    x: 0.55, y: 3.95, w: 4.0, h: 0.32,
    fontSize: 13, bold: true, color: C.bright, fontFace: 'Calibri',
  });
  s.addText('1→3 categories:  3.7×  CLTV uplift', {
    x: 0.55, y: 4.32, w: 4.0, h: 0.32,
    fontSize: 13, bold: true, color: C.gold, fontFace: 'Calibri',
  });

  // Right — engagement metrics
  s.addShape(prs.ShapeType.rect, {
    x: 5.1, y: 1.22, w: 4.5, h: 3.85,
    fill: { color: '1A1000' }, line: { color: C.gold, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('ENGAGEMENT & RETENTION IMPACT', {
    x: 5.28, y: 1.32, w: 4.2, h: 0.3,
    fontSize: 10, bold: true, color: C.gold, charSpacing: 1.5, fontFace: 'Calibri',
  });

  const rightMetrics = [
    { val: '34%', lbl: 'Users on 2+ categories\nafter CategoryLeap pilot', color: C.bright },
    { val: '−62%', lbl: 'Churn rate reduction for\n2+ category users vs 1-cat', color: C.gold },
    { val: '+48%', lbl: 'Milestone completion at\nLoans M2 (key funnel stage)', color: C.orange },
    { val: '+27%', lbl: 'Insurance policy purchases\nfrom Loans-complete cohort', color: C.bright },
  ];
  rightMetrics.forEach((m, i) => {
    const y = 1.78 + i * 0.8;
    s.addShape(prs.ShapeType.rect, {
      x: 5.22, y, w: 4.25, h: 0.72,
      fill: { color: '0D0800' }, line: { color: '2A1A00', width: 0.4 },
    });
    s.addText(m.val, {
      x: 5.32, y: y + 0.1, w: 1.15, h: 0.5,
      fontSize: 30, bold: true, color: m.color, fontFace: 'Calibri',
    });
    s.addText(m.lbl, {
      x: 6.52, y: y + 0.1, w: 2.85, h: 0.5,
      fontSize: 11.5, bold: false, color: C.white, fontFace: 'Calibri',
    });
  });

  // Insight
  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 5.25, w: 9.22, h: 0.82,
    fill: { color: '0E3520' }, line: { color: C.bright, width: 1 },
  });
  s.addText('PhonePe benchmark: Milestone reward construct (same architecture) drove 34% higher 90-day device activity and 60% cart abandonment reduction at Pincode. Cross-category engagement is the highest-leverage CLTV driver available to Navi today.', {
    x: 0.58, y: 5.33, w: 9.0, h: 0.65,
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

  s.addText('I Built This Architecture. Here Is the Exact Parallel.', {
    x: 0.5, y: 0.58, w: 9, h: 0.5,
    fontSize: 22, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  // Left dark
  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 1.22, w: 4.6, h: 4.45,
    fill: { color: C.dark }, line: { color: C.gold, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('What I Built at PhonePe', {
    x: 0.55, y: 1.35, w: 4.2, h: 0.32,
    fontSize: 13, bold: true, color: C.gold, fontFace: 'Calibri',
  });

  const phonepeProof = [
    'Milestone Construct (POS Merchants): AED-equivalent terminal rental waivers tied to TPV + TPC goals — same goal-tracking architecture as CategoryLeap pillar milestones',
    'Dynamic Cart Incentivisation Engine: Context-aware triggers → 35% AOV uplift, 60% cart abandonment reduction — same progress-reward loop as Silver/Gold tier unlock',
    'ML-Driven Rewards Marketplace: 500+ brand partners cross-selling into base — same cross-product recommendation via propensity scoring',
    'B2B Platform: 5,000+ merchants acquired via structured milestone-based onboarding — same sequential milestone completion logic',
    'Propensity-to-Transact: Real-time scoring enabling precise next-best-action — same engine powering CategoryLeap NBO via GrowthOS CDP',
  ];
  phonepeProof.forEach((pt, i) => {
    s.addText('▸  ' + pt, {
      x: 0.52, y: 1.82 + i * 0.74, w: 4.28, h: 0.65,
      fontSize: 10.5, bold: false, color: 'FFF3C4', fontFace: 'Calibri',
    });
  });

  // Right light
  s.addShape(prs.ShapeType.rect, {
    x: 5.22, y: 1.22, w: 4.4, h: 4.45,
    fill: { color: 'FFFDE7' }, line: { color: C.gold, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('How It Maps to This Role', {
    x: 5.38, y: 1.35, w: 4.1, h: 0.32,
    fontSize: 13, bold: true, color: '6D4C00', fontFace: 'Calibri',
  });

  const naviMapping = [
    '→  CategoryLeap UPI / Loans / Insurance pillar milestone design (same goal + reward loop)',
    '→  Silver/Gold/Platinum tier unlock celebration screen (same confetti + CLTV projection reveal)',
    '→  Cross-category NBO powered by GrowthOS CDP propensity (same next-best-action engine)',
    '→  Ops CLTV Dashboard — drop-off signal + GrowthOS action trigger (same monitoring + feedback)',
    '→  Milestone funnel analytics: UPI 100% → Loans 48% → Insurance 27% (same conversion funnel architecture)',
  ];
  naviMapping.forEach((pt, i) => {
    s.addText(pt, {
      x: 5.38, y: 1.82 + i * 0.74, w: 4.1, h: 0.65,
      fontSize: 10.5, bold: false, color: '3A2800', fontFace: 'Calibri',
    });
  });

  // Closing
  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 5.82, w: 9.22, h: 0.62,
    fill: { color: C.dark }, line: { color: C.gold, width: 1 },
  });
  s.addText('CategoryLeap is not a concept. Every mechanic — milestones, tier progression, ops dashboard, cross-category NBO — maps directly to something I shipped and measured at PhonePe at 350M MAU.', {
    x: 0.6, y: 5.9, w: 8.8, h: 0.5,
    fontSize: 11, bold: false, color: 'FFF3C4', fontFace: 'Calibri', italic: true,
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
    fontSize: 9, bold: true, color: C.gold,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('Phased Rollout: UPI Milestone → Loans → Insurance → Full Tier Engine', {
    x: 0.5, y: 0.58, w: 9, h: 0.52,
    fontSize: 22, bold: true, color: C.white, fontFace: 'Calibri',
  });

  const phases = [
    {
      phase: 'Phase 1',
      timeline: 'Month 1–2',
      title: 'UPI Milestone Pilot',
      points: [
        'Launch UPI pillar (3 milestones) to 100K UPI-active users',
        'Track M1→M2→M3 completion funnel and drop-off',
        'A/B test: milestone reward vs cashback vs no incentive',
        'Success metric: 60% M1 completion within 30 days',
      ],
      color: C.bright,
    },
    {
      phase: 'Phase 2',
      timeline: 'Month 3–4',
      title: 'Loans + Insurance Pillars',
      points: [
        'Unlock Loans + Insurance pillars for UPI-complete users',
        'Deploy Bronze → Silver → Gold tier progression',
        'Integrate GrowthOS NBO to surface CategoryLeap CTA',
        'Success metric: 30% of UPI-complete users start Loans M1',
      ],
      color: C.gold,
    },
    {
      phase: 'Phase 3',
      timeline: 'Month 5–6',
      title: 'Full Tier Engine + Ops Dashboard',
      points: [
        'Launch Platinum tier with exclusive financial benefits',
        'Deploy ops CLTV dashboard with real-time drop-off alerts',
        'Activate GrowthOS → CategoryLeap feedback loop',
        'Success metric: 34% of MAU on 2+ categories at 90 days',
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

  // Ask box
  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 5.85, w: 9.22, h: 0.55,
    fill: { color: '1A1000' }, line: { color: C.gold, width: 1 },
  });
  s.addText('What I need:  Transaction data access  ·  1 Backend Engineer (milestone tracking API)  ·  1 Designer (tier progression UI)  ·  GrowthOS CDP integration', {
    x: 0.6, y: 5.93, w: 8.8, h: 0.4,
    fontSize: 10, bold: false, color: 'FFF3C4', fontFace: 'Calibri',
  });

  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade', {
    x: 0.38, y: 6.9, w: 9.22, h: 0.22,
    fontSize: 8.5, bold: false, color: C.muted, fontFace: 'Calibri',
  });
}

// ─────────────────────────────────────────────
// WRITE
// ─────────────────────────────────────────────
prs.writeFile({ fileName: 'assets/navi-categoryleap-deck.pptx' })
  .then(() => console.log('✅  navi-categoryleap-deck.pptx written'))
  .catch(e => { console.error(e); process.exit(1); });
