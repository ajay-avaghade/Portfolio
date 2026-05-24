// build-revolut-growthiq-deck.js
// Run: node assets/build-revolut-growthiq-deck.js
// Output: assets/revolut-growthiq-deck.pptx

const pptx = require('pptxgenjs');
const prs = new pptx();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0D0F14',
  hero:   '141822',
  bright: '0668E1',
  purple: '7C3AED',
  gold:   'F59E0B',
  green:  '10B981',
  orange: 'F97316',
  white:  'FFFFFF',
  lgray:  'F0F2F7',
  muted:  '64748b',
  ink:    '0f172a',
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
      line: { color: C.bright, width: 0.4, transparency: 88 },
      rotate: 35,
    });
  }

  s.addText('REVOLUT  ·  JUNIOR PRODUCT MARKETING MANAGER — LOYALTY', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 8.5, bold: false, color: C.bright,
    charSpacing: 2, fontFace: 'Calibri',
  });

  s.addText('GrowthIQ', {
    x: 0.5, y: 0.65, w: 9, h: 1.1,
    fontSize: 66, bold: true, color: C.white, fontFace: 'Calibri',
    shadow: makeShadow(),
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.72, w: 1.6, h: 0.07,
    fill: { color: C.bright }, line: { color: C.bright },
  });

  s.addText('The Upgrade Intelligence Layer Revolut is Missing', {
    x: 0.5, y: 1.88, w: 7.5, h: 0.38,
    fontSize: 18, bold: false, color: 'C8D4F0', fontFace: 'Calibri',
  });

  s.addText('CDP propensity scoring · Story variant engine · Offer calibration · Campaign analytics', {
    x: 0.5, y: 2.38, w: 7.5, h: 0.4,
    fontSize: 12.5, bold: false, color: '4A6FA5', fontFace: 'Calibri',
  });

  // Big stat box
  s.addShape(prs.ShapeType.rect, {
    x: 7.1, y: 4.2, w: 2.6, h: 2.0,
    fill: { color: '0A1428' }, line: { color: C.bright, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('+22%', {
    x: 7.1, y: 4.35, w: 2.6, h: 0.72,
    fontSize: 42, bold: true, color: C.bright, align: 'center', fontFace: 'Calibri',
  });
  s.addText('plan upgrade CVR\nvs blanket campaign', {
    x: 7.1, y: 5.05, w: 2.6, h: 0.55,
    fontSize: 10.5, bold: false, color: '7CA8D8', align: 'center', fontFace: 'Calibri',
  });

  s.addText('Presented by Ajay Avaghade  ·  Growth & MarTech PM', {
    x: 0.5, y: 6.9, w: 6.5, h: 0.25,
    fontSize: 9.5, bold: false, color: C.muted, fontFace: 'Calibri',
  });

  // Watermark
  s.addText('R', {
    x: 5.8, y: 1.2, w: 4.5, h: 5,
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

  s.addText('Revolut Shows the Same Upgrade Banner to 70M Users Regardless of Intent, Fit, or Financial Benefit.', {
    x: 0.5, y: 0.57, w: 9, h: 0.9,
    fontSize: 23, bold: true, color: C.white, fontFace: 'Calibri',
  });

  const cols = [
    { icon: '📢', stat: '70M', label: 'Users receive\nidentical upgrade CTA', sub: 'Standard, Premium, Metal — same banner\nregardless of usage signals or plan fit' },
    { icon: '💸', stat: '−89%', label: 'Offer budget wasted\non non-convertible users', sub: 'HIGH-intent users (upgrade anyway) and LOW-intent\nusers (will not convert) receive costly incentives' },
    { icon: '📉', stat: '11%', label: 'Baseline upgrade CVR\non blanket campaigns', sub: 'Industry-leading CVR sits at 22–28% with\npropensity-targeted messaging — gap is 2× revenue' },
  ];

  cols.forEach((col, i) => {
    const x = 0.45 + i * 3.1;
    s.addShape(prs.ShapeType.rect, {
      x, y: 1.68, w: 2.9, h: 3.3,
      fill: { color: '0A1428' }, line: { color: C.bright, width: 0.8 },
      shadow: makeShadow(),
    });
    s.addText(col.icon, { x, y: 1.85, w: 2.9, h: 0.45, fontSize: 22, align: 'center' });
    s.addText(col.stat, {
      x, y: 2.32, w: 2.9, h: 0.72,
      fontSize: 38, bold: true, color: C.gold, align: 'center', fontFace: 'Calibri',
    });
    s.addText(col.label, {
      x: x + 0.12, y: 3.05, w: 2.65, h: 0.62,
      fontSize: 11.5, bold: true, color: C.white, align: 'center', fontFace: 'Calibri',
    });
    s.addText(col.sub, {
      x: x + 0.1, y: 3.7, w: 2.7, h: 0.75,
      fontSize: 9.5, bold: false, color: '7CA8D8', align: 'center', fontFace: 'Calibri',
    });
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.45, y: 5.2, w: 9.2, h: 0.9,
    fill: { color: '0A1428' }, line: { color: C.bright, width: 1 },
  });
  s.addText('Root cause: No intelligence between CDP data and campaign execution. Revolut knows exactly which users would benefit financially from upgrading — but that signal never reaches the upgrade nudge. The result is wasted offer budget, low CVR, and premium revenue left on the table.', {
    x: 0.65, y: 5.28, w: 8.9, h: 0.72,
    fontSize: 10.5, bold: false, color: 'C8D4F0', fontFace: 'Calibri',
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

  s.addText('The Upgrade Decision Is Already Made in the Data. GrowthIQ Just Surfaces It at the Right Moment.', {
    x: 0.5, y: 0.57, w: 9, h: 0.82,
    fontSize: 22, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  // Left — current
  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 1.55, w: 4.2, h: 3.95,
    fill: { color: 'FFF3F3' }, line: { color: 'E57373', width: 1.2 },
    shadow: makeShadow(),
  });
  s.addText('❌  Current: One Banner, 70M Users', {
    x: 0.53, y: 1.67, w: 4.0, h: 0.35,
    fontSize: 13.5, bold: true, color: 'C62828', fontFace: 'Calibri',
  });
  const currentPts = [
    'Same "Upgrade to Premium" CTA for every user',
    'Offer size fixed — regardless of user value or intent',
    'No signal from usage data (FX limit hits, blocked features)',
    'HIGH-intent users receive discount they would not need',
    'LOW-intent users receive pressure and ignore it',
    'No narrative personalisation — one message for all segments',
  ];
  currentPts.forEach((pt, i) => {
    s.addText('• ' + pt, {
      x: 0.58, y: 2.12 + i * 0.53, w: 3.88, h: 0.46,
      fontSize: 11.5, color: '4A1A1A', fontFace: 'Calibri',
    });
  });

  // VS
  s.addShape(prs.ShapeType.ellipse, {
    x: 4.5, y: 2.92, w: 0.7, h: 0.7,
    fill: { color: C.ink }, line: { color: C.ink },
  });
  s.addText('VS', {
    x: 4.5, y: 2.92, w: 0.7, h: 0.7,
    fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle', fontFace: 'Calibri',
  });

  // Right — proposed
  s.addShape(prs.ShapeType.rect, {
    x: 5.12, y: 1.55, w: 4.3, h: 3.95,
    fill: { color: 'EEF3FF' }, line: { color: C.bright, width: 1.5 },
    shadow: makeShadow(),
  });
  s.addText('✅  GrowthIQ: Signal-Led Intelligence', {
    x: 5.27, y: 1.67, w: 4.1, h: 0.35,
    fontSize: 13.5, bold: true, color: '1A3A8A', fontFace: 'Calibri',
  });
  const proposedPts = [
    'CDP scores upgrade propensity per user per tier (Plus / Premium / Metal)',
    'HIGH (>0.75): no offer — they convert on value alone',
    'MED (0.55–0.74): minimum calibrated incentive (1 month free)',
    'LOW (<0.55): suppressed — do not spend, do not pressure',
    'Story variant engine picks Aspiration / Fear / Social Proof per user',
    'Budget concentrates on the convertible middle — not the extremes',
  ];
  proposedPts.forEach((pt, i) => {
    s.addText('• ' + pt, {
      x: 5.27, y: 2.12 + i * 0.53, w: 4.05, h: 0.46,
      fontSize: 11.5, color: '1A2A5A', fontFace: 'Calibri',
    });
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 5.65, w: 9.1, h: 0.72,
    fill: { color: C.dark }, line: { color: C.bright, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('Key insight: Revolut already has the data to know which users would financially benefit from upgrading. GrowthIQ is the layer that turns that data into the right message, to the right person, with the right offer — automatically.', {
    x: 0.58, y: 5.73, w: 8.8, h: 0.58,
    fontSize: 10.5, bold: false, color: 'C8D4F0', fontFace: 'Calibri',
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

  s.addText('5-Layer Stack: From Usage Signal to Calibrated Upgrade Revenue', {
    x: 0.5, y: 0.57, w: 9, h: 0.52,
    fontSize: 22, bold: true, color: C.white, fontFace: 'Calibri',
  });

  const steps = [
    { n: '01', title: 'CDP — Unified Usage Profile', body: 'Ingest real-time signals: FX limit hits, locked feature views, ATM blocks, spend velocity, crypto activity, travel patterns. Build a scored profile per user updated on every session.', color: C.bright },
    { n: '02', title: 'Propensity Scoring Per Tier', body: 'Score each user\'s probability to upgrade to Plus, Premium, and Metal independently. Threshold: HIGH >0.75, MED 0.55-0.74, LOW <0.55. Recalculate weekly.', color: C.bright },
    { n: '03', title: 'Story Variant Selection', body: 'Fear: user has hit FX limit 3+x or been blocked. Aspiration: high spend, travel-active, engaged crypto user. Social Proof: 3+ contacts on Premium, cohort comparison signal.', color: C.purple },
    { n: '04', title: 'Offer Calibration Engine', body: 'HIGH band: no offer, value message only. MED band: minimum effective incentive — 1 month free trial. LOW band: suppressed entirely. Offer cost only fires when CVR probability justifies it.', color: C.purple },
    { n: '05', title: 'Analytics Feedback Loop', body: 'Track CVR by variant x band x tier. Winner retrains propensity model. HIGH threshold adjusts upward if no-offer CVR holds. Every campaign cycle gets smarter than the last.', color: C.gold },
  ];

  steps.forEach((st, i) => {
    const y = 1.35 + i * 1.0;
    s.addShape(prs.ShapeType.rect, {
      x: 0.38, y, w: 9.24, h: 0.88,
      fill: { color: '0D0F14' }, line: { color: st.color, width: 0.7 },
    });
    s.addText(st.n, {
      x: 0.52, y: y + 0.12, w: 0.55, h: 0.62,
      fontSize: 26, bold: true, color: st.color, fontFace: 'Calibri',
    });
    s.addText(st.title, {
      x: 1.18, y: y + 0.1, w: 3.1, h: 0.35,
      fontSize: 13, bold: true, color: C.white, fontFace: 'Calibri',
    });
    s.addText(st.body, {
      x: 1.18, y: y + 0.44, w: 8.3, h: 0.38,
      fontSize: 10, bold: false, color: '99B4CC', fontFace: 'Calibri',
    });
  });

  s.addText('PhonePe benchmark: Propensity-to-Transact ML (same architecture) cut marketing burn 32% at 350M MAU — shipped and measured, not projected.', {
    x: 0.5, y: 6.85, w: 9.2, h: 0.24,
    fontSize: 8.5, bold: false, color: C.muted, fontFace: 'Calibri', italic: true,
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

  s.addText('4 Key Surfaces — Built on Revolut Design Language', {
    x: 0.5, y: 0.57, w: 9, h: 0.48,
    fontSize: 22, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  const screens = [
    {
      num: '01', title: 'CDP\nUser Profile',
      desc: 'Upgrade propensity bars per tier, behavioural signal feed with dot-coded severity, auto-selected story variant, offer band + NBO chip.',
      ascii: '┌──────────┐\n│SM  72%   │\n│Prem HIGH │\n│FX hit 3x │\n│Variant:  │\n│FEAR ⚠️   │\n│1mo free  │\n└──────────┘',
    },
    {
      num: '02', title: 'Campaign\nBuilder',
      desc: 'Segment by MED band, variant card selection (Fear / Aspiration / Social Proof), offer config, message preview, one-click launch with confetti.',
      ascii: '┌──────────┐\n│1.82M usr │\n│MED band  │\n│[FEAR ✓]  │\n│[Aspira ] │\n│[Social ] │\n│LAUNCH →  │\n└──────────┘',
    },
    {
      num: '03', title: 'Consumer\nUpgrade Nudge',
      desc: 'In-app Fear variant card: "You left £23 on the table." Exact £ breakdown. CTA: 1 month free. Shown only to MED-band users at peak engagement moment.',
      ascii: '┌──────────┐\n│⚠️ Standard│\n│ cost you │\n│  £23 ↑  │\n│FX:−£23.4│\n│Prem: £0  │\n│Try Free→ │\n└──────────┘',
    },
    {
      num: '04', title: 'Analytics\nDashboard',
      desc: 'CVR by variant x band, budget efficiency waterfall (blanket £2.84M vs GrowthIQ £1.82M = £1.02M saved), cohort upgrade rate, model retrain recommendation.',
      ascii: '┌──────────┐\n│CVR  24%  │\n│Fear  31% │\n│Aspir 24% │\n│Soc   19% │\n│Saved     │\n│£1.02M    │\n└──────────┘',
    },
  ];

  screens.forEach((sc, i) => {
    const x = 0.38 + i * 2.36;
    s.addShape(prs.ShapeType.rect, {
      x, y: 1.18, w: 2.18, h: 5.12,
      fill: { color: C.dark }, line: { color: C.bright, width: 1 },
      shadow: makeShadow(),
    });
    s.addText(sc.num, {
      x: x + 0.12, y: 1.28, w: 0.5, h: 0.3,
      fontSize: 10, bold: true, color: C.bright, fontFace: 'Calibri',
    });
    s.addText(sc.title, {
      x: x + 0.1, y: 1.56, w: 1.98, h: 0.55,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri',
    });
    s.addText(sc.ascii, {
      x: x + 0.08, y: 2.16, w: 2.02, h: 1.95,
      fontSize: 8, bold: false, color: C.bright, fontFace: 'Courier New',
    });
    s.addText(sc.desc, {
      x: x + 0.1, y: 4.16, w: 2.0, h: 1.88,
      fontSize: 9.5, bold: false, color: '99B4CC', fontFace: 'Calibri',
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

  s.addText('Projected Impact — Grounded in PhonePe Proof at 350M MAU Scale', {
    x: 0.5, y: 0.57, w: 9, h: 0.52,
    fontSize: 22, bold: true, color: C.white, fontFace: 'Calibri',
  });

  // Left — subscriber impact
  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 1.25, w: 4.45, h: 0.38,
    fill: { color: '0A1428' }, line: { color: C.bright, width: 0.5 },
  });
  s.addText('SUBSCRIBER & PRODUCT IMPACT', {
    x: 0.55, y: 1.3, w: 4.1, h: 0.28,
    fontSize: 9, bold: true, color: C.bright, charSpacing: 1.5, fontFace: 'Calibri',
  });

  const leftM = [
    { val: '+22%', lbl: 'Plan upgrade CVR', sub: 'MED-band targeted cohort vs 11% blanket baseline' },
    { val: '+38%', lbl: 'HIGH-band upgrade rate (no offer)', sub: 'Value-only message converts at 2× previous no-offer control' },
    { val: '−34%', lbl: 'Upgrade campaign spend', sub: 'Suppressing LOW band + no offer for HIGH = £1.02M saved per cycle' },
    { val: '+£8.40', lbl: 'Monthly revenue uplift per converted user', sub: 'Standard (£0) to Premium (£9.99) net of 1-month free trial cost' },
  ];
  leftM.forEach((m, i) => {
    const y = 1.78 + i * 0.9;
    s.addShape(prs.ShapeType.rect, {
      x: 0.38, y, w: 4.45, h: 0.82,
      fill: { color: '0A1428' }, line: { color: '1A2A48', width: 0.5 },
    });
    s.addText(m.val, {
      x: 0.5, y: y + 0.08, w: 1.3, h: 0.45,
      fontSize: 28, bold: true, color: C.bright, fontFace: 'Calibri',
    });
    s.addText(m.lbl, {
      x: 1.85, y: y + 0.08, w: 2.88, h: 0.28,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri',
    });
    s.addText(m.sub, {
      x: 1.85, y: y + 0.38, w: 2.88, h: 0.32,
      fontSize: 9.5, bold: false, color: '7CA8D8', fontFace: 'Calibri',
    });
  });

  // Right — company impact
  s.addShape(prs.ShapeType.rect, {
    x: 5.1, y: 1.25, w: 4.5, h: 0.38,
    fill: { color: '1A1000' }, line: { color: C.gold, width: 0.5 },
  });
  s.addText('COMPANY / REVENUE IMPACT', {
    x: 5.28, y: 1.3, w: 4.2, h: 0.28,
    fontSize: 9, bold: true, color: C.gold, charSpacing: 1.5, fontFace: 'Calibri',
  });

  const rightM = [
    { val: '−32%', lbl: 'Marketing spend per conversion', sub: 'PhonePe benchmark: shipped and measured at 350M MAU' },
    { val: '£1.02M', lbl: 'Campaign budget saved per cycle', sub: 'Blanket £2.84M vs GrowthIQ £1.82M concentrated on MED band' },
    { val: '3×', lbl: 'CLTV uplift: Standard to Premium user', sub: 'Premium users spend more, churn less, and use more Revolut products' },
    { val: '+26%', lbl: 'User engagement increase', sub: 'PhonePe benchmark: ML-driven targeting vs static campaigns' },
  ];
  rightM.forEach((m, i) => {
    const y = 1.78 + i * 0.9;
    s.addShape(prs.ShapeType.rect, {
      x: 5.1, y, w: 4.5, h: 0.82,
      fill: { color: '100A00' }, line: { color: '2A1A00', width: 0.5 },
    });
    s.addText(m.val, {
      x: 5.22, y: y + 0.08, w: 1.5, h: 0.45,
      fontSize: 26, bold: true, color: C.gold, fontFace: 'Calibri',
    });
    s.addText(m.lbl, {
      x: 6.76, y: y + 0.08, w: 2.7, h: 0.28,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri',
    });
    s.addText(m.sub, {
      x: 6.76, y: y + 0.38, w: 2.7, h: 0.32,
      fontSize: 9.5, bold: false, color: 'B89A30', fontFace: 'Calibri',
    });
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 5.6, w: 9.22, h: 0.72,
    fill: { color: '0A1428' }, line: { color: C.bright, width: 1 },
  });
  s.addText('PhonePe proof: Propensity-to-Transact ML replaced £1,000+ Cr/yr of static targeting with real-time scoring — 32% marketing burn reduction, sustained GMV. Same architecture. Same outcome, applied to Revolut plan upgrade revenue.', {
    x: 0.58, y: 5.68, w: 9.0, h: 0.56,
    fontSize: 10.5, bold: false, color: 'C8D4F0', fontFace: 'Calibri',
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

  s.addText('Every GrowthIQ Layer Maps Directly to Something I Built and Measured at PhonePe.', {
    x: 0.5, y: 0.57, w: 9, h: 0.52,
    fontSize: 22, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 1.22, w: 4.6, h: 4.52,
    fill: { color: C.dark }, line: { color: C.bright, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('What I Built at PhonePe', {
    x: 0.55, y: 1.35, w: 4.2, h: 0.32,
    fontSize: 13, bold: true, color: C.bright, fontFace: 'Calibri',
  });

  const ppProof = [
    'Propensity-to-Transact ML: Replaced £1,000+ Cr/yr static targeting with real-time user scoring — 32% marketing burn reduction at 350M MAU. Exact GrowthIQ CDP propensity architecture.',
    'Story Variant Engine (GrowthOS): Aspiration / Fear / Social Proof narrative selection per user segment — deployed in CRM builder with brand self-serve tooling.',
    'Offer Calibration Engine: HIGH-band suppression, MED-band minimum offer, LOW-band suppressed — £106 Cr saved per campaign cycle at Navi, 32% CAC reduction at PhonePe.',
    'Rewards Marketplace Redesign: Rebuilt £100 Cr/yr static portfolio into ML-driven ecosystem — 500+ brand partners, 26% engagement uplift, 11% YoY revenue growth.',
    'Dynamic Incentivisation (Pincode): Context-aware triggers at cart exit — 60% abandonment reduction. Same real-time signal architecture as GrowthIQ Fear variant trigger.',
  ];
  ppProof.forEach((pt, i) => {
    s.addText('▸  ' + pt, {
      x: 0.52, y: 1.82 + i * 0.76, w: 4.3, h: 0.68,
      fontSize: 10.5, bold: false, color: 'C8D4F0', fontFace: 'Calibri',
    });
  });

  s.addShape(prs.ShapeType.rect, {
    x: 5.2, y: 1.22, w: 4.42, h: 4.52,
    fill: { color: 'EEF3FF' }, line: { color: C.bright, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('How It Maps to This Role', {
    x: 5.36, y: 1.35, w: 4.1, h: 0.32,
    fontSize: 13, bold: true, color: '1A3A8A', fontFace: 'Calibri',
  });

  const mapping = [
    '→  GrowthIQ CDP propensity scoring (same real-time ML, same band architecture)',
    '→  GrowthIQ Story Variant Engine (same Aspiration / Fear / Social Proof selection logic)',
    '→  GrowthIQ Offer Calibration (same suppress HIGH, calibrate MED, suppress LOW)',
    '→  Revolut PMM campaign tooling (same self-serve brand CRM builder, same analytics loop)',
    '→  Revolut upgrade GTM (same context-aware trigger at peak engagement moment)',
  ];
  mapping.forEach((pt, i) => {
    s.addText(pt, {
      x: 5.36, y: 1.82 + i * 0.76, w: 4.16, h: 0.68,
      fontSize: 10.5, bold: false, color: '1A2A5A', fontFace: 'Calibri',
    });
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 5.88, w: 9.22, h: 0.56,
    fill: { color: C.dark }, line: { color: C.bright, width: 1 },
  });
  s.addText('GrowthIQ is not a concept I designed for this application. It is an architecture I built, shipped, and measured — at a scale 5× larger than Revolut today.', {
    x: 0.6, y: 5.96, w: 8.8, h: 0.42,
    fontSize: 10.5, bold: false, color: 'C8D4F0', fontFace: 'Calibri', italic: true,
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

  s.addText('CDP Foundation → Intelligence Engine → Full Calibration Loop', {
    x: 0.5, y: 0.57, w: 9, h: 0.52,
    fontSize: 22, bold: true, color: C.white, fontFace: 'Calibri',
  });

  const phases = [
    {
      phase: 'Phase 1', timeline: 'Month 1–2',
      title: 'CDP Signal Pipeline',
      points: [
        'Ingest: FX limit hits, blocked features, spend velocity, travel signals',
        'Build propensity score per user per tier — Plus, Premium, Metal',
        'Band users: HIGH / MED / LOW — weekly recalculation',
        'Success metric: 80% of MAU scored within 6 weeks of launch',
      ],
      color: C.bright,
    },
    {
      phase: 'Phase 2', timeline: 'Month 3–4',
      title: 'Story Variant + Campaign Builder',
      points: [
        'A/B test 3 story variants vs current control across MED band',
        'Deploy PMM self-serve campaign builder (segment, variant, offer, launch)',
        'Run first GrowthIQ campaign on Premium tier — 1.82M MED users',
        'Success metric: +15% CVR lift vs blanket campaign baseline at 4 weeks',
      ],
      color: C.purple,
    },
    {
      phase: 'Phase 3', timeline: 'Month 5–6',
      title: 'Full Calibration + Analytics Loop',
      points: [
        'Deploy offer calibration: HIGH no-offer, MED incentive, LOW suppress',
        'Launch analytics dashboard with variant x band attribution',
        'Retrain propensity model on 2 campaign cycles of observed CVR data',
        'Success metric: £1M+ campaign budget saving vs blanket spend confirmed',
      ],
      color: C.gold,
    },
  ];

  phases.forEach((ph, i) => {
    const x = 0.38 + i * 3.14;
    s.addShape(prs.ShapeType.rect, {
      x, y: 1.28, w: 2.98, h: 4.42,
      fill: { color: '0D0F14' }, line: { color: ph.color, width: 1.2 },
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
      x: x + 0.12, y: 1.75, w: 2.72, h: 0.52,
      fontSize: 12.5, bold: true, color: C.white, fontFace: 'Calibri',
    });
    ph.points.forEach((pt, j) => {
      s.addText('• ' + pt, {
        x: x + 0.14, y: 2.35 + j * 0.72, w: 2.7, h: 0.66,
        fontSize: 10, bold: false, color: '99B4CC', fontFace: 'Calibri',
      });
    });
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 5.85, w: 9.22, h: 0.55,
    fill: { color: '0A1428' }, line: { color: C.bright, width: 1 },
  });
  s.addText('What I need:  CDP data access  ·  1 Data Engineer (signal ingestion)  ·  1 ML Engineer (propensity model)  ·  PMM campaign builder UI (existing Revolut CMS or new)', {
    x: 0.6, y: 5.93, w: 8.8, h: 0.4,
    fontSize: 10, bold: false, color: 'C8D4F0', fontFace: 'Calibri',
  });

  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade', {
    x: 0.38, y: 6.9, w: 9.22, h: 0.22,
    fontSize: 8.5, bold: false, color: C.muted, fontFace: 'Calibri',
  });
}

prs.writeFile({ fileName: 'assets/revolut-growthiq-deck.pptx' })
  .then(() => console.log('✅  revolut-growthiq-deck.pptx written'))
  .catch(e => { console.error(e); process.exit(1); });
