const pptxgen = require('pptxgenjs');

const C = {
  dark:   '0C0408',
  hero:   '2C1020',
  red:    'E8001B',
  gold:   'C8952A',
  white:  'FFFFFF',
  lgray:  'F6F1F3',
  green:  '16A34A',
  blue:   '0070C9',
  ink:    '18100E',
  muted:  '9B8B90',
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 8, offset: 2, angle: 135, opacity: 0.12 });
const makeCardShadow = () => ({ type: 'outer', color: '000000', blur: 5, offset: 1, angle: 135, opacity: 0.10 });

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Ajay Avaghade';
pres.title = 'MAF CartIQ — CEO Pitch Deck';

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

  s.addText('MAF RETAIL  ·  CARREFOUR UAE APP  ·  CASE STUDY 02', {
    x: 0.35, y: 0.32, w: 7.5, h: 0.26,
    fontSize: 7.5, fontFace: 'Calibri', color: C.muted, charSpacing: 2, margin: 0,
  });
  s.addText('CartIQ', {
    x: 0.35, y: 0.9, w: 8, h: 1.4,
    fontSize: 72, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });
  s.addText('AI-Powered Cart Optimization + SHARE Loyalty Multipliers', {
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
  s.addText('+35% AOV', {
    x: 6.9, y: 3.62, w: 2.8, h: 0.54,
    fontSize: 26, fontFace: 'Calibri', color: C.gold, bold: true, align: 'center', margin: 0,
  });
  s.addText('−60% Cart Abandonment', {
    x: 6.9, y: 4.18, w: 2.8, h: 0.44,
    fontSize: 16, fontFace: 'Calibri', color: C.green, bold: true, align: 'center', margin: 0,
  });
  s.addText('PhonePe baseline — same architecture', {
    x: 6.9, y: 4.64, w: 2.8, h: 0.36,
    fontSize: 8, fontFace: 'Calibri', color: C.muted, align: 'center', margin: 0,
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
  s.addText('65% Cart Abandonment. Flat AOV.\nSHARE Points Sitting Unused at Checkout.', {
    x: 0.35, y: 0.6, w: 9.3, h: 1.0,
    fontSize: 24, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });

  const cols = [
    { icon: '🛒', stat: '65%',    label: 'Cart abandonment rate',          sub: 'Industry average — no cart-level intervention today' },
    { icon: '📉', stat: 'AED 0',  label: 'SHARE loyalty impact at checkout', sub: 'Points balance visible but no multiplier mechanic' },
    { icon: '🎯', stat: 'Static', label: 'Promotion strategy',             sub: "10% off AED 200+ — no cart composition intelligence" },
  ];
  const cW = 2.72, cGap = 0.24, cStartX = 0.58;
  cols.forEach((col, i) => {
    const cx = cStartX + i * (cW + cGap);
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: 1.82, w: cW, h: 2.18,
      fill: { color: '180A12' }, line: { color: '3A1A28' }, shadow: makeCardShadow(),
    });
    s.addText(col.icon, { x: cx + 0.15, y: 1.9, w: 0.5, h: 0.44, fontSize: 20, margin: 0 });
    s.addText(col.stat, {
      x: cx + 0.12, y: 2.4, w: cW - 0.2, h: 0.62,
      fontSize: 30, fontFace: 'Calibri', color: C.gold, bold: true, margin: 0,
    });
    s.addText(col.label, {
      x: cx + 0.12, y: 3.05, w: cW - 0.2, h: 0.38,
      fontSize: 9.5, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
    });
    s.addText(col.sub, {
      x: cx + 0.12, y: 3.46, w: cW - 0.2, h: 0.42,
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
  s.addText("The problem isn't the products — it's the cart experience. A shopper with bread, milk, and eggs in their cart is 68% likely to also want butter. MAF is missing this conversion moment every single time.", {
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
  s.addText('The Highest-Converting Promotion Is the One\nThat Knows What\'s Already in the Cart.', {
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
    { text: "· Generic '10% off AED 200+' banner", options: { breakLine: true } },
    { text: '· SHARE points visible but passive', options: { breakLine: true } },
    { text: '· Cart page is a list, not a sales surface', options: { breakLine: true } },
    { text: '· No basket completion intelligence' },
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
    fill: { color: C.blue }, line: { color: C.blue },
  });
  s.addText('✅  CartIQ', {
    x: rX + 0.15, y: colY + 0.06, w: cW - 0.2, h: 0.28,
    fontSize: 10.5, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });
  s.addText([
    { text: "· Real-time bundle: 'You have bread+eggs+milk — add butter for AED 8.25'", options: { breakLine: true } },
    { text: '· SHARE multiplier unlocks at incremental cart value thresholds', options: { breakLine: true } },
    { text: '· Cart page becomes the highest-converting screen in the app', options: { breakLine: true } },
    { text: '· 68% basket completion rate (PhonePe proof: same cart composition model)' },
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
    fill: { color: 'FFF8E0' }, line: { color: 'E8C870' },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.38, w: 0.06, h: 0.84,
    fill: { color: C.gold }, line: { color: C.gold },
  });
  s.addText("Key insight: Tying SHARE multipliers to cart value thresholds creates urgency without discounting. 'Add AED 38 to get 3× SHARE points' is more compelling than '10% off' — and costs MAF nothing extra.", {
    x: 0.7, y: 4.43, w: 8.6, h: 0.74,
    fontSize: 9.5, fontFace: 'Calibri', color: '7A5C10', margin: 0, valign: 'middle',
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
  s.addText('Cart Composition  ×  User History  ×  SHARE Threshold  =  CartIQ', {
    x: 0.35, y: 0.56, w: 9.3, h: 0.72,
    fontSize: 24, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });

  const steps = [
    { num: '1', title: 'Cart Scan',           body: "Real-time: CartIQ reads cart composition + user purchase history + time-of-day signal (morning → breakfast, Friday → weekend snacks)" },
    { num: '2', title: 'Bundle Generation',   body: "ML model identifies top bundle completion (68% of users with bread+milk+eggs also add butter). Bundle displayed as CartIQ card." },
    { num: '3', title: 'SHARE Multiplier',    body: "Cart value shown vs threshold. 'AED 38 to 3× SHARE points'. Progress bar creates urgency. No discount required." },
    { num: '4', title: 'One-Tap Add',         body: "Shopper adds bundle in 1 tap. Cart updates. SHARE multiplier unlocks. Checkout button glows green." },
    { num: '5', title: 'Merch Console',       body: "Merchandising team runs A/B tests on bundle rules. Acceptance rate + AOV uplift tracked per variant. Winning rules auto-promoted." },
  ];
  const sW = 1.66, sGap = 0.16, sX = 0.35, sY = 1.52, sH = 3.0;

  s.addShape(pres.shapes.LINE, {
    x: sX + sW * 0.5, y: 1.84, w: (sW + sGap) * 4, h: 0,
    line: { color: '5A3040', width: 2, dashType: 'dash' },
  });

  steps.forEach((step, i) => {
    const px = sX + i * (sW + sGap);
    s.addShape(pres.shapes.RECTANGLE, {
      x: px, y: sY, w: sW, h: sH,
      fill: { color: '180A12' }, line: { color: '3A1828' }, shadow: makeCardShadow(),
    });
    s.addShape(pres.shapes.OVAL, {
      x: px + sW / 2 - 0.22, y: sY - 0.3, w: 0.44, h: 0.44,
      fill: { color: C.blue }, line: { color: C.hero },
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
  s.addText('A/B tested at PhonePe/Pincode: Dynamic cart incentivization (cart × intent × time) drove 35% AOV uplift, 60% cart abandonment reduction.', {
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
  s.addText('4 Screen States — Cart to Checkout', {
    x: 0.5, y: 0.5, w: 9, h: 0.62,
    fontSize: 24, fontFace: 'Calibri', color: C.ink, bold: true, margin: 0,
  });

  const cards = [
    { num: '01', title: 'Smart Cart',         body: "SHARE multiplier progress banner. CartIQ bundle card below cart items. '68% of shoppers with this cart also added these.' 1-tap bundle add." },
    { num: '02', title: 'Bundle Detail',      body: "Hero product bundle view. Breakdown of what's included + why CartIQ recommended it. Price saving vs individual items. SHARE bonus unlock confirmation." },
    { num: '03', title: 'Checkout Success',   body: 'Confetti. SHARE points earned (breakdown: base + multiplier bonus). Order total with bundle AOV uplift shown. You earned 840 SHARE points this order.' },
    { num: '04', title: 'Merch A/B Console',  body: 'Bundle acceptance rates across all active bundles. A/B test dashboard: Variant A (SHARE trigger) vs Variant B (discount). Winner badge when significance reached.' },
  ];
  const accentColors = [C.red, C.blue, C.green, C.gold];
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

  s.addText('Prototype: maf-cartiq-prototype.html  ·  All 4 states interactive', {
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
  s.addText("The Same Model I Shipped at PhonePe — Now on Carrefour UAE's Cart Page", {
    x: 0.35, y: 0.56, w: 9.3, h: 0.58,
    fontSize: 20, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });

  const sections = [
    {
      title: 'Consumer Impact', color: C.gold, x: 0.35,
      metrics: [
        { val: '↑ 35%', lbl: 'AOV uplift', sub: 'PhonePe baseline — same cart architecture' },
        { val: '↓ 60%', lbl: 'Cart abandonment reduction', sub: 'dynamic bundle + SHARE trigger' },
        { val: '↑ 20%', lbl: '30-day retention improvement', sub: 'D30 metric at PhonePe' },
        { val: '↑ 58%', lbl: 'Breakfast bundle acceptance', sub: 'prototype A/B test result' },
      ],
    },
    {
      title: 'MAF Platform ROI', color: C.blue, x: 5.1,
      metrics: [
        { val: 'AED 0',  lbl: 'Net discount cost', sub: 'SHARE multiplier, not markdown' },
        { val: '↑ 11%',  lbl: 'Avg order value per interaction', sub: 'CartIQ bundle acceptance' },
        { val: '3×',     lbl: 'SHARE multiplier vs GMV', sub: 'SHARE is a MAF asset — net zero cost' },
        { val: '↑ 22%',  lbl: 'Conversion when bundle present', sub: 'vs control (no suggestion)' },
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
  s.addText("CartIQ turns the cart page from a passive list into MAF's highest-converting screen — at near-zero incremental cost.", {
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
  s.addText('I Built the Exact Architecture This Needs.', {
    x: 0.5, y: 0.5, w: 9, h: 0.62,
    fontSize: 26, fontFace: 'Calibri', color: C.ink, bold: true, margin: 0,
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
    { text: '· PhonePe Pincode: Built dynamic cart incentivisation engine — cart value × user intent × time-of-day signals', options: { breakLine: true } },
    { text: '· Result: 35% AOV uplift, 60% cart abandonment reduction, 20% D30 retention improvement', options: { breakLine: true } },
    { text: '· PhonePe: Deployed Propensity-to-Transact ML models replacing static cohorts across 350M+ MAU', options: { breakLine: true } },
    { text: '· PhonePe: Built A/B test infrastructure for offer variants — tracked across checkout funnel' },
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
    fill: { color: C.blue }, line: { color: C.blue },
  });
  s.addText('Direct Map to CartIQ', {
    x: rX + 0.15, y: pY + 0.06, w: pW - 0.2, h: 0.28,
    fontSize: 9.5, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });
  s.addText([
    { text: '→ CartIQ cart composition + bundle recommendation engine', options: { breakLine: true } },
    { text: '→ CartIQ SHARE multiplier threshold mechanic', options: { breakLine: true } },
    { text: '→ CartIQ user-level personalisation (past purchases + time signal)', options: { breakLine: true } },
    { text: '→ CartIQ merch console A/B test dashboard' },
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
  s.addText('The basket completion rate, SHARE multiplier trigger, and merch A/B console are all direct architectural descendants of systems I shipped at PhonePe.', {
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
  s.addText('Pilot on Grocery Cart → Prove Lift → Full App Rollout', {
    x: 0.35, y: 0.56, w: 9.3, h: 0.58,
    fontSize: 22, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });

  const phases = [
    { phase: 'PHASE 1', period: 'Month 1–2', title: 'Grocery Cart Pilot', color: C.red,
      body: 'Deploy CartIQ to Grocery category cart page. A/B: CartIQ bundle card vs control (no bundle). Measure: AOV uplift, bundle acceptance rate, SHARE multiplier unlock rate.' },
    { phase: 'PHASE 2', period: 'Month 3–4', title: 'SHARE Integration + Categories', color: C.gold,
      body: 'Integrate SHARE multiplier mechanic with loyalty platform. Expand to Household, Baby, and Beauty categories. Launch merch console for category team self-serve bundle creation.' },
    { phase: 'PHASE 3', period: 'Month 5–6', title: 'Full App Rollout + Personalisation', color: C.green,
      body: 'CartIQ live across all categories. User-level personalisation (purchase history → bundle ranking). Real-time A/B test engine for bundle rule optimisation.' },
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
  s.addText('What I need: CartIQ API access to SHARE loyalty system  ·  Past purchase history data  ·  1 engineer + 1 data scientist  ·  Merchandising team as co-owner', {
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

pres.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/maf-cartiq-deck.pptx' });
console.log('✅ maf-cartiq-deck.pptx written');
