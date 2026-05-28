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
pres.layout = 'LAYOUT_16x9'; // 10" x 5.625"
pres.author = 'Ajay Avaghade';
pres.title = 'MAF SellerOS — CEO Pitch Deck';

// ─────────────────────────────────────────────────────────────
// SLIDE 1 — COVER
// ─────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.dark };

  // Diagonal texture lines
  for (let i = 0; i < 9; i++) {
    s.addShape(pres.shapes.RECTANGLE, {
      x: -1.2 + i * 1.5, y: -0.5, w: 0.06, h: 9,
      fill: { color: '200A14' }, line: { color: '200A14' }, rotate: 25,
    });
  }

  // Left red accent bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.12, h: 5.625,
    fill: { color: C.red }, line: { color: C.red },
  });

  // Tag
  s.addText('MAF RETAIL  ·  MARKETPLACE PRODUCT  ·  CASE STUDY 01', {
    x: 0.35, y: 0.32, w: 7.5, h: 0.26,
    fontSize: 7.5, fontFace: 'Calibri', color: C.muted,
    charSpacing: 2, margin: 0,
  });

  // Product name
  s.addText('SellerOS', {
    x: 0.35, y: 0.9, w: 8, h: 1.4,
    fontSize: 72, fontFace: 'Calibri', color: C.white,
    bold: true, margin: 0,
  });

  // Subtitle
  s.addText('Turning SME Marketplace Sellers from Dormant to Thriving', {
    x: 0.35, y: 2.32, w: 7.5, h: 0.56,
    fontSize: 17, fontFace: 'Calibri', color: C.gold,
    margin: 0,
  });

  // Presenter
  s.addText('Ajay Avaghade  ·  Product Manager', {
    x: 0.35, y: 2.96, w: 6, h: 0.36,
    fontSize: 11, fontFace: 'Calibri', color: C.muted, margin: 0,
  });

  // Big stat box bottom-right
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.9, y: 3.5, w: 2.8, h: 1.65,
    fill: { color: '180A12' }, line: { color: C.gold },
  });
  s.addText('+60%', {
    x: 6.9, y: 3.62, w: 2.8, h: 0.78,
    fontSize: 42, fontFace: 'Calibri', color: C.gold,
    bold: true, align: 'center', margin: 0,
  });
  s.addText('active seller rate at 90 days', {
    x: 6.9, y: 4.42, w: 2.8, h: 0.5,
    fontSize: 9, fontFace: 'Calibri', color: C.muted,
    align: 'center', margin: 0,
  });

  // Bottom bar
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
    fontSize: 8, fontFace: 'Calibri', color: C.red,
    bold: true, charSpacing: 3, margin: 0,
  });

  s.addText('30% of Marketplace Sellers Go Inactive Within 90 Days\n— Killing GMV and the SME Mission', {
    x: 0.35, y: 0.6, w: 9.3, h: 1.1,
    fontSize: 22, fontFace: 'Calibri', color: C.white,
    bold: true, margin: 0,
  });

  // 3 stat columns
  const cols = [
    { icon: '📦', stat: '30%',        label: 'Seller inactivity rate',          sub: 'CAPEX deployed, GMV never generated' },
    { icon: '💸', stat: 'AED 640/mo', label: 'Avg GMV loss per inactive seller', sub: 'Compounded across 500+ seller cohort' },
    { icon: '🎯', stat: '0',          label: 'Actionable signals to sellers',    sub: 'No pricing intel, no campaign tools, no milestone incentive' },
  ];
  const cW = 2.72, cGap = 0.24, cStartX = 0.58;
  cols.forEach((col, i) => {
    const cx = cStartX + i * (cW + cGap);
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: 1.88, w: cW, h: 2.12,
      fill: { color: '180A12' }, line: { color: '3A1A28' },
      shadow: makeCardShadow(),
    });
    s.addText(col.icon, {
      x: cx + 0.15, y: 1.98, w: 0.5, h: 0.44, fontSize: 20, margin: 0,
    });
    s.addText(col.stat, {
      x: cx + 0.12, y: 2.48, w: cW - 0.2, h: 0.62,
      fontSize: 28, fontFace: 'Calibri', color: C.gold,
      bold: true, margin: 0,
    });
    s.addText(col.label, {
      x: cx + 0.12, y: 3.13, w: cW - 0.2, h: 0.4,
      fontSize: 9.5, fontFace: 'Calibri', color: C.white,
      bold: true, margin: 0,
    });
    s.addText(col.sub, {
      x: cx + 0.12, y: 3.55, w: cW - 0.2, h: 0.36,
      fontSize: 8, fontFace: 'Calibri', color: C.muted, margin: 0,
    });
  });

  // Bottom insight box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 4.22, w: 9.3, h: 0.88,
    fill: { color: '180A12' }, line: { color: '3A1020' },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 4.22, w: 0.06, h: 0.88,
    fill: { color: C.gold }, line: { color: C.gold },
  });
  s.addText('Root cause: Sellers lack pricing signals, campaign tools, and any milestone incentive to stay active. Category managers spend 80% of their time fighting fires instead of growing the category.', {
    x: 0.55, y: 4.27, w: 8.9, h: 0.78,
    fontSize: 9.5, fontFace: 'Calibri', color: 'C0A8B0',
    margin: 0, valign: 'middle',
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
    fontSize: 8, fontFace: 'Calibri', color: C.red,
    bold: true, charSpacing: 3, margin: 0,
  });

  s.addText("Sellers Don't Need Another Dashboard. They Need a Coach.", {
    x: 0.5, y: 0.5, w: 9, h: 0.88,
    fontSize: 26, fontFace: 'Calibri', color: C.ink,
    bold: true, margin: 0,
  });

  const colY = 1.55, colH = 2.62, lX = 0.5, rX = 5.22, cW = 4.28;

  // Left dark panel
  s.addShape(pres.shapes.RECTANGLE, {
    x: lX, y: colY, w: cW, h: colH,
    fill: { color: C.dark }, line: { color: '2A1020' },
    shadow: makeCardShadow(),
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: lX, y: colY, w: cW, h: 0.36,
    fill: { color: C.hero }, line: { color: C.hero },
  });
  s.addText('❌  Status Quo', {
    x: lX + 0.15, y: colY + 0.06, w: cW - 0.2, h: 0.28,
    fontSize: 10.5, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });
  s.addText([
    { text: '· No pricing signal → sellers price blind', options: { breakLine: true } },
    { text: '· No campaign tool → zero promotion capability', options: { breakLine: true } },
    { text: '· Flat onboarding → no behaviour change after day 1', options: { breakLine: true } },
    { text: '· Category manager reactive, not proactive' },
  ], {
    x: lX + 0.15, y: colY + 0.52, w: cW - 0.25, h: colH - 0.68,
    fontSize: 9.5, fontFace: 'Calibri', color: 'C0A8B0', margin: 0, valign: 'top',
  });

  // Right light panel
  s.addShape(pres.shapes.RECTANGLE, {
    x: rX, y: colY, w: cW, h: colH,
    fill: { color: C.lgray }, line: { color: 'DDD0D4' },
    shadow: makeCardShadow(),
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: rX, y: colY, w: cW, h: 0.36,
    fill: { color: C.green }, line: { color: C.green },
  });
  s.addText('✅  SellerOS', {
    x: rX + 0.15, y: colY + 0.06, w: cW - 0.2, h: 0.28,
    fontSize: 10.5, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });
  s.addText([
    { text: "· AI pricing nudge: 'Your SKU is 18% above median'", options: { breakLine: true } },
    { text: '· 3-click campaign builder + Eid-timed promotions', options: { breakLine: true } },
    { text: '· Milestone commission credit: hit AED 12K → earn AED 200 waiver', options: { breakLine: true } },
    { text: '· Health score + ops console proactively flags breaches' },
  ], {
    x: rX + 0.15, y: colY + 0.52, w: cW - 0.25, h: colH - 0.68,
    fontSize: 9.5, fontFace: 'Calibri', color: C.ink, margin: 0, valign: 'top',
  });

  // Arrow badge
  s.addShape(pres.shapes.OVAL, {
    x: 4.72, y: 2.58, w: 0.56, h: 0.56,
    fill: { color: C.red }, line: { color: C.white },
    shadow: makeCardShadow(),
  });
  s.addText('→', {
    x: 4.72, y: 2.58, w: 0.56, h: 0.56,
    fontSize: 18, fontFace: 'Calibri', color: C.white,
    bold: true, align: 'center', valign: 'middle', margin: 0,
  });

  // Bottom callout
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.38, w: 9, h: 0.84,
    fill: { color: 'FFF8E0' }, line: { color: 'E8C870' },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.38, w: 0.06, h: 0.84,
    fill: { color: C.gold }, line: { color: C.gold },
  });
  s.addText('Key insight: The milestone waiver costs MAF nothing net — by the time the seller hits the GMV goal, the additional commission has already paid for the credit many times over.', {
    x: 0.7, y: 4.43, w: 8.6, h: 0.74,
    fontSize: 9.5, fontFace: 'Calibri', color: '7A5C10',
    margin: 0, valign: 'middle',
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
    fontSize: 8, fontFace: 'Calibri', color: C.gold,
    bold: true, charSpacing: 3, margin: 0,
  });
  s.addText('Three Levers, One Health Score', {
    x: 0.35, y: 0.56, w: 9.3, h: 0.72,
    fontSize: 28, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });

  const steps = [
    { num: '1', title: 'Onboarding',       body: 'Seller joins marketplace. SellerOS health score activated. Milestone window opens: 90 days to AED 12K GMV.' },
    { num: '2', title: 'Pricing AI',        body: 'Daily: AI scans category median. SKUs overpriced by >10% trigger alert. Seller gets 1-click price recommendation.' },
    { num: '3', title: 'Campaign Builder',  body: 'Seller launches promotion in 3 clicks: SKU → discount % → audience → live. No ops team required.' },
    { num: '4', title: 'Milestone Engine',  body: 'Progress bar visible in dashboard. Nudges at 50%, 80%, 95% completion. Hit goal → AED 200 commission credit auto-applied.' },
    { num: '5', title: 'Ops Console',       body: 'Category manager sees seller health grid. Breach alert → send nudge or escalate. Real-time, no Excel lag.' },
  ];
  const sW = 1.66, sGap = 0.16, sX = 0.35, sY = 1.52, sH = 3.0;

  // Connecting dashed line
  s.addShape(pres.shapes.LINE, {
    x: sX + sW * 0.5, y: 1.84, w: (sW + sGap) * 4, h: 0,
    line: { color: '5A3040', width: 2, dashType: 'dash' },
  });

  steps.forEach((step, i) => {
    const px = sX + i * (sW + sGap);
    s.addShape(pres.shapes.RECTANGLE, {
      x: px, y: sY, w: sW, h: sH,
      fill: { color: '180A12' }, line: { color: '3A1828' },
      shadow: makeCardShadow(),
    });
    // Number badge
    s.addShape(pres.shapes.OVAL, {
      x: px + sW / 2 - 0.22, y: sY - 0.3, w: 0.44, h: 0.44,
      fill: { color: C.red }, line: { color: C.hero },
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
      fontSize: 8.5, fontFace: 'Calibri', color: 'C0A8B0',
      margin: 0, valign: 'top',
    });
  });

  // Bottom note
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 4.78, w: 9.3, h: 0.47,
    fill: { color: '120810' }, line: { color: '2A1020' },
  });
  s.addText('A/B tested at PhonePe: Milestone waiver drove 60% improvement in 90-day device activation vs flat discount control.', {
    x: 0.5, y: 4.82, w: 9.0, h: 0.39,
    fontSize: 8.5, fontFace: 'Calibri', color: C.muted,
    italic: true, margin: 0, valign: 'middle',
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
  s.addText('4 Screen States — Seller Journey End to End', {
    x: 0.5, y: 0.5, w: 9, h: 0.62,
    fontSize: 24, fontFace: 'Calibri', color: C.ink, bold: true, margin: 0,
  });

  const cards = [
    { num: '01', title: 'Seller Home',       body: 'Health score 72/100. GMV vs target. Milestone progress bar at 70%. Actionable alert: 2 SKUs overpriced.' },
    { num: '02', title: 'AI Pricing Engine',  body: 'SKU-by-SKU recommendation. Your price vs category median vs optimal. Est. revenue impact per change. 1-click apply.' },
    { num: '03', title: 'Campaign Live',      body: '3-click campaign builder → success state. Reach: 45K shoppers. Expected GMV: +AED 1.2K. Milestone contribution shown.' },
    { num: '04', title: 'Ops Console',        body: 'Category manager view. Seller health grid. Breach (red), At Risk (orange), Healthy (green). 1-click nudge or escalation.' },
  ];
  const accentColors = [C.red, C.gold, C.green, '0070C9'];
  const cW = 2.1, cGap = 0.16, cStartX = 0.42, cY = 1.3, cH = 3.48;

  cards.forEach((card, i) => {
    const cx = cStartX + i * (cW + cGap);
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cY, w: cW, h: cH,
      fill: { color: C.lgray }, line: { color: 'DDD0D4' },
      shadow: makeCardShadow(),
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cY, w: cW, h: 0.06,
      fill: { color: accentColors[i] }, line: { color: accentColors[i] },
    });
    // Number tag
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
      fontSize: 9, fontFace: 'Calibri', color: '4A3840',
      margin: 0, valign: 'top',
    });
  });

  s.addText('Prototype: maf-selleros-prototype.html  ·  All 4 states interactive', {
    x: 0.5, y: 4.98, w: 9, h: 0.28,
    fontSize: 7.5, fontFace: 'Calibri', color: C.muted,
    italic: true, align: 'center', margin: 0,
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
  s.addText('Built on PhonePe Proof Points — Adjusted for MAF Scale', {
    x: 0.35, y: 0.56, w: 9.3, h: 0.58,
    fontSize: 22, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });

  const sections = [
    {
      title: 'Seller Impact', color: C.gold, x: 0.35,
      metrics: [
        { val: '↑ 60%',       lbl: 'Active rate at 90 days',     sub: 'vs 40% baseline' },
        { val: '↑ 25%',       lbl: 'Avg monthly GMV per seller', sub: 'pricing AI + campaigns' },
        { val: '↓ 35%',       lbl: 'Early churn rate (1–3 mo)',  sub: 'milestone creates lock-in' },
        { val: '+AED 640/mo', lbl: 'GMV uplift from pricing AI', sub: 'per seller, projected' },
      ],
    },
    {
      title: 'MAF Platform ROI', color: C.green, x: 5.1,
      metrics: [
        { val: 'AED 200',    lbl: 'Max commission credit cost', sub: 'per milestone hit' },
        { val: 'AED 12,000', lbl: 'Min GMV before credit fires', sub: 'net-zero cost structure' },
        { val: '~1.7%',      lbl: 'Credit as % of GMV generated', sub: 'effectively free' },
        { val: '↑ 18%',      lbl: 'Category mgr capacity freed', sub: 'reactive → proactive' },
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
        fill: { color: '160A12' }, line: { color: '2E1422' },
        shadow: makeCardShadow(),
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

  // Divider
  s.addShape(pres.shapes.LINE, {
    x: 5.0, y: 1.3, w: 0, h: 3.42,
    line: { color: '3A1828', width: 1, dashType: 'dash' },
  });

  // Bottom
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 4.92, w: 9.3, h: 0.4,
    fill: { color: '180A12' }, line: { color: '2A1020' },
  });
  s.addText('The milestone credit only fires AFTER the seller has already hit the GMV target — making this a zero-net-cost stickiness lever.', {
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
  s.addText("I Built This at PhonePe. Here's the Direct Map.", {
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
    { text: '· 5,000+ B2B merchants — self-serve PG integration platform, 23% lower CAC', options: { breakLine: true } },
    { text: '· A/B tested milestone rental waiver vs cashback — 60% improvement in 90-day device activation', options: { breakLine: true } },
    { text: '· Rebuilt ₹100Cr rewards portfolio — ML marketplace, 500+ brand partners, 11% YoY revenue growth', options: { breakLine: true } },
    { text: '· Replaced 6 siloed workflows — unified ops layer, TAT 2d→4h, −68% escalation tickets' },
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
  s.addText('Direct Map to SellerOS', {
    x: rX + 0.15, y: pY + 0.06, w: pW - 0.2, h: 0.28,
    fontSize: 9.5, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });
  s.addText([
    { text: '→ SellerOS self-serve campaign builder + onboarding milestone', options: { breakLine: true } },
    { text: '→ SellerOS milestone commission credit mechanic', options: { breakLine: true } },
    { text: '→ SellerOS AI pricing engine + brand self-serve tooling', options: { breakLine: true } },
    { text: '→ SellerOS ops console with health score + breach alerting' },
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
  s.addText("This isn’t a case study. This is architecture I’ve shipped — applied to MAF’s marketplace problem.", {
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
  s.addText('Three Phases — Zero Risk Pilot to Full Category Rollout', {
    x: 0.35, y: 0.56, w: 9.3, h: 0.58,
    fontSize: 22, fontFace: 'Calibri', color: C.white, bold: true, margin: 0,
  });

  const phases = [
    { phase: 'PHASE 1', period: 'Month 1–2', title: 'Pilot: 50 Sellers / 1 Category', color: C.red,
      body: 'Deploy SellerOS to 50 newly onboarded Gourmet & Specialty sellers. Instrument health score, pricing alerts, milestone tracking. Measure 90-day activation rate vs control group.' },
    { phase: 'PHASE 2', period: 'Month 3–4', title: 'Scale + A/B Test', color: C.gold,
      body: "Run Milestone Credit vs Flat Discount A/B across 500 sellers. Launch campaign builder for ops-lite cohort. Tune pricing AI thresholds based on pilot data." },
    { phase: 'PHASE 3', period: 'Month 5–6', title: 'Full Marketplace Rollout', color: C.green,
      body: "Roll out SellerOS to all new seller onboarding. Integrate Ma’an SME programme as priority cohort. Scale ops console to all category managers." },
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

  // Ask box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 4.16, w: 9.3, h: 0.68,
    fill: { color: '120810' }, line: { color: '3A1020' },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 4.16, w: 0.06, h: 0.68,
    fill: { color: C.gold }, line: { color: C.gold },
  });
  s.addText('What I need: Seller transaction + inventory data access  ·  1 engineer (pricing API + milestone tracking)  ·  Alignment with Category Management team', {
    x: 0.55, y: 4.2, w: 8.9, h: 0.6,
    fontSize: 9, fontFace: 'Calibri', color: C.muted, margin: 0, valign: 'middle',
  });

  // Footer
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.22, w: 10, h: 0.4,
    fill: { color: '120810' }, line: { color: '120810' },
  });
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439', {
    x: 0.35, y: 5.25, w: 9.3, h: 0.32,
    fontSize: 9, fontFace: 'Calibri', color: C.muted, margin: 0, valign: 'middle',
  });
}

pres.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/maf-selleros-deck.pptx' });
console.log('✅ maf-selleros-deck.pptx written');
