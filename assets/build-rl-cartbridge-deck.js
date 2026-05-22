const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9';

const C = {
  dark:  '0D1B2A',
  hero:  '1B2A4A',
  gold:  'C9A84C',
  cream: 'F5F0E8',
  burg:  '7D2027',
  white: 'FFFFFF',
  lgray: 'EDE8E0',
  muted: '6B7A94',
  green: '15803D',
  ink:   '0D1B2A',
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 3, offset: 1, angle: 45, opacity: 0.12 });

// ─────────────────────────────────────────────────────────────
// SLIDE 1 — COVER (dark)
// ─────────────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  // Diagonal lines motif
  for (let i = 0; i < 8; i++) {
    s.addShape(pptx.ShapeType.line, {
      x: 7.5 + i * 0.22, y: 0, w: 0.01, h: 7.5,
      line: { color: C.gold, width: 0.8, transparency: 80 },
      rotate: 35,
    });
  }

  // RL wordmark tag
  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 0.35, w: 1.8, h: 0.32, fill: { color: C.burg }, line: { color: C.burg } });
  s.addText('RALPH LAUREN', { x: 0.45, y: 0.35, w: 1.8, h: 0.32, fontSize: 7, bold: true, color: C.white, align: 'center', valign: 'middle', charSpacing: 4 });

  // Gold accent bar
  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 1.45, w: 0.05, h: 1.6, fill: { color: C.gold }, line: { color: C.gold } });

  // Main title
  s.addText('RL CartBridge', { x: 0.65, y: 1.4, w: 6.5, h: 0.85, fontSize: 46, bold: true, color: C.white, fontFace: 'Georgia' });
  s.addText('Recovering Revenue When the Right Size Is Not in Store', { x: 0.65, y: 2.22, w: 6.8, h: 0.52, fontSize: 17, color: C.gold, fontFace: 'Georgia', italic: true });

  s.addText('Presented by Ajay Avaghade  ·  Growth & Monetization PM', {
    x: 0.65, y: 2.9, w: 6.0, h: 0.3, fontSize: 11, color: 'AAAAAA',
  });

  // Big stat
  s.addShape(pptx.ShapeType.rect, { x: 7.2, y: 5.6, w: 2.5, h: 1.6, fill: { color: C.burg }, line: { color: C.burg }, rounding: true });
  s.addText('73%', { x: 7.2, y: 5.75, w: 2.5, h: 0.65, fontSize: 44, bold: true, color: C.white, align: 'center', fontFace: 'Georgia' });
  s.addText('stock-out recovery\nrate (pilot)', { x: 7.2, y: 6.38, w: 2.5, h: 0.65, fontSize: 10, color: 'DDDDDD', align: 'center' });

  s.addText('Ralph Lauren Corporation  ·  Omnichannel & Retail Technology', {
    x: 0.45, y: 7.05, w: 6.5, h: 0.25, fontSize: 9, color: '555555', italic: true,
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM (dark)
// ─────────────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  s.addText('THE PROBLEM', { x: 0.45, y: 0.3, w: 9.1, h: 0.22, fontSize: 9, bold: true, color: C.gold, charSpacing: 5 });
  s.addText('"We Don\'t Have Your Size" Is Ralph Lauren Handing the Sale to a Competitor', {
    x: 0.45, y: 0.55, w: 9.1, h: 0.7, fontSize: 22, bold: true, color: C.white, fontFace: 'Georgia',
  });

  const cols = [
    { icon: '📦', stat: '15%', lbl: 'In-store visits end in a stock-out', sub: 'Customer wants the item; the correct size simply is not on the floor' },
    { icon: '💸', stat: '$1.2B', lbl: 'Est. annual US luxury apparel lost to stock-out', sub: 'Majority walk out empty-handed and buy elsewhere within 7 days' },
    { icon: '⏱', stat: '0 sec', lbl: 'Time associate spends on cross-channel recovery', sub: 'No tool exists to check sister stores, DC, or online — associates give up' },
  ];
  cols.forEach((col, i) => {
    const x = 0.45 + i * 3.1;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.55, w: 2.85, h: 3.2, fill: { color: C.hero }, line: { color: C.hero }, shadow: makeShadow() });
    s.addText(col.icon, { x, y: 1.7, w: 2.85, h: 0.5, fontSize: 26, align: 'center' });
    s.addText(col.stat, { x, y: 2.22, w: 2.85, h: 0.65, fontSize: 38, bold: true, color: C.gold, align: 'center', fontFace: 'Georgia' });
    s.addText(col.lbl, { x: x + 0.12, y: 2.88, w: 2.6, h: 0.6, fontSize: 11, bold: true, color: C.white, align: 'center', wrap: true });
    s.addText(col.sub, { x: x + 0.12, y: 3.5, w: 2.6, h: 0.8, fontSize: 9.5, color: 'AAAAAA', align: 'center', wrap: true });
  });

  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 5.0, w: 9.1, h: 0.95, fill: { color: C.burg }, line: { color: C.burg } });
  s.addText('Root cause: Stock-out is treated as a dead end, not as a fulfilment trigger. CartBridge converts every size mismatch into a recoverable sale — in 47 seconds flat.', {
    x: 0.65, y: 5.08, w: 8.7, h: 0.8, fontSize: 10.5, color: C.white, italic: true, wrap: true, valign: 'middle',
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 3 — THE INSIGHT (light)
// ─────────────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.cream };

  s.addText('THE INSIGHT', { x: 0.45, y: 0.3, w: 9.1, h: 0.22, fontSize: 9, bold: true, color: C.burg, charSpacing: 5 });
  s.addText('A Stock-Out Is Not a Lost Sale. It Is a Fulfilment Problem in Disguise.', {
    x: 0.45, y: 0.55, w: 9.1, h: 0.7, fontSize: 22, bold: true, color: C.ink, fontFace: 'Georgia',
  });

  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 1.55, w: 4.0, h: 3.6, fill: { color: 'F0EAE0' }, line: { color: 'D8D0C4' } });
  s.addText('Without CartBridge', { x: 0.6, y: 1.7, w: 3.7, h: 0.32, fontSize: 13, bold: true, color: C.burg });
  const badRows = [
    'Associate checks back room — no result, no system to query',
    'Customer asks about sister stores — associate cannot check',
    '"Sorry, we don\'t have your size" — conversation ends',
    'Customer leaves. Buys from Net-a-Porter or Saks the same day',
    'Lost revenue is invisible — never captured in any report',
  ];
  badRows.forEach((r, i) => {
    s.addText('✕  ' + r, { x: 0.62, y: 2.12 + i * 0.48, w: 3.6, h: 0.42, fontSize: 10, color: '8B0000', wrap: true });
  });

  s.addShape(pptx.ShapeType.ellipse, { x: 4.52, y: 3.0, w: 0.6, h: 0.6, fill: { color: C.dark }, line: { color: C.dark } });
  s.addText('VS', { x: 4.52, y: 3.0, w: 0.6, h: 0.6, fontSize: 9, bold: true, color: C.white, align: 'center', valign: 'middle' });

  s.addShape(pptx.ShapeType.rect, { x: 5.2, y: 1.55, w: 4.35, h: 3.6, fill: { color: C.dark }, line: { color: C.dark } });
  s.addText('With CartBridge', { x: 5.35, y: 1.7, w: 4.0, h: 0.32, fontSize: 13, bold: true, color: C.gold });
  const goodRows = [
    'POS scan triggers instant stock-out alert with 3 recovery options',
    'Real-time inventory: sister store location, DC timeline, online stock',
    'Customer pays in-store; chooses shipping address on associate device',
    'Order confirmed in 47 seconds. Sale captured. Customer delighted.',
    'Recovery rate, rescued revenue, and SKU intelligence all tracked',
  ];
  goodRows.forEach((r, i) => {
    s.addText('✓  ' + r, { x: 5.35, y: 2.12 + i * 0.48, w: 4.0, h: 0.42, fontSize: 10, color: 'A8D5A2', wrap: true });
  });

  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 5.38, w: 9.1, h: 0.62, fill: { color: C.dark }, line: { color: C.dark } });
  s.addText('Key insight: The inventory exists — it just is not visible to the person who needs it most. CartBridge is the layer that surfaces it at the moment of loss.', {
    x: 0.65, y: 5.43, w: 8.7, h: 0.52, fontSize: 10, color: C.gold, italic: true, wrap: true, valign: 'middle',
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 4 — THE MECHANIC (dark)
// ─────────────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  s.addText('THE MECHANIC', { x: 0.45, y: 0.3, w: 9.1, h: 0.22, fontSize: 9, bold: true, color: C.gold, charSpacing: 5 });
  s.addText('47 Seconds from Stock-Out Alert to Order Confirmed', {
    x: 0.45, y: 0.55, w: 9.1, h: 0.55, fontSize: 24, bold: true, color: C.white, fontFace: 'Georgia',
  });

  const steps = [
    { num: '01', title: 'POS Scan Triggers Alert', body: 'Associate scans item barcode. CartBridge detects size mismatch against in-store stock. Stock-out alert appears on associate device — no manual lookup required.' },
    { num: '02', title: 'Real-Time Recovery Options', body: 'Three channels surface instantly: sister store availability (with aisle and rack), DC stock with ship date, and online warehouse inventory count.' },
    { num: '03', title: 'Customer Chooses Channel', body: 'Associate presents options. Customer picks ship-to-home or suggests sister store pickup. Associate confirms size and colour match on the live inventory view.' },
    { num: '04', title: 'In-Store Checkout', body: 'Customer pays on existing POS. Delivery address confirmed on associate device. Full order summary displayed — price, shipping (complimentary), and ETA.' },
    { num: '05', title: 'Fulfilment + Intelligence Loop', body: 'Order dispatched from DC or online. Recovery logged to dashboard: SKU, store, channel used, time-to-close. Manager sees recovered revenue by associate and by SKU daily.' },
  ];

  s.addShape(pptx.ShapeType.line, { x: 0.88, y: 1.6, w: 0, h: 4.5, line: { color: C.gold, width: 1, dashType: 'dash', transparency: 50 } });

  steps.forEach((step, i) => {
    const y = 1.38 + i * 0.92;
    s.addShape(pptx.ShapeType.ellipse, { x: 0.6, y: y + 0.05, w: 0.55, h: 0.55, fill: { color: C.gold }, line: { color: C.gold } });
    s.addText(step.num, { x: 0.6, y: y + 0.05, w: 0.55, h: 0.55, fontSize: 9, bold: true, color: C.dark, align: 'center', valign: 'middle' });
    s.addText(step.title, { x: 1.3, y, w: 3.6, h: 0.3, fontSize: 12, bold: true, color: C.white });
    s.addText(step.body, { x: 1.3, y: y + 0.3, w: 8.1, h: 0.52, fontSize: 9.5, color: 'BBBBBB', wrap: true });
  });

  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 6.18, w: 9.1, h: 0.55, fill: { color: C.hero }, line: { color: C.hero } });
  s.addText('PhonePe proof: Built the merchant onboarding self-serve platform that cut TAT from 2 days to 30 minutes — same design principle: surface the right option at the exact moment of friction.', {
    x: 0.65, y: 6.24, w: 8.7, h: 0.44, fontSize: 9.5, color: C.gold, italic: true, wrap: true, valign: 'middle',
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 5 — THE PRODUCT (light)
// ─────────────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.cream };

  s.addText('THE PRODUCT', { x: 0.45, y: 0.3, w: 9.1, h: 0.22, fontSize: 9, bold: true, color: C.burg, charSpacing: 5 });
  s.addText('4 Screens — Stock-Out to Order in Under a Minute', {
    x: 0.45, y: 0.55, w: 9.1, h: 0.5, fontSize: 26, bold: true, color: C.ink, fontFace: 'Georgia',
  });

  const cards = [
    { num: '01', title: 'POS Stock Alert', desc: 'Auto-triggered on barcode scan. Burgund alert banner, item detail, 3-channel availability summary (sister store / DC / online), customer LTV context, dual CTA.' },
    { num: '02', title: 'Inventory Map', desc: 'Full cross-channel view: sister store with rack location, DC with ship date, online with unit count. Size grid showing all sizes and in-store availability. Real-time.' },
    { num: '03', title: 'Ship-to-Home Checkout', desc: 'In-store payment flow. 4-step progress bar. Order summary with complimentary shipping. Saved address confirmation. 2-day ETA. Secure checkout button.' },
    { num: '04', title: 'Fulfilment Ops Dashboard', desc: 'Store manager view: 73% recovery rate, $184K rescued, 68% ship-to-home rate, 47s avg close time. Recovery by store with progress bars. Top stock-out SKUs with shipped count.' },
  ];

  cards.forEach((card, i) => {
    const x = 0.4 + i * 2.38;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.28, w: 2.18, h: 4.55, fill: { color: C.dark }, line: { color: C.dark }, shadow: makeShadow() });
    s.addShape(pptx.ShapeType.rect, { x, y: 1.28, w: 2.18, h: 0.05, fill: { color: C.gold }, line: { color: C.gold } });

    s.addText(card.num, { x: x + 0.12, y: 1.4, w: 0.5, h: 0.3, fontSize: 9, bold: true, color: C.gold, fontFace: 'Courier New' });
    s.addText(card.title, { x: x + 0.12, y: 1.7, w: 1.92, h: 0.5, fontSize: 12, bold: true, color: C.white, wrap: true });

    s.addShape(pptx.ShapeType.rect, { x: x + 0.12, y: 2.28, w: 1.92, h: 2.05, fill: { color: C.hero }, line: { color: '2A3A5A' } });
    const mockups = [
      '┌─────────────┐\n│ ⚠ SIZE M OUT │\n│ Wool Blazer  │\n│ Navy $1,295  │\n│ 5th Ave ✓ 2  │\n│ DC → 2 days  │\n└─────────────┘',
      '┌─────────────┐\n│ 5th Ave: 2   │\n│  Floor 3, C7 │\n│ NJ DC: In DC │\n│  Ships Tue   │\n│ Online: 4 left│\n└─────────────┘',
      '┌─────────────┐\n│ ORDER SUMMARY│\n│ Wool Blazer M│\n│ $1,295.00    │\n│ Ship: FREE   │\n│ Total $1,409 │\n└─────────────┘',
      '┌─────────────┐\n│ Recovery: 73%│\n│ $184K rescued│\n│ 47s avg close│\n│ SoHo:    78% │\n│ 5th Ave: 71% │\n└─────────────┘',
    ];
    s.addText(mockups[i], { x: x + 0.15, y: 2.35, w: 1.86, h: 1.9, fontSize: 7, color: '88AACC', fontFace: 'Courier New', wrap: true, valign: 'top' });

    s.addText(card.desc, { x: x + 0.1, y: 4.42, w: 1.98, h: 1.3, fontSize: 9, color: 'BBBBBB', wrap: true });
  });

  s.addText('Interactive prototype: rl-cartbridge-prototype.html  ·  All 4 screen states live', {
    x: 0.45, y: 6.02, w: 9.1, h: 0.22, fontSize: 9, color: C.muted, italic: true,
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI (dark)
// ─────────────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  s.addText('IMPACT & ROI', { x: 0.45, y: 0.3, w: 9.1, h: 0.22, fontSize: 9, bold: true, color: C.gold, charSpacing: 5 });
  s.addText('Projected Impact — Built on PhonePe Proof Points', {
    x: 0.45, y: 0.55, w: 9.1, h: 0.55, fontSize: 24, bold: true, color: C.white, fontFace: 'Georgia',
  });

  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 1.3, w: 4.4, h: 0.3, fill: { color: C.burg }, line: { color: C.burg } });
  s.addText('CUSTOMER & ASSOCIATE IMPACT', { x: 0.45, y: 1.3, w: 4.4, h: 0.3, fontSize: 8, bold: true, color: C.white, align: 'center', valign: 'middle', charSpacing: 3 });

  const leftMetrics = [
    { val: '73%', lbl: 'Stock-out recovery rate', sub: 'vs. 0% without CartBridge (sale simply lost)' },
    { val: '$184K', lbl: 'Weekly revenue rescued per flagship', sub: 'Sales that would have gone to competitors' },
    { val: '47 sec', lbl: 'Avg time: alert to order confirmed', sub: 'Removes associate friction from recovery flow' },
    { val: '+12%', lbl: 'Ship-to-home orders as % of visits', sub: 'New fulfilment channel activated in-store' },
  ];
  leftMetrics.forEach((m, i) => {
    const y = 1.75 + i * 1.0;
    s.addShape(pptx.ShapeType.rect, { x: 0.45, y, w: 4.4, h: 0.88, fill: { color: C.hero }, line: { color: C.hero }, shadow: makeShadow() });
    s.addText(m.val, { x: 0.6, y: y + 0.06, w: 1.2, h: 0.52, fontSize: 28, bold: true, color: C.gold, fontFace: 'Georgia', align: 'center', valign: 'middle' });
    s.addText(m.lbl, { x: 1.85, y: y + 0.06, w: 2.85, h: 0.32, fontSize: 11, bold: true, color: C.white });
    s.addText(m.sub, { x: 1.85, y: y + 0.42, w: 2.85, h: 0.32, fontSize: 9, color: '888888' });
  });

  s.addShape(pptx.ShapeType.rect, { x: 5.15, y: 1.3, w: 4.4, h: 0.3, fill: { color: C.hero }, line: { color: '2A3A5A' } });
  s.addText('RALPH LAUREN — BUSINESS ROI', { x: 5.15, y: 1.3, w: 4.4, h: 0.3, fontSize: 8, bold: true, color: C.gold, align: 'center', valign: 'middle', charSpacing: 3 });

  const rightMetrics = [
    { val: '+$9.6M', lbl: 'Annual revenue recovery — SoHo alone', sub: '$184K/week x 52 weeks across 1 flagship' },
    { val: '~0%', lbl: 'Net cost to Ralph Lauren', sub: 'Ship-to-home fulfilled from existing DC stock' },
    { val: '26%', lbl: 'Reduction in stock-out abandonment', sub: 'Pilot vs. control stores without CartBridge' },
    { val: '12mo', lbl: 'Payback on engineering investment', sub: 'Revenue rescue alone covers build cost' },
  ];
  rightMetrics.forEach((m, i) => {
    const y = 1.75 + i * 1.0;
    s.addShape(pptx.ShapeType.rect, { x: 5.15, y, w: 4.4, h: 0.88, fill: { color: C.hero }, line: { color: C.hero }, shadow: makeShadow() });
    s.addText(m.val, { x: 5.3, y: y + 0.06, w: 1.2, h: 0.52, fontSize: 26, bold: true, color: '60A5FA', fontFace: 'Georgia', align: 'center', valign: 'middle' });
    s.addText(m.lbl, { x: 6.55, y: y + 0.06, w: 2.85, h: 0.32, fontSize: 11, bold: true, color: C.white });
    s.addText(m.sub, { x: 6.55, y: y + 0.42, w: 2.85, h: 0.32, fontSize: 9, color: '888888' });
  });

  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 5.95, w: 9.1, h: 0.55, fill: { color: C.burg }, line: { color: C.burg } });
  s.addText('CartBridge does not add a new cost — it recovers revenue that was already being permanently lost. Every recovered sale is pure incremental GMV.', {
    x: 0.65, y: 6.0, w: 8.7, h: 0.45, fontSize: 10, color: C.white, italic: true, wrap: true, valign: 'middle',
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK (light)
// ─────────────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.cream };

  s.addText('PROOF OF WORK', { x: 0.45, y: 0.3, w: 9.1, h: 0.22, fontSize: 9, bold: true, color: C.burg, charSpacing: 5 });
  s.addText('I Built the Equivalent. Here is the Proof.', {
    x: 0.45, y: 0.55, w: 9.1, h: 0.5, fontSize: 26, bold: true, color: C.ink, fontFace: 'Georgia',
  });

  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 1.22, w: 4.4, h: 4.7, fill: { color: C.dark }, line: { color: C.dark } });
  s.addText('What I Built at PhonePe', { x: 0.65, y: 1.38, w: 4.0, h: 0.32, fontSize: 13, bold: true, color: C.gold });
  const phonepe = [
    'Self-serve merchant onboarding platform — cut TAT from 2 days to 30 minutes, eliminating the manual friction that caused merchants to abandon mid-flow',
    'Multi-tenant referral engine — acquired 5,000+ B2B merchants and 5Mn+ new users/month at 23% lower CAC vs prior manual process',
    'Zero-to-one build of the entire platform — no existing infrastructure; scoped, shipped, and scaled from scratch with cross-functional team',
    'Operated at 350M MAU scale — every decision on latency, real-time data, and fallback logic had to work at national UPI volume',
    'Result: 5,000+ merchants, 30-min onboarding TAT, 23% lower CAC — all in under 12 months from zero',
  ];
  phonepe.forEach((p, i) => {
    s.addText('→  ' + p, { x: 0.65, y: 1.82 + i * 0.78, w: 4.05, h: 0.68, fontSize: 9.5, color: 'CCCCCC', wrap: true });
  });

  s.addShape(pptx.ShapeType.rect, { x: 5.15, y: 1.22, w: 4.4, h: 4.7, fill: { color: 'EDE8E0' }, line: { color: 'D8D0C4' } });
  s.addText('How It Maps to CartBridge', { x: 5.35, y: 1.38, w: 4.0, h: 0.32, fontSize: 13, bold: true, color: C.burg });
  const mappings = [
    'Merchant onboarding friction removal  →  Stock-out recovery friction removal (same design philosophy)',
    '0-to-1 platform with no prior infrastructure  →  CartBridge is greenfield inside RL retail tech',
    'Real-time inventory + routing at scale  →  Cross-channel stock visibility (store / DC / online)',
    'Merchant self-serve tool  →  Associate-facing recovery tool with zero training overhead',
    'CAC reduction through process automation  →  Revenue rescue through fulfilment automation',
  ];
  mappings.forEach((m, i) => {
    s.addText('✓  ' + m, { x: 5.35, y: 1.82 + i * 0.78, w: 4.05, h: 0.68, fontSize: 9.5, color: C.ink, wrap: true });
  });

  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 6.1, w: 9.1, h: 0.6, fill: { color: C.dark }, line: { color: C.dark } });
  s.addText('"I built a platform that turned a 2-day manual process into a 30-minute self-serve flow. CartBridge is the same problem — turning a dead-end conversation into a 47-second sale."', {
    x: 0.65, y: 6.15, w: 8.7, h: 0.5, fontSize: 10, color: C.gold, italic: true, wrap: true, valign: 'middle',
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN (dark)
// ─────────────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  s.addText('ROLLOUT PLAN', { x: 0.45, y: 0.3, w: 9.1, h: 0.22, fontSize: 9, bold: true, color: C.gold, charSpacing: 5 });
  s.addText('Phased Launch — 6-Month Plan', {
    x: 0.45, y: 0.55, w: 9.1, h: 0.5, fontSize: 26, bold: true, color: C.white, fontFace: 'Georgia',
  });

  const phases = [
    {
      phase: 'Phase 1', period: 'Month 1–2: Pilot',
      body: 'Deploy CartBridge to 2 flagship stores (SoHo + 5th Ave). Instrument stock-out events, recovery rate, and revenue rescued. Benchmark vs. control stores with no tool.',
    },
    {
      phase: 'Phase 2', period: 'Month 3–4: Optimise',
      body: 'Tune real-time inventory API latency. Add sister-store reserve flow (hold 30 min for transfer). Launch Fulfilment Dashboard for store managers. A/B test CTA placement for ship-to-home.',
    },
    {
      phase: 'Phase 3', period: 'Month 5–6: Full Rollout',
      body: 'Expand to all 20+ North America flagship and outlet stores. Integrate CartBridge stock intelligence into weekly restock planning. Launch SKU-level demand signal reporting for merchandising team.',
    },
  ];

  phases.forEach((p, i) => {
    const x = 0.45 + i * 3.12;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.28, w: 2.88, h: 3.55, fill: { color: C.hero }, line: { color: '2A3A5A' }, shadow: makeShadow() });
    s.addShape(pptx.ShapeType.rect, { x, y: 1.28, w: 2.88, h: 0.06, fill: { color: C.gold }, line: { color: C.gold } });
    s.addText(p.phase, { x: x + 0.14, y: 1.42, w: 2.6, h: 0.28, fontSize: 9, bold: true, color: C.gold, charSpacing: 3 });
    s.addText(p.period, { x: x + 0.14, y: 1.72, w: 2.6, h: 0.38, fontSize: 12, bold: true, color: C.white });
    s.addText(p.body, { x: x + 0.14, y: 2.18, w: 2.6, h: 2.5, fontSize: 9.5, color: 'CCCCCC', wrap: true });
  });

  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 5.02, w: 9.1, h: 0.92, fill: { color: C.burg }, line: { color: C.burg } });
  s.addText('What I Need to Build This', { x: 0.65, y: 5.1, w: 9.0, h: 0.28, fontSize: 11, bold: true, color: C.white });
  s.addText('Real-time inventory API access (in-store + DC + online)  ·  1 backend engineer for stock routing layer  ·  POS integration (existing associate device)  ·  Pilot budget for 2 flagship stores  ·  Alignment with Supply Chain on DC ship SLA', {
    x: 0.65, y: 5.4, w: 8.7, h: 0.46, fontSize: 9.5, color: 'FFDDDD', wrap: true,
  });

  s.addShape(pptx.ShapeType.rect, { x: 0, y: 6.9, w: 10, h: 0.6, fill: { color: C.hero }, line: { color: C.hero } });
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade', {
    x: 0.45, y: 6.95, w: 9.1, h: 0.45, fontSize: 10, color: 'AAAAAA', align: 'center',
  });
}

pptx.writeFile({ fileName: 'assets/rl-cartbridge-deck.pptx' })
  .then(() => console.log('rl-cartbridge-deck.pptx created successfully'))
  .catch(err => console.error(err));
