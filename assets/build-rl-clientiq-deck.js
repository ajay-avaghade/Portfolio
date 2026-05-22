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

  // Diagonal lines motif (top-right)
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
  s.addText('RL ClientIQ', { x: 0.65, y: 1.4, w: 6.5, h: 0.85, fontSize: 46, bold: true, color: C.white, fontFace: 'Georgia' });
  s.addText('Turning Every Associate Into a Personal Stylist', { x: 0.65, y: 2.22, w: 6.8, h: 0.48, fontSize: 19, color: C.gold, fontFace: 'Georgia', italic: true });

  // Subtitle block
  s.addText('Presented by Ajay Avaghade  ·  Growth & Monetization PM', {
    x: 0.65, y: 2.85, w: 6.0, h: 0.3, fontSize: 11, color: 'AAAAAA',
  });

  // Big stat — bottom right
  s.addShape(pptx.ShapeType.rect, { x: 7.2, y: 5.6, w: 2.5, h: 1.6, fill: { color: C.burg }, line: { color: C.burg }, rounding: true });
  s.addText('28%', { x: 7.2, y: 5.75, w: 2.5, h: 0.65, fontSize: 44, bold: true, color: C.white, align: 'center', fontFace: 'Georgia' });
  s.addText('avg basket uplift\nwith ClientIQ', { x: 7.2, y: 6.38, w: 2.5, h: 0.65, fontSize: 10, color: 'DDDDDD', align: 'center' });

  // Bottom company line
  s.addText('Ralph Lauren Corporation  ·  Clienteling & Retail Technology', {
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
  s.addText('Associates Fly Blind — and High-Value Customers Walk Out', {
    x: 0.45, y: 0.55, w: 9.1, h: 0.7, fontSize: 26, bold: true, color: C.white, fontFace: 'Georgia',
  });

  // 3 stat columns
  const cols = [
    { icon: '👁', stat: '0%', lbl: 'Purchase history visible at POS', sub: 'Associates greet VIP customers with zero context' },
    { icon: '🛍', stat: '$0', lbl: 'Wishlist surfaced during a visit', sub: 'Items customers want sit idle in digital lists, never shown in-store' },
    { icon: '📉', stat: '60%', lbl: 'Cross-sell missed in-store', sub: 'Associates recommend on intuition — not on the customer\'s actual style profile' },
  ];
  cols.forEach((col, i) => {
    const x = 0.45 + i * 3.1;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.55, w: 2.85, h: 3.2, fill: { color: C.hero }, line: { color: C.hero }, shadow: makeShadow() });
    s.addText(col.icon, { x, y: 1.7, w: 2.85, h: 0.5, fontSize: 26, align: 'center' });
    s.addText(col.stat, { x, y: 2.22, w: 2.85, h: 0.65, fontSize: 40, bold: true, color: C.gold, align: 'center', fontFace: 'Georgia' });
    s.addText(col.lbl, { x: x + 0.12, y: 2.88, w: 2.6, h: 0.6, fontSize: 11, bold: true, color: C.white, align: 'center', wrap: true });
    s.addText(col.sub, { x: x + 0.12, y: 3.5, w: 2.6, h: 0.8, fontSize: 9.5, color: 'AAAAAA', align: 'center', wrap: true });
  });

  // Insight box
  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 5.0, w: 9.1, h: 0.95, fill: { color: C.burg }, line: { color: C.burg } });
  s.addText('Root cause: The associate relationship is the brand\'s greatest asset — but it is entirely context-free. ClientIQ puts every customer\'s LTV, purchase history, size profile, and wishlist into the associate\'s hands at the moment of greeting.', {
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
  s.addText('Generic Service Burns Relationships. Context-Driven Service Builds Them.', {
    x: 0.45, y: 0.55, w: 9.1, h: 0.7, fontSize: 22, bold: true, color: C.ink, fontFace: 'Georgia',
  });

  // Left column — Current
  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 1.55, w: 4.0, h: 3.6, fill: { color: 'F0EAE0' }, line: { color: 'D8D0C4' } });
  s.addText('Without ClientIQ', { x: 0.6, y: 1.7, w: 3.7, h: 0.32, fontSize: 13, bold: true, color: C.burg });
  const badRows = [
    'Associate has no customer context at greeting',
    'Recommends by gut feel — misses actual size and style',
    'Wishlist items sit unseen — customer leaves without them',
    'Lost sale is invisible — no recovery mechanism',
    'VIP customer gets the same experience as a walk-in',
  ];
  badRows.forEach((r, i) => {
    s.addText('✕  ' + r, { x: 0.62, y: 2.12 + i * 0.48, w: 3.6, h: 0.42, fontSize: 10, color: '8B0000', wrap: true });
  });

  // VS divider
  s.addShape(pptx.ShapeType.ellipse, { x: 4.52, y: 3.0, w: 0.6, h: 0.6, fill: { color: C.dark }, line: { color: C.dark } });
  s.addText('VS', { x: 4.52, y: 3.0, w: 0.6, h: 0.6, fontSize: 9, bold: true, color: C.white, align: 'center', valign: 'middle' });

  // Right column — Proposed
  s.addShape(pptx.ShapeType.rect, { x: 5.2, y: 1.55, w: 4.35, h: 3.6, fill: { color: C.dark }, line: { color: C.dark } });
  s.addText('With ClientIQ', { x: 5.35, y: 1.7, w: 4.0, h: 0.32, fontSize: 13, bold: true, color: C.gold });
  const goodRows = [
    'Customer 360 loaded on QR scan — LTV, history, sizes',
    'Recommends from actual wishlist + style profile on file',
    'In-store wishlist lookup + instant stock check by size',
    'Ship-to-home closes the sale even if stock is out',
    'VIP customers receive the personalised service they deserve',
  ];
  goodRows.forEach((r, i) => {
    s.addText('✓  ' + r, { x: 5.35, y: 2.12 + i * 0.48, w: 4.0, h: 0.42, fontSize: 10, color: 'A8D5A2', wrap: true });
  });

  // Callout
  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 5.38, w: 9.1, h: 0.62, fill: { color: C.dark }, line: { color: C.dark } });
  s.addText('Key insight: The most loyal Ralph Lauren customers already share their style data through digital channels — ClientIQ is simply the bridge that surfaces it at the point of human connection.', {
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
  s.addText('How ClientIQ Works — Five Moments That Change the Sale', {
    x: 0.45, y: 0.55, w: 9.1, h: 0.55, fontSize: 24, bold: true, color: C.white, fontFace: 'Georgia',
  });

  const steps = [
    { num: '01', title: 'Customer Arrives', body: 'Associate opens ClientIQ. Scans QR code from loyalty card or app. Customer 360 loads in under 2 seconds — no login required.' },
    { num: '02', title: 'Instant Context', body: 'LTV, last 3 purchases, saved sizes, wishlist items, and NPS score all visible. Associate greets by name with genuine product intelligence.' },
    { num: '03', title: 'Wishlist Lookup', body: 'Associate surfaces wishlist items and checks in-store stock by exact size. FIND button shows aisle, rack, floor — or triggers ship-to-home.' },
    { num: '04', title: 'Product Match', body: 'AI suggests complementary items based on purchase history and current season. Associate can pull items for a fitting room with a single tap.' },
    { num: '05', title: 'Sale Closed', body: 'Transaction logged back to customer profile. NPS prompt sent 24h later. Manager dashboard updates basket uplift and loyalty enrollment KPIs.' },
  ];

  // Dashed connector line
  s.addShape(pptx.ShapeType.line, { x: 0.88, y: 1.6, w: 0, h: 4.5, line: { color: C.gold, width: 1, dashType: 'dash', transparency: 50 } });

  steps.forEach((step, i) => {
    const y = 1.38 + i * 0.92;
    s.addShape(pptx.ShapeType.ellipse, { x: 0.6, y: y + 0.05, w: 0.55, h: 0.55, fill: { color: C.gold }, line: { color: C.gold } });
    s.addText(step.num, { x: 0.6, y: y + 0.05, w: 0.55, h: 0.55, fontSize: 9, bold: true, color: C.dark, align: 'center', valign: 'middle' });
    s.addText(step.title, { x: 1.3, y: y, w: 3.6, h: 0.3, fontSize: 12, bold: true, color: C.white });
    s.addText(step.body, { x: 1.3, y: y + 0.3, w: 8.1, h: 0.52, fontSize: 9.5, color: 'BBBBBB', wrap: true });
  });

  // A/B proof note
  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 6.18, w: 9.1, h: 0.55, fill: { color: C.hero }, line: { color: C.hero } });
  s.addText('PhonePe proof: ML propensity + real-time user profiling at 350M MAU scale — same core mechanic as ClientIQ, applied to digital payments. Drove 22% checkout conversion lift.', {
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
  s.addText('4 Screens — Built and Validated', {
    x: 0.45, y: 0.55, w: 9.1, h: 0.5, fontSize: 26, bold: true, color: C.ink, fontFace: 'Georgia',
  });

  const cards = [
    { num: '01', title: 'Associate Home', desc: 'Daily summary: visitor queue, basket target, today\'s appointments. Persistent greeting card surfaces VIP context before the customer enters.' },
    { num: '02', title: 'Customer 360', desc: 'Full profile on scan: LTV, order count, NPS, size profile chips, 3-item purchase history, 2-item wishlist with FIND buttons.' },
    { num: '03', title: 'Product Match', desc: 'Wishlist item card with real-time availability across in-store, sister stores, DC, and online. Retrieve-in-store and ship-to-home action buttons.' },
    { num: '04', title: 'Manager Dashboard', desc: 'Store ops view: basket uplift, ClientIQ usage %, ship-to-home rate, loyalty enrollment. Associate leaderboard with revenue impact and adoption nudges.' },
  ];

  cards.forEach((card, i) => {
    const x = 0.4 + i * 2.38;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.28, w: 2.18, h: 4.55, fill: { color: C.dark }, line: { color: C.dark }, shadow: makeShadow() });
    s.addShape(pptx.ShapeType.rect, { x, y: 1.28, w: 2.18, h: 0.05, fill: { color: C.gold }, line: { color: C.gold } });

    // Screen number
    s.addText(card.num, { x: x + 0.12, y: 1.4, w: 0.5, h: 0.3, fontSize: 9, bold: true, color: C.gold, fontFace: 'Courier New' });
    s.addText(card.title, { x: x + 0.12, y: 1.7, w: 1.92, h: 0.5, fontSize: 12, bold: true, color: C.white, wrap: true });

    // ASCII mockup block
    s.addShape(pptx.ShapeType.rect, { x: x + 0.12, y: 2.28, w: 1.92, h: 2.05, fill: { color: C.hero }, line: { color: '2A3A5A' } });
    const mockups = [
      '┌──────────────┐\n│ Sophie — SoHo │\n│ 14 visitors  │\n│ $4.2K basket  │\n│ 3 appts today│\n└──────────────┘',
      '┌──────────────┐\n│ Alexandra Chen│\n│ PURPLE LABEL  │\n│ LTV $12,847   │\n│ 47 orders     │\n│ Size: US 4    │\n└──────────────┘',
      '┌──────────────┐\n│ Cashmere V-Neck│\n│ $695          │\n│ ✓ In store    │\n│   Floor 2, B4 │\n│ [Retrieve]    │\n└──────────────┘',
      '┌──────────────┐\n│ +28% basket   │\n│ 84% usage     │\n│               │\n│ Sophie  +$1.8K│\n│ Marco   +$1.2K│\n└──────────────┘',
    ];
    s.addText(mockups[i], { x: x + 0.15, y: 2.35, w: 1.86, h: 1.9, fontSize: 7, color: '88AACC', fontFace: 'Courier New', wrap: true, valign: 'top' });

    s.addText(card.desc, { x: x + 0.1, y: 4.42, w: 1.98, h: 1.3, fontSize: 9, color: 'BBBBBB', wrap: true });
  });

  s.addText('Interactive prototype: rl-clientiq-prototype.html  ·  All 4 screen states live', {
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

  // Left: Associate / Customer
  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 1.3, w: 4.4, h: 0.3, fill: { color: C.burg }, line: { color: C.burg } });
  s.addText('CUSTOMER & ASSOCIATE IMPACT', { x: 0.45, y: 1.3, w: 4.4, h: 0.3, fontSize: 8, bold: true, color: C.white, align: 'center', valign: 'middle', charSpacing: 3 });

  const leftMetrics = [
    { val: '+28%', lbl: 'Avg basket size per visit', sub: 'vs. visits without ClientIQ' },
    { val: '+34%', lbl: 'Wishlist-to-purchase conversion', sub: 'Items surfaced in-store vs. dormant' },
    { val: '3×', lbl: 'Loyalty enrollment rate', sub: 'Associate-prompted vs. self-serve' },
    { val: '−40%', lbl: 'Lost sales from stock-out', sub: 'Ship-to-home closes the gap' },
  ];
  leftMetrics.forEach((m, i) => {
    const y = 1.75 + i * 1.0;
    s.addShape(pptx.ShapeType.rect, { x: 0.45, y, w: 4.4, h: 0.88, fill: { color: C.hero }, line: { color: C.hero }, shadow: makeShadow() });
    s.addText(m.val, { x: 0.6, y: y + 0.06, w: 1.1, h: 0.52, fontSize: 30, bold: true, color: C.gold, fontFace: 'Georgia', align: 'center', valign: 'middle' });
    s.addText(m.lbl, { x: 1.75, y: y + 0.06, w: 2.95, h: 0.32, fontSize: 11, bold: true, color: C.white });
    s.addText(m.sub, { x: 1.75, y: y + 0.42, w: 2.95, h: 0.32, fontSize: 9, color: '888888' });
  });

  // Right: Company ROI
  s.addShape(pptx.ShapeType.rect, { x: 5.15, y: 1.3, w: 4.4, h: 0.3, fill: { color: C.hero }, line: { color: '2A3A5A' } });
  s.addText('RALPH LAUREN — BUSINESS ROI', { x: 5.15, y: 1.3, w: 4.4, h: 0.3, fontSize: 8, bold: true, color: C.gold, align: 'center', valign: 'middle', charSpacing: 3 });

  const rightMetrics = [
    { val: '+$420', lbl: 'Revenue per VIP store visit', sub: 'Purple Label / Collection tiers' },
    { val: '84%', lbl: 'Associate adoption rate (week 4)', sub: 'PhonePe parallel: ML tool adoption' },
    { val: '12%', lbl: 'Ship-to-home orders as share of visits', sub: 'Zero lost sales from stockout' },
    { val: '18mo', lbl: 'Payback period on dev investment', sub: 'Basket uplift alone covers build cost' },
  ];
  rightMetrics.forEach((m, i) => {
    const y = 1.75 + i * 1.0;
    s.addShape(pptx.ShapeType.rect, { x: 5.15, y, w: 4.4, h: 0.88, fill: { color: C.hero }, line: { color: C.hero }, shadow: makeShadow() });
    s.addText(m.val, { x: 5.3, y: y + 0.06, w: 1.2, h: 0.52, fontSize: 28, bold: true, color: '60A5FA', fontFace: 'Georgia', align: 'center', valign: 'middle' });
    s.addText(m.lbl, { x: 6.55, y: y + 0.06, w: 2.85, h: 0.32, fontSize: 11, bold: true, color: C.white });
    s.addText(m.sub, { x: 6.55, y: y + 0.42, w: 2.85, h: 0.32, fontSize: 9, color: '888888' });
  });

  // Insight
  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 5.95, w: 9.1, h: 0.55, fill: { color: C.burg }, line: { color: C.burg } });
  s.addText('ClientIQ turns every in-store visit from a single transaction into a relationship event — compounding LTV across every subsequent visit.', {
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
  s.addText('I Built This. Here\'s the Parallel.', {
    x: 0.45, y: 0.55, w: 9.1, h: 0.5, fontSize: 26, bold: true, color: C.ink, fontFace: 'Georgia',
  });

  // Left: PhonePe context (dark)
  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 1.22, w: 4.4, h: 4.7, fill: { color: C.dark }, line: { color: C.dark } });
  s.addText('What I Built at PhonePe', { x: 0.65, y: 1.38, w: 4.0, h: 0.32, fontSize: 13, bold: true, color: C.gold });
  const phonepe = [
    'Propensity-to-Transact ML models scoring 350M+ users in real time — replacing manual cohort selection with user-level targeting',
    'End-to-end payment offers surfacing + redemption UX redesign — cross-functional delivery across Tech, UX, Finance, Legal',
    'Dynamic cart incentivisation engine for Pincode Q-commerce — context-aware triggers across cart value, user intent, time signals',
    '500+ brand partners onboarded on self-serve campaign platform — audience segmentation + ML-driven marketplace tooling',
    'Result: 22% checkout conversion lift, 32% marketing burn reduction, 35% AOV uplift, 26% user engagement increase',
  ];
  phonepe.forEach((p, i) => {
    s.addText('→  ' + p, { x: 0.65, y: 1.82 + i * 0.78, w: 4.05, h: 0.68, fontSize: 9.5, color: 'CCCCCC', wrap: true });
  });

  // Right: Mapping to RL JD
  s.addShape(pptx.ShapeType.rect, { x: 5.15, y: 1.22, w: 4.4, h: 4.7, fill: { color: 'EDE8E0' }, line: { color: 'D8D0C4' } });
  s.addText('How It Maps to This Role', { x: 5.35, y: 1.38, w: 4.0, h: 0.32, fontSize: 13, bold: true, color: C.burg });
  const mappings = [
    'User-level ML targeting at scale  →  ClientIQ customer 360 personalisation engine',
    'Cross-functional PM (Tech/UX/Finance/Legal)  →  RL product org (Tech, Store Ops, Loyalty, CRM)',
    'Dynamic incentivisation engine  →  Real-time wishlist + product match recommendation surface',
    'Self-serve brand campaign platform  →  Associate-facing clienteling tool with adoption tracking',
    'Checkout conversion + GMV metrics  →  Basket uplift, loyalty enrollment, ship-to-home rate',
  ];
  mappings.forEach((m, i) => {
    s.addText('✓  ' + m, { x: 5.35, y: 1.82 + i * 0.78, w: 4.05, h: 0.68, fontSize: 9.5, color: C.ink, wrap: true });
  });

  // Closing quote
  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 6.1, w: 9.1, h: 0.6, fill: { color: C.dark }, line: { color: C.dark } });
  s.addText('"The data infrastructure for ClientIQ already exists inside Ralph Lauren — I know how to wire it to the associate experience because I built the equivalent at scale."', {
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
      body: 'Deploy ClientIQ to 10 associates across 2 flagship stores (SoHo + 5th Ave). Instrument basket uplift, wishlist conversion, and loyalty enrollment. Benchmark vs. control group associates without ClientIQ.',
    },
    {
      phase: 'Phase 2', period: 'Month 3–4: Scale + Iterate',
      body: 'Roll out to all 50+ associates across pilot stores. Tune ML recommendation model on real purchase data. Launch Manager Dashboard with adoption leaderboard. A/B test ship-to-home CTA placement.',
    },
    {
      phase: 'Phase 3', period: 'Month 5–6: Full Rollout',
      body: 'Expand to all North America flagship stores. Integrate with loyalty CRM and RFID inventory systems. Launch push-nudge engine for associates approaching basket targets.',
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

  // Ask box
  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 5.02, w: 9.1, h: 0.92, fill: { color: C.burg }, line: { color: C.burg } });
  s.addText('What I Need to Build This', { x: 0.65, y: 5.1, w: 9.0, h: 0.28, fontSize: 11, bold: true, color: C.white });
  s.addText('Access to loyalty CRM and purchase history API  ·  1 iOS engineer for associate app  ·  1 UX designer  ·  Alignment with Store Ops on associate workflow  ·  Pilot budget for 2 flagship stores', {
    x: 0.65, y: 5.4, w: 8.7, h: 0.46, fontSize: 9.5, color: 'FFDDDD', wrap: true,
  });

  // Contact footer
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 6.9, w: 10, h: 0.6, fill: { color: C.hero }, line: { color: C.hero } });
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade', {
    x: 0.45, y: 6.95, w: 9.1, h: 0.45, fontSize: 10, color: 'AAAAAA', align: 'center',
  });
}

pptx.writeFile({ fileName: 'assets/rl-clientiq-deck.pptx' })
  .then(() => console.log('rl-clientiq-deck.pptx created successfully'))
  .catch(err => console.error(err));
