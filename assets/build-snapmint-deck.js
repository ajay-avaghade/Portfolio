const pptx = require('pptxgenjs');
const prs = new pptx();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0A0F2E',
  hero:   '111827',
  bright: '4F46E5',
  accent: '6366F1',
  gold:   'F59E0B',
  white:  'FFFFFF',
  lgray:  'EEF2FF',
  muted:  '94A3B8',
  ink:    '0F172A',
};

const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });

// ─────────────────────────────────────────────
// SLIDE 1 — COVER
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  // Diagonal stripe pattern
  for(let i = 0; i < 14; i++){
    s.addShape(prs.ShapeType.line, {
      x: -1 + i*0.9, y: 0, w: 0.01, h: 7.5,
      line: { color: 'FFFFFF', width: 0.4, transparency: 92 },
      rotate: 30,
    });
  }

  // Indigo accent block top-right
  s.addShape(prs.ShapeType.rect, {
    x: 7.5, y: 0, w: 2.5, h: 7.5,
    fill: { color: C.bright, transparency: 88 },
    line: { color: C.bright, transparency: 85, width: 1 },
  });

  // Tag
  s.addText('SNAPMINT — PRODUCT CASE STUDY', {
    x: 0.6, y: 0.55, w: 5, h: 0.28,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 2.5,
  });

  // Title
  s.addText('SnapPay:', {
    x: 0.6, y: 1.1, w: 9, h: 1.0,
    fontSize: 54, bold: true, color: C.white, charSpacing: -1,
  });
  s.addText('Credit on UPI', {
    x: 0.6, y: 2.0, w: 9, h: 0.9,
    fontSize: 42, bold: true, color: C.gold,
  });

  // Subtitle
  s.addText('Solving India\'s ₹25,000 Cr credit-access gap at the point of every UPI payment', {
    x: 0.6, y: 3.05, w: 7.2, h: 0.45,
    fontSize: 14, color: C.muted,
  });

  // Presenter
  s.addText('Ajay Avaghade  ·  Growth & Monetization PM', {
    x: 0.6, y: 6.6, w: 5, h: 0.28,
    fontSize: 10, color: C.muted,
  });

  // Big stat
  s.addShape(prs.ShapeType.rect, {
    x: 7.0, y: 4.8, w: 2.7, h: 2.4,
    fill: { color: C.bright, transparency: 82 },
    line: { color: C.bright, transparency: 70, width: 1 },
    rectRadius: 0.18,
    shadow: makeShadow(),
  });
  s.addText('300M', {
    x: 7.0, y: 5.0, w: 2.7, h: 0.8,
    fontSize: 38, bold: true, color: C.white, align: 'center',
  });
  s.addText('credit-eligible Indians\nwith no UPI credit option', {
    x: 7.0, y: 5.8, w: 2.7, h: 0.7,
    fontSize: 9.5, color: C.muted, align: 'center',
  });
}

// ─────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.hero };

  s.addText('THE PROBLEM', {
    x: 0.55, y: 0.38, w: 9, h: 0.25,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 2.5,
  });
  s.addText('300M Indians Have Credit.\nZero Have a Way to Spend It via UPI.', {
    x: 0.55, y: 0.68, w: 9, h: 1.1,
    fontSize: 28, bold: true, color: C.white, lineSpacing: 36,
  });

  // 3 stat columns
  const cols3 = [
    { icon:'🏦', stat:'300M', lbl:'Credit-eligible Indians', sub:'Yet fewer than 35M actively use credit cards — a 265M person gap' },
    { icon:'⛔', stat:'~40%', lbl:'UPI payments fail due to\ninsufficient balance', sub:'Credit exists but cannot reach the point of payment' },
    { icon:'💤', stat:'₹X Cr', lbl:'Idle credit sitting unused\nevery month', sub:'Snapmint\'s approved limits go unspent — no UPI channel to deploy them' },
  ];
  cols3.forEach((col, i) => {
    const x = 0.5 + i * 3.18;
    s.addShape(prs.ShapeType.rect, {
      x, y: 2.1, w: 2.9, h: 3.6,
      fill: { color: 'FFFFFF', transparency: 94 },
      line: { color: C.bright, transparency: 78, width: 1 },
      rectRadius: 0.14,
    });
    s.addText(col.icon, { x, y: 2.25, w: 2.9, h: 0.55, fontSize: 28, align:'center' });
    s.addText(col.stat, {
      x, y: 2.85, w: 2.9, h: 0.72,
      fontSize: 36, bold: true, color: C.gold, align:'center',
    });
    s.addText(col.lbl, {
      x, y: 3.6, w: 2.9, h: 0.55,
      fontSize: 11, bold: true, color: C.white, align:'center',
    });
    s.addText(col.sub, {
      x, y: 4.18, w: 2.9, h: 0.9,
      fontSize: 9.5, color: C.muted, align:'center',
    });
  });

  // Insight box
  s.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 5.95, w: 9, h: 0.98,
    fill: { color: C.bright, transparency: 88 },
    line: { color: C.bright, transparency: 75, width: 1 },
    rectRadius: 0.1,
  });
  s.addText('Root cause: NPCI enabled Credit Line on UPI in 2023. The rails exist. No consumer fintech has built the product on top of them for the 265M credit-eligible, card-absent segment.', {
    x: 0.7, y: 6.05, w: 8.6, h: 0.78,
    fontSize: 10.5, color: C.white, italic: true,
  });
}

// ─────────────────────────────────────────────
// SLIDE 3 — THE INSIGHT
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE INSIGHT', {
    x: 0.55, y: 0.38, w: 9, h: 0.25,
    fontSize: 9, bold: true, color: C.bright, charSpacing: 2.5,
  });
  s.addText('"Insufficient Funds" Is Not a Bank Problem.\nIt\'s a Credit Access Problem.', {
    x: 0.55, y: 0.68, w: 9, h: 1.1,
    fontSize: 26, bold: true, color: C.ink, lineSpacing: 34,
  });

  // Left column — Current state
  s.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 2.05, w: 4.2, h: 3.65,
    fill: { color: 'FEE2E2', transparency: 10 },
    line: { color: 'EF4444', transparency: 70, width: 1 },
    rectRadius: 0.14,
    shadow: makeShadow(),
  });
  s.addText('❌  Today — Insufficient Funds', {
    x: 0.65, y: 2.22, w: 3.9, h: 0.4,
    fontSize: 12, bold: true, color: '991B1B',
  });
  [
    'User scans QR · amount ₹3,200 · bank balance ₹410',
    '"Insufficient Funds" error — payment fails',
    'User leaves the store frustrated',
    'Snapmint\'s ₹25,000 approved limit sits idle',
    'Merchant loses the sale · platform loses GMV',
  ].forEach((pt, i) => {
    s.addText(`• ${pt}`, {
      x: 0.7, y: 2.75 + i * 0.52, w: 3.85, h: 0.45,
      fontSize: 11, color: '7F1D1D',
    });
  });

  // VS divider
  s.addShape(prs.ShapeType.ellipse, {
    x: 4.55, y: 3.45, w: 0.9, h: 0.9,
    fill: { color: C.bright },
    line: { color: C.bright, width: 0 },
    shadow: makeShadow(),
  });
  s.addText('VS', {
    x: 4.55, y: 3.52, w: 0.9, h: 0.76,
    fontSize: 12, bold: true, color: C.white, align:'center',
  });

  // Right column — Proposed
  s.addShape(prs.ShapeType.rect, {
    x: 5.3, y: 2.05, w: 4.2, h: 3.65,
    fill: { color: 'EEF2FF', transparency: 10 },
    line: { color: C.bright, transparency: 65, width: 1 },
    rectRadius: 0.14,
    shadow: makeShadow(),
  });
  s.addText('✅  SnapPay — Credit at the QR', {
    x: 5.45, y: 2.22, w: 3.9, h: 0.4,
    fontSize: 12, bold: true, color: '1E1B4B',
  });
  [
    'Same scan · same ₹410 balance · same ₹3,200 amount',
    'SnapPay intercepts: "Use your Snapmint credit"',
    'User selects 3-month EMI at ₹0 cost · one tap',
    'Payment completes · merchant gets paid instantly',
    'EMI auto-scheduled via UPI AutoPay mandate',
  ].forEach((pt, i) => {
    s.addText(`• ${pt}`, {
      x: 5.5, y: 2.75 + i * 0.52, w: 3.85, h: 0.45,
      fontSize: 11, color: '1E1B4B',
    });
  });

  // Callout
  s.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 5.95, w: 9, h: 0.98,
    fill: { color: C.bright, transparency: 88 },
    line: { color: C.bright, transparency: 75, width: 1 },
    rectRadius: 0.1,
  });
  s.addText('Key insight: SnapPay makes Snapmint\'s idle credit limit an always-on payment instrument — active at every UPI merchant QR in India. No new credit card. No new account. Just credit, finally at the point of payment.', {
    x: 0.7, y: 6.05, w: 8.6, h: 0.78,
    fontSize: 10.5, color: C.ink, italic: true,
  });
}

// ─────────────────────────────────────────────
// SLIDE 4 — THE MECHANIC
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE MECHANIC', {
    x: 0.55, y: 0.38, w: 9, h: 0.25,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 2.5,
  });
  s.addText('How SnapPay Works: Credit at Every QR in India', {
    x: 0.55, y: 0.68, w: 9, h: 0.7,
    fontSize: 26, bold: true, color: C.white,
  });

  const steps = [
    { num:'1', title:'Credit Line Linked', body:'User onboards SnapPay — links bank, sets UPI PIN. Snapmint credit limit mapped to @snapmint UPI handle. NPCI mandate registered.' },
    { num:'2', title:'QR Scan + Shortfall', body:'User scans merchant QR. Payment amount exceeds bank balance. SnapPay intercepts before the "Insufficient Funds" error.' },
    { num:'3', title:'Credit Surfaced', body:'SnapPay bottom sheet: exact shortfall amount, 3/6/9 month EMI options, ₹0 cost display. One-tap selection.' },
    { num:'4', title:'Payment Completes', body:'UPI credit debit settled in real-time via NPCI. Merchant receives full ₹3,200. Transaction confirmed in <2 seconds.' },
    { num:'5', title:'AutoPay Mandate', body:'Post-payment: UPI AutoPay mandate set up for EMI auto-debit. D-3 alerts. Zero manual repayment friction.' },
  ];

  steps.forEach((step, i) => {
    const x = 0.45 + i * 1.86;
    // Circle
    s.addShape(prs.ShapeType.ellipse, {
      x: x + 0.55, y: 1.65, w: 0.72, h: 0.72,
      fill: { color: i === 2 ? C.gold : C.bright },
      line: { color: 'FFFFFF', transparency: 80, width: 1 },
    });
    s.addText(step.num, {
      x: x + 0.55, y: 1.72, w: 0.72, h: 0.58,
      fontSize: 16, bold: true, color: C.white, align:'center',
    });
    // Connector (except last)
    if(i < 4){
      s.addShape(prs.ShapeType.line, {
        x: x + 1.27, y: 2.02, w: 1.0, h: 0,
        line: { color: C.bright, width: 1.5, transparency: 55, dashType:'dash' },
      });
    }
    // Card
    s.addShape(prs.ShapeType.rect, {
      x, y: 2.6, w: 1.74, h: 3.4,
      fill: { color: 'FFFFFF', transparency: 93 },
      line: { color: C.bright, transparency: 72, width: 1 },
      rectRadius: 0.12,
    });
    s.addText(step.title, {
      x: x + 0.08, y: 2.75, w: 1.58, h: 0.55,
      fontSize: 11, bold: true, color: C.white, align:'center',
    });
    s.addText(step.body, {
      x: x + 0.08, y: 3.35, w: 1.58, h: 2.45,
      fontSize: 9.5, color: C.muted, align:'center',
    });
  });

  // A/B proof note
  s.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.2, w: 9, h: 0.65,
    fill: { color: C.gold, transparency: 90 },
    line: { color: C.gold, transparency: 75, width: 1 },
    rectRadius: 0.08,
  });
  s.addText('NPCI precedent: RuPay Credit Card on UPI enabled Sept 2023 · PhonePe, GPay integrating — Snapmint can capture the credit-native segment before big players fully optimise for it', {
    x: 0.7, y: 6.28, w: 8.6, h: 0.5,
    fontSize: 10, color: C.gold, italic: true,
  });
}

// ─────────────────────────────────────────────
// SLIDE 5 — THE PRODUCT
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE PRODUCT', {
    x: 0.55, y: 0.38, w: 9, h: 0.25,
    fontSize: 9, bold: true, color: C.bright, charSpacing: 2.5,
  });
  s.addText('4 Screen States — Designed for India\'s Next Credit User', {
    x: 0.55, y: 0.68, w: 9, h: 0.55,
    fontSize: 24, bold: true, color: C.ink,
  });

  const screens = [
    { num:'01', title:'UPI Credit Setup',  desc:'Snapmint home with SnapPay activation banner. 3-step progress: credit approved → bank link → UPI live. Credit limit displayed as always-available payment instrument.', mock:'⬛ ▓ ▓ ▓\n━━━━━━━━\n◉ Setup\n▓▓▓▓░░' },
    { num:'02', title:'Credit at the QR', desc:'Insufficient funds intercept: merchant name, amount, bank balance. SnapPay bottom sheet with EMI options (3/6/9 months). Zero-cost label prominent. One-tap payment CTA.', mock:'⬛ ₹3,200\n⚠ Bal ₹410\n━━━━━━━━\n💳 SnapPay\n[Pay Now]' },
    { num:'03', title:'Payment Confirmed', desc:'Confetti celebration. Full receipt: merchant, amount, EMI plan, zero cost confirmation, UPI ref. EMI schedule (3 upcoming dates). AutoPay setup CTA.', mock:'✅ Paid!\n₹3,200\n3 × ₹1,067\n₹0 cost\n🔁 AutoPay' },
    { num:'04', title:'AutoPay Mandate',   desc:'NPCI UPI AutoPay mandate details: bank, auto-debit amount, first debit date, total payments. EMI calendar with 3 upcoming dates. Pre-debit alert (D-3) confirmation.', mock:'🔁 AutoPay\n₹1,067/mo\n15 Jun ●\n15 Jul ○\n15 Aug ○' },
  ];

  screens.forEach((sc, i) => {
    const x = 0.4 + i * 2.32;
    s.addShape(prs.ShapeType.rect, {
      x, y: 1.45, w: 2.12, h: 5.55,
      fill: { color: C.white },
      line: { color: C.bright, transparency: 72, width: 1 },
      rectRadius: 0.14,
      shadow: makeShadow(),
    });
    s.addText(sc.num, {
      x: x + 0.1, y: 1.6, w: 0.5, h: 0.32,
      fontSize: 10, bold: true, color: C.bright,
    });
    s.addText(sc.title, {
      x: x + 0.08, y: 1.92, w: 1.96, h: 0.45,
      fontSize: 11.5, bold: true, color: C.ink,
    });
    // Mock block
    s.addShape(prs.ShapeType.rect, {
      x: x + 0.12, y: 2.42, w: 1.88, h: 1.55,
      fill: { color: C.dark },
      rectRadius: 0.08,
    });
    s.addText(sc.mock, {
      x: x + 0.14, y: 2.48, w: 1.84, h: 1.43,
      fontSize: 8.5, color: C.white, align:'center',
    });
    s.addText(sc.desc, {
      x: x + 0.1, y: 4.05, w: 1.92, h: 2.72,
      fontSize: 9, color: C.muted,
    });
  });

  s.addText('Live prototype: snapmint-snappay-prototype.html — interactive demo of all 4 screen states', {
    x: 0.55, y: 7.12, w: 9, h: 0.25,
    fontSize: 9, color: C.muted, italic: true,
  });
}

// ─────────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.hero };

  s.addText('IMPACT & ROI', {
    x: 0.55, y: 0.38, w: 9, h: 0.25,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 2.5,
  });
  s.addText('Projected Impact — Built on NPCI\'s Credit Line on UPI Rails', {
    x: 0.55, y: 0.68, w: 9, h: 0.6,
    fontSize: 24, bold: true, color: C.white,
  });

  const userMetrics = [
    { val:'↑ 40%', lbl:'Payment completion rate at insufficient-balance moments', sub:'vs 0% today — credit closes the gap' },
    { val:'↑ 25%', lbl:'Monthly credit utilisation per active Snapmint customer', sub:'idle limit converted to active spend' },
    { val:'↓ 30%', lbl:'EMI repayment failures via AutoPay mandate', sub:'vs manual repayment baseline' },
    { val:'3x',    lbl:'GMV per user unlocked vs bank-balance-only UPI', sub:'credit multiplies purchasing power' },
  ];
  const bizMetrics = [
    { val:'265M',    lbl:'Addressable users — credit-eligible, card-absent', sub:'Snapmint\'s unique distribution advantage' },
    { val:'~0%',     lbl:'Incremental credit risk on NPCI-settled transactions', sub:'UPI debit settled T+0; EMI risk already priced' },
    { val:'₹ 0 CAC', lbl:'UPI adoption cost for existing 10M Snapmint base', sub:'product-led, no new acquisition spend' },
    { val:'↑ 60%',   lbl:'LTV uplift per customer via higher credit utilisation', sub:'more transactions = stronger repayment data' },
  ];

  const sectionLabel = (txt, x) => s.addText(txt, {
    x, y: 1.55, w: 4.3, h: 0.28,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 1.5,
  });
  sectionLabel('USER IMPACT', 0.45);
  sectionLabel('SNAPMINT ROI', 5.05);

  [[userMetrics, 0.45],[bizMetrics, 5.05]].forEach(([metrics, x]) => {
    metrics.forEach((m, i) => {
      const y = 1.95 + i * 1.18;
      s.addShape(prs.ShapeType.rect, {
        x, y, w: 4.3, h: 1.08,
        fill: { color: 'FFFFFF', transparency: 94 },
        line: { color: C.bright, transparency: 78, width: 1 },
        rectRadius: 0.1,
      });
      s.addText(m.val, {
        x: x + 0.15, y: y + 0.08, w: 1.3, h: 0.55,
        fontSize: 26, bold: true, color: C.gold,
      });
      s.addText(m.lbl, {
        x: x + 1.5, y: y + 0.08, w: 2.65, h: 0.42,
        fontSize: 10.5, bold: true, color: C.white,
      });
      s.addText(m.sub, {
        x: x + 1.5, y: y + 0.54, w: 2.65, h: 0.42,
        fontSize: 9, color: C.muted,
      });
    });
  });

  // Divider
  s.addShape(prs.ShapeType.line, {
    x: 4.93, y: 1.55, w: 0, h: 5.1,
    line: { color: C.bright, width: 1, transparency: 75, dashType:'dash' },
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.45, y: 6.62, w: 9.1, h: 0.72,
    fill: { color: C.bright, transparency: 88 },
    line: { color: C.bright, transparency: 75, width: 1 },
    rectRadius: 0.08,
  });
  s.addText('SnapPay turns Snapmint\'s existing approved credit into an active payment instrument — no new lending risk, no new CAC, just the UPI rails NPCI already built.', {
    x: 0.65, y: 6.7, w: 8.8, h: 0.56,
    fontSize: 10, color: C.white, italic: true,
  });
}

// ─────────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('PROOF OF WORK', {
    x: 0.55, y: 0.38, w: 9, h: 0.25,
    fontSize: 9, bold: true, color: C.bright, charSpacing: 2.5,
  });
  s.addText('I Built Inside India\'s Largest UPI App. Here\'s the Evidence.', {
    x: 0.55, y: 0.68, w: 9, h: 0.6,
    fontSize: 24, bold: true, color: C.ink,
  });

  // Left — dark proof column
  s.addShape(prs.ShapeType.rect, {
    x: 0.45, y: 1.5, w: 4.3, h: 5.45,
    fill: { color: C.dark },
    rectRadius: 0.16,
    shadow: makeShadow(),
  });
  s.addText('PhonePe — What I Actually Built', {
    x: 0.65, y: 1.68, w: 3.9, h: 0.4,
    fontSize: 12, bold: true, color: C.gold,
  });
  [
    'Built products on top of UPI at 350M MAU — PhonePe processes 50%+ of India\'s UPI volume',
    'Deployed Propensity-to-Transact ML on ₹1,000+ Cr/yr payment marketing — real-time scoring at UPI transaction level',
    'Built self-serve PG/PSP integration platform — 5,000+ B2B merchants, 30-min onboarding, 5Mn+ users/month',
    'Designed cart incentivisation engine with context-aware triggers (cart × intent × time) — 60% abandonment reduction',
    'Worked alongside NPCI integration, UPI mandate, and compliance teams as an embedded product stakeholder',
  ].forEach((pt, i) => {
    s.addText(`• ${pt}`, {
      x: 0.65, y: 2.22 + i * 0.78, w: 3.9, h: 0.72,
      fontSize: 9.5, color: C.muted,
    });
  });

  // Right — JD mapping column
  s.addShape(prs.ShapeType.rect, {
    x: 5.05, y: 1.5, w: 4.45, h: 5.45,
    fill: { color: C.white },
    line: { color: C.bright, transparency: 72, width: 1 },
    rectRadius: 0.16,
    shadow: makeShadow(),
  });
  s.addText('Snapmint JD — How It Maps', {
    x: 5.22, y: 1.68, w: 4.1, h: 0.4,
    fontSize: 12, bold: true, color: C.bright,
  });
  [
    'Build UPI from scratch → 0→1 product launches: Pincode (quick commerce), Kotak Cherry (WealthTech founding team)',
    'PSP/bank integrations → Self-serve PG integration platform (5,000+ merchants, real PSP partnerships)',
    'Drive UPI adoption & engagement → ML-powered growth: 32% burn reduction, 22% conversion lift',
    'UPI mandates & compliance → Embedded with NPCI mandate flows; deep regulatory context from PhonePe inside',
    'Data-driven, execution-focused → All results metrics-backed; shipped, not theorised',
  ].forEach((pt, i) => {
    s.addText(`• ${pt}`, {
      x: 5.22, y: 2.22 + i * 0.78, w: 4.1, h: 0.72,
      fontSize: 9.5, color: C.ink,
    });
  });

  // Quote
  s.addShape(prs.ShapeType.rect, {
    x: 0.45, y: 7.12, w: 9.1, h: 0.62,
    fill: { color: C.bright, transparency: 90 },
    line: { color: C.bright, transparency: 78, width: 1 },
    rectRadius: 0.08,
  });
  s.addText('"This isn\'t a concept. I built the infrastructure, the growth mechanics, and the merchant integrations this role needs — from inside the UPI ecosystem that processes ₹20L Cr/month."', {
    x: 0.65, y: 7.18, w: 8.8, h: 0.5,
    fontSize: 9.5, color: C.ink, italic: true,
  });
}

// ─────────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('ROLLOUT PLAN', {
    x: 0.55, y: 0.38, w: 9, h: 0.25,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 2.5,
  });
  s.addText('From NPCI Registration to 10M Users — A Phased Plan', {
    x: 0.55, y: 0.68, w: 9, h: 0.6,
    fontSize: 24, bold: true, color: C.white,
  });

  const phases = [
    {
      phase:'PHASE 1', period:'Month 1–2', title:'Foundation + PSP',
      points:[
        'Register Snapmint as TPAP with NPCI for Credit Line on UPI',
        'Partner with PSP (Razorpay / Cashfree) for UPI credit settlement',
        'Build core SnapPay setup flow + UPI handle allocation',
        'Internal pilot: 500 Snapmint employees',
      ],
    },
    {
      phase:'PHASE 2', period:'Month 3–4', title:'Pilot + QR Intercept',
      points:[
        'Launch QR-scan intercept for insufficient-balance scenarios',
        'Pilot with 10,000 existing Snapmint customers in 2 cities',
        'Instrument: payment completion rate, EMI selection, AutoPay setup',
        'A/B test: intercept timing — pre-payment vs post-failure',
      ],
    },
    {
      phase:'PHASE 3', period:'Month 5–6', title:'Scale to 10M',
      points:[
        'Roll out to full 10M Snapmint customer base',
        'Activate merchant QR SDK for Snapmint partner network',
        'Launch D-3 AutoPay nudge engine via push + UPI notification',
        'Expand: UPI-first credit acquisition for new user segment',
      ],
    },
  ];

  phases.forEach((ph, i) => {
    const x = 0.45 + i * 3.15;
    s.addShape(prs.ShapeType.rect, {
      x, y: 1.55, w: 2.9, h: 4.55,
      fill: { color: 'FFFFFF', transparency: 94 },
      line: { color: i === 1 ? C.gold : C.bright, transparency: i === 1 ? 60 : 76, width: i === 1 ? 1.5 : 1 },
      rectRadius: 0.14,
    });
    s.addText(ph.phase, {
      x: x + 0.12, y: 1.7, w: 2.66, h: 0.26,
      fontSize: 9, bold: true, color: i === 1 ? C.gold : C.muted, charSpacing: 1.5,
    });
    s.addText(ph.period, {
      x: x + 0.12, y: 1.96, w: 2.66, h: 0.3,
      fontSize: 11, bold: true, color: C.white,
    });
    s.addText(ph.title, {
      x: x + 0.12, y: 2.28, w: 2.66, h: 0.35,
      fontSize: 13, bold: true, color: i === 1 ? C.gold : C.accent,
    });
    ph.points.forEach((pt, j) => {
      s.addText(`• ${pt}`, {
        x: x + 0.12, y: 2.75 + j * 0.72, w: 2.66, h: 0.66,
        fontSize: 9.5, color: C.muted,
      });
    });
  });

  // Ask box
  s.addShape(prs.ShapeType.rect, {
    x: 0.45, y: 6.3, w: 9.1, h: 0.85,
    fill: { color: C.bright, transparency: 86 },
    line: { color: C.bright, transparency: 72, width: 1 },
    rectRadius: 0.1,
  });
  s.addText('What I need to build this:', {
    x: 0.65, y: 6.38, w: 2.2, h: 0.28,
    fontSize: 10, bold: true, color: C.white,
  });
  s.addText('NPCI TPAP registration support  ·  PSP partnership (Razorpay / Cashfree)  ·  1 backend engineer for UPI credit flow  ·  Access to Snapmint credit limit API', {
    x: 0.65, y: 6.66, w: 8.7, h: 0.35,
    fontSize: 9.5, color: C.muted,
  });

  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439', {
    x: 0.45, y: 7.22, w: 9, h: 0.24,
    fontSize: 9, color: C.muted, align:'center',
  });
}

// ─────────────────────────────────────────────
// WRITE FILE
// ─────────────────────────────────────────────
prs.writeFile({ fileName: 'assets/snapmint-snappay-deck.pptx' })
  .then(() => console.log('✅  Deck written: snapmint-snappay-deck.pptx'))
  .catch(e => { console.error(e); process.exit(1); });
