const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '07132B', // Almosafer deep navy
  hero:   '0C2240', // Rich dark slate
  bright: 'FF5A00', // Almosafer Orange
  gold:   'F5A623', // Gold Accent
  green:  '10B981', // Emerald green
  white:  'FFFFFF',
  lgray:  'F1F5F9',
  muted:  '8B92A5',
};

const makeShadow = () => ({
  type: 'outer', color: '000000', blur: 4, offset: 1, angle: 45, opacity: 0.14
});

// SLIDE 1: Cover
let slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: '25%', h: '100%', fill: { color: C.bright } });
slide.addText('ALMOSAFER TECH', { x: 0.5, y: 0.5, w: 2.5, fontSize: 13, color: C.dark, bold: true, align: 'center' });
slide.addText('SmartLock: Dynamic Price-Freeze Engine', { x: 3.0, y: 2.3, w: 6.5, fontSize: 28, bold: true, color: C.white });
slide.addText('De-risking purchase decisions to reclaim dropped booking funnels', { x: 3.0, y: 3.0, w: 6.5, fontSize: 14, color: C.muted });
slide.addText('By Ajay Avaghade', { x: 3.0, y: 6.3, w: 4, fontSize: 11, color: C.muted });
slide.addText('+22%', { x: 7.2, y: 4.8, w: 2.5, fontSize: 48, bold: true, color: C.gold, align: 'right' });
slide.addText('Checkout conversion lift', { x: 7.2, y: 5.6, w: 2.5, fontSize: 12, color: C.muted, align: 'right' });

// SLIDE 2: The Problem
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('THE PROBLEM', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('High-Consideration Travel Triggers Cart Abandonment', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.white });

slide.addText('68%', { x: 0.5, y: 2.3, w: 2.5, fontSize: 38, bold: true, color: C.bright });
slide.addText('Cart Abandonment Rate', { x: 0.5, y: 3.0, w: 2.5, fontSize: 13, bold: true, color: C.white });
slide.addText('Users drop off to compare prices across aggregators and consult travel companions.', { x: 0.5, y: 3.4, w: 2.5, fontSize: 11, color: C.muted });

slide.addText('4.8 Hrs', { x: 3.5, y: 2.3, w: 2.5, fontSize: 38, bold: true, color: C.bright });
slide.addText('Avg Decision Window', { x: 3.5, y: 3.0, w: 2.5, fontSize: 13, bold: true, color: C.white });
slide.addText('The gap between a user\'s first high-intent flight search and their eventual booking.', { x: 3.5, y: 3.4, w: 2.5, fontSize: 11, color: C.muted });

slide.addText('14%', { x: 6.5, y: 2.3, w: 2.5, fontSize: 38, bold: true, color: C.bright });
slide.addText('Price Hikes in Window', { x: 6.5, y: 3.0, w: 2.5, fontSize: 13, bold: true, color: C.white });
slide.addText('Average fare spike during the consideration window, driving permanent drop-offs.', { x: 6.5, y: 3.4, w: 2.5, fontSize: 11, color: C.muted });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 5.3, w: 9, h: 1.3, fill: { color: C.hero }, rectRadius: 0.08 });
slide.addText('Insight: Users don\'t abandon flights because they lose interest; they abandon to manage price risk. By offering a cheap way to freeze prices, Almosafer can capture search intent and lock users into our ecosystem.', { x: 0.8, y: 5.4, w: 8.4, h: 1.1, fontSize: 12.5, color: C.white, lineSpacing: 18 });

// SLIDE 3: The Insight
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addText('THE INSIGHT', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Transforming Price Anxiety into Booking Insurance', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.dark });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2.2, w: 4.0, h: 4.2, fill: { color: C.white }, shadow: makeShadow(), rectRadius: 0.08 });
slide.addText('❌ Legacy Checkout Funnel', { x: 0.9, y: 2.6, w: 3.2, fontSize: 15, bold: true, color: C.dark });
slide.addText('• Binary decision: buy now or risk paying more\n• Artificial urgency prompts erode customer trust\n• Drop-offs to external price comparison portals\n• 100% loss of abandoned cart transaction history', { x: 0.9, y: 3.2, w: 3.2, h: 2.8, valign: 'top', fontSize: 12, color: C.muted, bullet: true, lineSpacing: 22 });

slide.addShape(prs.ShapeType.rect, { x: 5.5, y: 2.2, w: 4.0, h: 4.2, fill: { color: C.hero }, shadow: makeShadow(), rectRadius: 0.08 });
slide.addText('✅ Insured SmartLock Experience', { x: 5.9, y: 2.6, w: 3.2, fontSize: 15, bold: true, color: C.bright });
slide.addText('• 48-hour price guarantee via a micro-lock-fee\n• Lowers commitment barrier for early researchers\n• Almosafer covers hikes, user gets savings if drops\n• Retains intent data, locking users into final booking', { x: 5.9, y: 3.2, w: 3.2, h: 2.8, valign: 'top', fontSize: 12, color: C.lgray, bullet: true, lineSpacing: 22 });

slide.addShape(prs.ShapeType.ellipse, { x: 4.6, y: 3.9, w: 0.8, h: 0.8, fill: { color: C.bright }, shadow: makeShadow() });
slide.addText('VS', { x: 4.6, y: 3.9, w: 0.8, h: 0.8, fontSize: 12, bold: true, color: C.white, align: 'center' });

// SLIDE 4: The Mechanic
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('THE MECHANIC', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Automated Hedging & Volatility Scoring Flow', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.white });

const steps = ['Route Search', 'Risk Evaluation', 'SmartLock Offered', 'Lock Fee Paid', 'Secure Booking'];
slide.addShape(prs.ShapeType.rect, { x: 1, y: 3.8, w: 8, h: 0.04, fill: { color: C.hero } });

steps.forEach((step, i) => {
  slide.addShape(prs.ShapeType.ellipse, { x: 0.8 + i*2, y: 3.6, w: 0.4, h: 0.4, fill: { color: C.bright } });
  slide.addText(String(i+1), { x: 0.8 + i*2, y: 3.6, w: 0.4, h: 0.4, fontSize: 11, bold: true, color: C.white, align: 'center' });
  slide.addText(step, { x: 0 + i*2, y: 4.2, w: 2, fontSize: 13, bold: true, color: C.white, align: 'center' });
});

slide.addText('1. User searches DXB-RUH flight\n2. Volatility API runs ML risk assessment\n3. Dynamic lock fee is calculated (e.g. 45 AED)\n4. Fee goes to risk pool; fare is locked\n5. User returns to complete booking at guaranteed rate', { x: 0.5, y: 5.3, w: 9, fontSize: 11, color: C.muted, lineSpacing: 16 });

// SLIDE 5: The Product
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addText('THE PRODUCT', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Seamless 4-Screen SmartLock Architecture', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.dark });

const screens = [
  { n: '01 — Flight Search', t: 'Rate Freeze Callout', d: 'Surfaces SmartLock option directly in flight card results during price volatility alerts.' },
  { n: '02 — Freeze Rate', t: 'Lock details drawer', d: 'Displays a live countdown timer, volatility score, clear refund terms, and pay CTAs.' },
  { n: '03 — Fare Secured', t: 'Confirmation receipt', d: 'Shows confirmation, savings tracking against live fares, and a 1-click booking path.' },
  { n: '04 — Risk Controller', t: 'Operations control panel', d: 'Empowers ops team to adjust dynamic risk margins and monitor route volatility rates.' }
];

screens.forEach((s, i) => {
  let cx = 0.5 + i*2.3;
  slide.addShape(prs.ShapeType.rect, { x: cx, y: 2.1, w: 2.15, h: 4.3, fill: { color: C.white }, shadow: makeShadow(), rectRadius: 0.08 });
  slide.addText(s.n, { x: cx+0.1, y: 2.3, w: 1.95, fontSize: 10, bold: true, color: C.bright });
  slide.addText(s.t, { x: cx+0.1, y: 2.6, w: 1.95, fontSize: 12, bold: true, color: C.dark });
  slide.addText(s.d, { x: cx+0.1, y: 3.0, w: 1.95, h: 2.5, valign: 'top', fontSize: 10.5, color: C.muted, lineSpacing: 16 });
});

// SLIDE 6: Impact & ROI
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('IMPACT & ROI', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Sizable Business Value & Reclaimed Conversions', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.white });

const metrics = [
  { val: '+22%', lbl: 'Checkout Conversion Lift', d: 'Encourages hesitant researchers to enter the checkout flow via low-barrier lock fee.' },
  { val: '-40%', lbl: 'Cart Abandonment', d: 'Locks in travel intent, stopping users from shopping around on competing aggregators.' },
  { val: '+15%', lbl: 'SmartLock Attach Rate', d: 'Adds high-margin ancillary lock fee revenue directly to Almosafer\'s balance sheet.' },
  { val: '77%', lbl: 'Net Risk Margin', d: 'Priced so total lock fees collected comfortably exceed payouts for fare differences.' }
];

metrics.forEach((m, i) => {
  let cx = 0.5 + (i%2)*4.5;
  let cy = 2.2 + Math.floor(i/2)*2.1;
  slide.addShape(prs.ShapeType.rect, { x: cx, y: cy, w: 4.2, h: 1.8, fill: { color: C.hero }, rectRadius: 0.08 });
  slide.addText(m.val, { x: cx+0.2, y: cy+0.2, w: 1.5, fontSize: 24, bold: true, color: C.bright });
  slide.addText(m.lbl, { x: cx+1.8, y: cy+0.2, w: 2.2, fontSize: 12, bold: true, color: C.white });
  slide.addText(m.d, { x: cx+0.2, y: cy+0.9, w: 3.8, fontSize: 11, color: C.muted, lineSpacing: 16 });
});

// SLIDE 7: Proof of Work
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 5, y: 0, w: 5, h: 7.5, fill: { color: C.lgray } });

slide.addText('MY EXPERIENCE', { x: 0.5, y: 0.5, w: 4, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('PhonePe Scale', { x: 0.5, y: 0.8, w: 4, fontSize: 22, bold: true, color: C.white });
slide.addText('• Revamped consumer offers discovery & checkout page to drive +22% conversion lift.\n• Deployed Propensity-to-Transact ML targeting models, reducing marketing burn by 32%.\n• Designed and scaled merchant PG self-serve integration platforms, acquiring 5000+ online/offline merchants.', { x: 0.5, y: 1.8, w: 4, h: 3.2, valign: 'top', fontSize: 12.5, color: C.muted, bullet: true, lineSpacing: 22 });

slide.addText('ALMOSAFER MATCH', { x: 5.5, y: 0.5, w: 4, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Why I fit this PM role', { x: 5.5, y: 0.8, w: 4, fontSize: 22, bold: true, color: C.dark });
slide.addText('• Proven capability of leveraging dynamic data/ML algorithms to steer consumer transactions.\n• Strong background in building merchant/API integrations, aligning with travel vendor connectivity.\n• Data-first analytical approach to balance risk margins against customer conversions.', { x: 5.5, y: 1.8, w: 4, h: 3.2, valign: 'top', fontSize: 12.5, color: C.dark, bullet: true, lineSpacing: 22 });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 5.5, w: 9, h: 1.4, fill: { color: C.bright }, rectRadius: 0.08 });
slide.addText('"Dynamic price risk shielding provides a powerful value proposition: turning conversion anxiety into highly profitable, risk-managed booking commitments."', { x: 0.8, y: 5.6, w: 8.4, h: 1.2, fontSize: 13, bold: true, color: C.white, align: 'center', italic: true, lineSpacing: 18 });

// SLIDE 8: Rollout Plan
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('ROLLOUT PLAN', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('SmartLock Phased Rollout Schedule', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.white });

const phases = [
  { t: 'Phase 1: Model Training', d: 'Ingest 12 months of historical DXB-RUH route pricing to construct volatility forecasting baselines.' },
  { t: 'Phase 2: Alpha Test', d: 'Launch rate lock to 5% of web and native app traffic on high-volume regional flights to test fee elasticity.' },
  { t: 'Phase 3: Full Scale', d: 'Roll out globally across all airline integrations, extending engine to hotel bookings.' }
];

phases.forEach((p, i) => {
  slide.addShape(prs.ShapeType.rect, { x: 0.5 + i*3.1, y: 2.2, w: 2.8, h: 2.4, fill: { color: C.hero }, rectRadius: 0.08, line: { color: C.bright, width: 2.5 } });
  slide.addText(p.t, { x: 0.7 + i*3.1, y: 2.5, w: 2.4, fontSize: 13, bold: true, color: C.white });
  slide.addText(p.d, { x: 0.7 + i*3.1, y: 3.1, w: 2.4, fontSize: 11, color: C.muted, lineSpacing: 16 });
});

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 5.0, w: 9, h: 1.2, fill: { color: C.hero }, rectRadius: 0.08 });
slide.addText('What I Need: A 20-minute chat to discuss Almosafer\'s checkout optimization roadmap.', { x: 0.8, y: 5.3, w: 8.4, fontSize: 13.5, color: C.bright, bold: true });
slide.addText('Ajay Avaghade | avaghadeajay009@gmail.com', { x: 0.5, y: 6.7, w: 9, fontSize: 10.5, color: C.muted, align: 'center' });

prs.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/almosafer-smartlock-deck.pptx' }).then(() => {
  console.log('SmartLock Deck generated!');
});
