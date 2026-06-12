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
slide.addText('RouteIQ: Dynamic Cart Bundling', { x: 3.0, y: 2.3, w: 6.5, fontSize: 28, bold: true, color: C.white });
slide.addText('Orchestrating real-time API integrations for high-margin ancillary attach rates', { x: 3.0, y: 3.0, w: 6.5, fontSize: 14, color: C.muted });
slide.addText('By Ajay Avaghade', { x: 3.0, y: 6.3, w: 4, fontSize: 11, color: C.muted });
slide.addText('+35%', { x: 7.2, y: 4.8, w: 2.5, fontSize: 48, bold: true, color: C.gold, align: 'right' });
slide.addText('Average Order Value (AOV)', { x: 7.2, y: 5.6, w: 2.5, fontSize: 12, color: C.muted, align: 'right' });

// SLIDE 2: The Problem
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('THE PROBLEM', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Siloed Booking Patterns & Outdated Campaign Operations', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.white });

slide.addText('82%', { x: 0.5, y: 2.3, w: 2.5, fontSize: 38, bold: true, color: C.bright });
slide.addText('Single-Vertical Silos', { x: 0.5, y: 3.0, w: 2.5, fontSize: 13, bold: true, color: C.white });
slide.addText('Users book flights but buy hotels and local ground transfers elsewhere, losing platform volume.', { x: 0.5, y: 3.4, w: 2.5, fontSize: 11, color: C.muted });

slide.addText('12%', { x: 3.5, y: 2.3, w: 2.5, fontSize: 38, bold: true, color: C.bright });
slide.addText('Generic Attach Rates', { x: 3.5, y: 3.0, w: 2.5, fontSize: 13, bold: true, color: C.white });
slide.addText('Soggy cross-sell conversion on travel insurance and car rentals due to rigid, unpersonalized CTAs.', { x: 3.5, y: 3.4, w: 2.5, fontSize: 11, color: C.muted });

slide.addText('2 Days', { x: 6.5, y: 2.3, w: 2.5, fontSize: 38, bold: true, color: C.bright });
slide.addText('Offer Campaign TAT', { x: 6.5, y: 3.0, w: 2.5, fontSize: 13, bold: true, color: C.white });
slide.addText('Manual back-office partner setups delay campaign launch velocity and limit promotional agility.', { x: 6.5, y: 3.4, w: 2.5, fontSize: 11, color: C.muted });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 5.3, w: 9, h: 1.3, fill: { color: C.hero }, rectRadius: 0.08 });
slide.addText('Insight: Conventional travel packages are static and forced. Real-time API connections with ground providers (Careem, local operators) and single-click checkout discounts convert siloed bookings into unified travel bundles.', { x: 0.8, y: 5.4, w: 8.4, h: 1.1, fontSize: 12.5, color: C.white, lineSpacing: 18 });

// SLIDE 3: The Insight
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addText('THE INSIGHT', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Dynamic Cart Bundling Powered by Real-Time Partner APIs', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.dark });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2.2, w: 4.0, h: 4.2, fill: { color: C.white }, shadow: makeShadow(), rectRadius: 0.08 });
slide.addText('❌ Conventional Packages', { x: 0.9, y: 2.6, w: 3.2, fontSize: 15, bold: true, color: C.dark });
slide.addText('• Hard-coded flight + hotel package deals\n• Spammy email marketing and banners\n• Multiple checkouts for cars/activities\n• Manual vendor coordination for promotions', { x: 0.9, y: 3.2, w: 3.2, h: 2.8, valign: 'top', fontSize: 12, color: C.muted, bullet: true, lineSpacing: 22 });

slide.addShape(prs.ShapeType.rect, { x: 5.5, y: 2.2, w: 4.0, h: 4.2, fill: { color: C.hero }, shadow: makeShadow(), rectRadius: 0.08 });
slide.addText('✅ RouteIQ Dynamic Cart', { x: 5.9, y: 2.6, w: 3.2, fontSize: 15, bold: true, color: C.bright });
slide.addText('• Context-aware cross-sell recommendations\n• Dynamic discounts adjusting in real-time\n• 1-click checkout with a single unified payment\n• Self-serve partner console with 30-min campaign TAT', { x: 5.9, y: 3.2, w: 3.2, h: 2.8, valign: 'top', fontSize: 12, color: C.lgray, bullet: true, lineSpacing: 22 });

slide.addShape(prs.ShapeType.ellipse, { x: 4.6, y: 3.9, w: 0.8, h: 0.8, fill: { color: C.bright }, shadow: makeShadow() });
slide.addText('VS', { x: 4.6, y: 3.9, w: 0.8, h: 0.8, fontSize: 12, bold: true, color: C.white, align: 'center' });

// SLIDE 4: The Mechanic
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('THE MECHANIC', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Intent Scoring & Self-Serve Partner Integration', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.white });

const steps = ['Add Primary Booking', 'Query Partner APIs', 'Calculate Discount', 'Unified 1-Click Pay', 'Self-Serve Launch'];
slide.addShape(prs.ShapeType.rect, { x: 1, y: 3.8, w: 8, h: 0.04, fill: { color: C.hero } });

steps.forEach((step, i) => {
  slide.addShape(prs.ShapeType.ellipse, { x: 0.8 + i*2, y: 3.6, w: 0.4, h: 0.4, fill: { color: C.bright } });
  slide.addText(String(i+1), { x: 0.8 + i*2, y: 3.6, w: 0.4, h: 0.4, fontSize: 11, bold: true, color: C.white, align: 'center' });
  slide.addText(step, { x: 0 + i*2, y: 4.2, w: 2, fontSize: 13, bold: true, color: C.white, align: 'center' });
});

slide.addText('1. User adds hotel in Riyadh to cart\n2. Engine calls Careem & attraction partner API endpoints\n3. RouteIQ evaluates intent and adds dynamic bundle markdown\n4. Unified payment settles flights, hotel, and taxi booking\n5. Local vendors launch campaigns independently via embeddable consoles', { x: 0.5, y: 5.3, w: 9, fontSize: 11, color: C.muted, lineSpacing: 16 });

// SLIDE 5: The Product
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addText('THE PRODUCT', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Real-Time Cross-Vertical Checkout UX', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.dark });

const screens = [
  { n: '01 — Hotel Detail', t: 'Dynamic cross-sell', d: 'Displays custom add-ons like taxi transfers or city tours matching user search parameters.' },
  { n: '02 — Cart Drawer', t: 'Unified bundling', d: 'Presents dynamic discount and real-time discount explanation (e.g. -110 SAR).' },
  { n: '03 — Confirmation', t: 'Unified receipt', d: 'Aggregates flight, hotel, and transfer booking confirmations into a single voucher.' },
  { n: '04 — Partner Portal', t: 'Self-serve operations', d: 'Allows vendors to list inventory, review API latency, and launch campaigns in 30 minutes.' }
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
slide.addText('Elevating Average Order Value and Operational Speed', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.white });

const metrics = [
  { val: '+35%', lbl: 'AOV Uplift', d: 'Drives hotel and ancillary additions to flight carts, expanding transaction size.' },
  { val: '+25%', lbl: 'Ancillary Attach Rate', d: 'Converts low-intent traffic into ground transfer and local activity bookings.' },
  { val: '30 Min', lbl: 'Promotional Campaign TAT', d: 'Slashes promotion deployment time from 2 days to 30 mins via self-serve tools.' },
  { val: '-23%', lbl: 'Lower Partner CAC', d: 'Passes transaction traffic to partners, reducing customer acquisition costs.' }
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
slide.addText('PhonePe Cart & PG Scale', { x: 0.5, y: 0.8, w: 4, fontSize: 22, bold: true, color: C.white });
slide.addText('• Launched cart-level real-time offer incentives on PhonePe Pincode, driving +35% AOV uplift.\n• Reduced quick commerce checkout cart abandonment by 60% and lifted 30-day retention by 20%.\n• Extended self-serve campaign engines to AppStore developers, shrinking offer live TAT to 30 mins.', { x: 0.5, y: 1.8, w: 4, h: 3.2, valign: 'top', fontSize: 12.5, color: C.muted, bullet: true, lineSpacing: 22 });

slide.addText('ALMOSAFER MATCH', { x: 5.5, y: 0.5, w: 4, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Why I fit this PM role', { x: 5.5, y: 0.8, w: 4, fontSize: 22, bold: true, color: C.dark });
slide.addText('• Direct experience building cart-level contextual triggers and dynamic merchant pricing models.\n• Led cross-functional Tech/UX/Ops sprints to integrate B2B platforms and self-serve consoles.\n• Deep understanding of API ecosystems for third-party service exchange and payouts.', { x: 5.5, y: 1.8, w: 4, h: 3.2, valign: 'top', fontSize: 12.5, color: C.dark, bullet: true, lineSpacing: 22 });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 5.5, w: 9, h: 1.4, fill: { color: C.bright }, rectRadius: 0.08 });
slide.addText('"Unlocking B2B self-serve consoles doesn\'t just streamline internal operations; it speeds up campaign velocity, allowing partner brands to drive checkout conversion."', { x: 0.8, y: 5.6, w: 8.4, h: 1.2, fontSize: 13, bold: true, color: C.white, align: 'center', italic: true, lineSpacing: 18 });

// SLIDE 8: Rollout Plan
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('ROLLOUT PLAN', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('RouteIQ Phased Integration Roadmap', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.white });

const phases = [
  { t: 'Phase 1: API Setup', d: 'Design and deploy unified partner API endpoints for transportation (Careem) and local tourism vendors.' },
  { t: 'Phase 2: Alpha Cart', d: 'Test real-time cart-bundling discount triggers on a subset of Riyadh hotel bookings to monitor AOV impacts.' },
  { t: 'Phase 3: B2B Dashboard', d: 'Extend self-serve promotion control panels to partners, allowing direct campaign configuration.' }
];

phases.forEach((p, i) => {
  slide.addShape(prs.ShapeType.rect, { x: 0.5 + i*3.1, y: 2.2, w: 2.8, h: 2.4, fill: { color: C.hero }, rectRadius: 0.08, line: { color: C.bright, width: 2.5 } });
  slide.addText(p.t, { x: 0.7 + i*3.1, y: 2.5, w: 2.4, fontSize: 13, bold: true, color: C.white });
  slide.addText(p.d, { x: 0.7 + i*3.1, y: 3.1, w: 2.4, fontSize: 11, color: C.muted, lineSpacing: 16 });
});

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 5.0, w: 9, h: 1.2, fill: { color: C.hero }, rectRadius: 0.08 });
slide.addText('What I Need: A 20-minute chat to discuss Almosafer\'s dynamic cart bundling architecture.', { x: 0.8, y: 5.3, w: 8.4, fontSize: 13.5, color: C.bright, bold: true });
slide.addText('Ajay Avaghade | avaghadeajay009@gmail.com', { x: 0.5, y: 6.7, w: 9, fontSize: 10.5, color: C.muted, align: 'center' });

prs.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/almosafer-routeiq-deck.pptx' }).then(() => {
  console.log('RouteIQ Deck generated!');
});
