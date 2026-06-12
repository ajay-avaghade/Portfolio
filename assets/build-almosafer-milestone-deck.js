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
slide.addText('Milestone Explorer: Gamified Loyalty', { x: 3.0, y: 2.3, w: 6.5, fontSize: 28, bold: true, color: C.white });
slide.addText('Toss-inspired travel trails utilizing user photos to complete destination steps', { x: 3.0, y: 3.0, w: 6.5, fontSize: 14, color: C.muted });
slide.addText('By Ajay Avaghade', { x: 3.0, y: 6.3, w: 4, fontSize: 11, color: C.muted });
slide.addText('+17%', { x: 7.2, y: 4.8, w: 2.5, fontSize: 48, bold: true, color: C.gold, align: 'right' });
slide.addText('Customer Lifetime Value (CLTV)', { x: 7.2, y: 5.6, w: 2.5, fontSize: 12, color: C.muted, align: 'right' });

// SLIDE 2: The Problem
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('THE PROBLEM', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Low Booking Frequency & Dormant Travel Customer Lifecycles', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.white });

slide.addText('1.4×', { x: 0.5, y: 2.3, w: 2.5, fontSize: 38, bold: true, color: C.bright });
slide.addText('Annual Booking Rate', { x: 0.5, y: 3.0, w: 2.5, fontSize: 13, bold: true, color: C.white });
slide.addText('Travel is inherently low-frequency. Users search, book, and go dormant for months.', { x: 0.5, y: 3.4, w: 2.5, fontSize: 11, color: C.muted });

slide.addText('78%', { x: 3.5, y: 2.3, w: 2.5, fontSize: 38, bold: true, color: C.bright });
slide.addText('Post-Trip Dormancy', { x: 3.5, y: 3.0, w: 2.5, fontSize: 13, bold: true, color: C.white });
slide.addText('Users ignore app notifications and emails once their current reservation ends.', { x: 3.5, y: 3.4, w: 2.5, fontSize: 11, color: C.muted });

slide.addText('0.2%', { x: 6.5, y: 2.3, w: 2.5, fontSize: 38, bold: true, color: C.bright });
slide.addText('Social Share Rate', { x: 6.5, y: 3.0, w: 2.5, fontSize: 13, bold: true, color: C.white });
slide.addText('Conventional loyalty points and milestones do not drive organic social media engagement.', { x: 6.5, y: 3.4, w: 2.5, fontSize: 11, color: C.muted });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 5.3, w: 9, h: 1.3, fill: { color: C.hero }, rectRadius: 0.08 });
slide.addText('Insight: Pedometer-style milestones (popularized by Toss Korea) can be adapted for travel. Mapping destinations as physical trail steps and prompting users to upload their favorite trip photos transforms a transactional utility into a visual travel diary.', { x: 0.8, y: 5.4, w: 8.4, h: 1.1, fontSize: 12.5, color: C.white, lineSpacing: 18 });

// SLIDE 3: The Insight
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addText('THE INSIGHT', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Toss Pedometer Loops Adapted for Visual Destination Trails', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.dark });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2.2, w: 4.0, h: 4.2, fill: { color: C.white }, shadow: makeShadow(), rectRadius: 0.08 });
slide.addText('❌ Legacy Loyalty Schemes', { x: 0.9, y: 2.6, w: 3.2, fontSize: 15, bold: true, color: C.dark });
slide.addText('• Accrual of points with high redemption barriers\n• Generic gold/platinum status level cards\n• Minimal engagement during post-trip intervals\n• Boring and rigid social referral templates', { x: 0.9, y: 3.2, w: 3.2, h: 2.8, valign: 'top', fontSize: 12, color: C.muted, bullet: true, lineSpacing: 22 });

slide.addShape(prs.ShapeType.rect, { x: 5.5, y: 2.2, w: 4.0, h: 4.2, fill: { color: C.hero }, shadow: makeShadow(), rectRadius: 0.08 });
slide.addText('✅ Milestone Explorer Loop', { x: 5.9, y: 2.6, w: 3.2, fontSize: 15, bold: true, color: C.bright });
slide.addText('• Destination-mapped travel routes (STA-sponsored)\n• Favorite trip photo indicates completed trail step\n• Unlocks high-value local vouchers (50% off tours)\n• Generates personalized postcards for social sharing', { x: 5.9, y: 3.2, w: 3.2, h: 2.8, valign: 'top', fontSize: 12, color: C.lgray, bullet: true, lineSpacing: 22 });

slide.addShape(prs.ShapeType.ellipse, { x: 4.6, y: 3.9, w: 0.8, h: 0.8, fill: { color: C.bright }, shadow: makeShadow() });
slide.addText('VS', { x: 4.6, y: 3.9, w: 0.8, h: 0.8, fontSize: 12, bold: true, color: C.white, align: 'center' });

// SLIDE 4: The Mechanic
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('THE MECHANIC', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Destination Map Progression & Safety Moderation', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.white });

const steps = ['Verify Booking', 'Complete Journey', 'Choose Best Photo', 'AI Moderation', 'Disburse Reward'];
slide.addShape(prs.ShapeType.rect, { x: 1, y: 3.8, w: 8, h: 0.04, fill: { color: C.hero } });

steps.forEach((step, i) => {
  slide.addShape(prs.ShapeType.ellipse, { x: 0.8 + i*2, y: 3.6, w: 0.4, h: 0.4, fill: { color: C.bright } });
  slide.addText(String(i+1), { x: 0.8 + i*2, y: 3.6, w: 0.4, h: 0.4, fontSize: 11, bold: true, color: C.white, align: 'center' });
  slide.addText(step, { x: 0 + i*2, y: 4.2, w: 2, fontSize: 13, bold: true, color: C.white, align: 'center' });
});

slide.addText('1. User registers flight/hotel booking for Heritage Trail (e.g. Jeddah)\n2. GPS and partner reservation system verify trip completion\n3. User uploads their favorite trip photo to mark the step\n4. AI Content Moderation evaluates image safety & relevance in milliseconds\n5. Partner-sponsored voucher (e.g. AlUla tour) is instantly issued', { x: 0.5, y: 5.3, w: 9, fontSize: 11, color: C.muted, lineSpacing: 16 });

// SLIDE 5: The Product
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addText('THE PRODUCT', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Toss-Inspired Photo Milestones & Partner Dashboard', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.dark });

const screens = [
  { n: '01 — Milestone Trail', t: 'Visual map interface', d: 'Displays Saudi Heritage Trail route path. Completed steps display user photos inside circle nodes.' },
  { n: '02 — Upload Step', t: 'Camera roll selection', d: 'Validates Radisson stay, prompt uploads, and provides quick crop previews for preferred photos.' },
  { n: '03 — Celebration', t: 'Postcard generator', d: 'Triggers confetti, reveals tourism board discount voucher, and generates shareable travel journal.' },
  { n: '04 — Campaign Ops', t: 'Moderation control', d: 'STA console displaying explorer cohorts, voucher redemption rates, and AI moderation queues.' }
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
slide.addText('Fostering Engagement, Lifetime Value, and Regional Sponsorships', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.white });

const metrics = [
  { val: '+17%', lbl: 'CLTV Expansion', d: 'Drives repeat domestic travel booking behavior, maximizing customer lifetime value.' },
  { val: '+30%', lbl: 'Repeat Bookings', d: 'Unlocking successive milestone vouchers encourages users to book their next tour.' },
  { val: '63%', lbl: 'Milestone Completion', d: 'Toss-inspired visuals and photo logs drive high user follow-through and campaign completions.' },
  { val: '2.4×', lbl: 'Social Sharing Gain', d: 'Postcard journal layouts prompt organic user uploads to WhatsApp and social media.' }
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
slide.addText('PhonePe Milestone Scale', { x: 0.5, y: 0.8, w: 4, fontSize: 22, bold: true, color: C.white });
slide.addText('• Launched PhonePe gamified milestones for users and merchants, driving 63% completion rate.\n• Boosted long-term merchant retention by 23% and increased overall customer lifetime value by 17%.\n• Collaborated with design and analytics to implement dynamic progress bars and visual rewards.', { x: 0.5, y: 1.8, w: 4, h: 3.2, valign: 'top', fontSize: 12.5, color: C.muted, bullet: true, lineSpacing: 22 });

slide.addText('ALMOSAFER MATCH', { x: 5.5, y: 0.5, w: 4, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Why I fit this PM role', { x: 5.5, y: 0.8, w: 4, fontSize: 22, bold: true, color: C.dark });
slide.addText('• Proven track record of launching highly successful gamified user engagement and retention features.\n• Obsessed with qualitative and quantitative research to identify customer blockers (Toss case study).\n• Expertise leading scrum teams across Tech, UX, and marketing to build complex consumer portals.', { x: 5.5, y: 1.8, w: 4, h: 3.2, valign: 'top', fontSize: 12.5, color: C.dark, bullet: true, lineSpacing: 22 });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 5.5, w: 9, h: 1.4, fill: { color: C.bright }, rectRadius: 0.08 });
slide.addText('"Gamification is most effective when it is personal. Linking travel loyalty rewards to a user-uploaded visual diary of their journey creates an emotional hook that drives repeat bookings."', { x: 0.8, y: 5.6, w: 8.4, h: 1.2, fontSize: 13, bold: true, color: C.white, align: 'center', italic: true, lineSpacing: 18 });

// SLIDE 8: Rollout Plan
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('ROLLOUT PLAN', { x: 0.5, y: 0.5, w: 9, fontSize: 9, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Milestone Explorer Pilot & Expansion Timeline', { x: 0.5, y: 0.8, w: 9, fontSize: 26, bold: true, color: C.white });

const phases = [
  { t: 'Phase 1: Heritage Trail', d: 'Co-launch Saudi Heritage Trail with STA, validating GPS tracking and photo uploads in Riyadh/Jeddah.' },
  { t: 'Phase 2: AI Guardrails', d: 'Train moderation models to reject off-topic photo uploads and automatically tag landscapes.' },
  { t: 'Phase 3: Global Trails', d: 'Expand milestones internationally (e.g. Dubai, London trails), integrating global airline & hotel partners.' }
];

phases.forEach((p, i) => {
  slide.addShape(prs.ShapeType.rect, { x: 0.5 + i*3.1, y: 2.2, w: 2.8, h: 2.4, fill: { color: C.hero }, rectRadius: 0.08, line: { color: C.bright, width: 2.5 } });
  slide.addText(p.t, { x: 0.7 + i*3.1, y: 2.5, w: 2.4, fontSize: 13, bold: true, color: C.white });
  slide.addText(p.d, { x: 0.7 + i*3.1, y: 3.1, w: 2.4, fontSize: 11, color: C.muted, lineSpacing: 16 });
});

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 5.0, w: 9, h: 1.2, fill: { color: C.hero }, rectRadius: 0.08 });
slide.addText('What I Need: A 20-minute chat to discuss Almosafer\'s loyalty and retention strategy.', { x: 0.8, y: 5.3, w: 8.4, fontSize: 13.5, color: C.bright, bold: true });
slide.addText('Ajay Avaghade | avaghadeajay009@gmail.com', { x: 0.5, y: 6.7, w: 9, fontSize: 10.5, color: C.muted, align: 'center' });

prs.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/almosafer-milestone-deck.pptx' }).then(() => {
  console.log('Milestone Deck generated!');
});
