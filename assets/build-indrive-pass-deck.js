const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0B0E13',
  hero:   '141810',
  bright: 'B5E033',
  orange: 'FF6B35',
  bronze: 'CD7F32',
  silver: '94A3B8',
  gold:   'F59E0B',
  green:  '22C55E',
  white:  'FFFFFF',
  lgray:  'F1F5F9',
  muted:  '94A3B8',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.15 });

// ─────────────────────────────────────────────
// SLIDE 1 — COVER
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  // diagonal lines watermark
  for (let i = 0; i < 12; i++) {
    sl.addShape(prs.ShapeType.line, {
      x: -0.5 + i * 1.0, y: 0, w: 4, h: 7.5,
      line: { color: C.orange, width: 0.5, transparency: 90 },
    });
  }

  // Pass icon watermark
  sl.addText('🎫', {
    x: 6.0, y: 0.2, w: 3.5, h: 3.5,
    fontSize: 180, transparency: 92,
  });

  // eyebrow
  sl.addText('INDRIVE  ·  SENIOR PRODUCT MANAGER — GROWTH & PERSONALIZATION', {
    x: 0.5, y: 0.5, w: 7.5, h: 0.3,
    fontSize: 9, color: C.orange, bold: true, charSpacing: 2,
  });

  // title
  sl.addText('inDrive Pass', {
    x: 0.5, y: 1.1, w: 7, h: 0.95,
    fontSize: 52, color: C.white, bold: true, fontFace: 'Arial',
  });

  // subtitle
  sl.addText('The subscription model that locks users in through value:\nearn credits on every ride, unlock higher tiers, adopt\nnew verticals — and lose it all if you switch.', {
    x: 0.5, y: 2.2, w: 7.2, h: 1.3,
    fontSize: 17, color: C.muted, lineSpacingMultiple: 1.35,
  });

  // presenter
  sl.addText('Presented by Ajay Avaghade  ·  Growth & Personalization PM', {
    x: 0.5, y: 3.7, w: 6, h: 0.3,
    fontSize: 11, color: C.bright,
  });

  // Tier strip bottom-right
  const tiers = [
    { label:'Bronze', val:'5%', color: C.bronze },
    { label:'Silver', val:'8%', color: C.silver },
    { label:'Gold', val:'12%', color: C.gold },
    { label:'Explorer', val:'15%', color: C.bright },
  ];
  tiers.forEach((t, i) => {
    sl.addShape(prs.ShapeType.roundRect, {
      x: 5.5 + i * 1.12, y: 5.4, w: 1.0, h: 1.75,
      fill: { color: C.hero },
      line: { color: t.color, width: 1.5 },
      rectRadius: 0.1,
    });
    sl.addText(t.val, { x: 5.5 + i * 1.12, y: 5.55, w: 1.0, h: 0.65, fontSize: 22, color: t.color, bold:true, align:'center', fontFace:'Arial' });
    sl.addText('back', { x: 5.5 + i * 1.12, y: 6.2, w: 1.0, h: 0.2, fontSize: 9, color: C.muted, align:'center' });
    sl.addText(t.label, { x: 5.5 + i * 1.12, y: 6.42, w: 1.0, h: 0.25, fontSize: 9, color: t.color, bold:true, align:'center' });
  });
}

// ─────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE PROBLEM', {
    x: 0.5, y: 0.4, w: 9, h: 0.3,
    fontSize: 9, color: C.orange, bold: true, charSpacing: 2,
  });
  sl.addText('inDrive Has Zero Switching Cost —\nA Competitor Promo Can Steal Any User With One Notification', {
    x: 0.5, y: 0.75, w: 9, h: 0.95,
    fontSize: 25, color: C.white, bold: true, lineSpacingMultiple: 1.2,
  });

  const cols3 = [
    { icon:'🔓', val:'Zero', lbl:'retention mechanics', sub:'No subscription, no credits, no tier. Every user equally acquirable by a competitor.', color: C.orange },
    { icon:'📊', val:'4%', lbl:'cross-vertical adoption', sub:'A user who rides never discovers Freight or Services — siloed verticals, siloed revenue.', color: C.silver },
    { icon:'💸', val:'High CAC', lbl:'repeat user re-acquisition', sub:'Without earned value, inDrive must subsidize every re-engagement from scratch.', color: C.bright },
  ];
  cols3.forEach((item, i) => {
    const x = 0.5 + i * 3.2;
    sl.addShape(prs.ShapeType.roundRect, {
      x, y: 1.9, w: 3.0, h: 3.0,
      fill: { color: '1A1A1A' },
      line: { color: item.color, width: 1.2 },
      rectRadius: 0.1,
      shadow: makeShadow(),
    });
    sl.addText(item.icon, { x, y: 2.0, w: 3.0, h: 0.5, fontSize: 24, align:'center', color: item.color });
    sl.addText(item.val, { x, y: 2.5, w: 3.0, h: 0.65, fontSize: 36, color: item.color, bold:true, align:'center', fontFace:'Arial' });
    sl.addText(item.lbl, { x, y: 3.15, w: 3.0, h: 0.4, fontSize: 12, color: C.white, bold:true, align:'center' });
    sl.addText(item.sub, { x: x+0.1, y: 3.6, w: 2.8, h: 0.85, fontSize: 10, color: C.muted, align:'center', lineSpacingMultiple:1.3 });
  });

  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.5, y: 5.15, w: 9.0, h: 0.95,
    fill: { color: '1A0800' },
    line: { color: C.orange, width: 1 },
    rectRadius: 0.08,
  });
  sl.addText('A user who takes 20 rides a month has exactly the same relationship with inDrive as a first-time user. There is no earned value, no tier, no credits that create the cost of switching. A single promotional push from Bolt or Uber can re-acquire them tonight. The fix is not a loyalty points gimmick — it is a subscription that makes your balance worth more than a competitor\'s discount.', {
    x: 0.7, y: 5.25, w: 8.6, h: 0.75,
    fontSize: 10, color: C.white, lineSpacingMultiple: 1.35,
  });
}

// ─────────────────────────────────────────────
// SLIDE 3 — THE INSIGHT
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE INSIGHT', {
    x: 0.5, y: 0.4, w: 9, h: 0.3,
    fontSize: 9, color: C.bronze, bold: true, charSpacing: 2,
  });
  sl.addText('The Best Lock-In Isn\'t a Penalty for Leaving —\nIt\'s Credits You\'ll Lose If You Do', {
    x: 0.5, y: 0.75, w: 9, h: 0.9,
    fontSize: 25, color: C.dark, bold: true, lineSpacingMultiple: 1.2,
  });

  // Left column
  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.4, y: 1.85, w: 4.3, h: 3.55,
    fill: { color: 'FFF8EE' },
    line: { color: C.bronze, width: 1.5 },
    rectRadius: 0.1,
  });
  sl.addText('❌  Generic Loyalty Tactics', {
    x: 0.6, y: 2.0, w: 3.9, h: 0.4,
    fontSize: 13, color: C.bronze, bold: true,
  });
  const leftPts = [
    'Loyalty points that expire and feel like fake currency',
    'Promo codes given once then never again',
    'Referral bonus that stops after 1 invite',
    'Discount for new users only — loyal users get nothing',
  ];
  leftPts.forEach((pt, i) => {
    sl.addText(`• ${pt}`, { x: 0.65, y: 2.5 + i * 0.65, w: 3.85, h: 0.6, fontSize: 11, color: '5A3A00', lineSpacingMultiple:1.25 });
  });

  // VS
  sl.addShape(prs.ShapeType.ellipse, {
    x: 4.5, y: 3.05, w: 0.9, h: 0.9,
    fill: { color: C.dark },
    line: { color: C.dark, width: 0 },
  });
  sl.addText('VS', { x: 4.5, y: 3.05, w: 0.9, h: 0.9, fontSize: 12, color: C.white, bold:true, align:'center', valign:'middle' });

  // Right column
  sl.addShape(prs.ShapeType.roundRect, {
    x: 5.2, y: 1.85, w: 4.4, h: 3.55,
    fill: { color: 'F5FFF0' },
    line: { color: C.bright, width: 1.5 },
    rectRadius: 0.1,
  });
  sl.addText('✅  inDrive Pass', {
    x: 5.4, y: 2.0, w: 4.0, h: 0.4,
    fontSize: 13, color: C.accent, bold: true,
  });
  const rightPts = [
    'Monthly Pass → X% cashback on every ride as inDrive Credits',
    'Credits only spendable on inDrive — balance lost if you switch',
    'More rides = higher tier = higher % back = stronger lock-in',
    'Cross-vertical milestones: use Freight + Services → Explorer tier',
    'The credits are the switching cost — no friction, just value',
  ];
  rightPts.forEach((pt, i) => {
    sl.addText(`• ${pt}`, { x: 5.4, y: 2.45 + i * 0.58, w: 4.0, h: 0.52, fontSize: 11, color: '1A3A00', lineSpacingMultiple:1.25 });
  });

  // Callout
  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.4, y: 5.6, w: 9.2, h: 0.8,
    fill: { color: C.dark },
    line: { color: C.gold, width: 1 },
    rectRadius: 0.08,
  });
  sl.addText('Key insight: Users do not leave platforms that owe them money. A $24.80 credit balance is worth more as a retention mechanism than any re-acquisition campaign — because it only costs inDrive on rides that have already been completed.', {
    x: 0.6, y: 5.7, w: 8.8, h: 0.6,
    fontSize: 10.5, color: C.white, lineSpacingMultiple: 1.3,
  });
}

// ─────────────────────────────────────────────
// SLIDE 4 — THE MECHANIC
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE MECHANIC', {
    x: 0.5, y: 0.4, w: 9, h: 0.3,
    fontSize: 9, color: C.gold, bold: true, charSpacing: 2,
  });
  sl.addText('Two Compounding Engines: Ride Tiers + Cross-Vertical Milestones', {
    x: 0.5, y: 0.75, w: 9, h: 0.5,
    fontSize: 23, color: C.white, bold: true,
  });

  // Tier ladder
  const tiers = [
    { t:'Bronze', rides:'5 rides/mo', pct:'5% back', ex:'', color: C.bronze, x: 0.4, w: 2.1 },
    { t:'Silver', rides:'10 rides/mo', pct:'8% back', ex:'', color: C.silver, x: 2.65, w: 2.1 },
    { t:'Gold', rides:'15 rides/mo', pct:'12% back', ex:'BidBand Band 3 default', color: C.gold, x: 4.9, w: 2.1 },
    { t:'Explorer', rides:'All 3 verticals', pct:'15% back', ex:'Priority matching + 6-mo rollover', color: C.bright, x: 7.15, w: 2.45 },
  ];
  tiers.forEach((tier) => {
    sl.addShape(prs.ShapeType.roundRect, {
      x: tier.x, y: 1.45, w: tier.w, h: 3.5,
      fill: { color: C.hero },
      line: { color: tier.color, width: 1.8 },
      rectRadius: 0.1,
      shadow: makeShadow(),
    });
    sl.addText(tier.t, { x: tier.x, y: 1.55, w: tier.w, h: 0.4, fontSize: 14, color: tier.color, bold:true, align:'center' });
    sl.addText(tier.pct, { x: tier.x, y: 2.0, w: tier.w, h: 0.55, fontSize: 24, color: tier.color, bold:true, align:'center', fontFace:'Arial' });
    sl.addText(tier.rides, { x: tier.x, y: 2.6, w: tier.w, h: 0.3, fontSize: 10, color: C.muted, align:'center' });
    if (tier.ex) {
      sl.addShape(prs.ShapeType.roundRect, {
        x: tier.x + 0.1, y: 3.0, w: tier.w - 0.2, h: 0.65,
        fill: { color: '1A2200' },
        line: { color: tier.color, width: 0.6, transparency:40 },
        rectRadius: 0.06,
      });
      sl.addText(tier.ex, {
        x: tier.x + 0.12, y: 3.05, w: tier.w - 0.24, h: 0.55,
        fontSize: 8.5, color: tier.color, align:'center', lineSpacingMultiple:1.25,
      });
    }
    // Arrow
    if (tier.x < 7.0) {
      sl.addText('→', { x: tier.x + tier.w, y: 2.8, w: 0.5, h: 0.3, fontSize: 14, color: C.muted, align:'center' });
    }
  });

  // Cross-vertical milestones section
  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.4, y: 5.1, w: 9.2, h: 1.55,
    fill: { color: '0D1500' },
    line: { color: C.bright, width: 1 },
    rectRadius: 0.1,
  });
  sl.addText('Cross-Vertical Milestone Engine', {
    x: 0.6, y: 5.2, w: 4.0, h: 0.3,
    fontSize: 11, color: C.bright, bold: true,
  });
  const pillars = [
    { icon:'🚗', name:'Rides', goal:'15 rides → Gold', color: C.gold },
    { icon:'🚛', name:'Freight', goal:'5 shipments → Freight tier', color: C.orange },
    { icon:'🔧', name:'Services', goal:'3 tasks → Services tier', color: C.silver },
    { icon:'🌟', name:'All-Access', goal:'All 3 pillars → Explorer 15%', color: C.bright },
  ];
  pillars.forEach((p, i) => {
    const x = 0.6 + i * 2.3;
    sl.addText(`${p.icon} ${p.name}`, { x, y: 5.55, w: 2.1, h: 0.3, fontSize: 10.5, color: p.color, bold:true });
    sl.addText(p.goal, { x, y: 5.85, w: 2.1, h: 0.55, fontSize: 9, color: C.muted, lineSpacingMultiple:1.25 });
  });
}

// ─────────────────────────────────────────────
// SLIDE 5 — THE PRODUCT
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE PRODUCT', {
    x: 0.5, y: 0.4, w: 9, h: 0.3,
    fontSize: 9, color: C.bronze, bold: true, charSpacing: 2,
  });
  sl.addText('4 Screen States — The Pass Lifecycle', {
    x: 0.5, y: 0.75, w: 9, h: 0.5,
    fontSize: 24, color: C.dark, bold: true,
  });

  const screens = [
    { n:'01', title:'Pass Home', desc:'Credits balance, current tier, ride milestone progress bars, cross-vertical tracker. All earned value visible in one glance.', mock:['Credits: $24.80','Silver Tier  ·  8% back','Rides: 14/15 for Gold','[===92%==========]','Try Freight — unlock next'], color: C.gold },
    { n:'02', title:'Active Ride', desc:'Live earning meter during the ride. Credits ticking up in real-time. Gold milestone progress bar showing this ride tips the threshold.', mock:['On your way  ·  Pass LIVE','Earning: +$1.31 this ride','8% of $16.40 fare','New balance: $26.11','15/15 after this ride!'], color: C.orange },
    { n:'03', title:'Gold Unlocked', desc:'Tier upgrade celebration. New benefits displayed clearly — 12% back, BidBand Band 3 default, Freight first ride free. Confetti.', mock:['Gold Tier Unlocked!','12% back per ride','BidBand Band 3 default','Freight: first shipment free','[View Gold Dashboard]'], color: C.gold },
    { n:'04', title:'Multi-Vertical', desc:'All-Access challenge tracker. Three pillars: Rides (done), Freight (in progress), Services (locked). Explorer benefits shown as goal.', mock:['All-Access Challenge','Rides: 15/15  Done','Freight: 2/5  In progress','Services: 0/3  Locked','Explorer: 2 of 3 complete'], color: C.bright },
  ];

  screens.forEach((sc, i) => {
    const x = 0.35 + i * 2.35;
    sl.addShape(prs.ShapeType.roundRect, {
      x, y: 1.45, w: 2.1, h: 4.95,
      fill: { color: C.dark },
      line: { color: sc.color, width: 2 },
      rectRadius: 0.12,
      shadow: makeShadow(),
    });
    sl.addText(sc.n, { x, y: 1.55, w: 2.1, h: 0.3, fontSize: 9, color: sc.color, bold:true, align:'center', charSpacing:1 });
    sl.addText(sc.title, { x, y: 1.85, w: 2.1, h: 0.35, fontSize: 12, color: C.white, bold:true, align:'center' });
    sl.addText(sc.desc, { x: x+0.1, y: 2.25, w: 1.9, h: 1.1, fontSize: 9, color: C.muted, align:'center', lineSpacingMultiple:1.3 });
    sl.addShape(prs.ShapeType.roundRect, {
      x: x+0.12, y: 3.4, w: 1.86, h: 1.7,
      fill: { color: '1A1F1A' },
      line: { color: sc.color, width: 0.6, transparency:50 },
      rectRadius: 0.07,
    });
    sc.mock.forEach((line, j) => {
      sl.addText(line, {
        x: x+0.17, y: 3.48 + j * 0.31, w: 1.76, h: 0.28,
        fontSize: 7.5, color: sc.color, fontFace:'Courier New', align:'center',
      });
    });
  });

  sl.addText('Prototype: indrive-pass-prototype.html  —  interactive live demo of all 4 screen states', {
    x: 0.5, y: 6.7, w: 9.0, h: 0.3,
    fontSize: 9.5, color: C.accent, italic: true, align: 'center',
  });
}

// ─────────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('IMPACT & ROI', {
    x: 0.5, y: 0.4, w: 9, h: 0.3,
    fontSize: 9, color: C.gold, bold: true, charSpacing: 2,
  });
  sl.addText('Projected Impact — Built on PhonePe Milestone Proof Points', {
    x: 0.5, y: 0.75, w: 9, h: 0.5,
    fontSize: 24, color: C.white, bold: true,
  });

  // Left
  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.4, y: 1.45, w: 0.12, h: 3.5,
    fill: { color: C.gold }, line: { color: C.gold, width: 0 }, rectRadius: 0.06,
  });
  sl.addText('USER RETENTION', { x: 0.7, y: 1.45, w: 4.2, h: 0.3, fontSize: 9, color: C.gold, bold:true, charSpacing:1.5 });
  const leftM = [
    { val:'2.8×', lbl:'90-day retention', sub:'Pass subscribers vs non-subscribers' },
    { val:'+34%', lbl:'cross-vertical adoption', sub:'Explorer milestone drives Freight + Services trial' },
    { val:'−42%', lbl:'CAC on second ride', sub:'Credits from ride 1 fund ride 2 — no subsidy needed' },
    { val:'12%', lbl:'Gold tier revenue uplift', sub:'Higher cashback band = higher ride frequency' },
  ];
  leftM.forEach((m, i) => {
    sl.addText(m.val, { x: 0.7, y: 1.85 + i * 0.85, w: 2.2, h: 0.55, fontSize: 32, color: C.gold, bold:true, fontFace:'Arial' });
    sl.addText(m.lbl, { x: 2.9, y: 1.9 + i * 0.85, w: 2.0, h: 0.25, fontSize: 10.5, color: C.white, bold:true });
    sl.addText(m.sub, { x: 2.9, y: 2.17 + i * 0.85, w: 2.0, h: 0.25, fontSize: 9, color: C.muted });
  });

  // Right
  sl.addShape(prs.ShapeType.roundRect, {
    x: 5.3, y: 1.45, w: 0.12, h: 3.5,
    fill: { color: C.bright }, line: { color: C.bright, width: 0 }, rectRadius: 0.06,
  });
  sl.addText('PLATFORM ROI', { x: 5.6, y: 1.45, w: 4.0, h: 0.3, fontSize: 9, color: C.bright, bold:true, charSpacing:1.5 });
  const rightM = [
    { val:'Low cost', lbl:'per retained user', sub:'Credits only paid on completed rides — no upfront burn' },
    { val:'3× CLTV', lbl:'Explorer vs single-vertical', sub:'Rides + Freight + Services vs rides-only revenue' },
    { val:'+28%', lbl:'GMV per Pass user', sub:'Higher tier = more frequent rides + cross-vertical spend' },
    { val:'Day 30', lbl:'payback on Pass cost', sub:'5% cashback earned back in ride volume by month 1' },
  ];
  rightM.forEach((m, i) => {
    sl.addText(m.val, { x: 5.6, y: 1.85 + i * 0.85, w: 2.2, h: 0.55, fontSize: 32, color: C.bright, bold:true, fontFace:'Arial' });
    sl.addText(m.lbl, { x: 7.8, y: 1.9 + i * 0.85, w: 2.0, h: 0.25, fontSize: 10.5, color: C.white, bold:true });
    sl.addText(m.sub, { x: 7.8, y: 2.17 + i * 0.85, w: 2.0, h: 0.25, fontSize: 9, color: C.muted });
  });

  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.4, y: 5.55, w: 9.2, h: 0.75,
    fill: { color: '1A1200' },
    line: { color: C.gold, width: 1 },
    rectRadius: 0.08,
  });
  sl.addText('Pass is not a cost center — it is a demand-generation engine. The credits inDrive gives back are funded by the GMV of rides already completed. Every percentage point of cashback is recouped many times over in reduced re-acquisition spend and cross-vertical GMV.', {
    x: 0.6, y: 5.65, w: 8.8, h: 0.55,
    fontSize: 10.5, color: C.white, lineSpacingMultiple: 1.3,
  });
}

// ─────────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('PROOF OF WORK', {
    x: 0.5, y: 0.4, w: 9, h: 0.3,
    fontSize: 9, color: C.bronze, bold: true, charSpacing: 2,
  });
  sl.addText('I Built This. Milestone Incentives, Cross-Category Adoption,\nCredit-Based Retention — All Shipped at PhonePe.', {
    x: 0.5, y: 0.75, w: 9, h: 0.8,
    fontSize: 23, color: C.dark, bold: true, lineSpacingMultiple: 1.2,
  });

  // Left: PhonePe
  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.4, y: 1.75, w: 4.4, h: 4.3,
    fill: { color: C.dark },
    line: { color: C.gold, width: 1 },
    rectRadius: 0.1,
    shadow: makeShadow(),
  });
  sl.addText('What I built at PhonePe', {
    x: 0.6, y: 1.9, w: 4.0, h: 0.35,
    fontSize: 12, color: C.gold, bold: true,
  });
  const leftPts = [
    'Milestone incentive engine: B2B merchants earning fee waivers by hitting TPV + TPC goals — same milestone-to-reward architecture as inDrive Pass tier progression',
    'ML propensity scoring replacing ₹1,000+ Cr/yr static rules — identical to Pass tier segmentation: higher-frequency users get higher cashback, lower-frequency get activation offers',
    'Dynamic cart incentivization on Pincode: AOV milestone engine with context-aware triggers — same as Pass\'s "1 ride to Gold" nudge at near-milestone moment',
    'Rebuilt ₹100 Cr/yr rewards portfolio with audience segmentation and brand self-serve tooling — same multi-partner cross-vertical engagement logic as Pass Explorer milestones',
    'A/B tested Rent Waiver vs Cashback mechanic: Waiver drove 34% higher goal completion — directly validates Pass cashback architecture over flat discount',
  ];
  leftPts.forEach((pt, i) => {
    sl.addText(`• ${pt}`, {
      x: 0.6, y: 2.35 + i * 0.68, w: 4.0, h: 0.62,
      fontSize: 9.5, color: C.muted, lineSpacingMultiple: 1.3,
    });
  });

  // Right: JD mapping
  sl.addShape(prs.ShapeType.roundRect, {
    x: 5.2, y: 1.75, w: 4.4, h: 4.3,
    fill: { color: 'F8F9FA' },
    line: { color: C.accent, width: 1 },
    rectRadius: 0.1,
    shadow: makeShadow(),
  });
  sl.addText('Maps to inDrive JD requirements', {
    x: 5.4, y: 1.9, w: 4.0, h: 0.35,
    fontSize: 12, color: C.accent, bold: true,
  });
  const rightPts = [
    '"Build growth loops for the product portfolio" → Pass is inDrive\'s first closed growth loop: ride → earn credits → ride again → tier up → adopt new vertical → Explorer',
    '"Translate behavior insights into segmented programmable journeys" → Pass tier segmentation by ride frequency; Explorer milestone by cross-vertical behavior signal',
    '"Cross sell, upsell, and retention" → Explorer challenge is the cross-sell mechanic; Gold tier is the upsell; credits balance is the retention mechanism',
    '"Partner with ML/DS for personalization at scale" → Tier qualification, milestone nudge timing, and credit calibration all require ML signals I have shipped in production',
    '"Identify, size and prioritize growth opportunities" → Pass addresses the highest-ROI retention gap; sized against re-acquisition CAC and cross-vertical CLTV uplift',
  ];
  rightPts.forEach((pt, i) => {
    sl.addText(`• ${pt}`, {
      x: 5.4, y: 2.35 + i * 0.68, w: 4.0, h: 0.62,
      fontSize: 9.5, color: '2D3748', lineSpacingMultiple: 1.3,
    });
  });

  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.4, y: 6.15, w: 9.2, h: 0.65,
    fill: { color: C.dark },
    line: { color: C.gold, width: 1 },
    rectRadius: 0.07,
  });
  sl.addText('"The milestone-to-reward engine, the tier progression, the cross-category adoption mechanic — these are not concepts. I shipped each of them at PhonePe across 350M MAU."', {
    x: 0.6, y: 6.23, w: 8.8, h: 0.5,
    fontSize: 10.5, color: C.white, italic: true, lineSpacingMultiple: 1.3,
  });
}

// ─────────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('ROLLOUT PLAN', {
    x: 0.5, y: 0.4, w: 9, h: 0.3,
    fontSize: 9, color: C.gold, bold: true, charSpacing: 2,
  });
  sl.addText('3 Phases — From Bronze Pilot to Full Explorer Launch', {
    x: 0.5, y: 0.75, w: 9, h: 0.5,
    fontSize: 24, color: C.white, bold: true,
  });

  const phases = [
    { ph:'Phase 1', time:'Month 1–2', title:'Bronze + Silver Pilot', bullets:['Launch inDrive Pass in 1 market: Bronze (5 rides / 5% back) + Silver (10 rides / 8% back)', 'Instrument: Pass subscriber vs non-subscriber retention, ride frequency, churn at 30/60/90 days', 'A/B test: Pass offer shown on home vs organic discovery only', 'Calibrate cashback % against GMV data to ensure positive unit economics'], color: C.silver },
    { ph:'Phase 2', time:'Month 3–4', title:'Gold + Explorer Mechanic', bullets:['Launch Gold tier (15 rides / 12% back) + BidBand Band 3 default for Gold users', 'Deploy cross-vertical milestone tracker: Freight + Services pillars', 'Tune milestone nudge timing (ride 13, 14, 15 proximity nudges)', 'Expand to 3 markets using pilot learnings on cashback calibration'], color: C.gold },
    { ph:'Phase 3', time:'Month 5–6', title:'Explorer + Global Rollout', bullets:['Launch Explorer tier (All-Access / 15% back / priority matching / 6-mo rollover)', 'Integrate Pass credits across Rides, Freight, and Services in one unified wallet', 'Roll out globally — adapt cashback % by market based on average fare data', 'Pass as onboarding hook: first-ride Pass trial to reduce new user churn'], color: C.bright },
  ];

  phases.forEach((ph, i) => {
    const x = 0.4 + i * 3.2;
    sl.addShape(prs.ShapeType.roundRect, {
      x, y: 1.45, w: 2.9, h: 3.9,
      fill: { color: '141810' },
      line: { color: ph.color, width: 1.5 },
      rectRadius: 0.1,
      shadow: makeShadow(),
    });
    sl.addText(ph.ph, { x, y: 1.55, w: 2.9, h: 0.3, fontSize: 9, color: ph.color, bold:true, align:'center', charSpacing:1 });
    sl.addText(ph.time, { x, y: 1.85, w: 2.9, h: 0.25, fontSize: 9, color: C.muted, align:'center' });
    sl.addText(ph.title, { x: x+0.12, y: 2.15, w: 2.66, h: 0.35, fontSize: 12, color: C.white, bold:true, align:'center' });
    ph.bullets.forEach((b, j) => {
      sl.addText(`• ${b}`, { x: x+0.15, y: 2.6 + j * 0.62, w: 2.6, h: 0.55, fontSize: 9, color: C.muted, lineSpacingMultiple:1.3 });
    });
  });

  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.4, y: 5.55, w: 9.2, h: 1.1,
    fill: { color: '0D1200' },
    line: { color: C.gold, width: 1.2 },
    rectRadius: 0.1,
  });
  sl.addText('What I need to build this:', {
    x: 0.6, y: 5.65, w: 9.0, h: 0.28,
    fontSize: 11, color: C.gold, bold: true,
  });
  sl.addText('Ride frequency + GMV data per user cohort  ·  DS partner for cashback unit economics calibration  ·  1 backend eng for credits wallet + tier API  ·  Alignment with Freight + Services PMs on cross-vertical milestone integration', {
    x: 0.6, y: 5.95, w: 8.8, h: 0.55,
    fontSize: 10, color: C.white, lineSpacingMultiple: 1.35,
  });

  sl.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  ajay-avaghade.github.io/Portfolio', {
    x: 0.5, y: 7.1, w: 9.0, h: 0.25,
    fontSize: 9, color: C.muted, align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/indrive-pass-deck.pptx' })
  .then(() => console.log('indrive-pass-deck.pptx created'))
  .catch(e => console.error(e));
