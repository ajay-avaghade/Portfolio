'use strict';
(async () => {
const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:'131921', hero:'232F3E', bright:'FF9900',
  blue:'007185', prime:'00A8E1', white:'FFFFFF',
  lgray:'F3F4F6', gold:'FEBD69', green:'067D62', ink:'0F172A', muted:'64748B',
};
const mk = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });
const W=10, H=7.5;

// SLIDE 1 — COVER
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  for(let i=0;i<14;i++){
    s.addShape('rect', { x:-0.5+i*0.85, y:0, w:0.5, h:H, fill:{color:C.hero}, line:{type:'none'}, rotate:12 });
  }
  s.addShape('rect', { x:7.6, y:0, w:2.4, h:1.1, fill:{color:C.prime}, line:{type:'none'} });
  s.addText('amazon.ae', { x:7.65, y:0.1, w:2.3, h:0.4, fontSize:14, bold:true, color:C.dark, fontFace:'Calibri', align:'center' });
  s.addText('HOME & LEISURE · MENA', { x:7.65, y:0.52, w:2.3, h:0.28, fontSize:8, bold:true, color:C.dark, fontFace:'Calibri', align:'center', charSpacing:2 });
  s.addText('InstallMatch', { x:0.5, y:1.4, w:7, h:1.1, fontSize:52, bold:true, color:C.white, fontFace:'Calibri' });
  s.addShape('rect', { x:0.5, y:2.6, w:1.4, h:0.08, fill:{color:C.prime}, line:{type:'none'} });
  s.addText('In-flow installation marketplace — verified UAE technicians booked alongside the order', {
    x:0.5, y:2.75, w:8.5, h:0.5, fontSize:17, color:C.gold, fontFace:'Calibri', italic:true
  });
  s.addText('Case Study 03 of 03  ·  Amazon MENA Senior PM, Home & Leisure', {
    x:0.5, y:3.35, w:7, h:0.3, fontSize:10, color:C.muted, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:5.75, w:5, h:0.06, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('Presented by Ajay Avaghade  ·  Growth & Category PM  ·  PhonePe → Amazon MENA', {
    x:0.5, y:5.9, w:7, h:0.35, fontSize:11, color:C.muted, fontFace:'Calibri'
  });
  s.addShape('rect', { x:7.6, y:5.1, w:2.1, h:2.0, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('100%', { x:7.6, y:5.2, w:2.1, h:0.8, fontSize:36, bold:true, color:C.prime, fontFace:'Calibri', align:'center' });
  s.addText('installation\ndemand lost\nto Google', { x:7.6, y:5.95, w:2.1, h:0.9, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center' });
}

// SLIDE 2 — THE PROBLEM
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('THE PROBLEM', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText("Amazon Sells the AC — Then Sends the Customer to Google", {
    x:0.5, y:0.62, w:9, h:0.9, fontSize:26, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.6, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });
  [
    { icon:'🔍', stat:'Google', lbl:"First thing after checkout", sub:'Every AC, dishwasher, ceiling fan\nbuyer immediately searches\n"AC installation Dubai" on Google' },
    { icon:'⚠️', stat:'Unverified', lbl:'Technicians they find', sub:'No background checks, no\ninsurance, no accountability —\nAmazon product gets damaged' },
    { icon:'📦', stat:'15%', lbl:'Return rate on large H&L', sub:'Items returned because customer\ncould not get installation — lost\nrevenue + reverse logistics cost' },
  ].forEach((col,i) => {
    const x = 0.5+i*3.1;
    s.addShape('rect', { x, y:1.78, w:2.9, h:3.5, fill:{color:C.hero}, line:{type:'none'}, shadow:mk() });
    s.addText(col.icon, { x, y:1.92, w:2.9, h:0.6, fontSize:28, align:'center' });
    s.addText(col.stat, { x, y:2.58, w:2.9, h:0.8, fontSize:34, bold:true, color:C.prime, fontFace:'Calibri', align:'center' });
    s.addText(col.lbl, { x:x+0.1, y:3.42, w:2.7, h:0.4, fontSize:12, bold:true, color:C.white, fontFace:'Calibri', align:'center' });
    s.addText(col.sub, { x:x+0.1, y:3.85, w:2.7, h:1.0, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center' });
  });
  s.addShape('rect', { x:0.5, y:5.48, w:9, h:1.8, fill:{color:'1a2332'}, line:{color:C.prime, width:1.5}, shadow:mk() });
  s.addText('Root Cause:', { x:0.7, y:5.62, w:2, h:0.35, fontSize:11, bold:true, color:C.prime, fontFace:'Calibri' });
  s.addText('Amazon closes the sale and immediately loses the customer relationship. They search for a technician, find one from a WhatsApp group, the installation goes wrong, the product gets damaged, and the return lands back at Amazon. InstallMatch keeps that entire post-purchase journey inside Amazon — and monetizes it.', {
    x:0.7, y:5.98, w:8.6, h:1.1, fontSize:11, color:C.muted, fontFace:'Calibri'
  });
}

// SLIDE 3 — THE INSIGHT
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.lgray} });
  s.addText('THE INSIGHT', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText("The Install Is Part of the Product. Amazon Just Does Not Sell It Yet.", {
    x:0.5, y:0.62, w:9, h:0.85, fontSize:25, bold:true, color:C.ink, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.52, w:9, h:0.04, fill:{color:C.prime}, line:{type:'none'} });
  s.addShape('rect', { x:0.5, y:1.72, w:4.2, h:5.2, fill:{color:C.white}, line:{color:'E5E7EB', width:1}, shadow:mk() });
  s.addText('❌  The Status Quo', { x:0.65, y:1.9, w:3.9, h:0.42, fontSize:14, bold:true, color:C.ink, fontFace:'Calibri' });
  [
    'Order confirmed. No installation path offered.',
    'Customer searches "AC installation Dubai" — leaves Amazon ecosystem',
    'Finds uninsured technician from WhatsApp group — product gets damaged',
    'Returns the unit to Amazon. 15% of H&L large-ticket orders follow this path.',
  ].forEach((t,i) => {
    s.addShape('rect', { x:0.65, y:2.48+i*0.76, w:0.22, h:0.22, fill:{color:'FEE2E2'}, line:{type:'none'} });
    s.addText('✕', { x:0.65, y:2.48+i*0.76, w:0.22, h:0.22, fontSize:10, bold:true, color:'DC2626', align:'center' });
    s.addText(t, { x:0.95, y:2.44+i*0.76, w:3.55, h:0.6, fontSize:11, color:C.muted, fontFace:'Calibri' });
  });
  s.addShape('ellipse', { x:4.55, y:3.6, w:0.9, h:0.9, fill:{color:C.prime}, line:{type:'none'} });
  s.addText('VS', { x:4.55, y:3.6, w:0.9, h:0.9, fontSize:14, bold:true, color:C.white, align:'center', valign:'middle', fontFace:'Calibri' });
  s.addShape('rect', { x:5.3, y:1.72, w:4.2, h:5.2, fill:{color:C.white}, line:{color:C.prime, width:2}, shadow:mk() });
  s.addText('✅  InstallMatch', { x:5.45, y:1.9, w:3.9, h:0.42, fontSize:14, bold:true, color:C.ink, fontFace:'Calibri' });
  [
    '36 verified, insured, background-checked UAE technicians — embedded on the PDP',
    'Customer books Mohammed Al Rashidi (4.9★, 2.1km) in the same checkout flow',
    'Installer tracked live on install day. 30-day workmanship guarantee.',
    'AED 149 added to order. Amazon earns commission. Returns drop 32%.',
  ].forEach((t,i) => {
    s.addShape('rect', { x:5.45, y:2.48+i*0.76, w:0.22, h:0.22, fill:{color:'DBEAFE'}, line:{type:'none'} });
    s.addText('✓', { x:5.45, y:2.48+i*0.76, w:0.22, h:0.22, fontSize:10, bold:true, color:C.prime, align:'center' });
    s.addText(t, { x:5.75, y:2.44+i*0.76, w:3.55, h:0.6, fontSize:11, color:C.muted, fontFace:'Calibri' });
  });
  s.addShape('rect', { x:0.5, y:7.05, w:9, h:0.3, fill:{color:'EFF8FF'}, line:{color:C.prime, width:1.5} });
  s.addText('Key Insight: The customer has AED 2,499 committed to the AC. AED 149 for verified install is not a spend decision — it is a relief decision.', {
    x:0.65, y:7.07, w:8.7, h:0.24, fontSize:9.5, color:C.ink, fontFace:'Calibri', italic:true
  });
}

// SLIDE 4 — THE MECHANIC
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('THE MECHANIC', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('Curated Marketplace × In-Flow Booking × Live Tracking = InstallMatch', {
    x:0.5, y:0.62, w:9, h:0.7, fontSize:23, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.38, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });
  [
    { n:'01', title:'PDP Trigger', body:'Product category triggers InstallMatch card. "36 verified technicians near Dubai Marina." Appears below Buy Now.' },
    { n:'02', title:'Installer Listing', body:'Sorted by: rating, distance, availability. Each card: name, stars, reviews, distance, languages, jobs done.' },
    { n:'03', title:'Slot Booking', body:'Date picker + time grid. Installer calendar synced in real-time. Address auto-filled from delivery address.' },
    { n:'04', title:'Bundled Checkout', body:'AC + install as single order. AED 149 line item. Delivery and install dates shown. One payment, one confirmation.' },
    { n:'05', title:'Live Tracking', body:'Install day: technician live on map. ETA shown. Progress steps (ordered → delivered → en route → installing → done).' },
  ].forEach((step,i) => {
    const x = 0.38+i*1.88;
    s.addShape('ellipse', { x:x+0.52, y:1.62, w:0.72, h:0.72, fill:{color:C.prime}, line:{type:'none'}, shadow:mk() });
    s.addText(step.n, { x:x+0.52, y:1.62, w:0.72, h:0.72, fontSize:13, bold:true, color:C.white, align:'center', valign:'middle', fontFace:'Calibri' });
    if(i<4) s.addShape('line', { x:x+1.24, y:1.98, w:0.64, h:0, line:{color:C.prime, width:1.5, dashType:'dash'} });
    s.addShape('rect', { x, y:2.52, w:1.78, h:3.55, fill:{color:C.hero}, line:{type:'none'}, shadow:mk() });
    s.addText(step.title, { x:x+0.05, y:2.66, w:1.68, h:0.45, fontSize:12, bold:true, color:C.white, fontFace:'Calibri', align:'center' });
    s.addText(step.body, { x:x+0.05, y:3.16, w:1.68, h:2.55, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center' });
  });
  s.addShape('rect', { x:0.5, y:6.32, w:9, h:0.95, fill:{color:'1a2332'}, line:{color:C.prime, width:1}, shadow:mk() });
  s.addText('Marketplace Design  ·  ', { x:0.7, y:6.46, w:2.4, h:0.4, fontSize:11, bold:true, color:C.prime, fontFace:'Calibri' });
  s.addText('At PhonePe, I built a B2B marketplace onboarding 5,000+ merchants with self-serve eligibility, pricing, and ops tools. InstallMatch applies the same architecture to a services marketplace — technicians as the supply side, homeowners as demand.', {
    x:3.0, y:6.44, w:6.3, h:0.62, fontSize:10, color:C.muted, fontFace:'Calibri'
  });
}

// SLIDE 5 — THE PRODUCT
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.lgray} });
  s.addText('THE PRODUCT', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('5 Screens — AC Order to Installation Day', {
    x:0.5, y:0.62, w:9, h:0.7, fontSize:28, bold:true, color:C.ink, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.38, w:9, h:0.04, fill:{color:C.prime}, line:{type:'none'} });
  [
    { n:'01', t:'PDP + Install Card', d:'InstallMatch card below Buy Now. "36 verified technicians near Dubai Marina." Bullet: AED 149+, same-day slots, 30-day guarantee.', m:['Daikin 1.5T Inverter', 'AED 2,499 22% off', '✦ prime | Free delivery', '────────────', 'INSTALLMATCH', '36 installers nearby', 'AED 149+, insured', '[Choose Installer]'] },
    { n:'02', t:'Installer Selection', d:'3 cards sorted by rating + distance. Mohammed Al Rashidi: 4.9★, 847 reviews, Amazon Certified Pro, today slots.', m:['Choose Installer (36)', '⭐ TOP MATCH', 'Mohammed Al Rashidi', '★★★★★ 4.9 (847)', '✓ Amazon Certified Pro', '2.1km | Today slots', 'AED 149 [Select →]', 'Ahmed Hassan 4.7★'] },
    { n:'03', t:'Slot Booking', d:'Date chips + time grid. Booked slots greyed out. Address pre-filled. Room type selector. Confirm sticky bar.', m:['Pick a Time Slot', 'Mohammed · AED 149', 'TUE 24 | WED 25 | FRI 27', '────────────', '9-11am [BOOKED]', '2-5pm [SELECT] ◉', '5-8pm [ ]', 'Living Room ◉'] },
    { n:'04', t:'Order Confirmed', d:'Confetti. AC delivery date + installer appointment in one confirmation. Total: AED 2,648. Track both in one view.', m:['Order + Install ✓', '#406-7920341', '────────────', 'Daikin AC → Thu 26', 'Mohammed → Tue 24', '2:00–5:00 PM', 'Total: AED 2,648', '[Track →]'] },
    { n:'05', t:'Live Tracking', d:'Install day. Mohammed on live map, ETA 18 min. Progress steps. "Call" and "Message" buttons. Arrival notification.', m:['Live Tracking', 'Mohammed on way!', 'ETA ~18 min | 3.2km', '────────────', '✓ Order confirmed', '✓ AC delivered', '→ Mohammed en route', '○ Installation pending'] },
  ].forEach((sc,i) => {
    const x = 0.3+i*1.9;
    s.addShape('rect', { x, y:1.55, w:1.78, h:5.7, fill:{color:C.white}, line:{color:'E5E7EB', width:1}, shadow:mk() });
    s.addShape('rect', { x, y:1.55, w:1.78, h:0.42, fill:{color:C.dark}, line:{type:'none'} });
    s.addText(`${sc.n} — ${sc.t}`, { x:x+0.06, y:1.6, w:1.66, h:0.32, fontSize:8, bold:true, color:C.gold, fontFace:'Calibri' });
    s.addText(sc.d, { x:x+0.06, y:2.02, w:1.66, h:1.05, fontSize:8.5, color:C.muted, fontFace:'Calibri' });
    s.addShape('rect', { x:x+0.06, y:3.12, w:1.66, h:3.88, fill:{color:'F8FAFC'}, line:{color:'E5E7EB', width:1} });
    s.addText(sc.m.join('\n'), { x:x+0.1, y:3.17, w:1.58, h:3.78, fontSize:7.5, color:C.ink, fontFace:'DM Mono', valign:'top' });
  });
}

// SLIDE 6 — IMPACT & ROI
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('IMPACT & ROI', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('Stop the Demand Leak — Monetize the Full H&L Journey', {
    x:0.5, y:0.62, w:9, h:0.7, fontSize:26, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.38, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });
  s.addShape('rect', { x:0.5, y:1.56, w:4.3, h:0.4, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('SHOPPER VALUE', { x:0.5, y:1.56, w:4.3, h:0.4, fontSize:10, bold:true, color:C.gold, align:'center', valign:'middle', fontFace:'Calibri' });
  [
    { v:'AED 149', l:'Installation cost', s:'vs AED 250-400 for unverified technician found on Google' },
    { v:'30-day', l:'Workmanship guarantee', s:'Free rebooking if install has any defect' },
    { v:'<24hr', l:'Same-day slots available', s:'If ordered before 12 PM on working days' },
    { v:'Live', l:'Technician tracking', s:'Real-time map view + ETA + call/message buttons' },
  ].forEach((m,i) => {
    const y = 2.06+i*1.15;
    s.addShape('rect', { x:0.5, y, w:4.3, h:1.0, fill:{color:C.hero}, line:{type:'none'}, shadow:mk() });
    s.addText(m.v, { x:0.6, y:y+0.07, w:4.1, h:0.52, fontSize:28, bold:true, color:C.prime, fontFace:'Calibri' });
    s.addText(m.l, { x:0.6, y:y+0.57, w:2.4, h:0.3, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(m.s, { x:0.6, y:y+0.57, w:4.1, h:0.3, fontSize:10, color:C.muted, fontFace:'Calibri', align:'right' });
  });
  s.addShape('line', { x:4.95, y:1.56, w:0, h:5.7, line:{color:C.hero, width:1, dashType:'dash'} });
  s.addShape('rect', { x:5.1, y:1.56, w:4.4, h:0.4, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('AMAZON H&L ROI', { x:5.1, y:1.56, w:4.4, h:0.4, fontSize:10, bold:true, color:C.gold, align:'center', valign:'middle', fontFace:'Calibri' });
  [
    { v:'18%', l:'Install attach rate (target)', s:'On appliances requiring installation' },
    { v:'AED 149', l:'Avg install revenue per order', s:'Commission-based — Amazon takes 15-25%' },
    { v:'32%', l:'Returns reduction', s:'Installed products have 32% lower return rate' },
    { v:'NPS +12', l:'Post-purchase NPS lift', s:'Customers with successful install rate higher' },
  ].forEach((m,i) => {
    const y = 2.06+i*1.15;
    s.addShape('rect', { x:5.1, y, w:4.4, h:1.0, fill:{color:C.hero}, line:{type:'none'}, shadow:mk() });
    s.addText(m.v, { x:5.2, y:y+0.07, w:4.2, h:0.52, fontSize:28, bold:true, color:C.bright, fontFace:'Calibri' });
    s.addText(m.l, { x:5.2, y:y+0.57, w:2.7, h:0.3, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(m.s, { x:5.2, y:y+0.57, w:4.2, h:0.3, fontSize:10, color:C.muted, fontFace:'Calibri', align:'right' });
  });
  s.addShape('rect', { x:0.5, y:6.74, w:9, h:0.6, fill:{color:'1a2332'}, line:{color:C.prime, width:1} });
  s.addText('InstallMatch turns a post-purchase pain point into a structured revenue stream — and stops 15% of H&L orders from becoming returns.', {
    x:0.65, y:6.82, w:8.7, h:0.38, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true
  });
}

// SLIDE 7 — PROOF OF WORK
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.lgray} });
  s.addText('PROOF OF WORK', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText("Built a B2B Marketplace at Scale. InstallMatch Is the Same Architecture.", {
    x:0.5, y:0.62, w:9, h:0.7, fontSize:24, bold:true, color:C.ink, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.38, w:9, h:0.04, fill:{color:C.prime}, line:{type:'none'} });
  s.addShape('rect', { x:0.5, y:1.56, w:4.3, h:5.1, fill:{color:C.dark}, line:{type:'none'}, shadow:mk() });
  s.addText('WHAT I BUILT AT PHONEPE', { x:0.65, y:1.7, w:4.0, h:0.35, fontSize:9, bold:true, color:C.bright, fontFace:'Calibri', charSpacing:3 });
  [
    'Built 0→1 self-serve PG integration platform: 5,000+ B2B merchants acquired, 23% lower CAC, 5M+ new users/month',
    'Designed marketplace onboarding: merchant discovery, eligibility check, pricing, legal sign-off, go-live — all self-serve',
    'Built ops intelligence layer replacing 6 siloed workflows — TAT from 2 days to 4 hours for merchant support tickets',
    'Onboarded 500+ brand partners to ML-driven rewards marketplace — same multi-sided platform problem as InstallMatch',
    'Owned full BRD-to-launch across payment, merchant, and partner tracks simultaneously — cross-functional at scale',
  ].forEach((item,i) => {
    s.addShape('rect', { x:0.65, y:2.2+i*0.74, w:0.2, h:0.2, fill:{color:C.bright}, line:{type:'none'} });
    s.addText(item, { x:0.92, y:2.16+i*0.74, w:3.75, h:0.62, fontSize:10, color:C.muted, fontFace:'Calibri' });
  });
  s.addShape('rect', { x:5.1, y:1.56, w:4.4, h:5.1, fill:{color:C.white}, line:{color:'E5E7EB', width:1}, shadow:mk() });
  s.addText('HOW IT MAPS TO INSTALLMATCH', { x:5.25, y:1.7, w:4.1, h:0.35, fontSize:9, bold:true, color:C.ink, fontFace:'Calibri', charSpacing:2 });
  [
    ['PhonePe: B2B marketplace, 5K merchants', '→ InstallMatch: technician marketplace, UAE'],
    ['PhonePe: Self-serve onboarding + ops', '→ InstallMatch: technician vetting + calendar sync'],
    ['PhonePe: Eligibility + payout engine', '→ InstallMatch: booking + commission payout'],
    ['PhonePe: 30-min ops TAT (was 2 days)', '→ InstallMatch: same-day install scheduling'],
    ['PhonePe: 0→1 platform, multi-partner', '→ InstallMatch: 0→1 services marketplace'],
  ].forEach((m,i) => {
    const y = 2.2+i*0.74;
    s.addShape('rect', { x:5.25, y, w:4.1, h:0.64, fill:{color:C.lgray}, line:{type:'none'} });
    s.addText(m[0], { x:5.35, y:y+0.06, w:3.9, h:0.22, fontSize:10, bold:true, color:C.ink, fontFace:'Calibri' });
    s.addText(m[1], { x:5.35, y:y+0.32, w:3.9, h:0.22, fontSize:10, color:C.prime, fontFace:'Calibri' });
  });
  s.addShape('rect', { x:0.5, y:6.88, w:9, h:0.46, fill:{color:'EFF8FF'}, line:{color:C.prime, width:1.5} });
  s.addText('"InstallMatch is the PhonePe merchant marketplace — applied to UAE installation services."  — Ajay Avaghade', {
    x:0.7, y:6.94, w:8.6, h:0.3, fontSize:11, color:C.ink, fontFace:'Calibri', italic:true
  });
}

// SLIDE 8 — ROLLOUT PLAN
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('ROLLOUT PLAN', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('From Dubai Pilot to MENA Installation Platform', {
    x:0.5, y:0.62, w:9, h:0.7, fontSize:28, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.38, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });
  [
    { label:'PHASE 1', period:'Month 1–2', title:'Dubai AC Pilot', color:C.hero, acc:C.gold, items:['InstallMatch for AC units — Dubai only', 'Onboard 50 background-verified technicians', 'Slot booking + live tracking MVP', 'KPIs: attach rate, NPS, return rate on installs'] },
    { label:'PHASE 2', period:'Month 3–4', title:'H&L Expansion', color:C.prime, acc:C.white, items:['Expand to all installable H&L: dishwashers, fans, TVs', 'Launch technician rating system + review prompts', 'Add WhatsApp updates: "Mohammed is 10 min away"', 'A/B test: install upsell on PDP vs. post-checkout'] },
    { label:'PHASE 3', period:'Month 5–6', title:'MENA Scale', color:C.hero, acc:C.gold, items:['Roll out to KSA, Kuwait, Egypt markets', 'Launch InstallMatch for furniture assembly category', 'Technician SLA dashboard for Amazon ops team', 'Publish MENA installation marketplace playbook'] },
  ].forEach((ph,i) => {
    const x = 0.4+i*3.1;
    s.addShape('rect', { x, y:1.56, w:2.95, h:4.55, fill:{color:ph.color}, line:{type:'none'}, shadow:mk() });
    s.addText(ph.label, { x:x+0.1, y:1.68, w:2.75, h:0.28, fontSize:9, bold:true, color:ph.acc, fontFace:'Calibri', charSpacing:3 });
    s.addText(ph.period, { x:x+0.1, y:1.96, w:2.75, h:0.26, fontSize:9, color:ph.color===C.prime?'dbeafe':C.muted, fontFace:'Calibri' });
    s.addText(ph.title, { x:x+0.1, y:2.26, w:2.75, h:0.5, fontSize:17, bold:true, color:ph.acc, fontFace:'Calibri' });
    ph.items.forEach((item,j) => {
      s.addShape('ellipse', { x:x+0.12, y:2.9+j*0.74, w:0.12, h:0.12, fill:{color:ph.acc}, line:{type:'none'} });
      s.addText(item, { x:x+0.3, y:2.84+j*0.74, w:2.55, h:0.64, fontSize:10, color:ph.color===C.prime?'dbeafe':C.muted, fontFace:'Calibri' });
    });
  });
  s.addShape('rect', { x:0.5, y:6.35, w:9, h:0.95, fill:{color:'1a2332'}, line:{color:C.prime, width:1.5}, shadow:mk() });
  s.addText('What I need to build InstallMatch:', { x:0.7, y:6.44, w:4, h:0.35, fontSize:12, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('UAE technician vetting + onboarding pipeline  ·  Calendar + booking API  ·  Live tracking integration  ·  1 PM + 2 Engineers + 1 Marketplace Ops Manager', {
    x:0.7, y:6.77, w:8.6, h:0.38, fontSize:10, color:C.muted, fontFace:'Calibri'
  });
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  ajay-avaghade.github.io/Portfolio', {
    x:0, y:7.2, w:W, h:0.26, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center'
  });
}

await prs.writeFile({ fileName: 'assets/amazon-installmatch-deck.pptx' });
console.log('Done: amazon-installmatch-deck.pptx');
})();
