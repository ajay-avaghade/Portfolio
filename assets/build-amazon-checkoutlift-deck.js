'use strict';
(async () => {
const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:'131921', hero:'232F3E', bright:'FF9900',
  gold:'FEBD69', white:'FFFFFF', lgray:'F3F4F6',
  green:'067D62', blue:'007185', ink:'0F172A', muted:'64748B',
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
  s.addShape('rect', { x:7.6, y:0, w:2.4, h:1.1, fill:{color:C.bright}, line:{type:'none'} });
  s.addText('amazon.ae', { x:7.65, y:0.1, w:2.3, h:0.4, fontSize:14, bold:true, color:C.dark, fontFace:'Calibri', align:'center' });
  s.addText('HOME & LEISURE · MENA', { x:7.65, y:0.52, w:2.3, h:0.28, fontSize:8, bold:true, color:C.dark, fontFace:'Calibri', align:'center', charSpacing:2 });
  s.addText('CheckoutLift', { x:0.5, y:1.4, w:7, h:1.1, fontSize:52, bold:true, color:C.white, fontFace:'Calibri' });
  s.addShape('rect', { x:0.5, y:2.6, w:1.4, h:0.08, fill:{color:C.bright}, line:{type:'none'} });
  s.addText('Cart-aware steal deal engine — surfacing the right H&L product at checkout', {
    x:0.5, y:2.75, w:8.5, h:0.5, fontSize:17, color:C.gold, fontFace:'Calibri', italic:true
  });
  s.addText('Case Study 01 of 03  ·  Amazon MENA Senior PM, Home & Leisure', {
    x:0.5, y:3.35, w:7, h:0.3, fontSize:10, color:C.muted, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:5.75, w:5, h:0.06, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('Presented by Ajay Avaghade  ·  Growth & Category PM  ·  PhonePe → Amazon MENA', {
    x:0.5, y:5.9, w:7, h:0.35, fontSize:11, color:C.muted, fontFace:'Calibri'
  });
  s.addShape('rect', { x:7.6, y:5.1, w:2.1, h:2.0, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('65%', { x:7.6, y:5.2, w:2.1, h:0.8, fontSize:40, bold:true, color:C.bright, fontFace:'Calibri', align:'center' });
  s.addText('checkout\nabandonment\ntoday', { x:7.6, y:5.95, w:2.1, h:0.9, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center' });
}

// SLIDE 2 — THE PROBLEM
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('THE PROBLEM', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText("Amazon.ae Leaves AED Millions on the Table at Every Checkout", {
    x:0.5, y:0.62, w:9, h:0.9, fontSize:26, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.6, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });
  const cols = [
    { icon:'🛒', stat:'65%', lbl:'Cart abandonment rate', sub:'Industry average at\ncheckout entry — no\ncart-level intervention today' },
    { icon:'📦', stat:'0', lbl:'Cart-aware upsells', sub:'Checkout page is a list\nof items — not a\nsales surface' },
    { icon:'🎯', stat:'68%', lbl:'Cross-buy probability', sub:'Shoppers with a Nescafe\norder also buy a\nNespresso machine — untapped' },
  ];
  cols.forEach((col,i) => {
    const x = 0.5 + i*3.1;
    s.addShape('rect', { x, y:1.78, w:2.9, h:3.5, fill:{color:C.hero}, line:{type:'none'}, shadow:mk() });
    s.addText(col.icon, { x, y:1.92, w:2.9, h:0.6, fontSize:28, align:'center' });
    s.addText(col.stat, { x, y:2.58, w:2.9, h:0.8, fontSize:38, bold:true, color:C.bright, fontFace:'Calibri', align:'center' });
    s.addText(col.lbl, { x:x+0.1, y:3.42, w:2.7, h:0.4, fontSize:12, bold:true, color:C.white, fontFace:'Calibri', align:'center' });
    s.addText(col.sub, { x:x+0.1, y:3.85, w:2.7, h:1.0, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center' });
  });
  s.addShape('rect', { x:0.5, y:5.48, w:9, h:1.8, fill:{color:'1a2332'}, line:{color:C.bright, width:1.5}, shadow:mk() });
  s.addText('Root Cause:', { x:0.7, y:5.62, w:2, h:0.35, fontSize:11, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('Amazon knows what is in the cart but does nothing with that signal at checkout. A shopper with Nescafe capsules has a 68% likelihood of also wanting a Nespresso machine — but Amazon shows them a generic checkout with no contextual suggestion. The steal deal moment is given away to Google search, competitor apps, and Instagram ads every single day.', {
    x:0.7, y:5.98, w:8.6, h:1.1, fontSize:11, color:C.muted, fontFace:'Calibri'
  });
}

// SLIDE 3 — THE INSIGHT
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.lgray} });
  s.addText('THE INSIGHT', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText("The Cart Is a Purchase-Intent Signal — Amazon Just Isn't Reading It", {
    x:0.5, y:0.62, w:9, h:0.85, fontSize:25, bold:true, color:C.ink, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.52, w:9, h:0.04, fill:{color:C.bright}, line:{type:'none'} });
  s.addShape('rect', { x:0.5, y:1.72, w:4.2, h:5.2, fill:{color:C.white}, line:{color:'E5E7EB', width:1}, shadow:mk() });
  s.addText('❌  Checkout Today', { x:0.65, y:1.9, w:3.9, h:0.42, fontSize:14, bold:true, color:C.ink, fontFace:'Calibri' });
  [
    'Cart shows 3 items — page is a list, nothing more',
    'No cross-category recommendation at the highest-intent moment',
    'Countdown deals shown on homepage, never at checkout',
    'Customer abandons, searches on Google, buys from Noon',
  ].forEach((t,i) => {
    s.addShape('rect', { x:0.65, y:2.48+i*0.76, w:0.22, h:0.22, fill:{color:'FEE2E2'}, line:{type:'none'} });
    s.addText('✕', { x:0.65, y:2.48+i*0.76, w:0.22, h:0.22, fontSize:10, bold:true, color:'DC2626', align:'center' });
    s.addText(t, { x:0.95, y:2.44+i*0.76, w:3.55, h:0.6, fontSize:11, color:C.muted, fontFace:'Calibri' });
  });
  s.addShape('ellipse', { x:4.55, y:3.6, w:0.9, h:0.9, fill:{color:C.bright}, line:{type:'none'} });
  s.addText('VS', { x:4.55, y:3.6, w:0.9, h:0.9, fontSize:14, bold:true, color:C.dark, align:'center', valign:'middle', fontFace:'Calibri' });
  s.addShape('rect', { x:5.3, y:1.72, w:4.2, h:5.2, fill:{color:C.white}, line:{color:C.bright, width:2}, shadow:mk() });
  s.addText('✅  CheckoutLift', { x:5.45, y:1.9, w:3.9, h:0.42, fontSize:14, bold:true, color:C.ink, fontFace:'Calibri' });
  [
    'Cart scanned: Nescafe capsules detected in basket',
    'ML model: 68% of this cart type also buys Nespresso',
    'CheckoutLift surfaces: "Before you checkout — AED 79 off Nespresso (38 min left)"',
    'One-tap add, order confirmed, AED 246 saved',
  ].forEach((t,i) => {
    s.addShape('rect', { x:5.45, y:2.48+i*0.76, w:0.22, h:0.22, fill:{color:'DCFCE7'}, line:{type:'none'} });
    s.addText('✓', { x:5.45, y:2.48+i*0.76, w:0.22, h:0.22, fontSize:10, bold:true, color:C.green, align:'center' });
    s.addText(t, { x:5.75, y:2.44+i*0.76, w:3.55, h:0.6, fontSize:11, color:C.muted, fontFace:'Calibri' });
  });
  s.addShape('rect', { x:0.5, y:7.05, w:9, h:0.3, fill:{color:'FFFBEB'}, line:{color:C.bright, width:1.5} });
  s.addText('Key Insight: The steal deal only needs to be right once — cart composition makes the recommendation precise enough to feel personal, not generic.', {
    x:0.65, y:7.07, w:8.7, h:0.24, fontSize:9.5, color:C.ink, fontFace:'Calibri', italic:true
  });
}

// SLIDE 4 — THE MECHANIC
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('THE MECHANIC', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('Cart Composition × Purchase History × Countdown = CheckoutLift', {
    x:0.5, y:0.62, w:9, h:0.7, fontSize:24, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.38, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });
  const steps = [
    { n:'01', title:'Cart Scan', body:'At checkout entry, engine scans cart SKUs. Category + brand signals extracted in real-time (<50ms).' },
    { n:'02', title:'Co-buy ML Match', body:'Model identifies top H&L co-purchase for this cart type. Confidence threshold: >55% to surface.' },
    { n:'03', title:'Deal Eligibility', body:'Checks inventory, discount budget, and countdown window. Only deals with ≥15% off + stock>10 units qualify.' },
    { n:'04', title:'Countdown Surface', body:'CheckoutLift card appears below cart items. Timer (30–60 min) creates urgency. AED savings shown prominently.' },
    { n:'05', title:'One-Tap Add + Learn', body:'Customer adds in one tap. Cart updates. Model records accept/skip signal for weekly retraining.' },
  ];
  steps.forEach((step,i) => {
    const x = 0.38 + i*1.88;
    s.addShape('ellipse', { x:x+0.52, y:1.62, w:0.72, h:0.72, fill:{color:C.bright}, line:{type:'none'}, shadow:mk() });
    s.addText(step.n, { x:x+0.52, y:1.62, w:0.72, h:0.72, fontSize:13, bold:true, color:C.dark, align:'center', valign:'middle', fontFace:'Calibri' });
    if(i<4) s.addShape('line', { x:x+1.24, y:1.98, w:0.64, h:0, line:{color:C.bright, width:1.5, dashType:'dash'} });
    s.addShape('rect', { x, y:2.52, w:1.78, h:3.55, fill:{color:C.hero}, line:{type:'none'}, shadow:mk() });
    s.addText(step.title, { x:x+0.05, y:2.66, w:1.68, h:0.45, fontSize:12, bold:true, color:C.white, fontFace:'Calibri', align:'center' });
    s.addText(step.body, { x:x+0.05, y:3.16, w:1.68, h:2.55, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center' });
  });
  s.addShape('rect', { x:0.5, y:6.32, w:9, h:0.95, fill:{color:'1a2332'}, line:{color:C.bright, width:1}, shadow:mk() });
  s.addText('PhonePe Proof  ·  ', { x:0.7, y:6.46, w:2.1, h:0.4, fontSize:11, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('Dynamic cart incentivization engine (Pincode, q-commerce): cart value x user intent x time signals → 35% AOV uplift, 60% cart abandonment reduction, 20% 30-day retention improvement. CheckoutLift applies the exact same mechanic to Amazon H&L large-ticket items.', {
    x:2.7, y:6.44, w:6.6, h:0.62, fontSize:10, color:C.muted, fontFace:'Calibri'
  });
}

// SLIDE 5 — THE PRODUCT
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.lgray} });
  s.addText('THE PRODUCT', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('5 Screens — End-to-End Steal Deal Journey', {
    x:0.5, y:0.62, w:9, h:0.7, fontSize:28, bold:true, color:C.ink, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.38, w:9, h:0.04, fill:{color:C.bright}, line:{type:'none'} });
  const screens = [
    { n:'01', t:'Home Feed', d:'CheckoutLift steal deal banner on home. Nespresso deal surfaced because Nescafe capsules bought last week.', m:['amazon.ae home', '────────────', 'CHECKOUTLIFT', 'Nespresso deal', 'based on your cart', 'AED 79 off today', '⏱ 38 min left', '[Shop Now]'] },
    { n:'02', t:'Cart + Deal Card', d:'"Before you checkout" card below 3 cart items. Countdown timer. AED savings front and center.', m:['Shopping Cart (3)', '────────────', 'Nescafe capsules', 'AED 89.00', '────────────', 'CHECKOUTLIFT', 'Nespresso AED 299', 'Save AED 246 ⏱38m', '[Add to Cart]'] },
    { n:'03', t:'Deal Detail', d:'Bottom sheet. Why recommended, ratings, features, CTA. "68% of shoppers with your cart also bought this."', m:['Nespresso Vertuo', '★★★★☆ 4.4 (891)', 'AED 299 (was 545)', '────────────', 'WHY THIS?', 'You have Nescafe', 'capsules in cart.', '[Add — AED 299]'] },
    { n:'04', t:'Cart Updated', d:'Deal added toast. Cart shows 4 items with steal deal highlighted and savings callout.', m:['Cart (4 items)', '✓ You saved AED 246', '────────────', 'Nespresso Vertuo', 'STEAL DEAL -AED246', 'AED 299.00', '────────────', '[Checkout]'] },
    { n:'05', t:'Order Confirmed', d:'Confetti. CheckoutLift savings callout. Total saved displayed prominently alongside order details.', m:['Order Placed! 🎉', '────────────', 'CHECKOUTLIFT', 'STEAL DEAL SAVED', 'AED 246 today', '────────────', 'Order total: 1,132', '[Track Order]'] },
  ];
  screens.forEach((sc,i) => {
    const x = 0.3 + i*1.9;
    s.addShape('rect', { x, y:1.55, w:1.78, h:5.7, fill:{color:C.white}, line:{color:'E5E7EB', width:1}, shadow:mk() });
    s.addShape('rect', { x, y:1.55, w:1.78, h:0.42, fill:{color:C.dark}, line:{type:'none'} });
    s.addText(`${sc.n} — ${sc.t}`, { x:x+0.06, y:1.6, w:1.66, h:0.32, fontSize:8.5, bold:true, color:C.gold, fontFace:'Calibri' });
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
  s.addText('PhonePe Proof Points Scaled to Amazon MENA H&L', {
    x:0.5, y:0.62, w:9, h:0.7, fontSize:26, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.38, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });
  s.addShape('rect', { x:0.5, y:1.56, w:4.3, h:0.4, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('SHOPPER IMPACT', { x:0.5, y:1.56, w:4.3, h:0.4, fontSize:10, bold:true, color:C.gold, align:'center', valign:'middle', fontFace:'Calibri' });
  [
    { v:'35%', l:'AOV uplift', s:'Pincode baseline: same cart × intent model' },
    { v:'AED 246', l:'Avg. steal deal saving', s:'Per successful CheckoutLift transaction' },
    { v:'68%', l:'Co-buy acceptance rate', s:'Cart-composition deals vs. generic banners' },
    { v:'22%', l:'Checkout conversion lift', s:'PhonePe offer redesign baseline' },
  ].forEach((m,i) => {
    const y = 2.06+i*1.15;
    s.addShape('rect', { x:0.5, y, w:4.3, h:1.0, fill:{color:C.hero}, line:{type:'none'}, shadow:mk() });
    s.addText(m.v, { x:0.6, y:y+0.07, w:4.1, h:0.52, fontSize:28, bold:true, color:C.bright, fontFace:'Calibri' });
    s.addText(m.l, { x:0.6, y:y+0.57, w:2.4, h:0.3, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(m.s, { x:0.6, y:y+0.57, w:4.1, h:0.3, fontSize:10, color:C.muted, fontFace:'Calibri', align:'right' });
  });
  s.addShape('line', { x:4.95, y:1.56, w:0, h:5.7, line:{color:C.hero, width:1, dashType:'dash'} });
  s.addShape('rect', { x:5.1, y:1.56, w:4.4, h:0.4, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('AMAZON H&L ROI', { x:5.1, y:1.56, w:4.4, h:0.4, fontSize:10, bold:true, color:C.gold, align:'center', valign:'middle', fontFace:'Calibri' });
  [
    { v:'60%', l:'Cart abandonment reduction', s:'Pincode: direct mechanic proof' },
    { v:'+1.3 items', l:'Units per order uplift', s:'From deal attach at checkout entry' },
    { v:'AED 0', l:'Extra discount budget needed', s:'Deal funding stays in existing margin structure' },
    { v:'20%', l:'30-day retention improvement', s:'Habitual steal deal shopper cohort forms' },
  ].forEach((m,i) => {
    const y = 2.06+i*1.15;
    s.addShape('rect', { x:5.1, y, w:4.4, h:1.0, fill:{color:C.hero}, line:{type:'none'}, shadow:mk() });
    s.addText(m.v, { x:5.2, y:y+0.07, w:4.2, h:0.52, fontSize:28, bold:true, color:C.bright, fontFace:'Calibri' });
    s.addText(m.l, { x:5.2, y:y+0.57, w:2.7, h:0.3, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(m.s, { x:5.2, y:y+0.57, w:4.2, h:0.3, fontSize:10, color:C.muted, fontFace:'Calibri', align:'right' });
  });
  s.addShape('rect', { x:0.5, y:6.74, w:9, h:0.6, fill:{color:'1a2332'}, line:{color:C.bright, width:1} });
  s.addText("CheckoutLift doesn't add discount budget — it captures revenue that Amazon was already losing to checkout abandonment and Google search.", {
    x:0.65, y:6.82, w:8.7, h:0.38, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true
  });
}

// SLIDE 7 — PROOF OF WORK
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.lgray} });
  s.addText('PROOF OF WORK', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText("I Built the Exact Same Engine at PhonePe. Here's the Map.", {
    x:0.5, y:0.62, w:9, h:0.7, fontSize:26, bold:true, color:C.ink, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.38, w:9, h:0.04, fill:{color:C.bright}, line:{type:'none'} });
  s.addShape('rect', { x:0.5, y:1.56, w:4.3, h:5.1, fill:{color:C.dark}, line:{type:'none'}, shadow:mk() });
  s.addText('WHAT I BUILT AT PHONEPE', { x:0.65, y:1.7, w:4.0, h:0.35, fontSize:9, bold:true, color:C.bright, fontFace:'Calibri', charSpacing:3 });
  [
    'Built dynamic cart incentivization engine for Pincode (q-commerce) — exact same cart x intent x time trigger model',
    '35% AOV uplift, 60% cart abandonment reduction, 20% 30-day retention improvement across A/B tested variants',
    'Owned full BRD-to-launch: cart signal ingestion, ML model, deal eligibility engine, checkout UX integration',
    'Ran 4 A/B variants simultaneously across trigger type, discount depth, countdown duration, and placement',
    '22% checkout conversion lift from offers placement redesign — same UX muscle needed for CheckoutLift card',
  ].forEach((item,i) => {
    s.addShape('rect', { x:0.65, y:2.2+i*0.74, w:0.2, h:0.2, fill:{color:C.bright}, line:{type:'none'} });
    s.addText(item, { x:0.92, y:2.16+i*0.74, w:3.75, h:0.62, fontSize:10, color:C.muted, fontFace:'Calibri' });
  });
  s.addShape('rect', { x:5.1, y:1.56, w:4.4, h:5.1, fill:{color:C.white}, line:{color:'E5E7EB', width:1}, shadow:mk() });
  s.addText('HOW IT MAPS TO CHECKOUTLIFT', { x:5.25, y:1.7, w:4.1, h:0.35, fontSize:9, bold:true, color:C.ink, fontFace:'Calibri', charSpacing:2 });
  [
    ['PhonePe: Cart x intent engine', '→ CheckoutLift: Cart composition model'],
    ['PhonePe: 60% abandonment drop', '→ Target: 50%+ at H&L checkout'],
    ['PhonePe: Countdown timer A/B', '→ CheckoutLift: 30-60 min steal deal window'],
    ['PhonePe: Eligibility + deal match', '→ CheckoutLift: SKU co-buy probability'],
    ['PhonePe: Ops dashboard for offers', '→ CheckoutLift: Merch team A/B console'],
  ].forEach((m,i) => {
    const y = 2.2+i*0.74;
    s.addShape('rect', { x:5.25, y, w:4.1, h:0.64, fill:{color:C.lgray}, line:{type:'none'} });
    s.addText(m[0], { x:5.35, y:y+0.06, w:3.9, h:0.22, fontSize:10, bold:true, color:C.ink, fontFace:'Calibri' });
    s.addText(m[1], { x:5.35, y:y+0.32, w:3.9, h:0.22, fontSize:10, color:C.green, fontFace:'Calibri' });
  });
  s.addShape('rect', { x:0.5, y:6.88, w:9, h:0.46, fill:{color:'FFFBEB'}, line:{color:C.bright, width:1.5} });
  s.addText('"CheckoutLift is the Pincode cart engine — applied to Amazon H&L large-ticket checkout."  — Ajay Avaghade', {
    x:0.7, y:6.94, w:8.6, h:0.3, fontSize:11, color:C.ink, fontFace:'Calibri', italic:true
  });
}

// SLIDE 8 — ROLLOUT PLAN
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('ROLLOUT PLAN', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('Three Phases — UAE Pilot to MENA Category Rollout', {
    x:0.5, y:0.62, w:9, h:0.7, fontSize:28, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.38, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });
  [
    { label:'PHASE 1', period:'Month 1–2', title:'UAE Pilot (H&L)', color:C.hero, acc:C.gold, items:['CheckoutLift on 3 H&L subcategories: AC, Washing Machines, Coffee', 'A/B test: CheckoutLift card vs. control (no card)', 'Instrument co-buy model: cart composition → SKU match', 'KPIs: AOV uplift, deal accept rate, checkout conversion'] },
    { label:'PHASE 2', period:'Month 3–4', title:'Scale + Optimize', color:C.bright, acc:C.dark, items:['Expand to all H&L checkout sessions (UAE)', 'Launch merch A/B console for category managers', 'Weekly model retraining loop with accept/skip signals', 'Add countdown duration optimization (30/45/60 min)'] },
    { label:'PHASE 3', period:'Month 5–6', title:'MENA Category Rollout', color:C.hero, acc:C.gold, items:['Roll out to KSA, Egypt, Kuwait — all Amazon.ae H&L', 'Extend co-buy model to Electronics and Home Decor', 'Launch steal deal notification engine (WhatsApp, push)', 'Publish CheckoutLift playbook for global marketplace teams'] },
  ].forEach((ph,i) => {
    const x = 0.4+i*3.1;
    s.addShape('rect', { x, y:1.56, w:2.95, h:4.55, fill:{color:ph.color}, line:{type:'none'}, shadow:mk() });
    s.addText(ph.label, { x:x+0.1, y:1.68, w:2.75, h:0.28, fontSize:9, bold:true, color:ph.acc, fontFace:'Calibri', charSpacing:3 });
    s.addText(ph.period, { x:x+0.1, y:1.96, w:2.75, h:0.26, fontSize:9, color:ph.color===C.bright?C.dark:C.muted, fontFace:'Calibri' });
    s.addText(ph.title, { x:x+0.1, y:2.26, w:2.75, h:0.5, fontSize:17, bold:true, color:ph.color===C.bright?C.dark:C.white, fontFace:'Calibri' });
    ph.items.forEach((item,j) => {
      s.addShape('ellipse', { x:x+0.12, y:2.9+j*0.74, w:0.12, h:0.12, fill:{color:ph.acc}, line:{type:'none'} });
      s.addText(item, { x:x+0.3, y:2.84+j*0.74, w:2.55, h:0.64, fontSize:10, color:ph.color===C.bright?C.dark:C.muted, fontFace:'Calibri' });
    });
  });
  s.addShape('rect', { x:0.5, y:6.35, w:9, h:0.95, fill:{color:'1a2332'}, line:{color:C.bright, width:1.5}, shadow:mk() });
  s.addText('What I need to build CheckoutLift:', { x:0.7, y:6.44, w:4, h:0.35, fontSize:12, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('H&L order + cart data access  ·  Co-buy ML model (can adapt PhonePe architecture)  ·  Checkout placement API  ·  1 PM + 2 Engineers + 1 Data Scientist for Phase 1', {
    x:0.7, y:6.77, w:8.6, h:0.38, fontSize:10, color:C.muted, fontFace:'Calibri'
  });
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  ajay-avaghade.github.io/Portfolio', {
    x:0, y:7.2, w:W, h:0.26, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center'
  });
}

await prs.writeFile({ fileName: 'assets/amazon-checkoutlift-deck.pptx' });
console.log('Done: amazon-checkoutlift-deck.pptx');
})();
