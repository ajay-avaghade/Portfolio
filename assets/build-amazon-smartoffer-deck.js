'use strict';
const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:  '131921', hero:  '232F3E', bright:'FF9900',
  gold:  'FEBD69', white: 'FFFFFF', lgray: 'F3F4F6',
  green: '15803D', ink:   '0F172A', muted: '64748B',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });
const W = 10, H = 7.5;

// ── SLIDE 1 — COVER ──────────────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:C.dark} });

  // Diagonal stripe motif
  for(let i=0;i<12;i++){
    s.addShape(prs.ShapeType.line, {
      x: -1 + i*1.1, y:0, w:0.6, h:H,
      line:{ color:C.hero, width:18, transparency:85 }, rotate:15
    });
  }

  // Orange accent block top-right
  s.addShape(prs.ShapeType.rect, { x:7.5, y:0, w:2.5, h:1.2, fill:{color:C.bright}, line:{type:'none'} });
  s.addText('amazon pay', { x:7.55, y:0.1, w:2.3, h:0.45, fontSize:14, bold:true, color:C.dark, fontFace:'Calibri', align:'center' });
  s.addText('PAYMENTS · INDIA', { x:7.55, y:0.55, w:2.3, h:0.3, fontSize:9, bold:true, color:C.dark, fontFace:'Calibri', align:'center', charSpacing:3 });

  // Product name
  s.addText('SmartOffer Engine', { x:0.5, y:1.5, w:7, h:1.1, fontSize:52, bold:true, color:C.white, fontFace:'Calibri' });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:2.7, w:1.2, h:0.08, fill:{color:C.bright}, line:{type:'none'} });
  s.addText('Precision bank discount surfacing powered by payment propensity ML', {
    x:0.5, y:2.85, w:8.5, h:0.5, fontSize:17, color:C.gold, fontFace:'Calibri', italic:true
  });

  // Presenter
  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.8, w:5, h:0.06, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('Presented by Ajay Avaghade  ·  Growth & Payments PM  ·  Amazon Emerging Stores', {
    x:0.5, y:5.95, w:7, h:0.35, fontSize:11, color:C.muted, fontFace:'Calibri'
  });

  // Big stat
  s.addShape(prs.ShapeType.rect, { x:7.5, y:5.0, w:2.2, h:2.0, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('40–60%', { x:7.5, y:5.15, w:2.2, h:0.9, fontSize:34, bold:true, color:C.bright, fontFace:'Calibri', align:'center' });
  s.addText('of IBD budget\nis leakage today', { x:7.5, y:6.0, w:2.2, h:0.8, fontSize:11, color:C.muted, fontFace:'Calibri', align:'center' });
}

// ── SLIDE 2 — THE PROBLEM ────────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('THE PROBLEM', { x:0.5, y:0.3, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText("Amazon's IBD Budget Is a Blunt Instrument — Not a Scalpel", {
    x:0.5, y:0.65, w:9, h:1.0, fontSize:28, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.7, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });

  // 3 stat columns
  const cols3 = [
    { icon:'💸', stat:'40–60%', lbl:'IBD budget is leakage', sub:'Paid to customers who\nwould have bought anyway' },
    { icon:'👁️', stat:'Same banner', lbl:'for every customer', sub:'High-intent buyer and\nfence-sitter see identical offer' },
    { icon:'📉', stat:'1.8× lower', lbl:'ROI vs. targeted spend', sub:'Flat IBD vs. propensity-\ntargeted discount programs' },
  ];
  cols3.forEach((col,i) => {
    const x = 0.5 + i*3.1;
    s.addShape(prs.ShapeType.rect, { x, y:1.85, w:2.9, h:3.4, fill:{color:C.hero}, line:{type:'none'}, shadow:makeShadow() });
    s.addText(col.icon, { x, y:2.0, w:2.9, h:0.6, fontSize:28, align:'center' });
    s.addText(col.stat, { x, y:2.65, w:2.9, h:0.8, fontSize:36, bold:true, color:C.bright, fontFace:'Calibri', align:'center' });
    s.addText(col.lbl, { x, y:3.5, w:2.9, h:0.45, fontSize:12, bold:true, color:C.white, fontFace:'Calibri', align:'center' });
    s.addText(col.sub, { x, y:3.95, w:2.9, h:0.9, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center' });
  });

  // Insight box
  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.5, w:9, h:1.65, fill:{color:'1a2332'}, line:{color:C.bright, width:1.5}, shadow:makeShadow() });
  s.addText('Root Cause:', { x:0.7, y:5.65, w:2, h:0.35, fontSize:11, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('Amazon surfaces the same IBD banner to all checkout users — no differentiation by intent, card eligibility, or price sensitivity. High-intent buyers (who convert without a nudge) claim the discount; undecided buyers (who need the nudge most) get no priority. The result: maximum budget spend, sub-optimal incremental GMV.', {
    x:0.7, y:6.0, w:8.6, h:0.95, fontSize:11, color:C.muted, fontFace:'Calibri'
  });
}

// ── SLIDE 3 — THE INSIGHT ────────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:C.lgray} });
  s.addText('THE INSIGHT', { x:0.5, y:0.3, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('Showing Fewer People the Discount Actually Increases ROI', {
    x:0.5, y:0.65, w:9, h:0.85, fontSize:26, bold:true, color:C.ink, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.55, w:9, h:0.04, fill:{color:C.bright}, line:{type:'none'} });

  // Left column — Current
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.75, w:4.2, h:4.6, fill:{color:C.white}, line:{color:'E5E7EB', width:1}, shadow:makeShadow() });
  s.addText('❌  Flat IBD (Current)', { x:0.65, y:1.95, w:3.9, h:0.45, fontSize:14, bold:true, color:C.ink, fontFace:'Calibri' });
  const leftPts = [
    'Banner shown to ALL checkout users regardless of intent',
    'No card vault eligibility check — offer shown even to unlinked cards',
    'High-intent buyer claims discount: +₹0 incremental GMV, –₹500 cost',
    'Budget exhausted on customers who needed no nudge',
  ];
  leftPts.forEach((t,i) => {
    s.addShape(prs.ShapeType.rect, { x:0.65, y:2.5+i*0.72, w:0.22, h:0.22, fill:{color:'FEE2E2'}, line:{type:'none'} });
    s.addText('✕', { x:0.65, y:2.5+i*0.72, w:0.22, h:0.22, fontSize:10, bold:true, color:'DC2626', align:'center' });
    s.addText(t, { x:0.95, y:2.48+i*0.72, w:3.55, h:0.5, fontSize:11, color:C.muted, fontFace:'Calibri' });
  });

  // VS divider
  s.addShape(prs.ShapeType.ellipse, { x:4.55, y:3.5, w:0.9, h:0.9, fill:{color:C.bright}, line:{type:'none'} });
  s.addText('VS', { x:4.55, y:3.5, w:0.9, h:0.9, fontSize:14, bold:true, color:C.dark, align:'center', valign:'middle', fontFace:'Calibri' });

  // Right column — Proposed
  s.addShape(prs.ShapeType.rect, { x:5.3, y:1.75, w:4.2, h:4.6, fill:{color:C.white}, line:{color:C.bright, width:2}, shadow:makeShadow() });
  s.addText('✅  SmartOffer (Proposed)', { x:5.45, y:1.95, w:3.9, h:0.45, fontSize:14, bold:true, color:C.ink, fontFace:'Calibri' });
  const rightPts = [
    'Propensity score computed at checkout entry — undecided users flagged',
    'Card vault eligibility checked in real-time before surfacing banner',
    'Undecided buyer gets targeted nudge: +₹4,500 GMV, –₹500 cost = 9× ROI',
    'Budget focused where it creates incremental conversion lift',
  ];
  rightPts.forEach((t,i) => {
    s.addShape(prs.ShapeType.rect, { x:5.45, y:2.5+i*0.72, w:0.22, h:0.22, fill:{color:'DCFCE7'}, line:{type:'none'} });
    s.addText('✓', { x:5.45, y:2.5+i*0.72, w:0.22, h:0.22, fontSize:10, bold:true, color:C.green, align:'center' });
    s.addText(t, { x:5.75, y:2.48+i*0.72, w:3.55, h:0.5, fontSize:11, color:C.muted, fontFace:'Calibri' });
  });

  // Callout box
  s.addShape(prs.ShapeType.rect, { x:0.5, y:6.55, w:9, h:0.75, fill:{color:'FFFBEB'}, line:{color:C.bright, width:1.5}, shadow:makeShadow() });
  s.addText('Key Insight: ', { x:0.7, y:6.68, w:1.3, h:0.4, fontSize:11, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('The discount only pays for itself when it creates incremental conversion. By suppressing it for high-intent buyers, every rupee of IBD budget generates measurably more GMV.', {
    x:1.9, y:6.68, w:7.4, h:0.45, fontSize:11, color:C.ink, fontFace:'Calibri'
  });
}

// ── SLIDE 4 — THE MECHANIC ───────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('THE MECHANIC', { x:0.5, y:0.3, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('5-Step Precision Surfacing Flow', {
    x:0.5, y:0.65, w:9, h:0.7, fontSize:30, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.4, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });

  const steps = [
    { n:'01', title:'Checkout Entry', body:'User enters checkout. SmartOffer engine triggers in <50ms.' },
    { n:'02', title:'Card Vault Check', body:'Real-time eligibility check: is a qualifying bank card linked to this account?' },
    { n:'03', title:'Propensity Score', body:'ML model scores P(convert | no discount). High-intent users suppressed; undecided users flagged.' },
    { n:'04', title:'Contextual Surface', body:'IBD banner rendered only for undecided + eligible users. Personalised: bank name, exact savings, order-specific.' },
    { n:'05', title:'Conversion & Attribution', body:'Transaction captured; discount applied. Cohort performance fed back to model for weekly retraining.' },
  ];
  steps.forEach((step,i) => {
    const x = 0.4 + i*1.88;
    s.addShape(prs.ShapeType.ellipse, { x:x+0.54, y:1.65, w:0.7, h:0.7, fill:{color:C.bright}, line:{type:'none'}, shadow:makeShadow() });
    s.addText(step.n, { x:x+0.54, y:1.65, w:0.7, h:0.7, fontSize:13, bold:true, color:C.dark, align:'center', valign:'middle', fontFace:'Calibri' });
    if(i<4){
      s.addShape(prs.ShapeType.line, { x:x+1.24, y:2.0, w:0.64, h:0, line:{color:C.bright, width:1.5, dashType:'dash'} });
    }
    s.addShape(prs.ShapeType.rect, { x, y:2.55, w:1.78, h:3.5, fill:{color:C.hero}, line:{type:'none'}, shadow:makeShadow() });
    s.addText(step.title, { x:x+0.05, y:2.7, w:1.68, h:0.45, fontSize:12, bold:true, color:C.white, fontFace:'Calibri', align:'center' });
    s.addText(step.body, { x:x+0.05, y:3.2, w:1.68, h:2.5, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center' });
  });

  // A/B proof note
  s.addShape(prs.ShapeType.rect, { x:0.5, y:6.35, w:9, h:0.9, fill:{color:'1a2332'}, line:{color:C.bright, width:1}, shadow:makeShadow() });
  s.addText('PhonePe Proof  ·  ', { x:0.7, y:6.5, w:2.1, h:0.5, fontSize:11, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('Exact same mechanic shipped at PhonePe: Propensity-to-Transact ML replaced manual cohort targeting across 350M MAU → 32% reduction in marketing burn, sustained GMV. SmartOffer applies identical logic to Amazon\'s IBD budget.', {
    x:2.7, y:6.5, w:6.6, h:0.6, fontSize:10, color:C.muted, fontFace:'Calibri'
  });
}

// ── SLIDE 5 — THE PRODUCT ────────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:C.lgray} });
  s.addText('THE PRODUCT', { x:0.5, y:0.3, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('4 Screen States — Prototype Built & Live', {
    x:0.5, y:0.65, w:9, h:0.7, fontSize:28, bold:true, color:C.ink, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.4, w:9, h:0.04, fill:{color:C.bright}, line:{type:'none'} });

  const screens = [
    { n:'01', title:'Smart Checkout', desc:'IBD banner surfaces for eligible undecided user only. High-intent buyer sees clean checkout — no banner, no discount leakage.', mock:['amazon checkout', '─────────────', '🎧 Sony WH-1000XM5', '₹24,990 → ₹22,491', '─────────────', '✦ PERSONALISED OFFER', 'HDFC · Save ₹2,499', '─────────────', '[Apply & Pay ₹22,491]'] },
    { n:'02', title:'Eligibility Detail', desc:'Card vault check + exact savings shown. Eligibility criteria transparent. Builds trust before asking for payment method switch.', mock:['← Back to checkout', '─────────────', 'HDFC REGALIA CARD', '**** **** **** 4521', '─────────────', '₹2,499 Savings', '✓ Card linked', '✓ Offer valid today', '[Use HDFC Card]'] },
    { n:'03', title:'Discount Applied', desc:'Savings confirmed with receipt. Offer reset date surfaced to set expectation for next purchase. Cross-sell nudge anchors next visit.', mock:['Order Placed! ✓', '─────────────', '💰 ₹2,499 Saved', '─────────────', 'HDFC Discount –₹2,499', 'Amount paid ₹22,491', '─────────────', '🎯 Offer resets June 1', '[Track Order]'] },
    { n:'04', title:'Ops Dashboard', desc:'Bank partner campaign manager: propensity cohort split, budget efficiency vs. flat targeting, per-bank conversion rates. Closes the loop.', mock:['SmartOffer Engine ●', '─────────────', 'Targeted: 12.4K', 'Conversions: 3,847', 'Budget saved: ₹18.2L', 'IBD ROI: 3.2×', '─────────────', 'HDFC conv: 31.2% ↑'] },
  ];
  screens.forEach((sc,i) => {
    const x = 0.35 + i*2.35;
    s.addShape(prs.ShapeType.rect, { x, y:1.6, w:2.2, h:5.65, fill:{color:C.white}, line:{color:'E5E7EB', width:1}, shadow:makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y:1.6, w:2.2, h:0.45, fill:{color:C.dark}, line:{type:'none'} });
    s.addText(`${sc.n} — ${sc.title}`, { x:x+0.07, y:1.65, w:2.06, h:0.35, fontSize:9, bold:true, color:C.gold, fontFace:'Calibri' });
    s.addText(sc.desc, { x:x+0.07, y:2.1, w:2.06, h:1.0, fontSize:9, color:C.muted, fontFace:'Calibri' });
    s.addShape(prs.ShapeType.rect, { x:x+0.07, y:3.15, w:2.06, h:3.85, fill:{color:'F8FAFC'}, line:{color:'E5E7EB', width:1} });
    s.addText(sc.mock.join('\n'), { x:x+0.12, y:3.2, w:1.96, h:3.75, fontSize:8, color:C.ink, fontFace:'DM Mono', valign:'top' });
  });
}

// ── SLIDE 6 — IMPACT & ROI ───────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('IMPACT & ROI', { x:0.5, y:0.3, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('Projected Impact — Built on PhonePe Proof Points', {
    x:0.5, y:0.65, w:9, h:0.7, fontSize:26, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.4, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });

  // Left block — Customer Impact
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.6, w:4.3, h:0.4, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('CUSTOMER IMPACT', { x:0.5, y:1.6, w:4.3, h:0.4, fontSize:10, bold:true, color:C.gold, align:'center', valign:'middle', fontFace:'Calibri' });
  const custMetrics = [
    { val:'18–22%', lbl:'Conversion lift', sub:'Among undecided cohort (vs. flat IBD)' },
    { val:'↑ Relevance', lbl:'Offer trust score', sub:'Personalised offer = higher perceived value' },
    { val:'2.4×', lbl:'Banner CTR', sub:'SmartOffer vs. flat targeting baseline' },
    { val:'₹2,499', lbl:'Avg. savings per order', sub:'Eligible customer receives correct offer' },
  ];
  custMetrics.forEach((m,i) => {
    const y = 2.1 + i*1.15;
    s.addShape(prs.ShapeType.rect, { x:0.5, y, w:4.3, h:1.0, fill:{color:C.hero}, line:{type:'none'}, shadow:makeShadow() });
    s.addText(m.val, { x:0.6, y:y+0.08, w:4.1, h:0.5, fontSize:28, bold:true, color:C.bright, fontFace:'Calibri' });
    s.addText(m.lbl, { x:0.6, y:y+0.56, w:2.5, h:0.3, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(m.sub, { x:0.6, y:y+0.56, w:4.1, h:0.3, fontSize:10, color:C.muted, fontFace:'Calibri', align:'right' });
  });

  // Dashed divider
  s.addShape(prs.ShapeType.line, { x:4.95, y:1.6, w:0, h:5.65, line:{color:C.hero, width:1, dashType:'dash'} });

  // Right block — Amazon ROI
  s.addShape(prs.ShapeType.rect, { x:5.1, y:1.6, w:4.4, h:0.4, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('AMAZON IBD ROI', { x:5.1, y:1.6, w:4.4, h:0.4, fontSize:10, bold:true, color:C.gold, align:'center', valign:'middle', fontFace:'Calibri' });
  const roiMetrics = [
    { val:'30–35%', lbl:'IBD budget leakage reduction', sub:'Discount suppressed for high-intent users' },
    { val:'3.2×', lbl:'IBD ROI improvement', sub:'GMV per discount rupee spent' },
    { val:'₹18.2L/day', lbl:'Budget saved (est.)', sub:'Redirected to new user acquisition' },
    { val:'32%', lbl:'Marketing burn reduction', sub:'PhonePe baseline: exact same model' },
  ];
  roiMetrics.forEach((m,i) => {
    const y = 2.1 + i*1.15;
    s.addShape(prs.ShapeType.rect, { x:5.1, y, w:4.4, h:1.0, fill:{color:C.hero}, line:{type:'none'}, shadow:makeShadow() });
    s.addText(m.val, { x:5.2, y:y+0.08, w:4.2, h:0.5, fontSize:28, bold:true, color:C.bright, fontFace:'Calibri' });
    s.addText(m.lbl, { x:5.2, y:y+0.56, w:2.8, h:0.3, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(m.sub, { x:5.2, y:y+0.56, w:4.2, h:0.3, fontSize:10, color:C.muted, fontFace:'Calibri', align:'right' });
  });

  // Insight
  s.addShape(prs.ShapeType.rect, { x:0.5, y:6.75, w:9, h:0.6, fill:{color:'1a2332'}, line:{color:C.bright, width:1} });
  s.addText('SmartOffer only redirects the discount budget — it doesn\'t increase spend. Every rupee saved on high-intent suppression funds incremental growth elsewhere.', {
    x:0.65, y:6.82, w:8.7, h:0.4, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true
  });
}

// ── SLIDE 7 — PROOF OF WORK ──────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:C.lgray} });
  s.addText('PROOF OF WORK', { x:0.5, y:0.3, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText("I've Shipped This. Here's the Evidence.", {
    x:0.5, y:0.65, w:9, h:0.7, fontSize:28, bold:true, color:C.ink, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.4, w:9, h:0.04, fill:{color:C.bright}, line:{type:'none'} });

  // Left — PhonePe (dark)
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.6, w:4.3, h:5.0, fill:{color:C.dark}, line:{type:'none'}, shadow:makeShadow() });
  s.addText('WHAT I BUILT AT PHONEPE', { x:0.65, y:1.75, w:4.0, h:0.35, fontSize:9, bold:true, color:C.bright, fontFace:'Calibri', charSpacing:3 });
  const phonepeItems = [
    'Redesigned payment eligibility surfacing across 350M MAU — same UX problem as SmartOffer checkout placement',
    'Shipped Propensity-to-Transact ML models replacing ₹1,000+ Cr/yr manual cohort targeting → 32% marketing burn reduction',
    'Owned the full offer redemption stack: eligibility check, banner surfacing, checkout integration, and attribution',
    'A/B tested targeting strategies across 8+ payment cohorts — developed the undecided vs. high-intent segmentation framework',
    '22% checkout conversion lift — directly comparable to SmartOffer projected 18-22% undecided cohort improvement',
  ];
  phonepeItems.forEach((item,i) => {
    s.addShape(prs.ShapeType.rect, { x:0.65, y:2.22+i*0.72, w:0.2, h:0.2, fill:{color:C.bright}, line:{type:'none'} });
    s.addText(item, { x:0.93, y:2.18+i*0.72, w:3.75, h:0.6, fontSize:10, color:C.muted, fontFace:'Calibri' });
  });

  // Right — JD Mapping (light)
  s.addShape(prs.ShapeType.rect, { x:5.1, y:1.6, w:4.4, h:5.0, fill:{color:C.white}, line:{color:'E5E7EB', width:1}, shadow:makeShadow() });
  s.addText('HOW IT MAPS TO THIS ROLE', { x:5.25, y:1.75, w:4.1, h:0.35, fontSize:9, bold:true, color:C.ink, fontFace:'Calibri', charSpacing:3 });
  const mappings = [
    ['JD: End-to-end IBD lifecycle', '→ Owned offers stack E2E at PhonePe'],
    ['JD: Data-driven customer behavior', '→ Built ML propensity models at scale'],
    ['JD: Launch in new markets', '→ Rolled out 0→1 offers in Pincode (Q-commerce)'],
    ['JD: Partner with engineering', '→ Owned PRD-to-delivery across 6+ squads'],
    ['JD: Present to senior leadership', '→ Pitched propensity model ROI to CxO'],
  ];
  mappings.forEach((m,i) => {
    const y = 2.22 + i*0.72;
    s.addShape(prs.ShapeType.rect, { x:5.25, y, w:4.1, h:0.62, fill:{color:C.lgray}, line:{type:'none'} });
    s.addText(m[0], { x:5.35, y:y+0.06, w:3.9, h:0.22, fontSize:10, bold:true, color:C.ink, fontFace:'Calibri' });
    s.addText(m[1], { x:5.35, y:y+0.3, w:3.9, h:0.22, fontSize:10, color:C.green, fontFace:'Calibri' });
  });

  // Quote box
  s.addShape(prs.ShapeType.rect, { x:0.5, y:6.85, w:9, h:0.5, fill:{color:'FFFBEB'}, line:{color:C.bright, width:1.5} });
  s.addText('"SmartOffer isn\'t a concept — it\'s what I shipped at PhonePe, now applied to Amazon\'s IBD problem."  — Ajay Avaghade', {
    x:0.7, y:6.92, w:8.6, h:0.33, fontSize:11, color:C.ink, fontFace:'Calibri', italic:true
  });
}

// ── SLIDE 8 — ROLLOUT PLAN ───────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('ROLLOUT PLAN', { x:0.5, y:0.3, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('3-Phase Deployment Across India → MENA → APAC', {
    x:0.5, y:0.65, w:9, h:0.7, fontSize:28, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.4, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });

  const phases = [
    { label:'PHASE 1', period:'Month 1–2', title:'India Pilot', color:C.hero, accent:C.gold, items:['Deploy SmartOffer to 50K checkout sessions/day in India', 'A/B test: SmartOffer suppression vs. flat IBD baseline', 'Instrument propensity model: card vault check + intent signals', 'Measure: leakage reduction, undecided cohort conversion'] },
    { label:'PHASE 2', period:'Month 3–4', title:'Scale + Retrain', color:C.bright, accent:C.dark, items:['Scale to full India traffic (~2M checkout sessions/day)', 'Weekly model retraining loop with conversion feedback', 'Onboard HDFC, Axis, ICICI, Kotak to SmartOffer campaign manager', 'Launch ops dashboard for bank partner self-serve reporting'] },
    { label:'PHASE 3', period:'Month 5–6', title:'MENA/APAC Expansion', color:C.hero, accent:C.gold, items:['Adapt model for MENA card networks (Mashreq, FAB, Riyad)', 'Launch SmartOffer in UAE, Saudi Arabia, Egypt', 'Extend to APAC: SG, MY with local bank partnership APIs', 'Publish IBD ROI playbook for global rollout template'] },
  ];
  phases.forEach((ph,i) => {
    const x = 0.4 + i*3.1;
    s.addShape(prs.ShapeType.rect, { x, y:1.6, w:2.95, h:4.5, fill:{color:ph.color}, line:{type:'none'}, shadow:makeShadow() });
    s.addText(ph.label, { x:x+0.1, y:1.72, w:2.75, h:0.28, fontSize:9, bold:true, color:ph.accent, fontFace:'Calibri', charSpacing:3 });
    s.addText(ph.period, { x:x+0.1, y:1.98, w:2.75, h:0.28, fontSize:9, color: ph.color===C.bright ? C.dark : C.muted, fontFace:'Calibri' });
    s.addText(ph.title, { x:x+0.1, y:2.3, w:2.75, h:0.5, fontSize:17, bold:true, color: ph.color===C.bright ? C.dark : C.white, fontFace:'Calibri' });
    ph.items.forEach((item,j) => {
      s.addShape(prs.ShapeType.ellipse, { x:x+0.12, y:2.95+j*0.75, w:0.12, h:0.12, fill:{color:ph.accent}, line:{type:'none'} });
      s.addText(item, { x:x+0.3, y:2.88+j*0.75, w:2.55, h:0.65, fontSize:10, color: ph.color===C.bright ? C.dark : C.muted, fontFace:'Calibri' });
    });
  });

  // Ask box
  s.addShape(prs.ShapeType.rect, { x:0.5, y:6.35, w:9, h:1.0, fill:{color:'1a2332'}, line:{color:C.bright, width:1.5}, shadow:makeShadow() });
  s.addText('What I need to build SmartOffer:', { x:0.7, y:6.45, w:4, h:0.35, fontSize:12, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('Access to checkout event logs  ·  ML infra for propensity scoring  ·  Bank partner API integration  ·  1 PM + 2 Engineers for Phase 1', {
    x:0.7, y:6.78, w:8.6, h:0.4, fontSize:10, color:C.muted, fontFace:'Calibri'
  });
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439', {
    x:0, y:7.2, w:W, h:0.28, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center'
  });
}

prs.writeFile({ fileName: 'assets/amazon-smartoffer-deck.pptx' })
  .then(() => console.log('✅ Deck written: amazon-smartoffer-deck.pptx'))
  .catch(e => console.error('❌', e));
