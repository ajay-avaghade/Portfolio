'use strict';
(async () => {
const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:'131921', hero:'232F3E', bright:'FF9900',
  green:'067D62', dkgreen:'044d3c', white:'FFFFFF',
  lgray:'F3F4F6', gold:'FEBD69', ink:'0F172A', muted:'64748B',
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
  s.addText('WarrantyIQ', { x:0.5, y:1.4, w:7, h:1.1, fontSize:52, bold:true, color:C.white, fontFace:'Calibri' });
  s.addShape('rect', { x:0.5, y:2.6, w:1.4, h:0.08, fill:{color:C.green}, line:{type:'none'} });
  s.addText('Contextual warranty attach engine — failure-probability intelligence at the PDP', {
    x:0.5, y:2.75, w:8.5, h:0.5, fontSize:17, color:C.gold, fontFace:'Calibri', italic:true
  });
  s.addText('Case Study 02 of 03  ·  Amazon MENA Senior PM, Home & Leisure', {
    x:0.5, y:3.35, w:7, h:0.3, fontSize:10, color:C.muted, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:5.75, w:5, h:0.06, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('Presented by Ajay Avaghade  ·  Growth & Category PM  ·  PhonePe → Amazon MENA', {
    x:0.5, y:5.9, w:7, h:0.35, fontSize:11, color:C.muted, fontFace:'Calibri'
  });
  s.addShape('rect', { x:7.6, y:5.1, w:2.1, h:2.0, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('3%', { x:7.6, y:5.2, w:2.1, h:0.8, fontSize:40, bold:true, color:C.bright, fontFace:'Calibri', align:'center' });
  s.addText('warranty\nattach rate\ntoday', { x:7.6, y:5.95, w:2.1, h:0.9, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center' });
}

// SLIDE 2 — THE PROBLEM
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('THE PROBLEM', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText("Amazon Sells AED 1,899 Washing Machines and Then Walks Away", {
    x:0.5, y:0.62, w:9, h:0.9, fontSize:25, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.6, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });
  [
    { icon:'🛡️', stat:'3%', lbl:'Warranty attach rate', sub:'On washing machines, ACs\nand refrigerators — industry\nnorm is 18–25% with active surfacing' },
    { icon:'🔧', stat:'AED 900', lbl:'Avg motor repair cost', sub:'When the machine fails\nat year 2–3 — the exact\nwindow manufacturer warranty expires' },
    { icon:'😤', stat:'Amazon', lbl:'Gets the blame', sub:'Customer bought from Amazon\n— Amazon gets the support\ncall, even though warranty expired' },
  ].forEach((col,i) => {
    const x = 0.5+i*3.1;
    s.addShape('rect', { x, y:1.78, w:2.9, h:3.5, fill:{color:C.hero}, line:{type:'none'}, shadow:mk() });
    s.addText(col.icon, { x, y:1.92, w:2.9, h:0.6, fontSize:28, align:'center' });
    s.addText(col.stat, { x, y:2.58, w:2.9, h:0.8, fontSize:38, bold:true, color:C.bright, fontFace:'Calibri', align:'center' });
    s.addText(col.lbl, { x:x+0.1, y:3.42, w:2.7, h:0.4, fontSize:12, bold:true, color:C.white, fontFace:'Calibri', align:'center' });
    s.addText(col.sub, { x:x+0.1, y:3.85, w:2.7, h:1.0, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center' });
  });
  s.addShape('rect', { x:0.5, y:5.48, w:9, h:1.8, fill:{color:'1a2332'}, line:{color:C.bright, width:1.5}, shadow:mk() });
  s.addText('Root Cause:', { x:0.7, y:5.62, w:2, h:0.35, fontSize:11, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('Manufacturer warranty is treated as a checkbox on the PDP, not a product feature. When it expires silently at year 1, Amazon has no mechanism to extend protection, no data on which products are most failure-prone, and no revenue from the AED 450–900 repair that follows. WarrantyIQ fixes all three.', {
    x:0.7, y:5.98, w:8.6, h:1.1, fontSize:11, color:C.muted, fontFace:'Calibri'
  });
}

// SLIDE 3 — THE INSIGHT
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.lgray} });
  s.addText('THE INSIGHT', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText("The Warranty Pitch Fails Because It Has No Intelligence — Just a Banner", {
    x:0.5, y:0.62, w:9, h:0.85, fontSize:24, bold:true, color:C.ink, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.52, w:9, h:0.04, fill:{color:C.green}, line:{type:'none'} });
  s.addShape('rect', { x:0.5, y:1.72, w:4.2, h:5.2, fill:{color:C.white}, line:{color:'E5E7EB', width:1}, shadow:mk() });
  s.addText('❌  Warranty Today', { x:0.65, y:1.9, w:3.9, h:0.42, fontSize:14, bold:true, color:C.ink, fontFace:'Calibri' });
  [
    '"Add 1-year warranty — AED 149" shown below the buy button',
    'No context: why this product? why this price? why now?',
    'No failure data surfaced: "Washing machine motors fail at year 2-3"',
    'Customer skips it. Machine fails 18 months later. NPS tanks.',
  ].forEach((t,i) => {
    s.addShape('rect', { x:0.65, y:2.48+i*0.76, w:0.22, h:0.22, fill:{color:'FEE2E2'}, line:{type:'none'} });
    s.addText('✕', { x:0.65, y:2.48+i*0.76, w:0.22, h:0.22, fontSize:10, bold:true, color:'DC2626', align:'center' });
    s.addText(t, { x:0.95, y:2.44+i*0.76, w:3.55, h:0.6, fontSize:11, color:C.muted, fontFace:'Calibri' });
  });
  s.addShape('ellipse', { x:4.55, y:3.6, w:0.9, h:0.9, fill:{color:C.bright}, line:{type:'none'} });
  s.addText('VS', { x:4.55, y:3.6, w:0.9, h:0.9, fontSize:14, bold:true, color:C.dark, align:'center', valign:'middle', fontFace:'Calibri' });
  s.addShape('rect', { x:5.3, y:1.72, w:4.2, h:5.2, fill:{color:C.white}, line:{color:C.green, width:2}, shadow:mk() });
  s.addText('✅  WarrantyIQ', { x:5.45, y:1.9, w:3.9, h:0.42, fontSize:14, bold:true, color:C.ink, fontFace:'Calibri' });
  [
    'Failure-probability model: "Washing machine motors typically fail at year 2–3"',
    'WHY THIS: "Repair cost without cover: AED 450–900. 2yr plan: AED 239."',
    'Three plan tiers with co-pay AED 0 and in-home technician promise',
    'Warranty activates on delivery. Dashboard + claim filing via WhatsApp.',
  ].forEach((t,i) => {
    s.addShape('rect', { x:5.45, y:2.48+i*0.76, w:0.22, h:0.22, fill:{color:'DCFCE7'}, line:{type:'none'} });
    s.addText('✓', { x:5.45, y:2.48+i*0.76, w:0.22, h:0.22, fontSize:10, bold:true, color:C.green, align:'center' });
    s.addText(t, { x:5.75, y:2.44+i*0.76, w:3.55, h:0.6, fontSize:11, color:C.muted, fontFace:'Calibri' });
  });
  s.addShape('rect', { x:0.5, y:7.05, w:9, h:0.3, fill:{color:'F0FFF8'}, line:{color:C.green, width:1.5} });
  s.addText('Key Insight: Context converts. "Motor fails at year 2-3. Repair: AED 900. Plan: AED 239." is a financial argument — not a feature upsell.', {
    x:0.65, y:7.07, w:8.7, h:0.24, fontSize:9.5, color:C.ink, fontFace:'Calibri', italic:true
  });
}

// SLIDE 4 — THE MECHANIC
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('THE MECHANIC', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('Failure Intelligence × Plan Tiers × Zero-Deductible Promise = WarrantyIQ', {
    x:0.5, y:0.62, w:9, h:0.7, fontSize:23, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.38, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });
  [
    { n:'01', title:'PDP Signal', body:'Product category + brand triggers WarrantyIQ. Category failure-rate model loaded (<50ms). Card rendered below Add to Cart.' },
    { n:'02', title:'Failure Context', body:'"Motor fails at year 2–3. Without cover: AED 450–900 repair." Converts the pitch from upsell to financial advice.' },
    { n:'03', title:'Three Plan Tiers', body:'1yr / 2yr (POPULAR) / 3yr. Price, monthly equiv, and key features per tier. 2yr pre-selected — best value signal.' },
    { n:'04', title:'Cart + Checkout', body:'Plan added as line item. Cart shows warranty alongside product. Checkout includes plan cost in total.' },
    { n:'05', title:'Auto-Activate + Dashboard', body:'On delivery, warranty activates automatically. Dashboard shows coverage period, claim filing, authorised service centres.' },
  ].forEach((step,i) => {
    const x = 0.38+i*1.88;
    s.addShape('ellipse', { x:x+0.52, y:1.62, w:0.72, h:0.72, fill:{color:C.green}, line:{type:'none'}, shadow:mk() });
    s.addText(step.n, { x:x+0.52, y:1.62, w:0.72, h:0.72, fontSize:13, bold:true, color:C.white, align:'center', valign:'middle', fontFace:'Calibri' });
    if(i<4) s.addShape('line', { x:x+1.24, y:1.98, w:0.64, h:0, line:{color:C.green, width:1.5, dashType:'dash'} });
    s.addShape('rect', { x, y:2.52, w:1.78, h:3.55, fill:{color:C.hero}, line:{type:'none'}, shadow:mk() });
    s.addText(step.title, { x:x+0.05, y:2.66, w:1.68, h:0.45, fontSize:12, bold:true, color:C.white, fontFace:'Calibri', align:'center' });
    s.addText(step.body, { x:x+0.05, y:3.16, w:1.68, h:2.55, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center' });
  });
  s.addShape('rect', { x:0.5, y:6.32, w:9, h:0.95, fill:{color:'1a2332'}, line:{color:C.green, width:1}, shadow:mk() });
  s.addText('Design Principle  ·  ', { x:0.7, y:6.46, w:2.2, h:0.4, fontSize:11, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('Every element of the WarrantyIQ card is designed to answer the one question that kills warranty attach: "Why do I need this?" Failure data, repair cost comparison, and plan math answer it before the customer can ask.', {
    x:2.8, y:6.44, w:6.5, h:0.62, fontSize:10, color:C.muted, fontFace:'Calibri'
  });
}

// SLIDE 5 — THE PRODUCT
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.lgray} });
  s.addText('THE PRODUCT', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('5 Screens — PDP to Warranty Dashboard', {
    x:0.5, y:0.62, w:9, h:0.7, fontSize:28, bold:true, color:C.ink, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.38, w:9, h:0.04, fill:{color:C.green}, line:{type:'none'} });
  [
    { n:'01', t:'PDP + Warranty Card', d:'WarrantyIQ card below Add-to-Cart. Failure context + 3 plan tiers inline. "POPULAR" badge on 2yr.', m:['Samsung 8kg Washer', 'AED 1,899 Save 27%', '────────────', 'WARRANTYIQ', '1yr AED149|2yr 239', '[POPULAR] 3yr 299', 'Motor fails yr 2-3', '[See All Plans]'] },
    { n:'02', t:'Plan Comparison', d:'Full plan comparison. Coverage grid: manufacturer vs WarrantyIQ. Radio selector for tier. In-home technician callout.', m:['WarrantyIQ Plans', '────────────', '✕ Mfr: no accident', '✓ WIQ: all failures', '✕ Mfr: year 1 only', '✓ WIQ: 3yr cover', '────────────', '2-Year ◉ AED 239'] },
    { n:'03', t:'Cart + Warranty Line', d:'Warranty added as separate line item with green "ACTIVE" badge. AED 0 deductible callout. Cart total updated.', m:['Shopping Cart (2)', '────────────', 'Samsung 8kg Washer', 'AED 1,899', '────────────', 'WarrantyIQ 2yr [ACTIVE]', 'AED 239 | AED 0 deduct', 'Total: AED 2,138'] },
    { n:'04', t:'Order + Warranty Live', d:'Order confirmed. Warranty activated banner. Coverage period, deductible, and claim count displayed immediately.', m:['Order Placed! ✓', '────────────', 'WARRANTYIQ ACTIVE', 'Samsung 8kg Washer', '2yr | AED 0 deduct', 'Expires: June 2028', '0 of unlimited claims', '[View Dashboard]'] },
    { n:'05', t:'Warranty Dashboard', d:'Coverage progress bar. Quick actions: file claim, book service, view policy. Authorised service centres nearby.', m:['My WarrantyIQ (1)', '────────────', 'Samsung Washer 🟢', 'Protected | 24 months', 'Expires Jun 2028', '────────────', 'File Claim | Book Svc', 'Samsung Svc — 2.4km'] },
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
  s.addText('High-Attach, High-Margin — Zero Net Customer Cost', {
    x:0.5, y:0.62, w:9, h:0.7, fontSize:26, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.38, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });
  s.addShape('rect', { x:0.5, y:1.56, w:4.3, h:0.4, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('SHOPPER VALUE', { x:0.5, y:1.56, w:4.3, h:0.4, fontSize:10, bold:true, color:C.gold, align:'center', valign:'middle', fontFace:'Calibri' });
  [
    { v:'AED 0', l:'Deductible on all claims', s:'No surprise costs when machine fails' },
    { v:'AED 239', l:'2yr plan cost', s:'vs AED 450-900 repair without cover' },
    { v:'24hr', l:'In-home technician SLA', s:'No logistics burden on the customer' },
    { v:'40%', l:'Post-purchase NPS uplift', s:'Protected customers rate higher across categories' },
  ].forEach((m,i) => {
    const y = 2.06+i*1.15;
    s.addShape('rect', { x:0.5, y, w:4.3, h:1.0, fill:{color:C.hero}, line:{type:'none'}, shadow:mk() });
    s.addText(m.v, { x:0.6, y:y+0.07, w:4.1, h:0.52, fontSize:28, bold:true, color:C.green, fontFace:'Calibri' });
    s.addText(m.l, { x:0.6, y:y+0.57, w:2.4, h:0.3, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(m.s, { x:0.6, y:y+0.57, w:4.1, h:0.3, fontSize:10, color:C.muted, fontFace:'Calibri', align:'right' });
  });
  s.addShape('line', { x:4.95, y:1.56, w:0, h:5.7, line:{color:C.hero, width:1, dashType:'dash'} });
  s.addShape('rect', { x:5.1, y:1.56, w:4.4, h:0.4, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('AMAZON H&L ROI', { x:5.1, y:1.56, w:4.4, h:0.4, fontSize:10, bold:true, color:C.gold, align:'center', valign:'middle', fontFace:'Calibri' });
  [
    { v:'28%', l:'Warranty attach rate (target)', s:'vs 3% today — industry norm with active surfacing' },
    { v:'AED 239', l:'Avg plan revenue per transaction', s:'High-margin revenue attached to H&L sales' },
    { v:'~15%', l:'Gross margin on warranty plans', s:'Net of service centre payout — structural profit' },
    { v:'32%', l:'Returns reduction (protected items)', s:'Customers claim repair instead of returning' },
  ].forEach((m,i) => {
    const y = 2.06+i*1.15;
    s.addShape('rect', { x:5.1, y, w:4.4, h:1.0, fill:{color:C.hero}, line:{type:'none'}, shadow:mk() });
    s.addText(m.v, { x:5.2, y:y+0.07, w:4.2, h:0.52, fontSize:28, bold:true, color:C.bright, fontFace:'Calibri' });
    s.addText(m.l, { x:5.2, y:y+0.57, w:2.7, h:0.3, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(m.s, { x:5.2, y:y+0.57, w:4.2, h:0.3, fontSize:10, color:C.muted, fontFace:'Calibri', align:'right' });
  });
  s.addShape('rect', { x:0.5, y:6.74, w:9, h:0.6, fill:{color:'1a2332'}, line:{color:C.green, width:1} });
  s.addText('WarrantyIQ converts a customer liability (post-purchase support calls) into a high-margin revenue stream — while improving NPS for the same transaction.', {
    x:0.65, y:6.82, w:8.7, h:0.38, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true
  });
}

// SLIDE 7 — PROOF OF WORK
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.lgray} });
  s.addText('PROOF OF WORK', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText("Marketplace Monetization at Scale — The Architecture Transfers Directly.", {
    x:0.5, y:0.62, w:9, h:0.7, fontSize:25, bold:true, color:C.ink, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.38, w:9, h:0.04, fill:{color:C.green}, line:{type:'none'} });
  s.addShape('rect', { x:0.5, y:1.56, w:4.3, h:5.1, fill:{color:C.dark}, line:{type:'none'}, shadow:mk() });
  s.addText('WHAT I BUILT AT PHONEPE', { x:0.65, y:1.7, w:4.0, h:0.35, fontSize:9, bold:true, color:C.bright, fontFace:'Calibri', charSpacing:3 });
  [
    'Rebuilt ₹100 Cr/yr rewards portfolio into ML-driven marketplace — 500+ brand partners, 11% YoY revenue growth',
    'Designed self-serve monetization layer: brands set audience, discount depth, and timing — no ops team required',
    'Owned full BRD-to-launch for each new merchant vertical: legal, finance, ops, and product sign-off in parallel',
    '500+ brand partner onboarding without adding headcount — scalable eligibility and payout engine',
    'Built offer redemption stack: eligibility, surfacing, checkout integration, attribution — end-to-end ownership',
  ].forEach((item,i) => {
    s.addShape('rect', { x:0.65, y:2.2+i*0.74, w:0.2, h:0.2, fill:{color:C.bright}, line:{type:'none'} });
    s.addText(item, { x:0.92, y:2.16+i*0.74, w:3.75, h:0.62, fontSize:10, color:C.muted, fontFace:'Calibri' });
  });
  s.addShape('rect', { x:5.1, y:1.56, w:4.4, h:5.1, fill:{color:C.white}, line:{color:'E5E7EB', width:1}, shadow:mk() });
  s.addText('HOW IT MAPS TO WARRANTYIQ', { x:5.25, y:1.7, w:4.1, h:0.35, fontSize:9, bold:true, color:C.ink, fontFace:'Calibri', charSpacing:2 });
  [
    ['PhonePe: ML-driven brand marketplace', '→ WarrantyIQ: failure-probability model per SKU'],
    ['PhonePe: Self-serve monetization layer', '→ WarrantyIQ: service centre partner onboarding'],
    ['PhonePe: Eligibility + payout engine', '→ WarrantyIQ: plan eligibility + claim payout'],
    ['PhonePe: BRD to launch, end-to-end', '→ WarrantyIQ: full lifecycle from spec to ops'],
    ['PhonePe: 11% YoY revenue from zero', '→ WarrantyIQ: ~15% margin on plan revenue'],
  ].forEach((m,i) => {
    const y = 2.2+i*0.74;
    s.addShape('rect', { x:5.25, y, w:4.1, h:0.64, fill:{color:C.lgray}, line:{type:'none'} });
    s.addText(m[0], { x:5.35, y:y+0.06, w:3.9, h:0.22, fontSize:10, bold:true, color:C.ink, fontFace:'Calibri' });
    s.addText(m[1], { x:5.35, y:y+0.32, w:3.9, h:0.22, fontSize:10, color:C.green, fontFace:'Calibri' });
  });
  s.addShape('rect', { x:0.5, y:6.88, w:9, h:0.46, fill:{color:'F0FFF8'}, line:{color:C.green, width:1.5} });
  s.addText('"WarrantyIQ is the same monetization architecture as PhonePe Rewards — applied to Amazon H&L post-purchase."  — Ajay Avaghade', {
    x:0.7, y:6.94, w:8.6, h:0.3, fontSize:11, color:C.ink, fontFace:'Calibri', italic:true
  });
}

// SLIDE 8 — ROLLOUT PLAN
{
  const s = prs.addSlide();
  s.addShape('rect', { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('ROLLOUT PLAN', { x:0.5, y:0.28, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('From UAE Appliances to MENA H&L Warranty Platform', {
    x:0.5, y:0.62, w:9, h:0.7, fontSize:27, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape('rect', { x:0.5, y:1.38, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });
  [
    { label:'PHASE 1', period:'Month 1–2', title:'UAE Appliance Pilot', color:C.hero, acc:C.gold, items:['WarrantyIQ on washing machines, ACs, refrigerators', '3 plan tiers: 1yr / 2yr / 3yr with in-home SLA', 'Onboard 3 UAE authorised service centre partners', 'KPIs: attach rate, plan tier mix, claim rate'] },
    { label:'PHASE 2', period:'Month 3–4', title:'Scale + Optimize', color:C.green, acc:C.white, items:['Expand to all large-ticket H&L (TVs, dishwashers)', 'Launch WhatsApp claim filing (sub-2min flow)', 'Train failure-probability model on UAE product return data', 'A/B test: failure context card vs. generic warranty banner'] },
    { label:'PHASE 3', period:'Month 5–6', title:'MENA H&L Platform', color:C.hero, acc:C.gold, items:['Roll out to KSA, Egypt, Kuwait appliance categories', 'Launch brand-funded warranty tier (Samsung, Daikin co-fund)', 'WarrantyIQ dashboard: customer lifetime claim analytics', 'Publish playbook for global Amazon H&L teams'] },
  ].forEach((ph,i) => {
    const x = 0.4+i*3.1;
    s.addShape('rect', { x, y:1.56, w:2.95, h:4.55, fill:{color:ph.color}, line:{type:'none'}, shadow:mk() });
    s.addText(ph.label, { x:x+0.1, y:1.68, w:2.75, h:0.28, fontSize:9, bold:true, color:ph.acc, fontFace:'Calibri', charSpacing:3 });
    s.addText(ph.period, { x:x+0.1, y:1.96, w:2.75, h:0.26, fontSize:9, color:ph.color===C.green?'dcfce7':C.muted, fontFace:'Calibri' });
    s.addText(ph.title, { x:x+0.1, y:2.26, w:2.75, h:0.5, fontSize:17, bold:true, color:ph.acc, fontFace:'Calibri' });
    ph.items.forEach((item,j) => {
      s.addShape('ellipse', { x:x+0.12, y:2.9+j*0.74, w:0.12, h:0.12, fill:{color:ph.acc}, line:{type:'none'} });
      s.addText(item, { x:x+0.3, y:2.84+j*0.74, w:2.55, h:0.64, fontSize:10, color:ph.color===C.green?'dcfce7':C.muted, fontFace:'Calibri' });
    });
  });
  s.addShape('rect', { x:0.5, y:6.35, w:9, h:0.95, fill:{color:'1a2332'}, line:{color:C.green, width:1.5}, shadow:mk() });
  s.addText('What I need to build WarrantyIQ:', { x:0.7, y:6.44, w:4, h:0.35, fontSize:12, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('H&L product failure-rate data access  ·  UAE service centre partnership agreements  ·  Checkout plan attach API  ·  1 PM + 2 Engineers + 1 Partnerships Manager', {
    x:0.7, y:6.77, w:8.6, h:0.38, fontSize:10, color:C.muted, fontFace:'Calibri'
  });
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  ajay-avaghade.github.io/Portfolio', {
    x:0, y:7.2, w:W, h:0.26, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center'
  });
}

await prs.writeFile({ fileName: 'assets/amazon-warrantyiq-deck.pptx' });
console.log('Done: amazon-warrantyiq-deck.pptx');
})();
