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
  for(let i=0;i<12;i++){
    s.addShape(prs.ShapeType.line, { x:-1+i*1.1, y:0, w:0.6, h:H, line:{color:C.hero, width:18, transparency:85}, rotate:15 });
  }
  s.addShape(prs.ShapeType.rect, { x:7.5, y:0, w:2.5, h:1.2, fill:{color:C.bright}, line:{type:'none'} });
  s.addText('amazon pay', { x:7.55, y:0.1, w:2.3, h:0.45, fontSize:14, bold:true, color:C.dark, fontFace:'Calibri', align:'center' });
  s.addText('PAYMENTS · INDIA', { x:7.55, y:0.55, w:2.3, h:0.3, fontSize:9, bold:true, color:C.dark, fontFace:'Calibri', align:'center', charSpacing:3 });

  s.addText('Amazon Pay Milestone', { x:0.5, y:1.5, w:7, h:1.0, fontSize:48, bold:true, color:C.white, fontFace:'Calibri' });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:2.6, w:1.2, h:0.08, fill:{color:C.bright}, line:{type:'none'} });
  s.addText('From one-time IBD activation to a 30-day habit loop — funded by the bank partner', {
    x:0.5, y:2.75, w:8.5, h:0.5, fontSize:16, color:C.gold, fontFace:'Calibri', italic:true
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.8, w:5, h:0.06, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('Presented by Ajay Avaghade  ·  Growth & Payments PM  ·  Amazon Emerging Stores', {
    x:0.5, y:5.95, w:7, h:0.35, fontSize:11, color:C.muted, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:7.5, y:5.0, w:2.2, h:2.0, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('60–70%', { x:7.5, y:5.15, w:2.2, h:0.9, fontSize:34, bold:true, color:C.bright, fontFace:'Calibri', align:'center' });
  s.addText('of IBD-activated\nusers make only\n1 transaction', { x:7.5, y:6.0, w:2.2, h:1.1, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center' });
}

// ── SLIDE 2 — THE PROBLEM ────────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('THE PROBLEM', { x:0.5, y:0.3, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('IBD Activates Customers Once — Then Loses Them', {
    x:0.5, y:0.65, w:9, h:0.85, fontSize:30, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.55, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });

  const cols3 = [
    { icon:'📉', stat:'60–70%', lbl:'IBD users make only 1 txn', sub:'Activated by the offer,\nnever habituated to the card' },
    { icon:'🔁', stat:'1.3×', lbl:'Avg. monthly repeat rate', sub:'vs. 4.2× for loyalty-\nprogramme users globally' },
    { icon:'💔', stat:'High churn', lbl:'Bank partner renewal risk', sub:'Low repeat usage = weak\npartnership ROI data at renewal' },
  ];
  cols3.forEach((col,i) => {
    const x = 0.5 + i*3.1;
    s.addShape(prs.ShapeType.rect, { x, y:1.75, w:2.9, h:3.5, fill:{color:C.hero}, line:{type:'none'}, shadow:makeShadow() });
    s.addText(col.icon, { x, y:1.9, w:2.9, h:0.6, fontSize:28, align:'center' });
    s.addText(col.stat, { x, y:2.55, w:2.9, h:0.8, fontSize:36, bold:true, color:C.bright, fontFace:'Calibri', align:'center' });
    s.addText(col.lbl, { x, y:3.4, w:2.9, h:0.45, fontSize:12, bold:true, color:C.white, fontFace:'Calibri', align:'center' });
    s.addText(col.sub, { x, y:3.85, w:2.9, h:0.9, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center' });
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.5, w:9, h:1.75, fill:{color:'1a2332'}, line:{color:C.bright, width:1.5}, shadow:makeShadow() });
  s.addText('Root Cause:', { x:0.7, y:5.65, w:2, h:0.35, fontSize:11, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText("Amazon's IBD program is a one-time event: customer gets the discount, uses it once, and has no reason to return with the same card. There is no engagement mechanic, no progress loop, no escalating reward. Bank partners fund the activation but receive no loyalty data in return — making partnership renewals a one-way value exchange.", {
    x:0.7, y:6.0, w:8.6, h:1.05, fontSize:11, color:C.muted, fontFace:'Calibri'
  });
}

// ── SLIDE 3 — THE INSIGHT ────────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:C.lgray} });
  s.addText('THE INSIGHT', { x:0.5, y:0.3, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText("A Discount That Grows With Usage Turns Activation Into Habit", {
    x:0.5, y:0.65, w:9, h:0.85, fontSize:25, bold:true, color:C.ink, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.55, w:9, h:0.04, fill:{color:C.bright}, line:{type:'none'} });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.75, w:4.2, h:4.7, fill:{color:C.white}, line:{color:'E5E7EB', width:1}, shadow:makeShadow() });
  s.addText('❌  Flat IBD (Current)', { x:0.65, y:1.95, w:3.9, h:0.45, fontSize:14, bold:true, color:C.ink, fontFace:'Calibri' });
  const leftPts = [
    'Same 10% discount whether customer transacts once or ten times',
    'No progress indicator — customer has nothing to work toward',
    'Discount feels like a promotion, not a relationship',
    'Bank partner funds the discount but gains no loyalty data',
  ];
  leftPts.forEach((t,i) => {
    s.addShape(prs.ShapeType.rect, { x:0.65, y:2.55+i*0.7, w:0.22, h:0.22, fill:{color:'FEE2E2'}, line:{type:'none'} });
    s.addText('✕', { x:0.65, y:2.55+i*0.7, w:0.22, h:0.22, fontSize:10, bold:true, color:'DC2626', align:'center' });
    s.addText(t, { x:0.95, y:2.53+i*0.7, w:3.55, h:0.5, fontSize:11, color:C.muted, fontFace:'Calibri' });
  });

  s.addShape(prs.ShapeType.ellipse, { x:4.55, y:3.55, w:0.9, h:0.9, fill:{color:C.bright}, line:{type:'none'} });
  s.addText('VS', { x:4.55, y:3.55, w:0.9, h:0.9, fontSize:14, bold:true, color:C.dark, align:'center', valign:'middle', fontFace:'Calibri' });

  s.addShape(prs.ShapeType.rect, { x:5.3, y:1.75, w:4.2, h:4.7, fill:{color:C.white}, line:{color:C.bright, width:2}, shadow:makeShadow() });
  s.addText('✅  Milestone (Proposed)', { x:5.45, y:1.95, w:3.9, h:0.45, fontSize:14, bold:true, color:C.ink, fontFace:'Calibri' });
  const rightPts = [
    'Tier 1 (3 txns): 5% off · Tier 2 (10 txns): 8% off · Tier 3 (20 txns): 12% off',
    'Progress bar lives on Amazon Pay home — visible every session',
    'Nudges at 50%, 80%, 95% of each milestone create urgency',
    'Bank partner funds escalating tiers; receives transaction frequency data',
  ];
  rightPts.forEach((t,i) => {
    s.addShape(prs.ShapeType.rect, { x:5.45, y:2.55+i*0.7, w:0.22, h:0.22, fill:{color:'DCFCE7'}, line:{type:'none'} });
    s.addText('✓', { x:5.45, y:2.55+i*0.7, w:0.22, h:0.22, fontSize:10, bold:true, color:C.green, align:'center' });
    s.addText(t, { x:5.75, y:2.53+i*0.7, w:3.55, h:0.5, fontSize:11, color:C.muted, fontFace:'Calibri' });
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:6.6, w:9, h:0.7, fill:{color:'FFFBEB'}, line:{color:C.bright, width:1.5} });
  s.addText('Key Insight: ', { x:0.7, y:6.73, w:1.3, h:0.4, fontSize:11, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('Same psychological mechanic as a loyalty stamp card — but funded by the bank partner\'s marketing budget, not Amazon\'s. Amazon gets retention; the bank gets usage frequency data worth renewing for.', {
    x:1.9, y:6.73, w:7.4, h:0.4, fontSize:11, color:C.ink, fontFace:'Calibri'
  });
}

// ── SLIDE 4 — THE MECHANIC ───────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('THE MECHANIC', { x:0.5, y:0.3, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('3 Tiers · 30 Days · 1 Habit Loop', {
    x:0.5, y:0.65, w:9, h:0.7, fontSize:30, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.4, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });

  const steps = [
    { n:'01', title:'IBD Activation', body:'Customer activates bank discount (existing flow). Milestone window opens: 30-day progress tracking begins.' },
    { n:'02', title:'Tier 1 — 5% Off', body:'3 transactions in 30 days. Unlocked immediately. Sets the first anchor for habitual usage.' },
    { n:'03', title:'Tier 2 — 8% Off', body:'10 transactions. Progress bar + nudges at 50%/80%/95%. Creates active engagement loop.' },
    { n:'04', title:'Tier 3 — 12% Off', body:'20 transactions. Highest tier. Customers at this level show 4× monthly GMV vs. Tier 1.' },
    { n:'05', title:'Bank Partner Loop', body:'Transaction data shared with bank partner. Renewal negotiations backed by frequency + GMV data.' },
  ];
  steps.forEach((step,i) => {
    const x = 0.4 + i*1.88;
    const tcolor = i===2 ? C.bright : (i===3 ? C.green : C.gold);
    s.addShape(prs.ShapeType.ellipse, { x:x+0.54, y:1.65, w:0.7, h:0.7, fill:{color:tcolor}, line:{type:'none'}, shadow:makeShadow() });
    s.addText(step.n, { x:x+0.54, y:1.65, w:0.7, h:0.7, fontSize:13, bold:true, color:C.dark, align:'center', valign:'middle', fontFace:'Calibri' });
    if(i<4){
      s.addShape(prs.ShapeType.line, { x:x+1.24, y:2.0, w:0.64, h:0, line:{color:C.hero, width:1.5, dashType:'dash'} });
    }
    s.addShape(prs.ShapeType.rect, { x, y:2.55, w:1.78, h:3.5, fill:{color:C.hero}, line:{type:'none'}, shadow:makeShadow() });
    s.addText(step.title, { x:x+0.05, y:2.7, w:1.68, h:0.45, fontSize:12, bold:true, color:tcolor, fontFace:'Calibri', align:'center' });
    s.addText(step.body, { x:x+0.05, y:3.2, w:1.68, h:2.5, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center' });
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:6.35, w:9, h:0.9, fill:{color:'1a2332'}, line:{color:C.bright, width:1}, shadow:makeShadow() });
  s.addText('PhonePe Proof  ·  ', { x:0.7, y:6.5, w:2.1, h:0.5, fontSize:11, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('Built the identical milestone mechanic at PhonePe for POS device merchants: Rent Waiver unlocked at TPV + TPC goals. 60% improvement in 90-day device activity. Cart incentivisation engine: 60% cart abandonment reduction, 20% improvement in 30-day retention. Same behavioural loop, new context.', {
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
    { n:'01', title:'Amazon Pay Home', desc:'Milestone banner surfaced on Pay home screen. Progress visible without navigation. Creates daily recall loop.', mock:['amazon pay', '─────────────', 'Balance: ₹1,240', 'HDFC ****4521 linked', '─────────────', '🏆 HDFC MILESTONE', '4 more → 8% off', '4/10 done [40%░░░]', '[View Milestones →]'] },
    { n:'02', title:'Milestone Tracker', desc:'3 tiers displayed with live progress bars, transaction history, and countdown. Nudge text shows exact gap.', mock:['HDFC Milestone', '9 days left', '─────────────', 'Tier 1 ✓ 5% OFF', '[████████] 100%', '─────────────', 'Tier 2 ⚡ 8% OFF', '[████░░░░] 40%', '6 more transactions'] },
    { n:'03', title:'Tier Unlocked', desc:'Confetti + new discount rate confirmed. Previous rate shown (anchoring). Next tier progress immediately visible.', mock:['🏆 MILESTONE!', '─────────────', 'Tier 2 Complete!', '─────────────', 'New rate: 8% OFF', '(was: 5% off)', '─────────────', 'Est. saving ₹820/mo', '[Keep Shopping →]'] },
    { n:'04', title:'Bank Partner Ops', desc:'Dashboard for HDFC partner team: tier completion rates, nudge performance, GMV by tier. Drives renewal conversations.', mock:['Milestone Program ●', 'HDFC · May 2025', '─────────────', 'Active users: 84.2K', 'Repeat rate: 3.1×', 'GMV this month: ₹8.4Cr', '─────────────', 'Tier 2 completion: 41%'] },
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
  s.addText('Projected Impact — Built on PhonePe & Virtusa Proof Points', {
    x:0.5, y:0.65, w:9, h:0.7, fontSize:24, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.4, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.6, w:4.3, h:0.4, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('CUSTOMER IMPACT', { x:0.5, y:1.6, w:4.3, h:0.4, fontSize:10, bold:true, color:C.gold, align:'center', valign:'middle', fontFace:'Calibri' });
  const custMetrics = [
    { val:'3×', lbl:'30-day repeat transaction rate', sub:'vs. flat IBD baseline' },
    { val:'25%', lbl:'Avg. monthly GMV per user', sub:'Tier 2 completers vs. Tier 1' },
    { val:'20%', lbl:'Improvement in 30-day retention', sub:'PhonePe cart engine baseline' },
    { val:'₹820/mo', lbl:'Avg. user savings at Tier 2', sub:'Based on ₹10,250 monthly HDFC spend' },
  ];
  custMetrics.forEach((m,i) => {
    const y = 2.1 + i*1.15;
    s.addShape(prs.ShapeType.rect, { x:0.5, y, w:4.3, h:1.0, fill:{color:C.hero}, line:{type:'none'}, shadow:makeShadow() });
    s.addText(m.val, { x:0.6, y:y+0.08, w:4.1, h:0.5, fontSize:28, bold:true, color:C.bright, fontFace:'Calibri' });
    s.addText(m.lbl, { x:0.6, y:y+0.56, w:2.5, h:0.3, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(m.sub, { x:0.6, y:y+0.56, w:4.1, h:0.3, fontSize:10, color:C.muted, fontFace:'Calibri', align:'right' });
  });

  s.addShape(prs.ShapeType.line, { x:4.95, y:1.6, w:0, h:5.65, line:{color:C.hero, width:1, dashType:'dash'} });

  s.addShape(prs.ShapeType.rect, { x:5.1, y:1.6, w:4.4, h:0.4, fill:{color:C.hero}, line:{type:'none'} });
  s.addText('AMAZON + BANK PARTNER ROI', { x:5.1, y:1.6, w:4.4, h:0.4, fontSize:10, bold:true, color:C.gold, align:'center', valign:'middle', fontFace:'Calibri' });
  const roiMetrics = [
    { val:'40%', lbl:'Bank partner renewal rate uplift', sub:'Usage frequency data vs. one-time activation' },
    { val:'2.8×', lbl:'GMV per Tier 2 user vs. Tier 1', sub:'Tier completion drives compounding spend' },
    { val:'₹0 net cost', lbl:'To Amazon for the discount', sub:'Bank partner funds all tier rewards' },
    { val:'60%', lbl:'90-day device activity improvement', sub:'Virtusa Milestone Waiver: identical mechanic' },
  ];
  roiMetrics.forEach((m,i) => {
    const y = 2.1 + i*1.15;
    s.addShape(prs.ShapeType.rect, { x:5.1, y, w:4.4, h:1.0, fill:{color:C.hero}, line:{type:'none'}, shadow:makeShadow() });
    s.addText(m.val, { x:5.2, y:y+0.08, w:4.2, h:0.5, fontSize:28, bold:true, color:C.bright, fontFace:'Calibri' });
    s.addText(m.lbl, { x:5.2, y:y+0.56, w:2.8, h:0.3, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(m.sub, { x:5.2, y:y+0.56, w:4.2, h:0.3, fontSize:10, color:C.muted, fontFace:'Calibri', align:'right' });
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:6.75, w:9, h:0.6, fill:{color:'1a2332'}, line:{color:C.bright, width:1} });
  s.addText('Amazon bears zero discount cost — the bank partner funds the tiers in exchange for transaction frequency data that makes their card more valuable to renew.', {
    x:0.65, y:6.82, w:8.7, h:0.4, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true
  });
}

// ── SLIDE 7 — PROOF OF WORK ──────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:C.lgray} });
  s.addText('PROOF OF WORK', { x:0.5, y:0.3, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText("I've Shipped This Mechanic Twice. Here's the Evidence.", {
    x:0.5, y:0.65, w:9, h:0.7, fontSize:25, bold:true, color:C.ink, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.4, w:9, h:0.04, fill:{color:C.bright}, line:{type:'none'} });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.6, w:4.3, h:5.0, fill:{color:C.dark}, line:{type:'none'}, shadow:makeShadow() });
  s.addText('WHAT I BUILT', { x:0.65, y:1.75, w:4.0, h:0.35, fontSize:9, bold:true, color:C.bright, fontFace:'Calibri', charSpacing:3 });
  const builtItems = [
    'PhonePe cart incentivisation engine: context-aware triggers across cart value × user intent × time → 35% AOV uplift, 60% cart abandonment reduction, 20% 30-day retention improvement',
    'PhonePe Nudge Engine: threshold-based push triggers at 50%/80%/95% progress — identical to Milestone nudge architecture',
    'Virtusa MerchantOS: Milestone Rental Waiver — 3-tier goal mechanic (TPV + TPC targets), ops dashboard, A/B tested vs. cashback variant. 60% improvement in 90-day device activity',
    'PhonePe ML-driven rewards marketplace: audience segmentation + brand self-serve → 26% user engagement improvement, 11% YoY revenue growth',
  ];
  builtItems.forEach((item,i) => {
    s.addShape(prs.ShapeType.rect, { x:0.65, y:2.25+i*0.88, w:0.2, h:0.2, fill:{color:C.bright}, line:{type:'none'} });
    s.addText(item, { x:0.93, y:2.21+i*0.88, w:3.75, h:0.75, fontSize:10, color:C.muted, fontFace:'Calibri' });
  });

  s.addShape(prs.ShapeType.rect, { x:5.1, y:1.6, w:4.4, h:5.0, fill:{color:C.white}, line:{color:'E5E7EB', width:1}, shadow:makeShadow() });
  s.addText('HOW IT MAPS TO THIS ROLE', { x:5.25, y:1.75, w:4.1, h:0.35, fontSize:9, bold:true, color:C.ink, fontFace:'Calibri', charSpacing:3 });
  const mappings = [
    ['JD: End-to-end IBD lifecycle ownership', '→ Built offers stack E2E at PhonePe'],
    ['JD: Drive customer behavior change', '→ Cart engine: 60% abandonment reduction'],
    ['JD: Launch in new MENA/APAC markets', '→ Milestone Waiver: MENA bank expansion plan'],
    ['JD: Partner with bank teams', '→ 500+ brand partners onboarded at PhonePe'],
    ['JD: Mentor PMs, foster innovation', '→ Led 6-squad PM team across offers charter'],
  ];
  mappings.forEach((m,i) => {
    const y = 2.25 + i*0.75;
    s.addShape(prs.ShapeType.rect, { x:5.25, y, w:4.1, h:0.65, fill:{color:C.lgray}, line:{type:'none'} });
    s.addText(m[0], { x:5.35, y:y+0.07, w:3.9, h:0.22, fontSize:10, bold:true, color:C.ink, fontFace:'Calibri' });
    s.addText(m[1], { x:5.35, y:y+0.34, w:3.9, h:0.22, fontSize:10, color:C.green, fontFace:'Calibri' });
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:6.85, w:9, h:0.5, fill:{color:'FFFBEB'}, line:{color:C.bright, width:1.5} });
  s.addText('"Milestone isn\'t a concept — I\'ve shipped the mechanic at PhonePe and Virtusa. Amazon is the next context."  — Ajay Avaghade', {
    x:0.7, y:6.92, w:8.6, h:0.33, fontSize:11, color:C.ink, fontFace:'Calibri', italic:true
  });
}

// ── SLIDE 8 — ROLLOUT PLAN ───────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:C.dark} });
  s.addText('ROLLOUT PLAN', { x:0.5, y:0.3, w:9, h:0.3, fontSize:9, bold:true, color:C.muted, fontFace:'Calibri', charSpacing:4 });
  s.addText('3-Phase Rollout: India → MENA → Full Emerging Markets', {
    x:0.5, y:0.65, w:9, h:0.7, fontSize:26, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.4, w:9, h:0.04, fill:{color:C.hero}, line:{type:'none'} });

  const phases = [
    { label:'PHASE 1', period:'Month 1–2', title:'India Pilot', color:C.hero, accent:C.gold, items:['Launch Milestone with HDFC + Axis (highest IBD volume)', 'A/B test: Milestone vs. flat IBD for activated users', 'Nudge engine: deploy 50%/80%/95% threshold triggers', 'Measure: 30-day repeat rate, Tier 2 completion rate'] },
    { label:'PHASE 2', period:'Month 3–4', title:'Scale + MENA Launch', color:C.bright, accent:C.dark, items:['Scale to all India bank partners (ICICI, SBI, Kotak)', 'Launch in UAE (Mashreq, FAB) and Saudi (Riyad, NCB)', 'Bank partner ops dashboard: self-serve tier management', 'Negotiate Tier 4 with HDFC based on pilot data'] },
    { label:'PHASE 3', period:'Month 5–6', title:'Full Emerging Market', color:C.hero, accent:C.gold, items:['Expand to Egypt, Singapore, Malaysia, Thailand', 'Adapt milestone thresholds per market (local txn frequency)', 'Publish Milestone playbook for global IBD programs', 'Annual bank partner review: Milestone as renewal anchor'] },
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

  s.addShape(prs.ShapeType.rect, { x:0.5, y:6.35, w:9, h:1.0, fill:{color:'1a2332'}, line:{color:C.bright, width:1.5}, shadow:makeShadow() });
  s.addText('What I need to build Amazon Pay Milestone:', { x:0.7, y:6.45, w:4.5, h:0.35, fontSize:12, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('Bank partner API for transaction frequency data  ·  Nudge engine infra  ·  Progress tracker component in Amazon Pay app  ·  1 PM + 2 Engineers + 1 Designer', {
    x:0.7, y:6.78, w:8.6, h:0.4, fontSize:10, color:C.muted, fontFace:'Calibri'
  });
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439', {
    x:0, y:7.2, w:W, h:0.28, fontSize:10, color:C.muted, fontFace:'Calibri', align:'center'
  });
}

prs.writeFile({ fileName: 'assets/amazon-milestone-deck.pptx' })
  .then(() => console.log('✅ Deck written: amazon-milestone-deck.pptx'))
  .catch(e => console.error('❌', e));
