const pptx = require('pptxgenjs');
const prs = new pptx();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:  '1A0A0A',
  hero:  '2D1515',
  red:   'E60012',
  acc:   'FF3347',
  gold:  'FF6B00',
  white: 'FFFFFF',
  lgray: 'FFF5F5',
  muted: '8B9E9B',
  ink:   '1A1A1A',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });

// ─── SLIDE 1 — COVER ───────────────────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.dark };
  for(let i=0;i<6;i++){
    sld.addShape(prs.ShapeType.rect, { x:i*1.85-0.4, y:0, w:0.9, h:7.5, fill:{ color:C.hero }, line:{ color:'000000', width:0 } });
  }
  sld.addText('PAYPAY — GROWTH PLATFORM', { x:0.5, y:0.42, w:7, h:0.28, fontSize:9, bold:true, color:C.red, charSpacing:3, fontFace:'Arial' });
  sld.addText('PayPay SmartBonus', { x:0.5, y:0.85, w:9, h:1.05, fontSize:50, bold:true, color:C.white, fontFace:'Arial' });
  sld.addText('From Blanket Cashback to Behaviour-Linked Personalised Rewards', { x:0.5, y:1.98, w:7.2, h:0.75, fontSize:16, color:C.acc, fontFace:'Arial' });
  sld.addText('Presented by Ajay Avaghade · Growth & Platform PM', { x:0.5, y:2.88, w:6, h:0.32, fontSize:11, color:C.muted, fontFace:'Arial' });
  sld.addShape(prs.ShapeType.roundRect, { x:7.1, y:5.4, w:2.55, h:1.85, fill:{ color:C.hero }, line:{ color:C.red, width:2 }, rectRadius:0.12 });
  sld.addText('¥1T+', { x:7.1, y:5.45, w:2.55, h:0.72, fontSize:36, bold:true, color:C.red, fontFace:'Arial', align:'center' });
  sld.addText('Annual bonus campaign spend\n(blanket — no personalisation)', { x:7.1, y:6.2, w:2.55, h:0.75, fontSize:9, color:C.muted, fontFace:'Arial', align:'center' });
  sld.addShape(prs.ShapeType.rect, { x:0, y:7.08, w:10, h:0.42, fill:{ color:C.hero }, line:{ color:'000000', width:0 } });
  sld.addText('CASE STUDY  ·  GROWTH & MONETISATION  ·  PAYPAY JAPAN', { x:0.3, y:7.12, w:9.4, h:0.28, fontSize:8, color:C.muted, charSpacing:2, fontFace:'Arial', align:'center' });
}

// ─── SLIDE 2 — THE PROBLEM ─────────────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.dark };
  sld.addText('THE PROBLEM', { x:0.5, y:0.3, w:9, h:0.28, fontSize:9, bold:true, color:C.red, charSpacing:4, fontFace:'Arial' });
  sld.addText('Blanket Cashback Burns Billions on Users\nWho Would Have Paid Anyway.', { x:0.5, y:0.6, w:9, h:0.95, fontSize:27, bold:true, color:C.white, fontFace:'Arial' });
  const cols2 = [
    { icon:'💸', stat:'¥1T+', lbl:'Annual bonus spend', sub:'Applied equally regardless of user behaviour or lapse risk' },
    { icon:'📉', stat:'0%', lbl:'Behaviour change for power users', sub:'Users transacting 40x/mo get the same bonus as dormant users' },
    { icon:'🎯', stat:'<30%', lbl:'Campaign spend that changes behaviour', sub:'Majority of bonus budget spent on users who would have transacted anyway' },
  ];
  cols2.forEach((col,i)=>{
    const x = 0.38 + i*3.18;
    sld.addShape(prs.ShapeType.roundRect, { x, y:1.78, w:2.92, h:3.55, fill:{ color:C.hero }, line:{ color:'333333', width:1 }, rectRadius:0.12, shadow:makeShadow() });
    sld.addText(col.icon, { x, y:1.95, w:2.92, h:0.55, fontSize:26, fontFace:'Arial', align:'center' });
    sld.addText(col.stat, { x, y:2.55, w:2.92, h:0.72, fontSize:34, bold:true, color:C.red, fontFace:'Arial', align:'center' });
    sld.addText(col.lbl, { x, y:3.3, w:2.92, h:0.38, fontSize:11, bold:true, color:C.white, fontFace:'Arial', align:'center' });
    sld.addText(col.sub, { x:x+0.14, y:3.72, w:2.64, h:1.4, fontSize:9, color:C.muted, fontFace:'Arial', align:'center' });
  });
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:5.5, w:9.24, h:1.4, fill:{ color:C.hero }, line:{ color:'333333', width:1 }, rectRadius:0.1 });
  sld.addText('Root cause: PayPay has no user-level propensity model. Every campaign is a flat rate applied to every user — power users get subsidised for doing nothing different, while lapsing users receive the same insufficient nudge they have been ignoring for weeks.', { x:0.58, y:5.6, w:8.84, h:1.1, fontSize:10.5, color:C.lgray, fontFace:'Arial', italic:true });
}

// ─── SLIDE 3 — THE INSIGHT ─────────────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.lgray };
  sld.addText('THE INSIGHT', { x:0.5, y:0.3, w:9, h:0.28, fontSize:9, bold:true, color:C.red, charSpacing:4, fontFace:'Arial' });
  sld.addText('Blanket Cashback Is a Loyalty Tax.\nSmartBonus Only Pays When It Changes Behaviour.', { x:0.5, y:0.6, w:9, h:0.95, fontSize:24, bold:true, color:C.ink, fontFace:'Arial' });
  // Left panel
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:1.72, w:4.25, h:4.55, fill:{ color:C.white }, line:{ color:'DDDDDD', width:1 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('Blanket Campaign (Current)', { x:0.55, y:1.9, w:3.9, h:0.38, fontSize:12, bold:true, color:'CC3300', fontFace:'Arial' });
  const curr = ['Same 10% rate for 70M users','Power user (40 txns/mo): would pay anyway — bonus wasted','Dormant user (1 txn/mo): 10% not enough to re-engage','No learning loop — same mistake next campaign','P&L hit: billions spent, behaviour unchanged'];
  curr.forEach((t,i)=>{ sld.addText(t, { x:0.58, y:2.42+i*0.6, w:3.9, h:0.54, fontSize:9.5, color:C.ink, fontFace:'Arial' }); });
  // VS
  sld.addShape(prs.ShapeType.ellipse, { x:4.65, y:3.58, w:0.7, h:0.7, fill:{ color:C.red }, line:{ color:'000000', width:0 } });
  sld.addText('VS', { x:4.65, y:3.58, w:0.7, h:0.7, fontSize:11, bold:true, color:C.white, fontFace:'Arial', align:'center', valign:'middle' });
  // Right panel
  sld.addShape(prs.ShapeType.roundRect, { x:5.37, y:1.72, w:4.25, h:4.55, fill:{ color:C.dark }, line:{ color:C.red, width:1.5 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('SmartBonus (Proposed)', { x:5.52, y:1.9, w:3.9, h:0.38, fontSize:12, bold:true, color:C.red, fontFace:'Arial' });
  const prop = ['ML score per user: recency x frequency x lapse risk','Power user (40 txns/mo): 0.5% — they need no nudge','Dormant user (1 txn/mo): 12% — minimum effective dose','Self-improving: each campaign updates the model','Result: same budget, 3x the behaviour-change ROI'];
  prop.forEach((t,i)=>{ sld.addText(t, { x:5.52, y:2.42+i*0.6, w:3.9, h:0.54, fontSize:9.5, color:C.lgray, fontFace:'Arial' }); });
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:6.42, w:9.24, h:0.85, fill:{ color:C.white }, line:{ color:C.red, width:1.5 }, rectRadius:0.08 });
  sld.addText('Key insight: The most loyal users need the smallest incentive. The most at-risk users need the largest. Blanket cashback gets both exactly wrong. SmartBonus makes the subsidy proportional to the need.', { x:0.58, y:6.52, w:8.84, h:0.68, fontSize:10, color:C.ink, fontFace:'Arial', italic:true });
}

// ─── SLIDE 4 — THE MECHANIC ────────────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.dark };
  sld.addText('THE MECHANIC', { x:0.5, y:0.3, w:9, h:0.28, fontSize:9, bold:true, color:C.red, charSpacing:4, fontFace:'Arial' });
  sld.addText('SmartBonus Engine + MissionBoard: Two Levers, One Flywheel', { x:0.5, y:0.6, w:9, h:0.65, fontSize:26, bold:true, color:C.white, fontFace:'Arial' });
  const steps = [
    { n:'01', title:'Score User', desc:'ML model scores every user daily: transaction recency x frequency x category variety x lapse probability. Produces a propensity score 0-100.' },
    { n:'02', title:'Assign Rate', desc:'Score maps to minimum effective bonus rate. Power users (score 80+): 0.5%. Regular (40-79): 3%. At-risk (10-39): 8%. Dormant (<10): 12-15%.' },
    { n:'03', title:'Mission Layer', desc:'MissionBoard gives users agency: complete 5 weekly missions (vending machine, new restaurant, morning payment) to unlock a higher rate tier for that session.' },
    { n:'04', title:'Deliver Bonus', desc:'Personalised rate applied automatically at payment. User sees their rate on home screen. No opt-in required — SmartBonus is always on.' },
    { n:'05', title:'Model Updates', desc:'Each completed mission and payment feeds back into the model. Campaign ROI is measured at user level. Budget allocation shifts to highest-marginal-return segments.' },
  ];
  steps.forEach((st,i)=>{
    const x = 0.38 + i*1.9;
    sld.addShape(prs.ShapeType.roundRect, { x, y:1.65, w:1.72, h:4.75, fill:{ color:C.hero }, line:{ color: i===2 ? C.gold : '333333', width: i===2 ? 2 : 1 }, rectRadius:0.1, shadow:makeShadow() });
    sld.addShape(prs.ShapeType.ellipse, { x:x+0.56, y:1.78, w:0.6, h:0.6, fill:{ color: i===2 ? C.gold : C.red }, line:{ color:'000000', width:0 } });
    sld.addText(st.n, { x:x+0.56, y:1.78, w:0.6, h:0.6, fontSize:11, bold:true, color:C.white, fontFace:'Arial', align:'center', valign:'middle' });
    sld.addText(st.title, { x:x+0.08, y:2.48, w:1.56, h:0.5, fontSize:10, bold:true, color: i===2 ? C.gold : C.white, fontFace:'Arial', align:'center' });
    sld.addText(st.desc, { x:x+0.1, y:3.05, w:1.52, h:3.3, fontSize:8, color:C.muted, fontFace:'Arial' });
  });
  sld.addText('Proof at PhonePe: Propensity-to-Transact ML replaced static targeting for 350M MAU — 32% marketing burn reduction while sustaining GMV growth.', { x:0.38, y:6.6, w:9.24, h:0.55, fontSize:9.5, color:C.muted, fontFace:'Arial', italic:true });
}

// ─── SLIDE 5 — THE PRODUCT ─────────────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.lgray };
  sld.addText('THE PRODUCT', { x:0.5, y:0.3, w:9, h:0.28, fontSize:9, bold:true, color:C.red, charSpacing:4, fontFace:'Arial' });
  sld.addText('4 Key Screen States — Built on Proven PhonePe Architecture', { x:0.5, y:0.6, w:9, h:0.65, fontSize:24, bold:true, color:C.ink, fontFace:'Arial' });
  const screens = [
    { n:'01', title:'Home', color:C.red, desc:'Personalised "Your bonus today: 3%" card surfaces prominently on the PayPay home screen alongside the MissionBoard widget showing 3/5 weekly missions complete.', ascii:'[PayPay Home]\n\nBalance: ¥12,430\n\n┌─ SmartBonus ─────┐\n│ ◉ 3% Today       │\n│ Your rate · tap  │\n│ +2 missions→ 5%  │\n└──────────────────┘\n\n┌─ MissionBoard ───┐\n│ ▓▓▓▓▓░░░ 3/5   │\n│ Tue deadline     │\n│ Unlock 5% ✨    │\n└──────────────────┘' },
    { n:'02', title:'SmartBonus Hub', color:C.red, desc:'Big rate display with scoring breakdown. 5 missions listed with done/active/locked states. Step tier progress bar shows path to Gold. Tap mission for detail.', ascii:'[SmartBonus]\n\nYour rate today\n    3%\nApplied on all txns\n✨ 2 missions → 5%\n\nWhy you got 3%\nFrequency ▓▓▓▓▓▓░ 68%\nVariety   ▓▓▓░░░░ 45%\nRecency   ▓▓▓▓▓▓▓ 80%\n\n✅ First txn today\n✅ Vending machine\n🍜 New restaurant cat.\n☀️ 3 txns before noon\n🔒 Weekend streak' },
    { n:'03', title:'Rate Unlocked', color:C.acc, desc:'Confetti celebration when all 5 missions complete. Rate upgrades from 3% to 5%. Step tier Gold unlocked. PayPay Step bonus applied automatically.', ascii:'[Mission Complete!]\n\n    ✓ (confetti)\n\n  Missions unlocked!\n  Vending Machine Pro\n\n┌─ Rate Upgraded ──┐\n│  3% ──────► 5%   │\n│  Next 10 txns    │\n└──────────────────┘\n\n  Missions: 5/5 ✨\n  Bonus earned: ¥620\n  Streak: 3 weeks 🏆\n\n  PayPay Step Gold\n  unlocked! +1.5% base' },
    { n:'04', title:'Ops Dashboard', color:C.gold, desc:'Internal view: segment bonus rate heatmap, campaign ROI vs blanket baseline, mission completion rates by type, and stuck-user nudge queue with push composer.', ascii:'[SmartBonus Ops]\n\n−29% Burn  +23% WAU\n3.1× ROI vs blanket\n\nRates by Segment\nPower    ▓░░░░  0.5%\nRegular  ▓▓▓░░  3.0%\nAt-risk  ▓▓▓▓▓  7.5%\nDormant  ▓▓▓▓▓▓ 12%\n\nMission Completion\nDay txn  ▓▓▓▓▓ 89%\nVending  ▓▓▓▓░ 71%\nRestaurant ▓▓░░ 44%\n\nNudge Queue: 48K' },
  ];
  screens.forEach((sc,i)=>{
    const x = 0.3 + i*2.42;
    sld.addShape(prs.ShapeType.roundRect, { x, y:1.4, w:2.22, h:5.7, fill:{ color:C.white }, line:{ color:'DDDDDD', width:1 }, rectRadius:0.1, shadow:makeShadow() });
    sld.addShape(prs.ShapeType.rect, { x, y:1.4, w:2.22, h:0.38, fill:{ color:sc.color }, line:{ color:'000000', width:0 } });
    sld.addText(`${sc.n}  ${sc.title}`, { x:x+0.08, y:1.42, w:2.06, h:0.32, fontSize:9, bold:true, color:C.white, fontFace:'Arial' });
    sld.addShape(prs.ShapeType.roundRect, { x:x+0.1, y:1.86, w:2.02, h:2.3, fill:{ color:C.dark }, line:{ color:C.hero, width:1 }, rectRadius:0.06 });
    sld.addText(sc.ascii, { x:x+0.14, y:1.9, w:1.94, h:2.22, fontSize:5.8, color:C.lgray, fontFace:'Courier New', valign:'top' });
    sld.addText(sc.desc, { x:x+0.1, y:4.26, w:2.02, h:2.65, fontSize:8.5, color:C.ink, fontFace:'Arial', valign:'top' });
  });
  sld.addText('Prototype: paypay-smartbonus-prototype.html  ·  Live interactive demo of all 4 states', { x:0.5, y:7.1, w:9, h:0.28, fontSize:8.5, color:C.muted, fontFace:'Arial', italic:true, align:'center' });
}

// ─── SLIDE 6 — IMPACT & ROI ────────────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.dark };
  sld.addText('IMPACT & ROI', { x:0.5, y:0.3, w:9, h:0.28, fontSize:9, bold:true, color:C.red, charSpacing:4, fontFace:'Arial' });
  sld.addText('Projected Impact — Anchored to PhonePe ML Proof Points', { x:0.5, y:0.6, w:9, h:0.65, fontSize:24, bold:true, color:C.white, fontFace:'Arial' });
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:1.4, w:4.5, h:5.38, fill:{ color:C.hero }, line:{ color:'333333', width:1 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('User Impact', { x:0.55, y:1.55, w:4.15, h:0.38, fontSize:12, bold:true, color:C.red, fontFace:'Arial' });
  const um = [
    { val:'+23%', lbl:'Weekly active payers', sub:'Mission layer drives habit across new payment contexts' },
    { val:'+40%', lbl:'Step tier upgrade rate', sub:'Guided path replaces abstract 30-txn goal with 5 weekly missions' },
    { val:'-20%', lbl:'30-day churn for completers', sub:'Mission completers are 2x more likely to stay active next month' },
    { val:'+5%', lbl:'Rate perceived as personal', sub:'Users who see a personalised rate report higher PayPay satisfaction' },
  ];
  um.forEach((m,i)=>{
    sld.addShape(prs.ShapeType.roundRect, { x:0.55, y:2.08+i*1.1, w:4.15, h:0.96, fill:{ color:C.dark }, line:{ color:'333333', width:1 }, rectRadius:0.07 });
    sld.addText(m.val, { x:0.7, y:2.12+i*1.1, w:1.3, h:0.52, fontSize:24, bold:true, color:C.red, fontFace:'Arial', align:'center', valign:'middle' });
    sld.addText(m.lbl, { x:2.08, y:2.15+i*1.1, w:2.45, h:0.28, fontSize:10, bold:true, color:C.white, fontFace:'Arial' });
    sld.addText(m.sub, { x:2.08, y:2.44+i*1.1, w:2.45, h:0.28, fontSize:8.5, color:C.muted, fontFace:'Arial' });
  });
  sld.addShape(prs.ShapeType.roundRect, { x:5.12, y:1.4, w:4.5, h:5.38, fill:{ color:C.hero }, line:{ color:'333333', width:1 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('PayPay ROI', { x:5.28, y:1.55, w:4.15, h:0.38, fontSize:12, bold:true, color:C.gold, fontFace:'Arial' });
  const pm = [
    { val:'-30%', lbl:'Campaign burn per GMV unit', sub:'Bonus routed only to users where it changes behaviour' },
    { val:'3.1x', lbl:'Incremental GMV per bonus yen', sub:'Same total spend, 3x the behaviour-change return' },
    { val:'+18%', lbl:'GMV from mission-active users', sub:'Mission completers transact in new categories PayPay did not own before' },
    { val:'<6mo', lbl:'Payback period for ML build', sub:'Model cost amortised within two campaign cycles at scale' },
  ];
  pm.forEach((m,i)=>{
    sld.addShape(prs.ShapeType.roundRect, { x:5.28, y:2.08+i*1.1, w:4.15, h:0.96, fill:{ color:C.dark }, line:{ color:'333333', width:1 }, rectRadius:0.07 });
    sld.addText(m.val, { x:5.4, y:2.12+i*1.1, w:1.3, h:0.52, fontSize:24, bold:true, color:C.gold, fontFace:'Arial', align:'center', valign:'middle' });
    sld.addText(m.lbl, { x:6.76, y:2.15+i*1.1, w:2.45, h:0.28, fontSize:10, bold:true, color:C.white, fontFace:'Arial' });
    sld.addText(m.sub, { x:6.76, y:2.44+i*1.1, w:2.45, h:0.28, fontSize:8.5, color:C.muted, fontFace:'Arial' });
  });
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:6.95, w:9.24, h:0.38, fill:{ color:C.hero }, line:{ color:'333333', width:1 }, rectRadius:0.07 });
  sld.addText('At PhonePe this exact ML architecture reduced a 1000+ Cr/yr marketing budget burn by 32% while sustaining GMV. SmartBonus applies the same model to PayPay at 70M users.', { x:0.58, y:6.98, w:8.84, h:0.3, fontSize:9, color:C.muted, fontFace:'Arial', italic:true });
}

// ─── SLIDE 7 — PROOF OF WORK ───────────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.lgray };
  sld.addText('PROOF OF WORK', { x:0.5, y:0.3, w:9, h:0.28, fontSize:9, bold:true, color:C.red, charSpacing:4, fontFace:'Arial' });
  sld.addText('I Built This. Here Is the Proof.', { x:0.5, y:0.6, w:9, h:0.65, fontSize:24, bold:true, color:C.ink, fontFace:'Arial' });
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:1.42, w:4.4, h:5.08, fill:{ color:C.dark }, line:{ color:'333333', width:1 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('What I Built at PhonePe', { x:0.55, y:1.58, w:4.06, h:0.38, fontSize:12, bold:true, color:C.red, fontFace:'Arial' });
  const pp = [
    'Deployed Propensity-to-Transact ML models replacing 1000+ Cr/yr of static rule-based targeting with real-time user-level scoring and personalised incentive matching',
    'Delivered 32% marketing burn reduction while sustaining GMV growth across 350M MAU — the exact same architecture as SmartBonus at PayPay',
    'Built dynamic cart incentivisation engine with context-aware triggers: cart value x user intent x time signals — 35% AOV uplift, 60% cart abandonment reduction',
    'Rebuilt 100 Cr/yr rewards portfolio into ML-driven marketplace with audience segmentation and brand self-serve tooling — 26% uplift in user engagement',
    'Led end-to-end offers redesign with A/B experimentation across Tech, UX, Finance, and Legal — 22% checkout conversion lift at national scale',
  ];
  pp.forEach((t,i)=>{ sld.addText(`${String.fromCharCode(9658)}  ${t}`, { x:0.55, y:2.12+i*0.82, w:4.06, h:0.75, fontSize:9, color:C.lgray, fontFace:'Arial' }); });
  sld.addShape(prs.ShapeType.roundRect, { x:5.22, y:1.42, w:4.4, h:5.08, fill:{ color:C.white }, line:{ color:'DDDDDD', width:1 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('How It Maps to This PayPay Role', { x:5.38, y:1.58, w:4.06, h:0.38, fontSize:12, bold:true, color:C.ink, fontFace:'Arial' });
  const mp = [
    { from:'Propensity-to-Transact ML at 350M MAU', to:'SmartBonus scoring engine at PayPay 70M MAU' },
    { from:'32% marketing burn reduction at PhonePe', to:'Projected 30% campaign efficiency gain at PayPay' },
    { from:'Context-aware trigger engine (cart x intent x time)', to:'SmartBonus rate assignment (score x lapse x category)' },
    { from:'MerchantOS milestone waiver (5-step goals)', to:'MissionBoard (5 weekly missions, Step tier unlock)' },
    { from:'ML marketplace with audience segmentation', to:'Mission-level personalisation and nudge engine' },
  ];
  mp.forEach((m,i)=>{
    sld.addText(`PhonePe: ${m.from}`, { x:5.38, y:2.12+i*0.82, w:4.06, h:0.3, fontSize:8.5, color:'888888', fontFace:'Arial' });
    sld.addText(`${String.fromCharCode(8594)}  PayPay: ${m.to}`, { x:5.38, y:2.42+i*0.82, w:4.06, h:0.3, fontSize:9, bold:true, color:C.ink, fontFace:'Arial' });
  });
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:6.65, w:9.24, h:0.95, fill:{ color:'FFEBEB' }, line:{ color:C.red, width:1.5 }, rectRadius:0.08 });
  sld.addText('"SmartBonus is not a concept. It is the same propensity-driven incentive architecture I deployed at PhonePe for 350M users — rebuilt for PayPay at 70M and growing."', { x:0.58, y:6.72, w:8.84, h:0.8, fontSize:10, color:C.ink, fontFace:'Arial', italic:true });
}

// ─── SLIDE 8 — ROLLOUT ─────────────────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.dark };
  sld.addText('ROLLOUT PLAN', { x:0.5, y:0.3, w:9, h:0.28, fontSize:9, bold:true, color:C.red, charSpacing:4, fontFace:'Arial' });
  sld.addText('Phased Rollout — 6 Months to Full Deployment', { x:0.5, y:0.6, w:9, h:0.65, fontSize:24, bold:true, color:C.white, fontFace:'Arial' });
  const phases = [
    { n:'Phase 1', period:'Month 1-2: ML Pilot', color:C.red,
      items:['Train Propensity-to-Transact model on 90 days of transaction history','A/B test SmartBonus rate vs flat 3% blanket on 5% of user base','Measure incremental GMV and burn efficiency vs control','Instrument mission completion tracking and Step tier attribution'] },
    { n:'Phase 2', period:'Month 3-4: MissionBoard Launch', color:C.gold,
      items:['Ship MissionBoard with 5 weekly missions to all SmartBonus users','Tune mission categories based on pilot drop-off data','Launch ops nudge engine for 40% and 80% mission completion thresholds','Expand SmartBonus to 50% of user base with winning rate curve'] },
    { n:'Phase 3', period:'Month 5-6: Full Rollout', color:'22C55E',
      items:['Roll SmartBonus + MissionBoard to all 70M users','Integrate with PayPay Step tier system for seamless rate display','Enable merchant-funded mission bonuses for new category discovery','Launch weekly campaign report for leadership: burn, GMV, mission ROI'] },
  ];
  phases.forEach((ph,i)=>{
    const x = 0.38 + i*3.22;
    sld.addShape(prs.ShapeType.roundRect, { x, y:1.42, w:3.0, h:4.45, fill:{ color:C.hero }, line:{ color:ph.color, width:1.5 }, rectRadius:0.1, shadow:makeShadow() });
    sld.addShape(prs.ShapeType.rect, { x, y:1.42, w:3.0, h:0.44, fill:{ color:ph.color }, line:{ color:'000000', width:0 } });
    sld.addText(`${ph.n}  x  ${ph.period}`, { x:x+0.1, y:1.44, w:2.8, h:0.38, fontSize:9.5, bold:true, color:C.dark, fontFace:'Arial' });
    ph.items.forEach((item,j)=>{ sld.addText(`${String.fromCharCode(9658)}  ${item}`, { x:x+0.12, y:2.0+j*0.72, w:2.76, h:0.66, fontSize:8.5, color:C.lgray, fontFace:'Arial' }); });
  });
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:6.0, w:9.24, h:1.08, fill:{ color:C.hero }, line:{ color:C.gold, width:1.5 }, rectRadius:0.1 });
  sld.addText('What I need to build this:', { x:0.6, y:6.07, w:3.0, h:0.28, fontSize:10, bold:true, color:C.gold, fontFace:'Arial' });
  sld.addText('Access to transaction history for ML training  x  Data science team partnership  x  1 backend engineer for scoring API  x  A/B test infrastructure access  x  PayPay Step tier team alignment', { x:0.6, y:6.38, w:8.84, h:0.55, fontSize:9.5, color:C.lgray, fontFace:'Arial' });
  sld.addShape(prs.ShapeType.rect, { x:0, y:7.08, w:10, h:0.42, fill:{ color:C.hero }, line:{ color:'000000', width:0 } });
  sld.addText('Ajay Avaghade  x  avaghadeajay009@gmail.com  x  +91 9561558439  x  linkedin.com/in/ajay-avaghade', { x:0.3, y:7.12, w:9.4, h:0.28, fontSize:9, color:C.muted, fontFace:'Arial', align:'center' });
}

prs.writeFile({ fileName:'assets/paypay-smartbonus-deck.pptx' })
  .then(()=>console.log('paypay-smartbonus-deck.pptx done'));
