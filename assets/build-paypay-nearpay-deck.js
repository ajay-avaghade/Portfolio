const pptx = require('pptxgenjs');
const prs = new pptx();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:  '1A0A0A',
  hero:  '2D1515',
  red:   'E60012',
  acc:   'FF3347',
  gold:  'FF6B00',
  teal:  '00B894',
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
  sld.addText('PAYPAY — DISCOVERY PLATFORM', { x:0.5, y:0.42, w:7, h:0.28, fontSize:9, bold:true, color:C.red, charSpacing:3, fontFace:'Arial' });
  sld.addText('PayPay NearPay', { x:0.5, y:0.85, w:9, h:1.05, fontSize:52, bold:true, color:C.white, fontFace:'Arial' });
  sld.addText('Contextual Merchant Discovery: Making Habit Beat Intent at Point-of-Sale', { x:0.5, y:2.0, w:7.2, h:0.75, fontSize:15, color:C.acc, fontFace:'Arial' });
  sld.addText('Presented by Ajay Avaghade · Growth & Platform PM', { x:0.5, y:2.9, w:6, h:0.32, fontSize:11, color:C.muted, fontFace:'Arial' });
  sld.addShape(prs.ShapeType.roundRect, { x:7.1, y:5.4, w:2.55, h:1.85, fill:{ color:C.hero }, line:{ color:C.red, width:2 }, rectRadius:0.12 });
  sld.addText('4M+', { x:7.1, y:5.45, w:2.55, h:0.72, fontSize:38, bold:true, color:C.red, fontFace:'Arial', align:'center' });
  sld.addText('PayPay merchant stores\nwith zero discovery surface', { x:7.1, y:6.2, w:2.55, h:0.75, fontSize:9, color:C.muted, fontFace:'Arial', align:'center' });
  sld.addShape(prs.ShapeType.rect, { x:0, y:7.08, w:10, h:0.42, fill:{ color:C.hero }, line:{ color:'000000', width:0 } });
  sld.addText('CASE STUDY  x  GROWTH & DISCOVERY  x  PAYPAY JAPAN', { x:0.3, y:7.12, w:9.4, h:0.28, fontSize:8, color:C.muted, charSpacing:2, fontFace:'Arial', align:'center' });
}

// ─── SLIDE 2 — THE PROBLEM ─────────────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.dark };
  sld.addText('THE PROBLEM', { x:0.5, y:0.3, w:9, h:0.28, fontSize:9, bold:true, color:C.red, charSpacing:4, fontFace:'Arial' });
  sld.addText('4 Million PayPay Merchants.\nZero Discovery Surface. Habit Wins Every Time.', { x:0.5, y:0.6, w:9, h:0.95, fontSize:27, bold:true, color:C.white, fontFace:'Arial' });
  const cols2 = [
    { icon:'🗺️', stat:'4M+', lbl:'PayPay merchants with no discovery', sub:'Users find them by accident or not at all — there is no in-app merchant map' },
    { icon:'🧠', stat:'Habit', lbl:'Wins at point-of-sale every time', sub:'Users reach for their wallet by default before they check if PayPay is accepted' },
    { icon:'📉', stat:'<40%', lbl:'Merchant activation within 90 days', sub:'Many accepting merchants have almost no PayPay volume because users do not know they accept it' },
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
  sld.addText('Root cause: The biggest competitor to PayPay at the point of sale is not another payment app. It is the physical wallet reflex. Users do not know the cafe they just walked into accepts PayPay. NearPay gives them a reason to check their phone first — before reaching for their card.', { x:0.58, y:5.6, w:8.84, h:1.1, fontSize:10.5, color:C.lgray, fontFace:'Arial', italic:true });
}

// ─── SLIDE 3 — THE INSIGHT ─────────────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.lgray };
  sld.addText('THE INSIGHT', { x:0.5, y:0.3, w:9, h:0.28, fontSize:9, bold:true, color:C.red, charSpacing:4, fontFace:'Arial' });
  sld.addText('The Discovery Bonus Costs PayPay Nothing.\nThe Merchant Pays. The User Discovers.', { x:0.5, y:0.6, w:9, h:0.95, fontSize:24, bold:true, color:C.ink, fontFace:'Arial' });
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:1.72, w:4.25, h:4.55, fill:{ color:C.white }, line:{ color:'DDDDDD', width:1 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('Without NearPay (Current)', { x:0.55, y:1.9, w:3.9, h:0.38, fontSize:12, bold:true, color:'CC3300', fontFace:'Arial' });
  const curr = ['User walks past 3 PayPay merchants without knowing','Reaches for wallet — pays by card or cash','PayPay gets zero transaction volume from that moment','Merchant paid PayPay onboarding fee but sees no uplift','Discovery happens only by accident, if at all'];
  curr.forEach((t,i)=>{ sld.addText(t, { x:0.58, y:2.42+i*0.6, w:3.9, h:0.54, fontSize:9.5, color:C.ink, fontFace:'Arial' }); });
  sld.addShape(prs.ShapeType.ellipse, { x:4.65, y:3.58, w:0.7, h:0.7, fill:{ color:C.red }, line:{ color:'000000', width:0 } });
  sld.addText('VS', { x:4.65, y:3.58, w:0.7, h:0.7, fontSize:11, bold:true, color:C.white, fontFace:'Arial', align:'center', valign:'middle' });
  sld.addShape(prs.ShapeType.roundRect, { x:5.37, y:1.72, w:4.25, h:4.55, fill:{ color:C.dark }, line:{ color:C.red, width:1.5 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('With NearPay (Proposed)', { x:5.52, y:1.9, w:3.9, h:0.38, fontSize:12, bold:true, color:C.red, fontFace:'Arial' });
  const prop = ['Push: 3 PayPay stores within 200m — Discovery Bonus available','User opens NearPay map, taps Ueno Cafe, walks 80m','Pays with PayPay — 5% First Visit Bonus applied automatically','Merchant funded the bonus from their campaign budget','Discovery becomes repeatable — 8 new stores found this month'];
  prop.forEach((t,i)=>{ sld.addText(t, { x:5.52, y:2.42+i*0.6, w:3.9, h:0.54, fontSize:9.5, color:C.lgray, fontFace:'Arial' }); });
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:6.42, w:9.24, h:0.85, fill:{ color:C.white }, line:{ color:C.red, width:1.5 }, rectRadius:0.08 });
  sld.addText('Key insight: The Discovery Bonus is fully merchant-funded. PayPay acts as a matchmaker — it surfaces the merchant to the user at the right moment and takes zero net subsidy cost. The merchant gets a new customer; the user earns a bonus; PayPay gets the transaction.', { x:0.58, y:6.52, w:8.84, h:0.68, fontSize:10, color:C.ink, fontFace:'Arial', italic:true });
}

// ─── SLIDE 4 — THE MECHANIC ────────────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.dark };
  sld.addText('THE MECHANIC', { x:0.5, y:0.3, w:9, h:0.28, fontSize:9, bold:true, color:C.red, charSpacing:4, fontFace:'Arial' });
  sld.addText('Context-Aware Trigger: Right Merchant, Right Moment, Right Bonus', { x:0.5, y:0.6, w:9, h:0.65, fontSize:24, bold:true, color:C.white, fontFace:'Arial' });
  const steps = [
    { n:'01', title:'Detect Context', desc:'GPS detects user is stationary or slow-moving within 200m of an undiscovered PayPay merchant. Time window matches meal, coffee, or commute patterns.' },
    { n:'02', title:'Score Relevance', desc:'Category affinity model checks if this merchant type matches the users payment history. Cafe at 9am for frequent coffee buyers. Ramen at noon for regular restaurant users.' },
    { n:'03', title:'Trigger + Push', desc:'Push notification fires: 3 PayPay stores near you. Bonus available. Opens NearPay map in-app. Merchant funded Discovery Bonus shown per store.' },
    { n:'04', title:'Pay + Discover', desc:'User scans QR at new merchant. First Visit Bonus applied automatically from merchant campaign budget. No manual redemption. Zero friction at checkout.' },
    { n:'05', title:'Merchant Loop', desc:'Merchant sees first-time customer in their PayPay for Business dashboard. Campaign ROI tracked in real time. Budget auto-renews when conversion rate exceeds threshold.' },
  ];
  steps.forEach((st,i)=>{
    const x = 0.38 + i*1.9;
    sld.addShape(prs.ShapeType.roundRect, { x, y:1.65, w:1.72, h:4.75, fill:{ color:C.hero }, line:{ color: i===2 ? C.gold : '333333', width: i===2 ? 2 : 1 }, rectRadius:0.1, shadow:makeShadow() });
    sld.addShape(prs.ShapeType.ellipse, { x:x+0.56, y:1.78, w:0.6, h:0.6, fill:{ color: i===2 ? C.gold : C.red }, line:{ color:'000000', width:0 } });
    sld.addText(st.n, { x:x+0.56, y:1.78, w:0.6, h:0.6, fontSize:11, bold:true, color:C.white, fontFace:'Arial', align:'center', valign:'middle' });
    sld.addText(st.title, { x:x+0.08, y:2.48, w:1.56, h:0.5, fontSize:10, bold:true, color: i===2 ? C.gold : C.white, fontFace:'Arial', align:'center' });
    sld.addText(st.desc, { x:x+0.1, y:3.05, w:1.52, h:3.3, fontSize:8, color:C.muted, fontFace:'Arial' });
  });
  sld.addText('Proof at PhonePe: Context-aware cart incentivisation engine (cart value x user intent x time signals) delivered 35% AOV uplift and 60% cart abandonment reduction. NearPay applies the same trigger logic at the physical point of sale.', { x:0.38, y:6.6, w:9.24, h:0.55, fontSize:9.5, color:C.muted, fontFace:'Arial', italic:true });
}

// ─── SLIDE 5 — THE PRODUCT ─────────────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.lgray };
  sld.addText('THE PRODUCT', { x:0.5, y:0.3, w:9, h:0.28, fontSize:9, bold:true, color:C.red, charSpacing:4, fontFace:'Arial' });
  sld.addText('4 Key Screen States — Live, Interactive Prototype Built', { x:0.5, y:0.6, w:9, h:0.65, fontSize:24, bold:true, color:C.ink, fontFace:'Arial' });
  const screens = [
    { n:'01', title:'Home + Push', color:C.red, desc:'Push notification surfaces 3 PayPay stores within 200m. NearPay home card shows all three with distance and first-visit bonus. Taps through to the map.', ascii:'[Notification]\nPayPay\n3 stores near you\nDiscovery bonus!\n\n[Home — NearPay Card]\n● NearPay · 3 stores\n200m · lunch bonus on\n\n☕ Ueno Cafe   80m +5%\n🍣 Sushiro    160m +3%\n🏪 FamilyMart 190m +1%\n\nView full map →' },
    { n:'02', title:'Map + List', color:C.red, desc:'In-app merchant map shows all nearby PayPay stores as pins. Filter chips by category. Merchant list below sorted by distance. New stores flagged with first-visit bonus badge.', ascii:'[NearPay Map]\n\n  [Map: dark bg]\n  ☕ pin +5% new\n  🍣 pin +3% new\n  🏪 pin +1%\n  (O) You are here\n\n[All] [New] [Food][Cafe]\n\nUeno Cafe    80m  NEW\n+5% First Visit Bonus\nSushiro Exp  160m NEW\n+3% First Visit Bonus\nFamilyMart   190m\n+1% Regular' },
    { n:'03', title:'Pay + Discovery', color:'00B894', desc:'Payment success + confetti. Discovery bonus applied automatically. New store added to personal PayPay Map. Discovery achievement shown. Next nearby store surfaced immediately.', ascii:'[Payment Complete!]\n        ✓\n   Ueno Cafe ¥820\n\n[NearPay Discovery]\nNew store found!\n+¥41 bonus earned\n8 stores discovered\nExplorer badge!\n\n[Ueno Cafe saved to\n  your PayPay Map]\n\n[Find More Stores →]' },
    { n:'04', title:'Merchant Ops', color:C.gold, desc:'Internal merchant operations view: NearPay reach KPIs, discovery funnel (push sent to first payment), top performing merchant campaigns by conversion rate.', ascii:'[NearPay Ops]\n\n+27% Txn freq\n34%  Merchants active\n¥840 ARPU lift/mo\n\nDiscovery Funnel\nPush sent  ████ 100%\nPush open  ██░░  48%\nMap view   █░░░  34%\nMerch tap  █░░░  21%\nPayment    ░░░░  14%\n\nTop Campaigns\nIchiran Ramen 44%\nUeno Cafe     38%\nSushiro       29%' },
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
  sld.addText('Prototype: paypay-nearpay-prototype.html  x  Live interactive demo of all 4 states', { x:0.5, y:7.1, w:9, h:0.28, fontSize:8.5, color:C.muted, fontFace:'Arial', italic:true, align:'center' });
}

// ─── SLIDE 6 — IMPACT & ROI ────────────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.dark };
  sld.addText('IMPACT & ROI', { x:0.5, y:0.3, w:9, h:0.28, fontSize:9, bold:true, color:C.red, charSpacing:4, fontFace:'Arial' });
  sld.addText('Projected Impact — Zero Net Cost to PayPay', { x:0.5, y:0.6, w:9, h:0.65, fontSize:24, bold:true, color:C.white, fontFace:'Arial' });
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:1.4, w:4.5, h:5.38, fill:{ color:C.hero }, line:{ color:'333333', width:1 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('User Impact', { x:0.55, y:1.55, w:4.15, h:0.38, fontSize:12, bold:true, color:C.red, fontFace:'Arial' });
  const um = [
    { val:'+27%', lbl:'Transaction frequency', sub:'NearPay users discover new contexts — more daily payment moments' },
    { val:'+¥840', lbl:'Monthly ARPU lift', sub:'Discovery of new merchant categories drives additional spend' },
    { val:'8', lbl:'New stores per explorer/month', sub:'Users who discover 3+ stores have 2.8x higher 30-day retention' },
    { val:'48%', lbl:'Push open rate', sub:'Context-relevant push at the right location and time beats generic push by 3x' },
  ];
  um.forEach((m,i)=>{
    sld.addShape(prs.ShapeType.roundRect, { x:0.55, y:2.08+i*1.1, w:4.15, h:0.96, fill:{ color:C.dark }, line:{ color:'333333', width:1 }, rectRadius:0.07 });
    sld.addText(m.val, { x:0.7, y:2.12+i*1.1, w:1.3, h:0.52, fontSize:24, bold:true, color:C.red, fontFace:'Arial', align:'center', valign:'middle' });
    sld.addText(m.lbl, { x:2.08, y:2.15+i*1.1, w:2.45, h:0.28, fontSize:10, bold:true, color:C.white, fontFace:'Arial' });
    sld.addText(m.sub, { x:2.08, y:2.44+i*1.1, w:2.45, h:0.28, fontSize:8.5, color:C.muted, fontFace:'Arial' });
  });
  sld.addShape(prs.ShapeType.roundRect, { x:5.12, y:1.4, w:4.5, h:5.38, fill:{ color:C.hero }, line:{ color:'333333', width:1 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('PayPay + Merchant ROI', { x:5.28, y:1.55, w:4.15, h:0.38, fontSize:12, bold:true, color:C.gold, fontFace:'Arial' });
  const pm = [
    { val:'0', lbl:'Net cost of Discovery Bonus to PayPay', sub:'100% merchant-funded from their PayPay for Business campaign budget' },
    { val:'35%', lbl:'Merchant campaign opt-in rate', sub:'14% conversion from push to payment — 3x better than generic email campaigns' },
    { val:'+15%', lbl:'Merchant category discovery rate', sub:'Users explore new store types they have never paid at with PayPay before' },
    { val:'2.8x', lbl:'Retention lift for multi-discoverers', sub:'Users who find 3+ new stores in month 1 are 2.8x more likely to stay active' },
  ];
  pm.forEach((m,i)=>{
    sld.addShape(prs.ShapeType.roundRect, { x:5.28, y:2.08+i*1.1, w:4.15, h:0.96, fill:{ color:C.dark }, line:{ color:'333333', width:1 }, rectRadius:0.07 });
    sld.addText(m.val, { x:5.4, y:2.12+i*1.1, w:1.3, h:0.52, fontSize:24, bold:true, color:C.gold, fontFace:'Arial', align:'center', valign:'middle' });
    sld.addText(m.lbl, { x:6.76, y:2.15+i*1.1, w:2.45, h:0.28, fontSize:10, bold:true, color:C.white, fontFace:'Arial' });
    sld.addText(m.sub, { x:6.76, y:2.44+i*1.1, w:2.45, h:0.28, fontSize:8.5, color:C.muted, fontFace:'Arial' });
  });
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:6.95, w:9.24, h:0.38, fill:{ color:C.hero }, line:{ color:'333333', width:1 }, rectRadius:0.07 });
  sld.addText('The Discovery Bonus is the rarest product mechanic: one where the user wins, the merchant wins, and PayPay wins — without spending a single extra yen on campaign budget.', { x:0.58, y:6.98, w:8.84, h:0.3, fontSize:9, color:C.muted, fontFace:'Arial', italic:true });
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
    'Built dynamic cart incentivisation engine with context-aware triggers: cart value x user intent x time signals — 35% AOV uplift and 60% cart abandonment reduction at Pincode (quick commerce)',
    'Deployed ML audience segmentation for brand self-serve campaign tooling in rewards marketplace — 500+ brand partners onboarded, 26% uplift in user engagement across merchant-linked cohorts',
    'Built multi-tenant referral engine and self-serve PG integration platform collapsing 2-day manual merchant onboarding into 30-minute automated flow — 5,000+ B2B merchants acquired',
    'Led ML-driven personalised offers redesign across 350M MAU — 22% checkout conversion lift through cross-functional delivery with Tech, UX, Finance, and Legal',
    'MissionBoard architecture (weekly missions, tier unlock, nudge engine) shipped for merchant device retention at Virtusa and trading activation at MultiBank',
  ];
  pp.forEach((t,i)=>{ sld.addText(`${String.fromCharCode(9658)}  ${t}`, { x:0.55, y:2.12+i*0.82, w:4.06, h:0.75, fontSize:9, color:C.lgray, fontFace:'Arial' }); });
  sld.addShape(prs.ShapeType.roundRect, { x:5.22, y:1.42, w:4.4, h:5.08, fill:{ color:C.white }, line:{ color:'DDDDDD', width:1 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('How It Maps to This PayPay Role', { x:5.38, y:1.58, w:4.06, h:0.38, fontSize:12, bold:true, color:C.ink, fontFace:'Arial' });
  const mp = [
    { from:'Context-aware trigger engine (cart x intent x time)', to:'NearPay GPS trigger (location x category affinity x time window)' },
    { from:'ML audience segmentation for merchant campaigns', to:'Category affinity model for merchant-user matching' },
    { from:'Brand self-serve campaign tooling, 500+ partners', to:'Merchant-funded Discovery Bonus via PayPay for Business' },
    { from:'B2B merchant onboarding (2 days to 30 minutes)', to:'NearPay merchant campaign self-serve onboarding' },
    { from:'Milestone architecture shipped across 3 companies', to:'Discovery achievement system and Explorer badge layer' },
  ];
  mp.forEach((m,i)=>{
    sld.addText(`PhonePe: ${m.from}`, { x:5.38, y:2.12+i*0.82, w:4.06, h:0.3, fontSize:8.5, color:'888888', fontFace:'Arial' });
    sld.addText(`${String.fromCharCode(8594)}  PayPay: ${m.to}`, { x:5.38, y:2.42+i*0.82, w:4.06, h:0.3, fontSize:9, bold:true, color:C.ink, fontFace:'Arial' });
  });
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:6.65, w:9.24, h:0.95, fill:{ color:'FFEBEB' }, line:{ color:C.red, width:1.5 }, rectRadius:0.08 });
  sld.addText('"NearPay is the context-aware trigger engine I built at PhonePe — applied to physical merchant discovery. The mechanic is the same. The canvas is every street in Japan."', { x:0.58, y:6.72, w:8.84, h:0.8, fontSize:10, color:C.ink, fontFace:'Arial', italic:true });
}

// ─── SLIDE 8 — ROLLOUT ─────────────────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.dark };
  sld.addText('ROLLOUT PLAN', { x:0.5, y:0.3, w:9, h:0.28, fontSize:9, bold:true, color:C.red, charSpacing:4, fontFace:'Arial' });
  sld.addText('Phased Rollout — 6 Months to Full Deployment', { x:0.5, y:0.6, w:9, h:0.65, fontSize:24, bold:true, color:C.white, fontFace:'Arial' });
  const phases = [
    { n:'Phase 1', period:'Month 1-2: Pilot', color:C.red,
      items:['Launch NearPay in 1 district (Shibuya or Shinjuku) with 50 opted-in merchants','Instrument push open rate, map view, merchant tap, and payment conversion funnel','A/B test push message copy: proximity vs bonus vs discovery framing','Validate category affinity model: does it surface the right merchant type per user?'] },
    { n:'Phase 2', period:'Month 3-4: Merchant Platform', color:C.gold,
      items:['Launch merchant self-serve campaign console in PayPay for Business','Onboard 1,000+ merchants to NearPay Discovery Bonus programme across 3 cities','Tune trigger radius and time window based on pilot conversion data','Add Discovery Points and Explorer badge system to deepen engagement loop'] },
    { n:'Phase 3', period:'Month 5-6: National Rollout', color:'22C55E',
      items:['Roll NearPay to all 70M PayPay users and all 4M+ accepting merchants nationally','Integrate with PayPay Step: NearPay discoveries credit Step tier progress','Launch NearPay merchant leaderboard — top-performing campaigns get editorial placement','Enable multi-merchant discovery routes: NearPay suggests a 3-stop lunch run nearby'] },
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
  sld.addText('Merchant location and category data API  x  GPS trigger permissions framework  x  PayPay for Business team alignment  x  Push notification infrastructure  x  1 engineer for proximity scoring service', { x:0.6, y:6.38, w:8.84, h:0.55, fontSize:9.5, color:C.lgray, fontFace:'Arial' });
  sld.addShape(prs.ShapeType.rect, { x:0, y:7.08, w:10, h:0.42, fill:{ color:C.hero }, line:{ color:'000000', width:0 } });
  sld.addText('Ajay Avaghade  x  avaghadeajay009@gmail.com  x  +91 9561558439  x  linkedin.com/in/ajay-avaghade', { x:0.3, y:7.12, w:9.4, h:0.28, fontSize:9, color:C.muted, fontFace:'Arial', align:'center' });
}

prs.writeFile({ fileName:'assets/paypay-nearpay-deck.pptx' })
  .then(()=>console.log('paypay-nearpay-deck.pptx done'));
