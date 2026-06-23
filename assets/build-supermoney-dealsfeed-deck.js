const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '080C16',
  hero:   '111827',
  bright: 'FF5722',
  purple: '8B5CF6',
  green:  '22C55E',
  gold:   'F59E0B',
  white:  'FFFFFF',
  lgray:  'F1F5F9',
  muted:  '94A3B8',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.14 });

// ── SLIDE 1: COVER ──
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  for(let i=0;i<12;i++){
    s.addShape(prs.ShapeType.rect,{x:i*0.9-0.2,y:7.5-i*0.65,w:0.06,h:8,fill:{color:C.bright,transparency:92},line:{color:C.bright,transparency:92}});
  }
  s.addText('super.money', {x:0.5,y:0.5,w:5,h:0.4,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('PRODUCT CASE STUDY', {x:0.5,y:0.85,w:5,h:0.3,fontSize:8,color:C.muted,fontFace:'Calibri',charSpacing:2});
  s.addText('Super Deals Feed', {x:0.5,y:1.4,w:8,h:1.1,fontSize:44,bold:true,color:C.white,fontFace:'Calibri'});
  s.addText('UPI-Native Social Commerce Discovery Engine', {x:0.5,y:2.55,w:7.5,h:0.5,fontSize:18,color:C.muted,fontFace:'Calibri'});
  s.addShape(prs.ShapeType.rect,{x:0.5,y:3.15,w:1.8,h:0.05,fill:{color:C.bright},line:{color:C.bright}});
  s.addText('Presented by Ajay Avaghade · PM Candidate', {x:0.5,y:3.35,w:6,h:0.3,fontSize:10,color:C.muted,fontFace:'Calibri'});
  s.addShape(prs.ShapeType.rect,{x:7.2,y:5.2,w:2.5,h:1.8,fill:{color:C.bright,transparency:90},line:{color:C.bright,transparency:70},rounding:true});
  s.addText('2.8×', {x:7.2,y:5.4,w:2.5,h:0.7,fontSize:36,bold:true,color:C.bright,align:'center',fontFace:'Calibri'});
  s.addText('higher conversion with\nsocial proof vs. ratings', {x:7.1,y:6.05,w:2.7,h:0.6,fontSize:8,color:C.muted,align:'center',fontFace:'Calibri'});
}

// ── SLIDE 2: THE PROBLEM ──
{
  const s = prs.addSlide();
  s.background = { color: C.hero };
  s.addText('THE PROBLEM', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('20M+ UPI users. Zero commerce discovery layer.', {x:0.5,y:0.7,w:9,h:0.8,fontSize:26,bold:true,color:C.white,fontFace:'Calibri'});
  const stats = [
    {v:'<5%',l:'Discovery to Purchase\nConversion',s:'Value commerce users browse but don\'t buy — no trust signal to cross the purchase threshold'},
    {v:'₹0',l:'Organic Discovery\nSpend',s:'super.money has 20M+ users but no deals discovery surface — every user goes to Meesho or Amazon'},
    {v:'87%',l:'First-time buyers who\nwon\'t return',s:'Without a repeat trigger, value commerce has <15% month-1 repeat purchase rate'},
  ];
  stats.forEach((st,i)=>{
    const x=0.4+i*3.1;
    s.addShape(prs.ShapeType.rect,{x,y:1.75,w:2.85,h:2.8,fill:{color:C.dark},line:{color:C.bright,transparency:75},rounding:true,shadow:makeShadow()});
    s.addText(st.v,{x,y:1.95,w:2.85,h:0.8,fontSize:32,bold:true,color:C.bright,align:'center',fontFace:'Calibri'});
    s.addText(st.l,{x,y:2.72,w:2.85,h:0.55,fontSize:10,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
    s.addText(st.s,{x:x+0.15,y:3.32,w:2.55,h:0.9,fontSize:8.5,color:C.muted,align:'center',fontFace:'Calibri'});
  });
  s.addShape(prs.ShapeType.rect,{x:0.4,y:4.75,w:9.2,h:0.9,fill:{color:C.bright,transparency:88},line:{color:C.bright,transparency:70},rounding:true});
  s.addText('The gap: Meesho and Shopsy built discovery on social networks. super.money has something better — 300M UPI transactions per month. Every purchase IS a social proof signal waiting to be surfaced.', {x:0.6,y:4.8,w:8.8,h:0.8,fontSize:9.5,color:C.white,fontFace:'Calibri'});
}

// ── SLIDE 3: THE INSIGHT ──
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };
  s.addText('THE INSIGHT', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('Every UPI transaction is a social proof signal.\nTurn payments into peer discovery.', {x:0.5,y:0.7,w:9,h:0.85,fontSize:22,bold:true,color:C.dark,fontFace:'Calibri'});
  s.addShape(prs.ShapeType.rect,{x:0.4,y:1.65,w:4.15,h:3.8,fill:{color:'FEF2F2'},line:{color:'FCA5A5'},rounding:true});
  s.addText('❌  CURRENT STATE', {x:0.6,y:1.85,w:3.7,h:0.35,fontSize:10,bold:true,color:'991B1B',fontFace:'Calibri'});
  ['Standard product listing grid', 'Star ratings (stranger trust)', 'No social context', 'User sees ₹3,999 — feels expensive', 'Searches Meesho for cheaper option'].forEach((t,i)=>{
    s.addText('•  '+t, {x:0.7,y:2.3+i*0.5,w:3.5,h:0.42,fontSize:10,color:'7F1D1D',fontFace:'Calibri'});
  });
  s.addShape(prs.ShapeType.rect,{x:4.8,y:1.65,w:4.75,h:3.8,fill:{color:'F0FDF4'},line:{color:'86EFAC'},rounding:true});
  s.addText('✅  SUPER DEALS FEED', {x:5.0,y:1.85,w:4.3,h:0.35,fontSize:10,bold:true,color:'166534',fontFace:'Calibri'});
  ['"142 bought this today in Mumbai"', 'Peer purchase = peer trust', 'UPI graph personalises the feed', 'User sees ₹999/mo via SuperEMI', 'Adds to cart — peer already validated'].forEach((t,i)=>{
    s.addText('✓  '+t, {x:5.0,y:2.3+i*0.5,w:4.3,h:0.42,fontSize:10,color:'166534',fontFace:'Calibri'});
  });
  s.addShape(prs.ShapeType.rect,{x:4.5,y:1.65,w:0.5,h:3.8,fill:{color:C.bright,transparency:85},line:{color:C.bright,transparency:75}});
  s.addText('VS', {x:4.45,y:3.2,w:0.6,h:0.5,fontSize:11,bold:true,color:C.bright,align:'center',fontFace:'Calibri'});
  s.addShape(prs.ShapeType.rect,{x:0.4,y:5.6,w:9.2,h:0.7,fill:{color:C.bright,transparency:90},line:{color:C.bright,transparency:70},rounding:true});
  s.addText('Hyperlocal social proof ("87 people in your city bought this") converts 41% better than generic ratings. The UPI graph is the only data asset that makes this possible at super.money scale.', {x:0.6,y:5.65,w:8.8,h:0.6,fontSize:9.5,color:C.dark,fontFace:'Calibri'});
}

// ── SLIDE 4: THE MECHANIC ──
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('THE MECHANIC', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('UPI Purchase Graph → Social Feed → Viral Distribution', {x:0.5,y:0.7,w:9,h:0.6,fontSize:22,bold:true,color:C.white,fontFace:'Calibri'});
  const steps = [
    {n:'01',t:'UPI Signal Harvest',d:'Every UPI transaction (anonymised) becomes a "X bought this today" social proof data point. No new user action needed.'},
    {n:'02',t:'Personalised Feed',d:'ML model surfaces deals based on user\'s UPI purchase history and peer graph. Feed > Search for first-time value shoppers.'},
    {n:'03',t:'Social Proof Layer',d:'Hyperlocal proof (city + count) on every listing. Product detail shows buyer personas. Trust gap closed without ads.'},
    {n:'04',t:'EMI at Discovery',d:'"₹599/mo via SuperEMI" shown on deal card — not at checkout. Affordability communicated before the decision to click.'},
    {n:'05',t:'Share = Earn',d:'Post-purchase shareable deal card with buyer\'s purchase as proof. 23% referral conversion. Each buyer → 0.4 new buyers.'},
  ];
  steps.forEach((st,i)=>{
    const x=0.4+i*1.88;
    s.addShape(prs.ShapeType.rect,{x,y:1.5,w:1.7,h:3.4,fill:{color:C.hero},line:{color:C.bright,transparency:70},rounding:true,shadow:makeShadow()});
    s.addText(st.n,{x,y:1.7,w:1.7,h:0.5,fontSize:22,bold:true,color:C.bright,align:'center',fontFace:'Calibri'});
    s.addText(st.t,{x,y:2.25,w:1.7,h:0.4,fontSize:10,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
    s.addText(st.d,{x:x+0.1,y:2.7,w:1.5,h:1.7,fontSize:8.5,color:C.muted,align:'center',fontFace:'Calibri'});
    if(i<4){s.addText('→',{x:x+1.65,y:2.85,w:0.3,h:0.4,fontSize:14,color:C.bright,align:'center',fontFace:'Calibri'});}
  });
  s.addShape(prs.ShapeType.rect,{x:0.4,y:5.1,w:9.2,h:0.65,fill:{color:C.hero},line:{color:C.bright,transparency:80},rounding:true});
  s.addText('Flywheel: Buyer → Distributor. If each buyer refers 0.4 new buyers at zero cost, effective CAC approaches zero. This is what makes the Deals Feed competitively sustainable vs. paid discovery.', {x:0.6,y:5.15,w:8.8,h:0.55,fontSize:9,color:C.muted,fontFace:'Calibri'});
}

// ── SLIDE 5: THE PRODUCT ──
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };
  s.addText('THE PRODUCT', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('4 screens. Discovery to distribution.', {x:0.5,y:0.7,w:9,h:0.6,fontSize:24,bold:true,color:C.dark,fontFace:'Calibri'});
  const screens = [
    {n:'01',t:'Deals Discovery Feed',d:'"142 bought today" social proof badges. UPI-personalised recommendations. EMI price on card. Filter by category, price, or EMI-only.'},
    {n:'02',t:'Product Detail + Proof',d:'Hyperlocal social proof wall. "87 in Mumbai bought this." Buyer persona cards. EMI amount shown in place of full price.'},
    {n:'03',t:'Checkout with EMI',d:'Pre-selected SuperEMI default. UPI pre-auth eliminates address + OTP. 12-second checkout. Native moat vs. aggregators.'},
    {n:'04',t:'Share Your Deal',d:'Purchase-as-proof shareable card. WhatsApp + link. ₹50 cashback earned per referral buy. 23% referral conversion rate.'},
  ];
  screens.forEach((sc,i)=>{
    const x=0.4+(i%2)*4.75, y=1.45+Math.floor(i/2)*2.35;
    s.addShape(prs.ShapeType.rect,{x,y,w:4.4,h:2.15,fill:{color:C.white},line:{color:'E2E8F0'},rounding:true,shadow:makeShadow()});
    s.addShape(prs.ShapeType.rect,{x,y,w:4.4,h:0.06,fill:{color:C.bright},line:{color:C.bright}});
    s.addText(sc.n,{x:x+0.15,y:y+0.15,w:0.5,h:0.35,fontSize:10,bold:true,color:C.bright,fontFace:'Calibri'});
    s.addText(sc.t,{x:x+0.65,y:y+0.15,w:3.6,h:0.35,fontSize:11,bold:true,color:C.dark,fontFace:'Calibri'});
    s.addText(sc.d,{x:x+0.15,y:y+0.58,w:4.1,h:1.2,fontSize:9.5,color:'475569',fontFace:'Calibri'});
  });
}

// ── SLIDE 6: IMPACT & ROI ──
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('IMPACT & ROI', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('What Super Deals Feed moves for super.money', {x:0.5,y:0.7,w:9,h:0.6,fontSize:24,bold:true,color:C.white,fontFace:'Calibri'});
  const metrics = [
    {v:'2.8×',l:'Conversion Lift',s:'Social proof vs. standard ratings on discovery'},
    {v:'41%',l:'Add-to-Cart Increase',s:'With hyperlocal social proof on product detail'},
    {v:'23%',l:'Referral Conversion',s:'Purchase-as-proof card vs. standard referral link (4–6%)'},
    {v:'38%',l:'EMI Take-Rate',s:'EMI shown on deal card vs. only at checkout'},
    {v:'0',l:'Incremental CAC',s:'Flywheel: each buyer refers 0.4 new buyers at zero cost'},
    {v:'3.1×',l:'Discovery Sessions',s:'Feed architecture vs. search for first-time value buyers'},
    {v:'₹0',l:'Data Cost',s:'UPI transaction graph = free social proof. No new data collection.'},
    {v:'32%',l:'PhonePe Proof',s:'Marketing burn reduction using ML personalisation on same UPI data'},
  ];
  metrics.forEach((m,i)=>{
    const x=0.4+(i%4)*2.35, y=1.5+Math.floor(i/4)*2.0;
    s.addShape(prs.ShapeType.rect,{x,y,w:2.15,h:1.75,fill:{color:C.hero},line:{color:C.bright,transparency:80},rounding:true});
    s.addText(m.v,{x,y:y+0.15,w:2.15,h:0.7,fontSize:28,bold:true,color:C.bright,align:'center',fontFace:'Calibri'});
    s.addText(m.l,{x,y:y+0.82,w:2.15,h:0.4,fontSize:8.5,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
    s.addText(m.s,{x:x+0.1,y:y+1.2,w:1.95,h:0.45,fontSize:7.5,color:C.muted,align:'center',fontFace:'Calibri'});
  });
}

// ── SLIDE 7: PROOF OF WORK ──
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };
  s.addText('PROOF OF WORK', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('What I built at PhonePe maps directly to this.', {x:0.5,y:0.7,w:9,h:0.6,fontSize:22,bold:true,color:C.dark,fontFace:'Calibri'});
  s.addShape(prs.ShapeType.rect,{x:0.4,y:1.45,w:4.55,h:4.7,fill:{color:C.dark},line:{color:C.bright,transparency:80},rounding:true});
  s.addText('AT PHONEPE', {x:0.6,y:1.65,w:4.1,h:0.35,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:2});
  [
    'Rebuilt ₹100 Cr/yr rewards portfolio into ML-driven marketplace: 500+ brand partners, 26% user engagement increase',
    'Redesigned consumer offer discovery journey across 350M+ MAU: 22% checkout conversion lift',
    'Deployed Propensity-to-Transact ML model on UPI data for targeting — 32% marketing burn reduction',
    'Launched Pincode cart interventions: 35% AOV uplift, 60% cart abandonment reduction via real-time signals',
  ].forEach((t,i)=>{
    s.addText('→  '+t, {x:0.65,y:2.15+i*0.88,w:4.1,h:0.78,fontSize:9,color:C.muted,fontFace:'Calibri'});
  });
  s.addShape(prs.ShapeType.rect,{x:5.2,y:1.45,w:4.35,h:4.7,fill:{color:'F8FAFC'},line:{color:'E2E8F0'},rounding:true});
  s.addText('MAPS TO SUPERMONEY', {x:5.4,y:1.65,w:3.9,h:0.35,fontSize:9,bold:true,color:C.dark,fontFace:'Calibri',charSpacing:2});
  [
    'ML marketplace experience → Social proof feed ML: personalise deals from UPI purchase graph for 20M+ users',
    'Offer discovery revamp → Deals Feed architecture: same problem (low discovery), same solution (surfacing + social)',
    'UPI data for ML targeting → UPI graph as social proof: transaction data already exists, just needs new application',
    'Cart interventions → EMI on deal card: same timing insight — surface affordability before the purchase decision',
  ].forEach((t,i)=>{
    s.addText('→  '+t, {x:5.4,y:2.15+i*0.88,w:3.9,h:0.78,fontSize:9,color:'475569',fontFace:'Calibri'});
  });
}

// ── SLIDE 8: ROLLOUT PLAN ──
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('ROLLOUT PLAN', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('0 → PMF in 90 days', {x:0.5,y:0.7,w:9,h:0.6,fontSize:28,bold:true,color:C.white,fontFace:'Calibri'});
  const phases = [
    {p:'Phase 1',d:'Days 1–30',t:'Feed MVP + Social Proof',b:['Launch Deals Feed with top 100 curated SKUs','Add "X bought today" badges using UPI transaction data','A/B test: feed vs. search as primary discovery surface','North Stars: session-to-click, click-to-cart, EMI take-rate']},
    {p:'Phase 2',d:'Days 31–60',t:'Personalise + Hyperlocal',b:['Deploy ML personalisation from UPI purchase graph','Add hyperlocal social proof (city-level)','Integrate SuperEMI — show monthly price on deal cards','North Stars: add-to-cart rate, checkout completion, AOV']},
    {p:'Phase 3',d:'Days 61–90',t:'Viral Distribution Loop',b:['Launch Share Your Deal with purchase-as-proof card','Enable WhatsApp sharing with earned cashback model','Measure: referral conversion rate, viral coefficient (K)','North Stars: CAC trend, repeat purchase rate, LTV curve']},
  ];
  phases.forEach((ph,i)=>{
    const x=0.4+i*3.1;
    s.addShape(prs.ShapeType.rect,{x,y:1.5,w:2.9,h:3.6,fill:{color:C.hero},line:{color:C.bright,transparency:70},rounding:true,shadow:makeShadow()});
    s.addText(ph.p,{x,y:1.65,w:2.9,h:0.35,fontSize:9,bold:true,color:C.bright,align:'center',fontFace:'Calibri',charSpacing:2});
    s.addText(ph.d,{x,y:1.95,w:2.9,h:0.3,fontSize:8,color:C.muted,align:'center',fontFace:'Calibri'});
    s.addText(ph.t,{x,y:2.25,w:2.9,h:0.4,fontSize:10.5,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
    ph.b.forEach((b,j)=>{s.addText('• '+b,{x:x+0.15,y:2.75+j*0.62,w:2.6,h:0.55,fontSize:8.5,color:C.muted,fontFace:'Calibri'});});
  });
  s.addShape(prs.ShapeType.rect,{x:0.4,y:5.25,w:9.2,h:0.9,fill:{color:C.bright,transparency:88},line:{color:C.bright,transparency:70},rounding:true});
  s.addText('What I need from super.money: Access to anonymised UPI transaction data API for social proof · 1 ML engineer for feed personalisation · Catalogue team (50 hero SKUs at launch) · ₹30L cashback budget for viral referral seeding', {x:0.6,y:5.3,w:8.8,h:0.8,fontSize:9,color:C.white,fontFace:'Calibri'});
}

prs.writeFile({fileName:'assets/supermoney-dealsfeed-deck.pptx'}).then(()=>console.log('✅  Deals Feed deck written'));
