const pptx = require('pptxgenjs');

const prs = new pptx();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '00302A',
  hero:   '00453C',
  teal:   '00BFA5',
  gold:   'D4A843',
  white:  'FFFFFF',
  lgray:  'F0F7F5',
  ink:    '0A1F1C',
  muted:  '8B9E9B',
};

const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });

// ─── SLIDE 1 — COVER (dark) ───────────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.dark };

  // Diagonal line motif
  for(let i=0;i<6;i++){
    sld.addShape(prs.ShapeType.line, { x:i*1.8-0.5, y:0, w:0, h:7.5, line:{ color:C.hero, width:18 } });
  }

  // Company tag
  sld.addText('ADCB — CORPORATE CARDS PRODUCT', {
    x:0.5, y:0.45, w:7, h:0.28,
    fontSize:9, bold:true, color:C.gold, charSpacing:3,
    fontFace:'Arial', align:'left',
  });

  // Product name
  sld.addText('ADCB CardLaunch', {
    x:0.5, y:0.9, w:8.5, h:1.1,
    fontSize:52, bold:true, color:C.white,
    fontFace:'Arial', align:'left',
  });

  // Subtitle
  sld.addText('From 3-Week Onboarding to 24 Minutes:\nA Self-Serve Corporate Card Programme Launch Engine', {
    x:0.5, y:2.05, w:7, h:0.85,
    fontSize:16, color:C.teal, bold:false,
    fontFace:'Arial', align:'left',
  });

  // Presenter
  sld.addText('Presented by Ajay Avaghade · Growth & Monetization PM', {
    x:0.5, y:3.05, w:6, h:0.32,
    fontSize:11, color:C.muted,
    fontFace:'Arial', align:'left',
  });

  // Big stat bottom right
  sld.addShape(prs.ShapeType.roundRect, { x:7.1, y:5.5, w:2.5, h:1.7, fill:{ color:C.hero }, line:{ color:C.gold, width:2 }, rectRadius:0.12 });
  sld.addText('3 WKS', { x:7.1, y:5.55, w:2.5, h:0.65, fontSize:36, bold:true, color:C.gold, fontFace:'Arial', align:'center' });
  sld.addText('avg onboarding time\n(current state)', { x:7.1, y:6.22, w:2.5, h:0.7, fontSize:9, color:C.muted, fontFace:'Arial', align:'center' });

  // Arrow
  sld.addText('↓', { x:7.1, y:5.2, w:2.5, h:0.35, fontSize:20, bold:true, color:C.teal, fontFace:'Arial', align:'center' });

  // Bottom bar
  sld.addShape(prs.ShapeType.rect, { x:0, y:7.1, w:10, h:0.4, fill:{ color:C.hero }, line:{ color:'000000', width:0 } });
  sld.addText('CASE STUDY  ·  CORPORATE BANKING PRODUCT  ·  ABU DHABI 2025', {
    x:0.3, y:7.12, w:9.4, h:0.28,
    fontSize:8, color:C.muted, charSpacing:2,
    fontFace:'Arial', align:'center',
  });
}

// ─── SLIDE 2 — THE PROBLEM (dark) ─────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.dark };

  sld.addText('THE PROBLEM', {
    x:0.5, y:0.3, w:9, h:0.28,
    fontSize:9, bold:true, color:C.gold, charSpacing:4, fontFace:'Arial',
  });

  sld.addText('Corporate Card Onboarding Takes 3 Weeks.\nIt Does Not Have To.', {
    x:0.5, y:0.62, w:9, h:1.0,
    fontSize:28, bold:true, color:C.white, fontFace:'Arial',
  });

  const cols = [
    { icon:'🕐', stat:'3 Weeks', lbl:'Average onboarding time', sub:'Manual KYB, paper forms, relationship manager back-and-forth' },
    { icon:'📋', stat:'14 Steps', lbl:'Manual onboarding steps', sub:'Trade licence, KYB, credit assessment, card issuance — all disconnected' },
    { icon:'💳', stat:'40–60%', lbl:'Drop-off before first swipe', sub:'Long TAT kills momentum — finance teams lose interest before cards arrive' },
  ];

  cols.forEach((col,i)=>{
    const x = 0.38 + i*3.18;
    sld.addShape(prs.ShapeType.roundRect, { x, y:1.82, w:2.9, h:3.5, fill:{ color:C.hero }, line:{ color:C.teal+'44', width:1 }, rectRadius:0.12, shadow:makeShadow() });
    sld.addText(col.icon, { x, y:2.0, w:2.9, h:0.55, fontSize:26, fontFace:'Arial', align:'center' });
    sld.addText(col.stat, { x, y:2.6, w:2.9, h:0.72, fontSize:34, bold:true, color:C.gold, fontFace:'Arial', align:'center' });
    sld.addText(col.lbl, { x, y:3.35, w:2.9, h:0.4, fontSize:11, bold:true, color:C.white, fontFace:'Arial', align:'center' });
    sld.addText(col.sub, { x:x+0.15, y:3.78, w:2.6, h:1.3, fontSize:9, color:C.muted, fontFace:'Arial', align:'center' });
  });

  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:5.52, w:9.24, h:1.4, fill:{ color:C.hero }, line:{ color:C.gold+'55', width:1 }, rectRadius:0.1 });
  sld.addText('Root cause: Corporate card onboarding is a relationship-manager-led manual process — no digital self-serve, no real-time KYB, no instant card issuance. Every delay is a lost activation.', {
    x:0.6, y:5.62, w:8.8, h:1.1,
    fontSize:10.5, color:C.lgray, fontFace:'Arial', italic:true,
  });
}

// ─── SLIDE 3 — THE INSIGHT (light) ────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.lgray };

  sld.addText('THE INSIGHT', {
    x:0.5, y:0.3, w:9, h:0.28,
    fontSize:9, bold:true, color:C.teal, charSpacing:4, fontFace:'Arial',
  });

  sld.addText('Corporate Onboarding Should Feel Like\nSigning Up for a Consumer App.', {
    x:0.5, y:0.6, w:9, h:0.95,
    fontSize:26, bold:true, color:C.ink, fontFace:'Arial',
  });

  // Left column — current
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:1.75, w:4.2, h:4.5, fill:{ color:C.white }, line:{ color:'DDDDDD', width:1 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('❌  Current State', { x:0.55, y:1.92, w:3.86, h:0.38, fontSize:13, bold:true, color:'CC3300', fontFace:'Arial' });
  const currItems = [
    'RM-led: client must wait for a human at every step',
    'Paper KYC/KYB submitted manually, verified days later',
    'Credit assessment is a black box — no visibility',
    'Cards take 5–7 business days to arrive physically',
    'Finance admin has no dashboard — calls RM for updates',
  ];
  currItems.forEach((txt,i)=>{
    sld.addText(`${i+1}. ${txt}`, { x:0.58, y:2.42+i*0.58, w:3.84, h:0.52, fontSize:10, color:C.ink, fontFace:'Arial' });
  });

  // VS divider
  sld.addShape(prs.ShapeType.ellipse, { x:4.65, y:3.5, w:0.7, h:0.7, fill:{ color:C.teal }, line:{ color:'000000', width:0 } });
  sld.addText('VS', { x:4.65, y:3.5, w:0.7, h:0.7, fontSize:11, bold:true, color:C.white, fontFace:'Arial', align:'center', valign:'middle' });

  // Right column — proposed
  sld.addShape(prs.ShapeType.roundRect, { x:5.42, y:1.75, w:4.2, h:4.5, fill:{ color:C.dark }, line:{ color:C.teal+'66', width:1.5 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('✅  CardLaunch', { x:5.58, y:1.92, w:3.86, h:0.38, fontSize:13, bold:true, color:C.teal, fontFace:'Arial' });
  const newItems = [
    'Finance admin self-serves: 4-step wizard, no RM needed',
    'Digital KYB in real time — trade licence OCR + eKYC',
    'AI credit scoring: decision in under 2 minutes',
    'Virtual cards issued instantly; physical cards in parallel',
    'Live dashboard: spend, limits, compliance — all in one place',
  ];
  newItems.forEach((txt,i)=>{
    sld.addText(`${i+1}. ${txt}`, { x:5.58, y:2.42+i*0.58, w:3.86, h:0.52, fontSize:10, color:C.lgray, fontFace:'Arial' });
  });

  // Callout
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:6.42, w:9.24, h:0.85, fill:{ color:C.teal+'22' }, line:{ color:C.teal, width:1.5 }, rectRadius:0.08 });
  sld.addText('Key insight: When a finance admin can onboard their entire team in 24 minutes — without calling the RM — activation rates compound. Every hour saved before first swipe is an hour closer to revenue.', {
    x:0.58, y:6.52, w:8.84, h:0.68,
    fontSize:10, color:C.ink, fontFace:'Arial', italic:true,
  });
}

// ─── SLIDE 4 — THE MECHANIC (dark) ────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.dark };

  sld.addText('THE MECHANIC', {
    x:0.5, y:0.3, w:9, h:0.28,
    fontSize:9, bold:true, color:C.gold, charSpacing:4, fontFace:'Arial',
  });

  sld.addText('4-Step Self-Serve Wizard:\nFrom Trade Licence to 25 Virtual Cards in 24 Minutes', {
    x:0.5, y:0.6, w:9, h:0.95,
    fontSize:26, bold:true, color:C.white, fontFace:'Arial',
  });

  const steps = [
    { n:'01', title:'Company Setup', desc:'Finance admin enters trade licence number. OCR auto-fills company details. Visa Commercial or Mastercard Corporate selected. Estimated credit limit surfaced instantly.' },
    { n:'02', title:'KYB & Credit', desc:'eKYC documents uploaded directly (no physical submission). AI credit scoring completes in under 2 minutes. Programme limit confirmed. Admin receives instant approval email.' },
    { n:'03', title:'Employee Upload', desc:'CSV upload or manual entry of employees. Role-based limits assigned per employee. Virtual cards issued in real time — available in Apple/Google Wallet within 60 seconds.' },
    { n:'04', title:'Go Live', desc:'Finance admin receives live CorporateLens dashboard access. Spend tracking, limit controls, and MCC restrictions active from Day 1. RM notified automatically — no call needed.' },
    { n:'05', title:'ADCB Ops', desc:'Compliance and risk team sees new programme in real-time queue. KYB flags auto-escalated. Card issuance and scheme reporting auto-triggered. Zero manual data entry required.' },
  ];

  steps.forEach((st,i)=>{
    const x = 0.38 + i*1.9;
    sld.addShape(prs.ShapeType.roundRect, { x, y:1.75, w:1.72, h:4.6, fill:{ color:C.hero }, line:{ color: i===2 ? C.gold : C.teal+'44', width: i===2 ? 2 : 1 }, rectRadius:0.1, shadow:makeShadow() });

    sld.addShape(prs.ShapeType.ellipse, { x:x+0.56, y:1.88, w:0.6, h:0.6, fill:{ color: i===2 ? C.gold : C.teal }, line:{ color:'000000', width:0 } });
    sld.addText(st.n, { x:x+0.56, y:1.88, w:0.6, h:0.6, fontSize:11, bold:true, color:C.dark, fontFace:'Arial', align:'center', valign:'middle' });

    sld.addText(st.title, { x:x+0.08, y:2.58, w:1.56, h:0.55, fontSize:10, bold:true, color: i===2 ? C.gold : C.white, fontFace:'Arial', align:'center' });
    sld.addText(st.desc, { x:x+0.1, y:3.18, w:1.52, h:3.1, fontSize:8, color:C.muted, fontFace:'Arial', align:'left' });

    // Connector dot between steps
    if(i<4){
      sld.addShape(prs.ShapeType.rect, { x:x+1.72, y:2.16, w:0.18, h:0.04, fill:{ color:C.teal+'88' }, line:{ color:'000000', width:0 } });
    }
  });

  sld.addText('Proof: B2B merchant onboarding at PhonePe reduced from 2 days to 30 minutes using the same self-serve principle — same architecture, new context.', {
    x:0.38, y:6.55, w:9.24, h:0.6,
    fontSize:9.5, color:C.muted, fontFace:'Arial', italic:true,
  });
}

// ─── SLIDE 5 — THE PRODUCT (light) ────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.lgray };

  sld.addText('THE PRODUCT', {
    x:0.5, y:0.3, w:9, h:0.28,
    fontSize:9, bold:true, color:C.teal, charSpacing:4, fontFace:'Arial',
  });

  sld.addText('4 Key Screen States — Built & Validated', {
    x:0.5, y:0.6, w:9, h:0.65,
    fontSize:24, bold:true, color:C.ink, fontFace:'Arial',
  });

  const screens = [
    {
      n:'01', title:'Company Setup', color:C.teal, desc:'Finance admin enters trade licence number — OCR auto-fills company name, registration number, and address. Credit limit options surfaced instantly based on revenue data.',
      ascii:'[ADCB CardLaunch Wizard]\n\nStep 1 of 4: Company Setup\n\n🏢 Al Futtaim Group LLC\n   Trade Lic: DED-2024-44821\n   ✅ Auto-verified\n\n💳 Programme Limit\n   ○ AED 250K\n   ● AED 500K ← selected\n   ○ AED 1M+\n\n[Next: Employee Setup →]',
    },
    {
      n:'02', title:'Employee Upload', color:C.gold, desc:'CSV bulk upload or manual entry. Role-based spending limits assigned per employee. Virtual cards issued instantly — physical cards dispatched in parallel within 24 hours.',
      ascii:'[Step 3 of 4: Employees]\n\n25 total  18 uploaded  7 virtual ✅\n\n📤 Drop CSV or add manually\n\nAhmad Al-Rashidi  AED 25K  ✅ Virtual\nFatima Al-Hashimi AED 15K  ✅ Virtual\nMohammed Al-Zaabi AED 20K  ✅ Virtual\nLaila Hassan      AED 10K  🚚 Physical\nKhalid Al-Nuaimi  AED 30K  ✅ Virtual\n\n[Issue Cards to All 25 →]',
    },
    {
      n:'03', title:'Programme Live!', color:'22C55E', desc:'Instant confirmation: all virtual cards active, employee notifications sent, CorporateLens dashboard access granted. RM notified automatically — zero call required.',
      ascii:'[🎉 Programme Live!]\n\n25 cards issued\n18 virtual · active now\nAED 500K limit live\n\n⏱ 24 minutes total\n   (was 3 weeks)\n\n📧 All employees notified\n   via email + SMS\n\n[→ Go to CorporateLens]',
    },
    {
      n:'04', title:'ADCB Ops View', color:C.teal, desc:'Internal ops dashboard shows all new programmes in real time. KYB flags auto-escalated for review. Card issuance and scheme reporting auto-triggered — zero manual data entry.',
      ascii:'[ADCB Ops — Onboarding]\n\nNew Programmes MTD: 14\nAvg Onboarding: 24 min\nCards Issued: 287\n\nKYB Queue\n✅ Al Futtaim    — Live\n✅ Emaar Prop    — Live\n⏳ DEWA          — KYB\n✅ DP World      — Live\n\n[View Full Pipeline →]',
    },
  ];

  screens.forEach((sc,i)=>{
    const x = 0.3 + i*2.42;
    sld.addShape(prs.ShapeType.roundRect, { x, y:1.42, w:2.22, h:5.65, fill:{ color:C.white }, line:{ color:sc.color+'44', width:1 }, rectRadius:0.1, shadow:makeShadow() });
    sld.addShape(prs.ShapeType.rect, { x, y:1.42, w:2.22, h:0.38, fill:{ color:sc.color }, line:{ color:'000000', width:0 }, rectRadius:0 });
    sld.addText(`${sc.n}  ${sc.title}`, { x:x+0.08, y:1.44, w:2.06, h:0.32, fontSize:9, bold:true, color:C.white, fontFace:'Arial' });

    sld.addShape(prs.ShapeType.roundRect, { x:x+0.1, y:1.88, w:2.02, h:2.3, fill:{ color:C.dark }, line:{ color:C.hero, width:1 }, rectRadius:0.06 });
    sld.addText(sc.ascii, { x:x+0.14, y:1.92, w:1.94, h:2.22, fontSize:5.8, color:C.lgray, fontFace:'Courier New', valign:'top' });

    sld.addText(sc.desc, { x:x+0.1, y:4.28, w:2.02, h:2.6, fontSize:8.5, color:C.ink, fontFace:'Arial', valign:'top' });
  });

  sld.addText('Prototype available: adcb-cardlaunch-prototype.html  ·  Interactive demo of all 4 states', {
    x:0.5, y:7.1, w:9, h:0.28,
    fontSize:8.5, color:C.muted, fontFace:'Arial', italic:true, align:'center',
  });
}

// ─── SLIDE 6 — IMPACT & ROI (dark) ────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.dark };

  sld.addText('IMPACT & ROI', {
    x:0.5, y:0.3, w:9, h:0.28,
    fontSize:9, bold:true, color:C.gold, charSpacing:4, fontFace:'Arial',
  });

  sld.addText('Projected Impact — Built on PhonePe B2B Onboarding Proof Points', {
    x:0.5, y:0.6, w:9, h:0.65,
    fontSize:22, bold:true, color:C.white, fontFace:'Arial',
  });

  // Left panel — Corporate client impact
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:1.4, w:4.5, h:5.35, fill:{ color:C.hero }, line:{ color:C.teal+'44', width:1 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('🏢  Corporate Client Impact', { x:0.55, y:1.55, w:4.15, h:0.38, fontSize:12, bold:true, color:C.teal, fontFace:'Arial' });

  const clientMetrics = [
    { val:'24 min', lbl:'Onboarding time', sub:'Down from 3 weeks (97% reduction)' },
    { val:'↓ 60%', lbl:'Finance admin effort', sub:'Zero RM calls needed for setup' },
    { val:'Day 1', lbl:'Cards active', sub:'Virtual cards live before first meeting ends' },
    { val:'1 screen', lbl:'Full programme visibility', sub:'Spend + limits + compliance in CorporateLens' },
  ];

  clientMetrics.forEach((m,i)=>{
    sld.addShape(prs.ShapeType.roundRect, { x:0.55, y:2.08+i*1.1, w:4.15, h:0.96, fill:{ color:C.dark }, line:{ color:C.teal+'33', width:1 }, rectRadius:0.07 });
    sld.addText(m.val, { x:0.7, y:2.12+i*1.1, w:1.3, h:0.52, fontSize:24, bold:true, color:C.teal, fontFace:'Arial', align:'center', valign:'middle' });
    sld.addText(m.lbl, { x:2.08, y:2.15+i*1.1, w:2.45, h:0.28, fontSize:10, bold:true, color:C.white, fontFace:'Arial' });
    sld.addText(m.sub, { x:2.08, y:2.44+i*1.1, w:2.45, h:0.28, fontSize:8.5, color:C.muted, fontFace:'Arial' });
  });

  // Right panel — ADCB ROI
  sld.addShape(prs.ShapeType.roundRect, { x:5.12, y:1.4, w:4.5, h:5.35, fill:{ color:C.hero }, line:{ color:C.gold+'44', width:1 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('📈  ADCB Bank ROI', { x:5.28, y:1.55, w:4.15, h:0.38, fontSize:12, bold:true, color:C.gold, fontFace:'Arial' });

  const bankMetrics = [
    { val:'↑ 40%', lbl:'First-swipe activation rate', sub:'Faster onboarding means less drop-off' },
    { val:'↑ 3×', lbl:'Programmes per RM per month', sub:'RM freed from admin; focuses on sales' },
    { val:'↓ 50%', lbl:'Operational cost per programme', sub:'No manual KYB, no paper, no callbacks' },
    { val:'↑ 25%', lbl:'Revenue per programme (LTV)', sub:'Higher activation = more spend volume' },
  ];

  bankMetrics.forEach((m,i)=>{
    sld.addShape(prs.ShapeType.roundRect, { x:5.28, y:2.08+i*1.1, w:4.15, h:0.96, fill:{ color:C.dark }, line:{ color:C.gold+'33', width:1 }, rectRadius:0.07 });
    sld.addText(m.val, { x:5.4, y:2.12+i*1.1, w:1.3, h:0.52, fontSize:24, bold:true, color:C.gold, fontFace:'Arial', align:'center', valign:'middle' });
    sld.addText(m.lbl, { x:6.75, y:2.15+i*1.1, w:2.45, h:0.28, fontSize:10, bold:true, color:C.white, fontFace:'Arial' });
    sld.addText(m.sub, { x:6.75, y:2.44+i*1.1, w:2.45, h:0.28, fontSize:8.5, color:C.muted, fontFace:'Arial' });
  });

  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:6.92, w:9.24, h:0.4, fill:{ color:C.hero }, line:{ color:C.gold+'55', width:1 }, rectRadius:0.07 });
  sld.addText('Every RM hour saved on admin is an hour redirected to new corporate client acquisition — the same compounding effect I drove at PhonePe.', {
    x:0.58, y:6.96, w:8.84, h:0.3,
    fontSize:9, color:C.muted, fontFace:'Arial', italic:true,
  });
}

// ─── SLIDE 7 — PROOF OF WORK (light) ──────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.lgray };

  sld.addText('PROOF OF WORK', {
    x:0.5, y:0.3, w:9, h:0.28,
    fontSize:9, bold:true, color:C.teal, charSpacing:4, fontFace:'Arial',
  });

  sld.addText('I Built This. Here is the Proof.', {
    x:0.5, y:0.6, w:9, h:0.65,
    fontSize:24, bold:true, color:C.ink, fontFace:'Arial',
  });

  // Left — dark panel PhonePe
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:1.42, w:4.4, h:5.05, fill:{ color:C.dark }, line:{ color:C.teal+'44', width:1 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('What I Built at PhonePe', { x:0.55, y:1.58, w:4.06, h:0.38, fontSize:12, bold:true, color:C.teal, fontFace:'Arial' });

  const phonepe = [
    'Built zero-to-one multi-tenant referral engine + self-serve PG integration platform replacing a 2-day manual merchant onboarding process with a 30-minute automated self-serve flow',
    'Acquired 5,000+ B2B merchants and 5Mn+ new users/month with 23% lower CAC versus the prior manual acquisition funnel',
    'Deployed Propensity-to-Transact ML models cutting marketing burn by 32% while sustaining GMV growth',
    'Redesigned offers eligibility and redemption UX across 350M MAU driving a 22% checkout conversion lift through A/B experimentation',
    'Led founding team at Kotak Cherry: scoped MVP, owned GTM, drove 100K+ downloads in the launch window',
  ];

  phonepe.forEach((txt,i)=>{
    sld.addText(`▸  ${txt}`, { x:0.55, y:2.1+i*0.82, w:4.06, h:0.75, fontSize:9, color:C.lgray, fontFace:'Arial' });
  });

  // Right — light panel JD mapping
  sld.addShape(prs.ShapeType.roundRect, { x:5.22, y:1.42, w:4.4, h:5.05, fill:{ color:C.white }, line:{ color:'DDDDDD', width:1 }, rectRadius:0.1, shadow:makeShadow() });
  sld.addText('How It Maps to This ADCB Role', { x:5.38, y:1.58, w:4.06, h:0.38, fontSize:12, bold:true, color:C.ink, fontFace:'Arial' });

  const mapping = [
    { from:'2-day → 30-min B2B onboarding at PhonePe', to:'3-week → 24-min corporate card launch (CardLaunch)' },
    { from:'5,000+ merchant self-serve acquisition', to:'Corporate client self-serve onboarding without RM' },
    { from:'ML targeting and audience segmentation', to:'AI credit scoring and real-time KYB automation' },
    { from:'Offers eligibility + redemption UX redesign', to:'CorporateLens control plane and spend dashboard' },
    { from:'Kotak Cherry 0→1 WealthTech launch', to:'ADCB CardLaunch 0→1 self-serve programme build' },
  ];

  mapping.forEach((m,i)=>{
    sld.addText(`PhonePe: ${m.from}`, { x:5.38, y:2.1+i*0.82, w:4.06, h:0.3, fontSize:8.5, color:'888888', fontFace:'Arial' });
    sld.addText(`→  ADCB: ${m.to}`, { x:5.38, y:2.4+i*0.82, w:4.06, h:0.3, fontSize:9, bold:true, color:C.ink, fontFace:'Arial' });
  });

  // Quote box
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:6.65, w:9.24, h:0.95, fill:{ color:C.teal+'1A' }, line:{ color:C.teal, width:1.5 }, rectRadius:0.08 });
  sld.addText('"CardLaunch is not a concept. It is the same self-serve onboarding architecture I shipped at PhonePe for B2B merchants — now applied to corporate card programme launch at a bank."', {
    x:0.58, y:6.72, w:8.84, h:0.8,
    fontSize:10, color:C.ink, fontFace:'Arial', italic:true,
  });
}

// ─── SLIDE 8 — ROLLOUT PLAN (dark) ────────────────────────────────────────
{
  const sld = prs.addSlide();
  sld.background = { color: C.dark };

  sld.addText('ROLLOUT PLAN', {
    x:0.5, y:0.3, w:9, h:0.28,
    fontSize:9, bold:true, color:C.gold, charSpacing:4, fontFace:'Arial',
  });

  sld.addText('Phased Rollout — 6 Months to Full Deployment', {
    x:0.5, y:0.6, w:9, h:0.65,
    fontSize:24, bold:true, color:C.white, fontFace:'Arial',
  });

  const phases = [
    {
      n:'Phase 1', period:'Month 1–2: Pilot',
      color: C.teal,
      items:['Deploy CardLaunch wizard to 10 new SME corporate clients', 'Instrument all onboarding steps — measure TAT, drop-off, and first-swipe rate vs control', 'Run wizard vs RM-assisted A/B test — isolate friction points', 'Align with Compliance and KYB team on eKYC document requirements'],
    },
    {
      n:'Phase 2', period:'Month 3–4: Scale',
      color: C.gold,
      items:['Roll out to all new corporate card applications with RM as optional fallback', 'Launch CorporateLens dashboard integration — spend controls go live from Day 1', 'Implement AI credit scoring API; reduce credit decision TAT from days to minutes', 'Tune wizard flow based on pilot drop-off data — expected 40%+ activation improvement'],
    },
    {
      n:'Phase 3', period:'Month 5–6: Full Rollout',
      color:'22C55E',
      items:['Migrate existing RM-onboarded clients to self-serve dashboard (CorporateLens)', 'Automate Visa/Mastercard scheme reporting via CardLaunch data pipeline', 'Expand to new segments: government entities, large enterprise (AED 1M+ programmes)', 'Launch RM productivity dashboard — track portfolio growth freed from admin tasks'],
    },
  ];

  phases.forEach((ph,i)=>{
    const x = 0.38 + i*3.22;
    sld.addShape(prs.ShapeType.roundRect, { x, y:1.42, w:3.0, h:4.4, fill:{ color:C.hero }, line:{ color:ph.color+'55', width:1.5 }, rectRadius:0.1, shadow:makeShadow() });
    sld.addShape(prs.ShapeType.rect, { x, y:1.42, w:3.0, h:0.44, fill:{ color:ph.color }, line:{ color:'000000', width:0 } });
    sld.addText(`${ph.n}  ·  ${ph.period}`, { x:x+0.1, y:1.44, w:2.8, h:0.38, fontSize:9.5, bold:true, color:C.dark, fontFace:'Arial' });
    ph.items.forEach((item,j)=>{
      sld.addText(`▸  ${item}`, { x:x+0.12, y:2.0+j*0.72, w:2.76, h:0.66, fontSize:8.5, color:C.lgray, fontFace:'Arial' });
    });
  });

  // What I need
  sld.addShape(prs.ShapeType.roundRect, { x:0.38, y:6.0, w:9.24, h:1.05, fill:{ color:C.hero }, line:{ color:C.gold+'66', width:1.5 }, rectRadius:0.1 });
  sld.addText('What I need to build this:', { x:0.6, y:6.07, w:2.8, h:0.28, fontSize:10, bold:true, color:C.gold, fontFace:'Arial' });
  sld.addText('Access to KYB/KYC data flows  ·  Alignment with Credit Risk and Cards Ops team  ·  1 backend engineer for wizard API  ·  1 designer for onboarding UX  ·  Visa/Mastercard sandbox access for scheme testing', {
    x:0.6, y:6.38, w:8.84, h:0.55,
    fontSize:9.5, color:C.lgray, fontFace:'Arial',
  });

  // Footer
  sld.addShape(prs.ShapeType.rect, { x:0, y:7.1, w:10, h:0.4, fill:{ color:C.hero }, line:{ color:'000000', width:0 } });
  sld.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade', {
    x:0.3, y:7.14, w:9.4, h:0.28,
    fontSize:9, color:C.muted, fontFace:'Arial', align:'center',
  });
}

// ─── WRITE FILE ────────────────────────────────────────────────────────────
prs.writeFile({ fileName:'assets/adcb-cardlaunch-deck.pptx' })
  .then(()=>console.log('adcb-cardlaunch-deck.pptx done'));
