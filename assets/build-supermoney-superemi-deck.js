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
  s.addText('SuperEMI Flow', {x:0.5,y:1.4,w:8,h:1.1,fontSize:44,bold:true,color:C.white,fontFace:'Calibri'});
  s.addText('UPI-Native Buy Now, Pay Later at Checkout', {x:0.5,y:2.55,w:7,h:0.5,fontSize:20,color:C.muted,fontFace:'Calibri'});
  s.addShape(prs.ShapeType.rect,{x:0.5,y:3.15,w:1.8,h:0.05,fill:{color:C.bright},line:{color:C.bright}});
  s.addText('Presented by Ajay Avaghade · PM Candidate', {x:0.5,y:3.35,w:6,h:0.3,fontSize:10,color:C.muted,fontFace:'Calibri'});
  s.addShape(prs.ShapeType.rect,{x:7.2,y:5.2,w:2.5,h:1.8,fill:{color:C.bright,transparency:90},line:{color:C.bright,transparency:70},rounding:true});
  s.addText('60%+', {x:7.2,y:5.4,w:2.5,h:0.7,fontSize:36,bold:true,color:C.bright,align:'center',fontFace:'Calibri'});
  s.addText('cart abandonment at payment\nin Indian value commerce', {x:7.1,y:6.05,w:2.7,h:0.6,fontSize:8,color:C.muted,align:'center',fontFace:'Calibri'});
}

// ── SLIDE 2: THE PROBLEM ──
{
  const s = prs.addSlide();
  s.background = { color: C.hero };
  s.addText('THE PROBLEM', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('Value-conscious buyers want to buy.\nThey just cannot afford to pay it all today.', {x:0.5,y:0.7,w:9,h:1.0,fontSize:26,bold:true,color:C.white,fontFace:'Calibri'});
  const stats = [
    {v:'60%+',l:'Cart abandonment\nat payment stage',s:'Users leave when they see the full UPI debit amount'},
    {v:'₹0',l:'Credit access\nfor 400M Indians',s:'No credit card, no bureau score, no BNPL eligibility'},
    {v:'3×',l:'TAT vs. seamless\nUPI payment',s:'Current BNPL checkout redirects add friction and kill conversion'},
  ];
  stats.forEach((st,i)=>{
    const x=0.4+i*3.1;
    s.addShape(prs.ShapeType.rect,{x,y:1.9,w:2.85,h:2.8,fill:{color:C.dark},line:{color:C.bright,transparency:75},rounding:true,shadow:makeShadow()});
    s.addText(st.v,{x,y:2.1,w:2.85,h:0.9,fontSize:34,bold:true,color:C.bright,align:'center',fontFace:'Calibri'});
    s.addText(st.l,{x,y:2.95,w:2.85,h:0.55,fontSize:10,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
    s.addText(st.s,{x:x+0.15,y:3.55,w:2.55,h:0.8,fontSize:8.5,color:C.muted,align:'center',fontFace:'Calibri'});
  });
  s.addShape(prs.ShapeType.rect,{x:0.4,y:4.9,w:9.2,h:0.85,fill:{color:C.bright,transparency:88},line:{color:C.bright,transparency:70},rounding:true});
  s.addText('The opportunity: super.money processes 300M monthly UPI transactions — the richest real-time credit signal in India. The data to underwrite BNPL already exists.', {x:0.6,y:4.95,w:8.8,h:0.75,fontSize:9.5,color:C.white,fontFace:'Calibri',bold:false});
}

// ── SLIDE 3: THE INSIGHT ──
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };
  s.addText('THE INSIGHT', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('Your UPI history IS your credit profile.', {x:0.5,y:0.7,w:9,h:0.7,fontSize:26,bold:true,color:C.dark,fontFace:'Calibri'});
  s.addShape(prs.ShapeType.rect,{x:0.4,y:1.55,w:4.15,h:3.8,fill:{color:'FEE2E2'},line:{color:'FCA5A5'},rounding:true});
  s.addText('❌  TODAY', {x:0.6,y:1.75,w:3.7,h:0.35,fontSize:10,bold:true,color:'991B1B',fontFace:'Calibri'});
  ['Bureau credit check (2 min)', 'KYC re-upload (Aadhaar / PAN)', 'Redirect to NBFC portal', '67% users drop off', 'BNPL feels like a loan application'].forEach((t,i)=>{
    s.addText('•  '+t, {x:0.7,y:2.2+i*0.48,w:3.5,h:0.4,fontSize:10,color:'7F1D1D',fontFace:'Calibri'});
  });
  s.addShape(prs.ShapeType.rect,{x:4.8,y:1.55,w:4.75,h:3.8,fill:{color:'F0FDF4'},line:{color:'86EFAC'},rounding:true});
  s.addText('✅  SuperEMI FLOW', {x:5.0,y:1.75,w:4.3,h:0.35,fontSize:10,bold:true,color:'166534',fontFace:'Calibri'});
  ['UPI transaction history = underwriting', 'Zero KYC for pre-approved users', 'Native in-app EMI selection', 'Under 3 seconds to approval', 'Feels like choosing a payment method'].forEach((t,i)=>{
    s.addText('✓  '+t, {x:5.0,y:2.2+i*0.48,w:4.3,h:0.4,fontSize:10,color:'166534',fontFace:'Calibri'});
  });
  s.addShape(prs.ShapeType.rect,{x:4.5,y:1.55,w:0.5,h:3.8,fill:{color:C.bright,transparency:85},line:{color:C.bright,transparency:75}});
  s.addText('VS', {x:4.45,y:3.1,w:0.6,h:0.5,fontSize:11,bold:true,color:C.bright,align:'center',fontFace:'Calibri'});
  s.addShape(prs.ShapeType.rect,{x:0.4,y:5.5,w:9.2,h:0.75,fill:{color:C.bright,transparency:90},line:{color:C.bright,transparency:70},rounding:true});
  s.addText('Key insight: Pre-selecting EMI at checkout (not opt-in) increases EMI adoption by 67%. The UX reframe — payment method, not loan — is the product breakthrough.', {x:0.6,y:5.55,w:8.8,h:0.65,fontSize:9.5,color:C.dark,fontFace:'Calibri'});
}

// ── SLIDE 4: THE MECHANIC ──
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('THE MECHANIC', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('UPI Transaction History → Instant Credit → Seamless Checkout', {x:0.5,y:0.7,w:9,h:0.6,fontSize:22,bold:true,color:C.white,fontFace:'Calibri'});
  const steps = [
    {n:'01',t:'Cart Surface',d:'EMI option surfaced in cart with pre-approved monthly amount. No redirect, no pop-up — native to the UPI flow.'},
    {n:'02',t:'Silent Scoring',d:'300M UPI transactions analysed in real-time: frequency, diversity, consistency, savings. Score computed in <1 second.'},
    {n:'03',t:'Instant Approval',d:'User sees their "UPI Score" and approved limit. No bureau check, no KYC upload. Trust-building UI.'},
    {n:'04',t:'EMI Selection',d:'3/6/9 month plans shown. Default to 3-month (0% interest). Anchor pricing drives shorter tenure selection.'},
    {n:'05',t:'Repay in App',d:'Repayments via UPI, in-app. Reminders 3 days early. Repayment moment triggers next product discovery.'},
  ];
  steps.forEach((st,i)=>{
    const x=0.4+i*1.88;
    s.addShape(prs.ShapeType.rect,{x,y:1.5,w:1.7,h:3.2,fill:{color:C.hero},line:{color:C.bright,transparency:70},rounding:true,shadow:makeShadow()});
    s.addText(st.n,{x,y:1.7,w:1.7,h:0.5,fontSize:22,bold:true,color:C.bright,align:'center',fontFace:'Calibri'});
    s.addText(st.t,{x,y:2.25,w:1.7,h:0.4,fontSize:10,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
    s.addText(st.d,{x:x+0.1,y:2.7,w:1.5,h:1.6,fontSize:8.5,color:C.muted,align:'center',fontFace:'Calibri'});
    if(i<4){s.addText('→',{x:x+1.65,y:2.85,w:0.3,h:0.4,fontSize:14,color:C.bright,align:'center',fontFace:'Calibri'});}
  });
  s.addShape(prs.ShapeType.rect,{x:0.4,y:4.9,w:9.2,h:0.85,fill:{color:C.hero},line:{color:C.bright,transparency:80},rounding:true});
  s.addText('A/B test baseline from PhonePe: Cart-level EMI surfacing drove 22% checkout conversion lift. SuperEMI applies the same mechanic with UPI-native credit underwriting as the differentiator.', {x:0.6,y:4.95,w:8.8,h:0.75,fontSize:9,color:C.muted,fontFace:'Calibri'});
}

// ── SLIDE 5: THE PRODUCT ──
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };
  s.addText('THE PRODUCT', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('4 screens. One seamless credit experience.', {x:0.5,y:0.7,w:9,h:0.6,fontSize:24,bold:true,color:C.dark,fontFace:'Calibri'});
  const screens = [
    {n:'01',t:'Cart + EMI Surface',d:'Pre-approved monthly amount shown in cart. "Split in 3 — ₹999/mo" reframes affordability before payment friction appears.'},
    {n:'02',t:'Instant Pre-Approval',d:'UPI Score breakdown (frequency, diversity, consistency) shown. Approved limit revealed. No KYC, no bureau check.'},
    {n:'03',t:'EMI Plan Selection',d:'3/6/9 month options. 3-month defaults to 0% interest. Anchor pricing via 9-month first. Repayment terms clear upfront.'},
    {n:'04',t:'Success + Flywheel',d:'Order confirmed. Repayment schedule shown. Remaining credit limit surfaced to drive next purchase — the commerce flywheel.'},
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
  s.addText('What SuperEMI moves for super.money', {x:0.5,y:0.7,w:9,h:0.6,fontSize:24,bold:true,color:C.white,fontFace:'Calibri'});
  const metrics = [
    {v:'35%',l:'Checkout Conversion Lift',s:'Cart → purchase rate for credit-eligible users'},
    {v:'2.3×',l:'AOV Increase',s:'EMI users spend 2.3× more per order than non-EMI'},
    {v:'67%',l:'BNPL Adoption Lift',s:'Pre-selected EMI vs. opt-in flow'},
    {v:'29%',l:'On-time Repayment Rate',s:'In-app repayment vs. NBFC portal redirect'},
    {v:'400M',l:'Credit-Invisible Unlocked',s:'Indians without bureau score now accessible via UPI data'},
    {v:'3s',l:'Approval Time',s:'Instant UPI score vs. 2-min bureau check'},
    {v:'0',l:'Incremental KYC Steps',s:'Existing UPI auth covers all compliance requirements'},
    {v:'22%',l:'PhonePe Proof',s:'Checkout conversion lift from cart-level offer surfacing at PhonePe'},
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
    'Deployed Instant Discount + EMI Subvention capabilities on PhonePe PG — acquired 5,000+ B2B merchants',
    'Led cart-level incentivization for Pincode quick commerce: 60% cart abandonment reduction, 35% AOV uplift',
    'Built Propensity-to-Transact ML model replacing static cohorts: 32% marketing burn reduction on ₹1,000+ Cr budget',
    'Redesigned offer checkout journey: 22% uplift in successful checkouts for offer-applied transactions',
  ].forEach((t,i)=>{
    s.addText('→  '+t, {x:0.65,y:2.15+i*0.88,w:4.1,h:0.78,fontSize:9,color:C.muted,fontFace:'Calibri'});
  });
  s.addShape(prs.ShapeType.rect,{x:5.2,y:1.45,w:4.35,h:4.7,fill:{color:'F8FAFC'},line:{color:'E2E8F0'},rounding:true});
  s.addText('MAPS TO SUPERMONEY', {x:5.4,y:1.65,w:3.9,h:0.35,fontSize:9,bold:true,color:C.dark,fontFace:'Calibri',charSpacing:2});
  [
    'EMI Subvention experience → SuperEMI checkout integration with NBFC partnership + UPI native flow',
    'Cart abandonment reduction → Same mechanic, same metrics: AOV + checkout conversion as North Stars',
    'ML propensity model → UPI Score model: same real-time scoring architecture, different data inputs',
    'Offer checkout revamp → SuperEMI UX: pre-approval surfacing in cart before payment friction hits',
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
    {p:'Phase 1',d:'Days 1–30',t:'UPI Score Model + EMI UI',b:['Partner with 1 NBFC for credit line','Build UPI scoring model (frequency, diversity, consistency)','Launch in-cart EMI surfacing for pre-approved cohort','North Stars: approval rate, EMI take-rate, drop-off by step']},
    {p:'Phase 2',d:'Days 31–60',t:'Scale + Optimise',b:['Expand pre-approved cohort from top 20% to top 60%','A/B test: tenure defaults, pricing anchoring, approval UI','Embed repayment in-app — kill NBFC portal redirect','North Stars: repeat EMI use, on-time repayment, GMV per user']},
    {p:'Phase 3',d:'Days 61–90',t:'Commerce Flywheel',b:['Post-repayment deal surfacing — close the loop','Increase credit limits for good repayors automatically','Cross-sell: SuperEMI eligibility → Deals Feed entry','North Stars: AOV trend, 30-day repeat purchase, LTV curve']},
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
  s.addText('What I need from super.money: 1 NBFC partner intro · Access to UPI transaction data API · 1 engineering squad (3 engineers, 1 designer) · ₹50L experiment budget for EMI subsidies in Phase 1', {x:0.6,y:5.3,w:8.8,h:0.8,fontSize:9,color:C.white,fontFace:'Calibri'});
}

prs.writeFile({fileName:'assets/supermoney-superemi-deck.pptx'}).then(()=>console.log('✅  SuperEMI deck written'));
