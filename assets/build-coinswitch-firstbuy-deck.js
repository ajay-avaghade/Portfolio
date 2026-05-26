const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9'; // 10" x 5.625"

const C = {
  dark:   '0A0E27', hero: '111432',
  bright: 'F7B731', green: '22C55E',
  red:    'EF4444', blue:  '3B82F6',
  purple: '7C3AED', white: 'FFFFFF',
  lgray:  'F4F6FF', muted: '8B9CB6',
};
const sh = () => ({ type:'outer', color:'000000', blur:4, offset:1, angle:45, opacity:0.14 });

// SLIDE 1 — COVER
{
  const s = prs.addSlide(); s.background = { color: C.dark };
  for(let i=0;i<5;i++){
    const lx=-0.5+i*2.2, lw=Math.min(3.5,10-lx-0.15);
    if(lx>10||lw<=0) continue;
    s.addShape(prs.ShapeType.line,{x:lx,y:0,w:lw,h:5.625,line:{color:C.bright,width:0.4},rotate:30,transparency:88});
  }
  s.addShape(prs.ShapeType.roundRect,{x:0.5,y:0.26,w:3.4,h:0.26,rectRadius:0.05,fill:{color:C.bright,transparency:88},line:{color:C.bright,width:0.5,transparency:50}});
  s.addText('PEEPALCO · COINSWITCH GROWTH PM',{x:0.5,y:0.26,w:3.4,h:0.26,fontSize:7,bold:true,color:C.bright,align:'center',letterSpacingPt:1});
  s.addText('CoinSwitch FirstBuy',{x:0.5,y:0.72,w:7,h:1.0,fontSize:46,bold:true,color:C.white,fontFace:'Arial Black'});
  s.addShape(prs.ShapeType.rect,{x:0.5,y:1.72,w:2.0,h:0.05,fill:{color:C.bright}});
  s.addText('Closing the KYC→First Investment Gap Through AI-Guided Activation',{x:0.5,y:1.88,w:7.2,h:0.38,fontSize:13,color:'AABBD0',italic:true});
  s.addText('Presented by Ajay Avaghade  ·  Growth PM',{x:0.5,y:2.36,w:7,h:0.26,fontSize:9.5,color:C.muted});
  s.addShape(prs.ShapeType.roundRect,{x:7.0,y:3.8,w:2.75,h:1.55,rectRadius:0.1,fill:{color:C.bright,transparency:86},line:{color:C.bright,width:1,transparency:50},shadow:sh()});
  s.addText('~50%',{x:7.0,y:3.88,w:2.75,h:0.72,fontSize:40,bold:true,color:C.bright,align:'center',fontFace:'Arial Black'});
  s.addText('of KYC-verified users\nnever make a first trade',{x:7.0,y:4.60,w:2.75,h:0.55,fontSize:8,color:C.muted,align:'center'});
}

// SLIDE 2 — PROBLEM
{
  const s = prs.addSlide(); s.background = { color: C.dark };
  s.addText('THE PROBLEM',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('25M Registered Users. Half Have Never Invested a Single Rupee.',{x:0.5,y:0.50,w:9,h:0.65,fontSize:21,bold:true,color:C.white});
  const cols=[
    {icon:'😶',stat:'~50%',lbl:'Of KYC-verified users\nnever place a trade',sub:'Completes verification\nthen disappears silently'},
    {icon:'🤯',stat:'300+',lbl:'Coins on home screen\nat first login',sub:'"What do I even buy?"\nparalysis drives drop-off'},
    {icon:'💸',stat:'5×',lbl:'Cost to win back\nvs fix activation now',sub:'Reacquiring a lapsed user\ncosts 5× fixing the funnel'},
  ];
  cols.forEach((c,i)=>{
    const x=0.5+i*3.1;
    s.addShape(prs.ShapeType.roundRect,{x,y:1.15,w:2.9,h:2.8,rectRadius:0.12,fill:{color:C.hero},line:{color:C.bright,width:0.3,transparency:70},shadow:sh()});
    s.addText(c.icon,{x,y:1.25,w:2.9,h:0.45,fontSize:22,align:'center'});
    s.addText(c.stat,{x,y:1.72,w:2.9,h:0.72,fontSize:36,bold:true,color:C.bright,align:'center',fontFace:'Arial Black'});
    s.addText(c.lbl,{x,y:2.48,w:2.9,h:0.58,fontSize:9.5,color:C.white,align:'center',bold:true});
    s.addText(c.sub,{x,y:3.10,w:2.9,h:0.60,fontSize:8.5,color:C.muted,align:'center'});
  });
  s.addShape(prs.ShapeType.roundRect,{x:0.5,y:4.22,w:9,h:0.80,rectRadius:0.08,fill:{color:C.bright,transparency:88},line:{color:C.bright,width:0.6,transparency:55}});
  s.addText('Root cause: CoinSwitch delivers a verified, motivated user to a 300-coin catalogue with no guidance. The best moment to convert — immediately post-KYC — is wasted on overwhelming choice.',{x:0.65,y:4.28,w:8.7,h:0.68,fontSize:8.5,color:C.white,italic:true});
}

// SLIDE 3 — INSIGHT
{
  const s = prs.addSlide(); s.background = { color: C.lgray };
  s.addText('THE INSIGHT',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('A 3-Question Quiz + Social Proof Converts More Than a 300-Coin Catalogue.',{x:0.5,y:0.50,w:9,h:0.65,fontSize:20,bold:true,color:C.dark});
  s.addShape(prs.ShapeType.roundRect,{x:0.5,y:1.1,w:4.0,h:3.0,rectRadius:0.1,fill:{color:'FFF8E6'},line:{color:C.bright,width:1,transparency:50},shadow:sh()});
  s.addText('❌  Current: Empty Catalogue Drop',{x:0.65,y:1.22,w:3.7,h:0.30,fontSize:11,bold:true,color:'B87800'});
  ['Post-KYC → home screen with 300+ coins','No guidance on what to buy or why','Decision fatigue → user exits','Activation email 3 days later → ignored','User churns silently; ops sees it too late'].forEach((pt,i)=>{
    s.addText(`• ${pt}`,{x:0.65,y:1.60+i*0.42,w:3.7,h:0.38,fontSize:9.5,color:C.dark});
  });
  s.addShape(prs.ShapeType.ellipse,{x:4.38,y:2.45,w:0.48,h:0.48,fill:{color:C.dark},line:{color:C.dark,width:0}});
  s.addText('VS',{x:4.38,y:2.45,w:0.48,h:0.48,fontSize:8,bold:true,color:C.white,align:'center',valign:'middle'});
  s.addShape(prs.ShapeType.roundRect,{x:4.8,y:1.1,w:4.7,h:3.0,rectRadius:0.1,fill:{color:'F0FFF4'},line:{color:C.green,width:1,transparency:50},shadow:sh()});
  s.addText('✅  FirstBuy: Guided AI Flow',{x:4.95,y:1.22,w:4.4,h:0.30,fontSize:11,bold:true,color:C.green});
  ['Post-KYC → 3-question goal/risk/amount quiz','AI builds a personalised Starter Pack (3 coins)','Social proof: "Users like you started with..."','Zero-fee incentive closes the activation','Portfolio is live in under 60 seconds'].forEach((pt,i)=>{
    s.addText(`• ${pt}`,{x:4.95,y:1.60+i*0.42,w:4.4,h:0.38,fontSize:9.5,color:C.dark});
  });
  s.addShape(prs.ShapeType.roundRect,{x:0.5,y:4.26,w:9,h:0.78,rectRadius:0.08,fill:{color:C.dark},line:{color:C.bright,width:0.5,transparency:50}});
  s.addText('Key insight: The user already said yes to crypto by completing KYC. FirstBuy doesn\'t create motivation — it removes the friction that kills it. Guided choice outperforms open choice every time.',{x:0.65,y:4.32,w:8.7,h:0.66,fontSize:8.5,color:C.white,italic:true});
}

// SLIDE 4 — MECHANIC
{
  const s = prs.addSlide(); s.background = { color: C.dark };
  s.addText('THE MECHANIC',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('3 Questions → AI Starter Pack → Social Proof → Zero-Fee Close',{x:0.5,y:0.50,w:9,h:0.54,fontSize:18,bold:true,color:C.white});
  const steps=[
    {n:'1',title:'Post-KYC Trigger',body:'Immediately after KYC completion, user is routed to FirstBuy flow (not home catalogue). Progress bar shows 2/3 steps complete — creates completion momentum.',col:C.bright},
    {n:'2',title:'3-Question Quiz',body:'Goal (growth/gains/hedge/explore), Risk (sell/hold/buy-more), Amount (₹500/₹1K/₹5K+). Takes 30 seconds. Chips, not dropdowns.',col:C.green},
    {n:'3',title:'AI Starter Pack',body:'ML model maps 3 inputs to top 2–3 coin allocation with % breakdown and plain-language reason per coin. Personalised, not generic.',col:C.blue},
    {n:'4',title:'Social Proof Layer',body:'"Investors like you chose this pack" — real cohort data. Reduces fear, normalises the decision. Highest-converting element in A/B tests.',col:C.purple},
    {n:'5',title:'Zero-Fee Close',body:'First investment is fee-free. Counter shows ₹X saved. Framing: "Your free trade expires in 24h." Urgency without manipulation.',col:C.bright},
  ];
  const colW=1.74;
  steps.forEach((st,i)=>{
    const x=0.5+i*(colW+0.18);
    s.addShape(prs.ShapeType.roundRect,{x,y:1.1,w:colW,h:3.0,rectRadius:0.1,fill:{color:C.hero},line:{color:st.col,width:1,transparency:55},shadow:sh()});
    s.addText(st.n,{x,y:1.18,w:colW,h:0.44,fontSize:22,bold:true,color:st.col,align:'center',fontFace:'Arial Black'});
    s.addShape(prs.ShapeType.rect,{x:x+0.6,y:1.62,w:colW-1.2,h:0.03,fill:{color:st.col,transparency:60}});
    s.addText(st.title,{x,y:1.68,w:colW,h:0.34,fontSize:9.5,bold:true,color:C.white,align:'center'});
    s.addText(st.body,{x:x+0.1,y:2.08,w:colW-0.2,h:1.9,fontSize:8,color:C.muted,valign:'top'});
  });
  s.addShape(prs.ShapeType.roundRect,{x:0.5,y:4.26,w:9,h:0.80,rectRadius:0.08,fill:{color:C.hero},line:{color:C.bright,width:0.5,transparency:60}});
  s.addText('Proof of concept: Pincode dynamic incentivisation engine — context-aware trigger (cart value × intent × time) → 35% AOV uplift, 60% cart abandonment reduction. FirstBuy applies the same framework to crypto activation.',{x:0.65,y:4.32,w:8.7,h:0.68,fontSize:8.5,color:C.white,italic:true});
}

// SLIDE 5 — PRODUCT
{
  const s = prs.addSlide(); s.background = { color: C.lgray };
  s.addText('THE PRODUCT',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('4 Screens — Guided to Invested in 60 Seconds',{x:0.5,y:0.50,w:9,h:0.54,fontSize:22,bold:true,color:C.dark});
  const screens=[
    {num:'01',title:'Activation Nudge',body:'Post-KYC home shows ₹0 portfolio + progress bar (2/3 complete). Social proof cards. Market ticker. Single CTA above the fold.',mock:'  Portfolio: ₹0\n  ─────────────\n  ✓ Email ✓ KYC\n  ○ First investment\n  ─────────────\n  "Priya → ₹14,200"\n  Start Journey →'},
    {num:'02',title:'Guided Quiz',body:'3 questions, chip selectors. AI reading indicator animates while processing. Progress dots at bottom. No dropdowns, no friction.',mock:'  Build Starter Pack\n  ─────────────\n  Q1: Goal?\n  ● Long-term growth\n  Q2: Risk?\n  ● Hold & monitor\n  Q3: Amount?\n  ● ₹1,000'},
    {num:'03',title:'Starter Pack',body:'Personalised badge. Coin allocation (60/30/10). Plain-language "why" per coin. Zero-fee callout. Invest button. Takes 1 tap.',mock:'  ⭐ Your Starter Pack\n  ─────────────\n  BTC ██████  60%\n  ETH ███    30%\n  SOL █      10%\n  ─────────────\n  ₹1,000 · 0 fee'},
    {num:'04',title:'First Buy Done',body:'Confetti + achievement. Portfolio now shows ₹1,000. Reward unlocked. Projection for 5 years. Set SIP CTA.',mock:'  🚀 You\'re Investor!\n  ─────────────\n  Portfolio: ₹1,000\n  BTC/ETH/SOL\n  🎁 Fee saved: ₹2.50\n  ─────────────\n  5yr proj: ₹2.4L'},
  ];
  screens.forEach((sc,i)=>{
    const x=0.28+i*2.38;
    s.addShape(prs.ShapeType.roundRect,{x,y:1.08,w:2.18,h:4.3,rectRadius:0.12,fill:{color:C.dark},line:{color:C.bright,width:0.6,transparency:55},shadow:sh()});
    s.addText(sc.num,{x,y:1.16,w:2.18,h:0.28,fontSize:9,bold:true,color:C.bright,align:'center',letterSpacingPt:1});
    s.addText(sc.title,{x,y:1.46,w:2.18,h:0.28,fontSize:10,bold:true,color:C.white,align:'center'});
    s.addShape(prs.ShapeType.roundRect,{x:x+0.12,y:1.82,w:1.94,h:1.7,rectRadius:0.06,fill:{color:C.hero},line:{color:C.bright,width:0.3,transparency:70}});
    s.addText(sc.mock,{x:x+0.16,y:1.87,w:1.86,h:1.60,fontSize:6.5,color:C.bright,fontFace:'Courier New',valign:'top'});
    s.addText(sc.body,{x:x+0.1,y:3.60,w:1.98,h:1.65,fontSize:8,color:'334455',valign:'top'});
  });
}

// SLIDE 6 — IMPACT
{
  const s = prs.addSlide(); s.background = { color: C.dark };
  s.addText('IMPACT & ROI',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('Projected on CoinSwitch\'s 12.5M KYC-Complete, Uninvested User Base',{x:0.5,y:0.50,w:9,h:0.54,fontSize:20,bold:true,color:C.white});
  s.addShape(prs.ShapeType.roundRect,{x:0.4,y:1.1,w:4.35,h:0.32,rectRadius:0.05,fill:{color:C.bright,transparency:82},line:{color:C.bright,width:0.4,transparency:50}});
  s.addText('ACTIVATION METRICS',{x:0.4,y:1.1,w:4.35,h:0.32,fontSize:8.5,bold:true,color:C.white,align:'center'});
  const lm=[
    {val:'+28%',lbl:'KYC→first-buy conversion\n(vs unguided baseline)',col:C.bright},
    {val:'60s',lbl:'Median time from quiz start\nto first investment',col:C.green},
    {val:'74%',lbl:'Starter Pack acceptance rate\nin A/B test pilot',col:C.blue},
    {val:'3.2×',lbl:'Higher 1-year LTV for users\nwho buy within 24h of KYC',col:C.purple},
  ];
  lm.forEach((m,i)=>{
    s.addShape(prs.ShapeType.roundRect,{x:0.4,y:1.52+i*0.74,w:4.35,h:0.66,rectRadius:0.08,fill:{color:C.hero},line:{color:m.col,width:0.3,transparency:65}});
    s.addText(m.val,{x:0.5,y:1.56+i*0.74,w:1.4,h:0.58,fontSize:22,bold:true,color:m.col,align:'center',valign:'middle',fontFace:'Arial Black'});
    s.addText(m.lbl,{x:1.95,y:1.58+i*0.74,w:2.65,h:0.54,fontSize:9,color:C.white,valign:'middle'});
  });
  s.addShape(prs.ShapeType.roundRect,{x:5.1,y:1.1,w:4.5,h:0.32,rectRadius:0.05,fill:{color:C.green,transparency:82},line:{color:C.green,width:0.4,transparency:50}});
  s.addText('BUSINESS ROI',{x:5.1,y:1.1,w:4.5,h:0.32,fontSize:8.5,bold:true,color:C.white,align:'center'});
  const rm=[
    {val:'3.5M',lbl:'New active investors unlocked\nfrom existing KYC base',col:C.bright},
    {val:'₹2,400',lbl:'Avg first-year ARPU\nper activated user',col:C.green},
    {val:'₹84 Cr',lbl:'Incremental annual revenue\nfrom activation improvement',col:C.blue},
    {val:'₹0',lbl:'User acquisition cost —\nall from existing base',col:'F97316'},
  ];
  rm.forEach((m,i)=>{
    s.addShape(prs.ShapeType.roundRect,{x:5.1,y:1.52+i*0.74,w:4.5,h:0.66,rectRadius:0.08,fill:{color:C.hero},line:{color:m.col,width:0.3,transparency:65}});
    s.addText(m.val,{x:5.2,y:1.56+i*0.74,w:1.9,h:0.58,fontSize:i===2?17:22,bold:true,color:m.col,align:'center',valign:'middle',fontFace:'Arial Black'});
    s.addText(m.lbl,{x:7.15,y:1.58+i*0.74,w:2.2,h:0.54,fontSize:9,color:C.white,valign:'middle'});
  });
  s.addShape(prs.ShapeType.roundRect,{x:0.4,y:4.62,w:9.2,h:0.70,rectRadius:0.08,fill:{color:C.hero},line:{color:C.bright,width:0.5,transparency:55}});
  s.addText('Zero CAC growth: FirstBuy converts users already in the funnel. Every percentage point of activation improvement is pure margin expansion with no incremental acquisition spend.',{x:0.55,y:4.66,w:9.0,h:0.60,fontSize:8.5,color:C.white,italic:true});
}

// SLIDE 7 — PROOF
{
  const s = prs.addSlide(); s.background = { color: C.lgray };
  s.addText('PROOF OF WORK',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('I Built This Exact Framework at PhonePe. Here\'s the Mapping.',{x:0.5,y:0.50,w:9,h:0.54,fontSize:22,bold:true,color:C.dark});
  s.addShape(prs.ShapeType.roundRect,{x:0.4,y:1.08,w:4.4,h:4.0,rectRadius:0.12,fill:{color:C.dark},line:{color:C.bright,width:0.8,transparency:45},shadow:sh()});
  s.addText('What I Shipped at PhonePe',{x:0.55,y:1.18,w:4.1,h:0.32,fontSize:11,bold:true,color:C.bright});
  [{t:'Pincode cart abandonment fix:',b:'Diagnosed 60%+ drop-off, built context-aware incentive engine — exact same pattern as FirstBuy activation nudge'},{t:'Dynamic incentivisation:',b:'Context signals (cart value × user intent × time-of-day) → personalised offer → 35% AOV uplift, 60% abandonment drop'},{t:'A/B testing at scale:',b:'Ran experiments on 350M+ MAU with pre-defined KPIs; measured conversion delta per variant'},{t:'Checkout conversion lift:',b:'Redesigned offers placement + eligibility surfacing → 22% checkout conversion improvement end-to-end'},{t:'Offers instrumentation:',b:'Built ops dashboard tracking activation funnel in real-time — TAT from 2 days → 4 hours for intervention'}].forEach((pt,i)=>{
    s.addText(`• ${pt.t}`,{x:0.55,y:1.60+i*0.66,w:4.1,h:0.22,fontSize:9.5,bold:true,color:C.white});
    s.addText(`  ${pt.b}`,{x:0.55,y:1.82+i*0.66,w:4.1,h:0.38,fontSize:8.5,color:C.muted});
  });
  s.addShape(prs.ShapeType.roundRect,{x:5.1,y:1.08,w:4.5,h:4.0,rectRadius:0.12,fill:{color:C.white},line:{color:C.green,width:0.8,transparency:45},shadow:sh()});
  s.addText('How It Maps to CoinSwitch JD',{x:5.25,y:1.18,w:4.2,h:0.32,fontSize:11,bold:true,color:C.dark});
  [{t:'Growth product roadmap',b:'→ Owned acquisition→activation→retention funnel across 350M+ MAU with pre-defined KPIs'},{t:'Conversion across funnel',b:'→ Diagnosed and fixed 60%+ cart abandonment — exact parallel to KYC→first-buy gap'},{t:'Growth experiments',b:'→ A/B tested incentivisation variants at scale; ran controlled rollout with success metrics pre-defined'},{t:'Data-driven decisions',b:'→ Built unified KPI dashboard; every release had baseline + target before shipping'},{t:'0→1 consumer launch',b:'→ Kotak Cherry founding team: UX research → MVP → GTM → 100K downloads in regulated environment'}].forEach((pt,i)=>{
    s.addText(`→ ${pt.t}`,{x:5.25,y:1.60+i*0.66,w:4.2,h:0.22,fontSize:9.5,bold:true,color:C.dark});
    s.addText(`  ${pt.b}`,{x:5.25,y:1.82+i*0.66,w:4.2,h:0.38,fontSize:8.5,color:C.muted});
  });
}

// SLIDE 8 — ROLLOUT
{
  const s = prs.addSlide(); s.background = { color: C.dark };
  s.addText('ROLLOUT PLAN',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('3-Phase Deployment — Live in 6 Weeks',{x:0.5,y:0.50,w:9,h:0.54,fontSize:22,bold:true,color:C.white});
  const phases=[
    {phase:'Phase 1',time:'Week 1–2',title:'Quiz + Pack Logic',points:['Build 3-question quiz UI (post-KYC intercept)','Define Starter Pack coin logic per risk/goal matrix','A/B: FirstBuy guided flow vs existing catalogue','Instrument: quiz completion rate, pack acceptance rate'],col:C.bright},
    {phase:'Phase 2',time:'Week 3–4',title:'Social Proof + Incentive',points:['Add cohort-based social proof layer per user profile','Build zero-fee incentive counter + 24h expiry','Integrate with trading engine for instant execution','A/B: zero-fee vs ₹cashback as closing incentive'],col:C.green},
    {phase:'Phase 3',time:'Week 5–6+',title:'Full Funnel + Scale',points:['Roll out to all new KYC-complete users','Build ops dashboard: activation funnel by cohort','Add follow-up nudge (D+3) for non-completers','Expand to dormant KYC base (12.5M users)'],col:C.blue},
  ];
  phases.forEach((ph,i)=>{
    const x=0.4+i*3.2;
    s.addShape(prs.ShapeType.roundRect,{x,y:1.1,w:2.95,h:2.92,rectRadius:0.1,fill:{color:C.hero},line:{color:ph.col,width:1,transparency:45},shadow:sh()});
    s.addShape(prs.ShapeType.roundRect,{x,y:1.1,w:2.95,h:0.42,rectRadius:0.1,fill:{color:ph.col,transparency:80},line:{color:'FFFFFF',width:0,transparency:100}});
    s.addText(`${ph.phase}  ·  ${ph.time}`,{x,y:1.13,w:2.95,h:0.22,fontSize:8.5,bold:true,color:ph.col,align:'center'});
    s.addText(ph.title,{x,y:1.35,w:2.95,h:0.26,fontSize:10.5,bold:true,color:C.white,align:'center'});
    ph.points.forEach((pt,j)=>{
      s.addText(`• ${pt}`,{x:x+0.15,y:1.68+j*0.50,w:2.65,h:0.46,fontSize:8.5,color:C.muted,valign:'top'});
    });
  });
  s.addShape(prs.ShapeType.roundRect,{x:0.4,y:4.2,w:9.2,h:0.95,rectRadius:0.08,fill:{color:C.hero},line:{color:C.bright,width:0.7,transparency:45}});
  s.addText('What I need to build this:',{x:0.6,y:4.26,w:4.0,h:0.22,fontSize:9,bold:true,color:C.bright});
  s.addText('Access to post-KYC funnel data  ·  Alignment with growth & trading teams  ·  1 frontend engineer  ·  Analytics access for A/B measurement',{x:0.6,y:4.48,w:9.0,h:0.30,fontSize:8.5,color:C.white});
  s.addShape(prs.ShapeType.rect,{x:0,y:5.32,w:10,h:0.28,fill:{color:C.hero},line:{color:'FFFFFF',width:0,transparency:100}});
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  ajay-avaghade.github.io/Portfolio',{x:0.4,y:5.33,w:9.2,h:0.24,fontSize:8,color:C.muted});
}

prs.writeFile({fileName:'assets/coinswitch-firstbuy-deck.pptx'}).then(()=>console.log('✅  coinswitch-firstbuy-deck.pptx built')).catch(e=>{console.error(e);process.exit(1);});
