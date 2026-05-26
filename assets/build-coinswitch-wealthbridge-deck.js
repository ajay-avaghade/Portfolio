const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0A0E27', hero: '111432',
  bright: '00D4AA', gold: 'F7B731',
  green:  '22C55E', purple: '7C3AED',
  blue:   '3B82F6', white:  'FFFFFF',
  lgray:  'F0F8FF', muted:  '8B9CB6',
};
const sh = () => ({ type:'outer', color:'000000', blur:4, offset:1, angle:45, opacity:0.14 });

// SLIDE 1
{
  const s = prs.addSlide(); s.background = { color: C.dark };
  for(let i=0;i<5;i++){
    const lx=-0.5+i*2.2, lw=Math.min(3.5,10-lx-0.15);
    if(lx>10||lw<=0) continue;
    s.addShape(prs.ShapeType.line,{x:lx,y:0,w:lw,h:5.625,line:{color:C.bright,width:0.4},rotate:30,transparency:88});
  }
  s.addShape(prs.ShapeType.roundRect,{x:0.5,y:0.26,w:3.4,h:0.26,rectRadius:0.05,fill:{color:C.bright,transparency:88},line:{color:C.bright,width:0.5,transparency:50}});
  s.addText('PEEPALCO · COINSWITCH GROWTH PM',{x:0.5,y:0.26,w:3.4,h:0.26,fontSize:7,bold:true,color:C.bright,align:'center',letterSpacingPt:1});
  s.addText('CoinSwitch WealthBridge',{x:0.5,y:0.72,w:7.5,h:1.0,fontSize:40,bold:true,color:C.white,fontFace:'Arial Black'});
  s.addShape(prs.ShapeType.rect,{x:0.5,y:1.72,w:2.0,h:0.05,fill:{color:C.bright}});
  s.addText('Behavioural Scoring Engine: Converting Crypto Users to Mutual Fund Investors',{x:0.5,y:1.88,w:7.2,h:0.38,fontSize:13,color:'AABBD0',italic:true});
  s.addText('Presented by Ajay Avaghade  ·  Growth PM',{x:0.5,y:2.36,w:7,h:0.26,fontSize:9.5,color:C.muted});
  s.addShape(prs.ShapeType.roundRect,{x:7.0,y:3.8,w:2.75,h:1.55,rectRadius:0.1,fill:{color:C.bright,transparency:86},line:{color:C.bright,width:1,transparency:50},shadow:sh()});
  s.addText('2.1M',{x:7.0,y:3.88,w:2.75,h:0.72,fontSize:40,bold:true,color:C.bright,align:'center',fontFace:'Arial Black'});
  s.addText('crypto users identified\nas MF-ready today',{x:7.0,y:4.60,w:2.75,h:0.55,fontSize:8,color:C.muted,align:'center'});
}

// SLIDE 2
{
  const s = prs.addSlide(); s.background = { color: C.dark };
  s.addText('THE PROBLEM',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('PeepalCo Has 25M Crypto Users. None Were Offered Their MF Product.',{x:0.5,y:0.50,w:9,h:0.65,fontSize:20,bold:true,color:C.white});
  const cols=[
    {icon:'🏦',stat:'₹54L Cr',lbl:'India MF AUM (May 2024)\ngrowing 40% YoY',sub:'The market is huge — PeepalCo\nhas zero share today'},
    {icon:'🔗',stat:'0',lbl:'Crypto-to-MF cross-sell\ncurrently triggered',sub:'25M crypto users sit in\na silo. No bridge exists.'},
    {icon:'📊',stat:'2.4×',lbl:'Higher LTV for users with\nboth crypto + MF accounts',sub:'Multi-product users churn\nat a fraction of the rate'},
  ];
  cols.forEach((c,i)=>{
    const x=0.5+i*3.1;
    s.addShape(prs.ShapeType.roundRect,{x,y:1.15,w:2.9,h:2.8,rectRadius:0.12,fill:{color:C.hero},line:{color:C.bright,width:0.3,transparency:70},shadow:sh()});
    s.addText(c.icon,{x,y:1.25,w:2.9,h:0.45,fontSize:22,align:'center'});
    s.addText(c.stat,{x,y:1.72,w:2.9,h:0.72,fontSize:i===0?24:36,bold:true,color:C.bright,align:'center',fontFace:'Arial Black'});
    s.addText(c.lbl,{x,y:2.48,w:2.9,h:0.58,fontSize:9.5,color:C.white,align:'center',bold:true});
    s.addText(c.sub,{x,y:3.10,w:2.9,h:0.60,fontSize:8.5,color:C.muted,align:'center'});
  });
  s.addShape(prs.ShapeType.roundRect,{x:0.5,y:4.22,w:9,h:0.80,rectRadius:0.08,fill:{color:C.bright,transparency:88},line:{color:C.bright,width:0.6,transparency:55}});
  s.addText('Root cause: PeepalCo launched MF but has no intelligence engine to identify which crypto users are ready for the transition, or what moment triggers highest conversion.',{x:0.65,y:4.28,w:8.7,h:0.68,fontSize:8.5,color:C.white,italic:true});
}

// SLIDE 3
{
  const s = prs.addSlide(); s.background = { color: C.lgray };
  s.addText('THE INSIGHT',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('Crypto Portfolio Behaviour Is the Best MF Readiness Signal Available.',{x:0.5,y:0.50,w:9,h:0.65,fontSize:21,bold:true,color:C.dark});
  s.addShape(prs.ShapeType.roundRect,{x:0.5,y:1.1,w:4.0,h:3.0,rectRadius:0.1,fill:{color:'F0F8FF'},line:{color:C.purple,width:1,transparency:50},shadow:sh()});
  s.addText('❌  Current: Spray Marketing',{x:0.65,y:1.22,w:3.7,h:0.30,fontSize:11,bold:true,color:C.purple});
  ['Same MF push email to all 25M users','No differentiation by portfolio maturity','No timing intelligence (always off-peak)','3-month-old user gets same pitch as 3-year user','Conversion rate: ~0.8% — noise level'].forEach((pt,i)=>{
    s.addText(`• ${pt}`,{x:0.65,y:1.60+i*0.42,w:3.7,h:0.38,fontSize:9.5,color:C.dark});
  });
  s.addShape(prs.ShapeType.ellipse,{x:4.38,y:2.45,w:0.48,h:0.48,fill:{color:C.dark},line:{color:C.dark,width:0}});
  s.addText('VS',{x:4.38,y:2.45,w:0.48,h:0.48,fontSize:8,bold:true,color:C.white,align:'center',valign:'middle'});
  s.addShape(prs.ShapeType.roundRect,{x:4.8,y:1.1,w:4.7,h:3.0,rectRadius:0.1,fill:{color:'F0FFF4'},line:{color:C.green,width:1,transparency:50},shadow:sh()});
  s.addText('✅  WealthBridge: Behavioural Scoring',{x:4.95,y:1.22,w:4.4,h:0.30,fontSize:11,bold:true,color:C.green});
  ['ML model reads 4 crypto portfolio signals','Scores MF-readiness (0–100) per user daily','Surfaces card only to score ≥ 60 users','Triggers on portfolio maturity event, not calendar','Conversion rate: 18% of flagged users'].forEach((pt,i)=>{
    s.addText(`• ${pt}`,{x:4.95,y:1.60+i*0.42,w:4.4,h:0.38,fontSize:9.5,color:C.dark});
  });
  s.addShape(prs.ShapeType.roundRect,{x:0.5,y:4.26,w:9,h:0.78,rectRadius:0.08,fill:{color:C.dark},line:{color:C.bright,width:0.5,transparency:50}});
  s.addText('Key insight: The 4 strongest MF-readiness signals — portfolio age, DCA behaviour, blue-chip ratio, dip-holding — are all readable from CoinSwitch\'s own data. No survey needed.',{x:0.65,y:4.32,w:8.7,h:0.66,fontSize:8.5,color:C.white,italic:true});
}

// SLIDE 4
{
  const s = prs.addSlide(); s.background = { color: C.dark };
  s.addText('THE MECHANIC',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('4 Behavioural Signals → ML Score → In-App Card → SIP Setup → AUM',{x:0.5,y:0.50,w:9,h:0.54,fontSize:17,bold:true,color:C.white});
  const steps=[
    {n:'1',title:'4 Signal Inputs',body:'Portfolio age (months active), DCA pattern (months invested / 12), Blue-chip ratio (BTC+ETH as % of portfolio), Dip-holding (held through ≥1 market decline)',col:C.bright},
    {n:'2',title:'MF Readiness Score',body:'ML model weights 4 signals → 0–100 score. Score ≥ 60: WealthBridge card appears in-app. Score ≥ 80: proactive push notification with personalised MF match.',col:C.gold},
    {n:'3',title:'Matched MF Card',body:'In-app card shows score + match reasoning. Specific MF recommended (not a generic list). Explains how it complements existing crypto portfolio.',col:C.purple},
    {n:'4',title:'One-Flow SIP Setup',body:'Amount selector → frequency → SEBI mandate → confirm. Direct plan (no distributor fee). eSIP auto-debit. Takes under 2 minutes.',col:C.green},
    {n:'5',title:'Closed Loop',body:'30-day post-activation tracking: SIP continuation rate, crypto portfolio change, multi-product LTV delta. Feeds model retraining quarterly.',col:C.blue},
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
  s.addText('Proof of concept: PhonePe Propensity-to-Transact ML model — same architecture, same closed loop. Mapped user intent signals to real-time targeting → 32% marketing spend reduction on ₹1,000 Cr/yr budget.',{x:0.65,y:4.32,w:8.7,h:0.68,fontSize:8.5,color:C.white,italic:true});
}

// SLIDE 5
{
  const s = prs.addSlide(); s.background = { color: C.lgray };
  s.addText('THE PRODUCT',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('4 Screens — Crypto Home to First SIP in Under 2 Minutes',{x:0.5,y:0.50,w:9,h:0.54,fontSize:22,bold:true,color:C.dark});
  const screens=[
    {num:'01',title:'MF-Ready Signal',body:'WealthBridge card surfaces on crypto home when score ≥ 60. Readiness score bar + "See My MF Match" CTA above holding list.',mock:'  Crypto: ₹84,200\n  ─────────────\n  🎯 MF-Ready!\n  Score: 82/100\n  ─────────────\n  See MF Match →\n  ─────────────\n  BTC ₹48,600'},
    {num:'02',title:'MF Match Screen',body:'4 readiness signals shown (why you qualify). Best-fit MF card with 3-yr CAGR, min SIP, expense ratio. Portfolio complementarity explained.',mock:'  WealthBridge\n  ─────────────\n  ✅ 14mo active\n  ✅ DCA 9/12 mo\n  ✅ 91% blue-chip\n  ─────────────\n  Mirae Large Cap\n  18.4% CAGR · ₹500'},
    {num:'03',title:'SIP Setup',body:'Amount chips (₹500/₹1K/₹2K/₹5K). Frequency selector (monthly/weekly). 5-year projection. SEBI-compliant mandate in 1 tap.',mock:'  Monthly SIP\n  ─────────────\n  ₹ 2,000 / mo\n  5yr proj: ₹2.1L\n  ─────────────\n  ✅ SEBI regulated\n  ✅ Direct plan\n  Confirm →'},
    {num:'04',title:'Ops Intelligence',body:'2.1M flagged users. Cross-sell funnel (flagged→card→match→SIP). Top cohort by score band. AUM activated + projected annual AUM.',mock:'  Cross-Sell Dash\n  ─────────────\n  Flagged: 2.1M\n  Conv: 18%\n  AUM: ₹420Cr\n  ─────────────\n  ARPU: +₹3,200/yr'},
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

// SLIDE 6
{
  const s = prs.addSlide(); s.background = { color: C.dark };
  s.addText('IMPACT & ROI',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('Projected on 2.1M Flagged MF-Ready CoinSwitch Users',{x:0.5,y:0.50,w:9,h:0.54,fontSize:22,bold:true,color:C.white});
  s.addShape(prs.ShapeType.roundRect,{x:0.4,y:1.1,w:4.35,h:0.32,rectRadius:0.05,fill:{color:C.bright,transparency:82},line:{color:C.bright,width:0.4,transparency:50}});
  s.addText('USER EXPERIENCE',{x:0.4,y:1.1,w:4.35,h:0.32,fontSize:8.5,bold:true,color:C.white,align:'center'});
  [{val:'82',lbl:'Avg MF-readiness score\nfor flagged cohort',col:C.bright},{val:'18%',lbl:'30-day SIP conversion rate\nfrom WealthBridge card',col:C.gold},{val:'2.4×',lbl:'Higher LTV for multi-product\nusers vs crypto-only',col:C.purple},{val:'<2 min',lbl:'Median time from card view\nto first SIP confirmed',col:C.green}].forEach((m,i)=>{
    s.addShape(prs.ShapeType.roundRect,{x:0.4,y:1.52+i*0.74,w:4.35,h:0.66,rectRadius:0.08,fill:{color:C.hero},line:{color:m.col,width:0.3,transparency:65}});
    s.addText(m.val,{x:0.5,y:1.56+i*0.74,w:1.4,h:0.58,fontSize:22,bold:true,color:m.col,align:'center',valign:'middle',fontFace:'Arial Black'});
    s.addText(m.lbl,{x:1.95,y:1.58+i*0.74,w:2.65,h:0.54,fontSize:9,color:C.white,valign:'middle'});
  });
  s.addShape(prs.ShapeType.roundRect,{x:5.1,y:1.1,w:4.5,h:0.32,rectRadius:0.05,fill:{color:C.green,transparency:82},line:{color:C.green,width:0.4,transparency:50}});
  s.addText('OPERATOR ROI',{x:5.1,y:1.1,w:4.5,h:0.32,fontSize:8.5,bold:true,color:C.white,align:'center'});
  [{val:'378K',lbl:'New MF SIP users\n(18% of 2.1M flagged)',col:C.bright},{val:'₹420 Cr',lbl:'New MF AUM activated\nin Year 1',col:C.gold},{val:'₹3,200',lbl:'Additional annual ARPU\nper cross-sold user',col:C.green},{val:'₹0',lbl:'Incremental CAC —\nall from existing user base',col:'F97316'}].forEach((m,i)=>{
    s.addShape(prs.ShapeType.roundRect,{x:5.1,y:1.52+i*0.74,w:4.5,h:0.66,rectRadius:0.08,fill:{color:C.hero},line:{color:m.col,width:0.3,transparency:65}});
    s.addText(m.val,{x:5.2,y:1.56+i*0.74,w:1.9,h:0.58,fontSize:22,bold:true,color:m.col,align:'center',valign:'middle',fontFace:'Arial Black'});
    s.addText(m.lbl,{x:7.15,y:1.58+i*0.74,w:2.2,h:0.54,fontSize:9,color:C.white,valign:'middle'});
  });
  s.addShape(prs.ShapeType.roundRect,{x:0.4,y:4.62,w:9.2,h:0.70,rectRadius:0.08,fill:{color:C.hero},line:{color:C.bright,width:0.5,transparency:55}});
  s.addText('WealthBridge directly enables PeepalCo\'s stated strategic priority: building India\'s leading multi-product wealth platform. The distribution is already built — 25M users. This adds the intelligence layer.',{x:0.55,y:4.66,w:9.0,h:0.60,fontSize:8.5,color:C.white,italic:true});
}

// SLIDE 7
{
  const s = prs.addSlide(); s.background = { color: C.lgray };
  s.addText('PROOF OF WORK',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('Same ML Architecture. Different Asset Class.',{x:0.5,y:0.50,w:9,h:0.54,fontSize:22,bold:true,color:C.dark});
  s.addShape(prs.ShapeType.roundRect,{x:0.4,y:1.08,w:4.4,h:4.0,rectRadius:0.12,fill:{color:C.dark},line:{color:C.bright,width:0.8,transparency:45},shadow:sh()});
  s.addText('What I Shipped at PhonePe',{x:0.55,y:1.18,w:4.1,h:0.32,fontSize:11,bold:true,color:C.bright});
  [{t:'Propensity-to-Transact ML:',b:'User-level intent scoring replacing static cohorts — exactly WealthBridge\'s behavioural scoring architecture'},{t:'₹100 Cr rewards rebuild:',b:'Rebuilt static portfolio into ML-driven marketplace with audience segmentation — 11% YoY revenue growth'},{t:'Brand self-serve tooling:',b:'500+ brands onboarded to targeted campaigns — same precision targeting as WealthBridge\'s cohort matching'},{t:'A/B testing infrastructure:',b:'Ran ML vs rule-based control across 350M MAU with instrumented KPIs at every funnel stage'},{t:'Cross-sell instrumentation:',b:'Built unified funnel dashboard for offer redemption — 26% uplift in user engagement on new products'}].forEach((pt,i)=>{
    s.addText(`• ${pt.t}`,{x:0.55,y:1.60+i*0.66,w:4.1,h:0.22,fontSize:9.5,bold:true,color:C.white});
    s.addText(`  ${pt.b}`,{x:0.55,y:1.82+i*0.66,w:4.1,h:0.38,fontSize:8.5,color:C.muted});
  });
  s.addShape(prs.ShapeType.roundRect,{x:5.1,y:1.08,w:4.5,h:4.0,rectRadius:0.12,fill:{color:C.white},line:{color:C.green,width:0.8,transparency:45},shadow:sh()});
  s.addText('How It Maps to PeepalCo JD',{x:5.25,y:1.18,w:4.2,h:0.32,fontSize:11,bold:true,color:C.dark});
  [{t:'Strategic growth opportunities',b:'→ WealthBridge opens ₹54L Cr MF market from existing 25M user base — clearest strategic lever available'},{t:'Work with Analytics teams',b:'→ Owned DS collaboration from feature engineering → production → post-launch KPI tracking at scale'},{t:'Measure conversion across funnel',b:'→ Built unified funnel dashboard: every step instrumented, TAT 2d → 4h for growth team intervention'},{t:'Grow market share',b:'→ CoinSwitch enters MF space at near-zero CAC by bridging existing crypto users'},{t:'Drive growth experiments',b:'→ WealthBridge is designed as an A/B experiment: ML scoring vs spray marketing — measurable in 30 days'}].forEach((pt,i)=>{
    s.addText(`→ ${pt.t}`,{x:5.25,y:1.60+i*0.66,w:4.2,h:0.22,fontSize:9.5,bold:true,color:C.dark});
    s.addText(`  ${pt.b}`,{x:5.25,y:1.82+i*0.66,w:4.2,h:0.38,fontSize:8.5,color:C.muted});
  });
}

// SLIDE 8
{
  const s = prs.addSlide(); s.background = { color: C.dark };
  s.addText('ROLLOUT PLAN',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('3-Phase Deployment — First AUM in 8 Weeks',{x:0.5,y:0.50,w:9,h:0.54,fontSize:22,bold:true,color:C.white});
  const phases=[
    {phase:'Phase 1',time:'Week 1–3',title:'Signal Model + Baseline',points:['Build MF-readiness scoring model on portfolio data','Define 4 signal weights via A/B on small cohort','Instrument cross-sell funnel: card view → SIP start','A/B: WealthBridge card vs generic MF push email'],col:C.bright},
    {phase:'Phase 2',time:'Week 4–6',title:'Match Engine + SIP Flow',points:['Build personalised MF recommendation per score band','Launch one-flow SIP setup with SEBI mandate integration','Deploy ops cross-sell dashboard (funnel + cohort AUM)','Tune score threshold (60/80) based on conversion data'],col:C.gold},
    {phase:'Phase 3',time:'Week 7–8+',title:'Scale + Multi-Product',points:['Roll out to all 2.1M flagged users','Expand to equity products (second bridge product)','Quarterly model retraining on new portfolio signals','Launch WealthBridge for new KYC users at onboarding'],col:C.green},
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
  s.addText('Access to portfolio age + transaction history  ·  MF product team alignment for SIP API  ·  1 DS engineer  ·  Analytics for funnel measurement',{x:0.6,y:4.48,w:9.0,h:0.30,fontSize:8.5,color:C.white});
  s.addShape(prs.ShapeType.rect,{x:0,y:5.32,w:10,h:0.28,fill:{color:C.hero},line:{color:'FFFFFF',width:0,transparency:100}});
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  ajay-avaghade.github.io/Portfolio',{x:0.4,y:5.33,w:9.2,h:0.24,fontSize:8,color:C.muted});
}

prs.writeFile({fileName:'assets/coinswitch-wealthbridge-deck.pptx'}).then(()=>console.log('✅  coinswitch-wealthbridge-deck.pptx built')).catch(e=>{console.error(e);process.exit(1);});
