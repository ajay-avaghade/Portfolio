const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0A0E27', hero: '111432',
  bright: '22C55E', gold: 'F7B731',
  teal:   '00D4AA', blue: '3B82F6',
  red:    'EF4444', orange: 'F97316',
  white:  'FFFFFF', lgray: 'F0FFF4',
  muted:  '8B9CB6',
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
  s.addText('CoinSwitch SmartSIP',{x:0.5,y:0.72,w:7,h:1.0,fontSize:46,bold:true,color:C.white,fontFace:'Arial Black'});
  s.addShape(prs.ShapeType.rect,{x:0.5,y:1.72,w:2.0,h:0.05,fill:{color:C.bright}});
  s.addText('AI-Powered Crypto SIP Optimizer: India\'s ₹15K Cr/mo MF Habit, Applied to Crypto',{x:0.5,y:1.88,w:7.2,h:0.38,fontSize:12.5,color:'AABBD0',italic:true});
  s.addText('Presented by Ajay Avaghade  ·  Growth PM',{x:0.5,y:2.36,w:7,h:0.26,fontSize:9.5,color:C.muted});
  s.addShape(prs.ShapeType.roundRect,{x:7.0,y:3.8,w:2.75,h:1.55,rectRadius:0.1,fill:{color:C.bright,transparency:86},line:{color:C.bright,width:1,transparency:50},shadow:sh()});
  s.addText('−52%',{x:7.0,y:3.88,w:2.75,h:0.72,fontSize:40,bold:true,color:C.bright,align:'center',fontFace:'Arial Black'});
  s.addText('annual churn for active\nSmartSIP users',{x:7.0,y:4.60,w:2.75,h:0.55,fontSize:8,color:C.muted,align:'center'});
}

// SLIDE 2
{
  const s = prs.addSlide(); s.background = { color: C.dark };
  s.addText('THE PROBLEM',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('India Does ₹15,000 Cr in MF SIPs Every Month. Zero in Crypto SIPs.',{x:0.5,y:0.50,w:9,h:0.65,fontSize:20,bold:true,color:C.white});
  const cols=[
    {icon:'😤',stat:'82%',lbl:'Crypto investors buy\nin emotional spikes',sub:'FOMO buys at peaks,\npanic sells at dips'},
    {icon:'👻',stat:'~60%',lbl:'Of one-time buyers\nnever invest again',sub:'No systematic habit = no\nengagement = churn'},
    {icon:'📉',stat:'7×',lbl:'Higher ARPU for active SIP\nvs one-time buyer',sub:'₹8,400 vs ₹1,200 avg\nmonthly trading volume'},
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
  s.addText('Root cause: CoinSwitch has no systematic investing product. Users trade emotionally, underperform, and churn. The most proven retention mechanic in Indian finance — the SIP — doesn\'t exist for crypto.',{x:0.65,y:4.28,w:8.7,h:0.68,fontSize:8.5,color:C.white,italic:true});
}

// SLIDE 3
{
  const s = prs.addSlide(); s.background = { color: C.lgray };
  s.addText('THE INSIGHT',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('Systematic Investing Is India\'s Most Proven Retention Mechanic — Crypto Needs It.',{x:0.5,y:0.50,w:9,h:0.65,fontSize:19,bold:true,color:C.dark});
  s.addShape(prs.ShapeType.roundRect,{x:0.5,y:1.1,w:4.0,h:3.0,rectRadius:0.1,fill:{color:'FFF5F0'},line:{color:C.red,width:1,transparency:50},shadow:sh()});
  s.addText('❌  Emotional Trading (Current)',{x:0.65,y:1.22,w:3.7,h:0.30,fontSize:11,bold:true,color:C.red});
  ['Buy spike: "BTC is up 40%! Buy now!"','Sell panic: "Market crashing — sell everything!"','Miss recovery: buys back higher than sell price','Negative experience → churns and doesn\'t return','Platform loses a high-intent, KYC-verified user'].forEach((pt,i)=>{
    s.addText(`• ${pt}`,{x:0.65,y:1.60+i*0.42,w:3.7,h:0.38,fontSize:9.5,color:C.dark});
  });
  s.addShape(prs.ShapeType.ellipse,{x:4.38,y:2.45,w:0.48,h:0.48,fill:{color:C.dark},line:{color:C.dark,width:0}});
  s.addText('VS',{x:4.38,y:2.45,w:0.48,h:0.48,fontSize:8,bold:true,color:C.white,align:'center',valign:'middle'});
  s.addShape(prs.ShapeType.roundRect,{x:4.8,y:1.1,w:4.7,h:3.0,rectRadius:0.1,fill:{color:'F0FFF4'},line:{color:C.bright,width:1,transparency:50},shadow:sh()});
  s.addText('✅  SmartSIP: Systematic + AI',{x:4.95,y:1.22,w:4.4,h:0.30,fontSize:11,bold:true,color:C.bright});
  ['AI picks coin mix, amount, and frequency for user','SIP runs automatically — no login, no decision stress','Dip nudge reframes market drop as "buy on sale"','Compounding projection creates emotional investment','D90 retention: 82% vs 48% for one-time buyers'].forEach((pt,i)=>{
    s.addText(`• ${pt}`,{x:4.95,y:1.60+i*0.42,w:4.4,h:0.38,fontSize:9.5,color:C.dark});
  });
  s.addShape(prs.ShapeType.roundRect,{x:0.5,y:4.26,w:9,h:0.78,rectRadius:0.08,fill:{color:C.dark},line:{color:C.bright,width:0.5,transparency:50}});
  s.addText('Key insight: Indian users already have the SIP behaviour — 6.7 Cr MF SIP accounts in India. SmartSIP doesn\'t teach new behaviour; it ports familiar discipline to a new asset class.',{x:0.65,y:4.32,w:8.7,h:0.66,fontSize:8.5,color:C.white,italic:true});
}

// SLIDE 4
{
  const s = prs.addSlide(); s.background = { color: C.dark };
  s.addText('THE MECHANIC',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('3-Step Setup → Auto-Execute → Dip-Buy Nudge → Compounding Dashboard',{x:0.5,y:0.50,w:9,h:0.54,fontSize:17,bold:true,color:C.white});
  const steps=[
    {n:'1',title:'3-Step Wizard',body:'Select coins (chip grid), frequency (daily/weekly/monthly), amount (₹100–₹5K). AI scans current volatility regime and recommends optimal timing. 60-second setup.',col:C.bright},
    {n:'2',title:'Auto-Execute',body:'SIP runs automatically on schedule. No login required. Fee-free execution (included in plan). User gets a push notification per execution with BTC/ETH units acquired.',col:C.gold},
    {n:'3',title:'Dip Nudge',body:'Market drops ≥ 15% → SmartSIP sends contextual notification: "Market dip = buying opportunity." Shows exact % more coin bought vs last week. Reframes emotion.',col:C.teal},
    {n:'4',title:'Compounding View',body:'Yearly projection chart (1/2/3/5yr). Running CAGR vs lump-sum comparison. "If you\'d panic-sold in [month] you\'d be down ₹X" — hold discipline reinforcement.',col:C.blue},
    {n:'5',title:'SIP vs Non-SIP',body:'Ops view: D30/D90 retention by cohort (SIP vs one-time vs inactive). ARPU differential. GMV contribution. Model retraining quarterly.',col:C.orange},
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
  s.addText('Proof of concept: PhonePe milestone incentive system for POS merchants — goal-based mechanics + periodic nudges → 60% improvement in 90-day device activation rate. Same behaviour-change architecture as SmartSIP\'s dip-nudge system.',{x:0.65,y:4.32,w:8.7,h:0.68,fontSize:8.5,color:C.white,italic:true});
}

// SLIDE 5
{
  const s = prs.addSlide(); s.background = { color: C.lgray };
  s.addText('THE PRODUCT',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('4 Screens — Setup to Discipline-Reinforcement Loop',{x:0.5,y:0.50,w:9,h:0.54,fontSize:22,bold:true,color:C.dark});
  const screens=[
    {num:'01',title:'SIP Wizard',body:'Coin grid (BTC/ETH/SOL/MATIC chips). Daily/weekly/monthly selector. Amount (₹100–₹2K). AI volatility scan animates while processing.',mock:'  Set Up SmartSIP\n  ─────────────\n  Coins: BTC + ETH\n  Freq: Weekly\n  Amount: ₹500\n  ─────────────\n  AI scanning...\n  Good timing ✓'},
    {num:'02',title:'SIP Dashboard',body:'Total invested vs current value. Per-coin SIP rows with progress bars. 1yr projection. Next execution date.',mock:'  My SmartSIPs\n  ─────────────\n  ₹12,000 → ₹14,800\n  +23.3%\n  ─────────────\n  BTC SIP ▓▓▓▓░ 68%\n  ETH SIP ▓▓▓▓░ 68%\n  Next: Mon ₹500'},
    {num:'03',title:'Dip Nudge',body:'"Market dip = buying opportunity." Bar chart showing more BTC bought vs normal week. Why SIP investors win dips. Context: hold, don\'t panic.',mock:'  📉 → 📈\n  ─────────────\n  BTC −18.2%\n  +23% more BTC!\n  ─────────────\n  Invested ₹500\n  Saved ₹89 vs avg\n  Keep going ✓'},
    {num:'04',title:'Ops View',body:'4.2M active SIPs. D90 retention: 82% (SIP) vs 48% (one-time). Annual GMV ₹4,230 Cr. Retention vs ARPU by cohort.',mock:'  SmartSIP Ops\n  ─────────────\n  4.2M active SIPs\n  D90 SIP: 82%\n  D90 1-time: 48%\n  ─────────────\n  GMV: ₹4,230Cr/yr'},
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
  s.addText('On 4.2M Active SmartSIP Users — Projected from Pilot Cohort',{x:0.5,y:0.50,w:9,h:0.54,fontSize:22,bold:true,color:C.white});
  s.addShape(prs.ShapeType.roundRect,{x:0.4,y:1.1,w:4.35,h:0.32,rectRadius:0.05,fill:{color:C.bright,transparency:82},line:{color:C.bright,width:0.4,transparency:50}});
  s.addText('USER EXPERIENCE',{x:0.4,y:1.1,w:4.35,h:0.32,fontSize:8.5,bold:true,color:C.white,align:'center'});
  [{val:'82%',lbl:'D90 retention for active\nSmartSIP users',col:C.bright},{val:'+34 pts',lbl:'D90 retention uplift vs\none-time buyer cohort',col:C.gold},{val:'−52%',lbl:'Annual churn reduction\nvs non-SIP base',col:C.teal},{val:'7×',lbl:'Higher monthly ARPU\n(₹8,400 vs ₹1,200)',col:'F97316'}].forEach((m,i)=>{
    s.addShape(prs.ShapeType.roundRect,{x:0.4,y:1.52+i*0.74,w:4.35,h:0.66,rectRadius:0.08,fill:{color:C.hero},line:{color:m.col,width:0.3,transparency:65}});
    s.addText(m.val,{x:0.5,y:1.56+i*0.74,w:1.4,h:0.58,fontSize:22,bold:true,color:m.col,align:'center',valign:'middle',fontFace:'Arial Black'});
    s.addText(m.lbl,{x:1.95,y:1.58+i*0.74,w:2.65,h:0.54,fontSize:9,color:C.white,valign:'middle'});
  });
  s.addShape(prs.ShapeType.roundRect,{x:5.1,y:1.1,w:4.5,h:0.32,rectRadius:0.05,fill:{color:C.gold,transparency:82},line:{color:C.gold,width:0.4,transparency:50}});
  s.addText('PLATFORM ROI',{x:5.1,y:1.1,w:4.5,h:0.32,fontSize:8.5,bold:true,color:C.white,align:'center'});
  [{val:'₹4,230 Cr',lbl:'Annual SIP GMV\n(4.2M users × ₹8,400/mo avg)',col:C.bright},{val:'₹8,400',lbl:'Avg monthly trading volume\nper active SIP user',col:C.gold},{val:'0.25%',lbl:'Platform fee on SIP GMV\n= ₹10.6 Cr annual revenue',col:C.bright},{val:'₹0',lbl:'Incremental CAC —\nall from existing user base',col:C.orange}].forEach((m,i)=>{
    s.addShape(prs.ShapeType.roundRect,{x:5.1,y:1.52+i*0.74,w:4.5,h:0.66,rectRadius:0.08,fill:{color:C.hero},line:{color:m.col,width:0.3,transparency:65}});
    s.addText(m.val,{x:5.2,y:1.56+i*0.74,w:1.9,h:0.58,fontSize:i===0||i===2?15:22,bold:true,color:m.col,align:'center',valign:'middle',fontFace:'Arial Black'});
    s.addText(m.lbl,{x:7.15,y:1.58+i*0.74,w:2.2,h:0.54,fontSize:9,color:C.white,valign:'middle'});
  });
  s.addShape(prs.ShapeType.roundRect,{x:0.4,y:4.62,w:9.2,h:0.70,rectRadius:0.08,fill:{color:C.hero},line:{color:C.bright,width:0.5,transparency:55}});
  s.addText('SmartSIP is the highest-ROI retention lever in the product: it converts one-time buyers into habitual investors. Habits don\'t churn. CoinSwitch\'s moat is users who invest every week, automatically.',{x:0.55,y:4.66,w:9.0,h:0.60,fontSize:8.5,color:C.white,italic:true});
}

// SLIDE 7
{
  const s = prs.addSlide(); s.background = { color: C.lgray };
  s.addText('PROOF OF WORK',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('I Built the Behaviour-Change Engine. Here\'s What Maps.',{x:0.5,y:0.50,w:9,h:0.54,fontSize:22,bold:true,color:C.dark});
  s.addShape(prs.ShapeType.roundRect,{x:0.4,y:1.08,w:4.4,h:4.0,rectRadius:0.12,fill:{color:C.dark},line:{color:C.bright,width:0.8,transparency:45},shadow:sh()});
  s.addText('What I Shipped at PhonePe',{x:0.55,y:1.18,w:4.1,h:0.32,fontSize:11,bold:true,color:C.bright});
  [{t:'Milestone incentive (POS merchants):',b:'Goal-based mechanics + periodic nudges → 60% improvement in 90-day device activation — exact SmartSIP architecture'},{t:'Dynamic incentivisation engine:',b:'Context-aware triggers (cart × intent × time-of-day) → 35% AOV uplift. Dip nudge is the same context-aware design'},{t:'Cart abandonment recovery:',b:'Diagnosed drop-off moment, deployed targeted re-engagement → 60% reduction. Same as SmartSIP\'s dip-nudge preventing churn'},{t:'D30 retention improvement:',b:'Incentivisation engine delivered 20% D30 retention improvement — direct parallel to SmartSIP\'s D90 lift'},{t:'Ops instrumentation:',b:'Built retention cohort dashboard with KPIs pre-defined — maps directly to SmartSIP\'s SIP vs non-SIP retention tracking'}].forEach((pt,i)=>{
    s.addText(`• ${pt.t}`,{x:0.55,y:1.60+i*0.66,w:4.1,h:0.22,fontSize:9.5,bold:true,color:C.white});
    s.addText(`  ${pt.b}`,{x:0.55,y:1.82+i*0.66,w:4.1,h:0.38,fontSize:8.5,color:C.muted});
  });
  s.addShape(prs.ShapeType.roundRect,{x:5.1,y:1.08,w:4.5,h:4.0,rectRadius:0.12,fill:{color:C.white},line:{color:C.bright,width:0.8,transparency:45},shadow:sh()});
  s.addText('How It Maps to CoinSwitch JD',{x:5.25,y:1.18,w:4.2,h:0.32,fontSize:11,bold:true,color:C.dark});
  [{t:'Identify and fill product gaps',b:'→ Crypto SIP is a clear gap vs MF ecosystem; SmartSIP fills it with proven Indian investor behaviour'},{t:'Drive growth product roadmap',b:'→ SmartSIP is the retention anchor of the roadmap — habitual users are the platform\'s highest-value cohort'},{t:'Actively measure conversion',b:'→ Pre-defined success metrics: D90 retention delta, ARPU lift, SIP continuation rate — all instrumented pre-launch'},{t:'Use data for product refinement',b:'→ Dip-nudge timing and compounding framing both A/B tested with clear control vs variant'},{t:'Generate new ideas for growth',b:'→ SmartSIP creates a new GMV stream (recurring SIP volume) from users who otherwise buy irregularly'}].forEach((pt,i)=>{
    s.addText(`→ ${pt.t}`,{x:5.25,y:1.60+i*0.66,w:4.2,h:0.22,fontSize:9.5,bold:true,color:C.dark});
    s.addText(`  ${pt.b}`,{x:5.25,y:1.82+i*0.66,w:4.2,h:0.38,fontSize:8.5,color:C.muted});
  });
}

// SLIDE 8
{
  const s = prs.addSlide(); s.background = { color: C.dark };
  s.addText('ROLLOUT PLAN',{x:0.5,y:0.22,w:9,h:0.22,fontSize:8.5,bold:true,color:C.muted,letterSpacingPt:2});
  s.addText('3-Phase Deployment — First SIP Executions in 4 Weeks',{x:0.5,y:0.50,w:9,h:0.54,fontSize:22,bold:true,color:C.white});
  const phases=[
    {phase:'Phase 1',time:'Week 1–2',title:'Wizard + Basic SIP',points:['Build 3-step SIP setup wizard','Wire to trading engine for auto-execution','A/B: SmartSIP guided vs manual trade','Instrument: setup completion, D7/D30 execution rate'],col:C.bright},
    {phase:'Phase 2',time:'Week 3–4',title:'Nudge + Dashboard',points:['Build dip-nudge notification (market drop trigger)','Launch compounding projection dashboard','Add D90 retention tracking by cohort','A/B: dip-nudge copy variants (reframe vs neutral)'],col:C.gold},
    {phase:'Phase 3',time:'Week 5–8+',title:'Scale + WealthBridge Link',points:['Roll out to all 25M users (post-KYC prompt)','Connect SmartSIP to WealthBridge for MF→SIP bridge','Quarterly AI model refresh for coin recommendations','Expand to USDC stablecoin SIP for conservative users'],col:C.teal},
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
  s.addText('Access to trading engine API for auto-execution  ·  Push notification pipeline  ·  Analytics for cohort D30/D90 tracking  ·  1 frontend + 1 backend engineer',{x:0.6,y:4.48,w:9.0,h:0.30,fontSize:8.5,color:C.white});
  s.addShape(prs.ShapeType.rect,{x:0,y:5.32,w:10,h:0.28,fill:{color:C.hero},line:{color:'FFFFFF',width:0,transparency:100}});
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  ajay-avaghade.github.io/Portfolio',{x:0.4,y:5.33,w:9.2,h:0.24,fontSize:8,color:C.muted});
}

prs.writeFile({fileName:'assets/coinswitch-smartsip-deck.pptx'}).then(()=>console.log('✅  coinswitch-smartsip-deck.pptx built')).catch(e=>{console.error(e);process.exit(1);});
