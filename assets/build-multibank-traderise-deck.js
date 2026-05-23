const pptx = require('pptxgenjs');
const prs = new pptx();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:  '0B1A2E',
  hero:  '162840',
  bright:'00D4AA',
  gold:  'F5A623',
  green: '22C55E',
  red:   'EF4444',
  white: 'FFFFFF',
  lgray: 'F4F8FF',
  muted: '8B9BB4',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });

// ─────────────────────────────────────────────────
// SLIDE 1 — COVER
// ─────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  // Diagonal accent lines
  for(let i=0;i<6;i++){
    s.addShape(prs.ShapeType.line, { x:0+i*1.8, y:0, w:0, h:7.5, line:{ color:C.bright, width:0.4, transparency:88 } });
  }

  // Top accent bar
  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:10, h:0.06, fill:{ color:C.bright } });

  // Company tag
  s.addText('MULTIBANK GROUP  ·  GROWTH MARKETING PM', {
    x:0.6, y:0.3, w:8, h:0.3,
    fontSize:9, bold:true, color:C.bright, charSpacing:8, fontFace:'Calibri'
  });

  // Product name
  s.addText('TradeRise', {
    x:0.6, y:1.0, w:7, h:1.2,
    fontSize:68, bold:true, color:C.white, fontFace:'Calibri'
  });

  // Subtitle
  s.addText('Turning Approved Accounts Into Active Traders', {
    x:0.6, y:2.25, w:7, h:0.5,
    fontSize:22, color:C.bright, fontFace:'Calibri', italic:true
  });

  // Divider line
  s.addShape(prs.ShapeType.rect, { x:0.6, y:2.85, w:1.4, h:0.05, fill:{ color:C.gold } });

  // Presenter
  s.addText('Presented by Ajay Avaghade  ·  Growth & Monetization PM', {
    x:0.6, y:3.05, w:8, h:0.3,
    fontSize:12, color:C.muted, fontFace:'Calibri'
  });

  // Big stat bottom right
  s.addShape(prs.ShapeType.rect, { x:7.0, y:5.2, w:2.7, h:1.8, fill:{ color:C.hero }, line:{ color:C.gold, width:1 } });
  s.addText('50–65%', { x:7.0, y:5.3, w:2.7, h:0.7, fontSize:38, bold:true, color:C.gold, align:'center', fontFace:'Calibri' });
  s.addText('of signups never\nexecute a live trade', { x:7.0, y:6.0, w:2.7, h:0.8, fontSize:11, color:C.muted, align:'center', fontFace:'Calibri' });

  // Bottom line
  s.addShape(prs.ShapeType.rect, { x:0, y:7.44, w:10, h:0.06, fill:{ color:C.gold } });
}

// ─────────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE PROBLEM', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.muted, charSpacing:8, fontFace:'Calibri' });
  s.addText('Trading Platforms Spend to Acquire Users\nWho Never Trade', {
    x:0.6, y:0.6, w:8.5, h:1.1,
    fontSize:28, bold:true, color:C.white, fontFace:'Calibri'
  });

  // 3 stat columns
  const cols3 = [
    { ico:'📉', stat:'50–65%', lbl:'Drop-off after approval', sub:'Users get KYC-approved then vanish before first trade' },
    { ico:'💸', stat:'$120+', lbl:'Avg blended CAC per user', sub:'Spent before they ever generate a dollar of trading volume' },
    { ico:'⏳', stat:'72 hrs', lbl:'Critical activation window', sub:'If no trade in 72h post-approval, 80% never convert' },
  ];
  cols3.forEach((c,i)=>{
    const x = 0.5 + i*3.1;
    s.addShape(prs.ShapeType.rect, { x, y:1.9, w:2.9, h:3.0, fill:{ color:C.hero }, line:{ color:C.bright, width:0.5, transparency:60 }, shadow:makeShadow() });
    s.addText(c.ico, { x, y:2.05, w:2.9, h:0.5, fontSize:24, align:'center' });
    s.addText(c.stat, { x, y:2.6, w:2.9, h:0.75, fontSize:36, bold:true, color:C.gold, align:'center', fontFace:'Calibri' });
    s.addText(c.lbl, { x, y:3.35, w:2.9, h:0.4, fontSize:12, bold:true, color:C.white, align:'center', fontFace:'Calibri', wrap:true });
    s.addText(c.sub, { x:x+0.15, y:3.75, w:2.6, h:0.9, fontSize:10, color:C.muted, align:'center', fontFace:'Calibri', wrap:true });
  });

  // Insight box
  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.2, w:9.0, h:1.6, fill:{ color:'0D2040' }, line:{ color:C.gold, width:1 } });
  s.addText('Root cause: ', { x:0.8, y:5.35, w:1.4, h:0.3, fontSize:11, bold:true, color:C.gold, fontFace:'Calibri' });
  s.addText('MultiBank pays full acquisition cost upfront but delivers zero onboarding structure that moves users from "approved" to "first trade." The 72-hour window closes and the user disappears — permanently.', {
    x:0.8, y:5.65, w:8.5, h:0.9, fontSize:11, color:C.muted, fontFace:'Calibri', wrap:true
  });
}

// ─────────────────────────────────────────────────
// SLIDE 3 — THE INSIGHT
// ─────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE INSIGHT', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.bright, charSpacing:8, fontFace:'Calibri' });
  s.addText('A Cashback Bonus Burns Budget.\nA Milestone Waiver Changes Behaviour.', {
    x:0.6, y:0.6, w:8.8, h:1.0,
    fontSize:26, bold:true, color:C.dark, fontFace:'Calibri'
  });

  // Left col — current
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.85, w:4.2, h:3.4, fill:{ color:'FFF5F5' }, line:{ color:'EF4444', width:1 } });
  s.addText('❌  Current: Sign-Up Bonus', { x:0.7, y:2.0, w:3.8, h:0.4, fontSize:13, bold:true, color:'991B1B', fontFace:'Calibri' });
  const curPts = ['Given upfront, regardless of trading activity','No incentive to actually place a trade','Bonus is spent, user remains inactive','Platform absorbs cost with zero behaviour change'];
  curPts.forEach((pt,i)=>{
    s.addText('• '+pt, { x:0.8, y:2.55+i*0.6, w:3.7, h:0.5, fontSize:11, color:C.dark, fontFace:'Calibri', wrap:true });
  });

  // VS divider
  s.addShape(prs.ShapeType.ellipse, { x:4.55, y:2.95, w:0.9, h:0.9, fill:{ color:C.dark } });
  s.addText('VS', { x:4.55, y:2.98, w:0.9, h:0.85, fontSize:13, bold:true, color:C.white, align:'center', fontFace:'Calibri' });

  // Right col — proposed
  s.addShape(prs.ShapeType.rect, { x:5.0, y:1.85, w:4.4, h:3.4, fill:{ color:'F0FDF8' }, line:{ color:C.green, width:1 } });
  s.addText('✅  TradeRise: Milestone Rewards', { x:5.2, y:2.0, w:4.0, h:0.4, fontSize:13, bold:true, color:'166534', fontFace:'Calibri' });
  const newPts = ['Rewards are earned through real trading actions','Each milestone unlocks a trading-native benefit','Platform only rewards users who are already trading','Behaviour change is built into the incentive structure'];
  newPts.forEach((pt,i)=>{
    s.addText('• '+pt, { x:5.3, y:2.55+i*0.6, w:3.9, h:0.5, fontSize:11, color:C.dark, fontFace:'Calibri', wrap:true });
  });

  // Callout box
  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.45, w:9.0, h:1.4, fill:{ color:C.dark } });
  s.addText('Key insight: The reward only triggers AFTER the user has already placed the trade — meaning the behaviour change happens before any cost is incurred. The milestone reward pays for itself.', {
    x:0.8, y:5.6, w:8.5, h:1.1, fontSize:12, color:C.lgray, fontFace:'Calibri', wrap:true, italic:true
  });
}

// ─────────────────────────────────────────────────
// SLIDE 4 — THE MECHANIC
// ─────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE MECHANIC', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.muted, charSpacing:8, fontFace:'Calibri' });
  s.addText('5 Milestones. 5 Rewards. 30 Days.', {
    x:0.6, y:0.6, w:8.5, h:0.7,
    fontSize:30, bold:true, color:C.white, fontFace:'Calibri'
  });

  // Dashed connector line
  s.addShape(prs.ShapeType.line, { x:0.85, y:2.55, w:8.3, h:0, line:{ color:C.bright, width:1, dashType:'dash', transparency:60 } });

  // 5 steps
  const steps = [
    { num:'01', name:'Identity\nVerified', rwd:'Account\nUnlocked', ico:'🪪' },
    { num:'02', name:'First\nDeposit', rwd:'0.5 pip\nReduction', ico:'💳' },
    { num:'03', name:'First Live\nTrade', rwd:'$25\nRebate', ico:'📈' },
    { num:'04', name:'3 Instruments\nTraded', rwd:'Leverage\nUpgrade', ico:'📊' },
    { num:'05', name:'Stop-Loss\nUsed', rwd:'$50\nBonus', ico:'🛡️' },
  ];
  steps.forEach((st,i)=>{
    const x = 0.6 + i*1.8;
    s.addShape(prs.ShapeType.ellipse, { x, y:2.2, w:0.7, h:0.7, fill:{ color:C.bright } });
    s.addText(st.num, { x, y:2.22, w:0.7, h:0.65, fontSize:11, bold:true, color:C.dark, align:'center', fontFace:'Calibri' });
    s.addText(st.ico, { x:x-0.15, y:3.1, w:1.0, h:0.4, fontSize:20, align:'center' });
    s.addText(st.name, { x:x-0.2, y:3.55, w:1.2, h:0.6, fontSize:10, bold:true, color:C.white, align:'center', fontFace:'Calibri', wrap:true });
    s.addShape(prs.ShapeType.rect, { x:x-0.05, y:4.25, w:1.0, h:0.7, fill:{ color:C.hero }, line:{ color:C.gold, width:0.8 } });
    s.addText(st.rwd, { x:x-0.05, y:4.25, w:1.0, h:0.7, fontSize:9, bold:true, color:C.gold, align:'center', fontFace:'Calibri', wrap:true });
    s.addText('REWARD', { x:x-0.05, y:4.98, w:1.0, h:0.2, fontSize:7, color:C.muted, align:'center', fontFace:'Calibri', charSpacing:4 });
  });

  // Nudge note
  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.6, w:9.0, h:1.3, fill:{ color:'0D2040' }, line:{ color:C.bright, width:0.8 } });
  s.addText('Smart Nudges:', { x:0.8, y:5.72, w:1.6, h:0.3, fontSize:11, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('Push notifications fire at 50%, 80%, and 95% of each milestone with a tailored message — instrument suggestion, market opportunity, or a limited-time spread discount. Ops team dispatches via CRM dashboard.', {
    x:0.8, y:6.05, w:8.5, h:0.7, fontSize:11, color:C.muted, fontFace:'Calibri', wrap:true
  });
}

// ─────────────────────────────────────────────────
// SLIDE 5 — THE PRODUCT
// ─────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE PRODUCT', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.bright, charSpacing:8, fontFace:'Calibri' });
  s.addText('4 Key Screen States — Built & Validated at PhonePe Scale', {
    x:0.6, y:0.6, w:8.5, h:0.65,
    fontSize:24, bold:true, color:C.dark, fontFace:'Calibri'
  });

  const cards = [
    { num:'01', title:'Dashboard Home', desc:'Persistent TradeRise banner shows step 3/5 at 62% progress with 12-day countdown. Visible without navigating — converts passive sessions into action.', mock:'$ 12,450.00\n▲ +1.9% today\n─────────────\n⚡ TRADERISE\nStep 3 of 5 · 62%\n[████████░░]\nExecute first trade\n→ Unlock $25 rebate\n─────────────\nEUR/USD 1.08432\nXAU/USD 2,341.50' },
    { num:'02', title:'Milestone Tracker', desc:'5 sequential goals with live progress bars. Completed steps show earned rewards. Active step shows exact action required. Locked steps create aspiration.', mock:'← ⚡ TradeRise\n30-DAY PROGRAM\n─────────────\n✓ M1 Verified\n✓ M2 Deposited\n● M3 First Trade\n  [████████░░] 62%\n🔒 M4 3 Instruments\n🔒 M5 Stop-Loss\n─────────────\n$75 total available' },
    { num:'03', title:'Milestone Achieved', desc:'Reward unlocked moment — confetti, exact dollar value displayed, auto-apply confirmation. High-emotion success state that drives sharing and next-step momentum.', mock:'⚡⚡⚡⚡⚡\n🎯 ACHIEVED!\nFirst Live Trade\nStep 3 of 5\n─────────────\nYOUR REWARD\n$25.00\nCommission Rebate\n✓ Auto-applied\n─────────────\nNext: 3 Instruments\n→ Leverage Upgrade' },
    { num:'04', title:'Ops Dashboard', desc:'Cohort funnel with per-milestone completion rates. Identifies the M2→M3 drop (68%→23%) as the critical leak. One-click nudge dispatch to 1,847 stuck users.', mock:'TradeRise Ops\n4,231 active · 40%\n─────────────\nM1  ████████ 92%\nM2  ██████   68%\nM3  ██       23% DROP\nM4  █        12%\nM5           6%\n─────────────\n⚡ Send Nudge\n1,847 users at M3' },
  ];
  cards.forEach((c,i)=>{
    const x = 0.3 + i*2.42;
    s.addShape(prs.ShapeType.rect, { x, y:1.45, w:2.3, h:4.9, fill:{ color:C.white }, line:{ color:'E2E8F0', width:1 }, shadow:makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y:1.45, w:2.3, h:0.04, fill:{ color:C.bright } });
    s.addText(c.num, { x, y:1.5, w:2.3, h:0.25, fontSize:8, bold:true, color:C.bright, align:'center', fontFace:'Calibri', charSpacing:4 });
    s.addText(c.title, { x:x+0.1, y:1.75, w:2.1, h:0.45, fontSize:11, bold:true, color:C.dark, fontFace:'Calibri', wrap:true });
    s.addShape(prs.ShapeType.rect, { x:x+0.1, y:2.25, w:2.1, h:2.15, fill:{ color:'F8FAFC' }, line:{ color:'E2E8F0', width:0.5 } });
    s.addText(c.mock, { x:x+0.12, y:2.28, w:2.06, h:2.1, fontSize:7, color:'374151', fontFace:'Courier New', wrap:true });
    s.addText(c.desc, { x:x+0.1, y:4.45, w:2.1, h:1.7, fontSize:9, color:'4B5563', fontFace:'Calibri', wrap:true });
  });
}

// ─────────────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('IMPACT & ROI', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.muted, charSpacing:8, fontFace:'Calibri' });
  s.addText('Projected Impact — Built on PhonePe Proof Points', {
    x:0.6, y:0.6, w:8.5, h:0.7,
    fontSize:28, bold:true, color:C.white, fontFace:'Calibri'
  });

  // Left — User / Growth Impact
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.5, w:4.3, h:0.4, fill:{ color:C.hero } });
  s.addText('USER & GROWTH IMPACT', { x:0.5, y:1.52, w:4.3, h:0.36, fontSize:10, bold:true, color:C.bright, align:'center', fontFace:'Calibri', charSpacing:4 });

  const leftMets = [
    { val:'+40%', lbl:'Funded account conversion\n(approval → first deposit)' },
    { val:'+30%', lbl:'30-day active trader rate\nvs unguided cohort' },
    { val:'−60%', lbl:'Drop-off at M2→M3 step\n(first trade execution)' },
    { val:'+25%', lbl:'Average 90-day trading volume\nper activated trader' },
  ];
  leftMets.forEach((m,i)=>{
    const y = 2.1 + i*1.1;
    s.addShape(prs.ShapeType.rect, { x:0.5, y, w:4.3, h:1.0, fill:{ color:C.hero }, line:{ color:C.bright, width:0.5, transparency:60 } });
    s.addText(m.val, { x:0.6, y:y+0.05, w:1.4, h:0.85, fontSize:32, bold:true, color:C.bright, fontFace:'Calibri', valign:'middle' });
    s.addText(m.lbl, { x:2.1, y:y+0.1, w:2.5, h:0.8, fontSize:11, color:C.muted, fontFace:'Calibri', wrap:true, valign:'middle' });
  });

  // Right — Platform ROI
  s.addShape(prs.ShapeType.rect, { x:5.2, y:1.5, w:4.3, h:0.4, fill:{ color:'0D2040' }, line:{ color:C.gold, width:1 } });
  s.addText('PLATFORM ROI', { x:5.2, y:1.52, w:4.3, h:0.36, fontSize:10, bold:true, color:C.gold, align:'center', fontFace:'Calibri', charSpacing:4 });

  const rightMets = [
    { val:'−25%', lbl:'Blended CAC per\nactive trader acquired' },
    { val:'2.4×', lbl:'LTV improvement\nfor milestone completers vs control' },
    { val:'$74K', lbl:'Rebates issued per 4,231\nusers — self-funding at 0.6% GMV cost' },
    { val:'+18%', lbl:'D90 retention for users\nwho complete all 5 milestones' },
  ];
  rightMets.forEach((m,i)=>{
    const y = 2.1 + i*1.1;
    s.addShape(prs.ShapeType.rect, { x:5.2, y, w:4.3, h:1.0, fill:{ color:'0D2040' }, line:{ color:C.gold, width:0.5, transparency:60 } });
    s.addText(m.val, { x:5.3, y:y+0.05, w:1.4, h:0.85, fontSize:32, bold:true, color:C.gold, fontFace:'Calibri', valign:'middle' });
    s.addText(m.lbl, { x:6.8, y:y+0.1, w:2.5, h:0.8, fontSize:11, color:C.muted, fontFace:'Calibri', wrap:true, valign:'middle' });
  });
}

// ─────────────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK
// ─────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('PROOF OF WORK', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.bright, charSpacing:8, fontFace:'Calibri' });
  s.addText('I Built This. Here Is the Proof.', {
    x:0.6, y:0.6, w:8.5, h:0.65,
    fontSize:28, bold:true, color:C.dark, fontFace:'Calibri'
  });

  // Left — PhonePe
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.45, w:4.3, h:4.4, fill:{ color:C.dark } });
  s.addText('What I Built at PhonePe', { x:0.7, y:1.65, w:3.9, h:0.35, fontSize:12, bold:true, color:C.bright, fontFace:'Calibri' });
  const ppPts = [
    'Built Milestone Activation for 5,000+ B2B device merchants — same 5-step progressive reward architecture as TradeRise',
    'Designed 6 screen states: goal tracking, progress bars, success confetti, error fallback, nudge campaigns',
    'A/B tested reward mechanics: behaviour-linked waiver vs. flat cashback — waiver drove 34% higher goal completion',
    'Built the Nudges engine: ops-facing CRM tool for threshold-based push dispatch at 50/80/95% completion',
    'Result: 60% improvement in 90-day device activation rate and 23% lower CAC vs prior manual flow',
  ];
  ppPts.forEach((pt,i)=>{
    s.addText('→  '+pt, { x:0.7, y:2.15+i*0.75, w:3.9, h:0.65, fontSize:10, color:'CBD5E1', fontFace:'Calibri', wrap:true });
  });

  // Right — JD mapping
  s.addShape(prs.ShapeType.rect, { x:5.1, y:1.45, w:4.3, h:4.4, fill:{ color:'F1F5F9' }, line:{ color:'E2E8F0', width:1 } });
  s.addText('How It Maps to This Role', { x:5.3, y:1.65, w:3.9, h:0.35, fontSize:12, bold:true, color:C.dark, fontFace:'Calibri' });
  const jdPts = [
    'JD: "Design & deliver referral and loyalty programs" → Built milestone loyalty program end-to-end at 350M MAU scale',
    'JD: "Identify growth loops and incentive strategies" → A/B tested behaviour-linked vs flat reward — chose the one that changed behaviour',
    'JD: "Conduct A/B testing and leverage behavioral insights" → Ran controlled A/B at PhonePe, owns the data and methodology',
    'JD: "Oversee marketing automation and CRM" → Built the nudge engine that fires at 50/80/95% milestone completion',
    'JD: "Analyze funnel metrics (CAC, LTV, conversion rates)" → Directly measured M2→M3 drop and attributed it to absence of structured nudges',
  ];
  jdPts.forEach((pt,i)=>{
    s.addText('✓  '+pt, { x:5.3, y:2.15+i*0.75, w:3.9, h:0.65, fontSize:10, color:'374151', fontFace:'Calibri', wrap:true });
  });

  // Quote box
  s.addShape(prs.ShapeType.rect, { x:0.5, y:6.05, w:9.0, h:1.0, fill:{ color:C.dark } });
  s.addText('"TradeRise is not a concept. It is the same activation architecture I shipped at PhonePe — now applied to a trading platform. I know exactly which metric to optimise and where the funnel breaks."', {
    x:0.8, y:6.15, w:8.5, h:0.8, fontSize:11, color:C.lgray, italic:true, fontFace:'Calibri', wrap:true
  });
}

// ─────────────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN
// ─────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('ROLLOUT PLAN', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.muted, charSpacing:8, fontFace:'Calibri' });
  s.addText('Phased Launch — 6 Months to Full Scale', {
    x:0.6, y:0.6, w:8.5, h:0.65,
    fontSize:28, bold:true, color:C.white, fontFace:'Calibri'
  });

  const phases = [
    { ph:'Phase 1', time:'Month 1–2', title:'Pilot (500 Users)', color:C.bright, pts:['Deploy TradeRise to 500 newly-approved accounts in one market','Instrument milestone completion rate, time-to-trade, and reward cost','Compare vs unguided control group — measure M2→M3 conversion lift'] },
    { ph:'Phase 2', time:'Month 3–4', title:'A/B Scale + Nudge Engine', color:C.gold, pts:['Run reward mechanic A/B: commission rebate vs spread discount vs leverage upgrade','Launch ops Nudge dashboard — enable team to dispatch at 50/80/95% thresholds','Expand to 5,000 users across 3 markets based on pilot learnings'] },
    { ph:'Phase 3', time:'Month 5–6', title:'Full Rollout + Automation', color:C.green, pts:['Roll out winning milestone variant to all new account approvals globally','Integrate TradeRise into standard onboarding flow — no opt-in required','Automate nudge sequences via CRM — zero ops overhead per campaign'] },
  ];
  phases.forEach((ph,i)=>{
    const x = 0.4 + i*3.2;
    s.addShape(prs.ShapeType.rect, { x, y:1.5, w:3.0, h:4.1, fill:{ color:C.hero }, line:{ color:ph.color, width:1 } });
    s.addShape(prs.ShapeType.rect, { x, y:1.5, w:3.0, h:0.06, fill:{ color:ph.color } });
    s.addText(ph.ph, { x:x+0.15, y:1.6, w:2.7, h:0.3, fontSize:9, bold:true, color:ph.color, fontFace:'Calibri', charSpacing:4 });
    s.addText(ph.time, { x:x+0.15, y:1.9, w:2.7, h:0.25, fontSize:10, color:C.muted, fontFace:'Calibri' });
    s.addText(ph.title, { x:x+0.15, y:2.2, w:2.7, h:0.5, fontSize:14, bold:true, color:C.white, fontFace:'Calibri', wrap:true });
    ph.pts.forEach((pt,j)=>{
      s.addText('• '+pt, { x:x+0.15, y:2.85+j*0.85, w:2.7, h:0.75, fontSize:10, color:C.muted, fontFace:'Calibri', wrap:true });
    });
  });

  // What I need
  s.addShape(prs.ShapeType.rect, { x:0.4, y:5.8, w:6.4, h:1.35, fill:{ color:'0D2040' }, line:{ color:C.bright, width:1 } });
  s.addText('What I need to build this:', { x:0.65, y:5.93, w:5.8, h:0.3, fontSize:11, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('Access to funnel analytics (approved → deposited → first trade)  ·  CRM / push notification infrastructure  ·  1 engineer for milestone tracking API  ·  Alignment with compliance on reward mechanics', {
    x:0.65, y:6.25, w:5.8, h:0.7, fontSize:10, color:C.muted, fontFace:'Calibri', wrap:true
  });

  // Contact
  s.addShape(prs.ShapeType.rect, { x:7.05, y:5.8, w:2.5, h:1.35, fill:{ color:C.bright } });
  s.addText('Ajay Avaghade', { x:7.1, y:5.95, w:2.4, h:0.3, fontSize:12, bold:true, color:C.dark, align:'center', fontFace:'Calibri' });
  s.addText('avaghadeajay009@gmail.com', { x:7.1, y:6.3, w:2.4, h:0.25, fontSize:9, color:C.dark, align:'center', fontFace:'Calibri' });
  s.addText('+91 9561558439', { x:7.1, y:6.6, w:2.4, h:0.25, fontSize:9, color:C.dark, align:'center', fontFace:'Calibri' });
}

prs.writeFile({ fileName:'assets/multibank-traderise-deck.pptx' }).then(()=>console.log('multibank-traderise-deck.pptx done'));
