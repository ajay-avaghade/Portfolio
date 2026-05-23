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
    s.addShape(prs.ShapeType.line, { x:0+i*1.8, y:0, w:0, h:7.5, line:{ color:C.gold, width:0.4, transparency:88 } });
  }

  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:10, h:0.06, fill:{ color:C.gold } });

  s.addText('MULTIBANK GROUP  ·  GROWTH MARKETING PM', {
    x:0.6, y:0.3, w:8, h:0.3,
    fontSize:9, bold:true, color:C.gold, charSpacing:8, fontFace:'Calibri'
  });

  s.addText('TradeCircle', {
    x:0.6, y:1.0, w:7, h:1.2,
    fontSize:68, bold:true, color:C.white, fontFace:'Calibri'
  });

  s.addText('Solving the Guilt Problem in Trading Referrals', {
    x:0.6, y:2.25, w:7.5, h:0.5,
    fontSize:22, color:C.gold, fontFace:'Calibri', italic:true
  });

  s.addShape(prs.ShapeType.rect, { x:0.6, y:2.85, w:1.4, h:0.05, fill:{ color:C.bright } });

  s.addText('Presented by Ajay Avaghade  ·  Growth & Monetization PM', {
    x:0.6, y:3.05, w:8, h:0.3,
    fontSize:12, color:C.muted, fontFace:'Calibri'
  });

  // Big stat
  s.addShape(prs.ShapeType.rect, { x:7.0, y:5.2, w:2.7, h:1.8, fill:{ color:C.hero }, line:{ color:C.gold, width:1 } });
  s.addText('~0×', { x:7.0, y:5.3, w:2.7, h:0.7, fontSize:42, bold:true, color:C.gold, align:'center', fontFace:'Calibri' });
  s.addText('viral coefficient for\nmost trading platforms', { x:7.0, y:6.0, w:2.7, h:0.8, fontSize:11, color:C.muted, align:'center', fontFace:'Calibri' });

  s.addShape(prs.ShapeType.rect, { x:0, y:7.44, w:10, h:0.06, fill:{ color:C.bright } });
}

// ─────────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE PROBLEM', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.muted, charSpacing:8, fontFace:'Calibri' });
  s.addText('Standard Trading Referrals Fail Because of Guilt.\nNobody Wants to Be Responsible for a Friend\'s Loss.', {
    x:0.6, y:0.6, w:8.5, h:1.1,
    fontSize:26, bold:true, color:C.white, fontFace:'Calibri'
  });

  const cols3 = [
    { ico:'😬', stat:'~0×', lbl:'Viral coefficient', sub:'Trading platforms have near-zero organic referrals. Nobody shares financial products with friends.' },
    { ico:'💸', stat:'$100+', lbl:'Blended CAC (paid only)', sub:'Without viral growth, every user is acquired through expensive paid channels — a structural ceiling.' },
    { ico:'🤫', stat:'73%', lbl:'Users who hide trading activity', sub:'Research shows most retail traders do not tell friends or family they trade CFDs or forex.' },
  ];
  cols3.forEach((c,i)=>{
    const x = 0.5 + i*3.1;
    s.addShape(prs.ShapeType.rect, { x, y:1.9, w:2.9, h:3.0, fill:{ color:C.hero }, line:{ color:C.gold, width:0.5, transparency:60 }, shadow:makeShadow() });
    s.addText(c.ico, { x, y:2.05, w:2.9, h:0.5, fontSize:24, align:'center' });
    s.addText(c.stat, { x, y:2.6, w:2.9, h:0.75, fontSize:36, bold:true, color:C.gold, align:'center', fontFace:'Calibri' });
    s.addText(c.lbl, { x, y:3.35, w:2.9, h:0.4, fontSize:12, bold:true, color:C.white, align:'center', fontFace:'Calibri', wrap:true });
    s.addText(c.sub, { x:x+0.15, y:3.75, w:2.6, h:0.9, fontSize:10, color:C.muted, align:'center', fontFace:'Calibri', wrap:true });
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.2, w:9.0, h:1.6, fill:{ color:'0D2040' }, line:{ color:C.gold, width:1 } });
  s.addText('Root cause: ', { x:0.8, y:5.35, w:1.4, h:0.3, fontSize:11, bold:true, color:C.gold, fontFace:'Calibri' });
  s.addText('If I refer a friend to a CFD trading platform and they lose money, that guilt falls on me. Standard referral mechanics ("share this link, get $50") completely ignore this psychological barrier. No wonder viral coefficient is near-zero.', {
    x:0.8, y:5.65, w:8.5, h:0.9, fontSize:11, color:C.muted, fontFace:'Calibri', wrap:true
  });
}

// ─────────────────────────────────────────────────
// SLIDE 3 — THE INSIGHT
// ─────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE INSIGHT', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.gold, charSpacing:8, fontFace:'Calibri' });
  s.addText('Remove the Guilt. Make Rewards Success-Linked.\nBoth Sides Only Win When the Friend Profits.', {
    x:0.6, y:0.6, w:8.8, h:1.0,
    fontSize:24, bold:true, color:C.dark, fontFace:'Calibri'
  });

  // Left — current
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.85, w:4.2, h:3.4, fill:{ color:'FFF5F5' }, line:{ color:'EF4444', width:1 } });
  s.addText('❌  Current: Link Referral ($50 bonus)', { x:0.7, y:2.0, w:3.8, h:0.4, fontSize:13, bold:true, color:'991B1B', fontFace:'Calibri' });
  const curPts = ['Bonus paid on signup — before any trading','Friend loses money → referrer feels guilty','Referrer goes silent, stops sharing the platform','No alignment: referrer earns even if friend loses'];
  curPts.forEach((pt,i)=>{
    s.addText('• '+pt, { x:0.8, y:2.55+i*0.6, w:3.7, h:0.5, fontSize:11, color:C.dark, fontFace:'Calibri', wrap:true });
  });

  // VS
  s.addShape(prs.ShapeType.ellipse, { x:4.55, y:2.95, w:0.9, h:0.9, fill:{ color:C.dark } });
  s.addText('VS', { x:4.55, y:2.98, w:0.9, h:0.85, fontSize:13, bold:true, color:C.white, align:'center', fontFace:'Calibri' });

  // Right — proposed
  s.addShape(prs.ShapeType.rect, { x:5.0, y:1.85, w:4.4, h:3.4, fill:{ color:'FFFBF0' }, line:{ color:C.gold, width:1 } });
  s.addText('✅  TradeCircle: Performance Referral', { x:5.2, y:2.0, w:4.0, h:0.4, fontSize:13, bold:true, color:'92400E', fontFace:'Calibri' });
  const newPts = ['Share an anonymised % return card — no dollar P&L exposed','Reward only triggers when the referred friend profits','Both sides win together — guilt is structurally removed','Referrer is incentivised to share advice, not just a link'];
  newPts.forEach((pt,i)=>{
    s.addText('• '+pt, { x:5.3, y:2.55+i*0.6, w:3.9, h:0.5, fontSize:11, color:C.dark, fontFace:'Calibri', wrap:true });
  });

  // Callout
  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.45, w:9.0, h:1.4, fill:{ color:C.dark } });
  s.addText('Key insight: By making rewards success-linked, we eliminate the guilt that kills trading referrals. The viral unit is not a link — it is a beautifully designed performance card that the referrer is proud to share.', {
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
  s.addText('Share Your Win. Build a Circle. Earn Together.', {
    x:0.6, y:0.6, w:8.5, h:0.7,
    fontSize:30, bold:true, color:C.white, fontFace:'Calibri'
  });

  s.addShape(prs.ShapeType.line, { x:0.85, y:2.55, w:8.3, h:0, line:{ color:C.gold, width:1, dashType:'dash', transparency:60 } });

  const steps = [
    { num:'01', name:'Week ends\nReturn posted', ico:'📊', desc:'App auto-generates a performance card' },
    { num:'02', name:'User shares\nthe card', ico:'📤', desc:'WhatsApp, Instagram, or direct link' },
    { num:'03', name:'Friend sees\nthe card', ico:'👀', desc:'Anonymised % return, instrument, style' },
    { num:'04', name:'Friend joins\nthe circle', ico:'🔵', desc:'Creates account using referral code' },
    { num:'05', name:'Friend profits\nboth earn', ico:'🎉', desc:'Spread rebates trigger on first profitable trade' },
  ];
  steps.forEach((st,i)=>{
    const x = 0.6 + i*1.8;
    s.addShape(prs.ShapeType.ellipse, { x, y:2.2, w:0.7, h:0.7, fill:{ color:C.gold } });
    s.addText(st.num, { x, y:2.22, w:0.7, h:0.65, fontSize:11, bold:true, color:C.dark, align:'center', fontFace:'Calibri' });
    s.addText(st.ico, { x:x-0.15, y:3.1, w:1.0, h:0.4, fontSize:20, align:'center' });
    s.addText(st.name, { x:x-0.2, y:3.55, w:1.2, h:0.6, fontSize:10, bold:true, color:C.white, align:'center', fontFace:'Calibri', wrap:true });
    s.addShape(prs.ShapeType.rect, { x:x-0.1, y:4.25, w:1.1, h:0.7, fill:{ color:C.hero }, line:{ color:C.bright, width:0.5 } });
    s.addText(st.desc, { x:x-0.1, y:4.28, w:1.1, h:0.65, fontSize:8, color:C.muted, align:'center', fontFace:'Calibri', wrap:true });
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.6, w:9.0, h:1.3, fill:{ color:'0D2040' }, line:{ color:C.gold, width:0.8 } });
  s.addText('The viral unit:', { x:0.8, y:5.72, w:1.3, h:0.3, fontSize:11, bold:true, color:C.gold, fontFace:'Calibri' });
  s.addText('The performance card is designed to be shared — anonymised % return (never dollar P&L), top instrument, and trading style. It signals success without exposing vulnerability. Network effects compound: as circles grow, platform stickiness and DAU both increase.', {
    x:0.8, y:6.05, w:8.5, h:0.7, fontSize:11, color:C.muted, fontFace:'Calibri', wrap:true
  });
}

// ─────────────────────────────────────────────────
// SLIDE 5 — THE PRODUCT
// ─────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE PRODUCT', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.gold, charSpacing:8, fontFace:'Calibri' });
  s.addText('4 Screens — Built Around the Referral Psychology Insight', {
    x:0.6, y:0.6, w:8.5, h:0.65,
    fontSize:24, bold:true, color:C.dark, fontFace:'Calibri'
  });

  const cards = [
    { num:'01', title:'Home + Circle Card', desc:'TradeCircle card shows week return (+4.2%), top instrument, 0 current members, and one-tap share CTA. Visible on home screen — no buried menu.', mock:'MultiBank\n$ 12,450.00\n────────────\n🔵 TRADECIRCLE\nThis week: +4.2%\nBest: XAU/USD\n0 members\n[📤 Share Card]\nEarn rebates when\nyour circle profits\n────────────\nXAU/USD 2,341.50' },
    { num:'02', title:'Performance Card', desc:'The shareable viral unit. Anonymised username, % return (not $ P&L), top instrument, trading style badge. Referral code embedded. Designed to feel like a win worth sharing.', mock:'MultiBank TradeCircle\n@TradeMaster_AK\n─────────────\nTHIS WEEK RETURN\n+4.2%\n─────────────\n[XAU/USD][Swing]\n[Dubai][14-mo]\n─────────────\nJoin my circle:\nCIRCLE-AK42\n─────────────\n% gains · Not advice' },
    { num:'03', title:'Friend Onboarding', desc:'Referred friend sees leader stats, how rewards work, and a suggested first trade matching the circle leader\'s instrument. No guilt framing — only "earn together" messaging.', mock:'🔵 CIRCLE INVITE\nYou\'re joining\n@TradeMaster_AK\nLast 7 days: +4.2%\nTop: XAU/USD\n─────────────\nHow it works:\n✓ You profit = both earn\n✓ Rewards on profit only\n─────────────\n[Join Circle]' },
    { num:'04', title:'Viral Ops Dashboard', desc:'Viral coefficient live tracker (1.8×, target 2.5×). Referral funnel by step. Insight: circles where leader made +5%+ convert 3.2× better. Enables growth team to surface high-performers.', mock:'TradeCircle Ops\n1.8× Viral Coeff\n12.4K Circles\n$18 CAC\n─────────────\nFunnel:\nShared    100%\nClicked    45%\nSignup     28%\nTraded     14%\nRebate      6%\n─────────────\nInsight: >5% return\n→ 3.2× conversion' },
  ];
  cards.forEach((c,i)=>{
    const x = 0.3 + i*2.42;
    s.addShape(prs.ShapeType.rect, { x, y:1.45, w:2.3, h:4.9, fill:{ color:C.white }, line:{ color:'E2E8F0', width:1 }, shadow:makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y:1.45, w:2.3, h:0.04, fill:{ color:C.gold } });
    s.addText(c.num, { x, y:1.5, w:2.3, h:0.25, fontSize:8, bold:true, color:C.gold, align:'center', fontFace:'Calibri', charSpacing:4 });
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
  s.addText('Projected Impact — Built on PhonePe Referral Engine Proof', {
    x:0.6, y:0.6, w:8.5, h:0.7,
    fontSize:28, bold:true, color:C.white, fontFace:'Calibri'
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.5, w:4.3, h:0.4, fill:{ color:C.hero } });
  s.addText('GROWTH IMPACT', { x:0.5, y:1.52, w:4.3, h:0.36, fontSize:10, bold:true, color:C.gold, align:'center', fontFace:'Calibri', charSpacing:4 });

  const leftMets = [
    { val:'3×', lbl:'Viral coefficient\nvs near-zero baseline' },
    { val:'−35%', lbl:'Blended CAC\nvs paid-only acquisition' },
    { val:'+25%', lbl:'D30 retention for\nreferred users vs organic' },
    { val:'45%', lbl:'Card-to-click rate\non anonymised performance cards' },
  ];
  leftMets.forEach((m,i)=>{
    const y = 2.1 + i*1.1;
    s.addShape(prs.ShapeType.rect, { x:0.5, y, w:4.3, h:1.0, fill:{ color:C.hero }, line:{ color:C.gold, width:0.5, transparency:60 } });
    s.addText(m.val, { x:0.6, y:y+0.05, w:1.4, h:0.85, fontSize:32, bold:true, color:C.gold, fontFace:'Calibri', valign:'middle' });
    s.addText(m.lbl, { x:2.1, y:y+0.1, w:2.5, h:0.8, fontSize:11, color:C.muted, fontFace:'Calibri', wrap:true, valign:'middle' });
  });

  s.addShape(prs.ShapeType.rect, { x:5.2, y:1.5, w:4.3, h:0.4, fill:{ color:'0D2040' }, line:{ color:C.bright, width:1 } });
  s.addText('PLATFORM ROI', { x:5.2, y:1.52, w:4.3, h:0.36, fontSize:10, bold:true, color:C.bright, align:'center', fontFace:'Calibri', charSpacing:4 });

  const rightMets = [
    { val:'$18', lbl:'Blended CAC per\nreferred user (vs $100+ paid)' },
    { val:'2.1×', lbl:'LTV for circle members\nvs solo-acquired users' },
    { val:'12.4K', lbl:'Active circles growing\nat 1.8× viral coefficient' },
    { val:'+14pt', lbl:'NPS improvement from\n"earning together" network effect' },
  ];
  rightMets.forEach((m,i)=>{
    const y = 2.1 + i*1.1;
    s.addShape(prs.ShapeType.rect, { x:5.2, y, w:4.3, h:1.0, fill:{ color:'0D2040' }, line:{ color:C.bright, width:0.5, transparency:60 } });
    s.addText(m.val, { x:5.3, y:y+0.05, w:1.4, h:0.85, fontSize:32, bold:true, color:C.bright, fontFace:'Calibri', valign:'middle' });
    s.addText(m.lbl, { x:6.8, y:y+0.1, w:2.5, h:0.8, fontSize:11, color:C.muted, fontFace:'Calibri', wrap:true, valign:'middle' });
  });
}

// ─────────────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK
// ─────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('PROOF OF WORK', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.gold, charSpacing:8, fontFace:'Calibri' });
  s.addText('I Built a Referral Engine at Scale. Here Is What I Shipped.', {
    x:0.6, y:0.6, w:8.5, h:0.65,
    fontSize:26, bold:true, color:C.dark, fontFace:'Calibri'
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.45, w:4.3, h:4.4, fill:{ color:C.dark } });
  s.addText('What I Built at PhonePe', { x:0.7, y:1.65, w:3.9, h:0.35, fontSize:12, bold:true, color:C.gold, fontFace:'Calibri' });
  const ppPts = [
    'Built zero-to-one multi-tenant referral engine for B2B merchant acquisition — 5Mn+ new users/month at 23% lower CAC',
    'Designed performance-linked incentive structure: rewards tied to actual business outcome (transactions), not just sign-up event',
    'Built audience segmentation and cohort tooling for 500+ brand partners — same segmentation logic as TradeCircle circle analytics',
    'Owned the full funnel: share → click → sign-up → first transaction → rebate trigger — measured at each drop-off step',
    'Rebuilt a static rewards portfolio into an ML-driven engagement system with personalised offer matching',
  ];
  ppPts.forEach((pt,i)=>{
    s.addText('→  '+pt, { x:0.7, y:2.15+i*0.75, w:3.9, h:0.65, fontSize:10, color:'CBD5E1', fontFace:'Calibri', wrap:true });
  });

  s.addShape(prs.ShapeType.rect, { x:5.1, y:1.45, w:4.3, h:4.4, fill:{ color:'F1F5F9' }, line:{ color:'E2E8F0', width:1 } });
  s.addText('How It Maps to This Role', { x:5.3, y:1.65, w:3.9, h:0.35, fontSize:12, bold:true, color:C.dark, fontFace:'Calibri' });
  const jdPts = [
    'JD: "Proven track record in designing and scaling referral programs" → Shipped a referral engine driving 5Mn+ users/month at 23% lower CAC',
    'JD: "Identify growth loops, viral mechanics, and network effects" → TradeCircle\'s "earn together" loop is a compounding network effect by design',
    'JD: "Understanding gamification and reward systems" → A/B tested reward mechanics, chose performance-linked over flat bonus based on data',
    'JD: "Analyze funnel metrics: CAC, LTV, conversion rates" → Owned full funnel attribution from share event to rebate trigger at PhonePe',
    'JD: "Marketing automation and CRM strategies" → Built automated nudge sequences keyed to milestone completion rates',
  ];
  jdPts.forEach((pt,i)=>{
    s.addText('✓  '+pt, { x:5.3, y:2.15+i*0.75, w:3.9, h:0.65, fontSize:10, color:'374151', fontFace:'Calibri', wrap:true });
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:6.05, w:9.0, h:1.0, fill:{ color:C.dark } });
  s.addText('"I built a referral platform that turned a 2-day manual process into a 30-minute self-serve flow for 5,000+ merchants. TradeCircle applies the same performance-linked incentive logic to a consumer trading context."', {
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
  s.addText('Phased Launch — From Pilot Circle to Viral Engine', {
    x:0.6, y:0.6, w:8.5, h:0.65,
    fontSize:28, bold:true, color:C.white, fontFace:'Calibri'
  });

  const phases = [
    { ph:'Phase 1', time:'Month 1–2', title:'Pilot (500 Users)', color:C.gold, pts:['Deploy to 500 active traders — measure share rate, click rate, and referred signup conversion','A/B test card designs: % return only vs % return + instrument badge','Measure guilt metric proxy: do referrers share more on winning weeks?'] },
    { ph:'Phase 2', time:'Month 3–4', title:'Reward Mechanic Test', color:C.bright, pts:['Test success-linked rebate vs flat sign-up bonus for both referrer and referred user','Instrument viral coefficient at each funnel step (share → click → KYC → trade → rebate)','Identify top circle leaders — build a leaderboard to surface high-performing traders'] },
    { ph:'Phase 3', time:'Month 5–6', title:'Full Scale + Network Effect', color:C.green, pts:['Roll out to all active traders globally as a standard in-app feature','Launch weekly performance card push notification to all users — drive sharing cadence','Add circle leaderboard and social proof mechanics to compound network effect'] },
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

  s.addShape(prs.ShapeType.rect, { x:0.4, y:5.8, w:6.4, h:1.35, fill:{ color:'0D2040' }, line:{ color:C.gold, width:1 } });
  s.addText('What I need to build this:', { x:0.65, y:5.93, w:5.8, h:0.3, fontSize:11, bold:true, color:C.gold, fontFace:'Calibri' });
  s.addText('Access to user trading P&L data (anonymised for card)  ·  Push notification / social share infrastructure  ·  1 engineer for referral tracking & rebate trigger API  ·  Compliance sign-off on performance card disclosure copy', {
    x:0.65, y:6.25, w:5.8, h:0.7, fontSize:10, color:C.muted, fontFace:'Calibri', wrap:true
  });

  s.addShape(prs.ShapeType.rect, { x:7.05, y:5.8, w:2.5, h:1.35, fill:{ color:C.gold } });
  s.addText('Ajay Avaghade', { x:7.1, y:5.95, w:2.4, h:0.3, fontSize:12, bold:true, color:C.dark, align:'center', fontFace:'Calibri' });
  s.addText('avaghadeajay009@gmail.com', { x:7.1, y:6.3, w:2.4, h:0.25, fontSize:9, color:C.dark, align:'center', fontFace:'Calibri' });
  s.addText('+91 9561558439', { x:7.1, y:6.6, w:2.4, h:0.25, fontSize:9, color:C.dark, align:'center', fontFace:'Calibri' });
}

prs.writeFile({ fileName:'assets/multibank-tradecircle-deck.pptx' }).then(()=>console.log('multibank-tradecircle-deck.pptx done'));
