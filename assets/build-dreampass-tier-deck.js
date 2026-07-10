const PptxGenJS = require('pptxgenjs');

(async () => {
  const prs = new PptxGenJS();
  prs.layout = 'LAYOUT_16x9';

  const C = {
    dark:  '130E01',
    hero:  '291A04',
    accent:'D97706',
    amber: 'B45309',
    white: 'FFFFFF',
    lgray: 'FFFBEB',
    muted: '94A3B8',
    gold:  'FCD34D',
    green: '10B981',
  };
  const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.14 });

  // SLIDE 1 — COVER
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.04, fill:{ color:C.accent } });
    s.addShape('rect', { x:0, y:7.46, w:'100%', h:0.04, fill:{ color:C.accent } });
    for(let i=0;i<6;i++){ s.addShape('line', { x:6.5+i*0.45, y:0, w:0, h:7.5, line:{ color:C.accent, width:0.5, transparency:85 } }); }
    s.addText('DREAMPASS · TREVOLUTION · CASE STUDY D', { x:0.5, y:0.3, w:9, h:0.25, color:C.muted, fontSize:8, bold:true, charSpacing:2 });
    s.addText('DreamTier', { x:0.5, y:1.0, w:9, h:1.4, color:C.white, fontSize:52, bold:true, fontFace:'Calibri' });
    s.addText('Freemium-to-Paid Upgrade Engine — Converting Free Members at the Moment of Travel Intent', { x:0.5, y:2.3, w:8.5, h:0.5, color:C.accent, fontSize:18, fontFace:'Calibri' });
    s.addText('Ajay Avaghade · Product Manager', { x:0.5, y:3.0, w:5, h:0.3, color:C.muted, fontSize:12 });
    s.addShape('rect', { x:6.8, y:1.8, w:2.8, h:1.4, fill:{ color:C.accent, transparency:88 }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
    s.addText('2.8×', { x:6.8, y:1.9, w:2.8, h:0.7, color:C.white, fontSize:36, bold:true, align:'center' });
    s.addText('ARPU lift: paid vs free member over 12 months', { x:6.8, y:2.5, w:2.8, h:0.5, color:C.accent, fontSize:10, align:'center' });
  }

  // SLIDE 2 — THE PROBLEM
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:0.06, h:'100%', fill:{ color:C.accent } });
    s.addText('THE PROBLEM', { x:0.3, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Free Members Experience the App, Book Trips, and Leave Without Ever Upgrading to Paid', { x:0.3, y:0.55, w:9.4, h:0.9, color:C.white, fontSize:23, bold:true, fontFace:'Calibri' });
    const cols = [
      { icon:'🔒', val:'4 of 6', lbl:'Benefits locked on free tier', sub:'Lounge access, eSIM, hotel cashback, and insurance are invisible until the user actively seeks an upgrade page' },
      { icon:'📉', val:'1×', lbl:'Coin earn rate on free tier', sub:'Free members earn at half the rate of paid members — but they don\'t know this until they\'re deep in the app' },
      { icon:'🕳️', val:'0%', lbl:'Contextual upgrade prompts today', sub:'No upgrade surface at the highest-intent moment: the booking confirmation page where lounge access is most relevant' },
    ];
    cols.forEach((c,i)=>{
      const x = 0.3 + i*3.2;
      s.addShape('rect', { x, y:1.65, w:3.0, h:2.6, fill:{ color:C.hero }, line:{ color:C.accent, width:1, transparency:60 }, shadow:makeShadow() });
      s.addText(c.icon, { x, y:1.75, w:3.0, h:0.5, fontSize:22, align:'center' });
      s.addText(c.val, { x, y:2.2, w:3.0, h:0.6, color:C.accent, fontSize:28, bold:true, align:'center' });
      s.addText(c.lbl, { x, y:2.75, w:3.0, h:0.3, color:C.white, fontSize:11, bold:true, align:'center' });
      s.addText(c.sub, { x:x+0.1, y:3.1, w:2.8, h:0.95, color:C.muted, fontSize:9.5 });
    });
    s.addShape('rect', { x:0.3, y:4.4, w:9.4, h:0.9, fill:{ color:C.accent, transparency:90 }, line:{ color:C.accent, width:1 } });
    s.addText('Root cause: The upgrade pitch is generic and untimed. Free members get a static upgrade page with a feature list — not a contextual "you are about to fly and your lounge is waiting" moment at peak intent.', { x:0.5, y:4.5, w:9.0, h:0.7, color:C.white, fontSize:10, italic:true });
  }

  // SLIDE 3 — THE INSIGHT
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.lgray } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('THE INSIGHT', { x:0.4, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('The Upgrade Moment Is Not a Settings Page. It\'s the Second After a Booking Is Confirmed.', { x:0.4, y:0.55, w:9.2, h:0.8, color:C.dark, fontSize:24, bold:true, fontFace:'Calibri' });
    s.addShape('rect', { x:0.4, y:1.5, w:4.4, h:3.3, fill:{ color:'FEF2F2' }, line:{ color:'FECACA', width:1 }, shadow:makeShadow() });
    s.addText('❌  Without DreamTier', { x:0.55, y:1.65, w:4.1, h:0.35, color:'991B1B', fontSize:11, bold:true });
    ['Generic "Upgrade" button buried in account settings','Feature list paywall with no contextual relevance','Free member books Dubai flight — no lounge mention','1× coin earn rate — member doesn\'t know they\'re missing half','No in-destination benefit reinforcement after upgrade'].forEach((t,i)=>{ s.addText(`• ${t}`, { x:0.6, y:2.1+i*0.42, w:4.1, h:0.38, color:'7F1D1D', fontSize:10 }); });
    s.addShape('rect', { x:5.1, y:1.5, w:4.6, h:3.3, fill:{ color:'FFFBEB' }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
    s.addText('✅  With DreamTier', { x:5.25, y:1.65, w:4.2, h:0.35, color:C.amber, fontSize:11, bold:true });
    ['Booking confirmed → "DXB lounge is available on your departure — unlock with Dreampass Paid"','Value-at-stake paywall: pay AED 499 / get AED 1,320 back','2× coin earn applied immediately — visible on booking screen','eSIM for destination auto-teased on booking confirmation','In-destination use of lounge → 76% renewal rate (vs 38% no-use)'].forEach((t,i)=>{ s.addText(`• ${t}`, { x:5.25, y:2.1+i*0.42, w:4.3, h:0.38, color:C.amber, fontSize:10 }); });
    s.addShape('rect', { x:0.4, y:5.0, w:9.2, h:0.7, fill:{ color:C.accent, transparency:85 }, line:{ color:C.accent, width:1 } });
    s.addText('Key insight: Paywalls that show "pay AED 499, get back AED 1,320" outperform feature-list paywalls by 31% on conversion. The ROI frame eliminates price anchoring — the user isn\'t paying for a subscription, they\'re making a 2.6× return investment.', { x:0.6, y:5.08, w:8.8, h:0.55, color:C.dark, fontSize:10, italic:true });
  }

  // SLIDE 4 — THE MECHANIC
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.hero } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.gold } });
    s.addText('THE MECHANIC', { x:0.4, y:0.25, w:9, h:0.25, color:C.gold, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Four Behavioural Triggers + Value-at-Stake Paywall Design', { x:0.4, y:0.55, w:9, h:0.6, color:C.white, fontSize:26, bold:true, fontFace:'Calibri' });
    const steps = [
      { icon:'✈️', title:'Trip Booking', body:'User books trip → immediate: locked lounge tease on confirmation page with "upgrade to access on your departure." Highest-intent upgrade moment — 14% CVR.' },
      { icon:'🔄', title:'Browse Signal', body:'User opens app 3+ times in 7 days without booking → "Your next trip is closer than you think — unlock lounge, eSIM, and 2× coins when you\'re ready."' },
      { icon:'🪙', title:'Coin Threshold', body:'User hits 500 coins on free tier → "You\'re earning at half speed. Upgrade to 2× and watch your balance grow." Shows side-by-side coin balance projection.' },
      { icon:'📅', title:'Seasonal Trigger', body:'D-30 before summer/winter peak → "Your lounge is waiting for peak season." Time-bound upgrade prompt with seasonal travel imagery and benefit preview.' },
    ];
    steps.forEach((st,i)=>{
      const x = 0.25 + i*2.4;
      s.addShape('rect', { x, y:1.3, w:2.2, h:3.5, fill:{ color:'130E01' }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
      s.addShape('rect', { x, y:1.3, w:2.2, h:0.45, fill:{ color:C.accent } });
      s.addText(st.icon, { x, y:1.35, w:2.2, h:0.35, fontSize:16, align:'center' });
      s.addText(st.title, { x, y:1.85, w:2.2, h:0.4, color:C.accent, fontSize:11, bold:true, align:'center' });
      s.addText(st.body, { x:x+0.12, y:2.32, w:2.0, h:2.35, color:C.muted, fontSize:9 });
    });
    s.addShape('rect', { x:0.25, y:5.0, w:9.5, h:0.7, fill:{ color:'130E01' }, line:{ color:C.gold, width:1, transparency:60 } });
    s.addText('Benchmark: PhonePe checkout redesign drove 22% conversion lift by surfacing the right offer at the right moment (cart value × user intent × time signals). DreamTier applies the same contextual trigger model to subscription upgrade timing.', { x:0.4, y:5.1, w:9.2, h:0.52, color:C.gold, fontSize:9.5, italic:true });
  }

  // SLIDE 5 — THE PRODUCT
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.lgray } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('THE PRODUCT', { x:0.4, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('5 Screen States — Free Member to Paid Welcome', { x:0.4, y:0.55, w:9.2, h:0.6, color:C.dark, fontSize:22, bold:true, fontFace:'Calibri' });
    const cards = [
      { num:'01', title:'Freemium Home', body:'Active: coin earn, exclusive deals. Locked (grayed): lounge, eSIM, hotel cashback, insurance. Upgrade banner: "AED 499 — unlock everything." 420 coins balance.' },
      { num:'02', title:'Booking Paywall', body:'Flight confirmed. Lounge tease: DXB Terminal 3 available on departure. Lock badge. eSIM teased below. "2× coins on this booking" note. Book button still active.' },
      { num:'03', title:'Tier Comparison', body:'Free vs Paid feature table. Value-at-stake: pay AED 499 / get AED 1,320 in benefits. Normal price AED 699 shown with strikethrough. Upgrade CTA. Skip link.' },
      { num:'04', title:'Welcome Paid', body:'Confetti. 4 unlock cards: lounge 6 visits / eSIM 5GB / hotel AED 600 / insurance all trips. Billing summary. 2× coin bonus on previous booking applied.' },
      { num:'05', title:'PM Upgrade View', body:'Dark PM screen. Trigger logic table (4 triggers + tags). Conversion funnel (100% → 14% complete upgrade). 4-metric outcome grid.' },
    ];
    cards.forEach((c,i)=>{
      const x = 0.25 + i*1.93;
      s.addShape('rect', { x, y:1.3, w:1.78, h:4.0, fill:{ color:C.white }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
      s.addShape('rect', { x, y:1.3, w:1.78, h:0.4, fill:{ color:C.accent, transparency:85 } });
      s.addText(c.num, { x, y:1.35, w:1.78, h:0.32, color:C.accent, fontSize:13, bold:true, align:'center' });
      s.addText(c.title, { x:x+0.08, y:1.8, w:1.62, h:0.4, color:C.dark, fontSize:11, bold:true });
      s.addText(c.body, { x:x+0.08, y:2.28, w:1.62, h:2.8, color:'475569', fontSize:9.5 });
    });
    s.addText('Prototype: dreampass-tier-prototype.html · All 5 states interactive', { x:0.4, y:5.5, w:9.2, h:0.3, color:C.muted, fontSize:9.5, italic:true });
  }

  // SLIDE 6 — IMPACT & ROI
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('IMPACT & ROI', { x:0.4, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Freemium Conversion as a Revenue Multiplier', { x:0.4, y:0.55, w:9.2, h:0.6, color:C.white, fontSize:22, bold:true, fontFace:'Calibri' });
    const left = [
      { val:'14%', lbl:'Free-to-paid CVR at booking moment', sub:'vs 3–5% from generic account-settings upgrade pages' },
      { val:'2.8×', lbl:'ARPU lift: paid vs free member', sub:'Paid members generate 2.5–3.2× annual revenue of free members' },
      { val:'+22%', lbl:'Trip frequency post-upgrade', sub:'Activated paid members book 1.4 more trips/yr — lounge reinforces value in-destination' },
    ];
    const right = [
      { val:'AED 499', lbl:'Annual paid subscription price', sub:'vs AED 1,320+ in benefits delivered — 2.6× ROI frame for member' },
      { val:'76%', lbl:'Renewal rate: member who used lounge on first trip', sub:'vs 38% for upgraded member who used zero benefits' },
      { val:'31%', lbl:'Value-at-stake paywall lift vs feature-list', sub:'ROI-framed paywall outperforms feature comparison on conversion' },
    ];
    [[left,'Growth Impact'],[right,'Unit Economics']].forEach(([items,title],col)=>{
      const xBase = col===0 ? 0.4 : 5.1;
      s.addShape('rect', { x:xBase, y:1.3, w:4.5, h:3.9, fill:{ color:C.hero }, line:{ color:C.accent, width:1, transparency:70 } });
      s.addText(title, { x:xBase+0.15, y:1.4, w:4.2, h:0.35, color:C.accent, fontSize:10, bold:true, charSpacing:1 });
      items.forEach((item,i)=>{
        s.addText(item.val, { x:xBase+0.15, y:1.85+i*1.1, w:4.2, h:0.55, color:col===0?C.accent:C.gold, fontSize:30, bold:true });
        s.addText(item.lbl, { x:xBase+0.15, y:2.37+i*1.1, w:4.2, h:0.3, color:C.white, fontSize:10, bold:true });
        s.addText(item.sub, { x:xBase+0.15, y:2.63+i*1.1, w:4.2, h:0.3, color:C.muted, fontSize:9 });
      });
    });
    s.addText('(Source: Recurly Subscription Benchmarks 2024; Product-Led Growth Collective 2023; CXL Institute Subscription Conversion 2023; Bond Brand Loyalty 2023)', { x:0.4, y:5.45, w:9.2, h:0.28, color:C.muted, fontSize:8.5, italic:true });
  }

  // SLIDE 7 — PROOF OF WORK
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.lgray } });
    s.addShape('rect', { x:0, y:0, w:4.7, h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('PROOF OF WORK', { x:0.3, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('I Built the Conversion\nEngine at PhonePe.', { x:0.3, y:0.6, w:4.0, h:0.9, color:C.white, fontSize:22, bold:true, fontFace:'Calibri' });
    const proofs = [
      'Checkout conversion redesign: context-aware offer placement (cart value × user intent × time signals) → 22% checkout conversion lift',
      'Pincode Q-commerce: dynamic cart incentivization engine → 35% AOV uplift, 60% cart abandonment reduction, 20% improvement in 30-day retention',
      'ML propensity-to-transact models replacing static rule-based targeting — real-time user-level scoring → 32% marketing burn reduction',
      'Kotak Cherry freemium-to-premium upsell design — WealthTech app MVP with tiered feature access and upgrade moment design',
    ];
    proofs.forEach((p,i)=>{
      s.addShape('rect', { x:0.3, y:1.65+i*1.1, w:4.1, h:0.95, fill:{ color:C.hero }, line:{ color:C.accent, width:1, transparency:70 } });
      s.addText(`• ${p}`, { x:0.45, y:1.72+i*1.1, w:3.8, h:0.8, color:C.muted, fontSize:9.5 });
    });
    const maps = [
      '→ DreamTier booking-moment paywall trigger design',
      '→ DreamTier contextual lounge/eSIM tease at cart/booking',
      '→ DreamTier value-at-stake framing (AED 499 / AED 1,320)',
      '→ DreamTier freemium lock UX + tier comparison page',
    ];
    maps.forEach((m,i)=>{
      s.addShape('rect', { x:5.0, y:1.65+i*1.1, w:4.6, h:0.95, fill:{ color:C.white }, line:{ color:C.accent, width:1 } });
      s.addText(m, { x:5.15, y:1.72+i*1.1, w:4.3, h:0.8, color:C.amber, fontSize:10, bold:true });
    });
    s.addShape('rect', { x:0.3, y:5.45, w:9.4, h:0.55, fill:{ color:C.accent, transparency:90 }, line:{ color:C.accent, width:1 } });
    s.addText('22% checkout conversion lift and 35% AOV uplift at PhonePe are the direct proof of contextual trigger design. DreamTier is the same architecture applied to subscription upgrade conversion.', { x:0.5, y:5.52, w:9.0, h:0.4, color:C.dark, fontSize:10.5, bold:true, italic:true });
  }

  // SLIDE 8 — ROLLOUT PLAN
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.gold } });
    s.addText('ROLLOUT PLAN', { x:0.4, y:0.25, w:9, h:0.25, color:C.gold, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Three Phases — Instrument Intent, Prove CVR, Scale Paywall', { x:0.4, y:0.55, w:9.2, h:0.6, color:C.white, fontSize:24, bold:true, fontFace:'Calibri' });
    const phases = [
      { phase:'Phase 1', period:'Month 1–2: Instrument', body:'Surface lounge tease card on booking confirmation for all free members flying from DXB/AUH/SHJ. Instrument: lounge tease seen, upgrade CTA clicked, paywall viewed, upgrade completed. Control: no tease card (existing flow).', ask:'Data: booking events for free members, flight departure airports, upgrade events' },
      { phase:'Phase 2', period:'Month 3–4: Prove CVR', body:'A/B test paywall framing: Value-at-stake (AED 499 / AED 1,320 value) vs Feature-list (Free vs Paid columns). A/B test trigger timing: immediate post-booking vs D+1 CRM follow-up. Target: 14% CVR at booking trigger. Tune coin earn messaging.', ask:'1 engineer: paywall A/B test framework + upgrade event tracking' },
      { phase:'Phase 3', period:'Month 5–6: Scale', body:'Roll DreamTier paywall surface to all 4 trigger types globally. Integrate in-destination benefit reinforcement: first lounge use → renewal pre-sell notification. Connect DreamActivate for new paid members (activation ladder starts at upgrade). Seasonal campaign: summer upgrade push.', ask:'Alignment: design (paywall UX), growth marketing (seasonal push), finance (freemium tier economics approval)' },
    ];
    phases.forEach((p,i)=>{
      const x = 0.3 + i*3.2;
      s.addShape('rect', { x, y:1.3, w:3.0, h:4.0, fill:{ color:C.hero }, line:{ color:C.gold, width:1, transparency:70 }, shadow:makeShadow() });
      s.addShape('rect', { x, y:1.3, w:3.0, h:0.38, fill:{ color:C.gold, transparency:80 } });
      s.addText(p.phase, { x, y:1.35, w:3.0, h:0.28, color:C.gold, fontSize:11, bold:true, align:'center' });
      s.addText(p.period, { x:x+0.1, y:1.76, w:2.8, h:0.35, color:C.accent, fontSize:10, bold:true });
      s.addText(p.body, { x:x+0.1, y:2.18, w:2.8, h:2.3, color:C.muted, fontSize:9.5 });
      s.addShape('rect', { x:x+0.1, y:4.6, w:2.8, h:0.55, fill:{ color:'130E01' }, line:{ color:C.accent, width:1, transparency:60 } });
      s.addText(`Need: ${p.ask}`, { x:x+0.15, y:4.65, w:2.7, h:0.45, color:C.accent, fontSize:8.5, italic:true });
    });
    s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  ajay-avaghade.github.io/Portfolio', { x:0.4, y:5.52, w:9.2, h:0.28, color:C.muted, fontSize:9, align:'center' });
  }

  await prs.writeFile({ fileName:'assets/dreampass-tier-deck.pptx' });
  console.log('Done: assets/dreampass-tier-deck.pptx');
})();
