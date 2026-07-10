const PptxGenJS = require('pptxgenjs');

(async () => {
  const prs = new PptxGenJS();
  prs.layout = 'LAYOUT_16x9';

  const C = {
    dark:  '071A10',
    hero:  '0A2E17',
    accent:'059669',
    green: '10B981',
    white: 'FFFFFF',
    lgray: 'F0FAF4',
    muted: '94A3B8',
    gold:  '4ADE80',
  };
  const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.14 });

  // SLIDE 1 — COVER
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.04, fill:{ color:C.accent } });
    s.addShape('rect', { x:0, y:7.46, w:'100%', h:0.04, fill:{ color:C.accent } });
    for(let i=0;i<6;i++){ s.addShape('line', { x:6.5+i*0.45, y:0, w:0, h:7.5, line:{ color:C.accent, width:0.5, transparency:85 } }); }
    s.addText('DREAMPASS · TREVOLUTION · CASE STUDY B', { x:0.5, y:0.3, w:9, h:0.25, color:C.muted, fontSize:8, bold:true, charSpacing:2 });
    s.addText('DreamLoop', { x:0.5, y:1.0, w:9, h:1.4, color:C.white, fontSize:52, bold:true, fontFace:'Calibri' });
    s.addText('Referral & Viral Acquisition Engine — Turning Every Booking Into a Growth Moment', { x:0.5, y:2.3, w:8.5, h:0.5, color:C.accent, fontSize:18, fontFace:'Calibri' });
    s.addText('Ajay Avaghade · Product Manager', { x:0.5, y:3.0, w:5, h:0.3, color:C.muted, fontSize:12 });
    s.addShape('rect', { x:6.8, y:1.8, w:2.8, h:1.4, fill:{ color:C.accent, transparency:88 }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
    s.addText('−40%', { x:6.8, y:1.9, w:2.8, h:0.7, color:C.white, fontSize:36, bold:true, align:'center' });
    s.addText('CAC vs paid acquisition for referred members', { x:6.8, y:2.5, w:2.8, h:0.5, color:C.accent, fontSize:10, align:'center' });
  }

  // SLIDE 2 — THE PROBLEM
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:0.06, h:'100%', fill:{ color:C.accent } });
    s.addText('THE PROBLEM', { x:0.3, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Dreampass Has No Referral Programme — Every Acquired Member Is a Missed Viral Loop', { x:0.3, y:0.55, w:9.4, h:0.9, color:C.white, fontSize:24, bold:true, fontFace:'Calibri' });
    const cols = [
      { icon:'📢', val:'0', lbl:'Referral programme today', sub:'No mechanism to turn satisfied members into growth agents — every acquisition is paid or organic search' },
      { icon:'✈️', val:'2.4×', lbl:'Post-booking referral CVR lift', sub:'Friends invited immediately after a booking convert at 2.4× the rate of cold invites — the moment is untapped' },
      { icon:'💸', val:'AED 200+', lbl:'Avg CAC from paid channels', sub:'Referred members cost 40–60% less to acquire and have 1.8× higher LTV due to social proof at join' },
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
    s.addText('Root cause: Dreampass acquires members in isolation. The post-booking emotional peak — the moment when a member is most excited and most likely to share — goes completely unexploited.', { x:0.5, y:4.5, w:9.0, h:0.7, color:C.white, fontSize:10, italic:true });
  }

  // SLIDE 3 — THE INSIGHT
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.lgray } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('THE INSIGHT', { x:0.4, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('The Best Time to Ask for a Referral Is Right After the Booking Confirmation.', { x:0.4, y:0.55, w:9.2, h:0.8, color:C.dark, fontSize:26, bold:true, fontFace:'Calibri' });
    s.addShape('rect', { x:0.4, y:1.5, w:4.4, h:3.3, fill:{ color:'FEF2F2' }, line:{ color:'FECACA', width:1 }, shadow:makeShadow() });
    s.addText('❌  Without DreamLoop', { x:0.55, y:1.65, w:4.1, h:0.35, color:'991B1B', fontSize:11, bold:true });
    ['No referral mechanism — members can\'t share','Post-booking moment is dead air','Each acquired member stays siloed','Paid CAC dominates acquisition mix','No peer social proof at the join moment'].forEach((t,i)=>{ s.addText(`• ${t}`, { x:0.6, y:2.1+i*0.42, w:4.1, h:0.38, color:'7F1D1D', fontSize:10 }); });
    s.addShape('rect', { x:5.1, y:1.5, w:4.6, h:3.3, fill:{ color:'F0FAF4' }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
    s.addText('✅  With DreamLoop', { x:5.25, y:1.65, w:4.2, h:0.35, color:C.accent, fontSize:11, bold:true });
    ['Booking confirmation triggers "Share with a friend" card','Referral card pre-built: trip destination + dual reward','Friend lands with full context — who invited, what they get','500 coins for referrer + 300 welcome coins for friend','Ongoing: 50 coins per trip the friend books — forever'].forEach((t,i)=>{ s.addText(`• ${t}`, { x:5.25, y:2.1+i*0.42, w:4.3, h:0.38, color:C.accent, fontSize:10 }); });
    s.addShape('rect', { x:0.4, y:5.0, w:9.2, h:0.7, fill:{ color:C.accent, transparency:85 }, line:{ color:C.accent, width:1 } });
    s.addText('Key insight: The dual-sided referral (both parties earn) increases completion rates by 34% vs single-sided programmes. The "ongoing earn" mechanic (50 coins/trip) creates a compounding loop — the referrer has skin in the friend\'s continued travel.', { x:0.6, y:5.08, w:8.8, h:0.55, color:C.dark, fontSize:10, italic:true });
  }

  // SLIDE 4 — THE MECHANIC
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.hero } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.gold } });
    s.addText('THE MECHANIC', { x:0.4, y:0.25, w:9, h:0.25, color:C.gold, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Five Steps — Booking Moment to Viral Loop Closed', { x:0.4, y:0.55, w:9, h:0.6, color:C.white, fontSize:26, bold:true, fontFace:'Calibri' });
    const steps = [
      { n:'1', title:'Booking Trigger', body:'Member books trip. Post-confirmation screen surfaces DreamLoop share card: destination + dual reward offer. Pre-built, shareable in 1 tap.' },
      { n:'2', title:'Share Card', body:'Personalised referral card: "Marcus just booked Dubai — join Dreampass to travel smarter." Deep link with referrer code. WhatsApp, Telegram, email, or copy link.' },
      { n:'3', title:'Friend Landing', body:'Friend opens link. Full context: who invited, what they get (300 welcome coins), what the referrer earns (500 coins). Join in 2 taps.' },
      { n:'4', title:'First Booking', body:'Friend books first trip. Referral bonus triggered. Both parties notified. Ongoing: referrer earns 50 coins on every future trip friend books — no expiry.' },
      { n:'5', title:'Loop Closes', body:'Referrer notified "Sarah booked!" with coin credit. Travel Circle updated. Badge at 3 referrals. Referrer\'s LTV now tied to friend\'s continued travel.' },
    ];
    steps.forEach((st,i)=>{
      const x = 0.25 + i*1.93;
      s.addShape('rect', { x, y:1.3, w:1.78, h:3.5, fill:{ color:'071A10' }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
      s.addShape('rect', { x, y:1.3, w:1.78, h:0.45, fill:{ color:C.accent } });
      s.addText(st.n, { x, y:1.35, w:1.78, h:0.35, color:C.white, fontSize:14, bold:true, align:'center' });
      s.addText(st.title, { x, y:1.85, w:1.78, h:0.4, color:C.accent, fontSize:11, bold:true, align:'center' });
      s.addText(st.body, { x:x+0.1, y:2.3, w:1.6, h:2.4, color:C.muted, fontSize:9 });
    });
    s.addShape('rect', { x:0.25, y:5.0, w:9.5, h:0.7, fill:{ color:'071A10' }, line:{ color:C.gold, width:1, transparency:60 } });
    s.addText('Benchmark: PhonePe merchant referral network drove 23% lower CAC across 5,000+ B2B merchants via peer-to-peer acquisition mechanics. DreamLoop applies the same flywheel to B2C travel membership.', { x:0.4, y:5.1, w:9.2, h:0.52, color:C.gold, fontSize:9.5, italic:true });
  }

  // SLIDE 5 — THE PRODUCT
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.lgray } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('THE PRODUCT', { x:0.4, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('5 Screen States — Referrer to Friend Journey End to End', { x:0.4, y:0.55, w:9.2, h:0.6, color:C.dark, fontSize:22, bold:true, fontFace:'Calibri' });
    const cards = [
      { num:'01', title:'Referral Home', body:'0 referrals. 500 coin earn per friend. Ongoing 50-coin/trip earn explained. Empty "Travel Circle" motivating first invite. Stats: 0/0/500.' },
      { num:'02', title:'Share Trip Moment', body:'Post-booking card: trip destination pre-filled. Referral card preview. 4 share channels. Copy link. "2.4× conversion" insight tip.' },
      { num:'03', title:'Friend Landing', body:'Who invited (Marcus). Welcome offer (300 coins). Dual-sided reward breakdown. Join CTA. Social proof from real travel context.' },
      { num:'04', title:'Friend First Booking', body:'Flight search. Coin earn banner: +642 booking coins + 300 welcome bonus = 942 coins. Referrer notification preview: Marcus earns 500.' },
      { num:'05', title:'Referrer Success', body:'Confetti. 500 coins credited. Travel Circle: 1 member (Sarah). Ongoing earn: 50/trip. Badge progress: 2 more referrals to Gold.' },
    ];
    cards.forEach((c,i)=>{
      const x = 0.25 + i*1.93;
      s.addShape('rect', { x, y:1.3, w:1.78, h:4.0, fill:{ color:C.white }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
      s.addShape('rect', { x, y:1.3, w:1.78, h:0.4, fill:{ color:C.accent, transparency:85 } });
      s.addText(c.num, { x, y:1.35, w:1.78, h:0.32, color:C.accent, fontSize:13, bold:true, align:'center' });
      s.addText(c.title, { x:x+0.08, y:1.8, w:1.62, h:0.4, color:C.dark, fontSize:11, bold:true });
      s.addText(c.body, { x:x+0.08, y:2.28, w:1.62, h:2.8, color:'475569', fontSize:9.5 });
    });
    s.addText('Prototype: dreampass-loop-prototype.html · All 5 states interactive', { x:0.4, y:5.5, w:9.2, h:0.3, color:C.muted, fontSize:9.5, italic:true });
  }

  // SLIDE 6 — IMPACT & ROI
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('IMPACT & ROI', { x:0.4, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Referral as a Compounding Acquisition Machine', { x:0.4, y:0.55, w:9.2, h:0.6, color:C.white, fontSize:22, bold:true, fontFace:'Calibri' });
    const left = [
      { val:'−40%', lbl:'CAC vs paid channels', sub:'Referred members cost 40–60% less to acquire' },
      { val:'2.1×', lbl:'12-month LTV vs cold acquired', sub:'Ongoing earn creates compounding referrer engagement' },
      { val:'34%', lbl:'Dual-sided programme completion lift', sub:'Both-earn vs single-earn referral mechanics' },
    ];
    const right = [
      { val:'AED 800', lbl:'Gross margin per referred member', sub:'At AED 499 subscription — cost of AED 500+300 coin set is ~AED 80 net' },
      { val:'2.4×', lbl:'Post-booking referral CVR', sub:'vs cold invite from ads or email' },
      { val:'50 coins', lbl:'Ongoing earn per friend trip', sub:'Self-reinforcing: referrer\'s coins grow with every trip friend books' },
    ];
    [[left,'Viral Growth Impact'],[right,'Unit Economics']].forEach(([items,title],col)=>{
      const xBase = col===0 ? 0.4 : 5.1;
      s.addShape('rect', { x:xBase, y:1.3, w:4.5, h:3.9, fill:{ color:C.hero }, line:{ color:C.accent, width:1, transparency:70 } });
      s.addText(title, { x:xBase+0.15, y:1.4, w:4.2, h:0.35, color:C.accent, fontSize:10, bold:true, charSpacing:1 });
      items.forEach((item,i)=>{
        s.addText(item.val, { x:xBase+0.15, y:1.85+i*1.1, w:4.2, h:0.55, color:col===0?C.accent:C.gold, fontSize:30, bold:true });
        s.addText(item.lbl, { x:xBase+0.15, y:2.37+i*1.1, w:4.2, h:0.3, color:C.white, fontSize:10, bold:true });
        s.addText(item.sub, { x:xBase+0.15, y:2.63+i*1.1, w:4.2, h:0.3, color:C.muted, fontSize:9 });
      });
    });
    s.addText('(Source: Extole Referral Marketing Benchmark 2024; Friendbuy State of Referral Marketing 2023; Nielsen Consumer Trust Report 2023)', { x:0.4, y:5.45, w:9.2, h:0.28, color:C.muted, fontSize:8.5, italic:true });
  }

  // SLIDE 7 — PROOF OF WORK
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.lgray } });
    s.addShape('rect', { x:0, y:0, w:4.7, h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('PROOF OF WORK', { x:0.3, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('I Built the Acquisition\nFlywheels. Here\'s the Map.', { x:0.3, y:0.6, w:4.0, h:0.9, color:C.white, fontSize:22, bold:true, fontFace:'Calibri' });
    const proofs = [
      'PhonePe self-serve PG platform — 5,000+ B2B merchants acquired via peer-to-peer referral and network effects',
      'ML-driven marketplace — 500+ brand partners onboarded; designed partner acquisition loop with self-serve tooling',
      'Pincode Q-commerce — 5M+ new users/month; 23% lower CAC via optimised acquisition mix and referral hooks',
      'Kotak Cherry — 100K+ downloads in launch window via word-of-mouth GTM with seeded referral programme',
    ];
    proofs.forEach((p,i)=>{
      s.addShape('rect', { x:0.3, y:1.65+i*1.1, w:4.1, h:0.95, fill:{ color:C.hero }, line:{ color:C.accent, width:1, transparency:70 } });
      s.addText(`• ${p}`, { x:0.45, y:1.72+i*1.1, w:3.8, h:0.8, color:C.muted, fontSize:9.5 });
    });
    const maps = [
      '→ DreamLoop dual-sided referral coin mechanics',
      '→ DreamLoop travel circle + partner self-serve share',
      '→ DreamLoop post-booking share card trigger',
      '→ DreamLoop friend landing with full referral context',
    ];
    maps.forEach((m,i)=>{
      s.addShape('rect', { x:5.0, y:1.65+i*1.1, w:4.6, h:0.95, fill:{ color:C.white }, line:{ color:C.accent, width:1 } });
      s.addText(m, { x:5.15, y:1.72+i*1.1, w:4.3, h:0.8, color:C.accent, fontSize:10, bold:true });
    });
    s.addShape('rect', { x:0.3, y:5.45, w:9.4, h:0.55, fill:{ color:C.accent, transparency:90 }, line:{ color:C.accent, width:1 } });
    s.addText('Every acquisition loop I have built has been peer-to-peer at its core. DreamLoop is the direct application of that architecture to travel membership.', { x:0.5, y:5.52, w:9.0, h:0.4, color:C.dark, fontSize:10.5, bold:true, italic:true });
  }

  // SLIDE 8 — ROLLOUT PLAN
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.gold } });
    s.addText('ROLLOUT PLAN', { x:0.4, y:0.25, w:9, h:0.25, color:C.gold, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Three Phases — Seed, Prove Viral Coefficient, Scale', { x:0.4, y:0.55, w:9.2, h:0.6, color:C.white, fontSize:24, bold:true, fontFace:'Calibri' });
    const phases = [
      { phase:'Phase 1', period:'Month 1–2: Seed', body:'Launch DreamLoop share card on post-booking confirmation for 1,000 most-active members. Instrument: share rate, friend click-through, join rate, first booking rate. Measure referral coefficient K.', ask:'Data: booking events, member share actions, friend join events' },
      { phase:'Phase 2', period:'Month 3–4: Prove K', body:'A/B test: 500 coins (current) vs 300 coins + tier-badge unlock. Test timing: immediate post-booking vs D1 CRM follow-up. Optimise for K > 0.3 (each member brings 0.3+ new members). Tune dual-sided reward.', ask:'1 engineer: referral link system + coin credit pipeline' },
      { phase:'Phase 3', period:'Month 5–6: Scale', body:'Roll DreamLoop to all members globally. Integrate Travel Circle badge into member profile. Trigger DreamActivate for referred friends at D0 (warm join path). Gold Circle badge at 3 referrals — unlocks bonus lounge visit.', ask:'Alignment: CRM, member ops, coin economy team for budget approval' },
    ];
    phases.forEach((p,i)=>{
      const x = 0.3 + i*3.2;
      s.addShape('rect', { x, y:1.3, w:3.0, h:4.0, fill:{ color:C.hero }, line:{ color:C.gold, width:1, transparency:70 }, shadow:makeShadow() });
      s.addShape('rect', { x, y:1.3, w:3.0, h:0.38, fill:{ color:C.gold, transparency:80 } });
      s.addText(p.phase, { x, y:1.35, w:3.0, h:0.28, color:C.gold, fontSize:11, bold:true, align:'center' });
      s.addText(p.period, { x:x+0.1, y:1.76, w:2.8, h:0.35, color:C.accent, fontSize:10, bold:true });
      s.addText(p.body, { x:x+0.1, y:2.18, w:2.8, h:2.3, color:C.muted, fontSize:9.5 });
      s.addShape('rect', { x:x+0.1, y:4.6, w:2.8, h:0.55, fill:{ color:'071A10' }, line:{ color:C.accent, width:1, transparency:60 } });
      s.addText(`Need: ${p.ask}`, { x:x+0.15, y:4.65, w:2.7, h:0.45, color:C.accent, fontSize:8.5, italic:true });
    });
    s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  ajay-avaghade.github.io/Portfolio', { x:0.4, y:5.52, w:9.2, h:0.28, color:C.muted, fontSize:9, align:'center' });
  }

  await prs.writeFile({ fileName:'assets/dreampass-loop-deck.pptx' });
  console.log('Done: assets/dreampass-loop-deck.pptx');
})();
