const PptxGenJS = require('pptxgenjs');

(async () => {
  const prs = new PptxGenJS();
  prs.layout = 'LAYOUT_16x9';

  const C = {
    dark:  '0C1A30',
    hero:  '0A2440',
    accent:'0891B2',
    teal:  '0E7490',
    white: 'FFFFFF',
    lgray: 'F0F7FA',
    muted: '94A3B8',
    gold:  'FCD34D',
    green: '10B981',
  };
  const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.14 });

  // ── SLIDE 1 — COVER ──────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.04, fill:{ color:C.accent } });
    s.addShape('rect', { x:0, y:7.46, w:'100%', h:0.04, fill:{ color:C.accent } });
    // diagonal accent lines
    for(let i=0;i<6;i++){
      s.addShape('line', { x:6.5+i*0.45, y:0, w:0, h:7.5, line:{ color:C.accent, width:0.5, transparency:85 } });
    }
    s.addText('DREAMPASS · TREVOLUTION · CASE STUDY A', { x:0.5, y:0.3, w:9, h:0.25, color:C.muted, fontSize:8, bold:true, charSpacing:2 });
    s.addText('DreamActivate', { x:0.5, y:1.0, w:9, h:1.4, color:C.white, fontSize:52, bold:true, fontFace:'Calibri' });
    s.addText('First-Benefit Activation Engine — Closing the 14-Day Renewal Cliff', { x:0.5, y:2.3, w:8.5, h:0.5, color:C.accent, fontSize:18, bold:false, fontFace:'Calibri' });
    s.addText('Ajay Avaghade · Product Manager', { x:0.5, y:3.0, w:5, h:0.3, color:C.muted, fontSize:12 });
    s.addShape('rect', { x:6.8, y:1.8, w:2.8, h:1.4, fill:{ color:C.accent, transparency:88 }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
    s.addText('+2.8×', { x:6.8, y:1.9, w:2.8, h:0.7, color:C.white, fontSize:36, bold:true, align:'center' });
    s.addText('renewal rate for activated members', { x:6.8, y:2.5, w:2.8, h:0.5, color:C.accent, fontSize:10, align:'center' });
  }

  // ── SLIDE 2 — THE PROBLEM ─────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:0.06, h:'100%', fill:{ color:C.accent } });
    s.addText('THE PROBLEM', { x:0.3, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Members Who Don\'t Redeem a Benefit Within 14 Days Almost Never Renew', { x:0.3, y:0.55, w:9.4, h:0.9, color:C.white, fontSize:26, bold:true, fontFace:'Calibri' });
    const cols = [
      { icon:'⚠️', val:'<14 days', lbl:'Activation cliff', sub:'Members who fail to use any benefit within 14 days have 3.2× higher churn at renewal' },
      { icon:'💤', val:'0 of 6', lbl:'Benefits used on avg', sub:'The average Dreampass member redeems zero benefits in their first month — value is invisible' },
      { icon:'🪙', val:'1,240', lbl:'Avg coins earned', sub:'Members accumulate coins from bookings but don\'t know how to spend them — creating a dead asset' },
    ];
    cols.forEach((c,i)=>{
      const x = 0.3 + i*3.2;
      s.addShape('rect', { x, y:1.65, w:3.0, h:2.6, fill:{ color:C.hero }, line:{ color:C.accent, width:1, transparency:60 }, shadow:makeShadow() });
      s.addText(c.icon, { x, y:1.75, w:3.0, h:0.5, fontSize:22, align:'center' });
      s.addText(c.val, { x, y:2.2, w:3.0, h:0.6, color:C.accent, fontSize:28, bold:true, align:'center' });
      s.addText(c.lbl, { x, y:2.75, w:3.0, h:0.3, color:C.white, fontSize:11, bold:true, align:'center' });
      s.addText(c.sub, { x:x+0.1, y:3.1, w:2.8, h:0.95, color:C.muted, fontSize:9.5, align:'left' });
    });
    s.addShape('rect', { x:0.3, y:4.4, w:9.4, h:0.9, fill:{ color:C.accent, transparency:90 }, line:{ color:C.accent, width:1 } });
    s.addText('Root cause: Dreampass has no activation architecture. Members join, coins pile up, and benefits sit untouched until renewal reminds them they\'ve been paying for nothing — at which point it\'s too late.', { x:0.5, y:4.5, w:9.0, h:0.7, color:C.white, fontSize:10, italic:true });
  }

  // ── SLIDE 3 — THE INSIGHT ─────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.lgray } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('THE INSIGHT', { x:0.4, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('The First Redemption Is the Product. Everything Else Is Marketing.', { x:0.4, y:0.55, w:9.2, h:0.8, color:C.dark, fontSize:26, bold:true, fontFace:'Calibri' });

    // Left: before
    s.addShape('rect', { x:0.4, y:1.5, w:4.4, h:3.3, fill:{ color:'FEF2F2' }, line:{ color:'FECACA', width:1 }, shadow:makeShadow() });
    s.addText('❌  Without DreamActivate', { x:0.55, y:1.65, w:4.1, h:0.35, color:'991B1B', fontSize:11, bold:true });
    const before = ['No benefit surfacing post-join','Coins feel like points — abstract, unclaimed','Member reaches renewal having never experienced value','Cancellation: "I forgot I even had this"','14-day window passes with zero engagement'];
    before.forEach((t,i)=>{ s.addText(`• ${t}`, { x:0.6, y:2.1+i*0.42, w:4.1, h:0.38, color:'7F1D1D', fontSize:10 }); });

    // Right: after
    s.addShape('rect', { x:5.1, y:1.5, w:4.6, h:3.3, fill:{ color:'F0F9FA' }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
    s.addText('✅  With DreamActivate', { x:5.25, y:1.65, w:4.2, h:0.35, color:C.teal, fontSize:11, bold:true });
    const after = ['Milestone bar on home: "Redeem your first benefit"','D3 nudge: "200 coins = free eSIM on your next trip"','D7 push: "4 lounge visits expiring — book a trip"','First redemption → renewal journey starts immediately','2.8× renewal rate vs non-activated members'];
    after.forEach((t,i)=>{ s.addText(`• ${t}`, { x:5.25, y:2.1+i*0.42, w:4.3, h:0.38, color:C.teal, fontSize:10 }); });

    s.addShape('rect', { x:0.4, y:5.0, w:9.2, h:0.7, fill:{ color:C.accent, transparency:85 }, line:{ color:C.accent, width:1 } });
    s.addText('Key insight: The milestone bar is a commitment device — it shifts the member\'s mental model from "subscription cost" to "progress toward a goal I\'ve already started." This is the same mechanic that drove D30 retention +15% at PhonePe.', { x:0.6, y:5.08, w:8.8, h:0.55, color:C.dark, fontSize:10, italic:true });
  }

  // ── SLIDE 4 — THE MECHANIC ────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.hero } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.gold } });
    s.addText('THE MECHANIC', { x:0.4, y:0.25, w:9, h:0.25, color:C.gold, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Five Triggers, One Activation Journey', { x:0.4, y:0.55, w:9, h:0.6, color:C.white, fontSize:26, bold:true, fontFace:'Calibri' });
    const steps = [
      { day:'D0', title:'Join', body:'Milestone bar activates on home. "Redeem your first benefit to start your renewal journey" CTA with coin balance shown.' },
      { day:'D3', title:'First Nudge', body:'Push: "Your 1,240 coins are worth an eSIM (200 coins) — activate it for free on your next trip." Deep-link to benefit discovery.' },
      { day:'D7', title:'Activation Cliff', body:'In-app banner + email: "7 days in — members who redeem now are 2.8× more likely to renew. Choose your first benefit." Urgency without alarm.' },
      { day:'D14', title:'Redemption', body:'Member redeems eSIM or lounge. Confetti state + renewal journey bar appears. Milestone: 25% toward next renewal unlocked.' },
      { day:'D90', title:'Renewal Path', body:'Quarterly check-in: benefit usage summary + renewal pre-sell. Members who used 2+ benefits renew at 76% vs 38% baseline.' },
    ];
    steps.forEach((st,i)=>{
      const x = 0.25 + i*1.93;
      s.addShape('rect', { x, y:1.3, w:1.78, h:3.5, fill:{ color:'0A1F3A' }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
      s.addShape('rect', { x, y:1.3, w:1.78, h:0.45, fill:{ color:C.accent } });
      s.addText(st.day, { x, y:1.35, w:1.78, h:0.35, color:C.white, fontSize:14, bold:true, align:'center' });
      s.addText(st.title, { x, y:1.85, w:1.78, h:0.4, color:C.accent, fontSize:11, bold:true, align:'center' });
      s.addText(st.body, { x:x+0.1, y:2.3, w:1.6, h:2.4, color:C.muted, fontSize:9, align:'left' });
    });
    s.addShape('rect', { x:0.25, y:5.0, w:9.5, h:0.7, fill:{ color:'0A1F3A' }, line:{ color:C.gold, width:1, transparency:60 } });
    s.addText('PhonePe proof point: Segment-to-channel intervention logic (user state → CRM vs. merchandising vs. offer) with 7-day activation cliff drove D30 retention +15% at PhonePe. DreamActivate applies the identical architecture to Dreampass member onboarding.', { x:0.4, y:5.1, w:9.2, h:0.52, color:C.gold, fontSize:9.5, italic:true });
  }

  // ── SLIDE 5 — THE PRODUCT ─────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.lgray } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('THE PRODUCT', { x:0.4, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('5 Screen States — Member Activation Journey End to End', { x:0.4, y:0.55, w:9.2, h:0.6, color:C.dark, fontSize:22, bold:true, fontFace:'Calibri' });
    const cards = [
      { num:'01', title:'Member Home', body:'1,240 coins. 85% progress to milestone. Activation step tracker: "Redeem first benefit" highlighted. "Choose Your First Benefit" CTA.' },
      { num:'02', title:'Benefit Discovery', body:'eSIM 200 coins (featured). Lounge 800. Hotel 400. Insurance 300. Sorted by ease of first redemption, not alphabetically.' },
      { num:'03', title:'Redemption Confirm', body:'eSIM selected. Balance summary (before/after). Renewal journey preview: "This completes 25% of your renewal milestone."' },
      { num:'04', title:'Activated!', body:'Confetti. 4-step renewal journey bar (2 complete). "2.8× more likely to renew" stat pill. Next milestone: 3 more benefits.' },
      { num:'05', title:'PM Lifecycle View', body:'Dark PM-mode screen. Trigger logic per day (D0→D90). 4-metric grid: activation rate, redemption rate, cliff saves, renewal lift.' },
    ];
    cards.forEach((c,i)=>{
      const x = 0.25 + i*1.93;
      s.addShape('rect', { x, y:1.3, w:1.78, h:4.0, fill:{ color:C.white }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
      s.addShape('rect', { x, y:1.3, w:1.78, h:0.4, fill:{ color:C.accent, transparency:85 } });
      s.addText(`${c.num}`, { x, y:1.35, w:1.78, h:0.32, color:C.accent, fontSize:13, bold:true, align:'center' });
      s.addText(c.title, { x:x+0.08, y:1.8, w:1.62, h:0.4, color:C.dark, fontSize:11, bold:true });
      s.addText(c.body, { x:x+0.08, y:2.28, w:1.62, h:2.8, color:'475569', fontSize:9.5 });
    });
    s.addText('Prototype: dreampass-activate-prototype.html · All 5 states interactive', { x:0.4, y:5.5, w:9.2, h:0.3, color:C.muted, fontSize:9.5, italic:true });
  }

  // ── SLIDE 6 — IMPACT & ROI ────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('IMPACT & ROI', { x:0.4, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Built on PhonePe Proof Points — Projected for Dreampass Scale', { x:0.4, y:0.55, w:9.2, h:0.6, color:C.white, fontSize:22, bold:true, fontFace:'Calibri' });

    const left = [
      { val:'2.8×', lbl:'Renewal rate lift', sub:'Activated vs non-activated members' },
      { val:'+15%', lbl:'D30 retention', sub:'PhonePe benchmark (segment-to-channel model)' },
      { val:'76%', lbl:'Renewal rate at 2+ benefits used', sub:'vs 38% baseline for zero-benefit members' },
    ];
    const right = [
      { val:'AED 0', lbl:'Marginal cost of first nudge sequence', sub:'CRM + push — existing channel infrastructure' },
      { val:'7 days', lbl:'Time to first signal', sub:'Activation cliff detection can run day 1' },
      { val:'+22%', lbl:'Lifetime bookings per member', sub:'Activated members book 1.4 more trips/yr on avg' },
    ];
    [[left,'Member Impact'],[right,'Dreampass Business ROI']].forEach(([items,title],col)=>{
      const xBase = col===0 ? 0.4 : 5.1;
      s.addShape('rect', { x:xBase, y:1.3, w:4.5, h:3.9, fill:{ color:C.hero }, line:{ color:C.accent, width:1, transparency:70 } });
      s.addText(title, { x:xBase+0.15, y:1.4, w:4.2, h:0.35, color:C.accent, fontSize:10, bold:true, charSpacing:1 });
      items.forEach((item,i)=>{
        s.addText(item.val, { x:xBase+0.15, y:1.85+i*1.1, w:4.2, h:0.55, color:col===0?C.accent:C.gold, fontSize:30, bold:true });
        s.addText(item.lbl, { x:xBase+0.15, y:2.37+i*1.1, w:4.2, h:0.3, color:C.white, fontSize:10, bold:true });
        s.addText(item.sub, { x:xBase+0.15, y:2.63+i*1.1, w:4.2, h:0.3, color:C.muted, fontSize:9 });
      });
    });
    s.addText('(Source: Antavo Global Customer Loyalty 2024; Bond Brand Loyalty 2023; PhonePe Internal — Ajay Avaghade)', { x:0.4, y:5.45, w:9.2, h:0.28, color:C.muted, fontSize:8.5, italic:true });
  }

  // ── SLIDE 7 — PROOF OF WORK ───────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.lgray } });
    s.addShape('rect', { x:0, y:0, w:4.7, h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('PROOF OF WORK', { x:0.3, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('I Built This Architecture\nat PhonePe.', { x:0.3, y:0.6, w:4.0, h:0.9, color:C.white, fontSize:22, bold:true, fontFace:'Calibri' });
    const proofs = [
      'Segment-to-channel lifecycle logic → D30 retention +15% at PhonePe (350M+ MAU)',
      '7-day activation cliff detection — bank-link + first transaction = 3× better D30 outcomes',
      '3-category breadth threshold → ~70% better D30 retention vs single-category users',
      'Masthead moratorium: negotiated no-ads for first 7 days — CLTV model vs week-one ad revenue',
    ];
    proofs.forEach((p,i)=>{
      s.addShape('rect', { x:0.3, y:1.65+i*1.1, w:4.1, h:0.95, fill:{ color:C.hero }, line:{ color:C.accent, width:1, transparency:70 } });
      s.addText(`• ${p}`, { x:0.45, y:1.72+i*1.1, w:3.8, h:0.8, color:C.muted, fontSize:9.5 });
    });
    const maps = [
      '→ DreamActivate trigger ladder (D0/D3/D7/D14/D90)',
      '→ DreamActivate 14-day cliff detection + push nudge',
      '→ DreamActivate 3-benefit milestone unlock mechanic',
      '→ DreamActivate benefit-first onboarding (no promotion noise)',
    ];
    maps.forEach((m,i)=>{
      s.addShape('rect', { x:5.0, y:1.65+i*1.1, w:4.6, h:0.95, fill:{ color:C.white }, line:{ color:C.accent, width:1 } });
      s.addText(m, { x:5.15, y:1.72+i*1.1, w:4.3, h:0.8, color:C.teal, fontSize:10, bold:true });
    });
    s.addShape('rect', { x:0.3, y:5.45, w:9.4, h:0.55, fill:{ color:C.accent, transparency:90 }, line:{ color:C.accent, width:1 } });
    s.addText('This is architecture I have shipped — applied directly to Dreampass\'s 14-day activation problem.', { x:0.5, y:5.52, w:9.0, h:0.4, color:C.dark, fontSize:10.5, bold:true, italic:true });
  }

  // ── SLIDE 8 — ROLLOUT PLAN ────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.gold } });
    s.addText('ROLLOUT PLAN', { x:0.4, y:0.25, w:9, h:0.25, color:C.gold, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Three Phases — Instrument, Prove, Scale', { x:0.4, y:0.55, w:9.2, h:0.6, color:C.white, fontSize:24, bold:true, fontFace:'Calibri' });
    const phases = [
      { phase:'Phase 1', period:'Month 1–2: Instrument', body:'Activate milestone bar + step tracker on home for all new members. Deploy D0/D3/D7 trigger sequence via CRM. Instrument redemption funnel: join → first nudge seen → benefit discovered → redeemed. Control group: 50% of new members (no activation UX).', ask:'Data: member join date, benefit redemption events, renewal flag' },
      { phase:'Phase 2', period:'Month 3–4: Prove', body:'Run A/B: Milestone bar vs flat "explore benefits" link. A/B test nudge copy: value ("200 coins = free eSIM") vs urgency ("14-day window"). Measure activation rate, cliff-save rate, 30-day engagement delta between test and control groups.', ask:'1 engineer: event tracking + trigger pipeline' },
      { phase:'Phase 3', period:'Month 5–6: Scale', body:'Roll DreamActivate to all new member onboarding globally. Launch 90-day renewal pre-sell sequence for activated members. Integrate DreamLoop referral trigger at D14 redemption moment (activated member = highest referral intent).', ask:'Alignment: CRM team, member ops, design for milestone bar implementation' },
    ];
    phases.forEach((p,i)=>{
      const x = 0.3 + i*3.2;
      s.addShape('rect', { x, y:1.3, w:3.0, h:4.0, fill:{ color:C.hero }, line:{ color:C.gold, width:1, transparency:70 }, shadow:makeShadow() });
      s.addShape('rect', { x, y:1.3, w:3.0, h:0.38, fill:{ color:C.gold, transparency:80 } });
      s.addText(p.phase, { x, y:1.35, w:3.0, h:0.28, color:C.gold, fontSize:11, bold:true, align:'center' });
      s.addText(p.period, { x:x+0.1, y:1.76, w:2.8, h:0.35, color:C.accent, fontSize:10, bold:true });
      s.addText(p.body, { x:x+0.1, y:2.18, w:2.8, h:2.3, color:C.muted, fontSize:9.5 });
      s.addShape('rect', { x:x+0.1, y:4.6, w:2.8, h:0.55, fill:{ color:'0A1F3A' }, line:{ color:C.accent, width:1, transparency:60 } });
      s.addText(`Need: ${p.ask}`, { x:x+0.15, y:4.65, w:2.7, h:0.45, color:C.accent, fontSize:8.5, italic:true });
    });
    s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  ajay-avaghade.github.io/Portfolio', { x:0.4, y:5.52, w:9.2, h:0.28, color:C.muted, fontSize:9, align:'center' });
  }

  await prs.writeFile({ fileName:'assets/dreampass-activate-deck.pptx' });
  console.log('Done: assets/dreampass-activate-deck.pptx');
})();
