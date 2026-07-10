const PptxGenJS = require('pptxgenjs');

(async () => {
  const prs = new PptxGenJS();
  prs.layout = 'LAYOUT_16x9';

  const C = {
    dark:  '0D071A',
    hero:  '1A0834',
    accent:'7C3AED',
    purple:'6D28D9',
    white: 'FFFFFF',
    lgray: 'F5F3FF',
    muted: '94A3B8',
    gold:  'C084FC',
    red:   'EF4444',
  };
  const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.14 });

  // SLIDE 1 — COVER
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.04, fill:{ color:C.accent } });
    s.addShape('rect', { x:0, y:7.46, w:'100%', h:0.04, fill:{ color:C.accent } });
    for(let i=0;i<6;i++){ s.addShape('line', { x:6.5+i*0.45, y:0, w:0, h:7.5, line:{ color:C.accent, width:0.5, transparency:85 } }); }
    s.addText('DREAMPASS · TREVOLUTION · CASE STUDY C', { x:0.5, y:0.3, w:9, h:0.25, color:C.muted, fontSize:8, bold:true, charSpacing:2 });
    s.addText('DreamSave', { x:0.5, y:1.0, w:9, h:1.4, color:C.white, fontSize:52, bold:true, fontFace:'Calibri' });
    s.addText('Churn Prevention & Renewal Retention Engine — Saving Members Before They Cancel', { x:0.5, y:2.3, w:8.5, h:0.5, color:C.accent, fontSize:18, fontFace:'Calibri' });
    s.addText('Ajay Avaghade · Product Manager', { x:0.5, y:3.0, w:5, h:0.3, color:C.muted, fontSize:12 });
    s.addShape('rect', { x:6.8, y:1.8, w:2.8, h:1.4, fill:{ color:C.accent, transparency:88 }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
    s.addText('−30%', { x:6.8, y:1.9, w:2.8, h:0.7, color:C.white, fontSize:36, bold:true, align:'center' });
    s.addText('churn rate at renewal with save intervention', { x:6.8, y:2.5, w:2.8, h:0.5, color:C.accent, fontSize:10, align:'center' });
  }

  // SLIDE 2 — THE PROBLEM
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:0.06, h:'100%', fill:{ color:C.accent } });
    s.addText('THE PROBLEM', { x:0.3, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Dreampass Loses Members at Renewal Who Never Experienced Value — Invisible Until It\'s Too Late', { x:0.3, y:0.55, w:9.4, h:0.9, color:C.white, fontSize:23, bold:true, fontFace:'Calibri' });
    const cols = [
      { icon:'💤', val:'0/6', lbl:'Benefits used by at-risk member', sub:'Members who use zero benefits in a 12-month cycle are 3.4× more likely to cancel at renewal' },
      { icon:'👻', val:'47 days', lbl:'Avg gap since last app open', sub:'At-risk members go dark well before renewal — there is no proactive outreach today to re-engage' },
      { icon:'🕐', val:'D-7', lbl:'When churn is detected today', sub:'The cancel flow is the first point of intervention — by then, the decision is already made' },
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
    s.addText('Root cause: Dreampass has no churn early-warning system. The platform sees a member cancel — not a member at risk of cancelling. The 45-day window before renewal is completely unutilised.', { x:0.5, y:4.5, w:9.0, h:0.7, color:C.white, fontSize:10, italic:true });
  }

  // SLIDE 3 — THE INSIGHT
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.lgray } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('THE INSIGHT', { x:0.4, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Churn Is Not a Cancel Decision — It\'s a 45-Day Drift. Intervene During the Drift.', { x:0.4, y:0.55, w:9.2, h:0.8, color:C.dark, fontSize:24, bold:true, fontFace:'Calibri' });
    s.addShape('rect', { x:0.4, y:1.5, w:4.4, h:3.3, fill:{ color:'FEF2F2' }, line:{ color:'FECACA', width:1 }, shadow:makeShadow() });
    s.addText('❌  Without DreamSave', { x:0.55, y:1.65, w:4.1, h:0.35, color:'991B1B', fontSize:11, bold:true });
    ['0 benefits used in 11-month cycle — no alert fired','App not opened in 47 days — no re-engagement push','Renewal page viewed 2× but not completed — no save offer','Cancel flow initiates — first intervention is too late','Member cancels: "I got no value from this"'].forEach((t,i)=>{ s.addText(`• ${t}`, { x:0.6, y:2.1+i*0.42, w:4.1, h:0.38, color:'7F1D1D', fontSize:10 }); });
    s.addShape('rect', { x:5.1, y:1.5, w:4.6, h:3.3, fill:{ color:'F5F3FF' }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
    s.addText('✅  With DreamSave', { x:5.25, y:1.65, w:4.2, h:0.35, color:C.accent, fontSize:11, bold:true });
    ['D-45: Push "4 lounge visits expiring — here\'s where to use them"','D-30: Email "Your unused AED 1,134 in travel value expires soon"','D-18: Home screen save offer: pause / discount / downgrade','D-7: Cancel gate shows full value-at-stake screen before allowing cancel','65% of members shown save offer accept — no cancel needed'].forEach((t,i)=>{ s.addText(`• ${t}`, { x:5.25, y:2.1+i*0.42, w:4.3, h:0.38, color:C.accent, fontSize:10 }); });
    s.addShape('rect', { x:0.4, y:5.0, w:9.2, h:0.7, fill:{ color:C.accent, transparency:85 }, line:{ color:C.accent, width:1 } });
    s.addText('Key insight: The pause option converts cancel intent into dormancy rather than churn. The member keeps all coins and benefits frozen — zero net cost to Dreampass, no revenue lost, and the member returns for peak travel season.', { x:0.6, y:5.08, w:8.8, h:0.55, color:C.dark, fontSize:10, italic:true });
  }

  // SLIDE 4 — THE MECHANIC
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.hero } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.gold } });
    s.addText('THE MECHANIC', { x:0.4, y:0.25, w:9, h:0.25, color:C.gold, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Four-Rung Intervention Ladder + Personalised Save Options', { x:0.4, y:0.55, w:9, h:0.6, color:C.white, fontSize:26, bold:true, fontFace:'Calibri' });
    const steps = [
      { n:'D-45', title:'Benefit Nudge', body:'Risk score fires: 0 benefits used + 45 days to renewal. Push: "Your 4 lounge visits expire soon — use on your next departure."' },
      { n:'D-30', title:'Value Summary', body:'Email + in-app banner: "AED 1,134 in unused benefits." Coins balance shown. Deep-link to benefit discovery. Personalised by usage profile.' },
      { n:'D-18', title:'Save Offer Surface', body:'Home screen card: "We have a save offer for you." 3 personalised options: pause 60 days / 30% renewal discount / Dreampass Lite downgrade.' },
      { n:'D-7', title:'Cancel Gate', body:'If cancel flow initiated: full value-at-stake screen. What you\'ll lose list (coins, lounge, eSIM, hotel, insurance). Save offer CTA prominently above cancel button.' },
    ];
    steps.forEach((st,i)=>{
      const x = 0.25 + i*2.4;
      s.addShape('rect', { x, y:1.3, w:2.2, h:3.5, fill:{ color:'0D071A' }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
      s.addShape('rect', { x, y:1.3, w:2.2, h:0.45, fill:{ color:C.accent } });
      s.addText(st.n, { x, y:1.35, w:2.2, h:0.35, color:C.white, fontSize:14, bold:true, align:'center' });
      s.addText(st.title, { x, y:1.85, w:2.2, h:0.4, color:C.accent, fontSize:11, bold:true, align:'center' });
      s.addText(st.body, { x:x+0.12, y:2.32, w:2.0, h:2.35, color:C.muted, fontSize:9 });
    });
    s.addShape('rect', { x:0.25, y:5.0, w:9.5, h:0.7, fill:{ color:'0D071A' }, line:{ color:C.gold, width:1, transparency:60 } });
    s.addText('PhonePe proof: Segment-to-channel model drove +15% D30 retention. The 3-category breadth threshold (users who engage with 3+ categories retain ~70% better) maps directly to DreamSave\'s multi-benefit engagement scoring.', { x:0.4, y:5.1, w:9.2, h:0.52, color:C.gold, fontSize:9.5, italic:true });
  }

  // SLIDE 5 — THE PRODUCT
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.lgray } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('THE PRODUCT', { x:0.4, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('5 Screen States — Churn Signal to Saved Member', { x:0.4, y:0.55, w:9.2, h:0.6, color:C.dark, fontSize:22, bold:true, fontFace:'Calibri' });
    const cards = [
      { num:'01', title:'At-Risk Home', body:'Renewal banner: 18 days left. Risk alert: 0/6 benefits used. 1,840 coins balance with expiry warning. Usage grid: 0 used / 6 available / 18 days.' },
      { num:'02', title:'Cancel Intent Gate', body:'Cancel initiated. Value-at-stake list: lounge, eSIM, hotel, coins, insurance. Preview of save offer. "See save offers" CTA above cancel button.' },
      { num:'03', title:'Save Options', body:'3 personalised options: Pause 60 days (best for you) / Renew at 30% off (AED 349) / Downgrade to Lite (AED 149). Valid 48 hours.' },
      { num:'04', title:'Pause Confirmed', body:'Pause summary: 60 days, AED 0 charged, all benefits frozen. Benefits kept list. Resume early option. "See you when you\'re back" state.' },
      { num:'05', title:'PM Risk View', body:'Dark PM screen. Churn risk score grid (74/100 high). Signal table: HIGH/MED badges + intervention action. 4-metric outcome grid.' },
    ];
    cards.forEach((c,i)=>{
      const x = 0.25 + i*1.93;
      s.addShape('rect', { x, y:1.3, w:1.78, h:4.0, fill:{ color:C.white }, line:{ color:C.accent, width:1 }, shadow:makeShadow() });
      s.addShape('rect', { x, y:1.3, w:1.78, h:0.4, fill:{ color:C.accent, transparency:85 } });
      s.addText(c.num, { x, y:1.35, w:1.78, h:0.32, color:C.accent, fontSize:13, bold:true, align:'center' });
      s.addText(c.title, { x:x+0.08, y:1.8, w:1.62, h:0.4, color:C.dark, fontSize:11, bold:true });
      s.addText(c.body, { x:x+0.08, y:2.28, w:1.62, h:2.8, color:'475569', fontSize:9.5 });
    });
    s.addText('Prototype: dreampass-save-prototype.html · All 5 states interactive', { x:0.4, y:5.5, w:9.2, h:0.3, color:C.muted, fontSize:9.5, italic:true });
  }

  // SLIDE 6 — IMPACT & ROI
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('IMPACT & ROI', { x:0.4, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Churn Prevention as a Zero-Net-Cost Retention Strategy', { x:0.4, y:0.55, w:9.2, h:0.6, color:C.white, fontSize:22, bold:true, fontFace:'Calibri' });
    const left = [
      { val:'−30%', lbl:'Churn rate at renewal', sub:'Members shown DreamSave intervention vs control' },
      { val:'65%', lbl:'Save offer acceptance rate', sub:'Of members who reach save offer screen' },
      { val:'+15%', lbl:'D30 retention benchmark', sub:'PhonePe segment-to-channel model (applied)' },
    ];
    const right = [
      { val:'AED 0', lbl:'Net cost of pause option', sub:'Coins freeze, no benefit consumed, no charge — zero cost to Dreampass' },
      { val:'AED 89', lbl:'Max cost of discount save option', sub:'30% off AED 499 renewal = AED 150 off vs AED 499 retained' },
      { val:'76%', lbl:'Renewal rate at 2+ benefits used', sub:'vs 38% for zero-benefit members — engagement = retention' },
    ];
    [[left,'Retention Impact'],[right,'Unit Economics of Saving']].forEach(([items,title],col)=>{
      const xBase = col===0 ? 0.4 : 5.1;
      s.addShape('rect', { x:xBase, y:1.3, w:4.5, h:3.9, fill:{ color:C.hero }, line:{ color:C.accent, width:1, transparency:70 } });
      s.addText(title, { x:xBase+0.15, y:1.4, w:4.2, h:0.35, color:C.accent, fontSize:10, bold:true, charSpacing:1 });
      items.forEach((item,i)=>{
        s.addText(item.val, { x:xBase+0.15, y:1.85+i*1.1, w:4.2, h:0.55, color:col===0?C.accent:C.gold, fontSize:30, bold:true });
        s.addText(item.lbl, { x:xBase+0.15, y:2.37+i*1.1, w:4.2, h:0.3, color:C.white, fontSize:10, bold:true });
        s.addText(item.sub, { x:xBase+0.15, y:2.63+i*1.1, w:4.2, h:0.3, color:C.muted, fontSize:9 });
      });
    });
    s.addText('(Source: Recurly Subscription Churn Report 2023; ProfitWell Churn Reduction Benchmark 2023; Antavo Global Customer Loyalty 2024; PhonePe Internal)', { x:0.4, y:5.45, w:9.2, h:0.28, color:C.muted, fontSize:8.5, italic:true });
  }

  // SLIDE 7 — PROOF OF WORK
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.lgray } });
    s.addShape('rect', { x:0, y:0, w:4.7, h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.accent } });
    s.addText('PROOF OF WORK', { x:0.3, y:0.25, w:9, h:0.25, color:C.accent, fontSize:9, bold:true, charSpacing:3 });
    s.addText('I Built the Retention\nEngine at PhonePe.', { x:0.3, y:0.6, w:4.0, h:0.9, color:C.white, fontSize:22, bold:true, fontFace:'Calibri' });
    const proofs = [
      '7-day activation cliff detection: bank-link + first transaction within day 7 drives 3× better D30 retention — designed the intervention sequence',
      '3-category breadth threshold: users engaging 3+ categories retain ~70% better — used as segment signal for churn risk scoring',
      'Masthead moratorium: negotiated no-ads for first 7 days — CLTV model showing long-term retention value vs week-one ad revenue',
      'ML personalisation: replaced static rule-based targeting with user-level propensity models — 32% marketing burn reduction',
    ];
    proofs.forEach((p,i)=>{
      s.addShape('rect', { x:0.3, y:1.65+i*1.1, w:4.1, h:0.95, fill:{ color:C.hero }, line:{ color:C.accent, width:1, transparency:70 } });
      s.addText(`• ${p}`, { x:0.45, y:1.72+i*1.1, w:3.8, h:0.8, color:C.muted, fontSize:9.5 });
    });
    const maps = [
      '→ DreamSave D-45 churn risk signal detection',
      '→ DreamSave 3-benefit engagement threshold scoring',
      '→ DreamSave pause option (CLTV > short-term revenue)',
      '→ DreamSave personalised save offer (risk-matched intervention)',
    ];
    maps.forEach((m,i)=>{
      s.addShape('rect', { x:5.0, y:1.65+i*1.1, w:4.6, h:0.95, fill:{ color:C.white }, line:{ color:C.accent, width:1 } });
      s.addText(m, { x:5.15, y:1.72+i*1.1, w:4.3, h:0.8, color:C.accent, fontSize:10, bold:true });
    });
    s.addShape('rect', { x:0.3, y:5.45, w:9.4, h:0.55, fill:{ color:C.accent, transparency:90 }, line:{ color:C.accent, width:1 } });
    s.addText('D30 retention +15% at PhonePe is the direct proof point for DreamSave. The signal detection, intervention ladder, and segment-to-channel logic are identical in architecture.', { x:0.5, y:5.52, w:9.0, h:0.4, color:C.dark, fontSize:10.5, bold:true, italic:true });
  }

  // SLIDE 8 — ROLLOUT PLAN
  {
    const s = prs.addSlide();
    s.addShape('rect', { x:0, y:0, w:'100%', h:'100%', fill:{ color:C.dark } });
    s.addShape('rect', { x:0, y:0, w:'100%', h:0.05, fill:{ color:C.gold } });
    s.addText('ROLLOUT PLAN', { x:0.4, y:0.25, w:9, h:0.25, color:C.gold, fontSize:9, bold:true, charSpacing:3 });
    s.addText('Three Phases — Signal, Prove Save Rate, Scale', { x:0.4, y:0.55, w:9.2, h:0.6, color:C.white, fontSize:24, bold:true, fontFace:'Calibri' });
    const phases = [
      { phase:'Phase 1', period:'Month 1–2: Signal', body:'Instrument churn risk scoring on all members renewing in next 90 days. Build risk score model: benefits used, app recency, renewal page visits, booking frequency. Flag high-risk cohort (score >60). Deploy D-45 push nudge to high-risk group. Control: no intervention.', ask:'Data: benefit redemption events, app open events, renewal page visits' },
      { phase:'Phase 2', period:'Month 3–4: Prove Save Rate', body:'Run A/B test: 3-rung intervention ladder (D-45/D-30/D-18) vs D-7 cancel-gate only. Test save offer types: pause vs discount vs downgrade — which converts highest per churn signal type. Target: 65% save rate on members who see offer.', ask:'1 engineer: risk score pipeline + save offer CRM integration' },
      { phase:'Phase 3', period:'Month 5–6: Scale', body:'Deploy DreamSave intervention ladder to all members globally. Auto-trigger personalised save offer based on risk signal type (price sensitivity → discount; low engagement → pause; low travel → Lite). Integrate with DreamActivate — saved members enter reactivation flow.', ask:'Alignment: CRM team, finance (pause mechanics), member ops (pause state management)' },
    ];
    phases.forEach((p,i)=>{
      const x = 0.3 + i*3.2;
      s.addShape('rect', { x, y:1.3, w:3.0, h:4.0, fill:{ color:C.hero }, line:{ color:C.gold, width:1, transparency:70 }, shadow:makeShadow() });
      s.addShape('rect', { x, y:1.3, w:3.0, h:0.38, fill:{ color:C.gold, transparency:80 } });
      s.addText(p.phase, { x, y:1.35, w:3.0, h:0.28, color:C.gold, fontSize:11, bold:true, align:'center' });
      s.addText(p.period, { x:x+0.1, y:1.76, w:2.8, h:0.35, color:C.accent, fontSize:10, bold:true });
      s.addText(p.body, { x:x+0.1, y:2.18, w:2.8, h:2.3, color:C.muted, fontSize:9.5 });
      s.addShape('rect', { x:x+0.1, y:4.6, w:2.8, h:0.55, fill:{ color:'0D071A' }, line:{ color:C.accent, width:1, transparency:60 } });
      s.addText(`Need: ${p.ask}`, { x:x+0.15, y:4.65, w:2.7, h:0.45, color:C.accent, fontSize:8.5, italic:true });
    });
    s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  ajay-avaghade.github.io/Portfolio', { x:0.4, y:5.52, w:9.2, h:0.28, color:C.muted, fontSize:9, align:'center' });
  }

  await prs.writeFile({ fileName:'assets/dreampass-save-deck.pptx' });
  console.log('Done: assets/dreampass-save-deck.pptx');
})();
