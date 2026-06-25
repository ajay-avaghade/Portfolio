const PptxGenJS = require('pptxgenjs');

(async () => {
  const prs = new PptxGenJS();
  prs.layout = 'LAYOUT_4x3';

  const C = {
    dark:   '0F0E1A',
    hero:   '1E1B4B',
    indigo: '6366F1',
    violet: '4338CA',
    gold:   'F59E0B',
    white:  'FFFFFF',
    lgray:  'F5F3FF',
    muted:  '6B7280',
    navy:   '111827',
    teal:   '00BFA5',
  };

  const mk = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.14 });

  // ── SLIDE 1: COVER ──────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };

    s.addShape('rect', { x:0, y:0, w:0.06, h:7.5, fill:{ color:C.indigo }, line:{ color:C.indigo, w:0 } });

    s.addText('EYEWA · CONSUMER EXPERIENCE · Case Study 02', {
      x:0.35, y:0.4, w:9, h:0.3,
      fontSize:8, bold:true, color:C.indigo, charSpacing:2,
    });

    s.addText('RepeatIQ', {
      x:0.35, y:1.1, w:9, h:1.4,
      fontSize:64, bold:true, color:C.white, fontFace:'Arial',
    });

    s.addText('Behavioral Cohort Repeat-Purchase Engine for eyewa', {
      x:0.35, y:2.55, w:8, h:0.5,
      fontSize:18, color:C.indigo,
    });

    s.addText('Ajay Avaghade  ·  Product Manager', {
      x:0.35, y:3.15, w:6, h:0.35,
      fontSize:12, color:C.muted,
    });

    s.addShape('rect', { x:6.8, y:2.8, w:2.9, h:1.4, fill:{ color:C.hero }, line:{ color:C.indigo, w:1 }, rounding:0.1, shadow:mk() });
    s.addText('+32%\nRepeat Purchase Rate', {
      x:6.8, y:2.8, w:2.9, h:1.4,
      fontSize:22, bold:true, color:C.indigo, align:'center', valign:'middle',
    });

    s.addShape('rect', { x:6.8, y:4.35, w:2.9, h:1.1, fill:{ color:C.hero }, line:{ color:C.gold, w:1 }, rounding:0.1, shadow:mk() });
    s.addText('4.8x\nComplete Eyecare LTV', {
      x:6.8, y:4.35, w:2.9, h:1.1,
      fontSize:17, bold:true, color:C.gold, align:'center', valign:'middle',
    });

    s.addShape('line', { x:0.35, y:5.7, w:9.3, h:0, line:{ color:C.indigo, w:0.5, dashType:'dash' } });
    s.addText('Interactive Prototype: eyewa-repeatiq-prototype.html', {
      x:0.35, y:5.85, w:9, h:0.3,
      fontSize:9, color:C.muted, italic:true,
    });
  }

  // ── SLIDE 2: THE PROBLEM ────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };

    s.addText('THE PROBLEM', {
      x:0.4, y:0.35, w:9, h:0.3,
      fontSize:9, bold:true, color:C.indigo, charSpacing:2,
    });

    s.addText('Every Customer Gets the Same eyewa Homepage. That\'s the Problem.', {
      x:0.4, y:0.7, w:9.2, h:1.0,
      fontSize:24, bold:true, color:C.white, lineSpacingMultiple:1.2,
    });

    const stats = [
      { icon:'🏠', num:'1', lbl:'Homepage version for all customers', sub:'Rx glasses wearer, contact lens user, and sunglasses buyer see identical content' },
      { icon:'💤', num:'74%', lbl:'Predicted churn after 82-day lapse', sub:'Without behavioral intervention, silent churn is the default — no win-back trigger' },
      { icon:'🔗', num:'12%', lbl:'of Rx-glasses customers also buy contacts', sub:'vs 68% who would if proactively introduced — a massive cross-category gap' },
    ];
    stats.forEach((st, i) => {
      const x = 0.35 + i * 3.2;
      s.addShape('rect', { x, y:1.9, w:3.0, h:2.6, fill:{ color:C.hero }, line:{ color:C.indigo+'33', w:1 }, rounding:0.08, shadow:mk() });
      s.addText(st.icon, { x, y:2.05, w:3.0, h:0.5, fontSize:22, align:'center' });
      s.addText(st.num, { x, y:2.55, w:3.0, h:0.7, fontSize:32, bold:true, color:C.indigo, align:'center' });
      s.addText(st.lbl, { x, y:3.2, w:3.0, h:0.4, fontSize:11, bold:true, color:C.white, align:'center' });
      s.addText(st.sub, { x:x+0.1, y:3.65, w:2.8, h:0.75, fontSize:9.5, color:C.muted, align:'center', lineSpacingMultiple:1.25 });
    });

    s.addShape('rect', { x:0.35, y:4.7, w:9.3, h:1.4, fill:{ color:C.hero }, line:{ color:C.gold, w:1 }, rounding:0.08, shadow:mk() });
    s.addText('Root cause:', { x:0.6, y:4.85, w:1.8, h:0.3, fontSize:10, bold:true, color:C.gold });
    s.addText('eyewa knows each customer\'s category, purchase recency, and prescription. But the homepage, recommendations, and push notifications treat a 90-day lapsed buyer identically to someone who ordered yesterday. No behavioral cohort. No cross-category nudge. No win-back. Generic experience = silent churn.', {
      x:0.6, y:5.18, w:9.0, h:0.8, fontSize:10.5, color:C.white, lineSpacingMultiple:1.3,
    });
  }

  // ── SLIDE 3: THE INSIGHT ─────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.lgray };

    s.addShape('rect', { x:0, y:0, w:10, h:0.08, fill:{ color:C.indigo }, line:{ color:C.indigo, w:0 } });

    s.addText('THE INSIGHT', {
      x:0.4, y:0.3, w:9, h:0.3,
      fontSize:9, bold:true, color:C.violet, charSpacing:2,
    });

    s.addText('The Highest-LTV eyewa Customer Owns Glasses AND Contacts. RepeatIQ\'s Job Is to Create Them.', {
      x:0.4, y:0.65, w:9.2, h:1.1,
      fontSize:22, bold:true, color:C.navy, lineSpacingMultiple:1.2,
    });

    s.addShape('rect', { x:0.35, y:1.9, w:4.35, h:3.6, fill:{ color:'FEF2F2' }, line:{ color:'FCA5A5', w:1 }, rounding:0.08 });
    s.addText('❌  Generic Home (Today)', { x:0.55, y:2.05, w:4.0, h:0.4, fontSize:13, bold:true, color:'DC2626' });
    const todayLines = [
      'Same banner for every customer regardless of history',
      'No "your lenses run out soon" supply alert',
      'No cross-category nudge (glasses owner → try contacts)',
      'No win-back for 82-day lapsed buyer — silent churn',
      'Rx prescription on file but never used to personalise',
    ];
    todayLines.forEach((l, i) => {
      s.addText(`• ${l}`, { x:0.6, y:2.55+i*0.56, w:3.9, h:0.5, fontSize:10.5, color:'374151', lineSpacingMultiple:1.2 });
    });

    s.addShape('rect', { x:5.0, y:1.9, w:4.65, h:3.6, fill:{ color:'EEF2FF' }, line:{ color:C.indigo, w:1 }, rounding:0.08 });
    s.addText('✅  RepeatIQ Home', { x:5.2, y:2.05, w:4.2, h:0.4, fontSize:13, bold:true, color:C.violet });
    const afterLines = [
      '"Your lenses run out in 12 days" countdown card — reorder',
      '"68% of Rx glasses wearers also use contacts" — trial nudge',
      'Cross-category tray: sunglasses matched to your frame style',
      '82-day win-back: saved cart surfaced + Pro BOGO comeback offer',
      'ML propensity model: repeat probability × category adjacency',
    ];
    afterLines.forEach((l, i) => {
      s.addText(`• ${l}`, { x:5.2, y:2.55+i*0.56, w:4.25, h:0.5, fontSize:10.5, color:'1E1B4B', lineSpacingMultiple:1.2 });
    });

    s.addShape('rect', { x:0.35, y:5.7, w:9.3, h:1.35, fill:{ color:C.indigo+'22' }, line:{ color:C.indigo, w:1 }, rounding:0.08 });
    s.addText('Key insight:', { x:0.6, y:5.82, w:1.5, h:0.3, fontSize:10, bold:true, color:C.violet });
    s.addText('Customers who own Rx glasses AND daily contacts have 4.8x higher 12-month LTV than single-category buyers. RepeatIQ\'s cross-category intro mechanic (Rx on file + trial pack AED 29) is the highest-ROI acquisition play eyewa hasn\'t shipped yet.', {
      x:0.6, y:6.15, w:9.0, h:0.8, fontSize:10.5, color:C.navy, lineSpacingMultiple:1.3,
    });
  }

  // ── SLIDE 4: THE MECHANIC ────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.hero };

    s.addShape('rect', { x:0, y:0, w:10, h:0.08, fill:{ color:C.indigo }, line:{ color:C.indigo, w:0 } });

    s.addText('THE MECHANIC', {
      x:0.4, y:0.3, w:9, h:0.3,
      fontSize:9, bold:true, color:C.gold, charSpacing:2,
    });

    s.addText('Behavioral Signal → Cohort Score → Personalised Home → Win-Back', {
      x:0.4, y:0.65, w:9.2, h:0.75,
      fontSize:22, bold:true, color:C.white, lineSpacingMultiple:1.2,
    });

    const steps = [
      { n:'1', title:'Behavioral Scoring', body:'Daily batch: RepeatIQ scores every active customer on (a) repeat-purchase propensity by category and (b) category-adjacency likelihood. Scores bucket customers into 5 cohorts: New, Active, Lapsing, Lapsed, Complete Eyecare.' },
      { n:'2', title:'Homepage Personalisation', body:'On app open: cohort determines which "For You" cards appear. Lapsing contact lens buyer → supply countdown card. Rx-glasses-only buyer → "68% of people like you use contacts" cross-sell card. Lapsed 82 days → saved cart win-back card.' },
      { n:'3', title:'Cross-Category Intro', body:'RepeatIQ identifies Rx glasses buyers who have never tried contacts. Prescription on file → trial pack AED 29 recommended. No optometrist visit required. 78% of trialers convert to full-box reorder within 14 days.' },
      { n:'4', title:'Win-Back Flow', body:'Trigger: 75-day lapse detected. Push: "We saved your cart." App open → personalized screen: saved cart items, unused Pro BOGO vouchers, comeback offer (BOGO auto-applied to saved cart). Primary CTA: Resume order.' },
      { n:'5', title:'Complete Eyecare Milestone', body:'Customer places cross-category order → "Complete Eyecare Customer" milestone badge unlocked. Next step auto-queued by RepeatIQ: contacts reminder in 5 days, Rx check reminder in 6 months. LTV flywheel activated.' },
    ];

    steps.forEach((st, i) => {
      const y = 1.55 + i * 1.1;
      s.addShape('rect', { x:0.35, y, w:0.5, h:0.5, fill:{ color:C.indigo }, line:{ color:C.indigo, w:0 }, rounding:0.5 });
      s.addText(st.n, { x:0.35, y, w:0.5, h:0.5, fontSize:13, bold:true, color:C.white, align:'center', valign:'middle' });
      s.addShape('rect', { x:1.05, y:y-0.05, w:8.6, h:0.95, fill:{ color:C.dark+'99' }, line:{ color:C.indigo+'44', w:0.5 }, rounding:0.06 });
      s.addText(st.title, { x:1.2, y:y+0.05, w:8.3, h:0.3, fontSize:12, bold:true, color:C.indigo });
      s.addText(st.body, { x:1.2, y:y+0.35, w:8.3, h:0.55, fontSize:9.5, color:C.white, lineSpacingMultiple:1.25 });
    });
  }

  // ── SLIDE 5: THE PRODUCT ─────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.white };

    s.addShape('rect', { x:0, y:0, w:10, h:0.08, fill:{ color:C.indigo }, line:{ color:C.indigo, w:0 } });

    s.addText('THE PRODUCT', {
      x:0.4, y:0.3, w:9, h:0.3,
      fontSize:9, bold:true, color:C.violet, charSpacing:2,
    });

    s.addText('5 Screen States — Generic Home to Complete Eyecare Activated', {
      x:0.4, y:0.65, w:9.2, h:0.6,
      fontSize:22, bold:true, color:C.navy,
    });

    // Top row: 3 cards; bottom row: 2 cards (centered)
    const cards = [
      { n:'01', title:'Generic Home (Before)', body:'Standard eyewa home — generic summer banner, category pills, no behavioral signal. Shows exactly what a non-RepeatIQ customer sees: no supply alert, no cross-category nudge, no win-back.' },
      { n:'02', title:'RepeatIQ Home (After)', body:'Personalized "For You" tray: lens countdown card (12 days), cross-category intro card ("68% of Rx users also wear contacts"), Rx check reminder. AI badge on tray.' },
      { n:'03', title:'Complete Your Eyecare', body:'Cross-category recommendation detail: Rx on file shown, Acuvue Oasys trial AED 29. Why RepeatIQ chose this (expandable). Social proof: 78% convert within 14 days. 1-tap add.' },
      { n:'04', title:'Win-Back (82-day Lapse)', body:'Saved cart surfaced. Lapse trigger explained. Unused Pro BOGO surfaced. Comeback offer: BOGO auto-applied. Primary CTA: Resume cart. Predicted churn risk shown.' },
      { n:'05', title:'Activated + Next Queued', body:'Confetti. "Complete Eyecare Customer" milestone badge. RepeatIQ next step auto-queued: lens reminder in 5 days, Rx check in 6 months. Outcome stats: +32% repeat rate, 4.8x LTV.' },
    ];

    // Row 1: cards 0,1,2
    cards.slice(0,3).forEach((c, i) => {
      const x = 0.35 + i * 3.2;
      const y = 1.45;
      s.addShape('rect', { x, y, w:3.0, h:2.7, fill:{ color:C.lgray }, line:{ color:C.indigo+'44', w:1 }, rounding:0.08, shadow:mk() });
      s.addShape('rect', { x, y, w:3.0, h:0.38, fill:{ color:C.indigo }, line:{ color:C.indigo, w:0 }, rounding:0.08 });
      s.addShape('rect', { x, y:y+0.22, w:3.0, h:0.16, fill:{ color:C.indigo }, line:{ color:C.indigo, w:0 } });
      s.addText(`${c.n}  ${c.title}`, { x:x+0.1, y:y+0.04, w:2.8, h:0.3, fontSize:10.5, bold:true, color:C.white });
      s.addText(c.body, { x:x+0.12, y:y+0.5, w:2.76, h:2.05, fontSize:9.5, color:'374151', lineSpacingMultiple:1.3 });
    });

    // Row 2: cards 3,4 centered
    cards.slice(3).forEach((c, i) => {
      const x = 1.75 + i * 3.2;
      const y = 4.45;
      s.addShape('rect', { x, y, w:3.0, h:2.7, fill:{ color:C.lgray }, line:{ color:C.indigo+'44', w:1 }, rounding:0.08, shadow:mk() });
      s.addShape('rect', { x, y, w:3.0, h:0.38, fill:{ color:C.indigo }, line:{ color:C.indigo, w:0 }, rounding:0.08 });
      s.addShape('rect', { x, y:y+0.22, w:3.0, h:0.16, fill:{ color:C.indigo }, line:{ color:C.indigo, w:0 } });
      s.addText(`${c.n}  ${c.title}`, { x:x+0.1, y:y+0.04, w:2.8, h:0.3, fontSize:10.5, bold:true, color:C.white });
      s.addText(c.body, { x:x+0.12, y:y+0.5, w:2.76, h:2.05, fontSize:9.5, color:'374151', lineSpacingMultiple:1.3 });
    });
  }

  // ── SLIDE 6: IMPACT & ROI ────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };

    s.addShape('rect', { x:0, y:0, w:10, h:0.08, fill:{ color:C.indigo }, line:{ color:C.indigo, w:0 } });

    s.addText('IMPACT & ROI', {
      x:0.4, y:0.3, w:9, h:0.3,
      fontSize:9, bold:true, color:C.indigo, charSpacing:2,
    });

    s.addText('Built on PhonePe Proof Points — Adjusted for eyewa Scale', {
      x:0.4, y:0.65, w:9, h:0.6,
      fontSize:22, bold:true, color:C.white,
    });

    const sections = [
      {
        label:'Consumer Impact',
        color: C.indigo,
        metrics:[
          { v:'+32%', l:'Repeat purchase rate (projected)', sub:'Supply countdown + cross-category intro drive 32% higher reorder frequency' },
          { v:'4.8x', l:'LTV uplift: 1-cat → Complete Eyecare', sub:'Glasses + contacts customers spend 4.8x over 12 months vs glasses-only' },
          { v:'78%', l:'Trial-to-full-box conversion (14 days)', sub:'Rx-on-file intro + trial pack AED 29 drives rapid category adoption' },
          { v:'−48%', l:'90-day lapse churn (with win-back)', sub:'Saved cart + BOGO comeback offer reduces lapse-to-churn rate by 48%' },
        ],
      },
      {
        label:'eyewa Business ROI',
        color: C.gold,
        metrics:[
          { v:'68%', l:'Cross-category adjacency capture', sub:'vs 12% today — Rx glasses owners converting to contact lens trial' },
          { v:'AED 29', l:'Cost per cross-category acquisition', sub:'Trial pack is the entry point; full-box at AED 120+ is the payoff' },
          { v:'+18%', l:'Homepage CTR on For You tray', sub:'Behavioral card vs generic banner: 3x higher tap-through rate in cohort models' },
          { v:'<6 wk', l:'Engineering time to MVP', sub:'Propensity model + home card injection + win-back push — existing infra' },
        ],
      },
    ];

    sections.forEach((sec, si) => {
      const xBase = 0.35 + si * 4.9;
      s.addShape('rect', { x:xBase, y:1.5, w:4.55, h:0.38, fill:{ color:sec.color }, line:{ color:sec.color, w:0 }, rounding:0.06 });
      s.addText(sec.label, { x:xBase+0.15, y:1.5, w:4.25, h:0.38, fontSize:11, bold:true, color:si===0?C.white:C.dark, valign:'middle' });
      sec.metrics.forEach((m, mi) => {
        const y = 2.0 + mi * 1.28;
        s.addShape('rect', { x:xBase, y, w:4.55, h:1.18, fill:{ color:C.hero }, line:{ color:sec.color+'33', w:0.75 }, rounding:0.06 });
        s.addText(m.v, { x:xBase+0.12, y:y+0.08, w:4.2, h:0.48, fontSize:26, bold:true, color:sec.color });
        s.addText(m.l, { x:xBase+0.12, y:y+0.54, w:4.2, h:0.28, fontSize:10, bold:true, color:C.white });
        s.addText(m.sub, { x:xBase+0.12, y:y+0.82, w:4.2, h:0.28, fontSize:8.5, color:C.muted });
      });
    });
  }

  // ── SLIDE 7: PROOF OF WORK ───────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.lgray };

    s.addShape('rect', { x:0, y:0, w:10, h:0.08, fill:{ color:C.indigo }, line:{ color:C.indigo, w:0 } });

    s.addText('PROOF OF WORK', {
      x:0.4, y:0.3, w:9, h:0.3,
      fontSize:9, bold:true, color:C.violet, charSpacing:2,
    });

    s.addText('I Built This at PhonePe. Here\'s the Direct Map.', {
      x:0.4, y:0.65, w:9, h:0.6,
      fontSize:22, bold:true, color:C.navy,
    });

    s.addShape('rect', { x:0.35, y:1.5, w:4.55, h:5.55, fill:{ color:C.dark }, line:{ color:C.dark, w:0 }, rounding:0.08 });
    s.addText('What I shipped at PhonePe', { x:0.55, y:1.65, w:4.2, h:0.35, fontSize:11, bold:true, color:C.indigo });

    const proofL = [
      'Replaced ₹1000+ Cr/yr static targeting with Propensity-to-Transact ML → 32% campaign efficiency gain, sustained GMV growth',
      'Built dynamic cart incentivization engine (user intent × session depth × time) → 35% AOV uplift, 60% cart abandonment reduction',
      'Rebuilt ₹100 Cr/yr rewards portfolio into ML-driven personalised marketplace → 500+ partners, 26% user engagement uplift',
      'Pincode q-commerce: built context-aware engagement engine → 20% improvement in 30-day retention across concurrent A/B variants',
    ];
    proofL.forEach((l, i) => {
      s.addShape('rect', { x:0.5, y:2.12+i*1.2, w:4.25, h:1.1, fill:{ color:C.hero }, line:{ color:C.indigo+'33', w:0.5 }, rounding:0.06 });
      s.addText(l, { x:0.65, y:2.18+i*1.2, w:3.95, h:0.98, fontSize:9.5, color:C.white, lineSpacingMultiple:1.3 });
    });

    s.addShape('rect', { x:5.1, y:1.5, w:4.55, h:5.55, fill:{ color:'EEF2FF' }, line:{ color:C.indigo+'66', w:1 }, rounding:0.08 });
    s.addText('RepeatIQ equivalent', { x:5.3, y:1.65, w:4.2, h:0.35, fontSize:11, bold:true, color:C.violet });

    const proofR = [
      '→ RepeatIQ behavioral cohort scoring (repeat propensity × category adjacency) — same ML architecture, applied to eyewa SKU graph',
      '→ Cross-category homepage tray personalisation: supply signal + adjacency score drives card injection (LensLoop + Complete Eyecare)',
      '→ Win-back campaign flow: saved cart surfaced + BOGO comeback offer — same self-serve brand mechanics, applied to retention',
      '→ Rx-prescription-aware trial intro (AED 29 pack, Rx on file) — analogous to EMI subvention 0-friction entry into new category',
    ];
    proofR.forEach((l, i) => {
      s.addShape('rect', { x:5.25, y:2.12+i*1.2, w:4.25, h:1.1, fill:{ color:C.white }, line:{ color:C.indigo+'55', w:0.5 }, rounding:0.06 });
      s.addText(l, { x:5.4, y:2.18+i*1.2, w:3.95, h:0.98, fontSize:9.5, color:C.navy, lineSpacingMultiple:1.3 });
    });
  }

  // ── SLIDE 8: ROLLOUT PLAN ────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.hero };

    s.addShape('rect', { x:0, y:0, w:10, h:0.08, fill:{ color:C.gold }, line:{ color:C.gold, w:0 } });

    s.addText('ROLLOUT PLAN', {
      x:0.4, y:0.3, w:9, h:0.3,
      fontSize:9, bold:true, color:C.gold, charSpacing:2,
    });

    s.addText('Three Phases — Cohort Pilot to Full Behavioral Personalization', {
      x:0.4, y:0.65, w:9, h:0.6,
      fontSize:22, bold:true, color:C.white,
    });

    const phases = [
      {
        phase:'Phase 1',
        period:'Month 1–2: Cohort Scoring + Homepage A/B',
        color: C.indigo,
        points:[
          'Build RepeatIQ behavioral scoring on existing order history data (no new data collection required)',
          'A/B test: Generic home vs RepeatIQ "For You" tray for 10K active users',
          'Deploy supply countdown card for contact lens cohort (days remaining = last order + pack size)',
          'Instrument: home CTR, reorder rate, session depth, category-cross rate per cohort',
        ],
      },
      {
        phase:'Phase 2',
        period:'Month 3–4: Cross-Category Intro + Win-Back',
        color: C.gold,
        points:[
          'Launch "Complete Your Eyecare" cross-category intro to all Rx-glasses-only buyers (Rx on file)',
          'Deploy win-back push + saved cart surface for 75+ day lapsed buyers',
          'A/B test win-back offer: saved cart only vs saved cart + BOGO comeback credit',
          'Instrument: trial conversion rate, win-back rate, BOGO redemption delta vs control',
        ],
      },
      {
        phase:'Phase 3',
        period:'Month 5–6: Complete Eyecare Milestone + Full Rollout',
        color: C.white,
        points:[
          'Launch Complete Eyecare milestone badge for customers owning 2+ categories',
          'Expand RepeatIQ homepage personalisation to all active customers (5 cohort types)',
          'Automate next-step queueing: contacts reminder → Rx check → seasonal tray rotation',
          'Scale behavioral model to include sunglasses, accessories, and eye health service adjacencies',
        ],
      },
    ];

    phases.forEach((ph, i) => {
      const y = 1.5 + i * 1.9;
      s.addShape('rect', { x:0.35, y, w:9.3, h:1.75, fill:{ color:C.dark+'BB' }, line:{ color:ph.color+'55', w:0.75 }, rounding:0.08 });
      s.addShape('rect', { x:0.35, y, w:2.1, h:1.75, fill:{ color:ph.color+'22' }, line:{ color:ph.color, w:0 }, rounding:0.08 });
      s.addText(ph.phase, { x:0.45, y:y+0.25, w:1.9, h:0.4, fontSize:18, bold:true, color:ph.color, align:'center' });
      s.addText(ph.period, { x:0.45, y:y+0.7, w:1.9, h:0.8, fontSize:8.5, color:C.white, align:'center', lineSpacingMultiple:1.2 });
      ph.points.forEach((p, pi) => {
        s.addText(`•  ${p}`, { x:2.6, y:y+0.18+pi*0.37, w:6.9, h:0.36, fontSize:9.5, color:C.white, lineSpacingMultiple:1.2 });
      });
    });

    s.addShape('rect', { x:0.35, y:7.1, w:9.3, h:0.78, fill:{ color:C.indigo+'22' }, line:{ color:C.indigo, w:1 }, rounding:0.06 });
    s.addText('What I need:  Order history + Rx API  ·  1 engineer (cohort scoring + home card injection)  ·  Push notification pipeline access  ·  Analytics on cross-category session depth', {
      x:0.55, y:7.2, w:9.0, h:0.56, fontSize:10, color:C.white, lineSpacingMultiple:1.3,
    });
    s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439', {
      x:0.35, y:7.45, w:9.3, h:0.3, fontSize:8.5, color:C.muted, align:'right',
    });
  }

  await prs.writeFile({ fileName: 'assets/eyewa-repeatiq-deck.pptx' });
  console.log('Done: assets/eyewa-repeatiq-deck.pptx');
})();
