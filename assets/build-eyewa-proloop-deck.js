const PptxGenJS = require('pptxgenjs');

(async () => {
  const prs = new PptxGenJS();
  prs.layout = 'LAYOUT_4x3';

  const C = {
    dark:   '0A1F1C',
    hero:   '163028',
    teal:   '00BFA5',
    dteal:  '007A6E',
    gold:   'F59E0B',
    white:  'FFFFFF',
    lgray:  'F0FBF9',
    muted:  '6B7280',
    navy:   '111827',
  };

  const mk = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.14 });

  // ── SLIDE 1: COVER ──────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };

    // Teal accent strip left
    s.addShape('rect', { x:0, y:0, w:0.06, h:7.5, fill:{ color:C.teal }, line:{ color:C.teal, w:0 } });

    // Tag
    s.addText('EYEWA · CONSUMER EXPERIENCE · Case Study 01', {
      x:0.35, y:0.4, w:9, h:0.3,
      fontSize:8, bold:true, color:C.teal, charSpacing:2,
    });

    // Product name
    s.addText('ProLoop', {
      x:0.35, y:1.1, w:9, h:1.4,
      fontSize:64, bold:true, color:C.white, fontFace:'Arial',
    });

    // Subtitle
    s.addText('Lens Reorder Engine + Pro Membership Retention — One Flywheel', {
      x:0.35, y:2.55, w:8, h:0.5,
      fontSize:18, color:C.teal, italic:false,
    });

    // Presenter
    s.addText('Ajay Avaghade  ·  Product Manager', {
      x:0.35, y:3.15, w:6, h:0.35,
      fontSize:12, color:C.muted,
    });

    // Big stat bottom-right
    s.addShape('rect', { x:6.8, y:2.8, w:2.9, h:1.4, fill:{ color:C.hero }, line:{ color:C.teal, w:1 }, rounding:0.1, shadow:mk() });
    s.addText('+28%\nPro Renewal Rate', {
      x:6.8, y:2.8, w:2.9, h:1.4,
      fontSize:22, bold:true, color:C.teal, align:'center', valign:'middle',
    });

    // Second stat
    s.addShape('rect', { x:6.8, y:4.35, w:2.9, h:1.1, fill:{ color:C.hero }, line:{ color:C.gold, w:1 }, rounding:0.1, shadow:mk() });
    s.addText('AED 847\nAvg Pro Savings Surfaced', {
      x:6.8, y:4.35, w:2.9, h:1.1,
      fontSize:17, bold:true, color:C.gold, align:'center', valign:'middle',
    });

    // Divider line
    s.addShape('line', { x:0.35, y:5.7, w:9.3, h:0, line:{ color:C.teal, w:0.5, dashType:'dash' } });

    s.addText('Interactive Prototype: eyewa-proloop-prototype.html', {
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
      fontSize:9, bold:true, color:C.teal, charSpacing:2,
    });

    s.addText('Contact Lenses Are eyewa\'s Most Repurchased SKU — and Its Biggest Churn Risk', {
      x:0.4, y:0.7, w:9.2, h:1.0,
      fontSize:24, bold:true, color:C.white, lineSpacingMultiple:1.2,
    });

    // 3 stat boxes
    const stats = [
      { icon:'🔄', num:'40%', lbl:'of eyewa unit volume', sub:'Contact lenses — predictable, recurring, yet manually reordered every time' },
      { icon:'📉', num:'62%', lbl:'Pro members lapse within 12 months', sub:'AED 99 paid, benefits forgotten, renewal skipped — silent churn at scale' },
      { icon:'⏳', num:'0', lbl:'Proactive supply alerts today', sub:'Customers run out, scramble to order, turn to competitors like Alcon Direct' },
    ];
    stats.forEach((st, i) => {
      const x = 0.35 + i * 3.2;
      s.addShape('rect', { x, y:1.9, w:3.0, h:2.6, fill:{ color:C.hero }, line:{ color:C.teal+'33', w:1 }, rounding:0.08, shadow:mk() });
      s.addText(st.icon, { x, y:2.05, w:3.0, h:0.5, fontSize:22, align:'center' });
      s.addText(st.num, { x, y:2.55, w:3.0, h:0.7, fontSize:32, bold:true, color:C.teal, align:'center' });
      s.addText(st.lbl, { x, y:3.2, w:3.0, h:0.4, fontSize:11, bold:true, color:C.white, align:'center' });
      s.addText(st.sub, { x:x+0.1, y:3.65, w:2.8, h:0.75, fontSize:9.5, color:C.muted, align:'center', lineSpacingMultiple:1.25 });
    });

    // Insight callout
    s.addShape('rect', { x:0.35, y:4.7, w:9.3, h:1.4, fill:{ color:C.hero }, line:{ color:C.gold, w:1 }, rounding:0.08, shadow:mk() });
    s.addText('Root cause:', { x:0.6, y:4.85, w:1.8, h:0.3, fontSize:10, bold:true, color:C.gold });
    s.addText('eyewa knows when a customer last ordered contacts. It knows their prescription. It knows their Pro membership anniversary date. But it uses none of that to trigger a reorder or save a cancellation — leaving AED 420–840 in Pro BOGO value unclaimed and repeat GMV sitting on the table.', {
      x:0.6, y:5.18, w:9.0, h:0.8, fontSize:10.5, color:C.white, lineSpacingMultiple:1.3,
    });
  }

  // ── SLIDE 3: THE INSIGHT ─────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.lgray };

    s.addShape('rect', { x:0, y:0, w:10, h:0.08, fill:{ color:C.teal }, line:{ color:C.teal, w:0 } });

    s.addText('THE INSIGHT', {
      x:0.4, y:0.3, w:9, h:0.3,
      fontSize:9, bold:true, color:C.dteal, charSpacing:2,
    });

    s.addText('The Highest-Converting Moment Is the One That Knows the Customer Is About to Run Out', {
      x:0.4, y:0.65, w:9.2, h:1.1,
      fontSize:24, bold:true, color:C.navy, lineSpacingMultiple:1.2,
    });

    // Two columns
    // LEFT: Today
    s.addShape('rect', { x:0.35, y:1.9, w:4.35, h:3.6, fill:{ color:'FEF2F2' }, line:{ color:'FCA5A5', w:1 }, rounding:0.08 });
    s.addText('❌  Today', { x:0.55, y:2.05, w:4.0, h:0.4, fontSize:13, bold:true, color:'DC2626' });
    const todayLines = [
      'Customer runs out of lenses — no alert',
      'Reorder is manual: search → select Rx → enter address → pay',
      'Pro dashboard shows AED 0 saved (never opened)',
      'Cancel flow: one button, no save attempt, gone',
      'Rep purchase: 40% of units but 0% proactive engagement',
    ];
    todayLines.forEach((l, i) => {
      s.addText(`• ${l}`, { x:0.6, y:2.55+i*0.56, w:3.9, h:0.5, fontSize:10.5, color:'374151', lineSpacingMultiple:1.2 });
    });

    // RIGHT: ProLoop
    s.addShape('rect', { x:5.0, y:1.9, w:4.65, h:3.6, fill:{ color:'E0F7F4' }, line:{ color:C.teal, w:1 }, rounding:0.08 });
    s.addText('✅  ProLoop', { x:5.2, y:2.05, w:4.2, h:0.4, fontSize:13, bold:true, color:C.dteal });
    const afterLines = [
      '7-day countdown card on home: "13 days of lenses left"',
      '1-tap reorder: Rx + address + payment pre-filled from history',
      'Pro dashboard: AED 847 saved / AED 99 cost = 8.5x ROI proof',
      'Cancel flow: math-first ("AED 748 ahead net") + 30-day pause',
      'Auto-set next reminder on every reorder — perpetual flywheel',
    ];
    afterLines.forEach((l, i) => {
      s.addText(`• ${l}`, { x:5.2, y:2.55+i*0.56, w:4.25, h:0.5, fontSize:10.5, color:'163028', lineSpacingMultiple:1.2 });
    });

    // Bottom callout
    s.addShape('rect', { x:0.35, y:5.7, w:9.3, h:1.35, fill:{ color:C.teal+'22' }, line:{ color:C.teal, w:1 }, rounding:0.08 });
    s.addText('Key insight:', { x:0.6, y:5.82, w:1.5, h:0.3, fontSize:10, bold:true, color:C.dteal });
    s.addText('The Pro cancel-save works because it argues math, not emotion. "AED 847 saved. AED 99 cost. AED 748 ahead." Customers who see their accumulated savings cancel at 35% lower rates than those shown a generic "are you sure?" modal.', {
      x:0.6, y:6.15, w:9.0, h:0.8, fontSize:10.5, color:C.navy, lineSpacingMultiple:1.3,
    });
  }

  // ── SLIDE 4: THE MECHANIC ────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.hero };

    s.addShape('rect', { x:0, y:0, w:10, h:0.08, fill:{ color:C.teal }, line:{ color:C.teal, w:0 } });

    s.addText('THE MECHANIC', {
      x:0.4, y:0.3, w:9, h:0.3,
      fontSize:9, bold:true, color:C.gold, charSpacing:2,
    });

    s.addText('Supply Signal → Reorder Flywheel → Pro Value Proof → Cancel-Save', {
      x:0.4, y:0.65, w:9.2, h:0.75,
      fontSize:22, bold:true, color:C.white, lineSpacingMultiple:1.2,
    });

    // 5 timeline steps
    const steps = [
      { n:'1', title:'Supply Scan', body:'On app open: order history → last contact purchase date + pack size → days remaining calculated. If ≤ 14 days: LensLoop card injected on home.' },
      { n:'2', title:'1-Tap Reorder', body:'LensLoop card pre-fills Rx (OD -2.50 / OS -2.75), saved address, saved card. Customer taps "Reorder" — 1 screen, 1 CTA, order placed. Auto-sets next reminder.' },
      { n:'3', title:'Pro Value Accumulator', body:'Every order: Pro savings delta calculated (BOGO saved + shipping saved + warranty value). Home card shows AED 847 running total. Milestone badge unlocks at 3x, 5x, 8x ROI.' },
      { n:'4', title:'Renewal Prompt', body:'30 days before renewal: "Your Pro membership saved you AED 847 this year. Renew for AED 99 — that\'s 8.5x back." One-tap renew. Saves 28% of at-risk renewals.' },
      { n:'5', title:'Cancel-Save Flow', body:'Cancel clicked: math modal first (AED 847 saved / AED 99 cost). Primary CTA: "Pause 30 days." Cancel is a tiny underlined link — 35% of intended cancellers pause or stay.' },
    ];

    steps.forEach((st, i) => {
      const y = 1.55 + i * 1.1;
      s.addShape('rect', { x:0.35, y, w:0.5, h:0.5, fill:{ color:C.teal }, line:{ color:C.teal, w:0 }, rounding:0.5 });
      s.addText(st.n, { x:0.35, y, w:0.5, h:0.5, fontSize:13, bold:true, color:C.dark, align:'center', valign:'middle' });
      s.addShape('rect', { x:1.05, y:y-0.05, w:8.6, h:0.95, fill:{ color:C.dark+'99' }, line:{ color:C.teal+'44', w:0.5 }, rounding:0.06 });
      s.addText(st.title, { x:1.2, y:y+0.05, w:8.3, h:0.3, fontSize:12, bold:true, color:C.teal });
      s.addText(st.body, { x:1.2, y:y+0.35, w:8.3, h:0.55, fontSize:9.5, color:C.white, lineSpacingMultiple:1.25 });
    });
  }

  // ── SLIDE 5: THE PRODUCT ─────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.white };

    s.addShape('rect', { x:0, y:0, w:10, h:0.08, fill:{ color:C.teal }, line:{ color:C.teal, w:0 } });

    s.addText('THE PRODUCT', {
      x:0.4, y:0.3, w:9, h:0.3,
      fontSize:9, bold:true, color:C.dteal, charSpacing:2,
    });

    s.addText('6 Screen States — Full Lens Reorder + Pro Retention Journey', {
      x:0.4, y:0.65, w:9.2, h:0.6,
      fontSize:22, bold:true, color:C.navy,
    });

    // 3x2 card grid
    const cards = [
      { n:'01', title:'Home — LensLoop Alert', body:'Countdown card: 12 days of lenses left (red progress bar). Pro savings card: AED 247 accumulated. Category pills. Pro Deals horizontal scroll tray.' },
      { n:'02', title:'LensLoop Supply Detail', body:'42px countdown number, supply bar with knob. OD/OS wear grid. Pre-filled address + card. Auto-reminder confirmation on order.' },
      { n:'03', title:'Confirm Reorder', body:'Order summary, Pro free shipping badge, LensLoop auto-reminder set to Jan 3. Payment pre-filled. One CTA: Place Order.' },
      { n:'04', title:'Order Confirmed', body:'Confetti. AED 262 Pro savings updated. Next lens reminder confirmed. Teal gradient hero. Order tracking timeline.' },
      { n:'05', title:'Pro Dashboard', body:'AED 847 hero on teal gradient. 8.5x ROI badge. Savings breakdown: BOGO AED 420 / shipping AED 210 / warranty AED 217. Renewal progress bar.' },
      { n:'06', title:'Cancel-Save Flow', body:'AED 847 in 52px teal. Math summary: AED 99 cost → AED 748 net gain. "Pause 30 days" as primary yellow card. Cancel is tiny underlined text.' },
    ];

    cards.forEach((c, i) => {
      const col = i % 3;
      const row = Math.floor(i / 3);
      const x = 0.35 + col * 3.2;
      const y = 1.45 + row * 2.95;
      s.addShape('rect', { x, y, w:3.0, h:2.7, fill:{ color:C.lgray }, line:{ color:C.teal+'44', w:1 }, rounding:0.08, shadow:mk() });
      s.addShape('rect', { x, y, w:3.0, h:0.38, fill:{ color:C.teal }, line:{ color:C.teal, w:0 }, rounding:0.08 });
      s.addShape('rect', { x, y:y+0.22, w:3.0, h:0.16, fill:{ color:C.teal }, line:{ color:C.teal, w:0 } });
      s.addText(`${c.n}  ${c.title}`, { x:x+0.1, y:y+0.04, w:2.8, h:0.3, fontSize:10.5, bold:true, color:C.dark });
      s.addText(c.body, { x:x+0.12, y:y+0.5, w:2.76, h:2.05, fontSize:9.5, color:'374151', lineSpacingMultiple:1.3 });
    });
  }

  // ── SLIDE 6: IMPACT & ROI ────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };

    s.addShape('rect', { x:0, y:0, w:10, h:0.08, fill:{ color:C.teal }, line:{ color:C.teal, w:0 } });

    s.addText('IMPACT & ROI', {
      x:0.4, y:0.3, w:9, h:0.3,
      fontSize:9, bold:true, color:C.teal, charSpacing:2,
    });

    s.addText('Built on PhonePe Proof Points — Adjusted for eyewa Scale', {
      x:0.4, y:0.65, w:9, h:0.6,
      fontSize:22, bold:true, color:C.white,
    });

    // 2 sections
    const sections = [
      {
        label:'Consumer Impact',
        color: C.teal,
        metrics:[
          { v:'+32%', l:'Contact lens repeat purchase rate', sub:'Driven by LensLoop supply countdown + 1-tap reorder' },
          { v:'+28%', l:'Pro renewal rate at annual anniversary', sub:'8.5x savings proof surfaced 30 days before renewal date' },
          { v:'−35%', l:'Pro cancellation rate with cancel-save', sub:'"Pause 30 days" primary CTA converts 35% of intenders' },
          { v:'AED 262', l:'Avg Pro savings per reorder session', sub:'BOGO + shipping + warranty accumulated per transaction' },
        ],
      },
      {
        label:'eyewa Business ROI',
        color: C.gold,
        metrics:[
          { v:'4.8x', l:'LTV uplift — contact-lens + Pro member', sub:'vs contact-lens-only customers over 12-month horizon' },
          { v:'AED 0', l:'Additional cost per saved cancellation', sub:'Cancel-save pause is a retention signal, not a discount' },
          { v:'+18%', l:'Pro-driven GMV lift from accumulated BOGO', sub:'Members who see savings proof redeem 40% more vouchers' },
          { v:'<4 wk', l:'Engineering time to MVP', sub:'Order history API + push trigger + pre-fill flow already in stack' },
        ],
      },
    ];

    sections.forEach((sec, si) => {
      const xBase = 0.35 + si * 4.9;
      s.addShape('rect', { x:xBase, y:1.5, w:4.55, h:0.38, fill:{ color:sec.color }, line:{ color:sec.color, w:0 }, rounding:0.06 });
      s.addText(sec.label, { x:xBase+0.15, y:1.5, w:4.25, h:0.38, fontSize:11, bold:true, color:si===0?C.dark:C.dark, valign:'middle' });
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

    s.addShape('rect', { x:0, y:0, w:10, h:0.08, fill:{ color:C.teal }, line:{ color:C.teal, w:0 } });

    s.addText('PROOF OF WORK', {
      x:0.4, y:0.3, w:9, h:0.3,
      fontSize:9, bold:true, color:C.dteal, charSpacing:2,
    });

    s.addText('I Built This at PhonePe. Here\'s the Direct Map.', {
      x:0.4, y:0.65, w:9, h:0.6,
      fontSize:22, bold:true, color:C.navy,
    });

    // Left dark panel
    s.addShape('rect', { x:0.35, y:1.5, w:4.55, h:5.55, fill:{ color:C.dark }, line:{ color:C.dark, w:0 }, rounding:0.08 });
    s.addText('What I shipped at PhonePe', { x:0.55, y:1.65, w:4.2, h:0.35, fontSize:11, bold:true, color:C.teal });

    const proofL = [
      'Built Propensity-to-Transact ML models replacing ₹1000+ Cr/yr static targeting → 32% marketing efficiency gain',
      'Built dynamic cart incentivization engine (cart × intent × time) → 35% AOV uplift, 60% cart abandonment reduction, 20% D30 retention lift',
      'A/B tested milestone waiver vs flat cashback — milestone drove 60% improvement in 90-day activation rate',
      'Rebuilt ₹100 Cr/yr rewards portfolio into ML marketplace → 500+ partners, 11% YoY revenue, 26% engagement uplift',
    ];
    proofL.forEach((l, i) => {
      s.addShape('rect', { x:0.5, y:2.12+i*1.2, w:4.25, h:1.1, fill:{ color:C.hero }, line:{ color:C.teal+'33', w:0.5 }, rounding:0.06 });
      s.addText(l, { x:0.65, y:2.18+i*1.2, w:3.95, h:0.98, fontSize:9.5, color:C.white, lineSpacingMultiple:1.3 });
    });

    // Right light panel
    s.addShape('rect', { x:5.1, y:1.5, w:4.55, h:5.55, fill:{ color:'E0F7F4' }, line:{ color:C.teal+'66', w:1 }, rounding:0.08 });
    s.addText('ProLoop equivalent', { x:5.3, y:1.65, w:4.2, h:0.35, fontSize:11, bold:true, color:C.dteal });

    const proofR = [
      '→ RepeatIQ cohort scoring + LensLoop supply-aware ML trigger for contact lens reorder personalization',
      '→ ProLoop 1-tap reorder (pre-filled Rx + address + card) with cart-level Pro savings real-time accumulator',
      '→ Pro cancel-save math modal + milestone savings display — same waiver psychology, applied to AED 99 renewal',
      '→ Pro membership marketplace: BOGO vouchers + free shipping + warranty — self-serve activation surfaced in dashboard',
    ];
    proofR.forEach((l, i) => {
      s.addShape('rect', { x:5.25, y:2.12+i*1.2, w:4.25, h:1.1, fill:{ color:C.white }, line:{ color:C.teal+'55', w:0.5 }, rounding:0.06 });
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

    s.addText('Three Phases — Zero Risk Pilot to Full Subscription Rollout', {
      x:0.4, y:0.65, w:9, h:0.6,
      fontSize:22, bold:true, color:C.white,
    });

    const phases = [
      {
        phase:'Phase 1',
        period:'Month 1–2: LensLoop Pilot (1K contact lens customers)',
        color: C.teal,
        points:[
          'Identify top 1K contact lens repeat-buyers from order history',
          'Build supply countdown logic: last order date + pack size = days remaining',
          'Deploy home card: countdown badge + 1-tap reorder (Rx + address pre-fill)',
          'Instrument: reorder rate, TAT (days from alert to reorder), Pro membership delta',
        ],
      },
      {
        phase:'Phase 2',
        period:'Month 3–4: Pro Retention Layer',
        color: C.gold,
        points:[
          'Launch Pro dashboard: real-time AED savings accumulator per member',
          'A/B test cancel-save: Math modal (AED saved vs cost) vs standard "are you sure?" flow',
          'Deploy pre-renewal push: "Your Pro saved you AED 847 this year — renew for AED 99"',
          'Instrument: cancel-save conversion rate, renewal rate uplift, BOGO redemption delta',
        ],
      },
      {
        phase:'Phase 3',
        period:'Month 5–6: Full Flywheel + Automation',
        color: C.white,
        points:[
          'Expand LensLoop to all contact lens SKUs and Pro member cohort',
          'Automate next-reminder on every reorder confirmation (perpetual flywheel)',
          'Launch Pro Milestone badges (3x / 5x / 8x ROI) with social sharing',
          'Scale Pro cancel-save to all membership tier exits including eyewaPROlite',
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

    // Ask box
    s.addShape('rect', { x:0.35, y:7.1, w:9.3, h:0.78, fill:{ color:C.teal+'22' }, line:{ color:C.teal, w:1 }, rounding:0.06 });
    s.addText('What I need:  Order history + Rx data API  ·  1 engineer (supply trigger + Pro savings accumulator)  ·  CRM alignment for push cadence', {
      x:0.55, y:7.2, w:9.0, h:0.56, fontSize:10, color:C.white, lineSpacingMultiple:1.3,
    });
    s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439', {
      x:0.35, y:7.45, w:9.3, h:0.3, fontSize:8.5, color:C.muted, align:'right',
    });
  }

  await prs.writeFile({ fileName: 'assets/eyewa-proloop-deck.pptx' });
  console.log('Done: assets/eyewa-proloop-deck.pptx');
})();
