const PptxGenJS = require('pptxgenjs');

(async () => {
  const prs = new PptxGenJS();
  prs.layout = 'LAYOUT_16x9';

  const C = {
    dark:   '0B1120',
    hero:   '0F1F3D',
    accent: '2563EB',
    gold:   'F59E0B',
    white:  'FFFFFF',
    lgray:  'F8FAFC',
    muted:  '94A3B8',
    green:  '059669',
    red:    'DC2626',
  };

  const makeShadow = () => ({ type: 'outer', color: '000000', blur: 4, offset: 2, angle: 45, opacity: 0.14 });

  // ── SLIDE 1: COVER ──────────────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 7.5, fill: { color: C.dark } });
    s.addShape('rect', { x: 0, y: 0, w: 3.5, h: 7.5, fill: { color: C.hero } });
    s.addShape('rect', { x: 3.5, y: 0, w: 0.04, h: 7.5, fill: { color: C.accent } });
    // decorative circles
    s.addShape('ellipse', { x: 6.5, y: -1, w: 3.5, h: 3.5, fill: { color: C.accent }, line: { color: C.accent }, shadow: makeShadow() });
    s.addShape('ellipse', { x: 7.5, y: 4.5, w: 2, h: 2, fill: { color: '1E3A5F' }, line: { color: '1E3A5F' } });

    s.addText('NOMUPAY  ·  MERCHANT EXPERIENCE  ·  Case Study 01', { x: 3.7, y: 0.4, w: 6, h: 0.3, fontSize: 8, bold: true, color: C.muted, charSpacing: 2 });
    s.addText('MerchantIQ', { x: 3.7, y: 0.9, w: 6, h: 1.4, fontSize: 52, bold: true, color: C.white, charSpacing: -1 });
    s.addText('Unified Multi-Acquirer Intelligence\n+ AI-Powered Self-Service', { x: 3.7, y: 2.4, w: 5.8, h: 1, fontSize: 17, color: C.gold, bold: false });
    s.addText('Ajay Avaghade  ·  Product Manager', { x: 3.7, y: 3.5, w: 4, h: 0.3, fontSize: 11, color: C.muted });

    // Left panel content
    s.addShape('rect', { x: 0.3, y: 0.5, w: 0.06, h: 0.5, fill: { color: C.accent } });
    s.addText('96.2%', { x: 0.5, y: 0.45, w: 2.7, h: 0.6, fontSize: 28, bold: true, color: C.white });
    s.addText('Auth rate visibility\nacross all acquirers', { x: 0.5, y: 0.95, w: 2.7, h: 0.5, fontSize: 10, color: C.muted });

    s.addShape('rect', { x: 0.3, y: 2.2, w: 0.06, h: 0.5, fill: { color: C.gold } });
    s.addText('74%', { x: 0.5, y: 2.15, w: 2.7, h: 0.6, fontSize: 28, bold: true, color: C.gold });
    s.addText('Support ticket\ndeflection rate', { x: 0.5, y: 2.65, w: 2.7, h: 0.5, fontSize: 10, color: C.muted });

    s.addShape('rect', { x: 0.3, y: 3.9, w: 0.06, h: 0.5, fill: { color: C.green } });
    s.addText('8 sec', { x: 0.5, y: 3.85, w: 2.7, h: 0.6, fontSize: 28, bold: true, color: C.white });
    s.addText('SupportIQ resolution\nvs 26h via ticket', { x: 0.5, y: 4.35, w: 2.7, h: 0.5, fontSize: 10, color: C.muted });

    s.addText('Live Prototype: ajay-avaghade.github.io/Portfolio/nomupay-merchantiq-prototype.html', { x: 3.7, y: 7.1, w: 6, h: 0.3, fontSize: 8, color: C.muted, hyperlink: { url: 'https://ajay-avaghade.github.io/Portfolio/nomupay-merchantiq-prototype.html' } });
  }

  // ── SLIDE 2: THE PROBLEM ────────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 7.5, fill: { color: C.dark } });
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.red } });

    s.addText('THE PROBLEM', { x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 9, bold: true, color: C.red, charSpacing: 2 });
    s.addText('Merchants Using Multiple Acquirers Are Flying Blind', { x: 0.5, y: 0.7, w: 9, h: 0.8, fontSize: 28, bold: true, color: C.white });
    s.addText('Logging into 3 portals to answer questions that should take 8 seconds is not a workflow — it is an oversight failure.', { x: 0.5, y: 1.55, w: 9, h: 0.4, fontSize: 13, color: C.muted });

    // 3 stat boxes
    const stats = [
      { n: '3+', l: 'Portals per merchant', s: 'Stripe, Adyen, PayU each\nrequire separate logins\nNo consolidated view exists', col: C.accent },
      { n: '26h', l: 'Avg ticket resolution time', s: '74% of queries are\ndata-retrieval questions\nAnswerable in 8 sec with AI', col: C.gold },
      { n: '2–3%', l: 'GMV lost to payment failures', s: 'Undetected routing issues\ncost merchants 2-3% of GMV\nannually (PYMNTS 2023)', col: C.red },
    ];
    stats.forEach((st, i) => {
      const x = 0.4 + i * 3.1;
      s.addShape('rect', { x, y: 2.3, w: 2.9, h: 3.8, fill: { color: '111827' }, line: { color: st.col, pt: 1 }, shadow: makeShadow() });
      s.addShape('rect', { x, y: 2.3, w: 2.9, h: 0.07, fill: { color: st.col } });
      s.addText(st.n, { x: x + 0.2, y: 2.55, w: 2.5, h: 1, fontSize: 40, bold: true, color: st.col });
      s.addText(st.l, { x: x + 0.2, y: 3.45, w: 2.5, h: 0.5, fontSize: 11, bold: true, color: C.white });
      s.addText(st.s, { x: x + 0.2, y: 3.95, w: 2.5, h: 1.4, fontSize: 10, color: C.muted, paraSpaceAfter: 4 });
    });

    // Bottom insight
    s.addShape('rect', { x: 0.4, y: 6.4, w: 9.2, h: 0.85, fill: { color: '111827' }, line: { color: C.accent, pt: 1 }, shadow: makeShadow() });
    s.addText('Root cause: Payment infrastructure is consolidating around multi-acquirer strategies, but merchant tooling has not kept up. The result is operational blindness — merchants cannot see their own payment health in real time, and support teams handle queries that live data should answer automatically.', { x: 0.65, y: 6.5, w: 8.7, h: 0.65, fontSize: 9.5, color: C.muted, italic: true });
  }

  // ── SLIDE 3: THE INSIGHT ────────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 7.5, fill: { color: C.lgray } });
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });

    s.addText('THE INSIGHT', { x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 9, bold: true, color: C.accent, charSpacing: 2 });
    s.addText('Merchants Don\'t Need Another Portal. They Need One Portal That Already Has the Answer.', { x: 0.5, y: 0.7, w: 9, h: 0.9, fontSize: 26, bold: true, color: C.dark });

    // Two columns
    s.addShape('rect', { x: 0.4, y: 1.85, w: 4.3, h: 4.5, fill: { color: 'FEE2E2' }, line: { color: 'FECACA', pt: 1 }, shadow: makeShadow() });
    s.addText('TODAY', { x: 0.6, y: 1.95, w: 3.9, h: 0.3, fontSize: 9, bold: true, color: C.red, charSpacing: 2 });
    const probs = [
      'Stripe portal: auth rate (no Adyen comparison)',
      'Adyen portal: settlements (no Stripe context)',
      'PayU portal: FX (no cross-acquirer view)',
      'Support ticket: "Where is my payout?" — 26h wait',
      'Month-end: manual CSV reconciliation across 3 sources',
    ];
    probs.forEach((p, i) => {
      s.addText('✗  ' + p, { x: 0.6, y: 2.4 + i * 0.7, w: 3.9, h: 0.55, fontSize: 11, color: '7F1D1D' });
    });

    s.addShape('rect', { x: 5.3, y: 1.85, w: 4.3, h: 4.5, fill: { color: 'D1FAE5' }, line: { color: 'A7F3D0', pt: 1 }, shadow: makeShadow() });
    s.addText('MerchantIQ', { x: 5.5, y: 1.95, w: 3.9, h: 0.3, fontSize: 9, bold: true, color: C.green, charSpacing: 2 });
    const sols = [
      'One dashboard: Stripe 97.1% vs Adyen 94.3% vs PayU 91.8%',
      'FX exposure + settlement ETA in a single view',
      'Acquirer alert: "Adyen down 1.2% — review routing"',
      'SupportIQ: "Your payout: £23,450, ETA 17:30 today" — 8s',
      'Auto-reconciliation export on batch completion',
    ];
    sols.forEach((p, i) => {
      s.addText('✓  ' + p, { x: 5.5, y: 2.4 + i * 0.7, w: 3.9, h: 0.55, fontSize: 11, color: '064E3B' });
    });

    s.addShape('rect', { x: 0.4, y: 6.55, w: 9.2, h: 0.75, fill: { color: C.accent }, shadow: makeShadow() });
    s.addText('Key insight: 74% of merchant support tickets are data-retrieval queries — "where is my payout?", "why did this fail?". SupportIQ answers in 8 seconds by pulling live data. No ticket opened. No agent cost. No 26-hour wait. Source: Zendesk CX Trends 2024.', { x: 0.6, y: 6.63, w: 9, h: 0.6, fontSize: 9.5, color: C.white, italic: true });
  }

  // ── SLIDE 4: THE MECHANIC ───────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 7.5, fill: { color: C.hero } });
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.gold } });

    s.addText('THE MECHANIC', { x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 9, bold: true, color: C.gold, charSpacing: 2 });
    s.addText('Two Layers, One Merchant Experience', { x: 0.5, y: 0.7, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white });

    const steps = [
      { n: '01', t: 'Data Consolidation', b: 'MerchantIQ pulls live transaction data from all acquirers (Stripe, Adyen, PayU) into a single data layer. No manual imports. No CSV uploads. Real-time via Nomupay API.' },
      { n: '02', t: 'Unified Dashboard', b: 'Auth rate, volume, FX exposure, payout status — one screen. Multi-acquirer performance comparison with colour-coded health indicators. Alert rail for anomalies (auth drops, connectivity issues).' },
      { n: '03', t: 'Multi-Acquirer Intelligence', b: 'Side-by-side acquirer performance: auth rate, volume, failure reasons per acquirer. AI routing suggestion: "Adyen down 1.2% — shift EU traffic to Stripe for next 4 hours." Merchant acts; Nomupay routes.' },
      { n: '04', t: 'SupportIQ AI Layer', b: 'Merchant asks "where is my payout?" — SupportIQ queries live batch data, returns structured answer: amount, status, ETA, breakdown. Resolved in 8 seconds. No ticket created. No agent involved.' },
      { n: '05', t: 'Zero-Ticket Resolution', b: 'Merchant confirms resolution. Satisfaction rating captured. Category logged in deflection dashboard. Category managers see which query types dominate — and proactively surface them as portal tooltips.' },
    ];

    steps.forEach((st, i) => {
      const x = 0.35 + i * 1.87;
      s.addShape('rect', { x, y: 1.55, w: 1.7, h: 5.6, fill: { color: '111827' }, line: { color: C.accent, pt: 1 }, shadow: makeShadow() });
      s.addShape('rect', { x, y: 1.55, w: 1.7, h: 0.07, fill: { color: C.accent } });
      s.addText(st.n, { x: x + 0.1, y: 1.65, w: 1.5, h: 0.55, fontSize: 26, bold: true, color: C.accent });
      s.addText(st.t, { x: x + 0.1, y: 2.2, w: 1.5, h: 0.7, fontSize: 10.5, bold: true, color: C.white });
      s.addText(st.b, { x: x + 0.1, y: 2.95, w: 1.5, h: 3.5, fontSize: 9, color: C.muted, paraSpaceAfter: 3 });
    });

    s.addText('PhonePe proof: Self-serve PG platform consolidated 3 bank integrations into one merchant view — same pattern applied to multi-acquirer intelligence. Also: ops intelligence layer reduced TAT from 2 days to 30 minutes by surfacing back-office data as merchant-facing insights.', { x: 0.4, y: 7.1, w: 9.2, h: 0.3, fontSize: 8, color: C.muted, italic: true });
  }

  // ── SLIDE 5: THE PRODUCT ────────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 7.5, fill: { color: C.lgray } });
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });

    s.addText('THE PRODUCT', { x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 9, bold: true, color: C.accent, charSpacing: 2 });
    s.addText('5 Screen States — Full Merchant Journey End to End', { x: 0.5, y: 0.7, w: 9, h: 0.6, fontSize: 26, bold: true, color: C.dark });

    const screens = [
      { n: '01', t: 'Payment Dashboard', b: 'Auth rate 96.2% prominently displayed. Acquirer health cards: Stripe 97.1% healthy, Adyen 94.3% alert, PayU 91.8% stable. FX exposure strip. SupportIQ entry point.' },
      { n: '02', t: 'Analytics Deep-Dive', b: '7-day auth rate chart per acquirer. Decline reason breakdown (insufficient funds 34%, expired card 22%, 3DS 18%). FX EUR/GBP exposure -2.1%.' },
      { n: '03', t: 'SupportIQ: Query Sent', b: 'Merchant types: "Where is my Tuesday payout? Batch #BP-20240624". SupportIQ AI typing indicator shows live data query in progress.' },
      { n: '04', t: 'SupportIQ: Resolved', b: 'AI returns: "Batch #BP-20240624 — £23,450 across 847 txns. 842 settled. 5 processing, ETA today 17:30 UK / 20:30 GST. No action needed." Resolved in 8s. No ticket raised.' },
      { n: '05', t: 'Deflection Insights', b: '847 queries resolved by SupportIQ this month. 0 tickets raised. 74% deflection rate. £8.5K support cost saved at £10/ticket. Satisfaction: 4.8/5 from 847 ratings.' },
    ];

    screens.forEach((sc, i) => {
      const x = 0.35 + i * 1.87;
      s.addShape('rect', { x, y: 1.55, w: 1.7, h: 5.6, fill: { color: C.white }, line: { color: 'E5E7EB', pt: 1 }, shadow: makeShadow() });
      s.addShape('rect', { x, y: 1.55, w: 1.7, h: 0.07, fill: { color: C.accent } });
      s.addText(sc.n, { x: x + 0.1, y: 1.65, w: 1.5, h: 0.55, fontSize: 26, bold: true, color: C.accent });
      s.addText(sc.t, { x: x + 0.1, y: 2.2, w: 1.5, h: 0.7, fontSize: 10.5, bold: true, color: C.dark });
      s.addText(sc.b, { x: x + 0.1, y: 2.95, w: 1.5, h: 3.5, fontSize: 9, color: '4B5563', paraSpaceAfter: 3 });
    });

    s.addText('Prototype: ajay-avaghade.github.io/Portfolio/nomupay-merchantiq-prototype.html  ·  All 5 states interactive with keyboard arrow navigation', { x: 0.4, y: 7.15, w: 9.2, h: 0.25, fontSize: 8, color: C.muted });
  }

  // ── SLIDE 6: IMPACT & ROI ───────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 7.5, fill: { color: C.dark } });
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });

    s.addText('IMPACT & ROI', { x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 9, bold: true, color: C.accent, charSpacing: 2 });
    s.addText('Built on Industry Benchmarks — Adjusted for Nomupay Merchant Scale', { x: 0.5, y: 0.7, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.white });

    // Merchant impact
    s.addShape('rect', { x: 0.4, y: 1.55, w: 4.3, h: 5.55, fill: { color: '111827' }, line: { color: C.accent, pt: 1 }, shadow: makeShadow() });
    s.addShape('rect', { x: 0.4, y: 1.55, w: 4.3, h: 0.07, fill: { color: C.accent } });
    s.addText('MERCHANT IMPACT', { x: 0.6, y: 1.7, w: 3.9, h: 0.3, fontSize: 9, bold: true, color: C.accent, charSpacing: 2 });
    const mImpact = [
      { v: '↓ 74%', l: 'Support tickets deflected by SupportIQ', src: 'Zendesk CX Trends 2024' },
      { v: '↑ 35%', l: 'Faster operational decisions (consolidated dashboard)', src: 'McKinsey Digital Banking 2022' },
      { v: '8 sec', l: 'Query resolution vs 26h via support ticket', src: 'Intercom Benchmark 2023' },
      { v: '↓ 15-20%', l: 'Revenue lost to undetected payment failures', src: 'PYMNTS Failure Index 2023' },
    ];
    mImpact.forEach((m, i) => {
      s.addText(m.v, { x: 0.6, y: 2.2 + i * 1.15, w: 3.9, h: 0.55, fontSize: 22, bold: true, color: C.accent });
      s.addText(m.l, { x: 0.6, y: 2.7 + i * 1.15, w: 3.9, h: 0.35, fontSize: 10, color: C.white });
      s.addText('Source: ' + m.src, { x: 0.6, y: 2.98 + i * 1.15, w: 3.9, h: 0.25, fontSize: 8, color: C.muted, italic: true });
    });

    // Nomupay ROI
    s.addShape('rect', { x: 5.3, y: 1.55, w: 4.3, h: 5.55, fill: { color: '111827' }, line: { color: C.gold, pt: 1 }, shadow: makeShadow() });
    s.addShape('rect', { x: 5.3, y: 1.55, w: 4.3, h: 0.07, fill: { color: C.gold } });
    s.addText('NOMUPAY ROI', { x: 5.5, y: 1.7, w: 3.9, h: 0.3, fontSize: 9, bold: true, color: C.gold, charSpacing: 2 });
    const nRoi = [
      { v: '£0', l: 'Incremental cost per SupportIQ query resolved', src: 'After amortised model cost' },
      { v: '£10-15', l: 'Cost per human-agent ticket (avoided)', src: 'Avg fintech B2B support cost' },
      { v: '↑ 2.4x', l: 'Merchant portal session time with MerchantIQ', src: 'Vs fragmented portal baseline' },
      { v: '↑ 40%', l: 'Merchant NPS uplift (single source of truth)', src: 'Industry retention research' },
    ];
    nRoi.forEach((m, i) => {
      s.addText(m.v, { x: 5.5, y: 2.2 + i * 1.15, w: 3.9, h: 0.55, fontSize: 22, bold: true, color: C.gold });
      s.addText(m.l, { x: 5.5, y: 2.7 + i * 1.15, w: 3.9, h: 0.35, fontSize: 10, color: C.white });
      s.addText('Source: ' + m.src, { x: 5.5, y: 2.98 + i * 1.15, w: 3.9, h: 0.25, fontSize: 8, color: C.muted, italic: true });
    });
  }

  // ── SLIDE 7: PROOF OF WORK ──────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 5, h: 7.5, fill: { color: C.dark } });
    s.addShape('rect', { x: 5, y: 0, w: 5, h: 7.5, fill: { color: C.lgray } });
    s.addShape('rect', { x: 4.97, y: 0, w: 0.06, h: 7.5, fill: { color: C.accent } });

    s.addText('PROOF OF WORK', { x: 0.4, y: 0.3, w: 4.2, h: 0.3, fontSize: 9, bold: true, color: C.accent, charSpacing: 2 });
    s.addText('I Built These Patterns at PhonePe.\nHere Is the Direct Map to MerchantIQ.', { x: 0.4, y: 0.7, w: 4.2, h: 1, fontSize: 18, bold: true, color: C.white });

    const proofs = [
      { pp: 'Self-serve PG platform: 3 bank integrations unified into one merchant view — 5,000+ B2B merchants, 23% lower CAC', miq: 'MerchantIQ: 3+ acquirer integrations unified into one dashboard for Nomupay merchants' },
      { pp: 'Ops intelligence layer: 6 siloed workflows into one console — TAT 2 days to 30 minutes', miq: 'SupportIQ: 26-hour support ticket eliminated by surfacing back-office data as a live AI response in 8 seconds' },
      { pp: 'Propensity-to-Transact ML models: real-time user-level scoring replacing manual cohorts — 32% efficiency gain', miq: 'SupportIQ AI: real-time data retrieval and interpretation replacing human agent lookup' },
      { pp: 'Milestone A/B test: defined baseline, instrumented funnel, read cohort curves — 60% activation uplift', miq: 'MerchantIQ deflection tracking: defined baseline ticket volume, instrumented SupportIQ funnel, reported on deflection rate' },
    ];

    proofs.forEach((p, i) => {
      const y = 1.95 + i * 1.3;
      s.addShape('rect', { x: 0.3, y, w: 4.3, h: 1.1, fill: { color: '111827' }, line: { color: C.accent, pt: 1 }, shadow: makeShadow() });
      s.addText(p.pp, { x: 0.45, y: y + 0.08, w: 4, h: 0.9, fontSize: 9, color: C.muted });

      s.addShape('rect', { x: 5.2, y, w: 4.4, h: 1.1, fill: { color: C.white }, line: { color: C.accent, pt: 1 }, shadow: makeShadow() });
      s.addShape('rect', { x: 5.2, y, w: 4.4, h: 0.06, fill: { color: C.accent } });
      s.addText('→  ' + p.miq, { x: 5.35, y: y + 0.12, w: 4.1, h: 0.9, fontSize: 9.5, color: '111827' });
    });

    s.addShape('rect', { x: 0.3, y: 7.1, w: 9.4, h: 0.3, fill: { color: C.accent }, shadow: makeShadow() });
    s.addText('This is not a case study conceived in theory. These are product patterns I have shipped — applied to Nomupay\'s merchant experience problem.', { x: 0.5, y: 7.12, w: 9, h: 0.26, fontSize: 8.5, color: C.white, italic: true });
  }

  // ── SLIDE 8: ROLLOUT PLAN ───────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 7.5, fill: { color: C.dark } });
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.gold } });

    s.addText('ROLLOUT PLAN', { x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 9, bold: true, color: C.gold, charSpacing: 2 });
    s.addText('Three Phases — Pilot Cohort to Full Merchant Base', { x: 0.5, y: 0.7, w: 9, h: 0.6, fontSize: 26, bold: true, color: C.white });

    const phases = [
      {
        ph: 'Phase 1', dur: 'Month 1–2: Pilot (50 merchants)', col: C.accent,
        pts: [
          'Deploy MerchantIQ dashboard to 50 Nomupay enterprise merchants with 2+ acquirers',
          'Enable SupportIQ for top 20 query types (payout status, failed transaction, FX rate)',
          'Instrument: ticket deflection rate, resolution time, merchant session time',
          'Measure 30-day merchant NPS vs non-MerchantIQ control group',
        ],
      },
      {
        ph: 'Phase 2', dur: 'Month 3–4: Scale + Optimise', col: C.gold,
        pts: [
          'Expand to 500 merchants — A/B test SupportIQ prompt styles for deflection rate optimisation',
          'Add FX exposure alerts and routing recommendations based on real-time acquirer data',
          'Enable batch payout status as SupportIQ query category (integrates with BatchPayIQ)',
          'Tune deflection thresholds — escalation to human agent when AI confidence below 85%',
        ],
      },
      {
        ph: 'Phase 3', dur: 'Month 5–6: Full Rollout', col: C.green,
        pts: [
          'Roll out MerchantIQ + SupportIQ to all Nomupay merchants globally',
          'Launch SupportIQ for API query resolution (developer experience extension)',
          'Introduce proactive anomaly alerts — push notification when acquirer drops below threshold',
          'Target: 90%+ deflection rate for tier-1 queries; merchant support headcount redeployment',
        ],
      },
    ];

    phases.forEach((ph, i) => {
      const x = 0.35 + i * 3.1;
      s.addShape('rect', { x, y: 1.55, w: 2.9, h: 5.6, fill: { color: '111827' }, line: { color: ph.col, pt: 1 }, shadow: makeShadow() });
      s.addShape('rect', { x, y: 1.55, w: 2.9, h: 0.07, fill: { color: ph.col } });
      s.addText(ph.ph, { x: x + 0.18, y: 1.68, w: 2.5, h: 0.5, fontSize: 20, bold: true, color: ph.col });
      s.addText(ph.dur, { x: x + 0.18, y: 2.18, w: 2.55, h: 0.45, fontSize: 10, bold: true, color: C.white });
      ph.pts.forEach((pt, j) => {
        s.addText('·  ' + pt, { x: x + 0.18, y: 2.75 + j * 0.9, w: 2.55, h: 0.8, fontSize: 9, color: C.muted, paraSpaceAfter: 2 });
      });
    });

    s.addShape('rect', { x: 0.35, y: 7.1, w: 9.3, h: 0.3, fill: { color: '1E293B' }, shadow: makeShadow() });
    s.addText('What I need to start: Read access to merchant support ticket taxonomy  ·  1 backend engineer (Nomupay data API integration)  ·  Alignment with Merchant Success team on query category priority', { x: 0.55, y: 7.14, w: 9, h: 0.22, fontSize: 8, color: C.muted });

    s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439', { x: 0.4, y: 7.33, w: 9.2, h: 0.17, fontSize: 8, color: C.muted });
  }

  await prs.writeFile({ fileName: 'assets/nomupay-merchantiq-deck.pptx' });
  console.log('Done: assets/nomupay-merchantiq-deck.pptx');
})();
