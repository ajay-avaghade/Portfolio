const pptx = require('pptxgenjs');
const prs = new pptx();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0A0A0A',
  hero:   '141414',
  bright: 'C2A582',
  orange: 'D4B895',
  amber:  'F59E0B',
  white:  'FFFFFF',
  lgray:  'FDFBF7',
  muted:  '8A8A8A',
  ink:    '111111',
  slate:  '404040',
  red:    'EF4444',
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 3, offset: 1, angle: 45, opacity: 0.12 });

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 1 — COVER (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  // Diagonal stripe accents
  for (let i = 0; i < 6; i++) {
    slide.addShape(prs.ShapeType.rect, {
      x: -0.5 + i * 2.2, y: -0.2, w: 0.18, h: 10,
      fill: { color: C.bright, transparency: 88 },
      line: { type: 'none' },
      rotate: 15,
    });
  }

  // Dezerv gold accent bar top
  slide.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 0.06,
    fill: { color: C.bright }, line: { type: 'none' },
  });

  // Company tag
  slide.addText('DEZERV — INVESTMENT EXECUTION PM', {
    x: 0.6, y: 0.28, w: 6, h: 0.28,
    fontSize: 9, bold: true, color: C.bright,
    charSpacing: 3, fontFace: 'Calibri',
  });

  // Product name
  slide.addText('SmartSIP Engine', {
    x: 0.6, y: 0.72, w: 8, h: 1.1,
    fontSize: 56, bold: true, color: C.white, fontFace: 'Calibri',
  });

  // Subtitle
  slide.addText('Predictive ML retry routing & automated fallback resolution\nthat drives zero-drop investment execution for affluent clients', {
    x: 0.6, y: 1.82, w: 7.4, h: 0.72,
    fontSize: 15, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
  });

  // Presenter
  slide.addText('Presented by Ajay Avaghade  ·  Product Manager — Investment Execution', {
    x: 0.6, y: 2.68, w: 6, h: 0.3,
    fontSize: 11, color: C.muted, fontFace: 'Calibri',
  });

  // Big stat box bottom right
  slide.addShape(prs.ShapeType.rect, {
    x: 7.2, y: 1.6, w: 2.4, h: 1.9,
    fill: { color: C.bright, transparency: 15 },
    line: { color: C.bright, pt: 1.5 },
    rectRadius: 0.1,
  });
  slide.addText('₹42.5Cr', {
    x: 7.2, y: 1.72, w: 2.4, h: 0.72,
    fontSize: 36, bold: true, color: C.bright, align: 'center', fontFace: 'Calibri',
  });
  slide.addText('abandoned AUM\nrecovered instantly\nvia SmartSIP', {
    x: 7.2, y: 2.44, w: 2.4, h: 0.72,
    fontSize: 9.5, color: C.muted, align: 'center', fontFace: 'Calibri', lineSpacingMultiple: 1.3,
  });

  // Bottom bar
  slide.addShape(prs.ShapeType.rect, {
    x: 0, y: 7.08, w: 10, h: 0.42,
    fill: { color: C.hero }, line: { type: 'none' },
  });
  slide.addText('Dezerv — Proprietary & Confidential  ·  Prepared by Ajay Avaghade  ·  2026', {
    x: 0.3, y: 7.1, w: 9.4, h: 0.3,
    fontSize: 8, color: C.muted, fontFace: 'Calibri',
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  slide.addText('THE PROBLEM', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('Failed Executions Erode Trust & Delay AUM Growth', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.9,
    fontSize: 26, bold: true, color: C.white, fontFace: 'Calibri', lineSpacingMultiple: 1.2,
  });

  // 3 stat columns
  const stats = [
    { icon: '📉', num: '68%', lbl: 'Failures from Insufficient Funds', sub: 'NACH mandate bounces due to timing mismatches in affluent clients\' primary accounts.' },
    { icon: '⏳', num: 'T+2', lbl: 'Current Ops Resolution TAT', sub: 'Ops team manually reconciles failures and reaches out to clients 48 hours later.' },
    { icon: '💸', num: '₹140Cr', lbl: 'Delayed Monthly AUM', sub: 'Capital sitting uninvested due to friction in retrying failed SIPs / mandates.' },
  ];
  stats.forEach((s, i) => {
    const x = 0.5 + i * 3.1;
    slide.addShape(prs.ShapeType.rect, {
      x, y: 1.72, w: 2.85, h: 3.4,
      fill: { color: C.hero }, line: { color: C.bright, pt: 1, transparency: 60 }, rectRadius: 0.1,
      shadow: makeShadow(),
    });
    slide.addText(s.icon, { x, y: 1.88, w: 2.85, h: 0.44, fontSize: 22, align: 'center' });
    slide.addText(s.num, {
      x, y: 2.36, w: 2.85, h: 0.72,
      fontSize: 38, bold: true, color: C.bright, align: 'center', fontFace: 'Calibri',
    });
    slide.addText(s.lbl, {
      x: x + 0.14, y: 3.12, w: 2.6, h: 0.52,
      fontSize: 11, bold: true, color: C.white, align: 'center', fontFace: 'Calibri', lineSpacingMultiple: 1.2,
    });
    slide.addText(s.sub, {
      x: x + 0.1, y: 3.68, w: 2.68, h: 1.1,
      fontSize: 9, color: C.muted, align: 'center', fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });

  // Insight box
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 5.38, w: 9, h: 1.18,
    fill: { color: C.hero }, line: { color: C.orange, pt: 1.5, transparency: 40 }, rectRadius: 0.1,
  });
  slide.addText('Root cause: Wealth management execution relies on archaic banking infrastructure (NACH, slow PGs). When a debit fails, the burden of resolution falls entirely on manual Ops reconciliation and high-friction client follow-ups. Affluent clients expect a "done-for-you" proactive experience, not a failed payment email.', {
    x: 0.72, y: 5.5, w: 8.6, h: 0.94,
    fontSize: 10.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.4,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 3 — THE INSIGHT (light)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.lgray };

  slide.addText('THE INSIGHT', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.slate, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('Manual Reconciliation vs Predictive Execution Routing', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.82,
    fontSize: 24, bold: true, color: C.ink, fontFace: 'Calibri', lineSpacingMultiple: 1.2,
  });

  // VS divider oval
  slide.addShape(prs.ShapeType.ellipse, {
    x: 4.6, y: 2.3, w: 0.8, h: 0.8,
    fill: { color: C.dark }, line: { type: 'none' },
  });
  slide.addText('VS', {
    x: 4.6, y: 2.36, w: 0.8, h: 0.62,
    fontSize: 11, bold: true, color: C.white, align: 'center', fontFace: 'Calibri',
  });

  // LEFT column — current
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.6, w: 4.0, h: 4.2,
    fill: { color: 'FFF5F5' }, line: { color: 'EF4444', pt: 1.5, transparency: 40 }, rectRadius: 0.1,
  });
  slide.addText('❌  Reactive Ops (Today)', {
    x: 0.68, y: 1.76, w: 3.68, h: 0.4,
    fontSize: 13, bold: true, color: C.ink, fontFace: 'Calibri',
  });
  const leftPts = [
    'Rely on single Payment Gateway; if it times out, the SIP is entirely dropped.',
    'Wait T+2 days for NACH bounce reports to filter into internal ledgers.',
    'Ops team manually calls / emails high-net-worth clients to retry payment.',
    'Client has to manually log in, find the failed mandate, and re-initiate.',
    'Lost market days = lost yield = eroded trust.',
  ];
  leftPts.forEach((pt, i) => {
    slide.addText('• ' + pt, {
      x: 0.68, y: 2.22 + i * 0.62, w: 3.6, h: 0.54,
      fontSize: 10.5, color: C.slate, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
    });
  });

  // RIGHT column — proposed
  slide.addShape(prs.ShapeType.rect, {
    x: 5.5, y: 1.6, w: 4.0, h: 4.2,
    fill: { color: 'FAF8F5' }, line: { color: C.bright, pt: 1.5, transparency: 30 }, rectRadius: 0.1,
  });
  slide.addText('✅  SmartSIP Engine (Proposed)', {
    x: 5.68, y: 1.76, w: 3.68, h: 0.4,
    fontSize: 13, bold: true, color: C.ink, fontFace: 'Calibri',
  });
  const rightPts = [
    'Dynamic routing: Instantly failover to secondary PG on timeout (BillDesk → Razorpay).',
    'Real-time hooks detect NACH failure instantly at T+0.',
    'Automated ML-triggered WhatsApp nudge with 1-click UPI Autopay fallback link.',
    'Client resolves the failure instantly from their phone with no manual Ops involvement.',
    'AUM is secured same-day, preserving the investment streak.',
  ];
  rightPts.forEach((pt, i) => {
    slide.addText('• ' + pt, {
      x: 5.68, y: 2.22 + i * 0.62, w: 3.6, h: 0.54,
      fontSize: 10.5, color: C.slate, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
    });
  });

  // Callout box
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.0, w: 9, h: 0.96,
    fill: { color: C.dark }, line: { type: 'none' }, rectRadius: 0.1,
  });
  slide.addText('Key insight: For affluent Indians, the barrier isn\'t capital—it\'s liquidity timing and platform friction. By shifting the resolution layer from Ops to an automated, instant UPI fallback, we convert a stressful "failed payment" into a frictionless "tap to approve" experience.', {
    x: 0.72, y: 6.1, w: 8.6, h: 0.76,
    fontSize: 10, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.4,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 4 — THE MECHANIC (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  slide.addText('THE MECHANIC', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('How SmartSIP Works: Detect, Reroute, Resolve', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.8,
    fontSize: 23, bold: true, color: C.white, fontFace: 'Calibri', lineSpacingMultiple: 1.2,
  });

  const steps = [
    { n: '01', title: 'Execution Initialization & Routing', desc: 'Dezerv initiates scheduled SIP debits. SmartSIP engine checks historical PG success rates. If BillDesk shows latency >3s, traffic is dynamically routed to Razorpay.' },
    { n: '02', title: 'Real-Time Failure Detection', desc: 'If a debit fails (e.g. Insufficient Funds code 04), the engine intercepts the webhook instantly instead of waiting for EOD batch reconciliation.' },
    { n: '03', title: 'Instant Nudge Orchestration', desc: 'Within 60 seconds of failure, a personalised WhatsApp and Push Notification is sent to the client: "SIP failed, tap here to resolve instantly."' },
    { n: '04', title: '1-Click UPI Fallback', desc: 'The client taps the link, which opens a pre-filled, secure UPI Autopay / intent flow. They authorise with their UPI PIN in seconds.' },
    { n: '05', title: 'T+0 Settlement Ledger Update', desc: 'Dezerv\'s internal ledgers are instantly updated. Ops dashboard clears the alert. The client\'s portfolio reflects the successful execution same-day.' },
  ];

  steps.forEach((s, i) => {
    const y = 1.6 + i * 0.98;
    // Step number circle
    slide.addShape(prs.ShapeType.ellipse, {
      x: 0.5, y: y, w: 0.48, h: 0.48,
      fill: { color: C.bright }, line: { type: 'none' },
    });
    slide.addText(s.n, {
      x: 0.5, y: y + 0.02, w: 0.48, h: 0.38,
      fontSize: 9, bold: true, color: C.dark, align: 'center', fontFace: 'Calibri',
    });
    // Connector line (skip last)
    if (i < 4) {
      slide.addShape(prs.ShapeType.rect, {
        x: 0.72, y: y + 0.48, w: 0.04, h: 0.5,
        fill: { color: C.bright, transparency: 60 }, line: { type: 'none' },
      });
    }
    slide.addText(s.title, {
      x: 1.18, y: y, w: 8.2, h: 0.28,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri',
    });
    slide.addText(s.desc, {
      x: 1.18, y: y + 0.28, w: 8.2, h: 0.62,
      fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });

  // A/B proof note
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.52, w: 9, h: 0.64,
    fill: { color: C.hero }, line: { color: C.orange, pt: 1, transparency: 40 }, rectRadius: 0.08,
  });
  slide.addText('PhonePe Proof: Built a self-serve PG integration with dynamic subvention routing. Deployed across 5,000+ merchants, ensuring zero-drop reliability and lowering operational TAT from 2 days to 30 mins.', {
    x: 0.72, y: 6.6, w: 8.6, h: 0.5,
    fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 5 — THE PRODUCT (light)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.lgray };

  slide.addText('THE PRODUCT', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.slate, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('4 Key Screen States — Ops Dashboard + Consumer Fallback', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 22, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  const screens = [
    {
      num: '01', title: 'Ops Command Center',
      desc: 'Real-time SIP execution monitoring. Live failure categorisation (Insufficient Funds vs PG Timeout). SmartSIP activation.',
      mock: '[ 18k Due | 1.4k Failed ]\n Insufficient Funds: 68%\n PG Downtime: 21%\n[Activate SmartSIP →]',
    },
    {
      num: '02', title: 'Routing Configuration',
      desc: 'No-code UI for defining failover logic. Setup PG-to-PG fallback. Configure instant WhatsApp UPI triggers for affluent clients.',
      mock: '[ Scenario 1: Timeout ]\n Billdesk → Razorpay\n[ Scenario 2: No Funds ]\n Trigger Instant UPI Push',
    },
    {
      num: '03', title: 'Consumer Resolution',
      desc: 'Premium, light-themed consumer UI. Contextual failure explanation with 1-click "Pay via UPI App" resolution button.',
      mock: '[! SIP Debit Failed ]\n ₹50,000 Dezerv PMS\n\n[ Pay via UPI App ]\n (Instant Resolution)',
    },
    {
      num: '04', title: 'Impact Dashboard',
      desc: 'Executive view of SmartSIP ROI. Tracking AUM recovered instantly, PG success rates, and funnel conversion of fallback links.',
      mock: '[ SmartSIP Impact ]\n ₹42.5 Cr AUM Recovered\n 68% Resolution Rate\n 99.8% PG Success\n Fallback TAT: 4 Mins',
    },
  ];

  screens.forEach((s, i) => {
    const x = 0.32 + i * 2.36;
    slide.addShape(prs.ShapeType.rect, {
      x, y: 1.4, w: 2.2, h: 5.3,
      fill: { color: C.white }, line: { color: C.bright, pt: 1.5, transparency: 40 }, rectRadius: 0.1,
      shadow: makeShadow(),
    });
    slide.addShape(prs.ShapeType.rect, {
      x, y: 1.4, w: 2.2, h: 0.36,
      fill: { color: C.dark }, line: { type: 'none' }, rectRadius: 0.1,
    });
    slide.addText(s.num + '  ' + s.title, {
      x: x + 0.1, y: 1.44, w: 2.05, h: 0.28,
      fontSize: 8.5, bold: true, color: C.bright, fontFace: 'Calibri',
    });
    // Mockup block
    slide.addShape(prs.ShapeType.rect, {
      x: x + 0.1, y: 1.86, w: 2.02, h: 2.3,
      fill: { color: 'FAF8F5' }, line: { color: C.bright, pt: 0.75, transparency: 60 }, rectRadius: 0.08,
    });
    slide.addText(s.mock, {
      x: x + 0.14, y: 1.9, w: 1.96, h: 2.2,
      fontSize: 7.5, color: C.slate, fontFace: 'Courier New', lineSpacingMultiple: 1.45,
    });
    slide.addText(s.desc, {
      x: x + 0.1, y: 4.24, w: 2.02, h: 2.2,
      fontSize: 9, color: C.slate, fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });

  slide.addText('Live interactive prototype: dezerv-smartsip-prototype.html', {
    x: 0.5, y: 6.88, w: 9, h: 0.28,
    fontSize: 8.5, color: C.slate, italic: true, fontFace: 'Calibri', align: 'center',
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  slide.addText('IMPACT & ROI', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('Projected Impact — Built on PhonePe Execution Proof Points', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 24, bold: true, color: C.white, fontFace: 'Calibri',
  });

  // Left section label
  slide.addText('OPERATIONAL IMPACT', {
    x: 0.5, y: 1.38, w: 4.4, h: 0.28,
    fontSize: 8.5, bold: true, color: C.bright, charSpacing: 2, fontFace: 'Calibri',
  });
  const leftMetrics = [
    { val: '4 Mins', lbl: 'Average Resolution TAT', sub: 'Down from T+2 days of manual ops reconciliation' },
    { val: '99.8%', lbl: 'PG Success Rate', sub: 'Driven by real-time dynamic gateway routing and failover' },
    { val: '−80%', lbl: 'Ops Manual Interventions', sub: 'Ops team focuses on true edge cases, not generic NACH failures' },
    { val: '68%', lbl: 'Nudge Conversion Rate', sub: 'Affluent clients instantly resolve via 1-click UPI Autopay link' },
  ];
  leftMetrics.forEach((m, i) => {
    const y = 1.72 + i * 1.22;
    slide.addShape(prs.ShapeType.rect, {
      x: 0.5, y, w: 4.3, h: 1.08,
      fill: { color: C.hero }, line: { color: C.bright, pt: 0.75, transparency: 50 }, rectRadius: 0.08,
    });
    slide.addText(m.val, {
      x: 0.68, y: y + 0.1, w: 1.6, h: 0.52,
      fontSize: 28, bold: true, color: C.bright, fontFace: 'Calibri',
    });
    slide.addText(m.lbl, {
      x: 2.32, y: y + 0.1, w: 2.3, h: 0.28,
      fontSize: 11, bold: true, color: C.white, fontFace: 'Calibri',
    });
    slide.addText(m.sub, {
      x: 2.32, y: y + 0.42, w: 2.3, h: 0.54,
      fontSize: 8.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
    });
  });

  // Right section label
  slide.addText('BUSINESS & AUM IMPACT', {
    x: 5.1, y: 1.38, w: 4.4, h: 0.28,
    fontSize: 8.5, bold: true, color: C.orange, charSpacing: 2, fontFace: 'Calibri',
  });
  const rightMetrics = [
    { val: '₹42.5Cr', lbl: 'Monthly AUM Recovered', sub: 'Capital instantly secured that would otherwise sit idle or drop off' },
    { val: 'T+0', lbl: 'Settlement Synchronisation', sub: 'Ledgers update same-day, preserving client investment streaks' },
    { val: '+18%', lbl: 'Client Retention Lift', sub: 'Removing stress from wealth building increases long-term trust' },
    { val: 'Zero', lbl: 'Friction Execution', sub: 'Fulfills Dezerv\'s promise of growing wealth without stress or time cost' },
  ];
  rightMetrics.forEach((m, i) => {
    const y = 1.72 + i * 1.22;
    slide.addShape(prs.ShapeType.rect, {
      x: 5.1, y, w: 4.3, h: 1.08,
      fill: { color: C.hero }, line: { color: C.orange, pt: 0.75, transparency: 50 }, rectRadius: 0.08,
    });
    slide.addText(m.val, {
      x: 5.28, y: y + 0.1, w: 1.6, h: 0.52,
      fontSize: 28, bold: true, color: C.orange, fontFace: 'Calibri',
    });
    slide.addText(m.lbl, {
      x: 6.92, y: y + 0.1, w: 2.3, h: 0.28,
      fontSize: 11, bold: true, color: C.white, fontFace: 'Calibri',
    });
    slide.addText(m.sub, {
      x: 6.92, y: y + 0.42, w: 2.3, h: 0.54,
      fontSize: 8.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
    });
  });

  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.56, w: 9, h: 0.6,
    fill: { color: C.hero }, line: { color: C.bright, pt: 1, transparency: 50 }, rectRadius: 0.08,
  });
  slide.addText('SmartSIP Engine transforms execution reliability from a backend Ops cost center into a core consumer trust feature. By owning outcomes, Dezerv ensures no affluent client is stressed by legacy banking failure modes.', {
    x: 0.72, y: 6.64, w: 8.6, h: 0.44,
    fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK (light)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.lgray };

  slide.addText('PROOF OF WORK', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.slate, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('I Built This Architecture. At PhonePe. At Scale.', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 23, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  // Left — dark column
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.38, w: 4.3, h: 5.3,
    fill: { color: C.dark }, line: { type: 'none' }, rectRadius: 0.1,
  });
  slide.addText('What I Built at PhonePe', {
    x: 0.7, y: 1.56, w: 3.9, h: 0.34,
    fontSize: 12, bold: true, color: C.bright, fontFace: 'Calibri',
  });
  const leftPts = [
    'Led execution architecture for third-party rewards, rebuilding ₹100 Cr/yr infrastructure into an ML-driven marketplace handling massive transaction volumes seamlessly.',
    'Built self-serve PG integration with dynamic subvention routing. Deployed to 5,000+ B2B merchants, acquiring 5Mn+ new users/month with zero-drop reliability.',
    'Redesigned the entire offer operations workflow, reducing manual reconciliation TAT from 2 days down to 30 minutes.',
    'Led cross-functional alignment across Engineering, Operations, and Compliance to ship complex financial products with strict regulatory constraints.',
  ];
  leftPts.forEach((pt, i) => {
    slide.addText('• ' + pt, {
      x: 0.7, y: 2.0 + i * 1.08, w: 3.9, h: 0.96,
      fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });

  // Right — light column
  slide.addShape(prs.ShapeType.rect, {
    x: 5.1, y: 1.38, w: 4.3, h: 5.3,
    fill: { color: C.white }, line: { color: C.bright, pt: 1.5, transparency: 40 }, rectRadius: 0.1,
    shadow: makeShadow(),
  });
  slide.addText('How It Maps to the Dezerv JD', {
    x: 5.3, y: 1.56, w: 3.9, h: 0.34,
    fontSize: 12, bold: true, color: C.ink, fontFace: 'Calibri',
  });
  const rightPts = [
    ['Execution Reliability', 'PG integration + dynamic routing at PhonePe → Dezerv Payments & SIP architecture'],
    ['Ops TAT Reduction', '2 days down to 30 mins at PhonePe → Automated settlements and reconciliations'],
    ['Owning the Outcome', 'Led P&L impact (32% burn reduction, 11% YoY growth) → Dezerv AUM accountability'],
    ['Cross-functional Leadership', 'Shipped products across tech, ops, and sales → Dezerv PM "wear many hats" culture'],
  ];
  rightPts.forEach((pt, i) => {
    const y = 2.0 + i * 1.08;
    slide.addText(pt[0], {
      x: 5.3, y, w: 3.9, h: 0.28,
      fontSize: 10.5, bold: true, color: C.ink, fontFace: 'Calibri',
    });
    slide.addText(pt[1], {
      x: 5.3, y: y + 0.28, w: 3.9, h: 0.68,
      fontSize: 9.5, color: C.slate, fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  slide.addText('ROLLOUT PLAN', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('Phased Deployment — Stability First, Then Scale', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 24, bold: true, color: C.white, fontFace: 'Calibri',
  });

  const phases = [
    {
      phase: 'Phase 1',
      time: 'Month 1–2 · Foundation',
      title: 'Ops Visibility & Gateway Routing',
      pts: [
        'Deploy Ops Command Center for real-time SIP failure tracking',
        'Implement dual-PG routing (e.g., BillDesk fallback to Razorpay)',
        'Audit historical NACH failure data to train intent models',
        'Reduce PG timeout drop-offs by 90% immediately',
      ],
      color: C.bright,
    },
    {
      phase: 'Phase 2',
      time: 'Month 3–4 · Nudge Automation',
      title: 'Instant UPI Fallback',
      pts: [
        'Integrate WhatsApp Business API for instant affluent client nudges',
        'Build secure 1-click UPI Autopay / Intent URL flows',
        'Deploy T+0 reconciliation ledgers for instant success feedback',
        'A/B test nudge copy for highest conversion (target: 65%+)',
      ],
      color: C.orange,
    },
    {
      phase: 'Phase 3',
      time: 'Month 5–6 · Optimisation',
      title: 'Predictive Execution at Scale',
      pts: [
        'Deploy ML models to predict SIP failures *before* initiation (based on historic liquidity timing)',
        'Automate retry scheduling around optimal client balance windows',
        'Scale architecture to handle 10x AUM throughput with zero degradation',
        'Release impact reports demonstrating AUM protected and yield preserved',
      ],
      color: '#A78BFA',
    },
  ];

  phases.forEach((p, i) => {
    const x = 0.5 + i * 3.1;
    slide.addShape(prs.ShapeType.rect, {
      x, y: 1.42, w: 2.88, h: 4.5,
      fill: { color: C.hero }, line: { color: p.color, pt: 1.5, transparency: 30 }, rectRadius: 0.1,
    });
    slide.addText(p.phase, {
      x: x + 0.14, y: 1.58, w: 2.62, h: 0.28,
      fontSize: 9, bold: true, color: p.color, charSpacing: 2, fontFace: 'Calibri',
    });
    slide.addText(p.time, {
      x: x + 0.14, y: 1.86, w: 2.62, h: 0.28,
      fontSize: 10, color: C.muted, fontFace: 'Calibri', italic: true,
    });
    slide.addText(p.title, {
      x: x + 0.14, y: 2.18, w: 2.62, h: 0.44,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri', lineSpacingMultiple: 1.2,
    });
    p.pts.forEach((pt, j) => {
      slide.addText('• ' + pt, {
        x: x + 0.14, y: 2.68 + j * 0.74, w: 2.62, h: 0.66,
        fontSize: 9, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
      });
    });
  });

  // Ask box
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.1, w: 9, h: 1.02,
    fill: { color: C.hero }, line: { color: C.bright, pt: 1.5, transparency: 30 }, rectRadius: 0.1,
  });
  slide.addText('What I need to build this at Dezerv:', {
    x: 0.72, y: 6.2, w: 3, h: 0.28,
    fontSize: 10, bold: true, color: C.bright, fontFace: 'Calibri',
  });
  slide.addText('Access to historic SIP execution drop-off funnels  ·  Alignment with Engineering on PG failover SLAs  ·  1 backend dev for WhatsApp API routing  ·  Ops team bandwidth for beta testing  ·  Clear runway to own the final P&L execution metrics', {
    x: 0.72, y: 6.5, w: 8.6, h: 0.54,
    fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
  });

  // Contact footer
  slide.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade', {
    x: 0.5, y: 7.2, w: 9, h: 0.26,
    fontSize: 8.5, color: C.muted, fontFace: 'Calibri', align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/dezerv-smartsip-deck.pptx' })
  .then(() => console.log('dezerv-smartsip-deck.pptx generated'))
  .catch(e => console.error(e));
