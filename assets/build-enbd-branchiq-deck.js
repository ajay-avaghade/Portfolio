const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';
const W = 10, H = 5.625;

const C = {
  dark:   '0D1B35',
  hero:   '162040',
  card:   '1E2D4E',
  gold:   'D4A843',
  teal:   '00957A',
  red:    'EF4444',
  amber:  'F59E0B',
  green:  '22C55E',
  white:  'F8FAFF',
  muted:  '7A8BAA',
  lgray:  'F0F4FF',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.15 });

// ─────────────────────────────────────────────
// SLIDE 1 — COVER
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  // Diagonal stripe motif
  for (let i = 0; i < 18; i++) {
    const lx = -1 + i * 0.75;
    const lw = Math.min(3.5, W - lx - 0.15);
    if (lx > W || lw <= 0) continue;
    s.addShape(prs.ShapeType.rect, {
      x: lx, y: 0, w: lw, h: H,
      fill: { color: 'FFFFFF', transparency: 97 },
      line: { color: 'FFFFFF', transparency: 100 },
    });
  }

  // Top tag
  s.addText('EMIRATES NBD × VIRTUSA', {
    x:0.5, y:0.4, w:9, h:0.28,
    fontSize:9, bold:true, color:C.gold, align:'left', charSpacing:2,
  });

  // Main title
  s.addText('ENBD BranchIQ', {
    x:0.5, y:0.95, w:7.5, h:0.85,
    fontSize:46, bold:true, color:C.white, align:'left',
  });

  s.addText('Smart Queue & Branch Service Orchestration', {
    x:0.5, y:1.78, w:7.5, h:0.4,
    fontSize:16, bold:false, color:C.muted, align:'left',
  });

  // Gold accent line
  s.addShape(prs.ShapeType.rect, {
    x:0.5, y:2.28, w:2.4, h:0.05,
    fill:{ color:C.gold }, line:{ color:C.gold },
  });

  // Presenter line
  s.addText('Presented by Ajay Avaghade · Service Product Owner', {
    x:0.5, y:2.5, w:7.5, h:0.25,
    fontSize:11, color:C.muted, align:'left',
  });

  // Big stat block bottom-right
  s.addShape(prs.ShapeType.rect, {
    x:7.2, y:3.4, w:2.5, h:1.8,
    fill:{ color:'102030' }, line:{ color:C.gold, pt:1.5 }, rectRadius:0.12,
    shadow: makeShadow(),
  });
  s.addText('14 min', {
    x:7.2, y:3.6, w:2.5, h:0.55,
    fontSize:30, bold:true, color:C.gold, align:'center',
  });
  s.addText('→ 4 min', {
    x:7.2, y:4.05, w:2.5, h:0.35,
    fontSize:18, bold:true, color:C.white, align:'center',
  });
  s.addText('Avg Branch Wait', {
    x:7.2, y:4.38, w:2.5, h:0.22,
    fontSize:9, color:C.muted, align:'center',
  });

  // Subtitle tag
  s.addText('Branch Transformation · AI Queue Intelligence · Mobile Pre-Booking', {
    x:0.5, y:5.1, w:6.5, h:0.25,
    fontSize:9, color:C.muted, align:'left',
  });
}

// ─────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE PROBLEM', {
    x:0.5, y:0.35, w:9, h:0.25,
    fontSize:9, bold:true, color:C.gold, charSpacing:2, align:'left',
  });
  s.addText('ENBD Branches Are Losing Customers in the Queue — Before They Even Reach a Desk', {
    x:0.5, y:0.7, w:9, h:0.75,
    fontSize:22, bold:true, color:C.white, align:'left',
  });

  const cols = [
    { icon:'⏳', stat:'34 min', label:'Avg walk-in wait', sub:'Peak: >45 min at high-traffic branches' },
    { icon:'😤', stat:'1 in 3', label:'Customers abandon queue', sub:'Leave without service — deflect to competitor' },
    { icon:'📋', stat:'0%', label:'Pre-booked appointments', sub:'All walk-ins; zero digital slot management' },
  ];
  cols.forEach((c,i) => {
    const x = 0.5 + i * 3.1;
    s.addShape(prs.ShapeType.rect, {
      x, y:1.6, w:2.9, h:2.6,
      fill:{ color:C.card }, line:{ color:'FFFFFF', transparency:92 }, rectRadius:0.1,
    });
    s.addText(c.icon, { x, y:1.75, w:2.9, h:0.45, fontSize:26, align:'center' });
    s.addText(c.stat, { x, y:2.22, w:2.9, h:0.5, fontSize:30, bold:true, color:C.gold, align:'center' });
    s.addText(c.label, { x, y:2.7, w:2.9, h:0.3, fontSize:11, bold:true, color:C.white, align:'center' });
    s.addText(c.sub, { x, y:3.02, w:2.9, h:0.5, fontSize:9, color:C.muted, align:'center' });
  });

  s.addShape(prs.ShapeType.rect, {
    x:0.5, y:4.4, w:9, h:0.9,
    fill:{ color:C.card }, line:{ color:C.gold, transparency:70 }, rectRadius:0.08,
  });
  s.addText('Root cause: Emirates NBD has no real-time queue visibility, no pre-booking capability, and no digital deflection layer. Every service encounter starts with a wait — and that wait is the bank\'s biggest NPS killer.', {
    x:0.65, y:4.5, w:8.7, h:0.7,
    fontSize:10, color:C.muted, align:'left',
  });
}

// ─────────────────────────────────────────────
// SLIDE 3 — THE INSIGHT
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE INSIGHT', {
    x:0.5, y:0.35, w:9, h:0.25,
    fontSize:9, bold:true, color:C.dark, charSpacing:2, align:'left',
  });
  s.addText('The Queue IS the Branch Experience. Fix the Queue, Fix the NPS.', {
    x:0.5, y:0.68, w:9, h:0.6,
    fontSize:22, bold:true, color:C.dark, align:'left',
  });

  // Left: Current
  s.addShape(prs.ShapeType.rect, {
    x:0.4, y:1.45, w:4.0, h:3.3,
    fill:{ color:'FFF0F0' }, line:{ color:'EF4444', transparency:60 }, rectRadius:0.1,
  });
  s.addText('❌  Walk-in Only (Today)', { x:0.5, y:1.6, w:3.8, h:0.3, fontSize:12, bold:true, color:'CC0000' });
  const curr = ['Customer arrives → joins queue with no info','Waits 20–45 min with no visibility','Staff have no advance preparation time','High abandonment at peak hours','Branch manager has no real-time data','Zero digital alternatives offered'];
  curr.forEach((t,i) => s.addText('· ' + t, { x:0.55, y:2.02+i*0.36, w:3.7, h:0.3, fontSize:10, color:'333333' }));

  // VS divider
  s.addShape(prs.ShapeType.ellipse, {
    x:4.5, y:2.7, w:0.8, h:0.8,
    fill:{ color:C.dark }, line:{ color:C.dark },
  });
  s.addText('VS', { x:4.5, y:2.72, w:0.8, h:0.35, fontSize:10, bold:true, color:C.white, align:'center' });

  // Right: Proposed
  s.addShape(prs.ShapeType.rect, {
    x:5.5, y:1.45, w:4.05, h:3.3,
    fill:{ color:'F0FDF4' }, line:{ color:'22C55E', transparency:60 }, rectRadius:0.1,
  });
  s.addText('✅  BranchIQ (Proposed)', { x:5.6, y:1.6, w:3.8, h:0.3, fontSize:12, bold:true, color:'15803D' });
  const prop = ['Mobile pre-booking → customer arrives on time','Real-time wait display in app → zero surprises','Staff see appointment context before meeting','AI capacity management alerts branch ops','Live dashboard for branch & network managers','Digital deflection for simple service requests'];
  prop.forEach((t,i) => s.addText('· ' + t, { x:5.65, y:2.02+i*0.36, w:3.8, h:0.3, fontSize:10, color:'1a1a1a' }));

  // Bottom callout
  s.addShape(prs.ShapeType.rect, {
    x:0.5, y:4.85, w:9, h:0.55,
    fill:{ color:C.dark }, line:{ color:C.dark }, rectRadius:0.08,
  });
  s.addText('Key insight: A pre-booked branch visit costs the bank nothing extra — and converts a frustrated walk-in into a high-NPS engagement. The queue is the product.', {
    x:0.6, y:4.92, w:8.8, h:0.4,
    fontSize:10, color:C.gold, align:'center',
  });
}

// ─────────────────────────────────────────────
// SLIDE 4 — THE MECHANIC
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE MECHANIC', {
    x:0.5, y:0.35, w:9, h:0.25,
    fontSize:9, bold:true, color:C.gold, charSpacing:2, align:'left',
  });
  s.addText('How BranchIQ Works: Pre-Book → Arrive → Track → Complete → Learn', {
    x:0.5, y:0.68, w:9, h:0.55,
    fontSize:20, bold:true, color:C.white, align:'left',
  });

  const steps = [
    { n:'01', title:'Real-Time\nWait Display', body:'Customer opens ENBD X app → sees all nearby branches with live queue lengths and wait times. Color-coded: green/amber/red.' },
    { n:'02', title:'Slot Booking\n& Confirmation', body:'Customer selects service type, picks a 30-min slot, receives QR code confirmation. RM is pre-briefed on service type.' },
    { n:'03', title:'In-Branch\nCheck-In', body:'Customer scans QR or taps NFC on arrival. Queue position shown in app. RM assigned. ~2 min wait vs 34 min walk-in.' },
    { n:'04', title:'Ops Console\nLive View', body:'Branch manager sees real-time queue, staff utilisation, breach alerts (>30 min threshold). One-click staff reassignment.' },
    { n:'05', title:'Post-Visit\nFeedback Loop', body:'Automated NPS survey sent 15 min post-visit. Verbatims tagged by NLP → piped to FeedbackLoop analytics backlog.' },
  ];

  steps.forEach((st, i) => {
    const x = 0.3 + i * 1.88;
    s.addShape(prs.ShapeType.rect, {
      x, y:1.45, w:1.72, h:3.5,
      fill:{ color:C.card }, line:{ color:'FFFFFF', transparency:90 }, rectRadius:0.1,
    });
    s.addShape(prs.ShapeType.ellipse, {
      x: x + 0.56, y:1.55, w:0.6, h:0.6,
      fill:{ color:C.gold }, line:{ color:C.gold },
    });
    s.addText(st.n, { x: x+0.56, y:1.57, w:0.6, h:0.4, fontSize:13, bold:true, color:C.dark, align:'center' });
    s.addText(st.title, { x, y:2.3, w:1.72, h:0.55, fontSize:10, bold:true, color:C.white, align:'center' });
    s.addText(st.body, { x: x+0.1, y:2.95, w:1.52, h:1.8, fontSize:8.5, color:C.muted, align:'left' });

    if (i < 4) {
      s.addText('→', { x: x+1.72, y:2.15, w:0.16, h:0.3, fontSize:11, color:C.gold, align:'center' });
    }
  });

  s.addShape(prs.ShapeType.rect, {
    x:0.5, y:5.1, w:9, h:0.3,
    fill:{ color:C.card }, line:{ color:'FFFFFF', transparency:95 }, rectRadius:0.06,
  });
  s.addText('Proof: Same deterministic workflow automation architecture as PhonePe ops intelligence layer — 6 siloed workflows → unified orchestration (TAT 2 days → 4 hours)', {
    x:0.6, y:5.14, w:8.8, h:0.22,
    fontSize:8.5, color:C.muted, align:'center',
  });
}

// ─────────────────────────────────────────────
// SLIDE 5 — THE PRODUCT
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE PRODUCT', {
    x:0.5, y:0.35, w:9, h:0.25,
    fontSize:9, bold:true, color:C.dark, charSpacing:2, align:'left',
  });
  s.addText('4 Key Screens — Built as Interactive Prototype', {
    x:0.5, y:0.68, w:9, h:0.45,
    fontSize:20, bold:true, color:C.dark, align:'left',
  });

  const screens = [
    { n:'S1', title:'Branch\nLocator', desc:'Live wait times, colour-coded by severity. User picks best branch before leaving home.', ascii:'[ 🟢 DIFC      8 min ]\n[ 🟡 Al Barsha 12 min ]\n[ 🔴 MoE      31 min ]' },
    { n:'S2', title:'Slot\nBooking', desc:'Service type grid, time chip selector, 30-min slots. QR code confirmation sent.', ascii:'[Home Loan  ] [Cards   ]\n[Investment ] [General ]\n\n⏰  10:30 AM  ← selected' },
    { n:'S3', title:'In-Branch\nCheck-In', desc:'QR scan, ticket A-047, queue #3, ~2 min wait. RM Aisha Al-Mansouri assigned.', ascii:'  🎉 Checked In!\n  Ticket: A-047\n  Position: #3\n  Est. Wait: ~2 min' },
    { n:'S4', title:'Ops Console\n(Web)', desc:'Alert banner: MoE 31 min breach. 47 in queue network-wide. Staff reassignment.', ascii:'⚠ MoE > 30 min threshold\n┌──────────────────────┐\n│ 47 queue │ 14m avg   │\n│ 68% pre-booked       │\n└──────────────────────┘' },
  ];

  screens.forEach((sc, i) => {
    const x = 0.3 + i * 2.35;
    s.addShape(prs.ShapeType.rect, {
      x, y:1.3, w:2.2, h:4.0,
      fill:{ color:'FFFFFF' }, line:{ color:C.dark, transparency:85 }, rectRadius:0.1,
      shadow: makeShadow(),
    });
    s.addShape(prs.ShapeType.rect, {
      x, y:1.3, w:2.2, h:0.38,
      fill:{ color:C.dark }, line:{ color:C.dark }, rectRadius:0.1,
    });
    s.addText(sc.n + ' — ' + sc.title.replace('\n',' '), {
      x: x+0.08, y:1.34, w:2.04, h:0.3,
      fontSize:9, bold:true, color:C.gold, align:'left',
    });

    s.addShape(prs.ShapeType.rect, {
      x: x+0.12, y:1.82, w:1.96, h:1.6,
      fill:{ color:'0A1224' }, line:{ color:'0A1224' }, rectRadius:0.06,
    });
    s.addText(sc.ascii, {
      x: x+0.14, y:1.86, w:1.92, h:1.52,
      fontSize:7, color:C.gold, fontFace:'Courier New', align:'left',
    });
    s.addText(sc.desc, {
      x: x+0.1, y:3.55, w:2.0, h:1.0,
      fontSize:8.5, color:'333333', align:'left',
    });
  });
}

// ─────────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('IMPACT & ROI', {
    x:0.5, y:0.35, w:9, h:0.25,
    fontSize:9, bold:true, color:C.gold, charSpacing:2, align:'left',
  });
  s.addText('Projected Impact — Grounded in PhonePe Delivery Proof Points', {
    x:0.5, y:0.68, w:9, h:0.5,
    fontSize:20, bold:true, color:C.white, align:'left',
  });

  // Left column — Customer & Branch impact
  s.addShape(prs.ShapeType.rect, {
    x:0.4, y:1.35, w:4.3, h:3.7,
    fill:{ color:C.card }, line:{ color:'FFFFFF', transparency:92 }, rectRadius:0.1,
  });
  s.addText('CUSTOMER & BRANCH IMPACT', { x:0.55, y:1.5, w:4.0, h:0.25, fontSize:9, bold:true, color:C.gold, charSpacing:1 });
  const left = [
    { v:'−76%', l:'Average wait time', s:'34 min → 8 min (pre-booked visits)' },
    { v:'−67%', l:'Queue abandonment', s:'From 1-in-3 to near-zero for pre-booked' },
    { v:'+18pt', l:'Branch NPS uplift', s:'Queue wait is top NPS detractor — direct causal link' },
    { v:'68%', l:'Pre-booked appointment rate', s:'Target by Month 6 — mirrors PhonePe PG onboarding adoption' },
  ];
  left.forEach((m,i) => {
    s.addText(m.v, { x:0.55, y:1.92+i*0.76, w:4.0, h:0.38, fontSize:24, bold:true, color:C.gold });
    s.addText(m.l, { x:0.55, y:2.22+i*0.76, w:4.0, h:0.22, fontSize:11, bold:true, color:C.white });
    s.addText(m.s, { x:0.55, y:2.42+i*0.76, w:4.0, h:0.2, fontSize:9, color:C.muted });
  });

  // Right column — Bank ROI
  s.addShape(prs.ShapeType.rect, {
    x:5.3, y:1.35, w:4.3, h:3.7,
    fill:{ color:C.card }, line:{ color:'FFFFFF', transparency:92 }, rectRadius:0.1,
  });
  s.addText('EMIRATES NBD BUSINESS ROI', { x:5.45, y:1.5, w:4.0, h:0.25, fontSize:9, bold:true, color:C.teal, charSpacing:1 });
  const right = [
    { v:'34%', l:'Digital deflection rate', s:'Simple queries handled via ENBD X — branch cost saved' },
    { v:'+AED 220', l:'Revenue per pre-booked visit', s:'RM-briefed visit converts at 3.4× walk-in rate' },
    { v:'−AED 1.8M', l:'Annual branch op cost saving', s:'Reduced staffing overflow, ticket printing, lobby mgmt' },
    { v:'4 months', l:'Payback period', s:'Based on 12-branch pilot rollout (Virtusa delivery model)' },
  ];
  right.forEach((m,i) => {
    s.addText(m.v, { x:5.45, y:1.92+i*0.76, w:4.0, h:0.38, fontSize:24, bold:true, color:C.teal });
    s.addText(m.l, { x:5.45, y:2.22+i*0.76, w:4.0, h:0.22, fontSize:11, bold:true, color:C.white });
    s.addText(m.s, { x:5.45, y:2.42+i*0.76, w:4.0, h:0.2, fontSize:9, color:C.muted });
  });

  s.addShape(prs.ShapeType.rect, {
    x:0.5, y:5.1, w:9, h:0.32,
    fill:{ color:C.card }, line:{ color:'FFFFFF', transparency:95 }, rectRadius:0.06,
  });
  s.addText('The pre-booked visit pays for itself: RM preparation time → higher product conversion → AED 220 incremental revenue per visit covers infrastructure cost in Month 4.', {
    x:0.6, y:5.14, w:8.8, h:0.24,
    fontSize:8.5, color:C.muted, align:'center',
  });
}

// ─────────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('PROOF OF WORK', {
    x:0.5, y:0.35, w:9, h:0.25,
    fontSize:9, bold:true, color:C.dark, charSpacing:2, align:'left',
  });
  s.addText("I Built This. Here's the PhonePe Architecture That Maps Directly.", {
    x:0.5, y:0.68, w:9, h:0.5,
    fontSize:20, bold:true, color:C.dark, align:'left',
  });

  // Left — PhonePe context
  s.addShape(prs.ShapeType.rect, {
    x:0.4, y:1.35, w:4.3, h:3.6,
    fill:{ color:C.dark }, line:{ color:C.dark }, rectRadius:0.1,
  });
  s.addText('WHAT I SHIPPED AT PHONEPE', { x:0.55, y:1.5, w:4.0, h:0.25, fontSize:9, bold:true, color:C.gold, charSpacing:1 });
  const pp = [
    'Replaced 6 siloed analyst workflows with unified ops intelligence layer',
    'TAT: 2 days → 4 hours. −68% escalation tickets within 60 days',
    'Built deterministic event orchestration: trigger → check → action → notify',
    'Designed ops console with live queue, status tracking, and breach alerting',
    'Context-aware nudge engine → staff intervention at threshold crossings',
    'Instrumented all KPIs before build — same approach as BranchIQ ops console',
  ];
  pp.forEach((t,i) => s.addText('· ' + t, { x:0.55, y:1.9+i*0.44, w:4.0, h:0.36, fontSize:9.5, color:C.white, align:'left' }));

  // Right — JD mapping
  s.addShape(prs.ShapeType.rect, {
    x:5.3, y:1.35, w:4.3, h:3.6,
    fill:{ color:'E8F5F2' }, line:{ color:C.teal, transparency:70 }, rectRadius:0.1,
  });
  s.addText('HOW IT MAPS TO THIS ROLE', { x:5.45, y:1.5, w:4.0, h:0.25, fontSize:9, bold:true, color:C.teal, charSpacing:1 });
  const jd = [
    '→ BranchIQ ops console: real-time queue + breach alerts',
    '→ Pre-booking = ops capacity management (same KPI design)',
    '→ Same orchestration: booking event → prep briefing → assign RM',
    '→ Same live dashboard pattern — branch vs network vs region view',
    '→ Digital deflection nudge = same nudge architecture, different channel',
    '→ Virtusa: I deliver Agile sprints with client stakeholder co-ownership',
  ];
  jd.forEach((t,i) => s.addText(t, { x:5.45, y:1.9+i*0.44, w:4.0, h:0.36, fontSize:9.5, color:'1a1a1a', align:'left' }));

  // Quote
  s.addShape(prs.ShapeType.rect, {
    x:0.5, y:5.07, w:9, h:0.38,
    fill:{ color:C.dark }, line:{ color:C.dark }, rectRadius:0.08,
  });
  s.addText('"This isn\'t a pitch about what I\'d build — it\'s a description of what I already shipped, applied to ENBD\'s exact problem."', {
    x:0.6, y:5.12, w:8.8, h:0.28,
    fontSize:9.5, italic:true, color:C.gold, align:'center',
  });
}

// ─────────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('ROLLOUT PLAN', {
    x:0.5, y:0.35, w:9, h:0.25,
    fontSize:9, bold:true, color:C.gold, charSpacing:2, align:'left',
  });
  s.addText('3-Phase Delivery — Virtusa Agile Model', {
    x:0.5, y:0.68, w:9, h:0.45,
    fontSize:20, bold:true, color:C.white, align:'left',
  });

  const phases = [
    {
      n:'Phase 1', sub:'Months 1–2',
      title:'Pilot — 3 Branches',
      items:[
        'Deploy BranchIQ (pre-booking + check-in) to DIFC, Downtown, JBR',
        'Instrument real-time queue and pre-booking KPIs from Day 1',
        'Ops console live for branch managers — breach alert threshold set',
        'Collect NPS before/after for A/B comparison vs walk-in baseline',
      ],
    },
    {
      n:'Phase 2', sub:'Months 3–4',
      title:'Expand + Integrate',
      items:[
        'Roll out to 12 branches — include Al Barsha (anomaly branch)',
        'Integrate FeedbackLoop NPS tagging → backlog auto-population',
        'Launch AdvisorPulse RM dashboard in DIFC and Downtown',
        'Digital deflection flow live — track deflection rate weekly',
      ],
    },
    {
      n:'Phase 3', sub:'Months 5–6',
      title:'Network-Wide + Optimise',
      items:[
        'Full network rollout (all Emirates NBD UAE branches)',
        'AI capacity management: ML-based staff allocation per slot band',
        'FeedbackLoop anomaly alerts live for Branch Quality team',
        'Target: 68% pre-booking rate, <8 min avg wait, NPS +18pt',
      ],
    },
  ];

  phases.forEach((p, i) => {
    const x = 0.3 + i * 3.12;
    s.addShape(prs.ShapeType.rect, {
      x, y:1.3, w:2.95, h:3.45,
      fill:{ color:C.card }, line:{ color:'FFFFFF', transparency:90 }, rectRadius:0.1,
    });
    s.addShape(prs.ShapeType.rect, {
      x, y:1.3, w:2.95, h:0.55,
      fill:{ color:i===0?C.gold:i===1?C.teal:'5B8AF5' }, line:{ color:'FFFFFF', transparency:100 }, rectRadius:0.1,
    });
    s.addText(p.n + ' · ' + p.sub, { x: x+0.1, y:1.35, w:2.75, h:0.22, fontSize:9, bold:true, color:C.dark });
    s.addText(p.title, { x: x+0.1, y:1.55, w:2.75, h:0.22, fontSize:9.5, bold:true, color:C.dark });
    p.items.forEach((item, j) => {
      s.addText('· ' + item, { x: x+0.12, y:2.0+j*0.56, w:2.72, h:0.5, fontSize:9, color:C.muted, align:'left' });
    });
  });

  // What I need box
  s.addShape(prs.ShapeType.rect, {
    x:0.4, y:4.87, w:9.2, h:0.55,
    fill:{ color:C.card }, line:{ color:C.gold, transparency:70 }, rectRadius:0.08,
  });
  s.addText('What I need to build this:  Access to branch queue & booking APIs  ·  Alignment with ENBD Digital & Branch Ops teams  ·  1 engineer (mobile) + 1 (ops console backend)  ·  Virtusa Agile sprint cadence', {
    x:0.55, y:4.93, w:9.0, h:0.4,
    fontSize:9, color:C.gold, align:'center',
  });

  s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  Planning UAE relocation Q2 2026', {
    x:0.5, y:5.38, w:9, h:0.2,
    fontSize:8.5, color:C.muted, align:'center',
  });
}

prs.writeFile({ fileName: 'assets/enbd-branchiq-deck.pptx' })
  .then(() => console.log('✅  enbd-branchiq-deck.pptx saved'))
  .catch(e => console.error(e));
