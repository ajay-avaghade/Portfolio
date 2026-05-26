const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';
const W = 10, H = 5.625;

const C = {
  dark:   '0D1B35',
  hero:   '162040',
  card:   '1E2D4E',
  teal:   '00957A',
  teallt: '00BFA0',
  gold:   'D4A843',
  blue:   '4A90D9',
  green:  '22C55E',
  amber:  'F59E0B',
  white:  'F8FAFF',
  muted:  '7A8BAA',
  lgray:  'F0F6FF',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.15 });

// ─────────────────────────────────────────────
// SLIDE 1 — COVER
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  for (let i = 0; i < 18; i++) {
    const lx = -1 + i * 0.75;
    const lw = Math.min(3.5, W - lx - 0.15);
    if (lx > W || lw <= 0) continue;
    s.addShape(prs.ShapeType.rect, { x:lx, y:0, w:lw, h:H, fill:{color:'FFFFFF',transparency:97}, line:{color:'FFFFFF',transparency:100} });
  }
  s.addText('EMIRATES NBD × VIRTUSA', { x:0.5, y:0.4, w:9, h:0.28, fontSize:9, bold:true, color:C.teal, align:'left', charSpacing:2 });
  s.addText('ENBD AdvisorPulse', { x:0.5, y:0.95, w:8, h:0.85, fontSize:46, bold:true, color:C.white, align:'left' });
  s.addText('AI-Powered RM Intelligence Dashboard', { x:0.5, y:1.78, w:8, h:0.4, fontSize:16, color:C.muted, align:'left' });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:2.28, w:2.4, h:0.05, fill:{color:C.teal}, line:{color:C.teal} });
  s.addText('Presented by Ajay Avaghade · Service Product Owner', { x:0.5, y:2.5, w:8, h:0.25, fontSize:11, color:C.muted, align:'left' });

  s.addShape(prs.ShapeType.rect, { x:7.2, y:3.4, w:2.5, h:1.8, fill:{color:'102030'}, line:{color:C.teal,pt:1.5}, rectRadius:0.12, shadow:makeShadow() });
  s.addText('87%', { x:7.2, y:3.6, w:2.5, h:0.55, fontSize:34, bold:true, color:C.teal, align:'center' });
  s.addText('Conversion\nConfidence', { x:7.2, y:4.08, w:2.5, h:0.45, fontSize:11, bold:true, color:C.white, align:'center' });
  s.addText('AI-predicted per visit', { x:7.2, y:4.5, w:2.5, h:0.22, fontSize:9, color:C.muted, align:'center' });
  s.addText('RM Intelligence · Life Event Signals · AI Conversation Guide', { x:0.5, y:5.1, w:6.5, h:0.25, fontSize:9, color:C.muted, align:'left' });
}

// ─────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('THE PROBLEM', { x:0.5, y:0.35, w:9, h:0.25, fontSize:9, bold:true, color:C.teal, charSpacing:2, align:'left' });
  s.addText('ENBD Relationship Managers Are Walking Into Customer Meetings Blind', { x:0.5, y:0.7, w:9, h:0.6, fontSize:21, bold:true, color:C.white, align:'left' });

  const cols = [
    { icon:'🗂️', stat:'6+', label:'Systems an RM checks', sub:'CRM, core banking, portfolio, loans — no unified view' },
    { icon:'⏱️', stat:'18 min', label:'Avg prep time per meeting', sub:'Wasted searching disconnected data silos' },
    { icon:'📉', stat:'1 in 4', label:'Visits miss upsell signal', sub:'Life events not surfaced — opportunity lost' },
  ];
  cols.forEach((c,i) => {
    const x = 0.5 + i * 3.1;
    s.addShape(prs.ShapeType.rect, { x, y:1.6, w:2.9, h:2.6, fill:{color:C.card}, line:{color:'FFFFFF',transparency:92}, rectRadius:0.1 });
    s.addText(c.icon, { x, y:1.75, w:2.9, h:0.45, fontSize:26, align:'center' });
    s.addText(c.stat, { x, y:2.22, w:2.9, h:0.5, fontSize:30, bold:true, color:C.teal, align:'center' });
    s.addText(c.label, { x, y:2.7, w:2.9, h:0.3, fontSize:11, bold:true, color:C.white, align:'center' });
    s.addText(c.sub, { x, y:3.02, w:2.9, h:0.5, fontSize:9, color:C.muted, align:'center' });
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:4.4, w:9, h:0.9, fill:{color:C.card}, line:{color:C.teal,transparency:70}, rectRadius:0.08 });
  s.addText('Root cause: RM context is fragmented across 6+ systems with no intelligent aggregation. An RM meeting a Platinum customer with an expiring term deposit has the same visibility as a teller handing out a token. That\'s a missed AED 600K conversation.', {
    x:0.65, y:4.5, w:8.7, h:0.7, fontSize:10, color:C.muted, align:'left',
  });
}

// ─────────────────────────────────────────────
// SLIDE 3 — THE INSIGHT
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };
  s.addText('THE INSIGHT', { x:0.5, y:0.35, w:9, h:0.25, fontSize:9, bold:true, color:C.dark, charSpacing:2, align:'left' });
  s.addText('An Informed RM Is a Revenue Event. Every Missed Signal Is a Missed Conversion.', { x:0.5, y:0.68, w:9, h:0.6, fontSize:21, bold:true, color:C.dark, align:'left' });

  s.addShape(prs.ShapeType.rect, { x:0.4, y:1.45, w:4.0, h:3.3, fill:{color:'FFF0F0'}, line:{color:'EF4444',transparency:60}, rectRadius:0.1 });
  s.addText('❌  RM Today (Uninformed)', { x:0.5, y:1.6, w:3.8, h:0.3, fontSize:12, bold:true, color:'CC0000' });
  const curr = ['Opens meeting with no context about customer','Spends 10 min asking questions already in CRM','Misses term deposit maturity in 4 days','No product recommendation framework','Outcome logged manually — no AI assistance','Upsell rate: 12% per visit'];
  curr.forEach((t,i) => s.addText('· '+t, { x:0.55, y:2.02+i*0.36, w:3.7, h:0.3, fontSize:10, color:'333333' }));

  s.addShape(prs.ShapeType.ellipse, { x:4.5, y:2.7, w:0.8, h:0.8, fill:{color:C.dark}, line:{color:C.dark} });
  s.addText('VS', { x:4.5, y:2.72, w:0.8, h:0.35, fontSize:10, bold:true, color:C.white, align:'center' });

  s.addShape(prs.ShapeType.rect, { x:5.5, y:1.45, w:4.05, h:3.3, fill:{color:'F0FDF9'}, line:{color:'00957A',transparency:60}, rectRadius:0.1 });
  s.addText('✅  RM with AdvisorPulse', { x:5.6, y:1.6, w:3.8, h:0.3, fontSize:12, bold:true, color:'00694F' });
  const prop = ['One-screen 360° view pre-loaded before meeting','AI-flagged: TD maturing in 4 days (AED 600K)','3 ranked conversation priorities with talk tracks','Product recommendation with conversion confidence %','Post-visit action items auto-populated','Upsell rate: 38% per visit (+217%)'];
  prop.forEach((t,i) => s.addText('· '+t, { x:5.65, y:2.02+i*0.36, w:3.8, h:0.3, fontSize:10, color:'1a1a1a' }));

  s.addShape(prs.ShapeType.rect, { x:0.5, y:4.85, w:9, h:0.55, fill:{color:C.dark}, line:{color:C.dark}, rectRadius:0.08 });
  s.addText('Key insight: The intelligence already exists in ENBD\'s data — it just isn\'t aggregated. AdvisorPulse is a synthesis layer, not a new data source. Build once, surface everywhere.', {
    x:0.6, y:4.92, w:8.8, h:0.4, fontSize:10, color:C.teal, align:'center',
  });
}

// ─────────────────────────────────────────────
// SLIDE 4 — THE MECHANIC
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('THE MECHANIC', { x:0.5, y:0.35, w:9, h:0.25, fontSize:9, bold:true, color:C.teal, charSpacing:2, align:'left' });
  s.addText('How AdvisorPulse Works: Data Synthesis → Signal Detection → Guided Conversation', { x:0.5, y:0.68, w:9, h:0.55, fontSize:19, bold:true, color:C.white, align:'left' });

  const steps = [
    { n:'01', title:'Data\nIngestion', body:'Core banking, CRM, portfolio, loan, digital activity pulled into unified customer profile. Updated nightly + real-time for high-signal events.' },
    { n:'02', title:'ML Signal\nDetection', body:'Life event models score each customer for 12 signal types: expiry events, property intent, portfolio gaps, life stage markers. Updated on each login.' },
    { n:'03', title:'Pre-Visit\nBriefing', body:'RM opens dashboard 15 min before appointment. Top 3 conversation priorities ranked by conversion confidence. Talking points pre-loaded.' },
    { n:'04', title:'In-Meeting\nGuide', body:'RM follows AI-suggested flow. Tabs for customer context, product recommendations, competitive benchmarks, and objection handling.' },
    { n:'05', title:'Post-Visit\nLog & CRM', body:'Action items auto-suggested. RM confirms outcomes. One-click push to Salesforce CRM. Next visit priorities pre-seeded.' },
  ];
  steps.forEach((st, i) => {
    const x = 0.3 + i * 1.88;
    s.addShape(prs.ShapeType.rect, { x, y:1.45, w:1.72, h:3.5, fill:{color:C.card}, line:{color:'FFFFFF',transparency:90}, rectRadius:0.1 });
    s.addShape(prs.ShapeType.ellipse, { x:x+0.56, y:1.55, w:0.6, h:0.6, fill:{color:C.teal}, line:{color:C.teal} });
    s.addText(st.n, { x:x+0.56, y:1.57, w:0.6, h:0.4, fontSize:13, bold:true, color:C.dark, align:'center' });
    s.addText(st.title, { x, y:2.3, w:1.72, h:0.55, fontSize:10, bold:true, color:C.white, align:'center' });
    s.addText(st.body, { x:x+0.1, y:2.95, w:1.52, h:1.8, fontSize:8.5, color:C.muted, align:'left' });
    if (i < 4) s.addText('→', { x:x+1.72, y:2.15, w:0.16, h:0.3, fontSize:11, color:C.teal, align:'center' });
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.1, w:9, h:0.3, fill:{color:C.card}, line:{color:'FFFFFF',transparency:95}, rectRadius:0.06 });
  s.addText('Proof: Same ML scoring architecture as PhonePe Propensity-to-Transact model — real-time user-level scoring replacing static rules across 350M+ MAU', {
    x:0.6, y:5.14, w:8.8, h:0.22, fontSize:8.5, color:C.muted, align:'center',
  });
}

// ─────────────────────────────────────────────
// SLIDE 5 — THE PRODUCT
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };
  s.addText('THE PRODUCT', { x:0.5, y:0.35, w:9, h:0.25, fontSize:9, bold:true, color:C.dark, charSpacing:2, align:'left' });
  s.addText('4 Web Dashboard Panels — Built as Interactive Prototype', { x:0.5, y:0.68, w:9, h:0.45, fontSize:20, bold:true, color:C.dark, align:'left' });

  const screens = [
    { n:'P1', title:'Appointment\nQueue', desc:'RM\'s daily view. Active meeting highlighted. AI-ready badge shows intelligence loaded. Completion tracking.', ascii:'[ SA — Sarah Al-Rashidi    ]\n[ 10:30 Khalid  ← ACTIVE  ]\n[ 11:30 Nour    In 52 min ]\n[ 13:00 Farrukh Pre-booked]' },
    { n:'P2', title:'Customer\n360° Card', desc:'Holdings, AUM, products, CSAT, relationship tenure. ML life-event signals with confidence scores.', ascii:'Khalid Al-Farsi · Platinum\nAUM: AED 2.4M\n🔥 TD maturing in 4 days\n🏠 Property intent: 74%' },
    { n:'P3', title:'AI Convo\nGuide', desc:'Top 3 priorities ranked by confidence. Talking points pre-loaded. Conversion probability bar per topic.', ascii:'[P1] TD Renewal  → 87%\n  · "AED 600K matures May 31"\n  · "New 5.8% structured"\n[P2] Portfolio gap → 61%' },
    { n:'P4', title:'Post-Visit\nLog & CRM', desc:'Action items auto-suggested. RM confirms outcomes. One-click Salesforce CRM push with full visit record.', ascii:'Visit: 45 min · 3 actions\n[✓] Send term sheet — today\n[ ] TD decision by May 30\n[→] Push to Salesforce CRM' },
  ];

  screens.forEach((sc, i) => {
    const x = 0.3 + i * 2.35;
    s.addShape(prs.ShapeType.rect, { x, y:1.3, w:2.2, h:4.0, fill:{color:'FFFFFF'}, line:{color:C.dark,transparency:85}, rectRadius:0.1, shadow:makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y:1.3, w:2.2, h:0.38, fill:{color:C.dark}, line:{color:C.dark}, rectRadius:0.1 });
    s.addText(sc.n + ' — ' + sc.title.replace('\n',' '), { x:x+0.08, y:1.34, w:2.04, h:0.3, fontSize:9, bold:true, color:C.teal, align:'left' });
    s.addShape(prs.ShapeType.rect, { x:x+0.12, y:1.82, w:1.96, h:1.6, fill:{color:'0A1224'}, line:{color:'0A1224'}, rectRadius:0.06 });
    s.addText(sc.ascii, { x:x+0.14, y:1.86, w:1.92, h:1.52, fontSize:7, color:C.teal, fontFace:'Courier New', align:'left' });
    s.addText(sc.desc, { x:x+0.1, y:3.55, w:2.0, h:1.0, fontSize:8.5, color:'333333', align:'left' });
  });
}

// ─────────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('IMPACT & ROI', { x:0.5, y:0.35, w:9, h:0.25, fontSize:9, bold:true, color:C.teal, charSpacing:2, align:'left' });
  s.addText('Projected Impact — Grounded in PhonePe ML Delivery Proof Points', { x:0.5, y:0.68, w:9, h:0.5, fontSize:20, bold:true, color:C.white, align:'left' });

  s.addShape(prs.ShapeType.rect, { x:0.4, y:1.35, w:4.3, h:3.7, fill:{color:C.card}, line:{color:'FFFFFF',transparency:92}, rectRadius:0.1 });
  s.addText('RM & CUSTOMER IMPACT', { x:0.55, y:1.5, w:4.0, h:0.25, fontSize:9, bold:true, color:C.teal, charSpacing:1 });
  const left = [
    { v:'+217%', l:'Visit-to-product conversion', s:'12% → 38% per RM visit with AI guidance' },
    { v:'−83%', l:'Pre-meeting prep time', s:'18 min → 3 min (AI pre-briefs, RM just reviews)' },
    { v:'94%', l:'Retention post-briefing', s:'Customers who receive personalised advice stay (TotalRewards parallel)' },
    { v:'+8pt', l:'Customer satisfaction (CSAT)', s:'Informed RM = customer feels valued, not processed' },
  ];
  left.forEach((m,i) => {
    s.addText(m.v, { x:0.55, y:1.92+i*0.76, w:4.0, h:0.38, fontSize:24, bold:true, color:C.teal });
    s.addText(m.l, { x:0.55, y:2.22+i*0.76, w:4.0, h:0.22, fontSize:11, bold:true, color:C.white });
    s.addText(m.s, { x:0.55, y:2.42+i*0.76, w:4.0, h:0.2, fontSize:9, color:C.muted });
  });

  s.addShape(prs.ShapeType.rect, { x:5.3, y:1.35, w:4.3, h:3.7, fill:{color:C.card}, line:{color:'FFFFFF',transparency:92}, rectRadius:0.1 });
  s.addText('EMIRATES NBD REVENUE ROI', { x:5.45, y:1.5, w:4.0, h:0.25, fontSize:9, bold:true, color:C.gold, charSpacing:1 });
  const right = [
    { v:'AED 3.2M', l:'Incremental AUM per RM per year', s:'Based on 38% conversion × 6 visits/day × 220 days × AED 280K avg product' },
    { v:'AED 1.8M', l:'Revenue saved from retained Platinum', s:'$560K replacement cost avoided — direct TotalRewards parallel' },
    { v:'3×', l:'ROI on AdvisorPulse deployment cost', s:'Virtusa delivery cost recovered in first quarter post-launch' },
    { v:'2 months', l:'Pilot to full-RM deployment', s:'Phased by segment: Priority → Preferred → Standard RM' },
  ];
  right.forEach((m,i) => {
    s.addText(m.v, { x:5.45, y:1.92+i*0.76, w:4.0, h:0.38, fontSize:24, bold:true, color:C.gold });
    s.addText(m.l, { x:5.45, y:2.22+i*0.76, w:4.0, h:0.22, fontSize:11, bold:true, color:C.white });
    s.addText(m.s, { x:5.45, y:2.42+i*0.76, w:4.0, h:0.2, fontSize:9, color:C.muted });
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.1, w:9, h:0.32, fill:{color:C.card}, line:{color:'FFFFFF',transparency:95}, rectRadius:0.06 });
  s.addText('The Khalid Al-Farsi scenario in this prototype represents AED 600K of AUM at renewal decision. AdvisorPulse ensures that conversation happens every time — not just when an RM happens to check the right system.', {
    x:0.6, y:5.14, w:8.8, h:0.24, fontSize:8.5, color:C.muted, align:'center',
  });
}

// ─────────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };
  s.addText('PROOF OF WORK', { x:0.5, y:0.35, w:9, h:0.25, fontSize:9, bold:true, color:C.dark, charSpacing:2, align:'left' });
  s.addText("I Built the ML Scoring Architecture. Here's the Direct Parallel.", { x:0.5, y:0.68, w:9, h:0.5, fontSize:20, bold:true, color:C.dark, align:'left' });

  s.addShape(prs.ShapeType.rect, { x:0.4, y:1.35, w:4.3, h:3.6, fill:{color:C.dark}, line:{color:C.dark}, rectRadius:0.1 });
  s.addText('WHAT I SHIPPED AT PHONEPE', { x:0.55, y:1.5, w:4.0, h:0.25, fontSize:9, bold:true, color:C.teal, charSpacing:1 });
  const pp = [
    'Deployed Propensity-to-Transact ML model: real-time user-level scoring',
    'Feature engineering through production rollout across 350M+ MAU',
    'Replaced static rule-based cohorts with dynamic intent signals',
    '−32% marketing spend while sustaining GMV growth',
    'Built dynamic incentivisation: cart × intent × time → personalised nudge',
    'Same ML architecture: signal → score → ranked recommendation → action',
  ];
  pp.forEach((t,i) => s.addText('· '+t, { x:0.55, y:1.9+i*0.44, w:4.0, h:0.36, fontSize:9.5, color:C.white, align:'left' }));

  s.addShape(prs.ShapeType.rect, { x:5.3, y:1.35, w:4.3, h:3.6, fill:{color:'E8F5F2'}, line:{color:C.teal,transparency:70}, rectRadius:0.1 });
  s.addText('HOW IT MAPS TO ADVISORPULSE', { x:5.45, y:1.5, w:4.0, h:0.25, fontSize:9, bold:true, color:C.teal, charSpacing:1 });
  const jd = [
    '→ User intent signals → customer life event signals (same architecture)',
    '→ Feature engineering: cart + intent → portfolio + maturity + activity',
    '→ Static cohorts → real-time customer-level propensity scores',
    '→ −32% wasted spend → −83% wasted RM prep time (same outcome class)',
    '→ Cart trigger → TD maturity trigger → same nudge-engine pattern',
    '→ Virtusa: I own feature spec + DS coordination + delivery sprint',
  ];
  jd.forEach((t,i) => s.addText(t, { x:5.45, y:1.9+i*0.44, w:4.0, h:0.36, fontSize:9.5, color:'1a1a1a', align:'left' }));

  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.07, w:9, h:0.38, fill:{color:C.dark}, line:{color:C.dark}, rectRadius:0.08 });
  s.addText('"Every PhonePe user had a propensity score in real time. Every ENBD customer can have a conversation score in real time. Same architecture, different domain."', {
    x:0.6, y:5.12, w:8.8, h:0.28, fontSize:9.5, italic:true, color:C.teal, align:'center',
  });
}

// ─────────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('ROLLOUT PLAN', { x:0.5, y:0.35, w:9, h:0.25, fontSize:9, bold:true, color:C.teal, charSpacing:2, align:'left' });
  s.addText('3-Phase Delivery — Virtusa Agile Model', { x:0.5, y:0.68, w:9, h:0.45, fontSize:20, bold:true, color:C.white, align:'left' });

  const phases = [
    {
      n:'Phase 1', sub:'Month 1–2',
      title:'Pilot — Priority Segment RMs',
      items:['Deploy AdvisorPulse to 20 Priority/Platinum RMs (DIFC + Downtown)', 'Wire up CRM, core banking, portfolio data feeds', 'Measure: prep time, conversion rate, CSAT vs control group RMs', 'Weekly sprint reviews with ENBD Digital and RM teams'],
    },
    {
      n:'Phase 2', sub:'Month 3–4',
      title:'Expand + ML Signal Tuning',
      items:['Roll out to Preferred segment RMs (50+ RMs, 5 branches)', 'Tune life event signal models on first 2 months of interaction data', 'Launch CRM auto-sync (Salesforce) — action items push on visit close', 'NPS delta measured: pilot RMs vs walk-in baseline'],
    },
    {
      n:'Phase 3', sub:'Month 5–6',
      title:'Network-Wide + Advanced Signals',
      items:['All RM segments and branches across ENBD UAE network', 'Add advanced signals: FeedbackLoop NPS scores + BranchIQ wait context', 'A/B test: AI-guided visit vs standard RM visit — revenue attribution', 'Target: +217% conversion, −83% prep time, AED 3.2M incremental AUM/RM/yr'],
    },
  ];
  phases.forEach((p, i) => {
    const x = 0.3 + i * 3.12;
    s.addShape(prs.ShapeType.rect, { x, y:1.3, w:2.95, h:3.45, fill:{color:C.card}, line:{color:'FFFFFF',transparency:90}, rectRadius:0.1 });
    s.addShape(prs.ShapeType.rect, { x, y:1.3, w:2.95, h:0.55, fill:{color:i===0?C.teal:i===1?C.gold:'5B8AF5'}, line:{color:'FFFFFF',transparency:100}, rectRadius:0.1 });
    s.addText(p.n + ' · ' + p.sub, { x:x+0.1, y:1.35, w:2.75, h:0.22, fontSize:9, bold:true, color:C.dark });
    s.addText(p.title, { x:x+0.1, y:1.55, w:2.75, h:0.22, fontSize:9.5, bold:true, color:C.dark });
    p.items.forEach((item, j) => s.addText('· '+item, { x:x+0.12, y:2.0+j*0.56, w:2.72, h:0.5, fontSize:9, color:C.muted, align:'left' }));
  });

  s.addShape(prs.ShapeType.rect, { x:0.4, y:4.87, w:9.2, h:0.55, fill:{color:C.card}, line:{color:C.teal,transparency:70}, rectRadius:0.08 });
  s.addText('What I need:  Access to CRM + core banking API schema  ·  DS team for signal model tuning  ·  ENBD RM team as design partners  ·  1 engineer (backend) + 1 (frontend)  ·  Virtusa Agile sprint cadence', {
    x:0.55, y:4.93, w:9.0, h:0.4, fontSize:9, color:C.teal, align:'center',
  });
  s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  Planning UAE relocation Q2 2026', {
    x:0.5, y:5.38, w:9, h:0.2, fontSize:8.5, color:C.muted, align:'center',
  });
}

prs.writeFile({ fileName: 'assets/enbd-advisorpulse-deck.pptx' })
  .then(() => console.log('✅  enbd-advisorpulse-deck.pptx saved'))
  .catch(e => console.error(e));
