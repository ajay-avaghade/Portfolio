const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';
const W = 10, H = 5.625;

const C = {
  dark:   '0D1B35',
  hero:   '162040',
  card:   '1E2D4E',
  blue:   '5B8AF5',
  bluelt: '7BA4FF',
  gold:   'D4A843',
  teal:   '00957A',
  red:    'EF4444',
  amber:  'F59E0B',
  green:  '22C55E',
  white:  'F8FAFF',
  muted:  '7A8BAA',
  lgray:  'EEF2FF',
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
  s.addText('EMIRATES NBD × VIRTUSA', { x:0.5, y:0.4, w:9, h:0.28, fontSize:9, bold:true, color:C.blue, align:'left', charSpacing:2 });
  s.addText('ENBD FeedbackLoop', { x:0.5, y:0.95, w:8, h:0.85, fontSize:46, bold:true, color:C.white, align:'left' });
  s.addText('Branch Service Quality Intelligence Platform', { x:0.5, y:1.78, w:8, h:0.4, fontSize:16, color:C.muted, align:'left' });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:2.28, w:2.4, h:0.05, fill:{color:C.blue}, line:{color:C.blue} });
  s.addText('Presented by Ajay Avaghade · Service Product Owner', { x:0.5, y:2.5, w:8, h:0.25, fontSize:11, color:C.muted, align:'left' });
  s.addShape(prs.ShapeType.rect, { x:7.2, y:3.4, w:2.5, h:1.8, fill:{color:'0A1224'}, line:{color:C.blue,pt:1.5}, rectRadius:0.12, shadow:makeShadow() });
  s.addText('+18pt', { x:7.2, y:3.6, w:2.5, h:0.55, fontSize:32, bold:true, color:C.blue, align:'center' });
  s.addText('NPS Uplift', { x:7.2, y:4.1, w:2.5, h:0.28, fontSize:13, bold:true, color:C.white, align:'center' });
  s.addText('from anomaly detection', { x:7.2, y:4.36, w:2.5, h:0.22, fontSize:9, color:C.muted, align:'center' });
  s.addText('NLP Analytics · Anomaly Detection · Jira Backlog Engine', { x:0.5, y:5.1, w:6.5, h:0.25, fontSize:9, color:C.muted, align:'left' });
}

// ─────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('THE PROBLEM', { x:0.5, y:0.35, w:9, h:0.25, fontSize:9, bold:true, color:C.blue, charSpacing:2, align:'left' });
  s.addText("ENBD Collects Customer Feedback. It Doesn't Act on It Fast Enough.", { x:0.5, y:0.7, w:9, h:0.6, fontSize:21, bold:true, color:C.white, align:'left' });

  const cols = [
    { icon:'📊', stat:'67', label:'Network NPS — not bad', sub:'But Al Barsha is at 41. Nobody knew until complaints hit a VP.' },
    { icon:'🗓️', stat:'3–4 wks', label:'Lag from feedback to fix', sub:'Verbatims sit in quarterly Excel exports, never actioned in-cycle' },
    { icon:'🔁', stat:'0%', label:'Feedback → Backlog link', sub:'No systematic path from customer verbatim to product/ops fix' },
  ];
  cols.forEach((c,i) => {
    const x = 0.5 + i * 3.1;
    s.addShape(prs.ShapeType.rect, { x, y:1.6, w:2.9, h:2.6, fill:{color:C.card}, line:{color:'FFFFFF',transparency:92}, rectRadius:0.1 });
    s.addText(c.icon, { x, y:1.75, w:2.9, h:0.45, fontSize:26, align:'center' });
    s.addText(c.stat, { x, y:2.22, w:2.9, h:0.5, fontSize:30, bold:true, color:C.blue, align:'center' });
    s.addText(c.label, { x, y:2.7, w:2.9, h:0.3, fontSize:11, bold:true, color:C.white, align:'center' });
    s.addText(c.sub, { x, y:3.02, w:2.9, h:0.5, fontSize:9, color:C.muted, align:'center' });
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:4.4, w:9, h:0.9, fill:{color:C.card}, line:{color:C.blue,transparency:70}, rectRadius:0.08 });
  s.addText('Root cause: Feedback is collected but siloed — NPS scores in one system, verbatims in another, no NLP tagging, no anomaly alerting, no auto-generated backlog. The gap between "customer complaint" and "Jira ticket" is measured in weeks, not hours.', {
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
  s.addText('Every Customer Complaint Is a Backlog Item in Disguise. Connect Them.', { x:0.5, y:0.68, w:9, h:0.6, fontSize:21, bold:true, color:C.dark, align:'left' });

  s.addShape(prs.ShapeType.rect, { x:0.4, y:1.45, w:4.0, h:3.3, fill:{color:'FFF0F0'}, line:{color:'EF4444',transparency:60}, rectRadius:0.1 });
  s.addText('❌  Today (Disconnected)', { x:0.5, y:1.6, w:3.8, h:0.3, fontSize:12, bold:true, color:'CC0000' });
  const curr = ['Survey results exported to Excel quarterly','Verbatims read manually by branch ops team','No NLP tagging — patterns invisible at scale','NPS drop at Al Barsha unknown for 4 weeks','No link between feedback and backlog tickets','Fix cycle: 3–4 weeks from complaint to action'];
  curr.forEach((t,i) => s.addText('· '+t, { x:0.55, y:2.02+i*0.36, w:3.7, h:0.3, fontSize:10, color:'333333' }));

  s.addShape(prs.ShapeType.ellipse, { x:4.5, y:2.7, w:0.8, h:0.8, fill:{color:C.dark}, line:{color:C.dark} });
  s.addText('VS', { x:4.5, y:2.72, w:0.8, h:0.35, fontSize:10, bold:true, color:C.white, align:'center' });

  s.addShape(prs.ShapeType.rect, { x:5.5, y:1.45, w:4.05, h:3.3, fill:{color:'F0F3FF'}, line:{color:'5B8AF5',transparency:60}, rectRadius:0.1 });
  s.addText('✅  FeedbackLoop (Proposed)', { x:5.6, y:1.6, w:3.8, h:0.3, fontSize:12, bold:true, color:'2952CC' });
  const prop = ['Real-time NPS + verbatim ingestion after every visit','NLP auto-tags verbatims into 12 service themes','Anomaly detection: NPS drop >0.5pt triggers instant alert','Root cause AI identifies top 3 drivers within 24h','Auto-generates prioritised backlog with Jira integration','Fix cycle: hours from alert to ticket creation'];
  prop.forEach((t,i) => s.addText('· '+t, { x:5.65, y:2.02+i*0.36, w:3.8, h:0.3, fontSize:10, color:'1a1a1a' }));

  s.addShape(prs.ShapeType.rect, { x:0.5, y:4.85, w:9, h:0.55, fill:{color:C.dark}, line:{color:C.dark}, rectRadius:0.08 });
  s.addText('Key insight: The Al Barsha NPS drop (4.2→2.8) had three root causes: outdated KYC template, staff shortfall, and fee communication gap. All three were detectable in the verbatim data. FeedbackLoop would have flagged this in 24 hours, not 4 weeks.', {
    x:0.6, y:4.92, w:8.8, h:0.4, fontSize:10, color:C.blue, align:'center',
  });
}

// ─────────────────────────────────────────────
// SLIDE 4 — THE MECHANIC
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('THE MECHANIC', { x:0.5, y:0.35, w:9, h:0.25, fontSize:9, bold:true, color:C.blue, charSpacing:2, align:'left' });
  s.addText('How FeedbackLoop Works: Collect → Tag → Detect → Root-Cause → Backlog', { x:0.5, y:0.68, w:9, h:0.55, fontSize:19, bold:true, color:C.white, align:'left' });

  const steps = [
    { n:'01', title:'Auto Survey\nTrigger', body:'15 min post-visit, automated NPS survey sent via ENBD X app or SMS. Response linked to branch, service type, and token number.' },
    { n:'02', title:'NLP Tag\nClassification', body:'Verbatim responses tagged across 12 themes (wait time, staff helpfulness, document clarity, fee transparency, etc.) in real time.' },
    { n:'03', title:'Anomaly\nDetection', body:'Sliding 7-day NPS window per branch × service type. Alert fires when branch drops >0.5pt below 7-day baseline. Email + Slack notify.' },
    { n:'04', title:'Root Cause\nAI Analysis', body:'AI correlates NPS drop with verbatim tag spikes, staffing data, and ops events. Generates 3 ranked root causes within 24h of alert.' },
    { n:'05', title:'Backlog\nGeneration', body:'Each root cause maps to a backlog item with impact score. RM or Branch Manager creates Jira ticket with one click. SLA auto-set.' },
  ];
  steps.forEach((st, i) => {
    const x = 0.3 + i * 1.88;
    s.addShape(prs.ShapeType.rect, { x, y:1.45, w:1.72, h:3.5, fill:{color:C.card}, line:{color:'FFFFFF',transparency:90}, rectRadius:0.1 });
    s.addShape(prs.ShapeType.ellipse, { x:x+0.56, y:1.55, w:0.6, h:0.6, fill:{color:C.blue}, line:{color:C.blue} });
    s.addText(st.n, { x:x+0.56, y:1.57, w:0.6, h:0.4, fontSize:13, bold:true, color:C.white, align:'center' });
    s.addText(st.title, { x, y:2.3, w:1.72, h:0.55, fontSize:10, bold:true, color:C.white, align:'center' });
    s.addText(st.body, { x:x+0.1, y:2.95, w:1.52, h:1.8, fontSize:8.5, color:C.muted, align:'left' });
    if (i < 4) s.addText('→', { x:x+1.72, y:2.15, w:0.16, h:0.3, fontSize:11, color:C.blue, align:'center' });
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.1, w:9, h:0.3, fill:{color:C.card}, line:{color:'FFFFFF',transparency:95}, rectRadius:0.06 });
  s.addText('Proof: Same fragmented→unified ops intelligence architecture as PhonePe ops layer — 6 siloed workflows → unified real-time layer (TAT 2 days → 4 hours)', {
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
    { n:'P1', title:'NPS Overview\nDashboard', desc:'Network NPS 67 with branch ranking. Al Barsha flagged red. Category breakdown by service type.', ascii:'NPS Network: 67 ↑+4pt\n🟢 JBR       82\n🟢 DIFC      78\n🔴 Al Barsha 41 ⚠' },
    { n:'P2', title:'NLP Tag\nAnalytics', desc:'Top themes ranked by volume + trend. Alert: "Unclear loan docs" up 44%. Verbatim explorer with filters.', ascii:'✅ Helpful staff    487 ↑12%\n✅ Fast resolution  312 ↑7%\n❌ Long wait        274 ↑31%\n❌ Unclear loan docs 118 ↑44%' },
    { n:'P3', title:'Anomaly\nAlert', desc:'Auto-detected NPS drop: Al Barsha Home Loan 4.2→2.8. AI root cause analysis: 3 issues ranked by volume.', ascii:'🚨 ALERT: Al Barsha\nHome Loan: 4.2 → 2.8\n\n#1 KYC docs outdated (57%)\n#2 Staffing shortfall (29%)' },
    { n:'P4', title:'Backlog\nFeed', desc:'3 Critical + 7 High items auto-generated. Jira integration. Impact score ranking. One-click ticket creation.', ascii:'🔴 ENBD-1847: KYC docs\n🔴 Fee schedule update\n🟡 ENBD-1849: Staffing\n🔵 + Create Jira ×3' },
  ];
  screens.forEach((sc, i) => {
    const x = 0.3 + i * 2.35;
    s.addShape(prs.ShapeType.rect, { x, y:1.3, w:2.2, h:4.0, fill:{color:'FFFFFF'}, line:{color:C.dark,transparency:85}, rectRadius:0.1, shadow:makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y:1.3, w:2.2, h:0.38, fill:{color:C.dark}, line:{color:C.dark}, rectRadius:0.1 });
    s.addText(sc.n + ' — ' + sc.title.replace('\n',' '), { x:x+0.08, y:1.34, w:2.04, h:0.3, fontSize:9, bold:true, color:C.blue, align:'left' });
    s.addShape(prs.ShapeType.rect, { x:x+0.12, y:1.82, w:1.96, h:1.6, fill:{color:'0A1224'}, line:{color:'0A1224'}, rectRadius:0.06 });
    s.addText(sc.ascii, { x:x+0.14, y:1.86, w:1.92, h:1.52, fontSize:7, color:C.blue, fontFace:'Courier New', align:'left' });
    s.addText(sc.desc, { x:x+0.1, y:3.55, w:2.0, h:1.0, fontSize:8.5, color:'333333', align:'left' });
  });
}

// ─────────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('IMPACT & ROI', { x:0.5, y:0.35, w:9, h:0.25, fontSize:9, bold:true, color:C.blue, charSpacing:2, align:'left' });
  s.addText('Projected Impact — Closing the Feedback-to-Fix Gap', { x:0.5, y:0.68, w:9, h:0.5, fontSize:20, bold:true, color:C.white, align:'left' });

  s.addShape(prs.ShapeType.rect, { x:0.4, y:1.35, w:4.3, h:3.7, fill:{color:C.card}, line:{color:'FFFFFF',transparency:92}, rectRadius:0.1 });
  s.addText('SERVICE QUALITY IMPACT', { x:0.55, y:1.5, w:4.0, h:0.25, fontSize:9, bold:true, color:C.blue, charSpacing:1 });
  const left = [
    { v:'24h', l:'Anomaly detection latency', s:'vs 4 weeks for Excel-based monthly NPS review today' },
    { v:'+18pt', l:'NPS uplift from fast fixes', s:'Al Barsha type incidents resolved before they escalate to VP level' },
    { v:'−82%', l:'Fix cycle time', s:'3–4 weeks → 48h from verbatim to Jira ticket to resolution plan' },
    { v:'100%', l:'Verbatim-to-backlog coverage', s:'Every negative theme auto-generates a candidate backlog item' },
  ];
  left.forEach((m,i) => {
    s.addText(m.v, { x:0.55, y:1.92+i*0.76, w:4.0, h:0.38, fontSize:24, bold:true, color:C.blue });
    s.addText(m.l, { x:0.55, y:2.22+i*0.76, w:4.0, h:0.22, fontSize:11, bold:true, color:C.white });
    s.addText(m.s, { x:0.55, y:2.42+i*0.76, w:4.0, h:0.2, fontSize:9, color:C.muted });
  });

  s.addShape(prs.ShapeType.rect, { x:5.3, y:1.35, w:4.3, h:3.7, fill:{color:C.card}, line:{color:'FFFFFF',transparency:92}, rectRadius:0.1 });
  s.addText('EMIRATES NBD BUSINESS ROI', { x:5.45, y:1.5, w:4.0, h:0.25, fontSize:9, bold:true, color:C.gold, charSpacing:1 });
  const right = [
    { v:'AED 2.8M', l:'Prevented revenue loss per incident', s:'Each Al Barsha-type NPS crisis affects ~3,000 customers and attrition risk' },
    { v:'34%', l:'Reduction in repeat complaints', s:'Root cause resolution vs symptomatic band-aids — PhonePe ops proof point' },
    { v:'4 hours', l:'Sprint-ready backlog from incident', s:'Virtusa Agile cadence: FeedbackLoop alert → sprint ticket in one morning' },
    { v:'3×', l:'ROI on deployment cost', s:'One prevented VP escalation covers 6 months of platform running cost' },
  ];
  right.forEach((m,i) => {
    s.addText(m.v, { x:5.45, y:1.92+i*0.76, w:4.0, h:0.38, fontSize:24, bold:true, color:C.gold });
    s.addText(m.l, { x:5.45, y:2.22+i*0.76, w:4.0, h:0.22, fontSize:11, bold:true, color:C.white });
    s.addText(m.s, { x:5.45, y:2.42+i*0.76, w:4.0, h:0.2, fontSize:9, color:C.muted });
  });

  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.1, w:9, h:0.32, fill:{color:C.card}, line:{color:'FFFFFF',transparency:95}, rectRadius:0.06 });
  s.addText('The Al Barsha scenario in this prototype was undetected for 4 weeks. With FeedbackLoop, it would have fired an alert in 24 hours — and three Jira tickets in 4 hours.', {
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
  s.addText("I Replaced 6 Fragmented Workflows With One Unified Layer. Same Pattern Here.", { x:0.5, y:0.68, w:9, h:0.5, fontSize:19, bold:true, color:C.dark, align:'left' });

  s.addShape(prs.ShapeType.rect, { x:0.4, y:1.35, w:4.3, h:3.6, fill:{color:C.dark}, line:{color:C.dark}, rectRadius:0.1 });
  s.addText('WHAT I SHIPPED AT PHONEPE', { x:0.55, y:1.5, w:4.0, h:0.25, fontSize:9, bold:true, color:C.blue, charSpacing:1 });
  const pp = [
    'Replaced 6 siloed analyst workflows with unified ops intelligence layer',
    'Each workflow had its own data source, process, and escalation path',
    'Built KPI instrumentation BEFORE the product — same approach here',
    'Result: TAT 2 days → 4 hours, −68% escalation tickets within 60 days',
    'Designed real-time breach alerting + ops console for live monitoring',
    'Cross-functional delivery: Tech, DS, Ops, Compliance, Finance stakeholders',
  ];
  pp.forEach((t,i) => s.addText('· '+t, { x:0.55, y:1.9+i*0.44, w:4.0, h:0.36, fontSize:9.5, color:C.white, align:'left' }));

  s.addShape(prs.ShapeType.rect, { x:5.3, y:1.35, w:4.3, h:3.6, fill:{color:'EAF0FF'}, line:{color:C.blue,transparency:70}, rectRadius:0.1 });
  s.addText('HOW IT MAPS TO FEEDBACKLOOP', { x:5.45, y:1.5, w:4.0, h:0.25, fontSize:9, bold:true, color:C.blue, charSpacing:1 });
  const jd = [
    '→ 6 siloed feedback systems → one FeedbackLoop intelligence layer',
    '→ Each system (NPS, verbatim, ops, HR) feeds into one unified model',
    '→ KPIs set before build: detection latency, fix cycle, backlog coverage',
    '→ 4 weeks → 24h detection; 3 weeks → 48h fix cycle (same improvement class)',
    '→ Anomaly alert = ops breach alert (same threshold-trigger architecture)',
    '→ Virtusa: I own PM delivery + stakeholder co-design with ENBD Quality team',
  ];
  jd.forEach((t,i) => s.addText(t, { x:5.45, y:1.9+i*0.44, w:4.0, h:0.36, fontSize:9.5, color:'1a1a1a', align:'left' }));

  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.07, w:9, h:0.38, fill:{color:C.dark}, line:{color:C.dark}, rectRadius:0.08 });
  s.addText('"At PhonePe I unified 6 fragmented workflows into one intelligence layer. FeedbackLoop does the same for ENBD\'s disconnected NPS, verbatim, and ops data silos."', {
    x:0.6, y:5.12, w:8.8, h:0.28, fontSize:9.5, italic:true, color:C.blue, align:'center',
  });
}

// ─────────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('ROLLOUT PLAN', { x:0.5, y:0.35, w:9, h:0.25, fontSize:9, bold:true, color:C.blue, charSpacing:2, align:'left' });
  s.addText('3-Phase Delivery — Virtusa Agile Model', { x:0.5, y:0.68, w:9, h:0.45, fontSize:20, bold:true, color:C.white, align:'left' });

  const phases = [
    {
      n:'Phase 1', sub:'Month 1–2',
      title:'Pilot — 5 Branches + NLP',
      items:['Wire NPS survey trigger to ENBD X app (post-visit automation)', 'Deploy NLP tagging engine across 12 service themes', 'Build anomaly detection: 7-day sliding window per branch × service', 'Pilot: DIFC, Downtown, JBR, Al Barsha, Deira — measure detection speed'],
    },
    {
      n:'Phase 2', sub:'Month 3–4',
      title:'Root Cause AI + Jira',
      items:['Launch AI root cause correlation (verbatim + ops + staffing data)', 'Jira integration: backlog auto-population with 1-click ticket creation', 'Connect FeedbackLoop to BranchIQ: queue wait = NPS predictor signal', 'Branch Manager dashboard live — weekly sprint review with ENBD Quality'],
    },
    {
      n:'Phase 3', sub:'Month 5–6',
      title:'Network-Wide + Closed Loop',
      items:['Full network rollout — all Emirates NBD UAE branches', 'Closed-loop tracking: ticket created → fix deployed → NPS re-measured', 'AdvisorPulse integration: poor RM visit NPS → RM coaching queue', 'Target: −82% fix cycle, +18pt NPS uplift, 24h anomaly detection'],
    },
  ];
  phases.forEach((p, i) => {
    const x = 0.3 + i * 3.12;
    s.addShape(prs.ShapeType.rect, { x, y:1.3, w:2.95, h:3.45, fill:{color:C.card}, line:{color:'FFFFFF',transparency:90}, rectRadius:0.1 });
    s.addShape(prs.ShapeType.rect, { x, y:1.3, w:2.95, h:0.55, fill:{color:i===0?C.blue:i===1?C.teal:C.gold}, line:{color:'FFFFFF',transparency:100}, rectRadius:0.1 });
    s.addText(p.n + ' · ' + p.sub, { x:x+0.1, y:1.35, w:2.75, h:0.22, fontSize:9, bold:true, color:i===0?C.white:C.dark });
    s.addText(p.title, { x:x+0.1, y:1.55, w:2.75, h:0.22, fontSize:9.5, bold:true, color:i===0?C.white:C.dark });
    p.items.forEach((item, j) => s.addText('· '+item, { x:x+0.12, y:2.0+j*0.56, w:2.72, h:0.5, fontSize:9, color:C.muted, align:'left' }));
  });

  s.addShape(prs.ShapeType.rect, { x:0.4, y:4.87, w:9.2, h:0.55, fill:{color:C.card}, line:{color:C.blue,transparency:70}, rectRadius:0.08 });
  s.addText('What I need:  NPS survey data API  ·  Branch ops data (queue, staffing) for correlation  ·  1 NLP engineer  ·  Jira API access  ·  ENBD Branch Quality team as design partners  ·  Virtusa Agile sprint cadence', {
    x:0.55, y:4.93, w:9.0, h:0.4, fontSize:9, color:C.blue, align:'center',
  });
  s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  Planning UAE relocation Q2 2026', {
    x:0.5, y:5.38, w:9, h:0.2, fontSize:8.5, color:C.muted, align:'center',
  });
}

prs.writeFile({ fileName: 'assets/enbd-feedbackloop-deck.pptx' })
  .then(() => console.log('✅  enbd-feedbackloop-deck.pptx saved'))
  .catch(e => console.error(e));
