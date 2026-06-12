const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0A1A2E',
  hero:   '0F2540',
  amber:  'F59E0B',
  teal:   '00BCD4',
  gold:   'F5A623',
  white:  'FFFFFF',
  lgray:  'E8EEF4',
  muted:  '8B9BB4',
  green:  '10B981',
  red:    'EF4444',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.14 });

// ══════════════════════════════════════════
// SLIDE 1 — COVER
// ══════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.dark };

  for(let i = 0; i < 5; i++) {
    slide.addShape(pptx.ShapeType.line, { x:7.8+(i*0.25), y:0, w:0, h:7.5, line:{color: C.amber, width:0.4, transparency:82} });
  }
  slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:10, h:0.08, fill:{color: C.amber} });

  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:0.4, w:0.55, h:0.55, fill:{color: C.amber}, line:{color: C.amber}, shadow: makeShadow() });
  slide.addText('TG', { x:0.5, y:0.4, w:0.55, h:0.55, fontSize:14, bold:true, color:C.dark, align:'center', valign:'middle' });
  slide.addText('TRANSGUARD GROUP', { x:1.15, y:0.48, w:3, h:0.22, fontSize:8, bold:true, color:C.amber, charSpacing:3 });
  slide.addText('Cash Services Division', { x:1.15, y:0.68, w:3, h:0.18, fontSize:8, color:C.muted });

  slide.addText('ClientPulse', { x:0.5, y:1.25, w:9, h:1.4, fontSize:68, bold:true, color:C.white, fontFace:'Calibri' });
  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:2.6, w:3.2, h:0.05, fill:{color: C.amber} });
  slide.addText('Account Health & Retention Intelligence', { x:0.5, y:2.75, w:8, h:0.45, fontSize:18, color:C.amber, bold:true });
  slide.addText('Proactive AM Playbooks · 60-Day Renewal Foresight · ACV Protection Engine', { x:0.5, y:3.25, w:8, h:0.3, fontSize:11, color:C.muted });

  slide.addText('Ajay Avaghade · Product Manager · Applied Case Study', { x:0.5, y:6.5, w:6, h:0.28, fontSize:10, color:C.muted });
  slide.addText('ajay-avaghade.github.io/Portfolio', { x:0.5, y:6.78, w:5, h:0.22, fontSize:9, color:C.amber });

  slide.addShape(pptx.ShapeType.rect, { x:7.2, y:5.4, w:2.6, h:1.8, fill:{color: C.hero}, line:{color: C.amber, width:1}, shadow: makeShadow() });
  slide.addText('AED 7.8M', { x:7.2, y:5.55, w:2.6, h:0.6, fontSize:26, bold:true, color:C.amber, align:'center', fontFace:'Courier New' });
  slide.addText('ACV at renewal risk', { x:7.2, y:6.1, w:2.6, h:0.2, fontSize:8, color:C.muted, align:'center' });
  slide.addText('in the next 60 days alone', { x:7.2, y:6.3, w:2.6, h:0.2, fontSize:8, color:C.gold, align:'center', bold:true });
}

// ══════════════════════════════════════════
// SLIDE 2 — THE PROBLEM
// ══════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.dark };

  slide.addText('THE PROBLEM', { x:0.5, y:0.3, w:9, h:0.25, fontSize:9, bold:true, color:C.amber, charSpacing:4 });
  slide.addText('Transguard loses enterprise clients at renewal because account health signals arrive too late', { x:0.5, y:0.6, w:9, h:0.9, fontSize:26, bold:true, color:C.white });

  const stats = [
    { n:'21 days', l:'Average lag between health signal and AM awareness', s:'By the time a client shows dissatisfaction, the renewal conversation is already at risk' },
    { n:'AED 7.8M', l:'Enterprise ACV at renewal risk in any 60-day window', s:'3 accounts at HIGH risk. No systematic early-warning triggers in place today' },
    { n:'14 open', l:'Support tickets unresolved at Emirates NBD renewal', s:'Clients measure relationship health by ticket resolution speed — not by what AMs believe' },
  ];
  stats.forEach((st, i) => {
    const x = 0.5 + i * 3.15;
    slide.addShape(pptx.ShapeType.rect, { x, y:1.8, w:2.9, h:2.8, fill:{color: C.hero}, line:{color: C.amber, width:0.5, transparency:60}, shadow: makeShadow() });
    slide.addShape(pptx.ShapeType.rect, { x, y:1.8, w:2.9, h:0.04, fill:{color: C.amber} });
    slide.addText(st.n, { x, y:1.9, w:2.9, h:0.8, fontSize:32, bold:true, color:C.amber, align:'center', fontFace:'Courier New' });
    slide.addText(st.l, { x:x+0.1, y:2.75, w:2.7, h:0.32, fontSize:10, bold:true, color:C.white, align:'center' });
    slide.addText(st.s, { x:x+0.1, y:3.1, w:2.7, h:1.3, fontSize:9, color:C.muted, align:'center' });
  });

  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:4.9, w:9, h:0.9, fill:{color:'111827'}, line:{color: C.gold, width:1} });
  slide.addText('Root cause: Account management is relationship-driven and reactive. There is no system that scores client health, trends it over time, and fires a proactive intervention before renewal.', { x:0.65, y:4.95, w:8.7, h:0.8, fontSize:10, color:C.gold, valign:'middle' });
}

// ══════════════════════════════════════════
// SLIDE 3 — THE INSIGHT
// ══════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.lgray };

  slide.addText('THE INSIGHT', { x:0.5, y:0.3, w:9, h:0.25, fontSize:9, bold:true, color:C.dark, charSpacing:4 });
  slide.addText('Replace relationship intuition with a real-time client health score and automated intervention playbooks', { x:0.5, y:0.6, w:9, h:0.8, fontSize:24, bold:true, color:C.dark });

  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:1.65, w:4.2, h:4.2, fill:{color:'FEF2F2'}, line:{color:'FECACA', width:1} });
  slide.addText('TODAY (WITHOUT CLIENTPULSE)', { x:0.6, y:1.75, w:4, h:0.25, fontSize:8, bold:true, color:'991B1B', charSpacing:2 });
  const problems = [
    '❌ AM checks in with client based on "feel"',
    '❌ Health signals (SLA dips, ticket spikes) invisible',
    '❌ Churn risk discovered at contract renewal meeting',
    '❌ No consistent playbook — every AM improvises',
    '❌ ACV at risk only visible to Finance, not AMs',
    '❌ Upsell conversations driven by quota, not data',
  ];
  problems.forEach((p, i) => {
    slide.addText(p, { x:0.65, y:2.2+(i*0.47), w:3.9, h:0.4, fontSize:10, color:'7F1D1D' });
  });

  slide.addShape(pptx.ShapeType.ellipse, { x:4.55, y:3.3, w:0.9, h:0.9, fill:{color:C.dark}, line:{color:C.amber, width:1.5} });
  slide.addText('VS', { x:4.55, y:3.3, w:0.9, h:0.9, fontSize:14, bold:true, color:C.white, align:'center', valign:'middle' });

  slide.addShape(pptx.ShapeType.rect, { x:5.3, y:1.65, w:4.2, h:4.2, fill:{color:'FFFBEB'}, line:{color:'FDE68A', width:1} });
  slide.addText('WITH CLIENTPULSE', { x:5.4, y:1.75, w:4, h:0.25, fontSize:8, bold:true, color:'92400E', charSpacing:2 });
  const solutions = [
    '✅ Daily health score per account (6 dimensions)',
    '✅ SLA dips, ticket spikes flagged within 24 hours',
    '✅ 60-day renewal horizon with risk banding',
    '✅ 5-step playbook auto-triggered for RED accounts',
    '✅ ACV at risk surfaced to AM and their director',
    '✅ Upsell triggers tied to health score improvement',
  ];
  solutions.forEach((s, i) => {
    slide.addText(s, { x:5.45, y:2.2+(i*0.47), w:3.9, h:0.4, fontSize:10, color:'92400E' });
  });

  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:6.05, w:9, h:0.7, fill:{color:C.dark}, line:{color:C.amber, width:0.5} });
  slide.addText('All data required already exists in Transguard systems. ClientPulse is a synthesis and alerting layer on top — no new data collection, no new systems, fast time-to-build.', { x:0.65, y:6.1, w:8.7, h:0.6, fontSize:10, color:C.white, valign:'middle' });
}

// ══════════════════════════════════════════
// SLIDE 4 — THE MECHANIC
// ══════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.dark };

  slide.addText('THE MECHANIC', { x:0.5, y:0.3, w:9, h:0.25, fontSize:9, bold:true, color:C.amber, charSpacing:4 });
  slide.addText('Six-dimension health score with automated playbook and renewal pipeline', { x:0.5, y:0.6, w:9, h:0.7, fontSize:26, bold:true, color:C.white });

  const steps = [
    { n:'01', t:'Six-Dimension Daily Scoring', d:'Score computed daily from: SCDM uptime, collection SLA %, ERP reconciliation rate, open support tickets, QBR attendance, and AM contact frequency. Each dimension weighted and combined into 0-100 score.' },
    { n:'02', t:'RED / AMBER / GREEN Banding', d:'Score >= 75: GREEN (healthy, auto-renewal eligible). Score 50-74: AMBER (attention needed, QBR triggered). Score < 50: RED (at-risk, 5-step playbook auto-fired). AM notified within 24 hours of band change.' },
    { n:'03', t:'Auto-Playbook Execution', d:'RED account triggers a 5-step playbook: root cause diagnosis → executive outreach → SLA review meeting → remediation offer → renewal sign-off. Each step has an auto-generated message template and deadline.' },
    { n:'04', t:'60-Day Renewal Horizon', d:'Pipeline view showing all renewals in the next 60 days with risk score, ACV at stake, and recommended AM action. Director-level visibility without needing 24 status calls.' },
    { n:'05', t:'Upsell Trigger at Recovery', d:'Accounts that improve from RED to AMBER after playbook execution are flagged as upsell candidates. The health recovery conversation naturally opens the door to adding CashPulse or expanded services.' },
  ];

  steps.forEach((step, i) => {
    const y = 1.6 + i * 1.0;
    slide.addShape(pptx.ShapeType.rect, { x:0.4, y, w:0.55, h:0.55, fill:{color: C.amber}, line:{color: C.amber} });
    slide.addText(step.n, { x:0.4, y, w:0.55, h:0.55, fontSize:13, bold:true, color:C.dark, align:'center', valign:'middle' });
    if(i < steps.length - 1) {
      slide.addShape(pptx.ShapeType.line, { x:0.665, y:y+0.55, w:0, h:0.45, line:{color: C.amber, width:1, dashType:'dash', transparency:50} });
    }
    slide.addText(step.t, { x:1.1, y:y+0.04, w:8.3, h:0.26, fontSize:12, bold:true, color:C.white });
    slide.addText(step.d, { x:1.1, y:y+0.3, w:8.3, h:0.45, fontSize:9, color:C.muted });
  });

  slide.addShape(pptx.ShapeType.rect, { x:0.4, y:6.65, w:9.2, h:0.52, fill:{color:'111827'}, line:{color: C.gold, width:0.8} });
  slide.addText('POC scope: 20 accounts, 4 AMs, 90 days. Measure: renewal rate delta vs control group, escalation reduction, and upsell conversion. Break-even on build cost from a single saved renewal.', { x:0.55, y:6.68, w:9, h:0.45, fontSize:9, color:C.gold, valign:'middle' });
}

// ══════════════════════════════════════════
// SLIDE 5 — THE PRODUCT
// ══════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.lgray };

  slide.addText('THE PRODUCT', { x:0.5, y:0.3, w:9, h:0.25, fontSize:9, bold:true, color:C.dark, charSpacing:4 });
  slide.addText('Four screens · AM platform · Director-level pipeline view', { x:0.5, y:0.6, w:9, h:0.6, fontSize:26, bold:true, color:C.dark });

  const screens = [
    { n:'01', t:'Account Portfolio', d:'Grid of all accounts with health scores in RED/AMBER/GREEN. Summary bar: 24 total, 3 at risk, 7 need attention, 14 healthy. ACV at stake displayed per account card.', mock:['  24 Total   3 At Risk  7 Attention  14 Healthy','Emirates NBD  [38 RED]  28d renewal  AED 4.2M','ENOC Group    [61 AMB]  55d renewal  AED 2.8M','Carrefour UAE [87 GRN]  142d renewal AED 1.9M'] },
    { n:'02', t:'Health Score Detail', d:'Per-account drill-down: 6-dimension score breakdown, 6-month trend chart, recent touchpoints log. Root cause surfaced automatically from the lowest-scoring dimensions.', mock:['Emirates NBD · Score: 38 ↓17pts · AT RISK','SCDM Uptime:    91.2%  ██████████░','Collection SLA: 78.4%  ████████░░░','Support Tickets: 14 open  ↑ 8 from last month'] },
    { n:'03', t:'Intervention Playbook', d:'5-step playbook auto-triggered for RED accounts. Step 2 active: auto-generated executive outreach email ready to send in one click. Step completion tracked with due dates.', mock:['✓ Step 1: Root Cause Diagnosis  [Done]','▶ Step 2: Executive Outreach   [Due Today]','○ Step 3: SLA Review Meeting   [15 Jun]','○ Step 4: Remediation Offer    [22 Jun]'] },
    { n:'04', t:'Renewal Risk Horizon', d:'60-day pipeline view of all upcoming renewals. Risk score, ACV, assigned AM, and recommended action per renewal. Director-level visibility across all accounts.', mock:['19d  Waitrose UAE   [HIGH]  AED 0.8M  Playbook','28d  Emirates NBD   [HIGH]  AED 4.2M  Playbook','41d  Carrefour DXB  [MED]   AED 0.6M  QBR Sched','55d  ENOC Group     [MED]   AED 2.8M  Upsell Rev'] },
  ];

  screens.forEach((sc, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.5 + col * 4.8;
    const y = 1.4 + row * 2.9;
    slide.addShape(pptx.ShapeType.rect, { x, y, w:4.5, h:2.7, fill:{color:C.white}, line:{color:'CBD5E1', width:0.5}, shadow: makeShadow() });
    slide.addShape(pptx.ShapeType.rect, { x, y, w:4.5, h:0.04, fill:{color: C.amber} });
    slide.addShape(pptx.ShapeType.rect, { x, y:y+0.04, w:4.5, h:0.38, fill:{color:C.dark} });
    slide.addText(`Screen ${sc.n} — ${sc.t}`, { x:x+0.12, y:y+0.06, w:4.3, h:0.3, fontSize:10, bold:true, color:C.amber });
    slide.addText(sc.d, { x:x+0.12, y:y+0.5, w:4.28, h:0.7, fontSize:8.5, color:'334155' });
    slide.addShape(pptx.ShapeType.rect, { x:x+0.12, y:y+1.28, w:4.28, h:1.26, fill:{color:'F8FAFC'}, line:{color:'E2E8F0', width:0.5} });
    sc.mock.forEach((line, li) => {
      slide.addText(line, { x:x+0.18, y:y+1.34+(li*0.28), w:4.18, h:0.26, fontSize:7.5, color:'475569', fontFace:'Courier New' });
    });
  });
}

// ══════════════════════════════════════════
// SLIDE 6 — IMPACT & ROI
// ══════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.dark };

  slide.addText('IMPACT & ROI', { x:0.5, y:0.3, w:9, h:0.25, fontSize:9, bold:true, color:C.amber, charSpacing:4 });
  slide.addText('Retention revenue protected + upsell conversion unlocked', { x:0.5, y:0.6, w:9, h:0.7, fontSize:26, bold:true, color:C.white });

  const leftMetrics = [
    { v:'22%', l:'Improvement in enterprise client retention', s:'Proactive playbooks catch at-risk accounts 60 days before renewal — enough time to intervene and save the contract.' },
    { v:'35%', l:'Fewer surprise churn events at renewal', s:'No more AED 4.2M contracts lost because the AM did not know the score was in freefall for 90 days.' },
    { v:'-60%', l:'Reduction in reactive escalation calls', s:'Clients whose issues are caught early through ClientPulse scoring do not escalate to Director level.' },
    { v:'18%', l:'Increase in upsell conversion rate', s:'Health-recovering accounts are the warmest upsell targets. ClientPulse identifies the exact moment to introduce CashPulse.' },
  ];
  const rightMetrics = [
    { v:'AED 4.2M', l:'Saved from a single Emirates NBD renewal', s:'The entire ClientPulse build cost is recovered from preventing one large enterprise account from churning.' },
    { v:'AED 85K', l:'Cost of proactive service credit (2% of ACV)', s:'Versus losing AED 4.2M contract entirely. The remediation economics are fundamentally asymmetric.' },
    { v:'20 days', l:'Build time to POC-ready', s:'No new data collection. Only integration with existing Transguard operational data and a scoring algorithm.' },
    { v:'1 quarter', l:'Payback period from first saved renewal', s:'Break-even on product cost from the first high-risk account retained. Every subsequent save is pure margin.' },
  ];

  slide.addText('CLIENT / RETENTION IMPACT', { x:0.5, y:1.5, w:4.5, h:0.25, fontSize:8, bold:true, color:C.amber, charSpacing:3 });
  slide.addText('TRANSGUARD ROI', { x:5.2, y:1.5, w:4.5, h:0.25, fontSize:8, bold:true, color:C.gold, charSpacing:3 });

  leftMetrics.forEach((m, i) => {
    const y = 1.85 + i * 1.15;
    slide.addShape(pptx.ShapeType.rect, { x:0.5, y, w:4.4, h:1.0, fill:{color: C.hero}, line:{color: C.amber, width:0.5, transparency:60} });
    slide.addText(m.v, { x:0.65, y:y+0.08, w:1.4, h:0.5, fontSize:26, bold:true, color:C.amber, fontFace:'Courier New' });
    slide.addText(m.l, { x:2.1, y:y+0.1, w:2.7, h:0.3, fontSize:9, bold:true, color:C.white });
    slide.addText(m.s, { x:2.1, y:y+0.4, w:2.7, h:0.45, fontSize:8, color:C.muted });
  });

  rightMetrics.forEach((m, i) => {
    const y = 1.85 + i * 1.15;
    slide.addShape(pptx.ShapeType.rect, { x:5.2, y, w:4.4, h:1.0, fill:{color: C.hero}, line:{color: C.gold, width:0.5, transparency:60} });
    slide.addText(m.v, { x:5.35, y:y+0.08, w:1.8, h:0.5, fontSize:22, bold:true, color:C.gold, fontFace:'Courier New' });
    slide.addText(m.l, { x:7.2, y:y+0.1, w:2.3, h:0.3, fontSize:9, bold:true, color:C.white });
    slide.addText(m.s, { x:7.2, y:y+0.4, w:2.3, h:0.45, fontSize:8, color:C.muted });
  });

  slide.addShape(pptx.ShapeType.line, { x:4.97, y:1.5, w:0, h:5.2, line:{color: C.amber, width:0.5, transparency:60} });
}

// ══════════════════════════════════════════
// SLIDE 7 — PROOF OF WORK
// ══════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.lgray };

  slide.addText('PROOF OF WORK', { x:0.5, y:0.3, w:9, h:0.25, fontSize:9, bold:true, color:C.dark, charSpacing:4 });
  slide.addText('What I built at PhonePe maps exactly to what ClientPulse requires', { x:0.5, y:0.6, w:9, h:0.7, fontSize:24, bold:true, color:C.dark });

  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:1.45, w:4.4, h:5.3, fill:{color:C.dark}, line:{color:C.amber, width:0.5} });
  slide.addText('WHAT I BUILT AT PHONEPE', { x:0.6, y:1.55, w:4.2, h:0.25, fontSize:8, bold:true, color:C.amber, charSpacing:2 });

  const phonepe = [
    { t:'Milestone Engine: Merchant Retention +23%', d:'Built gamified milestone system that shaped merchant behaviour over time via data-triggered nudges. Reduced merchant escalations by 85%. ClientPulse uses the same mechanic: data-triggered intervention at the right moment.' },
    { t:'Lifecycle Retention: -60% Cart Abandonment', d:'Built context-aware, real-time triggers for Pincode quick commerce. Same architecture as ClientPulse health score triggers — different domain, identical system design principle.' },
    { t:'B2B Self-Serve: 5,000+ Merchants, -23% CAC', d:'Zero-to-one B2B platform launch. Merchants self-serve their own offer configuration with no AM dependency. ClientPulse reduces AM dependency for account health monitoring using the same self-serve philosophy.' },
    { t:'ML Personalization: 26% Engagement Uplift', d:'Deployed ML models for audience segmentation and timing optimization. ClientPulse health scoring uses the same weighting logic: dimensional scores are like engagement signals, combined into a single propensity metric.' },
  ];
  phonepe.forEach((p, i) => {
    slide.addShape(pptx.ShapeType.rect, { x:0.65, y:2.0+(i*1.15), w:4.1, h:1.0, fill:{color: C.hero}, line:{color: C.amber, width:0.3, transparency:60} });
    slide.addText(p.t, { x:0.78, y:2.08+(i*1.15), w:3.9, h:0.26, fontSize:10, bold:true, color:C.amber });
    slide.addText(p.d, { x:0.78, y:2.35+(i*1.15), w:3.9, h:0.55, fontSize:8.5, color:C.muted });
  });

  slide.addShape(pptx.ShapeType.rect, { x:5.1, y:1.45, w:4.4, h:5.3, fill:{color:C.white}, line:{color:'CBD5E1', width:0.5} });
  slide.addText('HOW IT MAPS TO CLIENTPULSE', { x:5.2, y:1.55, w:4.2, h:0.25, fontSize:8, bold:true, color:C.dark, charSpacing:2 });

  const mapping = [
    { t:'Merchant Milestone = AM Intervention Playbook', d:'PhonePe milestones triggered specific merchant actions via data signals. ClientPulse playbooks trigger specific AM actions via health score signals. Same architecture — the domain shifts from consumer merchants to enterprise clients.' },
    { t:'Cart Abandonment Engine = Churn Prevention Engine', d:'Both systems identify a user at risk of abandoning (cart / contract) and fire a contextual intervention at the right moment. The intervention design logic is identical.' },
    { t:'B2B Self-Serve = AM Self-Serve Dashboard', d:'PhonePe removed the need for AMs to manually configure merchant offers. ClientPulse removes the need for Directors to call AMs for status updates. Same self-serve principle, applied to the AM layer.' },
    { t:'Audience Segmentation = Health Score Banding', d:'Propensity-to-transact model segments users into HIGH/MED/LOW. ClientPulse health score segments clients into GREEN/AMBER/RED. The algorithmic logic and response triggers are structurally identical.' },
  ];
  mapping.forEach((m, i) => {
    slide.addShape(pptx.ShapeType.rect, { x:5.25, y:2.0+(i*1.15), w:4.1, h:1.0, fill:{color:C.lgray}, line:{color:'CBD5E1', width:0.5} });
    slide.addText(m.t, { x:5.38, y:2.08+(i*1.15), w:3.9, h:0.26, fontSize:10, bold:true, color:C.dark });
    slide.addText(m.d, { x:5.38, y:2.35+(i*1.15), w:3.9, h:0.55, fontSize:8.5, color:'475569' });
  });

  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:6.95, w:9, h:0.75, fill:{color:C.dark}, line:{color: C.amber, width:0.5} });
  slide.addText('"I have already built and shipped every component of ClientPulse — just applied to a different business domain. The playbook design, the scoring architecture, and the AM tooling are all proven patterns from my work."', { x:0.65, y:7.0, w:8.7, h:0.65, fontSize:10, color:C.amber, valign:'middle', italic:true });
}

// ══════════════════════════════════════════
// SLIDE 8 — ROLLOUT PLAN
// ══════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.dark };

  slide.addText('ROLLOUT PLAN', { x:0.5, y:0.3, w:9, h:0.25, fontSize:9, bold:true, color:C.amber, charSpacing:4 });
  slide.addText('POC in 20 days · First renewal saved in 30 days · Full rollout in 90 days', { x:0.5, y:0.6, w:9, h:0.7, fontSize:24, bold:true, color:C.white });

  const phases = [
    { n:'Phase 01', t:'POC Launch', dur:'Days 1–30', items:['20 accounts, 4 AMs enrolled','Daily health score computed from existing data','Playbook fired for Emirates NBD + Waitrose','Success metric: renewal save rate + AM feedback NPS'], color: C.amber },
    { n:'Phase 02', t:'Full AM Rollout', dur:'Days 31–90', items:['All 24 enterprise accounts on ClientPulse','Director renewal horizon view live','Playbook library expanded (upsell, QBR, expansion)','Integration with Transguard CRM / Salesforce'], color: C.teal },
    { n:'Phase 03', t:'Scale & Product Value', dur:'Days 91–180', items:['Cross-sell CashPulse to ClientPulse renewals','Build client-facing health portal (shared view)','Upsell trigger system goes live (est. +18% conversion)','Target: AED 4.2M+ ACV retained in first 180 days'], color: C.gold },
  ];

  phases.forEach((ph, i) => {
    const x = 0.5 + i * 3.18;
    slide.addShape(pptx.ShapeType.rect, { x, y:1.5, w:3.0, h:4.5, fill:{color: C.hero}, line:{color: ph.color, width:1}, shadow: makeShadow() });
    slide.addShape(pptx.ShapeType.rect, { x, y:1.5, w:3.0, h:0.06, fill:{color: ph.color} });
    slide.addText(ph.n, { x:x+0.15, y:1.6, w:2.7, h:0.22, fontSize:8, bold:true, color:ph.color, charSpacing:2 });
    slide.addText(ph.t, { x:x+0.15, y:1.85, w:2.7, h:0.45, fontSize:18, bold:true, color:C.white });
    slide.addShape(pptx.ShapeType.rect, { x:x+0.15, y:2.35, w:1.4, h:0.24, fill:{color: ph.color}, line:{color: ph.color} });
    slide.addText(ph.dur, { x:x+0.15, y:2.35, w:1.4, h:0.24, fontSize:8, bold:true, color:C.dark, align:'center', valign:'middle' });
    ph.items.forEach((item, j) => {
      slide.addText(`• ${item}`, { x:x+0.15, y:2.78+(j*0.65), w:2.75, h:0.58, fontSize:9, color:C.muted });
    });
  });

  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:6.2, w:9, h:1.15, fill:{color: C.hero}, line:{color: C.amber, width:1} });
  slide.addText('What I need from Transguard to start:', { x:0.65, y:6.28, w:8.7, h:0.28, fontSize:10, bold:true, color:C.amber });
  slide.addText('Access to Transguard operational data (SCDM uptime, collection SLA, ticket system) · List of 20 accounts for POC · 2 weeks of AM time to define scoring weights · 90-day mandate with leadership visibility', { x:0.65, y:6.58, w:8.7, h:0.58, fontSize:9, color:C.muted });

  slide.addText('Ajay Avaghade · avaghadeajay009@gmail.com · ajay-avaghade.github.io/Portfolio', { x:0.5, y:7.38, w:9, h:0.25, fontSize:9, color:C.muted, align:'center' });
}

pptx.writeFile({ fileName: 'assets/transguard-clientpulse-deck.pptx' })
  .then(() => console.log('ClientPulse deck saved: assets/transguard-clientpulse-deck.pptx'))
  .catch(err => console.error(err));
