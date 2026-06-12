const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0A1A2E',
  hero:   '0F2540',
  teal:   '00BCD4',
  gold:   'F5A623',
  white:  'FFFFFF',
  lgray:  'E8EEF4',
  muted:  '8B9BB4',
  green:  '10B981',
  red:    'EF4444',
  amber:  'F59E0B',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.14 });

// ══════════════════════════════════════════
// SLIDE 1 — COVER
// ══════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.dark };

  // Diagonal accent lines
  for(let i = 0; i < 6; i++) {
    slide.addShape(pptx.ShapeType.line, { x:7.5+(i*0.22), y:0, w:0, h:7.5, line:{color: C.teal, width:0.4, transparency:85} });
  }

  // Teal top strip
  slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:10, h:0.08, fill:{color: C.teal} });

  // TG logo mark
  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:0.4, w:0.55, h:0.55, fill:{color: C.teal}, line:{color: C.teal}, shadow: makeShadow() });
  slide.addText('TG', { x:0.5, y:0.4, w:0.55, h:0.55, fontSize:14, bold:true, color:C.dark, align:'center', valign:'middle' });
  slide.addText('TRANSGUARD GROUP', { x:1.15, y:0.48, w:3, h:0.22, fontSize:8, bold:true, color:C.teal, charSpacing:3 });
  slide.addText('Cash Services Division', { x:1.15, y:0.68, w:3, h:0.18, fontSize:8, color:C.muted });

  // Main product name
  slide.addText('CashPulse', { x:0.5, y:1.25, w:9, h:1.4, fontSize:72, bold:true, color:C.white, fontFace:'Calibri' });
  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:2.6, w:3.2, h:0.05, fill:{color: C.teal} });
  slide.addText('Retailer Cash Intelligence Platform', { x:0.5, y:2.75, w:7, h:0.45, fontSize:18, color:C.teal, bold:true });
  slide.addText('Real-Time SCDM Visibility · ERP Auto-Reconciliation · Client Health Intelligence', { x:0.5, y:3.25, w:8, h:0.3, fontSize:11, color:C.muted });

  // Presenter
  slide.addText('Ajay Avaghade · Product Manager · Applied Case Study', { x:0.5, y:6.5, w:6, h:0.28, fontSize:10, color:C.muted });
  slide.addText('ajay-avaghade.github.io/Portfolio', { x:0.5, y:6.78, w:5, h:0.22, fontSize:9, color:C.teal });

  // Big stat bottom right
  slide.addShape(pptx.ShapeType.rect, { x:7.2, y:5.4, w:2.6, h:1.8, fill:{color: C.hero}, line:{color: C.teal, width:1}, shadow: makeShadow() });
  slide.addText('AED 12.4M', { x:7.2, y:5.55, w:2.6, h:0.6, fontSize:24, bold:true, color:C.teal, align:'center', fontFace:'Courier New' });
  slide.addText('daily cash processed', { x:7.2, y:6.1, w:2.6, h:0.2, fontSize:8, color:C.muted, align:'center' });
  slide.addText('zero visibility to clients today', { x:7.2, y:6.3, w:2.6, h:0.2, fontSize:8, color:C.gold, align:'center', bold:true });
}

// ══════════════════════════════════════════
// SLIDE 2 — THE PROBLEM
// ══════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.dark };

  slide.addText('THE PROBLEM', { x:0.5, y:0.3, w:9, h:0.25, fontSize:9, bold:true, color:C.teal, charSpacing:4 });
  slide.addText('Retailers using Transguard SCDMs operate with a 24-hour cash blind spot', { x:0.5, y:0.6, w:9, h:0.9, fontSize:28, bold:true, color:C.white });

  const stats = [
    { n:'24 hrs', l:'Manual reconciliation delay', s:'Nightly SAP batch runs surface exceptions next morning — too late to act' },
    { n:'AED 14M+', l:'Daily cash volume untracked in real-time', s:'Retailers cannot see their effective cash position until the following business day' },
    { n:'0%', l:'Client visibility into SCDM data today', s:'All operational data sits inside Transguard systems — none is shared back to clients in real-time' },
  ];
  stats.forEach((st, i) => {
    const x = 0.5 + i * 3.15;
    slide.addShape(pptx.ShapeType.rect, { x, y:1.8, w:2.9, h:2.8, fill:{color: C.hero}, line:{color: C.teal, width:0.5, transparency:60}, shadow: makeShadow() });
    slide.addShape(pptx.ShapeType.rect, { x, y:1.8, w:2.9, h:0.04, fill:{color: C.teal} });
    slide.addText(st.n, { x, y:1.9, w:2.9, h:0.8, fontSize:36, bold:true, color:C.teal, align:'center', fontFace:'Courier New' });
    slide.addText(st.l, { x:x+0.1, y:2.75, w:2.7, h:0.3, fontSize:10, bold:true, color:C.white, align:'center' });
    slide.addText(st.s, { x:x+0.1, y:3.1, w:2.7, h:1.3, fontSize:9, color:C.muted, align:'center' });
  });

  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:4.9, w:9, h:0.9, fill:{color:'111827'}, line:{color: C.gold, width:1} });
  slide.addText('Core insight: Transguard already has all this data. The product is a presentation layer on existing systems — not a new data collection effort.', { x:0.65, y:4.95, w:8.7, h:0.8, fontSize:10, color:C.gold, valign:'middle' });
}

// ══════════════════════════════════════════
// SLIDE 3 — THE INSIGHT
// ══════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.lgray };

  slide.addText('THE INSIGHT', { x:0.5, y:0.3, w:9, h:0.25, fontSize:9, bold:true, color:C.dark, charSpacing:4 });
  slide.addText('Transguard is sitting on a data gold mine — and giving none of it back to clients', { x:0.5, y:0.6, w:9, h:0.8, fontSize:26, bold:true, color:C.dark });

  // Left: Current state
  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:1.65, w:4.2, h:4.2, fill:{color:'FEF2F2'}, line:{color:'FECACA', width:1} });
  slide.addText('TODAY (WITHOUT CASHPULSE)', { x:0.6, y:1.75, w:4, h:0.25, fontSize:8, bold:true, color:'991B1B', charSpacing:2 });
  const problems = [
    '❌ Cash deposited → 24hr batch reconciliation',
    '❌ Fill levels checked only at physical collection',
    '❌ ERP exceptions discovered next morning',
    '❌ Client calls Transguard AM for status updates',
    '❌ No cross-store cash position visibility',
    '❌ Collection scheduling based on time, not fill velocity',
  ];
  problems.forEach((p, i) => {
    slide.addText(p, { x:0.65, y:2.2+(i*0.47), w:3.9, h:0.4, fontSize:10, color:'7F1D1D' });
  });

  // VS divider
  slide.addShape(pptx.ShapeType.ellipse, { x:4.55, y:3.3, w:0.9, h:0.9, fill:{color:C.dark}, line:{color:C.teal, width:1.5} });
  slide.addText('VS', { x:4.55, y:3.3, w:0.9, h:0.9, fontSize:14, bold:true, color:C.white, align:'center', valign:'middle' });

  // Right: CashPulse state
  slide.addShape(pptx.ShapeType.rect, { x:5.3, y:1.65, w:4.2, h:4.2, fill:{color:'F0FDF4'}, line:{color:'A7F3D0', width:1} });
  slide.addText('WITH CASHPULSE', { x:5.4, y:1.75, w:4, h:0.25, fontSize:8, bold:true, color:'065F46', charSpacing:2 });
  const solutions = [
    '✅ Real-time deposit credit confirmation (<5 min)',
    '✅ Live SCDM fill levels per machine per store',
    '✅ ERP exceptions flagged and resolved same session',
    '✅ Client self-serves all data — zero AM calls needed',
    '✅ Network-wide cash position dashboard',
    '✅ Predictive collection scheduling from fill velocity',
  ];
  solutions.forEach((s, i) => {
    slide.addText(s, { x:5.45, y:2.2+(i*0.47), w:3.9, h:0.4, fontSize:10, color:'065F46' });
  });

  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:6.05, w:9, h:0.7, fill:{color:C.dark}, line:{color:C.teal, width:0.5} });
  slide.addText('The product does not require new data infrastructure. It surfaces what already exists inside Transguard systems — making Transguard indispensable to each client.', { x:0.65, y:6.1, w:8.7, h:0.6, fontSize:10, color:C.white, valign:'middle' });
}

// ══════════════════════════════════════════
// SLIDE 4 — THE MECHANIC
// ══════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.dark };

  slide.addText('THE MECHANIC', { x:0.5, y:0.3, w:9, h:0.25, fontSize:9, bold:true, color:C.teal, charSpacing:4 });
  slide.addText('Five-layer data pipeline from SCDM deposit to client action', { x:0.5, y:0.6, w:9, h:0.7, fontSize:26, bold:true, color:C.white });

  const steps = [
    { n:'01', t:'SCDM Event Capture', d:'Every deposit event is captured in real-time from all 847 SCDMs across the UAE via Transguard secure network. Fill level, timestamp, and denomination data extracted.' },
    { n:'02', t:'Instant Credit Notification', d:'Within <5 minutes of deposit: client receives real-time push notification confirming amount, machine ID, and credit to their designated account. Zero wait time.' },
    { n:'03', t:'ERP Auto-Match Engine', d:'Deposit events are matched against client ERP entries (SAP, Oracle, Dynamics) via API. Matched: flagged green. Exception: flagged red with root cause + resolution path sent to client finance.' },
    { n:'04', t:'Predictive Collection Scheduling', d:'Fill velocity per machine extrapolated over rolling 7-day window. CashPulse predicts 90% capacity time and auto-schedules collection 4 hours in advance. Reduces rush call-outs.' },
    { n:'05', t:'Client Dashboard + AM Health View', d:'Client-facing: cash position, fill levels, ERP status. Transguard internal: client health scores, SLA metrics, collection schedule. Both views in one platform — two stakeholders, one source of truth.' },
  ];

  steps.forEach((step, i) => {
    const y = 1.6 + i * 1.0;
    slide.addShape(pptx.ShapeType.rect, { x:0.4, y, w:0.55, h:0.55, fill:{color: C.teal}, line:{color: C.teal} });
    slide.addText(step.n, { x:0.4, y, w:0.55, h:0.55, fontSize:13, bold:true, color:C.dark, align:'center', valign:'middle' });
    if(i < steps.length - 1) {
      slide.addShape(pptx.ShapeType.line, { x:0.665, y:y+0.55, w:0, h:0.45, line:{color: C.teal, width:1, dashType:'dash', transparency:50} });
    }
    slide.addText(step.t, { x:1.1, y:y+0.04, w:8.3, h:0.26, fontSize:12, bold:true, color:C.white });
    slide.addText(step.d, { x:1.1, y:y+0.3, w:8.3, h:0.45, fontSize:9, color:C.muted });
  });

  slide.addShape(pptx.ShapeType.rect, { x:0.4, y:6.65, w:9.2, h:0.52, fill:{color:'111827'}, line:{color: C.amber, width:0.8} });
  slide.addText('POC validation: Deploy with Carrefour UAE (34 SCDMs, SAP S/4HANA) + ENOC Group (118 SCDMs) over 90 days. Measure: reconciliation time delta, AM call reduction, client satisfaction score.', { x:0.55, y:6.68, w:9, h:0.45, fontSize:9, color:C.amber, valign:'middle' });
}

// ══════════════════════════════════════════
// SLIDE 5 — THE PRODUCT
// ══════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.lgray };

  slide.addText('THE PRODUCT', { x:0.5, y:0.3, w:9, h:0.25, fontSize:9, bold:true, color:C.dark, charSpacing:4 });
  slide.addText('Four screens · One platform · Two stakeholders', { x:0.5, y:0.6, w:9, h:0.6, fontSize:26, bold:true, color:C.dark });

  const screens = [
    { n:'01', t:'Network Dashboard', d:'All clients, all SCDMs, all health scores in one view. KPI strip: daily volume, avg fill, active alerts. Client table with health badge, fill mini-bar, ERP sync status.', mock:['[KPI STRIP: 847 SCDMs · AED 12.4M · 73% Fill · 3 Alerts]','Carrefour UAE   ●Healthy  87%  AED 2.8M  ✓ Live','ENOC Group      ●Attention 61% AED 5.1M  ✓ Live','Emirates NBD    ●At Risk  31%  AED 3.3M  ⚠ Delayed'] },
    { n:'02', t:'Store Deep-Dive', d:'Per-branch, per-machine fill levels, last deposit amounts, velocity data, next predicted collection window. ERP sync status per deposit in real-time.', mock:['SCDM-CF-001  ●Operational  92% ████████░', 'SCDM-CF-002  ⚠ High Fill   78% ███████░░','SCDM-CF-003  ●Operational  61% ██████░░░','DEP-2847  22,100 AED  ✓ SAP-8821 Matched'] },
    { n:'03', t:'ERP Reconciliation', d:'Auto-match engine runs every 5 minutes. 142 matched, 3 exceptions, 8 pending. Exception detail shows root cause and resolution path. Export to CSV for client finance.', mock:['✅ 142 Matched  ⚠ 3 Exceptions  ⏳ 8 Pending','DEP-2847  Carrefour  22,100  ✓ SAP-8821','DEP-2845  ENBD  187,500  ⚠ Exception ←','DEP-2839  Carrefour  31,750  ✓ SAP-8815'] },
    { n:'04', t:'Ops Command Center', d:'Transguard internal view: client health scores, live collection schedule, SLA performance meters, and active alert feed. AM dashboard with proactive intervention triggers.', mock:['Client Health: Carrefour 87 ●  ENOC 62 ●','Collection: 07:30 CF-Barsha ✓  10:00 ENOC →','SLA: CIT Timeliness 98.2%  SCDM Uptime 99.1%','⚠ ENBD DAFZA — fill critical · ⚠ ENOC SAP lag'] },
  ];

  screens.forEach((sc, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.5 + col * 4.8;
    const y = 1.4 + row * 2.9;
    slide.addShape(pptx.ShapeType.rect, { x, y, w:4.5, h:2.7, fill:{color:C.white}, line:{color:'CBD5E1', width:0.5}, shadow: makeShadow() });
    slide.addShape(pptx.ShapeType.rect, { x, y, w:4.5, h:0.04, fill:{color: C.teal} });
    slide.addShape(pptx.ShapeType.rect, { x, y:y+0.04, w:4.5, h:0.38, fill:{color:C.dark} });
    slide.addText(`Screen ${sc.n} — ${sc.t}`, { x:x+0.12, y:y+0.06, w:4.3, h:0.3, fontSize:10, bold:true, color:C.teal });
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

  slide.addText('IMPACT & ROI', { x:0.5, y:0.3, w:9, h:0.25, fontSize:9, bold:true, color:C.teal, charSpacing:4 });
  slide.addText('New revenue line + client retention flywheel', { x:0.5, y:0.6, w:9, h:0.7, fontSize:26, bold:true, color:C.white });

  const leftMetrics = [
    { v:'40%', l:'Reduction in manual reconciliation time', s:'24-hr batch → <5 min auto-match. Client finance teams reclaim 3+ hrs/day.' },
    { v:'28%', l:'Improvement in SCDM contract retention', s:'Real-time data creates daily dependency. Clients cannot leave a platform they use every morning.' },
    { v:'-60%', l:'Reduction in status-update AM calls', s:'Self-serve dashboard replaces all manual requests for fill levels, deposit confirmations, and SLA reports.' },
    { v:'94%', l:'ERP exception resolution same-session', s:'Exceptions surfaced within 5 minutes, not 24 hours. Client treasury closes books with confidence.' },
  ];
  const rightMetrics = [
    { v:'AED 3.2M', l:'New annual recurring revenue (Year 1)', s:'15 anchor clients × AED 213K dashboard license + API fee. Incremental to existing CIT/SCDM contracts.' },
    { v:'AED 12.4M', l:'Platform AUM under management', s:'Daily cash volume on the CashPulse network creates audit trail and insurance-grade reconciliation data.' },
    { v:'6x', l:'ROI on product development cost', s:'Build cost estimated AED 420K. Year 1 license revenue AED 3.2M. Payback in under 2 quarters.' },
    { v:'30 days', l:'Time-to-POC from green light', s:'No new data infrastructure. Integration via existing Transguard SCDM telemetry and client ERP APIs.' },
  ];

  slide.addText('CLIENT IMPACT', { x:0.5, y:1.5, w:4.5, h:0.25, fontSize:8, bold:true, color:C.teal, charSpacing:3 });
  slide.addText('TRANSGUARD ROI', { x:5.2, y:1.5, w:4.5, h:0.25, fontSize:8, bold:true, color:C.gold, charSpacing:3 });

  leftMetrics.forEach((m, i) => {
    const y = 1.85 + i * 1.15;
    slide.addShape(pptx.ShapeType.rect, { x:0.5, y, w:4.4, h:1.0, fill:{color: C.hero}, line:{color: C.teal, width:0.5, transparency:60} });
    slide.addText(m.v, { x:0.65, y:y+0.08, w:1.4, h:0.5, fontSize:26, bold:true, color:C.teal, fontFace:'Courier New' });
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

  slide.addShape(pptx.ShapeType.line, { x:4.97, y:1.5, w:0, h:5.2, line:{color: C.teal, width:0.5, transparency:60} });
}

// ══════════════════════════════════════════
// SLIDE 7 — PROOF OF WORK
// ══════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.lgray };

  slide.addText('PROOF OF WORK', { x:0.5, y:0.3, w:9, h:0.25, fontSize:9, bold:true, color:C.dark, charSpacing:4 });
  slide.addText('What I built at PhonePe maps exactly to what CashPulse requires', { x:0.5, y:0.6, w:9, h:0.7, fontSize:24, bold:true, color:C.dark });

  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:1.45, w:4.4, h:5.3, fill:{color:C.dark}, line:{color:C.teal, width:0.5} });
  slide.addText('WHAT I BUILT AT PHONEPE', { x:0.6, y:1.55, w:4.2, h:0.25, fontSize:8, bold:true, color:C.teal, charSpacing:2 });

  const phonepe = [
    { t:'Self-Serve Merchant Dashboards', d:'Built merchant-facing intelligence layer on top of back-office ops data — merchants configure and monitor offer engines in real-time. Same pattern as CashPulse giving clients visibility into SCDM data.' },
    { t:'Offer Ops TAT: 2 Days → 30 Minutes', d:'Reduced B2B workflow TAT from 2 days to 30 minutes by surfacing operational data as a self-serve platform. CashPulse reduces reconciliation from 24 hrs to <5 mins using the same principle.' },
    { t:'5,000+ B2B Partners Onboarded', d:'Built zero-to-one B2B integration platform with Instant Discount and EMI subvention capabilities. Proven ability to design and launch B2B SaaS products that scale to enterprise clients.' },
    { t:'ML-Driven Targeting: -32% Marketing Burn', d:'Deployed propensity-to-transact models replacing static rules with real-time user-level scoring. CashPulse predictive collection scheduling uses the same ML architecture applied to fill velocity.' },
  ];
  phonepe.forEach((p, i) => {
    slide.addShape(pptx.ShapeType.rect, { x:0.65, y:2.0+(i*1.15), w:4.1, h:1.0, fill:{color: C.hero}, line:{color: C.teal, width:0.3, transparency:60} });
    slide.addText(p.t, { x:0.78, y:2.08+(i*1.15), w:3.9, h:0.26, fontSize:10, bold:true, color:C.teal });
    slide.addText(p.d, { x:0.78, y:2.35+(i*1.15), w:3.9, h:0.55, fontSize:8.5, color:C.muted });
  });

  slide.addShape(pptx.ShapeType.rect, { x:5.1, y:1.45, w:4.4, h:5.3, fill:{color:C.white}, line:{color:'CBD5E1', width:0.5} });
  slide.addText('HOW IT MAPS TO CASHPULSE', { x:5.2, y:1.55, w:4.2, h:0.25, fontSize:8, bold:true, color:C.dark, charSpacing:2 });

  const mapping = [
    { t:'CashPulse Dashboard = PhonePe Merchant Dashboard', d:'Identical design principle: take opaque back-office data and surface it as a client-facing intelligence layer. The only difference is the domain — cash management vs offers management.' },
    { t:'ERP Reconciliation = Offer Ops Automation', d:'Both replace a manual, multi-day human process with a real-time automated workflow. The PM skill is identical: map the workflow, identify the data dependency, build the integration layer.' },
    { t:'Enterprise Client Onboarding = B2B Platform Launch', d:'PhonePe PG platform onboarded 5,000+ merchants at zero engineering dependency. CashPulse POC requires the same skill: design a low-friction integration experience for enterprise clients.' },
    { t:'Predictive Scheduling = Propensity-to-Transact ML', d:'The ML mechanic is the same: use historical behavioural data to predict future state and trigger an automated action. Applied to fill velocity instead of user purchase intent.' },
  ];
  mapping.forEach((m, i) => {
    slide.addShape(pptx.ShapeType.rect, { x:5.25, y:2.0+(i*1.15), w:4.1, h:1.0, fill:{color:C.lgray}, line:{color:'CBD5E1', width:0.5} });
    slide.addText(m.t, { x:5.38, y:2.08+(i*1.15), w:3.9, h:0.26, fontSize:10, bold:true, color:C.dark });
    slide.addText(m.d, { x:5.38, y:2.35+(i*1.15), w:3.9, h:0.55, fontSize:8.5, color:'475569' });
  });

  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:6.95, w:9, h:0.75, fill:{color:C.dark}, line:{color: C.teal, width:0.5} });
  slide.addText('"The product gap I saw at Transguard is the same gap I already closed at PhonePe — just applied to a different data layer. I have shipped this exact pattern before."', { x:0.65, y:7.0, w:8.7, h:0.65, fontSize:10, color:C.teal, valign:'middle', italic:true });
}

// ══════════════════════════════════════════
// SLIDE 8 — ROLLOUT PLAN
// ══════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.dark };

  slide.addText('ROLLOUT PLAN', { x:0.5, y:0.3, w:9, h:0.25, fontSize:9, bold:true, color:C.teal, charSpacing:4 });
  slide.addText('POC in 30 days · Revenue in 90 days · Scale in 180 days', { x:0.5, y:0.6, w:9, h:0.7, fontSize:26, bold:true, color:C.white });

  const phases = [
    { n:'Phase 01', t:'POC Launch', dur:'Days 1–90', items:['Anchor clients: Carrefour UAE + ENOC Group','150 SCDMs on CashPulse network','ERP integration: SAP S/4HANA (Carrefour)','Success metric: <5 min reconciliation time + NPS >8'], color: C.teal },
    { n:'Phase 02', t:'Full Productization', dur:'Days 91–180', items:['Roll out to all 6 enterprise accounts (847 SCDMs)','Launch dashboard licensing model (AED 213K/client/yr)','Add Oracle + Microsoft Dynamics ERP connectors','Hire 1 implementation engineer'], color: C.gold },
    { n:'Phase 03', t:'Market Expansion', dur:'Days 181–365', items:['License CashPulse to non-Transguard SCDM operators','Explore white-label to G4S / Brinks UAE','Pursue banking sector (ATM reconciliation use case)','Target: AED 3.2M ARR by end of Year 1'], color: C.green },
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

  slide.addShape(pptx.ShapeType.rect, { x:0.5, y:6.2, w:9, h:1.15, fill:{color: C.hero}, line:{color: C.teal, width:1} });
  slide.addText('What I need from Transguard to start:', { x:0.65, y:6.28, w:8.7, h:0.28, fontSize:10, bold:true, color:C.teal });
  slide.addText('1 anchor client introduction (Carrefour or ENOC) · Access to SCDM telemetry API documentation · Product engineering resource (1 BE, 1 FE) · 90-day POC mandate with Go/No-Go review', { x:0.65, y:6.58, w:8.7, h:0.58, fontSize:9, color:C.muted });

  slide.addText('Ajay Avaghade · avaghadeajay009@gmail.com · ajay-avaghade.github.io/Portfolio', { x:0.5, y:7.38, w:9, h:0.25, fontSize:9, color:C.muted, align:'center' });
}

pptx.writeFile({ fileName: 'assets/transguard-cashpulse-deck.pptx' })
  .then(() => console.log('CashPulse deck saved: assets/transguard-cashpulse-deck.pptx'))
  .catch(err => console.error(err));
