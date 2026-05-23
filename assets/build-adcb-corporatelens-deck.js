const pptx = require('pptxgenjs');
const prs = new pptx();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:  '00302A',
  hero:  '00453C',
  bright:'00BFA5',
  gold:  'D4A843',
  green: '22C55E',
  red:   'EF4444',
  amber: 'F59E0B',
  white: 'FFFFFF',
  lgray: 'F0F7F5',
  muted: '7EB5AE',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });

// ─── SLIDE 1 — COVER ───
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  for(let i=0;i<6;i++){
    s.addShape(prs.ShapeType.line, { x:0+i*1.8, y:0, w:0, h:7.5, line:{ color:C.bright, width:0.4, transparency:88 } });
  }
  s.addShape(prs.ShapeType.rect, { x:0, y:0, w:10, h:0.06, fill:{ color:C.bright } });
  s.addText('ADCB  ·  CORPORATE CARDS PM  ·  TRANSACTION BANKING', {
    x:0.6, y:0.3, w:8.5, h:0.3, fontSize:9, bold:true, color:C.bright, charSpacing:8, fontFace:'Calibri'
  });
  s.addText('CorporateLens', { x:0.6, y:1.0, w:8, h:1.2, fontSize:64, bold:true, color:C.white, fontFace:'Calibri' });
  s.addText('Real-Time Spend Control & Compliance for Corporate Card Programmes', {
    x:0.6, y:2.25, w:7.5, h:0.5, fontSize:20, color:C.gold, fontFace:'Calibri', italic:true
  });
  s.addShape(prs.ShapeType.rect, { x:0.6, y:2.85, w:1.4, h:0.05, fill:{ color:C.gold } });
  s.addText('Presented by Ajay Avaghade  ·  Corporate Cards & Transaction Banking PM', {
    x:0.6, y:3.05, w:8, h:0.3, fontSize:12, color:C.muted, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:7.0, y:5.2, w:2.7, h:1.8, fill:{ color:C.hero }, line:{ color:C.gold, width:1 } });
  s.addText('48 hrs', { x:7.0, y:5.3, w:2.7, h:0.7, fontSize:38, bold:true, color:C.gold, align:'center', fontFace:'Calibri' });
  s.addText('to change a card limit\nthe manual way', { x:7.0, y:6.0, w:2.7, h:0.8, fontSize:11, color:C.muted, align:'center', fontFace:'Calibri' });
  s.addShape(prs.ShapeType.rect, { x:0, y:7.44, w:10, h:0.06, fill:{ color:C.gold } });
}

// ─── SLIDE 2 — THE PROBLEM ───
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('THE PROBLEM', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.muted, charSpacing:8, fontFace:'Calibri' });
  s.addText('Corporate Finance Teams Are Flying Blind.\nNo Real-Time Visibility. No Self-Serve Controls.', {
    x:0.6, y:0.6, w:8.8, h:1.1, fontSize:26, bold:true, color:C.white, fontFace:'Calibri'
  });
  const cols3 = [
    { ico:'📞', stat:'48 hrs', lbl:'To change a card limit', sub:'Every control change requires calling the relationship manager and waiting for back-office processing' },
    { ico:'📄', stat:'Month-end', lbl:'Only time data is available', sub:'Corporate clients receive PDF statements 5–7 days after month close — too late to act' },
    { ico:'🔒', stat:'0%', lbl:'Self-serve control rate', sub:'No ability to freeze a card, block a merchant category, or set per-project limits without bank intervention' },
  ];
  cols3.forEach((c,i)=>{
    const x = 0.5 + i*3.1;
    s.addShape(prs.ShapeType.rect, { x, y:1.9, w:2.9, h:3.0, fill:{ color:C.hero }, line:{ color:C.bright, width:0.5, transparency:60 }, shadow:makeShadow() });
    s.addText(c.ico, { x, y:2.05, w:2.9, h:0.5, fontSize:24, align:'center' });
    s.addText(c.stat, { x, y:2.6, w:2.9, h:0.75, fontSize:34, bold:true, color:C.gold, align:'center', fontFace:'Calibri' });
    s.addText(c.lbl, { x, y:3.35, w:2.9, h:0.4, fontSize:12, bold:true, color:C.white, align:'center', fontFace:'Calibri', wrap:true });
    s.addText(c.sub, { x:x+0.15, y:3.75, w:2.6, h:0.9, fontSize:10, color:C.muted, align:'center', fontFace:'Calibri', wrap:true });
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.2, w:9.0, h:1.6, fill:{ color:'001F1A' }, line:{ color:C.gold, width:1 } });
  s.addText('Root cause: ', { x:0.8, y:5.35, w:1.4, h:0.3, fontSize:11, bold:true, color:C.gold, fontFace:'Calibri' });
  s.addText('Corporate card programs at UAE banks are operated as static, relationship-managed products. There is no digital control plane. Every change goes through a human, creating delays and errors that erode trust and card utilisation.', {
    x:0.8, y:5.65, w:8.5, h:0.9, fontSize:11, color:C.muted, fontFace:'Calibri', wrap:true
  });
}

// ─── SLIDE 3 — THE INSIGHT ───
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };
  s.addText('THE INSIGHT', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.bright, charSpacing:8, fontFace:'Calibri' });
  s.addText('A Control Change Should Take 8 Seconds.\nNot 48 Hours.', {
    x:0.6, y:0.6, w:8.8, h:1.0, fontSize:28, bold:true, color:C.dark, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.85, w:4.2, h:3.4, fill:{ color:'FFF5F5' }, line:{ color:'EF4444', width:1 } });
  s.addText('❌  Current State: Relationship-Managed', { x:0.7, y:2.0, w:3.8, h:0.4, fontSize:13, bold:true, color:'991B1B', fontFace:'Calibri' });
  ['Finance admin calls RM to adjust limit','RM emails back-office operations team','Back-office processes request manually','Card limit updated in core banking system','Confirmation sent — 48 hours later'].forEach((pt,i)=>{
    s.addText((i+1)+'. '+pt, { x:0.8, y:2.55+i*0.55, w:3.7, h:0.45, fontSize:11, color:C.dark, fontFace:'Calibri', wrap:true });
  });
  s.addShape(prs.ShapeType.ellipse, { x:4.55, y:2.95, w:0.9, h:0.9, fill:{ color:C.dark } });
  s.addText('VS', { x:4.55, y:2.98, w:0.9, h:0.85, fontSize:13, bold:true, color:C.white, align:'center', fontFace:'Calibri' });
  s.addShape(prs.ShapeType.rect, { x:5.0, y:1.85, w:4.4, h:3.4, fill:{ color:'F0FDF8' }, line:{ color:C.green, width:1 } });
  s.addText('✅  CorporateLens: Self-Serve Portal', { x:5.2, y:2.0, w:4.0, h:0.4, fontSize:13, bold:true, color:'166534', fontFace:'Calibri' });
  ['Finance admin opens CorporateLens mobile portal','Selects employee card and taps "Adjust Limit"','Sets new limit and confirms with biometric','Change pushed to Visa authorisation host in real-time','Limit live in 8 seconds — zero human intervention'].forEach((pt,i)=>{
    s.addText((i+1)+'. '+pt, { x:5.3, y:2.55+i*0.55, w:3.9, h:0.45, fontSize:11, color:C.dark, fontFace:'Calibri', wrap:true });
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:5.45, w:9.0, h:1.4, fill:{ color:C.dark } });
  s.addText('Key insight: The Visa Commercial API already supports real-time limit updates. The product gap is not technical — it is a missing self-serve control plane on top of an existing capability. CorporateLens builds that layer.', {
    x:0.8, y:5.6, w:8.5, h:1.1, fontSize:12, color:C.lgray, fontFace:'Calibri', wrap:true, italic:true
  });
}

// ─── SLIDE 4 — THE MECHANIC ───
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('THE MECHANIC', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.muted, charSpacing:8, fontFace:'Calibri' });
  s.addText('Three Layers. One Product. Zero Calls to the RM.', {
    x:0.6, y:0.6, w:8.5, h:0.7, fontSize:28, bold:true, color:C.white, fontFace:'Calibri'
  });
  const layers = [
    { title:'Layer 1 — Real-Time Data', color:C.bright, pts:['Live transaction feed across all employee cards','Authorisation holds visible as they happen','Spend velocity alerts at 70%, 90%, 100% of limit','MCC-level categorisation for policy compliance checks'] },
    { title:'Layer 2 — Self-Serve Controls', color:C.gold, pts:['Per-card monthly limit adjustment — live in 8 seconds via Visa API','MCC block toggles (travel, entertainment, cash advances)','One-tap card freeze/unfreeze for lost or misused cards','Per-employee spending reports on demand'] },
    { title:'Layer 3 — Scheme Compliance', color:C.green, pts:['Visa mandate tracker — EMV 3DS, tokenisation, reporting v2','Dispute and chargeback ratio monitoring vs Visa thresholds','Flagged transaction review queue for risk team','Audit trail for all control changes — governance-ready'] },
  ];
  layers.forEach((l,i)=>{
    const x = 0.4 + i*3.2;
    s.addShape(prs.ShapeType.rect, { x, y:1.5, w:3.0, h:4.8, fill:{ color:C.hero }, line:{ color:l.color, width:1 } });
    s.addShape(prs.ShapeType.rect, { x, y:1.5, w:3.0, h:0.06, fill:{ color:l.color } });
    s.addText(l.title, { x:x+0.15, y:1.65, w:2.7, h:0.5, fontSize:12, bold:true, color:l.color, fontFace:'Calibri', wrap:true });
    l.pts.forEach((pt,j)=>{
      s.addText('• '+pt, { x:x+0.15, y:2.3+j*0.95, w:2.7, h:0.85, fontSize:10, color:C.muted, fontFace:'Calibri', wrap:true });
    });
  });
  s.addShape(prs.ShapeType.rect, { x:0.4, y:6.5, w:9.2, h:0.8, fill:{ color:'001A16' }, line:{ color:C.bright, width:0.8 } });
  s.addText('All three layers built on existing Visa Commercial APIs and ADCB core banking infrastructure — no new scheme integrations required. The product gap is the portal, not the plumbing.', {
    x:0.65, y:6.6, w:8.8, h:0.6, fontSize:10, color:C.muted, fontFace:'Calibri', wrap:true, italic:true
  });
}

// ─── SLIDE 5 — THE PRODUCT ───
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };
  s.addText('THE PRODUCT', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.bright, charSpacing:8, fontFace:'Calibri' });
  s.addText('4 Key Screens — From Dashboard to Scheme Ops', {
    x:0.6, y:0.6, w:8.5, h:0.65, fontSize:24, bold:true, color:C.dark, fontFace:'Calibri'
  });
  const cards = [
    { num:'01', title:'Admin Dashboard', desc:'Company-wide spend summary (AED 847K MTD), 23 active cards, 4 near-limit alerts. Real-time transaction feed. One-tap access to card controls and MIS download.', mock:'ADCB CorporateLens\nEmaar Properties\n────────────\nAED 847K spend\n23 active cards\n4 near limit\n────────────\n⚠️ Ahmad 74% limit\n🚨 Flagged: Laila\n────────────\n[+ Issue Card] [⬇ MIS]' },
    { num:'02', title:'Card Controls Panel', desc:'Per-employee card detail — current limit, MTD spend, MCC toggle controls (travel/hotels/restaurants allowed; gambling/cash blocked). Pending authorisations visible in real-time.', mock:'← Ahmad Al-Rashidi\nVisa •••• 4821\n● Active\n────────────\nLimit: AED 25,000\nSpent: AED 18,450\n[████████░░] 74%\n[✏️ Adjust Limit]\n────────────\nMCC Controls:\n✅ Travel  ✅ Hotels\n❌ Gambling  ❌ Cash' },
    { num:'03', title:'Limit Updated (Live)', desc:'Admin adjusts limit from AED 25K to AED 30K. Confirmation shows change is live on Visa authorisation host in 8 seconds — versus 48-hour manual process. Available balance updated instantly.', mock:'✅ Limit Updated\nLive in 8 seconds\n────────────\nAhmad Al-Rashidi\nAED 25K → AED 30K\n● Live on Visa host\n────────────\nPrevious:  AED 25,000\nNew limit: AED 30,000\nAvailable: AED 11,550\nTime:      8 seconds\nPreviously:  48 hours' },
    { num:'04', title:'Risk & Scheme Ops', desc:'ADCB ops view — flagged transactions, Visa mandate compliance tracker (EMV 3DS: 45 days remaining), scheme scorecard. Single dashboard for risk monitoring and regulatory readiness.', mock:'CorporateLens Ops\n23 programs · 2 flagged\n────────────\n🚨 Laila: Out-country\n⚠️ Khalid: Over-limit\n────────────\nVisa Mandates:\nEMV 3DS    45 days\nReporting v2  120 days\nTokenisation  ✓ Done\n────────────\nAuth rate: 94.2%' },
  ];
  cards.forEach((c,i)=>{
    const x = 0.3 + i*2.42;
    s.addShape(prs.ShapeType.rect, { x, y:1.45, w:2.3, h:4.9, fill:{ color:C.white }, line:{ color:'E2E8F0', width:1 }, shadow:makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y:1.45, w:2.3, h:0.04, fill:{ color:C.bright } });
    s.addText(c.num, { x, y:1.5, w:2.3, h:0.25, fontSize:8, bold:true, color:C.bright, align:'center', fontFace:'Calibri', charSpacing:4 });
    s.addText(c.title, { x:x+0.1, y:1.75, w:2.1, h:0.45, fontSize:11, bold:true, color:C.dark, fontFace:'Calibri', wrap:true });
    s.addShape(prs.ShapeType.rect, { x:x+0.1, y:2.25, w:2.1, h:2.15, fill:{ color:'F8FAFC' }, line:{ color:'E2E8F0', width:0.5 } });
    s.addText(c.mock, { x:x+0.12, y:2.28, w:2.06, h:2.1, fontSize:7, color:'374151', fontFace:'Courier New', wrap:true });
    s.addText(c.desc, { x:x+0.1, y:4.45, w:2.1, h:1.7, fontSize:9, color:'4B5563', fontFace:'Calibri', wrap:true });
  });
}

// ─── SLIDE 6 — IMPACT & ROI ───
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('IMPACT & ROI', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.muted, charSpacing:8, fontFace:'Calibri' });
  s.addText('Projected Impact Across Corporate, Risk, and Operations', {
    x:0.6, y:0.6, w:8.5, h:0.7, fontSize:26, bold:true, color:C.white, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.5, w:4.3, h:0.4, fill:{ color:C.hero } });
  s.addText('CORPORATE CLIENT IMPACT', { x:0.5, y:1.52, w:4.3, h:0.36, fontSize:10, bold:true, color:C.bright, align:'center', fontFace:'Calibri', charSpacing:4 });
  [
    { val:'−97%', lbl:'Time to change card controls\n(48 hours → 8 seconds)' },
    { val:'+28%', lbl:'Card utilisation rate vs control group\n(more spend confidence)' },
    { val:'−40%', lbl:'RM support calls for routine\ncontrol changes' },
    { val:'+35%', lbl:'Finance admin satisfaction\n(NPS proxy for B2B SaaS)' },
  ].forEach((m,i)=>{
    const y = 2.1 + i*1.1;
    s.addShape(prs.ShapeType.rect, { x:0.5, y, w:4.3, h:1.0, fill:{ color:C.hero }, line:{ color:C.bright, width:0.5, transparency:60 } });
    s.addText(m.val, { x:0.6, y:y+0.05, w:1.4, h:0.85, fontSize:30, bold:true, color:C.bright, fontFace:'Calibri', valign:'middle' });
    s.addText(m.lbl, { x:2.1, y:y+0.1, w:2.5, h:0.8, fontSize:11, color:C.muted, fontFace:'Calibri', wrap:true, valign:'middle' });
  });
  s.addShape(prs.ShapeType.rect, { x:5.2, y:1.5, w:4.3, h:0.4, fill:{ color:'001A16' }, line:{ color:C.gold, width:1 } });
  s.addText('BANK (ADCB) ROI', { x:5.2, y:1.52, w:4.3, h:0.36, fontSize:10, bold:true, color:C.gold, align:'center', fontFace:'Calibri', charSpacing:4 });
  [
    { val:'−40%', lbl:'Operational cost for card\ncontrol changes (RM + back-office)' },
    { val:'+28%', lbl:'Monthly card spend per programme\n(higher limit confidence = more usage)' },
    { val:'100%', lbl:'Scheme mandate audit readiness\n(compliance tracker always current)' },
    { val:'+18%', lbl:'Corporate card programme\nretention at annual renewal' },
  ].forEach((m,i)=>{
    const y = 2.1 + i*1.1;
    s.addShape(prs.ShapeType.rect, { x:5.2, y, w:4.3, h:1.0, fill:{ color:'001A16' }, line:{ color:C.gold, width:0.5, transparency:60 } });
    s.addText(m.val, { x:5.3, y:y+0.05, w:1.4, h:0.85, fontSize:30, bold:true, color:C.gold, fontFace:'Calibri', valign:'middle' });
    s.addText(m.lbl, { x:6.8, y:y+0.1, w:2.5, h:0.8, fontSize:11, color:C.muted, fontFace:'Calibri', wrap:true, valign:'middle' });
  });
}

// ─── SLIDE 7 — PROOF OF WORK ───
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };
  s.addText('PROOF OF WORK', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.bright, charSpacing:8, fontFace:'Calibri' });
  s.addText('I Built the Controls Engine. Here Is the Proof.', {
    x:0.6, y:0.6, w:8.5, h:0.65, fontSize:26, bold:true, color:C.dark, fontFace:'Calibri'
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:1.45, w:4.3, h:4.4, fill:{ color:C.dark } });
  s.addText('What I Built at PhonePe', { x:0.7, y:1.65, w:3.9, h:0.35, fontSize:12, bold:true, color:C.bright, fontFace:'Calibri' });
  [
    'Redesigned end-to-end offers eligibility and controls engine for 350M MAU — defining which offers surface to which users, with real-time overrides — same architecture as CorporateLens card controls layer',
    'Led deployment of Propensity-to-Transact ML models replacing static targeting rules with real-time user-level scoring — same AI categorisation logic as CorporateLens spend analytics engine',
    'Managed RBI regulatory compliance requirements for UPI payment products — adjacent domain to Visa/Mastercard scheme mandate management',
    'Rebuilt a fragmented rewards stack with clear product ownership, backlog prioritisation, and delivery across Tech, Ops, Finance, and Legal — same cross-functional governance model as ADCB',
    'Owned post-launch product iteration: monitoring dashboards, root-cause analysis, and continuous enhancement roadmap — matches ADCB production ownership requirement',
  ].forEach((pt,i)=>{
    s.addText('→  '+pt, { x:0.7, y:2.15+i*0.75, w:3.9, h:0.65, fontSize:10, color:'CBD5E1', fontFace:'Calibri', wrap:true });
  });
  s.addShape(prs.ShapeType.rect, { x:5.1, y:1.45, w:4.3, h:4.4, fill:{ color:'F1F5F9' }, line:{ color:'E2E8F0', width:1 } });
  s.addText('How It Maps to This Role', { x:5.3, y:1.65, w:3.9, h:0.35, fontSize:12, bold:true, color:C.dark, fontFace:'Calibri' });
  [
    'JD: "Product roadmap and backlog ownership" → Owned full offers product backlog at PhonePe — prioritisation, sprint planning, and release readiness across 10+ engineering squads',
    'JD: "Translate business needs into user stories and acceptance criteria" → Wrote structured requirements for ML model deployment and UX redesign across 350M MAU — fully agile delivery',
    'JD: "Coordinate with scheme representatives and vendors" → Managed RBI and NPCI compliance mandates at PhonePe — analogous to Visa/Mastercard scheme coordination',
    'JD: "UAT scope and formal sign-off" → Owned UAT processes for every major offers release at PhonePe — defined test scenarios, coordinated execution, and signed off on production deployment',
    'JD: "Retain product ownership post-launch" → Built monitoring dashboards and owned root-cause analysis for all live products — not a hand-off PM',
  ].forEach((pt,i)=>{
    s.addText('✓  '+pt, { x:5.3, y:2.15+i*0.75, w:3.9, h:0.65, fontSize:10, color:'374151', fontFace:'Calibri', wrap:true });
  });
  s.addShape(prs.ShapeType.rect, { x:0.5, y:6.05, w:9.0, h:1.0, fill:{ color:C.dark } });
  s.addText('"The CorporateLens control plane is the same architecture as the offers eligibility engine I redesigned at PhonePe. The technology is different; the product thinking — controls, rules, real-time override, compliance audit trail — is identical."', {
    x:0.8, y:6.15, w:8.5, h:0.8, fontSize:11, color:C.lgray, italic:true, fontFace:'Calibri', wrap:true
  });
}

// ─── SLIDE 8 — ROLLOUT PLAN ───
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('ROLLOUT PLAN', { x:0.6, y:0.3, w:8, h:0.25, fontSize:9, bold:true, color:C.muted, charSpacing:8, fontFace:'Calibri' });
  s.addText('Phased Delivery — Agile, Scheme-Compliant, Risk-Controlled', {
    x:0.6, y:0.6, w:8.5, h:0.65, fontSize:26, bold:true, color:C.white, fontFace:'Calibri'
  });
  const phases = [
    { ph:'Phase 1', time:'Month 1–3', title:'MVP — Real-Time Dashboard', color:C.bright, pts:['Deliver spend dashboard and authorisation feed for pilot corporate clients','Requirements gathering, user stories, and UAT with 5 pilot companies','Instrument card utilisation and RM call volume as baseline metrics'] },
    { ph:'Phase 2', time:'Month 4–6', title:'Self-Serve Controls', color:C.gold, pts:['Deploy per-card limit adjustment via Visa Commercial API — live in seconds','MCC toggle controls and one-tap freeze — scheme-compliance reviewed','Expand to 50 corporate programmes; A/B test portal adoption vs RM-managed'] },
    { ph:'Phase 3', time:'Month 7–9', title:'Scheme Ops & Scale', color:C.green, pts:['Launch scheme mandate tracker integrated with compliance team','Roll out to all ADCB corporate card programmes across MENA','Post-launch ownership: monitoring, root-cause analysis, and roadmap v2'] },
  ];
  phases.forEach((ph,i)=>{
    const x = 0.4 + i*3.2;
    s.addShape(prs.ShapeType.rect, { x, y:1.5, w:3.0, h:4.1, fill:{ color:C.hero }, line:{ color:ph.color, width:1 } });
    s.addShape(prs.ShapeType.rect, { x, y:1.5, w:3.0, h:0.06, fill:{ color:ph.color } });
    s.addText(ph.ph, { x:x+0.15, y:1.6, w:2.7, h:0.3, fontSize:9, bold:true, color:ph.color, fontFace:'Calibri', charSpacing:4 });
    s.addText(ph.time, { x:x+0.15, y:1.9, w:2.7, h:0.25, fontSize:10, color:C.muted, fontFace:'Calibri' });
    s.addText(ph.title, { x:x+0.15, y:2.2, w:2.7, h:0.5, fontSize:14, bold:true, color:C.white, fontFace:'Calibri', wrap:true });
    ph.pts.forEach((pt,j)=>{
      s.addText('• '+pt, { x:x+0.15, y:2.85+j*0.85, w:2.7, h:0.75, fontSize:10, color:C.muted, fontFace:'Calibri', wrap:true });
    });
  });
  s.addShape(prs.ShapeType.rect, { x:0.4, y:5.8, w:6.4, h:1.35, fill:{ color:'001A16' }, line:{ color:C.bright, width:1 } });
  s.addText('What I need to build this:', { x:0.65, y:5.93, w:5.8, h:0.3, fontSize:11, bold:true, color:C.bright, fontFace:'Calibri' });
  s.addText('Access to Visa Commercial API documentation and scheme mandate calendar  ·  Alignment with ADCB core banking and operations teams  ·  2 engineers for portal + API integration  ·  Pilot corporate client for requirements and UAT', {
    x:0.65, y:6.25, w:5.8, h:0.7, fontSize:10, color:C.muted, fontFace:'Calibri', wrap:true
  });
  s.addShape(prs.ShapeType.rect, { x:7.05, y:5.8, w:2.5, h:1.35, fill:{ color:C.bright } });
  s.addText('Ajay Avaghade', { x:7.1, y:5.95, w:2.4, h:0.3, fontSize:12, bold:true, color:C.dark, align:'center', fontFace:'Calibri' });
  s.addText('avaghadeajay009@gmail.com', { x:7.1, y:6.3, w:2.4, h:0.25, fontSize:9, color:C.dark, align:'center', fontFace:'Calibri' });
  s.addText('+91 9561558439', { x:7.1, y:6.6, w:2.4, h:0.25, fontSize:9, color:C.dark, align:'center', fontFace:'Calibri' });
}

prs.writeFile({ fileName:'assets/adcb-corporatelens-deck.pptx' }).then(()=>console.log('adcb-corporatelens-deck.pptx done'));
