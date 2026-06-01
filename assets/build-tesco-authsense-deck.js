const PptxGenJS = require('pptxgenjs');

const C = {
  dark:  '0A1628',
  hero:  '102040',
  blue:  '00539F',
  red:   'E2001A',
  gold:  'F59E0B',
  white: 'FFFFFF',
  lgray: 'F0F4FF',
  muted: '6B7A99',
  green: '16A34A',
  redhero:'1A0408',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.13 });

async function buildAuthSense() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = 'AuthSense — Tesco Payment Authorization Intelligence';

  // ─── SLIDE 1: COVER ───────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.dark} });
    for(let i=0;i<6;i++){
      s.addShape(pptx.ShapeType.line, { x:-0.5+i*1.8, y:0, w:2.4, h:7.5, line:{color:C.red, width:0.5, transparency:85} });
    }
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.18, h:'100%', fill:{color:C.red} });
    s.addText('TESCO TECHNOLOGY · PAYMENTS PLATFORM · CASE STUDY 01', {
      x:0.35, y:0.22, w:9, h:0.25,
      fontSize:8, bold:true, color:C.muted, charSpacing:3, fontFace:'Calibri'
    });
    s.addText('AuthSense', {
      x:0.35, y:0.7, w:7, h:1.2,
      fontSize:52, bold:true, color:C.white, fontFace:'Calibri'
    });
    s.addText('Real-Time Payment Authorization Intelligence — 12-Minute Detection vs 4-Hour Lag', {
      x:0.35, y:1.85, w:8.5, h:0.55,
      fontSize:16, color:C.gold, fontFace:'Calibri', italic:true
    });
    s.addText('Ajay Avaghade · Product Manager', {
      x:0.35, y:2.55, w:5, h:0.3,
      fontSize:11, color:C.muted, fontFace:'Calibri'
    });
    s.addShape(pptx.ShapeType.roundRect, {
      x:7.1, y:5.5, w:2.65, h:1.6,
      fill:{color:C.red}, rectRadius:0.1, shadow:makeShadow()
    });
    s.addText('−40%', {
      x:7.1, y:5.56, w:2.65, h:0.75,
      fontSize:36, bold:true, color:C.white, align:'center', fontFace:'Calibri'
    });
    s.addText('MTTR on auth\nincidents', {
      x:7.1, y:6.25, w:2.65, h:0.65,
      fontSize:9, color:'FFAAAA', align:'center', fontFace:'Calibri'
    });
  }

  // ─── SLIDE 2: THE PROBLEM ──────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.dark} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.red} });
    s.addText('THE PROBLEM', {
      x:0.3, y:0.22, w:4, h:0.22,
      fontSize:9, bold:true, color:C.red, charSpacing:3, fontFace:'Calibri'
    });
    s.addText('When Authorization Rates Degrade, Tesco Finds Out From Customer Complaints — Not from the Platform', {
      x:0.3, y:0.5, w:9.4, h:0.9,
      fontSize:22, bold:true, color:C.white, fontFace:'Calibri'
    });

    const cols = [
      { icon:'⏳', stat:'4+ Hours', lbl:'Avg Detection Lag', sub:'Auth rate degradations typically detected from customer complaints or end-of-day reports — hours after onset' },
      { icon:'💸', stat:'£2,400/min', lbl:'Estimated GMV Impact', sub:'At Tesco online transaction volume, a 1% auth rate drop equates to thousands of pounds per minute in failed orders' },
      { icon:'🔍', stat:'Zero', lbl:'Real-Time PSP View', sub:'No current tool shows auth rates broken down by PSP × payment method × issuer in real-time. Ops flies blind.' },
    ];
    cols.forEach((c2,i) => {
      const x = 0.3 + i*3.25;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.75, w:3.0, h:3.15, fill:{color:C.hero}, rectRadius:0.1, shadow:makeShadow() });
      s.addText(c2.icon, { x, y:1.85, w:3.0, h:0.5, fontSize:24, align:'center' });
      s.addText(c2.stat, { x, y:2.4, w:3.0, h:0.55, fontSize:24, bold:true, color:C.red, align:'center', fontFace:'Calibri' });
      s.addText(c2.lbl, { x, y:2.94, w:3.0, h:0.3, fontSize:11, bold:true, color:C.white, align:'center', fontFace:'Calibri' });
      s.addText(c2.sub, { x:x+0.12, y:3.28, w:2.76, h:1.3, fontSize:10, color:C.muted, align:'center', fontFace:'Calibri', valign:'top' });
    });

    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:5.1, w:9.4, h:1.0, fill:{color:'1A0810'}, rectRadius:0.06 });
    s.addText('The JD explicitly calls this out: "ensuring payment processing metrics are in good health and proposing initiatives for optimising authorisation success rates." AuthSense is that initiative — built as a working product.', {
      x:0.5, y:5.17, w:9.0, h:0.85,
      fontSize:10, color:C.muted, fontFace:'Calibri', italic:true, valign:'middle'
    });
  }

  // ─── SLIDE 3: THE INSIGHT ──────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.white} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.red} });
    s.addText('THE INSIGHT', {
      x:0.3, y:0.22, w:4, h:0.22,
      fontSize:9, bold:true, color:C.red, charSpacing:3, fontFace:'Calibri'
    });
    s.addText('The Platform Already Has All the Data to Detect Issues in 12 Minutes. It Just Has No One Watching.', {
      x:0.3, y:0.5, w:9.4, h:0.9,
      fontSize:22, bold:true, color:C.dark, fontFace:'Calibri'
    });

    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:1.65, w:4.5, h:3.8, fill:{color:'FFF8F8'}, rectRadius:0.1, line:{color:'FECACA', width:0.75} });
    s.addText('❌  Today', { x:0.5, y:1.8, w:4.1, h:0.35, fontSize:13, bold:true, color:C.red, fontFace:'Calibri' });
    const todayItems = [
      'Auth rates monitored in aggregate — PSP level invisible',
      'Detection: customer complaint or daily batch report (4+ hrs)',
      'Root cause: manual investigation across PSP portals',
      'Remediation: ad-hoc, depends on who is on call',
      'No GMV impact quantification during incident',
    ];
    todayItems.forEach((txt,i) => {
      s.addText('• '+txt, { x:0.5, y:2.26+i*0.57, w:4.1, h:0.45, fontSize:11, color:'374151', fontFace:'Calibri' });
    });

    s.addShape(pptx.ShapeType.ellipse, { x:4.62, y:3.2, w:0.56, h:0.56, fill:{color:C.red} });
    s.addText('VS', { x:4.62, y:3.2, w:0.56, h:0.56, fontSize:9, bold:true, color:C.white, align:'center', valign:'middle', fontFace:'Calibri' });

    s.addShape(pptx.ShapeType.roundRect, { x:5.0, y:1.65, w:4.7, h:3.8, fill:{color:'EFF6FF'}, rectRadius:0.1, line:{color:C.blue, width:0.75} });
    s.addText('✅  AuthSense', { x:5.15, y:1.8, w:4.3, h:0.35, fontSize:13, bold:true, color:C.blue, fontFace:'Calibri' });
    const asItems = [
      '3σ anomaly detection per PSP × method × issuer dimension',
      'Detection in 12 minutes — automatic alert with GMV impact',
      'Root cause isolated to 3 dimensions in seconds',
      'Playbook-driven remediation — any analyst can execute',
      'Real-time GMV impact shown throughout incident lifecycle',
    ];
    asItems.forEach((txt,i) => {
      s.addText('• '+txt, { x:5.15, y:2.26+i*0.57, w:4.35, h:0.45, fontSize:11, color:'1E3A5F', fontFace:'Calibri' });
    });

    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:5.65, w:9.4, h:0.65, fill:{color:'EFF6FF'}, rectRadius:0.06 });
    s.addText('Key insight: Every auth failure is already logged. The data exists. AuthSense is the intelligence layer that turns raw PSP logs into a real-time ops console — exactly as the PhonePe ops intelligence layer did for offer operations.', {
      x:0.5, y:5.7, w:9.0, h:0.55, fontSize:10, color:C.blue, fontFace:'Calibri', italic:true, valign:'middle'
    });
  }

  // ─── SLIDE 4: THE MECHANIC ─────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.hero} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.gold} });
    s.addText('THE MECHANIC', {
      x:0.3, y:0.22, w:4, h:0.22,
      fontSize:9, bold:true, color:C.gold, charSpacing:3, fontFace:'Calibri'
    });
    s.addText('Detect. Diagnose. Remediate. Close. Repeat.', {
      x:0.3, y:0.5, w:9.4, h:0.65,
      fontSize:28, bold:true, color:C.white, fontFace:'Calibri'
    });

    const steps = [
      { n:'1', title:'Continuous Monitoring', body:'Every PSP transaction feeds AuthSense in real-time. Auth rates computed per PSP × payment method × issuer dimension every 60 seconds. Baseline from rolling 90-minute window.' },
      { n:'2', title:'3σ Anomaly Detection', body:'If auth rate drops 3+ standard deviations below baseline, alert fires automatically. Detection time: typically 12 minutes after onset. GMV impact estimated instantly at current transaction volume.' },
      { n:'3', title:'Issuer Heatmap', body:'8 major UK issuers × 5 payment methods grid. Color-coded failure rates. Root cause isolated from "something is wrong" to "Barclays Visa Debit via Worldpay" — in seconds, not hours.' },
      { n:'4', title:'Playbook Remediation', body:'Each incident type has a defined 5-step playbook. Re-route to backup PSP, adjust retry timing, escalate to issuer. Any analyst executes — no dependency on senior engineer on call.' },
    ];

    steps.forEach((st,i) => {
      const x = 0.25 + i*2.42;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.4, w:2.25, h:4.3, fill:{color:C.dark}, rectRadius:0.1, shadow:makeShadow() });
      s.addShape(pptx.ShapeType.ellipse, { x:x+0.78, y:1.52, w:0.68, h:0.68, fill:{color:C.red} });
      s.addText(st.n, { x:x+0.78, y:1.52, w:0.68, h:0.68, fontSize:16, bold:true, color:C.white, align:'center', valign:'middle', fontFace:'Calibri' });
      s.addText(st.title, { x:x+0.1, y:2.32, w:2.05, h:0.45, fontSize:12, bold:true, color:C.gold, align:'center', fontFace:'Calibri' });
      s.addText(st.body, { x:x+0.1, y:2.82, w:2.05, h:2.5, fontSize:10, color:C.muted, fontFace:'Calibri', valign:'top' });
      if(i<3){
        s.addText('→', { x:x+2.18, y:3.3, w:0.3, h:0.4, fontSize:14, color:C.red, align:'center', fontFace:'Calibri', bold:true });
      }
    });

    s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:5.88, w:9.5, h:0.65, fill:{color:'0D2040'}, rectRadius:0.06 });
    s.addText('PhonePe proof: Unified ops intelligence layer rebuilt 6 siloed workflows into one console with automated breach classification — TAT 2 days → 4 hours, −68% escalation tickets. Identical architecture.', {
      x:0.45, y:5.93, w:9.1, h:0.55,
      fontSize:10, color:C.muted, fontFace:'Calibri', italic:true, valign:'middle'
    });
  }

  // ─── SLIDE 5: THE PRODUCT ──────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.white} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.red} });
    s.addText('THE PRODUCT', {
      x:0.3, y:0.22, w:4, h:0.22,
      fontSize:9, bold:true, color:C.red, charSpacing:3, fontFace:'Calibri'
    });
    s.addText('4 Screens — From Live Overview to Incident Resolution', {
      x:0.3, y:0.5, w:9.4, h:0.65,
      fontSize:24, bold:true, color:C.dark, fontFace:'Calibri'
    });

    const cards = [
      { n:'01', title:'Auth Overview', desc:'Live PSP-level auth rates, transaction volume, latency. 24h trend chart. Anomaly alert banner with GMV impact estimate. One-click to alert detail.', mono:'Worldpay  94.1% ▼\nAdyen     99.2% ✓\nStripe    98.7% ✓\n⚠ ALT-0847 ACTIVE' },
      { n:'02', title:'Alert Detail', desc:'Affected PSP × method × issuer breakdown. Estimated GMV impact per minute. Recommended action (re-route to backup PSP) with 1-click execution.', mono:'PSP:    Worldpay\nMethod: Visa Debit\nRate:   91.2% ▼\nImpact: £2,400/min' },
      { n:'03', title:'Issuer Heatmap', desc:'8 issuers × 5 methods grid. Color-coded failure rates. Recommended routing action per cell. Procurement evidence over time.', mono:'       Visa  MC\nBarcl  87% ■ 98%\nHSBC   93% □ 96%\nLloyds 99% ✓ 99%' },
      { n:'04', title:'Remediation', desc:'5-step incident playbook. Step-locked progression. Incident log with timestamp and GMV impact. MTTR tracked per incident type for SLA reporting.', mono:'Step 1 ✓ Confirmed\nStep 2 ✓ Isolated\nStep 3 → Re-route\nStep 4 □ Verify' },
    ];

    cards.forEach((c2,i) => {
      const x = 0.25 + i*2.44;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.38, w:2.28, h:4.95, fill:{color:'FFF8F8'}, rectRadius:0.1, line:{color:'FECACA', width:0.5}, shadow:makeShadow() });
      s.addShape(pptx.ShapeType.roundRect, { x:x+0.06, y:1.45, w:2.16, h:0.3, fill:{color:C.red}, rectRadius:0.05 });
      s.addText(c2.n+' · '+c2.title, { x:x+0.06, y:1.45, w:2.16, h:0.3, fontSize:9, bold:true, color:C.white, align:'center', fontFace:'Calibri' });
      s.addText(c2.desc, { x:x+0.1, y:1.85, w:2.08, h:1.6, fontSize:10, color:'374151', fontFace:'Calibri', valign:'top' });
      s.addShape(pptx.ShapeType.roundRect, { x:x+0.1, y:3.5, w:2.08, h:1.0, fill:{color:C.dark}, rectRadius:0.06 });
      s.addText(c2.mono, { x:x+0.14, y:3.55, w:2.0, h:0.9, fontSize:9, color:'F87171', fontFace:'Courier New', valign:'top' });
    });

    s.addText('Prototype: tesco-authsense-prototype.html · All 4 screens interactive · Alert → re-route → playbook flow live', {
      x:0.3, y:6.68, w:9.4, h:0.3,
      fontSize:9, color:C.muted, italic:true, fontFace:'Calibri', align:'center'
    });
  }

  // ─── SLIDE 6: IMPACT & ROI ─────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.dark} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.red} });
    s.addText('IMPACT & ROI', {
      x:0.3, y:0.22, w:4, h:0.22,
      fontSize:9, bold:true, color:C.red, charSpacing:3, fontFace:'Calibri'
    });
    s.addText('Every Minute Saved on Detection and Resolution Is Recovered GMV', {
      x:0.3, y:0.5, w:9.4, h:0.65,
      fontSize:24, bold:true, color:C.white, fontFace:'Calibri'
    });

    const leftMetrics = [
      { v:'12 min', l:'Anomaly detection', s:'From 4+ hours via complaints to 12 minutes via 3σ threshold' },
      { v:'−40%', l:'MTTR per incident', s:'Playbook-driven remediation vs ad-hoc on-call investigation' },
      { v:'3 dims', l:'Root cause speed', s:'PSP × method × issuer breakdown narrows cause in seconds' },
      { v:'100%', l:'Incidents documented', s:'Full audit trail for regulatory SLA reporting' },
    ];
    const rightMetrics = [
      { v:'£2,400/min', l:'GMV per incident', s:'Recovered GMV for each minute of detection/resolution time saved' },
      { v:'−60%', l:'False positive alerts', s:'Per-method thresholds vs aggregate global baseline' },
      { v:'+35%', l:'Est. tender win rate', s:'Live auth data as compliance evidence in city tender submissions' },
      { v:'PSP leverage', l:'Contract negotiation', s:'Issuer heatmap data quantifies PSP performance gaps for renegotiation' },
    ];

    [leftMetrics, rightMetrics].forEach((metrics, col) => {
      const xBase = col===0 ? 0.3 : 5.1;
      s.addText(col===0 ? 'Ops & Detection' : 'Business & Commercial', {
        x:xBase, y:1.35, w:4.5, h:0.28,
        fontSize:10, bold:true, color:col===0?C.red:C.gold, charSpacing:2, fontFace:'Calibri'
      });
      metrics.forEach((m,i) => {
        const y = 1.72 + i*0.92;
        s.addShape(pptx.ShapeType.roundRect, { x:xBase, y, w:4.55, h:0.82, fill:{color:C.hero}, rectRadius:0.06, shadow:makeShadow() });
        s.addText(m.v, { x:xBase+0.12, y:y+0.06, w:1.65, h:0.5, fontSize:col===0?20:16, bold:true, color:col===0?C.red:C.gold, fontFace:'Calibri', valign:'middle' });
        s.addText(m.l, { x:xBase+1.72, y:y+0.06, w:2.68, h:0.28, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
        s.addText(m.s, { x:xBase+1.72, y:y+0.36, w:2.68, h:0.36, fontSize:9, color:C.muted, fontFace:'Calibri', valign:'top' });
      });
    });
    s.addShape(pptx.ShapeType.rect, { x:4.88, y:1.35, w:0.02, h:5.35, fill:{color:'1E3A5F'} });

    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:6.5, w:9.4, h:0.75, fill:{color:'1A0810'}, rectRadius:0.06 });
    s.addText('At £2,400 GMV per minute of incident duration, shaving 3 hours off detection time alone recovers £432,000 per major incident. AuthSense pays for itself on the first serious auth degradation.', {
      x:0.5, y:6.56, w:9.0, h:0.6, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true, valign:'middle'
    });
  }

  // ─── SLIDE 7: PROOF OF WORK ────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:4.9, h:'100%', fill:{color:C.dark} });
    s.addShape(pptx.ShapeType.rect, { x:4.9, y:0, w:5.1, h:'100%', fill:{color:C.lgray} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.red} });

    s.addText('PROOF OF WORK', {
      x:0.3, y:0.22, w:4.4, h:0.22,
      fontSize:9, bold:true, color:C.red, charSpacing:3, fontFace:'Calibri'
    });
    s.addText('The PhonePe Ops Intelligence Layer — Applied to Tesco Payments', {
      x:0.3, y:0.5, w:9.4, h:0.65,
      fontSize:20, bold:true, color:C.white, fontFace:'Calibri'
    });

    const leftProof = [
      'Rebuilt 6 siloed ops workflows (offer ops, merchant ops, dispute ops) into unified intelligence console with automated breach classification → TAT 2 days → 4 hours, −68% escalation tickets',
      'Deployed Propensity-to-Transact ML with real-time per-user scoring — same baseline anomaly detection principle now applied to PSP auth rates',
      'Built ₹100Cr rewards portfolio ops with health score per brand — same methodology as AuthSense health score per payment method',
      'Owned checkout conversion redesign (22% lift) — deep understanding of payment failure impact on UX and GMV',
    ];
    const rightProof = [
      '→ AuthSense unified console — same architecture, 6 siloed PSP portals → one real-time dashboard with automated anomaly classification',
      '→ AuthSense 3σ anomaly detection — same real-time scoring principle, auth rate deviation instead of purchase intent',
      '→ AuthSense per-method health monitoring — same health score pattern, payment method instead of brand partner',
      '→ AuthSense GMV impact quantification — deep fluency in the business value of every auth recovery percentage point',
    ];

    leftProof.forEach((txt,i) => {
      s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:1.4+i*1.2, w:4.45, h:1.1, fill:{color:C.hero}, rectRadius:0.07, shadow:makeShadow() });
      s.addText(txt, { x:0.4, y:1.46+i*1.2, w:4.15, h:0.98, fontSize:9.5, color:C.muted, fontFace:'Calibri', valign:'middle' });
    });
    rightProof.forEach((txt,i) => {
      s.addShape(pptx.ShapeType.roundRect, { x:5.05, y:1.4+i*1.2, w:4.65, h:1.1, fill:{color:C.white}, rectRadius:0.07, shadow:makeShadow() });
      s.addText(txt, { x:5.2, y:1.46+i*1.2, w:4.35, h:0.98, fontSize:9.5, color:'1E3A5F', fontFace:'Calibri', valign:'middle' });
    });

    s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:6.25, w:9.5, h:0.7, fill:{color:C.red}, rectRadius:0.06 });
    s.addText('The PhonePe ops intelligence layer cut TAT from 2 days to 4 hours and eliminated 68% of escalations. AuthSense applies the same architecture to payment ops — with the same expected outcome.', {
      x:0.4, y:6.3, w:9.2, h:0.6, fontSize:10.5, color:C.white, fontFace:'Calibri', italic:true, valign:'middle'
    });
  }

  // ─── SLIDE 8: ROLLOUT PLAN ─────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.dark} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.gold} });
    s.addText('ROLLOUT PLAN', {
      x:0.3, y:0.22, w:4, h:0.22,
      fontSize:9, bold:true, color:C.gold, charSpacing:3, fontFace:'Calibri'
    });
    s.addText('Three Phases — Read-Only to Full Remediation Ownership', {
      x:0.3, y:0.5, w:9.4, h:0.65,
      fontSize:26, bold:true, color:C.white, fontFace:'Calibri'
    });

    const phases = [
      { n:'01', title:'Phase 1 — Month 1–2: Read-Only Monitoring', body:'Deploy AuthSense in read-only mode alongside existing PSP portals. Build PSP × method × issuer baseline data. Validate 3σ alert thresholds vs historical incidents. Measure detection time vs current lag.', color:C.red },
      { n:'02', title:'Phase 2 — Month 3–4: Live Alerting + Playbooks', body:'Enable real-time alerting with GMV impact estimates. Build 5 incident-type playbooks with ops team. Run A/B: AuthSense-detected incidents vs complaint-detected incidents — measure MTTR delta.', color:C.gold },
      { n:'03', title:'Phase 3 — Month 5–6: Primary Ops Surface', body:'AuthSense becomes primary payment ops monitoring surface, replacing individual PSP portal monitoring. Ops SLA updated to 12-minute detection target. Issuer heatmap data fed into PSP contract review cycle.', color:C.green },
    ];

    phases.forEach((p,i) => {
      const x = 0.25 + i*3.28;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.38, w:3.08, h:4.1, fill:{color:C.hero}, rectRadius:0.1, shadow:makeShadow() });
      s.addShape(pptx.ShapeType.roundRect, { x:x+0.08, y:1.48, w:2.92, h:0.36, fill:{color:p.color}, rectRadius:0.05 });
      s.addText('Phase '+p.n, { x:x+0.08, y:1.48, w:2.92, h:0.36, fontSize:10, bold:true, color:C.dark, align:'center', fontFace:'Calibri' });
      s.addText(p.title, { x:x+0.12, y:1.93, w:2.84, h:0.55, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
      s.addText(p.body, { x:x+0.12, y:2.55, w:2.84, h:2.8, fontSize:10, color:C.muted, fontFace:'Calibri', valign:'top' });
    });

    s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:5.68, w:9.5, h:1.35, fill:{color:'0D2040'}, rectRadius:0.08 });
    s.addText('What I Need to Ship This', { x:0.45, y:5.76, w:4, h:0.28, fontSize:10, bold:true, color:C.gold, fontFace:'Calibri' });
    s.addText([
      '• Access to PSP transaction logs (Worldpay, Adyen, Stripe, Braintree) — read access for baseline model',
      '• Data engineering support to pipe PSP webhooks into AuthSense dashboard (weeks 2–4)',
      '• Ops team engagement to co-build incident playbooks (1 hour workshop per incident type)',
    ].join('\n'), {
      x:0.45, y:6.06, w:9.1, h:0.85,
      fontSize:9.5, color:C.muted, fontFace:'Calibri'
    });
    s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439', {
      x:0, y:7.18, w:'100%', h:0.25,
      fontSize:9, color:C.muted, align:'center', fontFace:'Calibri'
    });
  }

  await pptx.writeFile({ fileName:'assets/tesco-authsense-deck.pptx' });
  console.log('✅ tesco-authsense-deck.pptx saved');
}

buildAuthSense().catch(console.error);
