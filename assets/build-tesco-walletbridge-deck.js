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
  greenl:'DCFCE7',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.13 });

async function buildWalletBridge() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = 'WalletBridge — Tesco Payments Platform';

  // ─── SLIDE 1: COVER ───────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.dark} });
    // Diagonal accent lines
    for(let i=0;i<6;i++){
      s.addShape(pptx.ShapeType.line, { x:-0.5+i*1.8, y:0, w:2.4, h:7.5, line:{color:C.blue, width:0.5, transparency:85} });
    }
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.18, h:'100%', fill:{color:C.blue} });
    s.addText('TESCO TECHNOLOGY · PAYMENTS PLATFORM · CASE STUDY 02', {
      x:0.35, y:0.22, w:9, h:0.25,
      fontSize:8, bold:true, color:C.muted, charSpacing:3, fontFace:'Calibri'
    });
    s.addText('WalletBridge', {
      x:0.35, y:0.7, w:7, h:1.2,
      fontSize:52, bold:true, color:C.white, fontFace:'Calibri'
    });
    s.addText('Self-Serve Payment Method Integration — 6 Weeks to 10 Days', {
      x:0.35, y:1.85, w:7.5, h:0.5,
      fontSize:18, color:C.gold, fontFace:'Calibri', italic:true
    });
    s.addText('Ajay Avaghade · Product Manager', {
      x:0.35, y:2.5, w:5, h:0.3,
      fontSize:11, color:C.muted, fontFace:'Calibri'
    });
    // Big stat bottom-right
    s.addShape(pptx.ShapeType.roundRect, {
      x:7.2, y:5.6, w:2.5, h:1.5,
      fill:{color:C.blue}, rectRadius:0.1, shadow:makeShadow()
    });
    s.addText('−75%', {
      x:7.2, y:5.65, w:2.5, h:0.75,
      fontSize:36, bold:true, color:C.white, align:'center', fontFace:'Calibri'
    });
    s.addText('integration time per\nnew payment method', {
      x:7.2, y:6.32, w:2.5, h:0.55,
      fontSize:9, color:'A8C4E0', align:'center', fontFace:'Calibri'
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
    s.addText('Every New Payment Method Requires a Full Engineering Sprint — Leaving BNPL, New Wallets, and Revenue Sitting in the Backlog', {
      x:0.3, y:0.5, w:9.4, h:1.1,
      fontSize:22, bold:true, color:C.white, fontFace:'Calibri'
    });

    const cols = [
      { icon:'⏱', stat:'6–8 Weeks', lbl:'Per Integration', sub:'Separate auth flows, webhook schemas, and settlement formats for every PSP/wallet' },
      { icon:'💸', stat:'4 Months', lbl:'Klarna in Backlog', sub:'BNPL unlocks +23% conversion at £150+ baskets. Queued for engineering for 4 months' },
      { icon:'🔍', stat:'5 Portals', lbl:'Zero Shared Health View', sub:'Ops monitors each payment method via separate PSP portals. No consolidated status page' },
    ];
    cols.forEach((c2,i) => {
      const x = 0.3 + i*3.25;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.9, w:3.0, h:3.0, fill:{color:C.hero}, rectRadius:0.1, shadow:makeShadow() });
      s.addText(c2.icon, { x, y:2.0, w:3.0, h:0.5, fontSize:24, align:'center' });
      s.addText(c2.stat, { x, y:2.55, w:3.0, h:0.55, fontSize:26, bold:true, color:C.blue, align:'center', fontFace:'Calibri' });
      s.addText(c2.lbl,  { x, y:3.1,  w:3.0, h:0.3,  fontSize:11, bold:true, color:C.white, align:'center', fontFace:'Calibri' });
      s.addText(c2.sub,  { x:x+0.12, y:3.42, w:2.76, h:1.1, fontSize:10, color:C.muted, align:'center', fontFace:'Calibri', valign:'top' });
    });

    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:5.15, w:9.4, h:0.95, fill:{color:'0D2040'}, rectRadius:0.06 });
    s.addText('Root cause: Each payment method is treated as a bespoke engineering project. No abstraction layer. No self-serve. No shared health monitoring. The Payments API team spends weeks on integrations it should take days to ship.', {
      x:0.5, y:5.22, w:9.0, h:0.8,
      fontSize:10, color:C.muted, fontFace:'Calibri', italic:true, valign:'middle'
    });
  }

  // ─── SLIDE 3: THE INSIGHT ──────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.white} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.blue} });
    s.addText('THE INSIGHT', {
      x:0.3, y:0.22, w:4, h:0.22,
      fontSize:9, bold:true, color:C.blue, charSpacing:3, fontFace:'Calibri'
    });
    s.addText('Payment Methods Are Config, Not Code — Treat Them That Way', {
      x:0.3, y:0.5, w:9.4, h:0.9,
      fontSize:26, bold:true, color:C.dark, fontFace:'Calibri'
    });

    // Left column - today
    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:1.65, w:4.5, h:3.8, fill:{color:'F8F9FA'}, rectRadius:0.1, line:{color:'E5E7EB', width:0.5} });
    s.addText('❌  Today', { x:0.5, y:1.8, w:4.1, h:0.35, fontSize:13, bold:true, color:'991B1B', fontFace:'Calibri' });
    const todayItems = [
      'Engineering sprint per method (6–8 weeks)',
      'No sandbox — test in staging or production',
      'PSP health checked via 5 separate portals',
      'New method = queued behind feature work',
      'Auth rate degradation detected via complaints',
    ];
    todayItems.forEach((txt,i) => {
      s.addText('• '+txt, { x:0.5, y:2.25+i*0.55, w:4.1, h:0.45, fontSize:11, color:'374151', fontFace:'Calibri' });
    });

    // VS divider
    s.addShape(pptx.ShapeType.ellipse, { x:4.62, y:3.2, w:0.56, h:0.56, fill:{color:C.blue} });
    s.addText('VS', { x:4.62, y:3.2, w:0.56, h:0.56, fontSize:9, bold:true, color:C.white, align:'center', valign:'middle', fontFace:'Calibri' });

    // Right column - WalletBridge
    s.addShape(pptx.ShapeType.roundRect, { x:5.0, y:1.65, w:4.7, h:3.8, fill:{color:'EFF6FF'}, rectRadius:0.1, line:{color:C.blue, width:0.75} });
    s.addText('✅  WalletBridge', { x:5.15, y:1.8, w:4.3, h:0.35, fontSize:13, bold:true, color:C.blue, fontFace:'Calibri' });
    const wbItems = [
      'Guided wizard: API creds → webhook → sandbox → live',
      'Built-in sandbox — test all 5 scenarios before go-live',
      'One health dashboard — all methods, all PSPs',
      'New method in 10 days. No sprint. No queue.',
      '3σ anomaly detection — 12-min detection lag',
    ];
    wbItems.forEach((txt,i) => {
      s.addText('• '+txt, { x:5.15, y:2.25+i*0.55, w:4.35, h:0.45, fontSize:11, color:'1E3A5F', fontFace:'Calibri' });
    });

    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:5.65, w:9.4, h:0.65, fill:{color:'EFF6FF'}, rectRadius:0.06 });
    s.addText('Key insight: WalletBridge is architecturally identical to the self-serve PG integration platform built at PhonePe — which onboarded 5,000+ B2B merchants with zero engineering dependency per integration.', {
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
    s.addText('Four Steps. Zero Engineering Sprints.', {
      x:0.3, y:0.5, w:9.4, h:0.65,
      fontSize:28, bold:true, color:C.white, fontFace:'Calibri'
    });

    const steps = [
      { n:'1', title:'Method Library', body:'PM or Ops selects payment method from curated library (Klarna, Clearpay, Revolut Pay, etc.). Integration type auto-detected. Estimated completion time shown upfront.' },
      { n:'2', title:'Integration Wizard', body:'4-step guided wizard: API credentials → webhook endpoint config → event mapping → go-live checklist. Each step validates before next unlocks. No PSP documentation required.' },
      { n:'3', title:'Sandbox Gate', body:'Mandatory 5-scenario test suite before any method reaches production: Auth Approved, Soft Decline, Hard Decline, Network Timeout, Refund. 100% test log retained for compliance audit.' },
      { n:'4', title:'Health Monitoring Auto-Enabled', body:'On go-live, the method auto-appears in the health dashboard with per-method alert thresholds. Auth rate, latency, and volume tracked in real-time. 3σ anomaly detection active immediately.' },
    ];

    steps.forEach((st,i) => {
      const x = 0.25 + i*2.42;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.4, w:2.25, h:4.3, fill:{color:C.dark}, rectRadius:0.1, shadow:makeShadow() });
      s.addShape(pptx.ShapeType.ellipse, { x:x+0.78, y:1.52, w:0.68, h:0.68, fill:{color:C.blue} });
      s.addText(st.n, { x:x+0.78, y:1.52, w:0.68, h:0.68, fontSize:16, bold:true, color:C.white, align:'center', valign:'middle', fontFace:'Calibri' });
      s.addText(st.title, { x:x+0.1, y:2.32, w:2.05, h:0.45, fontSize:12, bold:true, color:C.gold, align:'center', fontFace:'Calibri' });
      s.addText(st.body, { x:x+0.1, y:2.82, w:2.05, h:2.5, fontSize:10, color:C.muted, fontFace:'Calibri', valign:'top' });
      if(i<3){
        s.addText('→', { x:x+2.18, y:3.3, w:0.3, h:0.4, fontSize:14, color:C.blue, align:'center', fontFace:'Calibri', bold:true });
      }
    });

    s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:5.88, w:9.5, h:0.65, fill:{color:'0D2040'}, rectRadius:0.06 });
    s.addText('PhonePe proof: Self-serve PG integration platform used this identical 4-step pattern to onboard 5,000+ B2B merchants — 23% lower CAC, zero engineering dependency per merchant.', {
      x:0.45, y:5.93, w:9.1, h:0.55,
      fontSize:10, color:C.muted, fontFace:'Calibri', italic:true, valign:'middle'
    });
  }

  // ─── SLIDE 5: THE PRODUCT ──────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.white} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.blue} });
    s.addText('THE PRODUCT', {
      x:0.3, y:0.22, w:4, h:0.22,
      fontSize:9, bold:true, color:C.blue, charSpacing:3, fontFace:'Calibri'
    });
    s.addText('4 Screens — From Method Discovery to Live Health Monitoring', {
      x:0.3, y:0.5, w:9.4, h:0.65,
      fontSize:24, bold:true, color:C.dark, fontFace:'Calibri'
    });

    const cards = [
      { n:'01', title:'Method Library', desc:'Active/pending/available payment methods in one view. Success rate per method. Real-time status. One-click to start new integration.', mono:'Visa 99.1% ✓\nApple Pay 98.9% ✓\nKlarna BNPL 60% →' },
      { n:'02', title:'Integration Wizard', desc:'Step-locked 4-stage onboarding: credentials → webhooks → sandbox → go-live. Each gate validates before next unlocks. No PSP docs needed.', mono:'Step 1 ✓ Credentials\nStep 2 ✓ Webhooks\nStep 3 → Sandbox' },
      { n:'03', title:'Sandbox Environment', desc:'Test any method × any scenario (5 types). Uses production API in test mode. Full transaction log as compliance audit trail.', mono:'Klarna · Approved ✓\nKlarna · Hard Dec. ✗\nVisa · Approved ✓' },
      { n:'04', title:'Health Dashboard', desc:'Live auth rate, p95 latency, volume per method. Per-method alert thresholds. 3σ anomaly detection auto-enabled on go-live.', mono:'Visa   99.1% ● ok\nAPay  98.9% ● ok\nKlarna 97.0% ● ok' },
    ];

    cards.forEach((c2,i) => {
      const x = 0.25 + i*2.44;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.38, w:2.28, h:4.95, fill:{color:C.lgray}, rectRadius:0.1, line:{color:'D0D8E8', width:0.5}, shadow:makeShadow() });
      s.addShape(pptx.ShapeType.roundRect, { x:x+0.06, y:1.45, w:2.16, h:0.3, fill:{color:C.blue}, rectRadius:0.05 });
      s.addText(c2.n+' · '+c2.title, { x:x+0.06, y:1.45, w:2.16, h:0.3, fontSize:9, bold:true, color:C.white, align:'center', fontFace:'Calibri' });
      s.addText(c2.desc, { x:x+0.1, y:1.85, w:2.08, h:1.6, fontSize:10, color:'374151', fontFace:'Calibri', valign:'top' });
      s.addShape(pptx.ShapeType.roundRect, { x:x+0.1, y:3.5, w:2.08, h:1.0, fill:{color:C.dark}, rectRadius:0.06 });
      s.addText(c2.mono, { x:x+0.14, y:3.55, w:2.0, h:0.9, fontSize:9, color:'7DD3A8', fontFace:'Courier New', valign:'top' });
    });

    s.addText('Prototype: tesco-walletbridge-prototype.html · All 4 screens interactive', {
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
    s.addText('Faster Integration. More Methods. Better Payment Health.', {
      x:0.3, y:0.5, w:9.4, h:0.65,
      fontSize:24, bold:true, color:C.white, fontFace:'Calibri'
    });

    const leftMetrics = [
      { v:'−75%', l:'Integration time', s:'6–8 weeks → 10 days per new payment method' },
      { v:'+4', l:'New methods/year', s:'Without adding engineering headcount' },
      { v:'99.2%', l:'Integration SLA target', s:'Per active method, monitored in real-time' },
      { v:'0', l:'Regressions on go-live', s:'5-scenario sandbox gate prevents all preventable failures' },
    ];
    const rightMetrics = [
      { v:'23%', l:'BNPL conversion uplift', s:'At £150+ basket — Klarna live in 10 days, not 4 months' },
      { v:'−60%', l:'False positive alerts', s:'Per-method thresholds vs global aggregate monitoring' },
      { v:'<5 min', l:'From alert to root cause', s:'3-dimension breakdown: PSP × method × issuer' },
      { v:'5 portals → 1', l:'Ops monitoring consolidation', s:'Single health dashboard replaces all PSP portal logins' },
    ];

    [leftMetrics, rightMetrics].forEach((metrics, col) => {
      const xBase = col===0 ? 0.3 : 5.1;
      s.addText(col===0 ? 'Integration Velocity' : 'Payment Health', {
        x:xBase, y:1.35, w:4.5, h:0.28,
        fontSize:10, bold:true, color:col===0?C.blue:C.gold, charSpacing:2, fontFace:'Calibri'
      });
      metrics.forEach((m,i) => {
        const y = 1.72 + i*0.92;
        s.addShape(pptx.ShapeType.roundRect, { x:xBase, y, w:4.55, h:0.82, fill:{color:C.hero}, rectRadius:0.06, shadow:makeShadow() });
        s.addText(m.v, { x:xBase+0.12, y:y+0.06, w:1.6, h:0.5, fontSize:22, bold:true, color:col===0?C.blue:C.gold, fontFace:'Calibri', valign:'middle' });
        s.addText(m.l, { x:xBase+1.7, y:y+0.06, w:2.7, h:0.28, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
        s.addText(m.s, { x:xBase+1.7, y:y+0.36, w:2.7, h:0.36, fontSize:9, color:C.muted, fontFace:'Calibri', valign:'top' });
      });
    });
    s.addShape(pptx.ShapeType.rect, { x:4.88, y:1.35, w:0.02, h:5.35, fill:{color:'1E3A5F'} });

    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:6.5, w:9.4, h:0.75, fill:{color:'0D2040'}, rectRadius:0.06 });
    s.addText('WalletBridge turns payment method integration from a recurring engineering cost into a once-built platform. Every future method is 10 days and zero engineering dependency.', {
      x:0.5, y:6.56, w:9.0, h:0.6, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true, valign:'middle'
    });
  }

  // ─── SLIDE 7: PROOF OF WORK ────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:4.9, h:'100%', fill:{color:C.dark} });
    s.addShape(pptx.ShapeType.rect, { x:4.9, y:0, w:5.1, h:'100%', fill:{color:C.lgray} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.blue} });

    s.addText('PROOF OF WORK', {
      x:0.3, y:0.22, w:4.4, h:0.22,
      fontSize:9, bold:true, color:C.blue, charSpacing:3, fontFace:'Calibri'
    });
    s.addText('I Built This Architecture at PhonePe', {
      x:0.3, y:0.5, w:9.4, h:0.65,
      fontSize:22, bold:true, color:C.white, fontFace:'Calibri'
    });

    const leftProof = [
      'Built zero-to-one self-serve PG integration platform with guided merchant onboarding wizard → 5,000+ B2B merchants acquired, 23% lower CAC',
      'Platform abstracted Instant Discount + EMI subvention config into self-serve steps → zero engineering dependency per merchant integration',
      'Rebuilt 6 siloed ops workflows into unified intelligence console → TAT 2 days → 4 hours, −68% escalation tickets',
      'Deployed Propensity-to-Transact ML with real-time scoring layer → 32% marketing burn reduction, sustained GMV across all markets',
    ];
    const rightProof = [
      '→ WalletBridge wizard — same self-serve onboarding architecture, payment methods instead of merchant capabilities',
      '→ WalletBridge abstraction layer — PSP config, webhook normalisation, error mapping into guided config steps',
      '→ WalletBridge health dashboard — same consolidation architecture, payment health instead of ops workflows',
      '→ WalletBridge 3σ anomaly detection — same real-time scoring principle, auth rate instead of propensity',
    ];

    leftProof.forEach((txt,i) => {
      s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:1.4+i*1.2, w:4.45, h:1.1, fill:{color:C.hero}, rectRadius:0.07, shadow:makeShadow() });
      s.addText(txt, { x:0.4, y:1.46+i*1.2, w:4.15, h:0.98, fontSize:9.5, color:C.muted, fontFace:'Calibri', valign:'middle' });
    });
    rightProof.forEach((txt,i) => {
      s.addShape(pptx.ShapeType.roundRect, { x:5.05, y:1.4+i*1.2, w:4.65, h:1.1, fill:{color:C.white}, rectRadius:0.07, shadow:makeShadow() });
      s.addText(txt, { x:5.2, y:1.46+i*1.2, w:4.35, h:0.98, fontSize:9.5, color:'1E3A5F', fontFace:'Calibri', valign:'middle' });
    });

    s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:6.25, w:9.5, h:0.7, fill:{color:C.blue}, rectRadius:0.06 });
    s.addText('This is not a conceptual case study. WalletBridge is the exact same platform architecture I shipped at PhonePe — applied to Tesco\'s payment method problem.', {
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
    s.addText('Three Phases — Low Risk Migration to Full Platform', {
      x:0.3, y:0.5, w:9.4, h:0.65,
      fontSize:26, bold:true, color:C.white, fontFace:'Calibri'
    });

    const phases = [
      { n:'01', title:'Phase 1 — Month 1–2: Pilot (Klarna BNPL)', body:'Complete Klarna integration as first WalletBridge-managed method. Measure integration time vs historical baseline. Validate sandbox gate catches all real-world edge cases. Deploy health monitoring and tune alert thresholds.', color:C.blue },
      { n:'02', title:'Phase 2 — Month 3–4: Platform & 2 More Methods', body:'Onboard 2 additional methods via wizard (Clearpay + Revolut Pay). Confirm sub-10-day integration time. A/B test per-method alert thresholds vs global baseline for false positive reduction. Publish integration time SLA.', color:C.gold },
      { n:'03', title:'Phase 3 — Month 5–6: Standard Onboarding Path', body:'WalletBridge becomes the default path for all future payment method integrations. Engineering no longer assigned to integration sprints. PM team owns method onboarding end to end. Health dashboard becomes primary ops monitoring surface.', color:C.green },
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
      '• Access to Tesco Payments API documentation and PSP contracts (Worldpay, Adyen, Stripe)',
      '• 1 backend engineer for webhook abstraction layer (weeks 2–6)',
      '• Alignment with PSP account managers on test credentials for sandbox environment',
    ].join('\n'), {
      x:0.45, y:6.06, w:9.1, h:0.85,
      fontSize:9.5, color:C.muted, fontFace:'Calibri'
    });
    s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439', {
      x:0, y:7.18, w:'100%', h:0.25,
      fontSize:9, color:C.muted, align:'center', fontFace:'Calibri'
    });
  }

  await pptx.writeFile({ fileName:'assets/tesco-walletbridge-deck.pptx' });
  console.log('✅ tesco-walletbridge-deck.pptx saved');
}

buildWalletBridge().catch(console.error);
