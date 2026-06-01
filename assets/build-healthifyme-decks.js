const PptxGenJS = require('pptxgenjs');

const C = {
  dark:   '051A0A',
  hero:   '0A2810',
  green:  '00C853',
  green2: '00E676',
  gold:   'FFB300',
  orange: 'FF6D00',
  red:    'F44336',
  white:  'FFFFFF',
  lgray:  'F0FFF4',
  muted:  '6B8A72',
  oranl:  'FBE9E7',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.13 });

// ─────────────────────────────────────────────────────────────────────────────
// DECK 1: UpgradeIQ
// ─────────────────────────────────────────────────────────────────────────────
async function buildUpgradeIQ() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';

  // SLIDE 1: COVER
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.dark} });
    for(let i=0;i<6;i++) s.addShape(pptx.ShapeType.line, { x:-0.5+i*1.8, y:0, w:2.4, h:7.5, line:{color:C.green, width:0.5, transparency:88} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.18, h:'100%', fill:{color:C.green} });
    s.addText('HEALTHIFYME · B2C AI PRODUCTS · CASE STUDY 01', { x:0.35, y:0.22, w:9, h:0.25, fontSize:8, bold:true, color:C.muted, charSpacing:3, fontFace:'Calibri' });
    s.addText('UpgradeIQ', { x:0.35, y:0.7, w:7, h:1.2, fontSize:52, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText('Subscription Upgrade Propensity Engine — Show the Right Offer at Peak Intent', { x:0.35, y:1.88, w:8, h:0.55, fontSize:16, color:C.gold, fontFace:'Calibri', italic:true });
    s.addText('Ajay Avaghade · Product Manager', { x:0.35, y:2.55, w:5, h:0.3, fontSize:11, color:C.muted, fontFace:'Calibri' });
    s.addShape(pptx.ShapeType.roundRect, { x:7.1, y:5.5, w:2.65, h:1.6, fill:{color:C.green}, rectRadius:0.1, shadow:makeShadow() });
    s.addText('+18pp', { x:7.1, y:5.56, w:2.65, h:0.75, fontSize:36, bold:true, color:C.dark, align:'center', fontFace:'Calibri' });
    s.addText('CVR lift vs\ngeneric banner', { x:7.1, y:6.25, w:2.65, h:0.65, fontSize:9, color:'004020', align:'center', fontFace:'Calibri' });
  }

  // SLIDE 2: THE PROBLEM
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.dark} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.red} });
    s.addText('THE PROBLEM', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.red, charSpacing:3, fontFace:'Calibri' });
    s.addText('HealthifyMe Shows "Upgrade to PRO" to Everyone — Including Users Who Have No Reason to Upgrade Yet', { x:0.3, y:0.5, w:9.4, h:0.9, fontSize:22, bold:true, color:C.white, fontFace:'Calibri' });
    const cols = [
      { icon:'📢', stat:'16%', lbl:'Generic Banner CVR', sub:'The current "Upgrade to PRO" banner converts at 16% regardless of user state, goal proximity, or engagement level' },
      { icon:'🎯', stat:'0%', lbl:'Propensity Signals Used', sub:'Goal proximity, streak length, engagement velocity — all available, none used to determine when to show the upgrade prompt' },
      { icon:'💸', stat:'−34%', lbl:'Missed CVR vs Potential', sub:'Users at goal proximity <1kg convert at 34% with a contextual offer. Showing them a generic banner wastes the highest-value moment' },
    ];
    cols.forEach((c2,i) => {
      const x = 0.3 + i*3.25;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.75, w:3.0, h:3.15, fill:{color:C.hero}, rectRadius:0.1, shadow:makeShadow() });
      s.addText(c2.icon, { x, y:1.85, w:3.0, h:0.5, fontSize:24, align:'center' });
      s.addText(c2.stat, { x, y:2.4, w:3.0, h:0.55, fontSize:26, bold:true, color:C.red, align:'center', fontFace:'Calibri' });
      s.addText(c2.lbl, { x, y:2.95, w:3.0, h:0.3, fontSize:11, bold:true, color:C.white, align:'center', fontFace:'Calibri' });
      s.addText(c2.sub, { x:x+0.12, y:3.3, w:2.76, h:1.28, fontSize:10, color:C.muted, align:'center', fontFace:'Calibri', valign:'top' });
    });
    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:5.1, w:9.4, h:0.95, fill:{color:'101A10'}, rectRadius:0.06 });
    s.addText('Root cause: Every free user gets the same upgrade prompt at the same moment. A user 2kg from their goal and on a 7-day streak has 5× the upgrade intent of a day-3 user — and gets the same banner. The opportunity cost is enormous.', { x:0.5, y:5.17, w:9.0, h:0.8, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true, valign:'middle' });
  }

  // SLIDE 3: THE INSIGHT
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.white} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.green} });
    s.addText('THE INSIGHT', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.green, charSpacing:3, fontFace:'Calibri' });
    s.addText('The Highest-Intent Upgrade Moment Is When the User Is Closest to Their Goal — Not When They First Open the App', { x:0.3, y:0.5, w:9.4, h:0.9, fontSize:21, bold:true, color:C.dark, fontFace:'Calibri' });
    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:1.65, w:4.5, h:3.8, fill:{color:'FFF8F8'}, rectRadius:0.1, line:{color:'FFCDD2', width:0.75} });
    s.addText('❌  Today', { x:0.5, y:1.8, w:4.1, h:0.35, fontSize:13, bold:true, color:C.red, fontFace:'Calibri' });
    ['Generic banner to all free users at all times', 'No goal proximity signal used', 'No streak / engagement velocity signal', 'Same value prop regardless of user state', 'Discount trained users — waits for offers'].forEach((t,i) => {
      s.addText('• '+t, { x:0.5, y:2.26+i*0.57, w:4.1, h:0.45, fontSize:11, color:'374151', fontFace:'Calibri' });
    });
    s.addShape(pptx.ShapeType.ellipse, { x:4.62, y:3.2, w:0.56, h:0.56, fill:{color:C.green} });
    s.addText('VS', { x:4.62, y:3.2, w:0.56, h:0.56, fontSize:9, bold:true, color:C.dark, align:'center', valign:'middle', fontFace:'Calibri' });
    s.addShape(pptx.ShapeType.roundRect, { x:5.0, y:1.65, w:4.7, h:3.8, fill:{color:C.lgray}, rectRadius:0.1, line:{color:C.green, width:0.75} });
    s.addText('✅  UpgradeIQ', { x:5.15, y:1.8, w:4.3, h:0.35, fontSize:13, bold:true, color:C.green, fontFace:'Calibri' });
    ['Goal proximity score computed per user in real-time', 'HIGH tier (<1kg from goal): show at 34% CVR moment', 'LOW tier (3kg+ from goal): suppress — not the right time', 'Value prop personalised: "23 days faster at your pace"', 'No discount required — goal motivation converts'].forEach((t,i) => {
      s.addText('• '+t, { x:5.15, y:2.26+i*0.57, w:4.35, h:0.45, fontSize:11, color:'1E3A2F', fontFace:'Calibri' });
    });
    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:5.65, w:9.4, h:0.65, fill:{color:C.lgray}, rectRadius:0.06 });
    s.addText('Key insight: Suppressing the upgrade prompt for LOW propensity users preserves their future upgrade intent. Showing it too early trains them to dismiss it — permanently lowering their conversion probability.', { x:0.5, y:5.7, w:9.0, h:0.55, fontSize:10, color:C.green, fontFace:'Calibri', italic:true, valign:'middle' });
  }

  // SLIDE 4: THE MECHANIC
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.hero} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.gold} });
    s.addText('THE MECHANIC', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.gold, charSpacing:3, fontFace:'Calibri' });
    s.addText('Score Every User Every Session. Show the Offer Only When the Score Says Now.', { x:0.3, y:0.5, w:9.4, h:0.65, fontSize:24, bold:true, color:C.white, fontFace:'Calibri' });
    const steps = [
      { n:'1', title:'Signal Collection', body:'Per-session: goal proximity (kg remaining), streak length, engagement velocity (logins/wk), feature demand (meal plan views, coach taps), historical upgrade cohort patterns.' },
      { n:'2', title:'Propensity Scoring', body:'ML model outputs HIGH / MED / LOW tier per user per session. HIGH = <1kg from goal + 5+ day streak. MED = 1–3kg. LOW = 3kg+ or low engagement. Score recomputed every login.' },
      { n:'3', title:'Trigger Decision', body:'HIGH tier: show UpgradeIQ card with personalised value frame ("23 days faster at your pace"). MED: show feature preview. LOW: suppress entirely — no upgrade prompt shown.' },
      { n:'4', title:'Personalised Frame', body:'Value proposition computed from user\'s actual data: "PRO users at YOUR streak + pace hit this goal 23 days sooner" is more compelling than any generic claim. No discount required.' },
    ];
    steps.forEach((st,i) => {
      const x = 0.25 + i*2.42;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.4, w:2.25, h:4.3, fill:{color:C.dark}, rectRadius:0.1, shadow:makeShadow() });
      s.addShape(pptx.ShapeType.ellipse, { x:x+0.78, y:1.52, w:0.68, h:0.68, fill:{color:C.green} });
      s.addText(st.n, { x:x+0.78, y:1.52, w:0.68, h:0.68, fontSize:16, bold:true, color:C.dark, align:'center', valign:'middle', fontFace:'Calibri' });
      s.addText(st.title, { x:x+0.1, y:2.32, w:2.05, h:0.45, fontSize:12, bold:true, color:C.gold, align:'center', fontFace:'Calibri' });
      s.addText(st.body, { x:x+0.1, y:2.82, w:2.05, h:2.5, fontSize:10, color:C.muted, fontFace:'Calibri', valign:'top' });
      if(i<3) s.addText('→', { x:x+2.18, y:3.3, w:0.3, h:0.4, fontSize:14, color:C.green, align:'center', fontFace:'Calibri', bold:true });
    });
    s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:5.88, w:9.5, h:0.65, fill:{color:'030E05'}, rectRadius:0.06 });
    s.addText('PhonePe proof: Propensity-to-Transact ML replacing manual cohorts with real-time per-user scoring → 32% marketing burn reduction while sustaining GMV. Identical architecture applied to subscription upgrade timing.', { x:0.45, y:5.93, w:9.1, h:0.55, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true, valign:'middle' });
  }

  // SLIDE 5: THE PRODUCT
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.white} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.green} });
    s.addText('THE PRODUCT', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.green, charSpacing:3, fontFace:'Calibri' });
    s.addText('4 Screens — Free User Home to PRO Activation with Confetti', { x:0.3, y:0.5, w:9.4, h:0.65, fontSize:24, bold:true, color:C.dark, fontFace:'Calibri' });
    const cards = [
      { n:'01', title:'Free User Home', desc:'UpgradeIQ card fires when goal proximity score = HIGH. Shows personalised stats: "2.1kg from goal, 7-day streak." Goal progress bar + streak badge visible.', mono:'Goal: 2.9/5kg ■\n7-day streak 🔥\nUpgradeIQ → PRO\n[23 days faster]' },
      { n:'02', title:'Upgrade Detail', desc:'Personalised value frame: "PRO users at your pace hit it 23 days faster." Social proof: "94 users with your pattern upgraded this week." Feature-by-feature goal unlocking comparison.', mono:'You: −0.3kg/wk\nFree: Aug 28\nPRO: Aug 5\n△ 23 days' },
      { n:'03', title:'Subscription Checkout', desc:'Annual plan pre-selected (Best Value). Milestone badge commitment device: "Earn your 5kg Challenge badge on hit." Confetti on PRO activation.', mono:'Monthly ₹999\nQuarterly ₹799\nAnnual ₹499 ★\n[Start PRO Now]' },
      { n:'04', title:'Growth Analytics', desc:'Propensity tier CVR breakdown: HIGH 34%, MED 18%, LOW 6%. A/B test: contextual trigger 34% vs generic banner 16% vs discount 21%.', mono:'HIGH  34% ↑\nMED   18% →\nLOW    6% ↓\nA/B ◆ +18pp' },
    ];
    cards.forEach((c2,i) => {
      const x = 0.25 + i*2.44;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.38, w:2.28, h:4.95, fill:{color:C.lgray}, rectRadius:0.1, line:{color:'C8E6C9', width:0.5}, shadow:makeShadow() });
      s.addShape(pptx.ShapeType.roundRect, { x:x+0.06, y:1.45, w:2.16, h:0.3, fill:{color:C.green}, rectRadius:0.05 });
      s.addText(c2.n+' · '+c2.title, { x:x+0.06, y:1.45, w:2.16, h:0.3, fontSize:9, bold:true, color:C.dark, align:'center', fontFace:'Calibri' });
      s.addText(c2.desc, { x:x+0.1, y:1.85, w:2.08, h:1.6, fontSize:10, color:'374151', fontFace:'Calibri', valign:'top' });
      s.addShape(pptx.ShapeType.roundRect, { x:x+0.1, y:3.5, w:2.08, h:1.0, fill:{color:C.dark}, rectRadius:0.06 });
      s.addText(c2.mono, { x:x+0.14, y:3.55, w:2.0, h:0.9, fontSize:9, color:'7DD3A8', fontFace:'Courier New', valign:'top' });
    });
    s.addText('Prototype: healthifyme-upgradeiq-prototype.html · All 4 screens interactive · Confetti on PRO activation', { x:0.3, y:6.68, w:9.4, h:0.3, fontSize:9, color:C.muted, italic:true, fontFace:'Calibri', align:'center' });
  }

  // SLIDE 6: IMPACT & ROI
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.dark} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.green} });
    s.addText('IMPACT & ROI', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.green, charSpacing:3, fontFace:'Calibri' });
    s.addText('34% Conversion at Peak Intent — Without a Single Rupee of Discount', { x:0.3, y:0.5, w:9.4, h:0.65, fontSize:24, bold:true, color:C.white, fontFace:'Calibri' });
    const lm = [
      { v:'34%', l:'CVR at HIGH tier', s:'Goal proximity <1kg + 5+ day streak' },
      { v:'+18pp', l:'Lift vs generic banner', s:'34% contextual vs 16% blanket prompt' },
      { v:'5.7×', l:'HIGH vs LOW tier CVR', s:'Same user, right moment = 5.7× conversion rate' },
      { v:'₹2.4L', l:'MRR added (30 days)', s:'From propensity-triggered upgrades alone' },
    ];
    const rm = [
      { v:'₹0', l:'Discount cost', s:'Goal-framed value prop converts without margin erosion' },
      { v:'−32%', l:'Marketing burn (PhonePe proof)', s:'Same propensity-to-transact architecture at 350M+ MAU' },
      { v:'60%', l:'Goal completion improvement', s:'Milestone mechanic — PhonePe device activation A/B proof' },
      { v:'84%', l:'PRO users hit goal (vs 62% free)', s:'Goal completion drives NPS and renewal, not just acquisition' },
    ];
    [lm,rm].forEach((metrics,col) => {
      const xBase = col===0 ? 0.3 : 5.1;
      s.addText(col===0 ? 'Conversion Metrics' : 'Business ROI', { x:xBase, y:1.35, w:4.5, h:0.28, fontSize:10, bold:true, color:col===0?C.green:C.gold, charSpacing:2, fontFace:'Calibri' });
      metrics.forEach((m,i) => {
        const y = 1.72 + i*0.92;
        s.addShape(pptx.ShapeType.roundRect, { x:xBase, y, w:4.55, h:0.82, fill:{color:C.hero}, rectRadius:0.06, shadow:makeShadow() });
        s.addText(m.v, { x:xBase+0.12, y:y+0.06, w:1.6, h:0.5, fontSize:21, bold:true, color:col===0?C.green:C.gold, fontFace:'Calibri', valign:'middle' });
        s.addText(m.l, { x:xBase+1.7, y:y+0.06, w:2.7, h:0.28, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
        s.addText(m.s, { x:xBase+1.7, y:y+0.36, w:2.7, h:0.36, fontSize:9, color:C.muted, fontFace:'Calibri', valign:'top' });
      });
    });
    s.addShape(pptx.ShapeType.rect, { x:4.88, y:1.35, w:0.02, h:5.35, fill:{color:'1A3020'} });
    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:6.5, w:9.4, h:0.75, fill:{color:'030E05'}, rectRadius:0.06 });
    s.addText('UpgradeIQ recovers the conversion rate that was always available — from users who wanted to upgrade but were never shown the offer at the right moment. No additional traffic. No discounts. Just timing.', { x:0.5, y:6.56, w:9.0, h:0.6, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true, valign:'middle' });
  }

  // SLIDE 7: PROOF OF WORK
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:4.9, h:'100%', fill:{color:C.dark} });
    s.addShape(pptx.ShapeType.rect, { x:4.9, y:0, w:5.1, h:'100%', fill:{color:C.lgray} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.green} });
    s.addText('PROOF OF WORK', { x:0.3, y:0.22, w:4.4, h:0.22, fontSize:9, bold:true, color:C.green, charSpacing:3, fontFace:'Calibri' });
    s.addText('I Built This Architecture at PhonePe — for Payments. Here It Is for Health.', { x:0.3, y:0.5, w:9.4, h:0.65, fontSize:20, bold:true, color:C.white, fontFace:'Calibri' });
    const lp = [
      'Led deployment of Propensity-to-Transact ML models replacing manual cohorts with real-time user-level scoring → 32% marketing burn reduction while sustaining GMV across all markets',
      'Built dynamic cart incentivization engine with context-aware triggers (cart value × user intent × time-of-day signals) → 35% AOV uplift, 60% cart abandonment reduction',
      'A/B tested milestone waiver vs flat discount — 60% improvement in 90-day device activation vs discount control. Zero net cost mechanic.',
      'Redesigned offers placement + eligibility surfacing (22% checkout conversion lift) — deep fluency in showing the right offer at the right moment to the right user.',
    ];
    const rp = [
      '→ UpgradeIQ propensity scoring — same real-time per-user ML scoring, subscription upgrade intent instead of purchase propensity',
      '→ UpgradeIQ trigger logic — same context-aware trigger stack: goal proximity × streak × time-of-day → subscription upgrade moment',
      '→ UpgradeIQ milestone badge — same milestone mechanic, goal achievement badge instead of rental waiver. Zero cost, high psychological impact.',
      '→ UpgradeIQ personalised value frame — same principle: show the right value prop at the right moment, not a generic banner.',
    ];
    lp.forEach((txt,i) => {
      s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:1.4+i*1.2, w:4.45, h:1.1, fill:{color:C.hero}, rectRadius:0.07, shadow:makeShadow() });
      s.addText(txt, { x:0.4, y:1.46+i*1.2, w:4.15, h:0.98, fontSize:9.5, color:C.muted, fontFace:'Calibri', valign:'middle' });
    });
    rp.forEach((txt,i) => {
      s.addShape(pptx.ShapeType.roundRect, { x:5.05, y:1.4+i*1.2, w:4.65, h:1.1, fill:{color:C.white}, rectRadius:0.07, shadow:makeShadow() });
      s.addText(txt, { x:5.2, y:1.46+i*1.2, w:4.35, h:0.98, fontSize:9.5, color:'1E3A2F', fontFace:'Calibri', valign:'middle' });
    });
    s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:6.25, w:9.5, h:0.7, fill:{color:C.green}, rectRadius:0.06 });
    s.addText('UpgradeIQ is the Propensity-to-Transact model I shipped at PhonePe — repointed from purchase intent to subscription upgrade intent. The architecture is identical. The proof exists.', { x:0.4, y:6.3, w:9.2, h:0.6, fontSize:10.5, color:C.dark, fontFace:'Calibri', italic:true, valign:'middle' });
  }

  // SLIDE 8: ROLLOUT PLAN
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.dark} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.gold} });
    s.addText('ROLLOUT PLAN', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.gold, charSpacing:3, fontFace:'Calibri' });
    s.addText('Three Phases — Baseline Model to Full Propensity-Triggered Upgrade Engine', { x:0.3, y:0.5, w:9.4, h:0.65, fontSize:23, bold:true, color:C.white, fontFace:'Calibri' });
    const phases = [
      { n:'01', title:'Phase 1 — Month 1–2: Signal Audit & Baseline', body:'Audit existing user signals: goal proximity, streak data, session frequency, feature interactions. Build propensity scoring baseline from historical upgrade cohort. Validate signal → conversion correlation before building ML.', color:C.green },
      { n:'02', title:'Phase 2 — Month 3–4: A/B Test Core Mechanic', body:'Deploy UpgradeIQ card for HIGH tier (50% test, 50% control gets generic banner). Measure CVR delta by propensity tier. Tune model thresholds. Build personalised value frame copy variants (pace-based vs social proof vs feature unlock).', color:C.gold },
      { n:'03', title:'Phase 3 — Month 5–6: Full Tier System Live', body:'Deploy full HIGH / MED / LOW suppression logic. Integrate milestone badge into checkout. Build ops analytics dashboard (CVR by tier, A/B tracker). Launch annual plan default selection with personalised goal date computation.', color:C.orange },
    ];
    phases.forEach((p,i) => {
      const x = 0.25 + i*3.28;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.38, w:3.08, h:4.1, fill:{color:C.hero}, rectRadius:0.1, shadow:makeShadow() });
      s.addShape(pptx.ShapeType.roundRect, { x:x+0.08, y:1.48, w:2.92, h:0.36, fill:{color:p.color}, rectRadius:0.05 });
      s.addText('Phase '+p.n, { x:x+0.08, y:1.48, w:2.92, h:0.36, fontSize:10, bold:true, color:C.dark, align:'center', fontFace:'Calibri' });
      s.addText(p.title, { x:x+0.12, y:1.93, w:2.84, h:0.55, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
      s.addText(p.body, { x:x+0.12, y:2.55, w:2.84, h:2.8, fontSize:10, color:C.muted, fontFace:'Calibri', valign:'top' });
    });
    s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:5.68, w:9.5, h:1.35, fill:{color:'030E05'}, rectRadius:0.08 });
    s.addText('What I Need', { x:0.45, y:5.76, w:4, h:0.28, fontSize:10, bold:true, color:C.gold, fontFace:'Calibri' });
    s.addText('• User session data: goal progress history, streak logs, feature interaction events\n• ML infra to serve per-user propensity score at session load (latency <100ms)\n• Product analytics: A/B framework to measure CVR by propensity tier and upgrade moment', { x:0.45, y:6.06, w:9.1, h:0.85, fontSize:9.5, color:C.muted, fontFace:'Calibri' });
    s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439', { x:0, y:7.18, w:'100%', h:0.25, fontSize:9, color:C.muted, align:'center', fontFace:'Calibri' });
  }

  await pptx.writeFile({ fileName:'assets/healthifyme-upgradeiq-deck.pptx' });
  console.log('✅ healthifyme-upgradeiq-deck.pptx saved');
}

// ─────────────────────────────────────────────────────────────────────────────
// DECK 2: RetentionOS
// ─────────────────────────────────────────────────────────────────────────────
async function buildRetentionOS() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';

  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.dark} });
    for(let i=0;i<6;i++) s.addShape(pptx.ShapeType.line, { x:-0.5+i*1.8, y:0, w:2.4, h:7.5, line:{color:C.red, width:0.5, transparency:88} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.18, h:'100%', fill:{color:C.red} });
    s.addText('HEALTHIFYME · B2C AI PRODUCTS · CASE STUDY 02', { x:0.35, y:0.22, w:9, h:0.25, fontSize:8, bold:true, color:C.muted, charSpacing:3, fontFace:'Calibri' });
    s.addText('RetentionOS', { x:0.35, y:0.7, w:7.5, h:1.2, fontSize:48, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText('PRO Subscriber Churn Prediction — Intervene 22 Days Before Cancellation', { x:0.35, y:1.88, w:8.5, h:0.55, fontSize:16, color:C.gold, fontFace:'Calibri', italic:true });
    s.addText('Ajay Avaghade · Product Manager', { x:0.35, y:2.55, w:5, h:0.3, fontSize:11, color:C.muted, fontFace:'Calibri' });
    s.addShape(pptx.ShapeType.roundRect, { x:7.1, y:5.5, w:2.65, h:1.6, fill:{color:C.red}, rectRadius:0.1, shadow:makeShadow() });
    s.addText('+12.9pp', { x:7.1, y:5.56, w:2.65, h:0.75, fontSize:34, bold:true, color:C.white, align:'center', fontFace:'Calibri' });
    s.addText('renewal rate lift\nintervention cohort', { x:7.1, y:6.25, w:2.65, h:0.65, fontSize:9, color:'FFAAAA', align:'center', fontFace:'Calibri' });
  }
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.dark} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.red} });
    s.addText('THE PROBLEM', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.red, charSpacing:3, fontFace:'Calibri' });
    s.addText('HealthifyMe Discovers PRO Subscriber Churn When the Renewal Fails — Not 3 Weeks Before When an Intervention Could Still Work', { x:0.3, y:0.5, w:9.4, h:0.9, fontSize:21, bold:true, color:C.white, fontFace:'Calibri' });
    const cols = [
      { icon:'⏳', stat:'Day 30', lbl:'When Churn Is Detected', sub:'Churn is detected at the renewal event — the cancellation. By then the user\'s motivation has been gone for 2–3 weeks.' },
      { icon:'📉', stat:'9.4%', lbl:'Subscribers At Risk', sub:'1,204 of 12,847 active PRO subscribers show declining engagement signals — pre-churn patterns that are observable today.' },
      { icon:'💸', stat:'₹18.2L', lbl:'Monthly Revenue at Stake', sub:'At-risk subscriber × avg annual plan value. Previously invisible — no product surfaces this number in real time.' },
    ];
    cols.forEach((c2,i) => {
      const x = 0.3 + i*3.25;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.72, w:3.0, h:3.2, fill:{color:C.hero}, rectRadius:0.1, shadow:makeShadow() });
      s.addText(c2.icon, { x, y:1.82, w:3.0, h:0.5, fontSize:24, align:'center' });
      s.addText(c2.stat, { x, y:2.37, w:3.0, h:0.55, fontSize:26, bold:true, color:C.red, align:'center', fontFace:'Calibri' });
      s.addText(c2.lbl, { x, y:2.92, w:3.0, h:0.3, fontSize:11, bold:true, color:C.white, align:'center', fontFace:'Calibri' });
      s.addText(c2.sub, { x:x+0.12, y:3.27, w:2.76, h:1.35, fontSize:10, color:C.muted, align:'center', fontFace:'Calibri', valign:'top' });
    });
    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:5.12, w:9.4, h:0.95, fill:{color:'101005'}, rectRadius:0.06 });
    s.addText('The data to predict churn 22 days early is already collected: login frequency, meal log rate, streak breaks, goal stalls. RetentionOS is the product layer that turns this data into revenue-saving interventions.', { x:0.5, y:5.19, w:9.0, h:0.8, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true, valign:'middle' });
  }
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.white} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.red} });
    s.addText('THE INSIGHT', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.red, charSpacing:3, fontFace:'Calibri' });
    s.addText('Engagement Drop Precedes Cancellation by 14 Days. Intervene at Day 8. Not Day 30.', { x:0.3, y:0.5, w:9.4, h:0.9, fontSize:23, bold:true, color:C.dark, fontFace:'Calibri' });
    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:1.65, w:4.5, h:3.8, fill:{color:'FFF8F8'}, rectRadius:0.1, line:{color:'FFCDD2', width:0.75} });
    s.addText('❌  Today', { x:0.5, y:1.8, w:4.1, h:0.35, fontSize:13, bold:true, color:C.red, fontFace:'Calibri' });
    ['Churn detected at renewal event (day 30)', 'No pre-churn engagement signals tracked', 'Win-back attempted after cancellation', 'Generic "we miss you" email — 4% recovery', 'No visibility into ₹18.2L revenue at risk'].forEach((t,i) => {
      s.addText('• '+t, { x:0.5, y:2.26+i*0.57, w:4.1, h:0.45, fontSize:11, color:'374151', fontFace:'Calibri' });
    });
    s.addShape(pptx.ShapeType.ellipse, { x:4.62, y:3.2, w:0.56, h:0.56, fill:{color:C.red} });
    s.addText('VS', { x:4.62, y:3.2, w:0.56, h:0.56, fontSize:9, bold:true, color:C.white, align:'center', valign:'middle', fontFace:'Calibri' });
    s.addShape(pptx.ShapeType.roundRect, { x:5.0, y:1.65, w:4.7, h:3.8, fill:{color:C.lgray}, rectRadius:0.1, line:{color:C.green, width:0.75} });
    s.addText('✅  RetentionOS', { x:5.15, y:1.8, w:4.3, h:0.35, fontSize:13, bold:true, color:C.green, fontFace:'Calibri' });
    ['Churn predicted at day 8 — 22 days before renewal', 'Engagement decline signals scored per subscriber', 'Intervention at peak recovery window (days 8–22)', 'Goal-anchored message — 34% recovery rate', '₹18.2L at-risk revenue visible in real time'].forEach((t,i) => {
      s.addText('• '+t, { x:5.15, y:2.26+i*0.57, w:4.35, h:0.45, fontSize:11, color:'1E3A2F', fontFace:'Calibri' });
    });
    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:5.65, w:9.4, h:0.65, fill:{color:C.lgray}, rectRadius:0.06 });
    s.addText('Key insight: Winning back a churning PRO subscriber at day 8 costs ₹0 (milestone challenge). Acquiring a new PRO subscriber costs ₹400–800. RetentionOS makes the cheaper path the default path.', { x:0.5, y:5.7, w:9.0, h:0.55, fontSize:10, color:C.green, fontFace:'Calibri', italic:true, valign:'middle' });
  }
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.hero} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.gold} });
    s.addText('THE MECHANIC', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.gold, charSpacing:3, fontFace:'Calibri' });
    s.addText('Score Every Subscriber Daily. Intervene Before They Have Decided to Leave.', { x:0.3, y:0.5, w:9.4, h:0.65, fontSize:24, bold:true, color:C.white, fontFace:'Calibri' });
    const steps = [
      { n:'1', title:'Signal Collection', body:'Daily: login frequency, meal log rate, streak status, goal progress velocity, AI coach interaction, feature usage. Compared against subscriber\'s own 30-day baseline.' },
      { n:'2', title:'Churn Scoring', body:'ML model outputs churn probability per subscriber per day. Score above threshold AND renewal in 7–21 days → classify as AT-RISK or CRITICAL. Predicted churn reason also computed.' },
      { n:'3', title:'Intervention Match', body:'Churn reason → intervention type: plateau → milestone challenge. Low feature use → feature unlock. Streak break → streak recovery. No one-size-fits-all message.' },
      { n:'4', title:'Recovery Tracking', body:'Post-intervention: track engagement recovery, renewal event, and churn outcome. Feed back into model. A/B test intervention types. MTTR equivalent tracked per pattern.' },
    ];
    steps.forEach((st,i) => {
      const x = 0.25 + i*2.42;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.4, w:2.25, h:4.3, fill:{color:C.dark}, rectRadius:0.1, shadow:makeShadow() });
      s.addShape(pptx.ShapeType.ellipse, { x:x+0.78, y:1.52, w:0.68, h:0.68, fill:{color:C.red} });
      s.addText(st.n, { x:x+0.78, y:1.52, w:0.68, h:0.68, fontSize:16, bold:true, color:C.white, align:'center', valign:'middle', fontFace:'Calibri' });
      s.addText(st.title, { x:x+0.1, y:2.32, w:2.05, h:0.45, fontSize:12, bold:true, color:C.gold, align:'center', fontFace:'Calibri' });
      s.addText(st.body, { x:x+0.1, y:2.82, w:2.05, h:2.5, fontSize:10, color:C.muted, fontFace:'Calibri', valign:'top' });
      if(i<3) s.addText('→', { x:x+2.18, y:3.3, w:0.3, h:0.4, fontSize:14, color:C.red, align:'center', fontFace:'Calibri', bold:true });
    });
    s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:5.88, w:9.5, h:0.65, fill:{color:'030E05'}, rectRadius:0.06 });
    s.addText('PhonePe proof: Unified ops intelligence layer rebuilt 6 siloed workflows into one console — TAT 2 days → 4 hours, −68% escalation tickets. RetentionOS applies the same architecture: consolidate signals → classify churn type → route intervention.', { x:0.45, y:5.93, w:9.1, h:0.55, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true, valign:'middle' });
  }
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.white} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.red} });
    s.addText('THE PRODUCT', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.red, charSpacing:3, fontFace:'Calibri' });
    s.addText('4 Screens — Cohort Risk View to Intervention Launch', { x:0.3, y:0.5, w:9.4, h:0.65, fontSize:24, bold:true, color:C.dark, fontFace:'Calibri' });
    const cards = [
      { n:'01', title:'Retention Health', desc:'4 risk tier cohorts (Healthy / Watch / At-Risk / Critical). Revenue at stake: ₹18.2L. Clickable alert: "47 crossed into AT-RISK today."', mono:'Healthy 10,218\nWatch   1,425\nAt-Risk 1,047 !\nCrit.     157 !!' },
      { n:'02', title:'Subscriber Detail', desc:'Priya M.: 74% churn probability. Engagement signals: login 5.8→0.7/wk, meal logs 3.2→0.4/day. Predicted reason: plateau week 8. CTA: Send Intervention.', mono:'Login  5.8→0.7\nLogs   3.2→0.4\nCoach  3→0\nChurn  74%' },
      { n:'03', title:'Intervention Builder', desc:'Type selector (Milestone / Discount / Feature / Coach). Live message preview. Expected recovery rate: 31%. A/B test toggle. Launch to 47 subscribers.', mono:'Challenge  31%\nDiscount   28%\nFeature    22%\nCoach      36%' },
      { n:'04', title:'Impact Analytics', desc:'Intervention cohort renewal: 84.1% vs control: 71.2%. +12.9pp lift. ₹4.2L recovered. Milestone Challenge = highest ROI (₹0 cost, 34% recovery).', mono:'Intervened 84.1%\nControl    71.2%\nDelta    +12.9pp\nRecovered ₹4.2L' },
    ];
    cards.forEach((c2,i) => {
      const x = 0.25 + i*2.44;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.38, w:2.28, h:4.95, fill:{color:'FFF8F8'}, rectRadius:0.1, line:{color:'FFCDD2', width:0.5}, shadow:makeShadow() });
      s.addShape(pptx.ShapeType.roundRect, { x:x+0.06, y:1.45, w:2.16, h:0.3, fill:{color:C.red}, rectRadius:0.05 });
      s.addText(c2.n+' · '+c2.title, { x:x+0.06, y:1.45, w:2.16, h:0.3, fontSize:9, bold:true, color:C.white, align:'center', fontFace:'Calibri' });
      s.addText(c2.desc, { x:x+0.1, y:1.85, w:2.08, h:1.6, fontSize:10, color:'374151', fontFace:'Calibri', valign:'top' });
      s.addShape(pptx.ShapeType.roundRect, { x:x+0.1, y:3.5, w:2.08, h:1.0, fill:{color:C.dark}, rectRadius:0.06 });
      s.addText(c2.mono, { x:x+0.14, y:3.55, w:2.0, h:0.9, fontSize:9, color:'F87171', fontFace:'Courier New', valign:'top' });
    });
    s.addText('Prototype: healthifyme-retentionos-prototype.html · Intervention type selector live · A/B toggle interactive', { x:0.3, y:6.68, w:9.4, h:0.3, fontSize:9, color:C.muted, italic:true, fontFace:'Calibri', align:'center' });
  }
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.dark} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.green} });
    s.addText('IMPACT & ROI', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.green, charSpacing:3, fontFace:'Calibri' });
    s.addText('₹4.2L Recovered per Month. Milestone Challenge Costs ₹0 and Recovers 34%.', { x:0.3, y:0.5, w:9.4, h:0.65, fontSize:23, bold:true, color:C.white, fontFace:'Calibri' });
    const lm = [
      { v:'+12.9pp', l:'Renewal rate lift', s:'Intervention cohort 84.1% vs control 71.2%' },
      { v:'34%', l:'Milestone recovery rate', s:'At ₹0 marginal cost per subscriber' },
      { v:'22 days', l:'Pre-churn detection window', s:'Engage before motivation loss becomes permanent' },
      { v:'4×', l:'ROI vs Human Coach Call', s:'Same recovery rate, zero marginal cost' },
    ];
    const rm = [
      { v:'₹4.2L', l:'Revenue recovered (30 days)', s:'Subscriptions that would have churned, retained' },
      { v:'₹0', l:'Milestone challenge cost', s:'vs ₹400–800 new subscriber acquisition cost' },
      { v:'−68%', l:'Escalation tickets (PhonePe)', s:'Ops intelligence architecture — same pattern' },
      { v:'+20%', l:'30-day retention (PhonePe)', s:'Same signal → intervention mechanics applied' },
    ];
    [lm,rm].forEach((metrics,col) => {
      const xBase = col===0 ? 0.3 : 5.1;
      s.addText(col===0 ? 'Retention Metrics' : 'Business ROI', { x:xBase, y:1.35, w:4.5, h:0.28, fontSize:10, bold:true, color:col===0?C.red:C.gold, charSpacing:2, fontFace:'Calibri' });
      metrics.forEach((m,i) => {
        const y = 1.72 + i*0.92;
        s.addShape(pptx.ShapeType.roundRect, { x:xBase, y, w:4.55, h:0.82, fill:{color:C.hero}, rectRadius:0.06, shadow:makeShadow() });
        s.addText(m.v, { x:xBase+0.12, y:y+0.06, w:1.6, h:0.5, fontSize:21, bold:true, color:col===0?C.red:C.gold, fontFace:'Calibri', valign:'middle' });
        s.addText(m.l, { x:xBase+1.7, y:y+0.06, w:2.7, h:0.28, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
        s.addText(m.s, { x:xBase+1.7, y:y+0.36, w:2.7, h:0.36, fontSize:9, color:C.muted, fontFace:'Calibri', valign:'top' });
      });
    });
    s.addShape(pptx.ShapeType.rect, { x:4.88, y:1.35, w:0.02, h:5.35, fill:{color:'1A3020'} });
    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:6.5, w:9.4, h:0.75, fill:{color:'0D1A0A'}, rectRadius:0.06 });
    s.addText('Retaining a churning PRO subscriber at ₹0 cost is always better than acquiring a new one at ₹400–800. RetentionOS makes the cheaper path the default — at scale, across 1,204 at-risk subscribers per month.', { x:0.5, y:6.56, w:9.0, h:0.6, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true, valign:'middle' });
  }
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:4.9, h:'100%', fill:{color:C.dark} });
    s.addShape(pptx.ShapeType.rect, { x:4.9, y:0, w:5.1, h:'100%', fill:{color:C.lgray} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.red} });
    s.addText('PROOF OF WORK', { x:0.3, y:0.22, w:4.4, h:0.22, fontSize:9, bold:true, color:C.red, charSpacing:3, fontFace:'Calibri' });
    s.addText('The PhonePe Ops Intelligence Layer — Applied to PRO Subscriber Retention', { x:0.3, y:0.5, w:9.4, h:0.65, fontSize:20, bold:true, color:C.white, fontFace:'Calibri' });
    const lp = [
      'Built unified ops intelligence layer consolidating 6 siloed workflows → automated breach classification and routing → TAT 2 days → 4 hours, −68% escalation tickets',
      'Dynamic cart incentivization: context-aware trigger (cart × intent × time) → 20% improvement in 30-day retention. Same signal-to-intervention architecture.',
      'ML propensity-to-transact: real-time per-user scoring replacing manual cohorts → 32% marketing burn reduction while sustaining GMV across all markets',
      'Milestone waiver A/B test: goal milestone mechanic → 60% improvement in 90-day device activation vs flat discount. Same mechanic as RetentionOS Milestone Challenge.',
    ];
    const rp = [
      '→ RetentionOS signal consolidation — same architecture: 6 signals (login, logs, streak, goal, coach, feature) → single churn risk score replacing manual ops review',
      '→ RetentionOS intervention timing — same 30-day retention logic: detect engagement drop → intervene at peak recovery window → measure renewal outcome',
      '→ RetentionOS churn propensity scoring — same real-time ML scoring, subscription churn probability instead of purchase propensity',
      '→ RetentionOS Milestone Challenge — same goal milestone mechanic, journey completion instead of device activation. ₹0 cost, 34% recovery rate.',
    ];
    lp.forEach((txt,i) => {
      s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:1.4+i*1.2, w:4.45, h:1.1, fill:{color:C.hero}, rectRadius:0.07, shadow:makeShadow() });
      s.addText(txt, { x:0.4, y:1.46+i*1.2, w:4.15, h:0.98, fontSize:9.5, color:C.muted, fontFace:'Calibri', valign:'middle' });
    });
    rp.forEach((txt,i) => {
      s.addShape(pptx.ShapeType.roundRect, { x:5.05, y:1.4+i*1.2, w:4.65, h:1.1, fill:{color:C.white}, rectRadius:0.07, shadow:makeShadow() });
      s.addText(txt, { x:5.2, y:1.46+i*1.2, w:4.35, h:0.98, fontSize:9.5, color:'1E3A2F', fontFace:'Calibri', valign:'middle' });
    });
    s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:6.25, w:9.5, h:0.7, fill:{color:C.red}, rectRadius:0.06 });
    s.addText('RetentionOS is the ops intelligence architecture I built at PhonePe — adapted from offer ops escalation to PRO subscriber churn. The signal processing, classification, and intervention routing are identical.', { x:0.4, y:6.3, w:9.2, h:0.6, fontSize:10.5, color:C.white, fontFace:'Calibri', italic:true, valign:'middle' });
  }
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.dark} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.gold} });
    s.addText('ROLLOUT PLAN', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.gold, charSpacing:3, fontFace:'Calibri' });
    s.addText('Three Phases — Signal Audit to Full Intervention Automation', { x:0.3, y:0.5, w:9.4, h:0.65, fontSize:26, bold:true, color:C.white, fontFace:'Calibri' });
    const phases = [
      { n:'01', title:'Phase 1 — Month 1–2: Signal Validation', body:'Map existing engagement signals (login, logs, streak, AI coach) against historical churn events. Validate 14-day pre-churn signal pattern. Build baseline churn probability model. Identify top 3 churn patterns (plateau, streak break, feature abandonment).', color:C.red },
      { n:'02', title:'Phase 2 — Month 3–4: Intervention A/B Tests', body:'Deploy RetentionOS for AT-RISK tier. Run 4 intervention type A/B tests (Milestone, Discount, Feature, Coach). Measure 7-day recovery rate per type. Build ops dashboard for subscriber health visibility. Tune model thresholds.', color:C.gold },
      { n:'03', title:'Phase 3 — Month 5–6: Automated Pipeline', body:'Automate daily churn scoring → intervention dispatch → recovery tracking pipeline. Milestone Challenge becomes default intervention for plateau pattern. Real-time ₹ at-risk dashboard live for product and finance teams. MTTR equivalent tracked monthly.', color:C.green },
    ];
    phases.forEach((p,i) => {
      const x = 0.25 + i*3.28;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.38, w:3.08, h:4.1, fill:{color:C.hero}, rectRadius:0.1, shadow:makeShadow() });
      s.addShape(pptx.ShapeType.roundRect, { x:x+0.08, y:1.48, w:2.92, h:0.36, fill:{color:p.color}, rectRadius:0.05 });
      s.addText('Phase '+p.n, { x:x+0.08, y:1.48, w:2.92, h:0.36, fontSize:10, bold:true, color:C.dark, align:'center', fontFace:'Calibri' });
      s.addText(p.title, { x:x+0.12, y:1.93, w:2.84, h:0.55, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
      s.addText(p.body, { x:x+0.12, y:2.55, w:2.84, h:2.8, fontSize:10, color:C.muted, fontFace:'Calibri', valign:'top' });
    });
    s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:5.68, w:9.5, h:1.35, fill:{color:'030E05'}, rectRadius:0.08 });
    s.addText('What I Need', { x:0.45, y:5.76, w:4, h:0.28, fontSize:10, bold:true, color:C.gold, fontFace:'Calibri' });
    s.addText('• Subscriber event logs: login timestamps, meal log events, streak data, AI coach session data\n• Historical renewal/churn outcomes for model training (12-month lookback)\n• ML scoring infra to run daily churn probability computation across 12,847 active subscribers', { x:0.45, y:6.06, w:9.1, h:0.85, fontSize:9.5, color:C.muted, fontFace:'Calibri' });
    s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439', { x:0, y:7.18, w:'100%', h:0.25, fontSize:9, color:C.muted, align:'center', fontFace:'Calibri' });
  }

  await pptx.writeFile({ fileName:'assets/healthifyme-retentionos-deck.pptx' });
  console.log('✅ healthifyme-retentionos-deck.pptx saved');
}

// ─────────────────────────────────────────────────────────────────────────────
// DECK 3: MealMatch
// ─────────────────────────────────────────────────────────────────────────────
async function buildMealMatch() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';

  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:'1A0A00'} });
    for(let i=0;i<6;i++) s.addShape(pptx.ShapeType.line, { x:-0.5+i*1.8, y:0, w:2.4, h:7.5, line:{color:C.orange, width:0.5, transparency:88} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.18, h:'100%', fill:{color:C.orange} });
    s.addText('HEALTHIFYME · B2C AI PRODUCTS · CASE STUDY 03', { x:0.35, y:0.22, w:9, h:0.25, fontSize:8, bold:true, color:C.muted, charSpacing:3, fontFace:'Calibri' });
    s.addText('MealMatch', { x:0.35, y:0.7, w:7, h:1.2, fontSize:52, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText('Regional AI Meal Recommendation — No Competitor Has India\'s Food Data. Show It.', { x:0.35, y:1.88, w:8.5, h:0.55, fontSize:15, color:C.gold, fontFace:'Calibri', italic:true });
    s.addText('Ajay Avaghade · Product Manager', { x:0.35, y:2.55, w:5, h:0.3, fontSize:11, color:C.muted, fontFace:'Calibri' });
    s.addShape(pptx.ShapeType.roundRect, { x:7.1, y:5.5, w:2.65, h:1.6, fill:{color:C.orange}, rectRadius:0.1, shadow:makeShadow() });
    s.addText('+30pp', { x:7.1, y:5.56, w:2.65, h:0.75, fontSize:36, bold:true, color:C.white, align:'center', fontFace:'Calibri' });
    s.addText('acceptance rate:\nregional vs generic', { x:7.1, y:6.25, w:2.65, h:0.65, fontSize:9, color:'FFD0B0', align:'center', fontFace:'Calibri' });
  }
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:'1A0A00'} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.orange} });
    s.addText('THE PROBLEM', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.orange, charSpacing:3, fontFace:'Calibri' });
    s.addText('HealthifyMe Has a 10L+ Indian Food Database — and Serves Generic "Low-Cal Meal" Suggestions That Get Ignored', { x:0.3, y:0.5, w:9.4, h:0.9, fontSize:21, bold:true, color:C.white, fontFace:'Calibri' });
    const cols = [
      { icon:'🍽', stat:'38%', lbl:'Generic Meal CVR', sub:'Generic "low-calorie" recommendations accepted only 38% of the time. Users ignore suggestions that don\'t match their actual cuisine habits.' },
      { icon:'📍', stat:'0%', lbl:'Regional Signal Used', sub:'HealthifyMe knows the user is in Bangalore eating South Indian 87% of the time — none of this signal informs meal recommendations today.' },
      { icon:'❌', stat:'−2.4×', lbl:'Weight Loss Rate', sub:'Users with <60% meal plan adherence lose weight 2.4× slower than those with ≥80% adherence. Low acceptance = low adherence = poor outcomes.' },
    ];
    cols.forEach((c2,i) => {
      const x = 0.3 + i*3.25;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.72, w:3.0, h:3.2, fill:{color:'2A1500'}, rectRadius:0.1, shadow:makeShadow() });
      s.addText(c2.icon, { x, y:1.82, w:3.0, h:0.5, fontSize:24, align:'center' });
      s.addText(c2.stat, { x, y:2.37, w:3.0, h:0.55, fontSize:26, bold:true, color:C.orange, align:'center', fontFace:'Calibri' });
      s.addText(c2.lbl, { x, y:2.92, w:3.0, h:0.3, fontSize:11, bold:true, color:C.white, align:'center', fontFace:'Calibri' });
      s.addText(c2.sub, { x:x+0.12, y:3.27, w:2.76, h:1.35, fontSize:10, color:C.muted, align:'center', fontFace:'Calibri', valign:'top' });
    });
    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:5.12, w:9.4, h:0.95, fill:{color:'100800'}, rectRadius:0.06 });
    s.addText('HealthifyMe\'s moat is the Indian food database. MyFitnessPal has 14M+ generic foods. HealthifyMe has 10L+ Indian foods — including idli, ragi mudde, pesarattu. This database is worth nothing if the recommendation engine doesn\'t use regional signals.', { x:0.5, y:5.19, w:9.0, h:0.8, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true, valign:'middle' });
  }
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.white} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.orange} });
    s.addText('THE INSIGHT', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.orange, charSpacing:3, fontFace:'Calibri' });
    s.addText('A Meal Recommendation That Matches What the User Actually Eats Is 1.8× More Likely to Be Logged Than One That Doesn\'t', { x:0.3, y:0.5, w:9.4, h:0.9, fontSize:21, bold:true, color:C.dark, fontFace:'Calibri' });
    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:1.65, w:4.5, h:3.8, fill:{color:'FBF0EC'}, rectRadius:0.1, line:{color:'FFCCBC', width:0.75} });
    s.addText('❌  Today', { x:0.5, y:1.8, w:4.1, h:0.35, fontSize:13, bold:true, color:C.orange, fontFace:'Calibri' });
    ['Generic "low-calorie" meal suggestions', 'No cuisine preference signal used', 'No macro budget calculated for current meal', 'Same suggestion for Bangalore and Delhi user', 'Manual food search required for logging'].forEach((t,i) => {
      s.addText('• '+t, { x:0.5, y:2.26+i*0.57, w:4.1, h:0.45, fontSize:11, color:'374151', fontFace:'Calibri' });
    });
    s.addShape(pptx.ShapeType.ellipse, { x:4.62, y:3.2, w:0.56, h:0.56, fill:{color:C.orange} });
    s.addText('VS', { x:4.62, y:3.2, w:0.56, h:0.56, fontSize:9, bold:true, color:C.white, align:'center', valign:'middle', fontFace:'Calibri' });
    s.addShape(pptx.ShapeType.roundRect, { x:5.0, y:1.65, w:4.7, h:3.8, fill:{color:'FFF8F5'}, rectRadius:0.1, line:{color:C.orange, width:0.75} });
    s.addText('✅  MealMatch', { x:5.15, y:1.8, w:4.3, h:0.35, fontSize:13, bold:true, color:C.orange, fontFace:'Calibri' });
    ['Regional cuisine preference (87% South Indian → show South Indian)', 'Macro budget: remaining calories × goal type', 'Time-of-day: dinner → heavier + comfort foods', '"Why MealMatch chose this" transparency builds trust', '1-tap log: confirm a suggestion vs search from scratch'].forEach((t,i) => {
      s.addText('• '+t, { x:5.15, y:2.26+i*0.57, w:4.35, h:0.45, fontSize:11, color:'3A1500', fontFace:'Calibri' });
    });
    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:5.65, w:9.4, h:0.65, fill:{color:'FFF0E8'}, rectRadius:0.06 });
    s.addText('Key insight: HealthifyMe\'s moat is Indian food data. MealMatch is the product that makes that moat visible to users — and to themselves. No competitor can copy a personalised South Indian dinner at 610 cal for a Bangalore user.', { x:0.5, y:5.7, w:9.0, h:0.55, fontSize:10, color:C.orange, fontFace:'Calibri', italic:true, valign:'middle' });
  }
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:'2A1500'} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.gold} });
    s.addText('THE MECHANIC', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.gold, charSpacing:3, fontFace:'Calibri' });
    s.addText('Cuisine × Macro × Time × Goal → The Right South Indian Meal at the Right Moment', { x:0.3, y:0.5, w:9.4, h:0.65, fontSize:23, bold:true, color:C.white, fontFace:'Calibri' });
    const steps = [
      { n:'1', title:'Signal Stack', body:'Cuisine preference (inferred from log history), macro budget (remaining cals for the meal), time-of-day, goal type (weight loss / muscle / maintenance), past acceptance rate per dish category.' },
      { n:'2', title:'Recommendation Model', body:'Filter food DB by: cuisine match + macro budget fit + goal alignment. Rank by: past acceptance × nutritional goal fit × seasonal/time preference. Return top 3 per meal slot.' },
      { n:'3', title:'1-Tap Log UX', body:'"Why MealMatch chose this" transparency: show 3 reasons. Accept → 1-tap log. No search box. No manual entry. Logging drops from 10-field form to a confirmation tap.' },
      { n:'4', title:'Adherence Feedback Loop', body:'Weekly adherence score (% of suggested meals logged). Users at ≥80% shown "on track" with goal projection update. Adherence score becomes a self-accountability tool driving PRO retention.' },
    ];
    steps.forEach((st,i) => {
      const x = 0.25 + i*2.42;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.4, w:2.25, h:4.3, fill:{color:'1A0A00'}, rectRadius:0.1, shadow:makeShadow() });
      s.addShape(pptx.ShapeType.ellipse, { x:x+0.78, y:1.52, w:0.68, h:0.68, fill:{color:C.orange} });
      s.addText(st.n, { x:x+0.78, y:1.52, w:0.68, h:0.68, fontSize:16, bold:true, color:C.white, align:'center', valign:'middle', fontFace:'Calibri' });
      s.addText(st.title, { x:x+0.1, y:2.32, w:2.05, h:0.45, fontSize:12, bold:true, color:C.gold, align:'center', fontFace:'Calibri' });
      s.addText(st.body, { x:x+0.1, y:2.82, w:2.05, h:2.5, fontSize:10, color:C.muted, fontFace:'Calibri', valign:'top' });
      if(i<3) s.addText('→', { x:x+2.18, y:3.3, w:0.3, h:0.4, fontSize:14, color:C.orange, align:'center', fontFace:'Calibri', bold:true });
    });
    s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:5.88, w:9.5, h:0.65, fill:{color:'100800'}, rectRadius:0.06 });
    s.addText('PhonePe proof: ML-driven marketplace with audience segmentation (500+ brand partners, 11% YoY revenue growth, 26% engagement increase). Same segmentation engine — food preference instead of brand audience.', { x:0.45, y:5.93, w:9.1, h:0.55, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true, valign:'middle' });
  }
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:C.white} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.orange} });
    s.addText('THE PRODUCT', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.orange, charSpacing:3, fontFace:'Calibri' });
    s.addText('4 Screens — South Indian Suggestion to Weekly Adherence Score', { x:0.3, y:0.5, w:9.4, h:0.65, fontSize:24, bold:true, color:C.dark, fontFace:'Calibri' });
    const cards = [
      { n:'01', title:'Meal Plan Home', desc:'"Suggested for Tonight" powered by regional + macro signals. South Indian dinner suggestions with calorie + protein callout. 1-tap to meal detail.', mono:'🥞 Dosa 340cal\n🍲 Rasam 280cal\n🫕 Ragi  420cal\n[Why these? →]' },
      { n:'02', title:'Meal Detail', desc:'"Why MealMatch chose this" — 4 transparent reasons (budget fit, cuisine match, protein target, probiotic value). Ingredient / Recipe / Alternatives tabs. 1-tap log.', mono:'Cals: 340\nProtein: 14g\nCarbs:  52g\n[Log This Meal]' },
      { n:'03', title:'Weekly Adherence', desc:'82% adherence score ring. Day-by-day meal dots (logged vs planned). "5 of 6 days on plan." Goal projection update at current adherence level.', mono:'Mon ●●● On plan\nTue ●●○ Partial\nWed ●●● On plan\n82% adherence' },
      { n:'04', title:'Recommendation Ops', desc:'Acceptance rate by cuisine (South Indian 68% vs generic 38%) and meal time (dinner 62%). A/B: regional vs generic. Regional personalisation = +30pp acceptance.', mono:'S.Indian 68% ↑\nN.Indian 45% →\nCont.    31% →\nGeneric  38% —' },
    ];
    cards.forEach((c2,i) => {
      const x = 0.25 + i*2.44;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.38, w:2.28, h:4.95, fill:{color:'FFF8F5'}, rectRadius:0.1, line:{color:'FFCCBC', width:0.5}, shadow:makeShadow() });
      s.addShape(pptx.ShapeType.roundRect, { x:x+0.06, y:1.45, w:2.16, h:0.3, fill:{color:C.orange}, rectRadius:0.05 });
      s.addText(c2.n+' · '+c2.title, { x:x+0.06, y:1.45, w:2.16, h:0.3, fontSize:9, bold:true, color:C.white, align:'center', fontFace:'Calibri' });
      s.addText(c2.desc, { x:x+0.1, y:1.85, w:2.08, h:1.6, fontSize:10, color:'374151', fontFace:'Calibri', valign:'top' });
      s.addShape(pptx.ShapeType.roundRect, { x:x+0.1, y:3.5, w:2.08, h:1.0, fill:{color:'1A0A00'}, rectRadius:0.06 });
      s.addText(c2.mono, { x:x+0.14, y:3.55, w:2.0, h:0.9, fontSize:9, color:'FFAB91', fontFace:'Courier New', valign:'top' });
    });
    s.addText('Prototype: healthifyme-mealmatch-prototype.html · 1-tap log live · Tab navigation interactive', { x:0.3, y:6.68, w:9.4, h:0.3, fontSize:9, color:C.muted, italic:true, fontFace:'Calibri', align:'center' });
  }
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:'1A0A00'} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.orange} });
    s.addText('IMPACT & ROI', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.orange, charSpacing:3, fontFace:'Calibri' });
    s.addText('Regional Personalisation Drives 68% Acceptance. Higher Adherence Drives 2.4x Better Outcomes.', { x:0.3, y:0.5, w:9.4, h:0.65, fontSize:22, bold:true, color:C.white, fontFace:'Calibri' });
    const lm = [
      { v:'68%', l:'Acceptance rate (regional)', s:'South Indian cuisine — vs 38% generic recommendations' },
      { v:'+30pp', l:'Acceptance lift vs generic', s:'Regional personalisation closes the engagement gap' },
      { v:'2.4×', l:'Weight loss rate at 80%+', s:'High adherence cohort vs <60% adherence cohort' },
      { v:'−68%', l:'Log abandonment', s:'1-tap confirm vs 10-field manual food search entry' },
    ];
    const rm = [
      { v:'+11%', l:'YoY revenue (PhonePe proof)', s:'Same segmentation engine: 500+ brand partners, 11% growth' },
      { v:'−34%', l:'PRO churn rate', s:'Weekly adherence screen viewers churn 34% less' },
      { v:'10L+', l:'Indian food database moat', s:'No competitor (MyFitnessPal, Noom) can replicate this' },
      { v:'5 regional', l:'Cuisines to expand next', s:'Bengali, Gujarati, Maharashtrian — next highest gap cuisines' },
    ];
    [lm,rm].forEach((metrics,col) => {
      const xBase = col===0 ? 0.3 : 5.1;
      s.addText(col===0 ? 'Engagement Metrics' : 'Business & Moat', { x:xBase, y:1.35, w:4.5, h:0.28, fontSize:10, bold:true, color:col===0?C.orange:C.gold, charSpacing:2, fontFace:'Calibri' });
      metrics.forEach((m,i) => {
        const y = 1.72 + i*0.92;
        s.addShape(pptx.ShapeType.roundRect, { x:xBase, y, w:4.55, h:0.82, fill:{color:'2A1500'}, rectRadius:0.06, shadow:makeShadow() });
        s.addText(m.v, { x:xBase+0.12, y:y+0.06, w:1.6, h:0.5, fontSize:21, bold:true, color:col===0?C.orange:C.gold, fontFace:'Calibri', valign:'middle' });
        s.addText(m.l, { x:xBase+1.7, y:y+0.06, w:2.7, h:0.28, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
        s.addText(m.s, { x:xBase+1.7, y:y+0.36, w:2.7, h:0.36, fontSize:9, color:C.muted, fontFace:'Calibri', valign:'top' });
      });
    });
    s.addShape(pptx.ShapeType.rect, { x:4.88, y:1.35, w:0.02, h:5.35, fill:{color:'3A1500'} });
    s.addShape(pptx.ShapeType.roundRect, { x:0.3, y:6.5, w:9.4, h:0.75, fill:{color:'100800'}, rectRadius:0.06 });
    s.addText('MealMatch turns HealthifyMe\'s Indian food database from a backend asset into a customer-visible product advantage. A Bangalore user getting a Dosa recommendation at 610 cal will never find that on MyFitnessPal.', { x:0.5, y:6.56, w:9.0, h:0.6, fontSize:10, color:C.muted, fontFace:'Calibri', italic:true, valign:'middle' });
  }
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:4.9, h:'100%', fill:{color:'1A0A00'} });
    s.addShape(pptx.ShapeType.rect, { x:4.9, y:0, w:5.1, h:'100%', fill:{color:'FFF8F5'} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.orange} });
    s.addText('PROOF OF WORK', { x:0.3, y:0.22, w:4.4, h:0.22, fontSize:9, bold:true, color:C.orange, charSpacing:3, fontFace:'Calibri' });
    s.addText('The ML Marketplace Architecture from PhonePe — Applied to HealthifyMe\'s Food Moat', { x:0.3, y:0.5, w:9.4, h:0.65, fontSize:19, bold:true, color:C.white, fontFace:'Calibri' });
    const lp = [
      '₹100 Cr/yr rewards portfolio rebuilt as ML-driven marketplace with audience segmentation and brand self-serve tooling → 500+ brand partners, 11% YoY revenue growth, 26% user engagement increase',
      'Dynamic cart incentivization with context-aware triggers (cart value × user intent × time-of-day) → 35% AOV uplift, 60% cart abandonment reduction. Same signal stack as MealMatch.',
      '22% checkout conversion lift from redesigning offers placement, eligibility surfacing, and redemption UX — deep fluency in surfacing the right item at the right moment',
      '20% improvement in 30-day retention: users who engage with personalised content churn less. MealMatch adherence screen is designed around this exact same principle.',
    ];
    const rp = [
      '→ MealMatch segmentation — same audience segmentation logic: cuisine preference tags instead of brand audience segments. 10L+ food items, tagged by region, goal fit, macro profile.',
      '→ MealMatch trigger signals — same context-aware signal stack: remaining macro budget × meal time × goal type replaces cart value × user intent × time-of-day.',
      '→ MealMatch transparency layer — "Why MealMatch chose this" mirrors the PhonePe eligibility surfacing redesign: show the user why this item is right for them, not just what it is.',
      '→ MealMatch adherence screen — same retention insight: users who engage with personalised content weekly have measurably lower churn. Adherence score = retention anchor.',
    ];
    lp.forEach((txt,i) => {
      s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:1.4+i*1.2, w:4.45, h:1.1, fill:{color:'2A1500'}, rectRadius:0.07, shadow:makeShadow() });
      s.addText(txt, { x:0.4, y:1.46+i*1.2, w:4.15, h:0.98, fontSize:9.5, color:C.muted, fontFace:'Calibri', valign:'middle' });
    });
    rp.forEach((txt,i) => {
      s.addShape(pptx.ShapeType.roundRect, { x:5.05, y:1.4+i*1.2, w:4.65, h:1.1, fill:{color:C.white}, rectRadius:0.07, shadow:makeShadow() });
      s.addText(txt, { x:5.2, y:1.46+i*1.2, w:4.35, h:0.98, fontSize:9.5, color:'3A1500', fontFace:'Calibri', valign:'middle' });
    });
    s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:6.25, w:9.5, h:0.7, fill:{color:C.orange}, rectRadius:0.06 });
    s.addText('MealMatch is the ML marketplace architecture I built at PhonePe — with one substitution: food preferences and macro goals instead of brand audiences and discount mechanics.', { x:0.4, y:6.3, w:9.2, h:0.6, fontSize:10.5, color:C.white, fontFace:'Calibri', italic:true, valign:'middle' });
  }
  {
    const s = pptx.addSlide();
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color:'1A0A00'} });
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.12, h:'100%', fill:{color:C.gold} });
    s.addText('ROLLOUT PLAN', { x:0.3, y:0.22, w:4, h:0.22, fontSize:9, bold:true, color:C.gold, charSpacing:3, fontFace:'Calibri' });
    s.addText('Three Phases — Regional Tagging to Full Personalised Meal Engine', { x:0.3, y:0.5, w:9.4, h:0.65, fontSize:26, bold:true, color:C.white, fontFace:'Calibri' });
    const phases = [
      { n:'01', title:'Phase 1 — Month 1–2: Regional Tag Audit', body:'Tag top 500 most-logged foods by regional cuisine (South Indian, North Indian, Bengali, Gujarati, Maharashtrian). Build cuisine preference inference from log history. Validate acceptance rate correlation with regional match.', color:C.orange },
      { n:'02', title:'Phase 2 — Month 3–4: MealMatch MVP', body:'Deploy top-3 regional meal suggestions for dinner slot to PRO users in 3 cities (Bangalore, Mumbai, Delhi). Measure acceptance rate vs control (generic suggestions). A/B test: regional personalised vs generic. Build 1-tap log UX.', color:C.gold },
      { n:'03', title:'Phase 3 — Month 5–6: Full Meal Plan', body:'Expand to all 3 meal slots (breakfast / lunch / dinner). Add macro-budget calculation per meal. Launch weekly adherence score and goal projection. Expand regional tagging to top 8 Indian regional cuisines. Build ops analytics for recommendation quality.', color:C.green },
    ];
    phases.forEach((p,i) => {
      const x = 0.25 + i*3.28;
      s.addShape(pptx.ShapeType.roundRect, { x, y:1.38, w:3.08, h:4.1, fill:{color:'2A1500'}, rectRadius:0.1, shadow:makeShadow() });
      s.addShape(pptx.ShapeType.roundRect, { x:x+0.08, y:1.48, w:2.92, h:0.36, fill:{color:p.color}, rectRadius:0.05 });
      s.addText('Phase '+p.n, { x:x+0.08, y:1.48, w:2.92, h:0.36, fontSize:10, bold:true, color:C.dark, align:'center', fontFace:'Calibri' });
      s.addText(p.title, { x:x+0.12, y:1.93, w:2.84, h:0.55, fontSize:11, bold:true, color:C.white, fontFace:'Calibri' });
      s.addText(p.body, { x:x+0.12, y:2.55, w:2.84, h:2.8, fontSize:10, color:C.muted, fontFace:'Calibri', valign:'top' });
    });
    s.addShape(pptx.ShapeType.roundRect, { x:0.25, y:5.68, w:9.5, h:1.35, fill:{color:'100800'}, rectRadius:0.08 });
    s.addText('What I Need', { x:0.45, y:5.76, w:4, h:0.28, fontSize:10, bold:true, color:C.gold, fontFace:'Calibri' });
    s.addText('• Food log history with dish-level data (to build cuisine preference model)\n• Food database access for regional tagging (top 500 most-logged foods)\n• Product analytics: meal suggestion impression → acceptance → log events to measure recommendation quality', { x:0.45, y:6.06, w:9.1, h:0.85, fontSize:9.5, color:C.muted, fontFace:'Calibri' });
    s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439', { x:0, y:7.18, w:'100%', h:0.25, fontSize:9, color:C.muted, align:'center', fontFace:'Calibri' });
  }

  await pptx.writeFile({ fileName:'assets/healthifyme-mealmatch-deck.pptx' });
  console.log('✅ healthifyme-mealmatch-deck.pptx saved');
}

Promise.all([
  buildUpgradeIQ(),
  buildRetentionOS(),
  buildMealMatch(),
]).then(() => console.log('🎉 All 3 HealthifyMe decks saved!'));
