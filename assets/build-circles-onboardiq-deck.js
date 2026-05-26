const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9'; // 10" x 5.625"

const C = {
  dark:   '060D1A',
  hero:   '0B1829',
  bright: '00C6FF',
  purple: '7C3AED',
  green:  '22C55E',
  orange: 'F97316',
  gold:   'F59E0B',
  white:  'FFFFFF',
  lgray:  'F0F6FF',
  muted:  '8B9CB6',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:1, angle:45, opacity:0.14 });

// ─── SLIDE 1 — COVER ────────────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  // Diagonal accent lines
  for (let i = 0; i < 5; i++) {
    const lx = -0.5 + i * 2.2;
    const lw = Math.min(3.5, 10 - lx - 0.15);
    if (lx > 10 || lw <= 0) continue;
    s.addShape(prs.ShapeType.line, {
      x: lx, y: 0, w: lw, h: 5.625,
      line: { color: C.purple, width: 0.4 }, rotate: 30, transparency: 88,
    });
  }

  // Top tag
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 0.26, w: 3.2, h: 0.26, rectRadius: 0.05,
    fill: { color: C.purple, transparency: 88 }, line: { color: C.purple, width: 0.5, transparency: 50 } });
  s.addText('CIRCLES — MENA DIGITAL BRAND', { x: 0.5, y: 0.26, w: 3.2, h: 0.26,
    fontSize: 7, bold: true, color: C.bright, align: 'center', letterSpacingPt: 1 });

  // Product name
  s.addText('OnboardIQ', { x: 0.5, y: 0.72, w: 7, h: 1.0,
    fontSize: 50, bold: true, color: C.white, fontFace: 'Arial Black' });

  // Underline
  s.addShape(prs.ShapeType.rect, { x: 0.5, y: 1.72, w: 2.2, h: 0.05, fill: { color: C.purple } });

  // Subtitle
  s.addText('AI-Guided Subscriber Onboarding: Right Plan, First Time', {
    x: 0.5, y: 1.88, w: 7.2, h: 0.38,
    fontSize: 13.5, color: 'AABBD0', italic: true,
  });

  // Presenter
  s.addText('Presented by Ajay Avaghade  ·  Growth & Acquisition PM', {
    x: 0.5, y: 2.36, w: 7, h: 0.26,
    fontSize: 9.5, color: C.muted,
  });

  // Big stat box
  s.addShape(prs.ShapeType.roundRect, { x: 7.0, y: 3.8, w: 2.75, h: 1.55, rectRadius: 0.1,
    fill: { color: C.purple, transparency: 86 }, line: { color: C.purple, width: 1, transparency: 50 }, shadow: makeShadow() });
  s.addText('−43%', { x: 7.0, y: 3.88, w: 2.75, h: 0.72,
    fontSize: 40, bold: true, color: C.purple, align: 'center', fontFace: 'Arial Black' });
  s.addText('plan mismatch rate vs\nmanual plan selection', {
    x: 7.0, y: 4.60, w: 2.75, h: 0.55,
    fontSize: 8, color: C.muted, align: 'center',
  });
}

// ─── SLIDE 2 — THE PROBLEM ───────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE PROBLEM', { x: 0.5, y: 0.22, w: 9, h: 0.22,
    fontSize: 8.5, bold: true, color: C.muted, letterSpacingPt: 2 });
  s.addText('New Subscribers Choose the Wrong Plan — Then Churn at 3× the Rate', {
    x: 0.5, y: 0.50, w: 9, h: 0.65,
    fontSize: 20, bold: true, color: C.white,
  });

  const cols3 = [
    { icon:'😕', stat:'34%', lbl:'New subscribers choose\na plan that doesn\'t fit', sub:'Based on perceived value,\nnot actual usage behaviour' },
    { icon:'📉', stat:'2.4×', lbl:'Higher 90-day churn rate\nfor mismatched-plan subs', sub:'Day-1 plan fit predicts\n90-day retention better than NPS' },
    { icon:'🔁', stat:'AED 220', lbl:'Avg win-back cost vs\nAED 0 onboarding fix', sub:'Mismatched subscribers cost\n5× more to recover than retain' },
  ];
  cols3.forEach((col, i) => {
    const x = 0.5 + i * 3.1;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.15, w: 2.9, h: 2.8, rectRadius: 0.12,
      fill: { color: C.hero }, line: { color: C.bright, width: 0.3, transparency: 70 }, shadow: makeShadow() });
    s.addText(col.icon, { x, y: 1.25, w: 2.9, h: 0.45, fontSize: 22, align: 'center' });
    s.addText(col.stat, { x, y: 1.72, w: 2.9, h: 0.72,
      fontSize: 36, bold: true, color: C.purple, align: 'center', fontFace: 'Arial Black' });
    s.addText(col.lbl, { x, y: 2.48, w: 2.9, h: 0.58,
      fontSize: 9.5, color: C.white, align: 'center', bold: true });
    s.addText(col.sub, { x, y: 3.10, w: 2.9, h: 0.60,
      fontSize: 8.5, color: C.muted, align: 'center' });
  });

  // Insight box
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 4.22, w: 9, h: 0.80, rectRadius: 0.08,
    fill: { color: C.purple, transparency: 88 }, line: { color: C.purple, width: 0.6, transparency: 55 } });
  s.addText('Root cause: Circles\' plan catalogue has 7 options. Without guidance, new subscribers default to the cheapest or "middle" option — not the right one. Day-1 mismatch is the #1 preventable churn driver.', {
    x: 0.65, y: 4.28, w: 8.7, h: 0.68,
    fontSize: 8.5, color: C.white, italic: true,
  });
}

// ─── SLIDE 3 — THE INSIGHT ───────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE INSIGHT', { x: 0.5, y: 0.22, w: 9, h: 0.22,
    fontSize: 8.5, bold: true, color: C.muted, letterSpacingPt: 2 });
  s.addText('A 3-Question Quiz Outperforms a 7-Plan Catalogue.', {
    x: 0.5, y: 0.50, w: 9, h: 0.65,
    fontSize: 22, bold: true, color: C.dark,
  });

  // Left panel
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 1.1, w: 4.0, h: 3.0, rectRadius: 0.1,
    fill: { color: 'FFF0F8' }, line: { color: C.purple, width: 1, transparency: 50 }, shadow: makeShadow() });
  s.addText('❌  Current: Plan Catalogue', { x: 0.65, y: 1.22, w: 3.7, h: 0.30,
    fontSize: 11, bold: true, color: C.purple });
  const lftPoints = [
    'Subscriber lands on page with 7 plan tiles',
    'No guidance on which plan fits their life',
    'Chooses by price or GB number — not fit',
    '34% pick the wrong plan in first month',
    'Circles discovers the mismatch at churn',
  ];
  lftPoints.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 0.65, y: 1.60 + i * 0.42, w: 3.7, h: 0.38,
      fontSize: 9.5, color: C.dark });
  });

  // VS
  s.addShape(prs.ShapeType.ellipse, { x: 4.38, y: 2.45, w: 0.48, h: 0.48,
    fill: { color: C.dark }, line: { color: C.dark, width: 0 } });
  s.addText('VS', { x: 4.38, y: 2.45, w: 0.48, h: 0.48,
    fontSize: 8, bold: true, color: C.white, align: 'center', valign: 'middle' });

  // Right panel
  s.addShape(prs.ShapeType.roundRect, { x: 4.8, y: 1.1, w: 4.7, h: 3.0, rectRadius: 0.1,
    fill: { color: 'F0FFF4' }, line: { color: C.green, width: 1, transparency: 50 }, shadow: makeShadow() });
  s.addText('✅  OnboardIQ: AI-Guided', { x: 4.95, y: 1.22, w: 4.4, h: 0.30,
    fontSize: 11, bold: true, color: C.green });
  const rgtPoints = [
    '3 behavioural questions (data / travel / voice)',
    'ML model maps answers to plan fit score per SKU',
    'Surfaces #1 best-match plan with reasons',
    '91% average plan fit score (AI-matched cohort)',
    '90-day churn: 4.1% (AI) vs 9.8% (manual)',
  ];
  rgtPoints.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 4.95, y: 1.60 + i * 0.42, w: 4.4, h: 0.38,
      fontSize: 9.5, color: C.dark });
  });

  // Callout
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 4.26, w: 9, h: 0.78, rectRadius: 0.08,
    fill: { color: C.dark, transparency: 0 }, line: { color: C.bright, width: 0.5, transparency: 50 } });
  s.addText('Key insight: OnboardIQ doesn\'t upsell — it optimises for plan fit. A subscriber on the right AED 89 plan with 91% fit retains at 2.4× the rate of a mismatched subscriber on any plan.', {
    x: 0.65, y: 4.32, w: 8.7, h: 0.66,
    fontSize: 8.5, color: C.white, italic: true,
  });
}

// ─── SLIDE 4 — THE MECHANIC ──────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE MECHANIC', { x: 0.5, y: 0.22, w: 9, h: 0.22,
    fontSize: 8.5, bold: true, color: C.muted, letterSpacingPt: 2 });
  s.addText('3 Questions → ML Scoring → Best Plan → Guided Activation', {
    x: 0.5, y: 0.50, w: 9, h: 0.54,
    fontSize: 18, bold: true, color: C.white,
  });

  const steps = [
    { n:'1', title:'3 Usage Signals', body:'Data intensity (light/medium/heavy), Travel frequency (none/occasional/monthly+), Voice preference (calls/apps/hybrid)', col: C.bright },
    { n:'2', title:'Plan Fit Scoring', body:'ML model maps 3 signals to 7 plan SKUs → outputs a Plan Fit Score (0–100) for each plan. Top match surfaced instantly.', col: C.purple },
    { n:'3', title:'Transparent Match', body:'Subscriber sees #1 plan + 3 reasons why it fits their profile. Not just specs — explains in plain language what they\'d actually use.', col: C.green },
    { n:'4', title:'1-Tap Activation', body:'eSIM link in recommendation screen. Activated in < 2 min. Guided 3-step post-activation checklist (eSIM → AutoPay → alerts).', col: C.orange },
    { n:'5', title:'30-Day Re-Score', body:'Usage data collected post-activation. If actual behaviour deviates from predicted, ChurnShield flags for early intervention.', col: C.gold },
  ];
  const colW = 1.74;
  steps.forEach((st, i) => {
    const x = 0.5 + i * (colW + 0.18);
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.1, w: colW, h: 3.0, rectRadius: 0.1,
      fill: { color: C.hero }, line: { color: st.col, width: 1, transparency: 55 }, shadow: makeShadow() });
    s.addText(st.n, { x, y: 1.18, w: colW, h: 0.44,
      fontSize: 22, bold: true, color: st.col, align: 'center', fontFace: 'Arial Black' });
    s.addShape(prs.ShapeType.rect, { x: x + 0.6, y: 1.62, w: colW - 1.2, h: 0.03, fill: { color: st.col, transparency: 60 } });
    s.addText(st.title, { x, y: 1.68, w: colW, h: 0.34,
      fontSize: 9.5, bold: true, color: C.white, align: 'center' });
    s.addText(st.body, { x: x + 0.1, y: 2.08, w: colW - 0.2, h: 1.9,
      fontSize: 8, color: C.muted, valign: 'top' });
  });

  // Note at bottom
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 4.26, w: 9, h: 0.80, rectRadius: 0.08,
    fill: { color: C.hero }, line: { color: C.bright, width: 0.5, transparency: 60 } });
  s.addText('SmartPlan link: Steps 1–2 use the same Plan Fit Score engine as Circles SmartPlan (ongoing plan health). OnboardIQ applies it at acquisition — same model, different trigger moment.', {
    x: 0.65, y: 4.32, w: 8.7, h: 0.68,
    fontSize: 8.5, color: C.white, italic: true,
  });
}

// ─── SLIDE 5 — THE PRODUCT ───────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE PRODUCT', { x: 0.5, y: 0.22, w: 9, h: 0.22,
    fontSize: 8.5, bold: true, color: C.muted, letterSpacingPt: 2 });
  s.addText('4 Screen States — Subscriber + Ops View', {
    x: 0.5, y: 0.50, w: 9, h: 0.54,
    fontSize: 22, bold: true, color: C.dark,
  });

  const screens = [
    {
      num: '01', title: 'AI Plan Picker',
      body: '3 usage questions with chip selectors (no dropdowns). Progress dots show completion. Single CTA: "Show My Perfect Plan".',
      mock: '  OnboardIQ\n  ─────────────\n  Q1: Data usage?\n  ● Social/Stream\n  Q2: Travel?\n  ● Few times/yr\n  Q3: Voice?\n  ● App-first\n  Show My Plan →',
    },
    {
      num: '02', title: 'AI Recommendation',
      body: 'Plan fit score badge (94%). Plan name, price, features. 4 "why this plan" reasons. Ranked second option visible. Activate CTA.',
      mock: '  ⭐ 94% Fit\n  ─────────────\n  Circles 30GB GCC\n  AED 89/mo\n  ✓ 30GB local\n  ✓ 10GB roaming\n  ✓ Unlimited social\n  Activate Now →',
    },
    {
      num: '03', title: 'Welcome Aboard',
      body: 'Confetti + rocket icon. eSIM QR code link. 3-step post-activation checklist. XP progress bar gamifies completion.',
      mock: '  🚀 Welcome!\n  ─────────────\n  Circles 30GB GCC\n  Active · AED 89/mo\n  ─────────────\n  1. Install eSIM\n  2. AutoPay setup\n  3. Usage alerts\n  Progress: 1/3',
    },
    {
      num: '04', title: 'Ops Intelligence',
      body: '−43% mismatch rate vs manual. Funnel (visit→quiz→activate). Top plan cohorts + ARPU. 90-day churn: 4.1% vs 9.8%.',
      mock: '  Acquisition Dash\n  ─────────────\n  12,480 onboarded\n  91% completion\n  −43% mismatch\n  ─────────────\n  Churn: 4.1% AI\n  vs 9.8% manual',
    },
  ];

  screens.forEach((sc, i) => {
    const x = 0.28 + i * 2.38;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.08, w: 2.18, h: 4.3, rectRadius: 0.12,
      fill: { color: C.dark }, line: { color: C.purple, width: 0.6, transparency: 55 }, shadow: makeShadow() });
    s.addText(`0${i+1}`, { x, y: 1.16, w: 2.18, h: 0.28,
      fontSize: 9, bold: true, color: C.purple, align: 'center', letterSpacingPt: 1 });
    s.addText(sc.title, { x, y: 1.46, w: 2.18, h: 0.28,
      fontSize: 10, bold: true, color: C.white, align: 'center' });
    s.addShape(prs.ShapeType.roundRect, { x: x + 0.12, y: 1.82, w: 1.94, h: 1.7, rectRadius: 0.06,
      fill: { color: C.hero }, line: { color: C.purple, width: 0.3, transparency: 70 } });
    s.addText(sc.mock, { x: x + 0.16, y: 1.87, w: 1.86, h: 1.60,
      fontSize: 6.5, color: C.bright, fontFace: 'Courier New', valign: 'top' });
    s.addText(sc.body, { x: x + 0.1, y: 3.60, w: 1.98, h: 1.65,
      fontSize: 8, color: '334455', valign: 'top' });
  });
}

// ─── SLIDE 6 — IMPACT & ROI ──────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('IMPACT & ROI', { x: 0.5, y: 0.22, w: 9, h: 0.22,
    fontSize: 8.5, bold: true, color: C.muted, letterSpacingPt: 2 });
  s.addText('Projected on MENA Launch Cohort (12K First-Month Subs)', {
    x: 0.5, y: 0.50, w: 9, h: 0.54,
    fontSize: 22, bold: true, color: C.white,
  });

  // Left header
  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 1.1, w: 4.35, h: 0.32, rectRadius: 0.05,
    fill: { color: C.purple, transparency: 82 }, line: { color: C.purple, width: 0.4, transparency: 50 } });
  s.addText('SUBSCRIBER EXPERIENCE', { x: 0.4, y: 1.1, w: 4.35, h: 0.32,
    fontSize: 8.5, bold: true, color: C.white, align: 'center' });

  const lm = [
    { val: '91', lbl: 'Average Plan Fit Score\n(AI-matched cohort)', col: C.purple },
    { val: '−43%', lbl: 'Plan mismatch rate\nvs manual selection', col: C.bright },
    { val: '60s', lbl: 'Time to plan decision\n(3 questions, instant match)', col: C.green },
    { val: '4.1%', lbl: '90-day churn rate\n(AI cohort vs 9.8% manual)', col: C.orange },
  ];
  lm.forEach((m, i) => {
    s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 1.52 + i * 0.74, w: 4.35, h: 0.66, rectRadius: 0.08,
      fill: { color: C.hero }, line: { color: m.col, width: 0.3, transparency: 65 } });
    s.addText(m.val, { x: 0.5, y: 1.56 + i * 0.74, w: 1.4, h: 0.58,
      fontSize: 22, bold: true, color: m.col, align: 'center', valign: 'middle', fontFace: 'Arial Black' });
    s.addText(m.lbl, { x: 1.95, y: 1.58 + i * 0.74, w: 2.65, h: 0.54,
      fontSize: 9, color: C.white, valign: 'middle' });
  });

  // Right header
  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.1, w: 4.5, h: 0.32, rectRadius: 0.05,
    fill: { color: C.bright, transparency: 82 }, line: { color: C.bright, width: 0.4, transparency: 50 } });
  s.addText('OPERATOR ROI', { x: 5.1, y: 1.1, w: 4.5, h: 0.32,
    fontSize: 8.5, bold: true, color: C.white, align: 'center' });

  const rm = [
    { val: '+AED 1.4M', lbl: 'Annual ARPU impact from\nreduced early-churn rate', col: C.bright },
    { val: '69%', lbl: 'Activation rate from\nplan page (quiz starters)', col: C.purple },
    { val: '−57%', lbl: 'Win-back cost saved\nper prevented early churn', col: C.green },
    { val: 'AED 0', lbl: 'Marginal cost per\nadditional OnboardIQ match', col: C.orange },
  ];
  rm.forEach((m, i) => {
    s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.52 + i * 0.74, w: 4.5, h: 0.66, rectRadius: 0.08,
      fill: { color: C.hero }, line: { color: m.col, width: 0.3, transparency: 65 } });
    s.addText(m.val, { x: 5.2, y: 1.56 + i * 0.74, w: 1.9, h: 0.58,
      fontSize: i===0 ? 16 : 22, bold: true, color: m.col, align: 'center', valign: 'middle', fontFace: 'Arial Black' });
    s.addText(m.lbl, { x: 7.15, y: 1.58 + i * 0.74, w: 2.2, h: 0.54,
      fontSize: 9, color: C.white, valign: 'middle' });
  });

  // Insight
  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 4.62, w: 9.2, h: 0.70, rectRadius: 0.08,
    fill: { color: C.hero }, line: { color: C.bright, width: 0.5, transparency: 55 } });
  s.addText('OnboardIQ also feeds ChurnShield: AI-matched subscribers start with a baseline Plan Fit Score, so ChurnShield\'s day-30 re-score has a reference point from Day 1.', {
    x: 0.55, y: 4.66, w: 9.0, h: 0.60,
    fontSize: 8.5, color: C.white, italic: true,
  });
}

// ─── SLIDE 7 — PROOF OF WORK ─────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('PROOF OF WORK', { x: 0.5, y: 0.22, w: 9, h: 0.22,
    fontSize: 8.5, bold: true, color: C.muted, letterSpacingPt: 2 });
  s.addText('I Built Every Piece of This. Here\'s the Mapping.', {
    x: 0.5, y: 0.50, w: 9, h: 0.54,
    fontSize: 22, bold: true, color: C.dark,
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 1.08, w: 4.4, h: 4.0, rectRadius: 0.12,
    fill: { color: C.dark }, line: { color: C.bright, width: 0.8, transparency: 45 }, shadow: makeShadow() });
  s.addText('What I Shipped at PhonePe', { x: 0.55, y: 1.18, w: 4.1, h: 0.32,
    fontSize: 11, bold: true, color: C.bright });

  const lPoints = [
    { t:'Dynamic incentivisation engine:', b:'Context-aware triggers — exactly the same signal-to-action architecture as OnboardIQ\'s 3-question → plan match' },
    { t:'A/B testing at scale:', b:'Ran Rent Waiver vs Cashback across 350M MAU with pre-defined KPIs — 60% improvement in 90-day activation' },
    { t:'0→1 consumer product launch:', b:'Kotak Cherry: UX research → MVP scope → GTM → 100K downloads. Rings/MENA launch is this, at scale.' },
    { t:'ML-driven segmentation:', b:'Propensity-to-Transact model: user-level scoring replacing static rules — same engine as plan-fit scoring' },
    { t:'Ops instrumentation:', b:'Replaced 6 siloed analyst workflows with unified KPI dashboard — TAT 2d → 4h (maps to OnboardIQ\'s acquisition ops view)' },
  ];
  lPoints.forEach((pt, i) => {
    s.addText(`• ${pt.t}`, { x: 0.55, y: 1.60 + i * 0.66, w: 4.1, h: 0.22,
      fontSize: 9.5, bold: true, color: C.white });
    s.addText(`  ${pt.b}`, { x: 0.55, y: 1.82 + i * 0.66, w: 4.1, h: 0.38,
      fontSize: 8.5, color: C.muted });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.08, w: 4.5, h: 4.0, rectRadius: 0.12,
    fill: { color: 'FFFFFF' }, line: { color: C.green, width: 0.8, transparency: 45 }, shadow: makeShadow() });
  s.addText('How It Maps to Circles JD', { x: 5.25, y: 1.18, w: 4.2, h: 0.32,
    fontSize: 11, bold: true, color: C.dark });

  const rPoints = [
    { t:'Digital product performance', b:'→ Funnel KPIs: plan-page visit → quiz start → activate → D30 retention — defined pre-launch' },
    { t:'Acquisition & growth', b:'→ 5,000+ B2B merchants acquired in one self-serve build; same zero-friction activation pattern' },
    { t:'Feature adoption', b:'→ Milestone incentive drove 60% device activation — same behaviour-change architecture as onboarding nudges' },
    { t:'Cross-functional delivery', b:'→ Owned Agile delivery across Tech, DS, Design, Compliance — mirrors Circles\' matrix' },
    { t:'0→1 consumer launches', b:'→ Kotak Cherry founding team: research → MVP → GTM → launch in regulated digital environment' },
  ];
  rPoints.forEach((pt, i) => {
    s.addText(`→ ${pt.t}`, { x: 5.25, y: 1.60 + i * 0.66, w: 4.2, h: 0.22,
      fontSize: 9.5, bold: true, color: C.dark });
    s.addText(`  ${pt.b}`, { x: 5.25, y: 1.82 + i * 0.66, w: 4.2, h: 0.38,
      fontSize: 8.5, color: C.muted });
  });
}

// ─── SLIDE 8 — ROLLOUT PLAN ──────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('ROLLOUT PLAN', { x: 0.5, y: 0.22, w: 9, h: 0.22,
    fontSize: 8.5, bold: true, color: C.muted, letterSpacingPt: 2 });
  s.addText('3-Phase Deployment — Launch-Ready in 8 Weeks', {
    x: 0.5, y: 0.50, w: 9, h: 0.54,
    fontSize: 22, bold: true, color: C.white,
  });

  const phases = [
    {
      phase: 'Phase 1', time: 'Week 1–3',
      title: 'Quiz + ML Baseline',
      points: ['Build 3-question quiz UI (web + app)', 'Train plan-fit scoring model on pilot data', 'A/B: OnboardIQ quiz vs manual plan catalogue', 'Measure plan fit score vs 30-day usage delta'],
      col: C.bright,
    },
    {
      phase: 'Phase 2', time: 'Week 4–6',
      title: 'Recommendation Engine',
      points: ['Deploy personalised plan match with reasons', 'Build ops acquisition dashboard (funnel + cohorts)', 'Integrate with eSIM activation flow', 'Launch guided 3-step post-activation checklist'],
      col: C.purple,
    },
    {
      phase: 'Phase 3', time: 'Week 7–8+',
      title: 'Live + Iteration',
      points: ['Full MENA launch rollout (default onboarding flow)', 'Connect to ChurnShield: Day-1 score feeds re-score', 'Monthly model retrain on growing activation data', 'Expand to plan upgrade recommendations for existing base'],
      col: C.green,
    },
  ];

  phases.forEach((ph, i) => {
    const x = 0.4 + i * 3.2;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.1, w: 2.95, h: 2.92, rectRadius: 0.1,
      fill: { color: C.hero }, line: { color: ph.col, width: 1, transparency: 45 }, shadow: makeShadow() });
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.1, w: 2.95, h: 0.42, rectRadius: 0.1,
      fill: { color: ph.col, transparency: 80 }, line: { color: 'FFFFFF', width: 0, transparency: 100 } });
    s.addText(`${ph.phase}  ·  ${ph.time}`, { x, y: 1.13, w: 2.95, h: 0.22,
      fontSize: 8.5, bold: true, color: ph.col, align: 'center' });
    s.addText(ph.title, { x, y: 1.35, w: 2.95, h: 0.26,
      fontSize: 10.5, bold: true, color: C.white, align: 'center' });
    ph.points.forEach((pt, j) => {
      s.addText(`• ${pt}`, { x: x + 0.15, y: 1.68 + j * 0.50, w: 2.65, h: 0.46,
        fontSize: 8.5, color: C.muted, valign: 'top' });
    });
  });

  // Ask box
  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 4.2, w: 9.2, h: 0.95, rectRadius: 0.08,
    fill: { color: C.hero }, line: { color: C.bright, width: 0.7, transparency: 45 } });
  s.addText('What I need to build this:', { x: 0.6, y: 4.26, w: 4.0, h: 0.22,
    fontSize: 9, bold: true, color: C.bright });
  s.addText('Access to plan catalogue data + historical subscriber usage  ·  Alignment with acquisition & digital product teams  ·  1 frontend engineer  ·  1 DS engineer for scoring model', {
    x: 0.6, y: 4.48, w: 9.0, h: 0.30,
    fontSize: 8.5, color: C.white,
  });

  // Footer
  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.32, w: 10, h: 0.28,
    fill: { color: C.hero }, line: { color: 'FFFFFF', width: 0, transparency: 100 } });
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  ajay-avaghade.github.io/Portfolio', {
    x: 0.4, y: 5.33, w: 9.2, h: 0.24,
    fontSize: 8, color: C.muted,
  });
}

// ─── WRITE ───────────────────────────────────────────────────────────────────
prs.writeFile({ fileName: 'assets/circles-onboardiq-deck.pptx' }).then(() => {
  console.log('✅  circles-onboardiq-deck.pptx built');
}).catch(e => { console.error(e); process.exit(1); });
