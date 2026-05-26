const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9'; // 10" x 5.625"

const C = {
  dark:   '060D1A',
  hero:   '0B1829',
  bright: '00C6FF',
  red:    'EF4444',
  green:  '22C55E',
  orange: 'F97316',
  purple: '8B5CF6',
  white:  'FFFFFF',
  lgray:  'F0F6FF',
  muted:  '8B9CB6',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:1, angle:45, opacity:0.14 });

// ─── SLIDE 1 — COVER ────────────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  // Diagonal accent lines (clipped to slide width)
  for (let i = 0; i < 5; i++) {
    const lx = -0.5 + i * 2.2;
    const lw = Math.min(3.5, 10 - lx - 0.15);
    if (lx > 10 || lw <= 0) continue;
    s.addShape(prs.ShapeType.line, {
      x: lx, y: 0, w: lw, h: 5.625,
      line: { color: C.red, width: 0.4 }, rotate: 30, transparency: 88,
    });
  }

  // Top tag
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 0.26, w: 3.2, h: 0.26, rectRadius: 0.05,
    fill: { color: C.red, transparency: 88 }, line: { color: C.red, width: 0.5, transparency: 50 } });
  s.addText('CIRCLES — MENA DIGITAL BRAND', { x: 0.5, y: 0.26, w: 3.2, h: 0.26,
    fontSize: 7, bold: true, color: C.bright, align: 'center', letterSpacingPt: 1 });

  // Product name
  s.addText('ChurnShield', { x: 0.5, y: 0.72, w: 6, h: 1.0,
    fontSize: 50, bold: true, color: C.white, fontFace: 'Arial Black' });

  // Underline accent
  s.addShape(prs.ShapeType.rect, { x: 0.5, y: 1.72, w: 2.2, h: 0.05, fill: { color: C.red } });

  // Subtitle
  s.addText('ML Churn Prediction & Proactive Retention for MENA Subscribers', {
    x: 0.5, y: 1.88, w: 7.2, h: 0.38,
    fontSize: 13.5, color: 'AABBD0', italic: true,
  });

  // Presenter line
  s.addText('Presented by Ajay Avaghade  ·  Growth & Retention PM', {
    x: 0.5, y: 2.36, w: 7, h: 0.26,
    fontSize: 9.5, color: C.muted,
  });

  // Big stat box bottom-right
  s.addShape(prs.ShapeType.roundRect, { x: 7.0, y: 3.8, w: 2.75, h: 1.55, rectRadius: 0.1,
    fill: { color: C.red, transparency: 86 }, line: { color: C.red, width: 1, transparency: 50 }, shadow: makeShadow() });
  s.addText('3.2×', { x: 7.0, y: 3.88, w: 2.75, h: 0.72,
    fontSize: 40, bold: true, color: C.red, align: 'center', fontFace: 'Arial Black' });
  s.addText('churn rate for mismatched-plan subscribers\nvs healthy-plan cohort', {
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
  s.addText('Subscribers Churn Silently — Operators Never Know Why', {
    x: 0.5, y: 0.50, w: 9, h: 0.65,
    fontSize: 22, bold: true, color: C.white,
  });

  // 3 stat columns
  const cols3 = [
    { icon:'🚪', stat:'31%', lbl:'Of subscribers churn within\n12 months of plan mismatch', sub:'Mismatched data/roaming plans\ndrive 3× the churn rate' },
    { icon:'🔇', stat:'89%', lbl:'Churn happens with zero\nprior complaint or contact', sub:'Silent churners: no app activity,\nno support ticket, then gone' },
    { icon:'💸', stat:'AED 2.1M', lbl:'Annual ARPU at risk\nper 250K subscriber base', sub:'12% of base mismatched ×\nAED 89 avg monthly ARPU' },
  ];
  cols3.forEach((col, i) => {
    const x = 0.5 + i * 3.1;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.15, w: 2.9, h: 2.8, rectRadius: 0.12,
      fill: { color: C.hero }, line: { color: C.bright, width: 0.3, transparency: 70 }, shadow: makeShadow() });
    s.addText(col.icon, { x, y: 1.25, w: 2.9, h: 0.45, fontSize: 22, align: 'center' });
    s.addText(col.stat, { x, y: 1.72, w: 2.9, h: 0.72,
      fontSize: i===2 ? 24 : 36, bold: true, color: C.red, align: 'center', fontFace: 'Arial Black' });
    s.addText(col.lbl, { x, y: 2.48, w: 2.9, h: 0.58,
      fontSize: 9.5, color: C.white, align: 'center', bold: true });
    s.addText(col.sub, { x, y: 3.10, w: 2.9, h: 0.60,
      fontSize: 8.5, color: C.muted, align: 'center' });
  });

  // Insight box
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 4.22, w: 9, h: 0.80, rectRadius: 0.08,
    fill: { color: C.red, transparency: 88 }, line: { color: C.red, width: 0.6, transparency: 55 } });
  s.addText('Root cause: Circles has the usage data. Subscribers just don\'t know they\'re on the wrong plan. By the time they complain, 89% have already decided to leave.', {
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
  s.addText('Reactive Support Misses the Window. Proactive ML Doesn\'t.', {
    x: 0.5, y: 0.50, w: 9, h: 0.65,
    fontSize: 22, bold: true, color: C.dark,
  });

  // Left panel — current
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 1.1, w: 4.0, h: 3.0, rectRadius: 0.1,
    fill: { color: 'FFF0F0' }, line: { color: C.red, width: 1, transparency: 50 }, shadow: makeShadow() });
  s.addText('❌  Current: Reactive Retention', { x: 0.65, y: 1.22, w: 3.7, h: 0.30,
    fontSize: 11, bold: true, color: C.red });
  const lftPoints = [
    'Subscriber hits data cap → frustration builds',
    'Pays unexpected roaming charge → anger spikes',
    'Contacts support or simply doesn\'t renew',
    'Ops sees the churn only after MNP request',
    'Win-back cost: 5-8× acquisition cost',
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

  // Right panel — proposed
  s.addShape(prs.ShapeType.roundRect, { x: 4.8, y: 1.1, w: 4.7, h: 3.0, rectRadius: 0.1,
    fill: { color: 'F0FFF4' }, line: { color: C.green, width: 1, transparency: 50 }, shadow: makeShadow() });
  s.addText('✅  ChurnShield: Proactive ML', { x: 4.95, y: 1.22, w: 4.4, h: 0.30,
    fontSize: 11, bold: true, color: C.green });
  const rgtPoints = [
    'ML scores every subscriber daily on 90-day usage',
    'Risk ≥ 60 triggers personalised in-app nudge',
    'Subscriber sees exactly why + best fix with 1 tap',
    'Plan switch or add-on offered at the right moment',
    'Churn risk drops 41% per converted subscriber',
  ];
  rgtPoints.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 4.95, y: 1.60 + i * 0.42, w: 4.4, h: 0.38,
      fontSize: 9.5, color: C.dark });
  });

  // Callout box
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 4.26, w: 9, h: 0.78, rectRadius: 0.08,
    fill: { color: C.dark, transparency: 0 }, line: { color: C.bright, width: 0.5, transparency: 50 } });
  s.addText('Key insight: ChurnShield closes the window where intervention is 5× cheaper than win-back. 61% of at-risk subscribers convert when reached before renewal vs 12% after MNP initiation.', {
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
  s.addText('5-Step ML Pipeline: Score → Segment → Intervene → Convert → Measure', {
    x: 0.5, y: 0.50, w: 9, h: 0.54,
    fontSize: 18, bold: true, color: C.white,
  });

  // 5 step cards
  const steps = [
    { n:'1', title:'Daily Scoring', body:'ML model processes 90-day usage vectors (data, roaming, voice, peak hours, billing proximity) → outputs Churn Risk Score 0–100 per subscriber', col: C.bright },
    { n:'2', title:'Segmentation', body:'Score ≥ 60: High-risk queue (real-time push). Score 40–59: Medium (billing-cycle nudge). Score < 40: Healthy (no action)', col: C.orange },
    { n:'3', title:'Personalised Trigger', body:'In-app alert shows subscriber\'s specific risk drivers + best-match intervention ranked by predicted conversion probability', col: C.purple },
    { n:'4', title:'1-Tap Resolution', body:'Subscriber accepts plan switch or add-on. Effective immediately. Churn risk score updates in real-time.', col: C.green },
    { n:'5', title:'Closed Loop', body:'Post-intervention 30-day tracking: re-score, measure churn delta, feed back to model for continuous improvement', col: C.red },
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

  // Bottom note
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 4.26, w: 9, h: 0.80, rectRadius: 0.08,
    fill: { color: C.hero }, line: { color: C.bright, width: 0.5, transparency: 60 } });
  s.addText('Proof of concept: PhonePe Propensity-to-Transact ML — same pipeline architecture, same closed-loop feedback. Shipped to 350M+ MAU, reduced marketing spend 32% while sustaining GMV growth.', {
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
  s.addText('4 Screen States — Subscriber Journey + Ops View', {
    x: 0.5, y: 0.50, w: 9, h: 0.54,
    fontSize: 22, bold: true, color: C.dark,
  });

  const screens = [
    {
      num: '01', title: 'Churn Risk Alert',
      body: 'Risk score (0–100) surfaces on home screen when subscriber hits threshold. 3 specific risk drivers shown. Single CTA: "See My Recommendation"',
      mock: '  Churn Risk\n  ─────────────\n  Score: 78/100\n  ❗ Data cap stress\n  ❗ Roaming charges\n  ◦ Voice mismatch\n  ─────────────\n  See Offer →',
    },
    {
      num: '02', title: 'ML Risk Analyzer',
      body: 'Risk circle + 4 driver bars with point contributions. ML insight explains cohort context. 2 ranked interventions with predicted churn reduction %',
      mock: '  Risk Profile\n  ─────────────\n  ████░░ 34 pts\n  ████░░ 24 pts\n  ██░░░░ 12 pts\n  █░░░░░  8 pts\n  ─────────────\n  ML: 3.2× base risk',
    },
    {
      num: '03', title: 'Offer Accepted',
      body: 'Confetti + shield icon. New plan name + benefits. Before/after risk score comparison (78→37). ARPU save shown. Auto-applied to billing.',
      mock: '  🛡️ Plan Switched!\n  ─────────────\n  Circles 50GB\n  +20GB · Roaming\n  AED 89/mo\n  ─────────────\n  Risk: 78 → 37\n  Saved: AED 12/mo',
    },
    {
      num: '04', title: 'Ops Intelligence',
      body: 'Risk distribution (high/medium/low). Intervention funnel with conversion rates. Per-type ARPU impact. AED 2.1M annualised ARPU rescued.',
      mock: '  Retention Dash\n  ─────────────\n  High:  8,420\n  Med:  12,300\n  Low: 229,280\n  ─────────────\n  Conv: 31% · ARPU↑\n  AED 2.1M rescued',
    },
  ];

  screens.forEach((sc, i) => {
    const x = 0.28 + i * 2.38;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.08, w: 2.18, h: 4.3, rectRadius: 0.12,
      fill: { color: C.dark }, line: { color: C.bright, width: 0.6, transparency: 55 }, shadow: makeShadow() });
    s.addText(`0${i+1}`, { x, y: 1.16, w: 2.18, h: 0.28,
      fontSize: 9, bold: true, color: C.bright, align: 'center', letterSpacingPt: 1 });
    s.addText(sc.title, { x, y: 1.46, w: 2.18, h: 0.28,
      fontSize: 10, bold: true, color: C.white, align: 'center' });
    s.addShape(prs.ShapeType.roundRect, { x: x + 0.12, y: 1.82, w: 1.94, h: 1.7, rectRadius: 0.06,
      fill: { color: C.hero }, line: { color: C.bright, width: 0.3, transparency: 70 } });
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
  s.addText('Projected on 250K-Subscriber MENA Base', {
    x: 0.5, y: 0.50, w: 9, h: 0.54,
    fontSize: 22, bold: true, color: C.white,
  });

  // Left column header
  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 1.1, w: 4.35, h: 0.32, rectRadius: 0.05,
    fill: { color: C.red, transparency: 82 }, line: { color: C.red, width: 0.4, transparency: 50 } });
  s.addText('SUBSCRIBER EXPERIENCE', { x: 0.4, y: 1.1, w: 4.35, h: 0.32,
    fontSize: 8.5, bold: true, color: C.white, align: 'center' });

  const lMetrics = [
    { val: '−41%', lbl: 'Churn risk reduction per\nconverted subscriber', col: C.red },
    { val: '61%', lbl: 'Intervention conversion rate\n(flagged → plan switch)', col: C.bright },
    { val: 'AED 12', lbl: 'Average monthly saving\nfor switched subscriber', col: C.green },
    { val: '3.2×', lbl: 'Faster churn identification\nvs reactive support', col: C.orange },
  ];
  lMetrics.forEach((m, i) => {
    s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 1.52 + i * 0.74, w: 4.35, h: 0.66, rectRadius: 0.08,
      fill: { color: C.hero }, line: { color: m.col, width: 0.3, transparency: 65 } });
    s.addText(m.val, { x: 0.5, y: 1.56 + i * 0.74, w: 1.4, h: 0.58,
      fontSize: 24, bold: true, color: m.col, align: 'center', valign: 'middle', fontFace: 'Arial Black' });
    s.addText(m.lbl, { x: 1.95, y: 1.58 + i * 0.74, w: 2.65, h: 0.54,
      fontSize: 9, color: C.white, valign: 'middle' });
  });

  // Right column header
  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.1, w: 4.5, h: 0.32, rectRadius: 0.05,
    fill: { color: C.bright, transparency: 82 }, line: { color: C.bright, width: 0.4, transparency: 50 } });
  s.addText('OPERATOR ROI', { x: 5.1, y: 1.1, w: 4.5, h: 0.32,
    fontSize: 8.5, bold: true, color: C.white, align: 'center' });

  const rMetrics = [
    { val: 'AED 2.1M', lbl: 'Annual ARPU rescued\nper 250K-subscriber base', col: C.bright },
    { val: '8,420', lbl: 'High-risk subscribers\nflagged monthly', col: C.purple },
    { val: '+AED 8.40', lbl: 'ARPU uplift per\nconverted subscriber', col: C.green },
    { val: 'AED 0', lbl: 'Intervention cost — only\ntriggers on ML-confirmed risk', col: C.orange },
  ];
  rMetrics.forEach((m, i) => {
    s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.52 + i * 0.74, w: 4.5, h: 0.66, rectRadius: 0.08,
      fill: { color: C.hero }, line: { color: m.col, width: 0.3, transparency: 65 } });
    s.addText(m.val, { x: 5.2, y: 1.56 + i * 0.74, w: 1.9, h: 0.58,
      fontSize: i === 0 ? 17 : 22, bold: true, color: m.col, align: 'center', valign: 'middle', fontFace: 'Arial Black' });
    s.addText(m.lbl, { x: 7.15, y: 1.58 + i * 0.74, w: 2.2, h: 0.54,
      fontSize: 9, color: C.white, valign: 'middle' });
  });

  // Insight box
  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 4.62, w: 9.2, h: 0.70, rectRadius: 0.08,
    fill: { color: C.hero }, line: { color: C.bright, width: 0.5, transparency: 55 } });
  s.addText('Zero-cost retention: ChurnShield only intervenes when ML confirms high risk. The offer only upgrades or adds value — Circles gains ARPU and the subscriber gets a better plan.', {
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
  s.addText('I Built This Pipeline. Here\'s the Exact Mapping.', {
    x: 0.5, y: 0.50, w: 9, h: 0.54,
    fontSize: 22, bold: true, color: C.dark,
  });

  // Left panel — PhonePe proof
  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 1.08, w: 4.4, h: 4.0, rectRadius: 0.12,
    fill: { color: C.dark }, line: { color: C.bright, width: 0.8, transparency: 45 }, shadow: makeShadow() });
  s.addText('What I Shipped at PhonePe', { x: 0.55, y: 1.18, w: 4.1, h: 0.32,
    fontSize: 11, bold: true, color: C.bright });

  const lPoints = [
    { t:'Propensity-to-Transact ML:', b:'User-level scoring across 350M+ MAU — feature engineering, model deployment, closed-loop feedback' },
    { t:'−32% marketing spend:', b:'Real-time intent scoring replaced static cohort targeting — no rule, no manual update' },
    { t:'Dynamic incentivisation:', b:'Context-aware triggers (cart value × intent × time) → 35% AOV uplift, 60% abandonment drop' },
    { t:'Ops intelligence layer:', b:'Unified KPI dashboard replacing 6 analyst workflows — TAT 2 days → 4 hours' },
    { t:'A/B testing at scale:', b:'ML-driven vs rule-based control — ran on 350M MAU, pre-defined success metrics' },
  ];
  lPoints.forEach((pt, i) => {
    s.addText(`• ${pt.t}`, { x: 0.55, y: 1.60 + i * 0.66, w: 4.1, h: 0.22,
      fontSize: 9.5, bold: true, color: C.white });
    s.addText(`  ${pt.b}`, { x: 0.55, y: 1.82 + i * 0.66, w: 4.1, h: 0.38,
      fontSize: 8.5, color: C.muted });
  });

  // Right panel — JD mapping
  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.08, w: 4.5, h: 4.0, rectRadius: 0.12,
    fill: { color: 'FFFFFF' }, line: { color: C.green, width: 0.8, transparency: 45 }, shadow: makeShadow() });
  s.addText('How It Maps to Circles JD', { x: 5.25, y: 1.18, w: 4.2, h: 0.32,
    fontSize: 11, bold: true, color: C.dark });

  const rPoints = [
    { t:'ML/AI product delivery', b:'→ Propensity model production rollout, feature engineering to post-launch KPI' },
    { t:'Subscriber retention', b:'→ Milestone incentive: 60% 90-day device activation improvement at PhonePe' },
    { t:'ARPU & growth metrics', b:'→ Owned ARPU, GMV, CAC across 350M+ MAU — pre-defined, instrumented from Day 1' },
    { t:'Agile cross-functional', b:'→ Owned delivery across Tech, DS, Design, Compliance, Commercial simultaneously' },
    { t:'KPI & instrumentation', b:'→ Defined success metrics before every build — not retro-fitted post-launch' },
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
  s.addText('3-Phase Deployment — 6 Months to Full Rollout', {
    x: 0.5, y: 0.50, w: 9, h: 0.54,
    fontSize: 22, bold: true, color: C.white,
  });

  const phases = [
    {
      phase: 'Phase 1', time: 'Month 1–2',
      title: 'ML Baseline + Pilot',
      points: ['Train churn model on 90-day historical usage data', 'Define risk score thresholds (60/40) with DS team', 'Deploy to 5K subscribers — measure 30-day churn delta', 'A/B: ChurnShield intervention vs no-contact control'],
      col: C.bright,
    },
    {
      phase: 'Phase 2', time: 'Month 3–4',
      title: 'Intervention Engine',
      points: ['Build personalised recommendation ranker', 'Launch in-app alert + SMS nudge for high-risk segment', 'Tune thresholds based on pilot conversion data', 'Ops dashboard for retention team: funnel + ARPU tracking'],
      col: C.orange,
    },
    {
      phase: 'Phase 3', time: 'Month 5–6',
      title: 'Full MENA Rollout',
      points: ['Roll out to full subscriber base (250K)', 'Closed-loop retraining: post-intervention scores feed model', 'Expand to medium-risk (score 40–59) cohort', 'Quarterly model refresh + threshold recalibration'],
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
  s.addText('Access to 90-day subscriber usage data  ·  Alignment with retention & billing teams  ·  1 DS engineer for model pipeline  ·  1 frontend engineer for in-app alert', {
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
prs.writeFile({ fileName: 'assets/circles-churnshield-deck.pptx' }).then(() => {
  console.log('✅  circles-churnshield-deck.pptx built');
}).catch(e => { console.error(e); process.exit(1); });
