const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9'; // 10" x 5.625"

// ── Slide dimensions: W=10", H=5.625" ──────────────────────────────────
// Safe zone: y top=0.22", y bottom max=5.35"
// Content area: y=1.0" to y=4.65"
// Bottom insight box: y=4.7" to y=5.35"

const C = {
  dark:   '060D1A',
  hero:   '0D1826',
  bright: '00C6FF',
  accent: '3B82F6',
  green:  '22C55E',
  orange: 'F97316',
  white:  'FFFFFF',
  lgray:  'F1F5F9',
  muted:  '7090A0',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.10 });

// ─── SLIDE 1 — COVER ───────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  // Diagonal accent lines (clipped to slide width)
  for (let i = 0; i < 5; i++) {
    const lx = -0.5 + i * 2.2;
    const lw = Math.min(3.5, 10 - lx - 0.15);
    if (lx > 10) continue;
    s.addShape(prs.ShapeType.line, {
      x: lx, y: 0, w: lw, h: 5.625,
      line: { color: C.bright, width: 0.4 }, rotate: 30, transparency: 90,
    });
  }

  // Top tag
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 0.26, w: 3.0, h: 0.26, rectRadius: 0.05,
    fill: { color: C.bright, transparency: 88 }, line: { color: C.bright, width: 0.5, transparency: 50 } });
  s.addText('CIRCLES — MENA DIGITAL BRAND', { x: 0.5, y: 0.26, w: 3.0, h: 0.26,
    fontSize: 7, bold: true, color: C.bright, align: 'center', letterSpacingPt: 1 });

  // Product name
  s.addText('SmartPlan', { x: 0.5, y: 0.72, w: 6, h: 1.1,
    fontSize: 60, bold: true, color: C.white, fontFace: 'Arial' });

  // Cyan underline
  s.addShape(prs.ShapeType.rect, { x: 0.5, y: 1.72, w: 2.0, h: 0.05, fill: { color: C.bright } });

  // Subtitle
  s.addText('AI-Powered Plan Personalisation for Digital Telco Subscribers', {
    x: 0.5, y: 1.88, w: 7.2, h: 0.38,
    fontSize: 13, color: 'A0C8E0', fontFace: 'Arial',
  });

  // Presenter
  s.addText('Presented by Ajay Avaghade  ·  Product Manager — AI/ML & Consumer Growth', {
    x: 0.5, y: 2.36, w: 7.0, h: 0.26,
    fontSize: 9.5, color: '5A7A8A', fontFace: 'Arial',
  });

  // Stat box — bottom right
  s.addShape(prs.ShapeType.roundRect, { x: 7.1, y: 3.8, w: 2.65, h: 1.55, rectRadius: 0.1,
    fill: { color: C.hero }, line: { color: C.bright, width: 1, transparency: 45 } });
  s.addText('34%', { x: 7.1, y: 3.88, w: 2.65, h: 0.75,
    fontSize: 42, bold: true, color: C.orange, align: 'center', fontFace: 'Courier New' });
  s.addText('of subscribers on the\nwrong plan', { x: 7.1, y: 4.62, w: 2.65, h: 0.55,
    fontSize: 8.5, color: '7090A0', align: 'center', fontFace: 'Arial' });
}

// ─── SLIDE 2 — THE PROBLEM ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE PROBLEM', { x: 0.5, y: 0.22, w: 9, h: 0.22,
    fontSize: 8, bold: true, color: C.bright, letterSpacingPt: 1.5 });
  s.addText('One-Size Plans Drive Churn Before You Can See It Coming', {
    x: 0.5, y: 0.46, w: 9, h: 0.54, fontSize: 22, bold: true, color: C.white, fontFace: 'Arial' });

  const cols3 = [
    { icon: '📊', stat: '34%', label: 'Plan Mismatch Rate', sub: 'Subscribers on the wrong plan in mature digital telco markets' },
    { icon: '📉', stat: '3.2×', label: 'Higher Churn Risk', sub: 'Mismatched users churn at 3.2× the rate of plan-fit subscribers' },
    { icon: '💸', stat: '28%', label: 'Preventable Churn', sub: 'Of all churn is addressable with right-fit plan detection at the right moment' },
  ];
  cols3.forEach((col, i) => {
    const x = 0.5 + i * 3.1;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.15, w: 2.9, h: 2.8, rectRadius: 0.12,
      fill: { color: C.hero }, line: { color: C.bright, width: 0.5, transparency: 70 } });
    s.addText(col.icon, { x, y: 1.28, w: 2.9, h: 0.42, fontSize: 20, align: 'center' });
    s.addText(col.stat, { x, y: 1.72, w: 2.9, h: 0.68,
      fontSize: 38, bold: true, color: C.bright, align: 'center', fontFace: 'Courier New' });
    s.addText(col.label, { x, y: 2.4, w: 2.9, h: 0.3,
      fontSize: 11, bold: true, color: C.white, align: 'center', fontFace: 'Arial' });
    s.addText(col.sub, { x: x+0.15, y: 2.72, w: 2.6, h: 0.7,
      fontSize: 8.5, color: C.muted, align: 'center', fontFace: 'Arial' });
  });

  // Insight box
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 4.2, w: 9.0, h: 0.82, rectRadius: 0.08,
    fill: { color: '0A1A28' }, line: { color: C.bright, width: 0.8, transparency: 40 } });
  s.addText('Root cause: Static plan portfolios treat all subscribers identically. Without usage intelligence, mismatched subscribers silently churn — and winback costs 5× more than prevention.', {
    x: 0.7, y: 4.24, w: 8.6, h: 0.74, fontSize: 9.5, color: '90B0C0', italic: true, fontFace: 'Arial' });
}

// ─── SLIDE 3 — THE INSIGHT ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE INSIGHT', { x: 0.5, y: 0.22, w: 9, h: 0.22,
    fontSize: 8, bold: true, color: C.accent, letterSpacingPt: 1.5 });
  s.addText("Churn Isn't a Price Problem — It's a Plan Fit Problem", {
    x: 0.5, y: 0.46, w: 9, h: 0.52, fontSize: 22, bold: true, color: C.dark, fontFace: 'Arial' });

  // Left col
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 1.1, w: 4.0, h: 3.0, rectRadius: 0.1,
    fill: { color: 'FEF2F2' }, line: { color: 'FECACA', width: 1 } });
  s.addText('❌  Static Plans (Current)', { x: 0.6, y: 1.22, w: 3.8, h: 0.34,
    fontSize: 11, bold: true, color: 'B91C1C', fontFace: 'Arial' });
  [
    'Same plan options for all subscribers — no usage intelligence',
    'No signal until bill shock or cancellation request arrives',
    'Winback costs 5× more than proactive plan-fit intervention',
    'ARPU stagnates — upsell requires expensive outbound campaigns',
    'PM has zero visibility into plan mismatch cohorts or churn risk',
  ].forEach((t, i) => {
    s.addText(`• ${t}`, { x: 0.7, y: 1.66 + i * 0.44, w: 3.6, h: 0.38,
      fontSize: 9, color: '7F1D1D', fontFace: 'Arial' });
  });

  // VS
  s.addShape(prs.ShapeType.ellipse, { x: 4.38, y: 2.45, w: 0.48, h: 0.48,
    fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
  s.addText('VS', { x: 4.38, y: 2.45, w: 0.48, h: 0.48,
    fontSize: 8, bold: true, color: C.white, align: 'center', valign: 'middle' });

  // Right col
  s.addShape(prs.ShapeType.roundRect, { x: 4.8, y: 1.1, w: 4.7, h: 3.0, rectRadius: 0.1,
    fill: { color: 'F0FAFB' }, line: { color: '67E8F9', width: 1 } });
  s.addText('✅  SmartPlan (Proposed)', { x: 4.9, y: 1.22, w: 4.5, h: 0.34,
    fontSize: 11, bold: true, color: '0E7490', fontFace: 'Arial' });
  [
    'ML scores each subscriber\'s 90-day usage across 4 dimensions',
    'Plan Fit Score (0–100) surfaces at 80% data used and renewal',
    'Personalised recommendation with projected savings shown in-app',
    'Right-sized plan reduces churn risk 3.2× at zero winback cost',
    'Ops dashboard gives PM full cohort visibility and nudge control',
  ].forEach((t, i) => {
    s.addText(`• ${t}`, { x: 5.0, y: 1.66 + i * 0.44, w: 4.3, h: 0.38,
      fontSize: 9, color: '0C4A6E', fontFace: 'Arial' });
  });

  // Callout
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 4.26, w: 9, h: 0.78, rectRadius: 0.08,
    fill: { color: 'E0F7FA' }, line: { color: C.bright, width: 1, transparency: 30 } });
  s.addText('Key insight: The switch only triggers after the subscriber\'s own usage data validates the recommendation — zero pressure, maximum personalisation, subscriber stays in control.', {
    x: 0.7, y: 4.3, w: 8.6, h: 0.7, fontSize: 9.5, color: '0E4A6E', italic: true, fontFace: 'Arial' });
}

// ─── SLIDE 4 — THE MECHANIC ────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE MECHANIC', { x: 0.5, y: 0.22, w: 9, h: 0.22,
    fontSize: 8, bold: true, color: C.bright, letterSpacingPt: 1.5 });
  s.addText('How SmartPlan Works: 5 Steps', {
    x: 0.5, y: 0.46, w: 9, h: 0.52, fontSize: 22, bold: true, color: C.white, fontFace: 'Arial' });

  const steps = [
    { n:'01', title:'Data\nIngestion', body:'90 days: data usage, peak hours, roaming frequency, voice mins, add-on utilisation' },
    { n:'02', title:'ML\nScoring', body:'Plan Fit Score (0–100) + cohort tag: Over-Data, Roaming Candidate, Under-Used' },
    { n:'03', title:'Recommendation\nEngine', body:'Maps user to optimal plan; calculates projected ARPU delta and personalised savings' },
    { n:'04', title:'Smart Nudge\nTrigger', body:'SmartPlan card surfaces at 80% data, billing renewal, post-travel, or bill shock' },
    { n:'05', title:'Switch\n+ Track', body:'One-tap switch; 90-day outcome tracked against Plan Fit, ARPU delta, churn rate' },
  ];

  steps.forEach((st, i) => {
    const x = 0.42 + i * 1.85;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.1, w: 1.74, h: 3.0, rectRadius: 0.1,
      fill: { color: C.hero }, line: { color: C.bright, width: 0.6, transparency: 65 } });
    s.addText(st.n, { x, y: 1.2, w: 1.74, h: 0.46,
      fontSize: 24, bold: true, color: C.bright, align: 'center', fontFace: 'Courier New' });
    s.addText(st.title, { x: x+0.08, y: 1.72, w: 1.58, h: 0.5,
      fontSize: 9.5, bold: true, color: C.white, align: 'center', fontFace: 'Arial' });
    s.addText(st.body, { x: x+0.08, y: 2.3, w: 1.58, h: 1.65,
      fontSize: 8.5, color: C.muted, align: 'center', fontFace: 'Arial' });
  });

  // Proof box
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 4.26, w: 9, h: 0.8, rectRadius: 0.08,
    fill: { color: '0A1A28' }, line: { color: C.orange, width: 0.8, transparency: 45 } });
  s.addText('A/B proof at PhonePe: Behaviour-linked incentive (Rent Waiver) vs flat discount → 60% improvement in 90-day device activation. Same trigger-and-track architecture as SmartPlan\'s nudge engine.', {
    x: 0.7, y: 4.3, w: 8.6, h: 0.72, fontSize: 9.5, color: 'FDBA74', italic: true, fontFace: 'Arial' });
}

// ─── SLIDE 5 — THE PRODUCT ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE PRODUCT', { x: 0.5, y: 0.22, w: 9, h: 0.22,
    fontSize: 8, bold: true, color: C.accent, letterSpacingPt: 1.5 });
  s.addText('4 Key Screen States — Built for the Circles MENA App', {
    x: 0.5, y: 0.46, w: 9, h: 0.5, fontSize: 22, bold: true, color: C.dark, fontFace: 'Arial' });

  const screens = [
    {
      n:'01', title:'Current Plan Home',
      desc:'Data ring + Plan Fit Score (62/100) + SmartPlan banner when fit < 75 or data > 80%',
      mock:['[circles.life]   🔔','─────────────────','  Active Plan','  Base 20GB · AED 89','  ████████████░ 18.2GB','  Plan Fit: 62/100 ⚠','─────────────────','  💡 Better plan found','  [View SmartPlan →]'],
    },
    {
      n:'02', title:'SmartPlan Analyzer',
      desc:'90-day usage breakdown + ML recommendation card showing Plan Fit uplift 62→94',
      mock:['← SmartPlan Analysis','─────────────────','  Usage (90 days)','  Data: ████████ 19.1G','  Roam: █████░░ 3 ctry','  Voice: █░░░░░  42min','─────────────────','  ✦ Smart 30GB AED 89','  Fit: 62 → 94 ↑↑↑'],
    },
    {
      n:'03', title:'Switch Confirmed',
      desc:'Confetti + before/after: +10GB, free roaming, same price, Plan Fit 94/100',
      mock:['  🎉 Plan Switched!','─────────────────','  Circles Smart 30GB','  Activates midnight','─────────────────','  Data    20→30 GB ↑','  Roaming  No→10ct ↑','  Cost  AED 89→89 ✓','  Fit    62→94/100 ✓'],
    },
    {
      n:'04', title:'Ops Intelligence',
      desc:'PM view: mismatch cohorts, ARPU at risk (AED 2.1M), churn segments, nudge launcher',
      mock:['SmartPlan Intel','─────────────────',' Mismatch  34% ⚠',' Churn Risk 3.2× ⚠',' Avg Fit Score: 62','─────────────────',' Overage  8,420 usr',' Unused  12,300 usr',' ARPU@Risk AED 2.1M'],
    },
  ];

  screens.forEach((sc, i) => {
    const x = 0.38 + i * 2.33;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.08, w: 2.18, h: 4.3, rectRadius: 0.12,
      fill: { color: C.dark }, line: { color: C.bright, width: 0.6, transparency: 55 } });
    s.addText(sc.n, { x, y: 1.18, w: 2.18, h: 0.3,
      fontSize: 10, bold: true, color: C.bright, align: 'center', letterSpacingPt: 1 });
    s.addText(sc.title, { x: x+0.08, y: 1.5, w: 2.02, h: 0.3,
      fontSize: 9.5, bold: true, color: C.white, align: 'center', fontFace: 'Arial' });
    s.addShape(prs.ShapeType.roundRect, { x: x+0.12, y: 1.86, w: 1.94, h: 1.7, rectRadius: 0.06,
      fill: { color: C.hero }, line: { color: '1A2840', width: 0.5 } });
    s.addText(sc.mock.join('\n'), { x: x+0.14, y: 1.9, w: 1.9, h: 1.62,
      fontSize: 5.8, color: '90B8C8', fontFace: 'Courier New', valign: 'top' });
    s.addText(sc.desc, { x: x+0.08, y: 3.62, w: 2.02, h: 0.92,
      fontSize: 8, color: '4A6A7A', align: 'center', fontFace: 'Arial' });
  });
}

// ─── SLIDE 6 — IMPACT & ROI ────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('IMPACT & ROI', { x: 0.5, y: 0.22, w: 9, h: 0.22,
    fontSize: 8, bold: true, color: C.bright, letterSpacingPt: 1.5 });
  s.addText('Projected Impact — Built on PhonePe Proof Points', {
    x: 0.5, y: 0.46, w: 9, h: 0.52, fontSize: 22, bold: true, color: C.white, fontFace: 'Arial' });

  // Left header
  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 1.1, w: 4.35, h: 0.32, rectRadius: 0.05,
    fill: { color: C.hero }, line: { color: C.bright, width: 0.5, transparency: 50 } });
  s.addText('SUBSCRIBER IMPACT', { x: 0.4, y: 1.1, w: 4.35, h: 0.32,
    fontSize: 8, bold: true, color: C.bright, align: 'center', letterSpacingPt: 1 });

  const leftM = [
    { val: '↑ 28%', lbl: 'Right-fit plan adoption within 90 days of SmartPlan launch' },
    { val: '↓ 31%', lbl: 'Churn rate in mismatched subscriber cohort post-intervention' },
    { val: '94/100', lbl: 'Average Plan Fit Score after switch (vs 62 baseline)' },
    { val: '< 4 min', lbl: 'Time from recommendation surfaced to plan switch confirmed' },
  ];
  leftM.forEach((m, i) => {
    s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 1.52 + i * 0.74, w: 4.35, h: 0.66, rectRadius: 0.08,
      fill: { color: C.hero }, line: { color: '1A3050', width: 0.5 } });
    s.addText(m.val, { x: 0.5, y: 1.56 + i * 0.74, w: 1.3, h: 0.58,
      fontSize: 18, bold: true, color: C.bright, align: 'center', fontFace: 'Courier New', valign: 'middle' });
    s.addText(m.lbl, { x: 1.95, y: 1.6 + i * 0.74, w: 2.65, h: 0.5,
      fontSize: 9, color: C.muted, fontFace: 'Arial', valign: 'middle' });
  });

  // Right header
  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.1, w: 4.5, h: 0.32, rectRadius: 0.05,
    fill: { color: '0A1C14' }, line: { color: C.green, width: 0.5, transparency: 50 } });
  s.addText('BUSINESS ROI', { x: 5.1, y: 1.1, w: 4.5, h: 0.32,
    fontSize: 8, bold: true, color: C.green, align: 'center', letterSpacingPt: 1 });

  const rightM = [
    { val: '+AED 8.40', lbl: 'ARPU uplift per switched subscriber per month via fit-based upsell' },
    { val: '3.2× ↓', lbl: 'Churn risk reduction in targeted cohort vs non-intervened control' },
    { val: 'AED 2.1M+', lbl: 'Annual ARPU at risk rescued per 250K-subscriber MENA base' },
    { val: '↓ 22%', lbl: 'Reduction in data-cap support contacts within 60 days of launch' },
  ];
  rightM.forEach((m, i) => {
    s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.52 + i * 0.74, w: 4.5, h: 0.66, rectRadius: 0.08,
      fill: { color: '080F08' }, line: { color: '0F2A1A', width: 0.5 } });
    s.addText(m.val, { x: 5.15, y: 1.56 + i * 0.74, w: 1.55, h: 0.58,
      fontSize: 16, bold: true, color: C.green, align: 'center', fontFace: 'Courier New', valign: 'middle' });
    s.addText(m.lbl, { x: 6.8, y: 1.6 + i * 0.74, w: 2.65, h: 0.5,
      fontSize: 9, color: '4A7A5A', fontFace: 'Arial', valign: 'middle' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 4.6, w: 9.2, h: 0.7, rectRadius: 0.08,
    fill: { color: '0A1A28' }, line: { color: C.bright, width: 0.8, transparency: 40 } });
  s.addText('SmartPlan only triggers after the subscriber\'s usage data validates the mismatch — zero-cost retention. ARPU uplift is a by-product of right-fit, not aggressive upselling.', {
    x: 0.6, y: 4.64, w: 9.0, h: 0.62, fontSize: 9.5, color: '90B0C0', italic: true, fontFace: 'Arial' });
}

// ─── SLIDE 7 — PROOF OF WORK ───────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('PROOF OF WORK', { x: 0.5, y: 0.22, w: 9, h: 0.22,
    fontSize: 8, bold: true, color: C.accent, letterSpacingPt: 1.5 });
  s.addText("I've Built Every Piece of This. Here's the Evidence.", {
    x: 0.5, y: 0.46, w: 9, h: 0.52, fontSize: 22, bold: true, color: C.dark, fontFace: 'Arial' });

  // Left dark panel
  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 1.08, w: 4.4, h: 4.0, rectRadius: 0.12,
    fill: { color: C.dark }, line: { color: C.bright, width: 0.5, transparency: 60 } });
  s.addText('What I shipped at PhonePe', { x: 0.55, y: 1.2, w: 4.1, h: 0.3,
    fontSize: 11, bold: true, color: C.bright, fontFace: 'Arial' });

  const leftItems = [
    { t:'ML Propensity-to-Transact models', d:'Replaced ₹1,000+ Cr/yr of static targeting → 32% spend reduction. Same scoring engine as Plan Fit Score.' },
    { t:'Context-aware trigger engine', d:'Cart × intent × time-of-day signals → 35% AOV uplift. Same architecture as SmartPlan nudge layer.' },
    { t:'Ops intelligence dashboard', d:'Replaced 6 analyst workflows → TAT 2 days to 4 hours. Same as SmartPlan Intelligence cohort view.' },
    { t:'Milestone A/B test framework', d:'Behaviour-linked incentive vs flat discount → 60% device activation uplift. Same BAT structure.' },
  ];
  leftItems.forEach((item, i) => {
    s.addText(`▸ ${item.t}`, { x: 0.6, y: 1.64 + i * 0.8, w: 4.0, h: 0.28,
      fontSize: 9.5, bold: true, color: C.white, fontFace: 'Arial' });
    s.addText(item.d, { x: 0.65, y: 1.94 + i * 0.8, w: 3.9, h: 0.4,
      fontSize: 8.5, color: C.muted, fontFace: 'Arial' });
  });

  // Right light panel
  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.08, w: 4.5, h: 4.0, rectRadius: 0.12,
    fill: { color: 'F0F8FB' }, line: { color: C.accent, width: 0.5, transparency: 50 } });
  s.addText('How it maps to this role', { x: 5.25, y: 1.2, w: 4.2, h: 0.3,
    fontSize: 11, bold: true, color: C.accent, fontFace: 'Arial' });

  const rightItems = [
    { t:'User-level ML scoring', d:'→ Plan Fit Score (0–100) replacing static plan assignment for MENA subscribers' },
    { t:'Funnel performance tracking', d:'→ KPI ownership across acquisition, plan fit, churn, and ARPU with pre-launch instrumentation' },
    { t:'Cross-functional Agile delivery', d:'→ Product, engineering, DS, commercial, and CS alignment for phased MENA rollout' },
    { t:'Data-to-recommendation loop', d:'→ 90-day usage → Plan Fit Score → personalised nudge → switch outcome tracking' },
  ];
  rightItems.forEach((item, i) => {
    s.addText(`▸ ${item.t}`, { x: 5.25, y: 1.64 + i * 0.8, w: 4.2, h: 0.28,
      fontSize: 9.5, bold: true, color: C.dark, fontFace: 'Arial' });
    s.addText(item.d, { x: 5.3, y: 1.94 + i * 0.8, w: 4.1, h: 0.4,
      fontSize: 8.5, color: '3A5A6A', fontFace: 'Arial' });
  });
}

// ─── SLIDE 8 — ROLLOUT ─────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('ROLLOUT PLAN', { x: 0.5, y: 0.22, w: 9, h: 0.22,
    fontSize: 8, bold: true, color: C.bright, letterSpacingPt: 1.5 });
  s.addText('Phased Delivery — MENA Digital Brand Launch', {
    x: 0.5, y: 0.46, w: 9, h: 0.52, fontSize: 22, bold: true, color: C.white, fontFace: 'Arial' });

  const phases = [
    {
      label:'Phase 1  ·  Month 1–2', title:'Pilot · 5K Subscribers',
      items:['Instrument Plan Fit Score on 5,000 MENA beta users','Baseline churn rate, ARPU, and support ticket volume','A/B test SmartPlan banner vs control (no intervention)','Validate 90-day usage data pipeline with ML team'],
      color: C.bright,
    },
    {
      label:'Phase 2  ·  Month 3–4', title:'Scale · 50K Users',
      items:['Expand to 50K; A/B test proactive vs reactive nudge','Tune Plan Fit thresholds using pilot churn data','Launch Ops Intelligence dashboard for PM & commercial','Instrument ARPU delta and 90-day churn outcomes'],
      color: C.accent,
    },
    {
      label:'Phase 3  ·  Month 5–6', title:'Full Rollout',
      items:['Deploy to full MENA base with winning nudge variant','Integrate Plan Fit Score into new subscriber onboarding','Launch upsell module as SmartPlan premium migration path','Publish ARPU and churn impact report for investors'],
      color: C.green,
    },
  ];

  phases.forEach((ph, i) => {
    const x = 0.42 + i * 3.12;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.1, w: 2.95, h: 2.92, rectRadius: 0.1,
      fill: { color: C.hero }, line: { color: ph.color, width: 1, transparency: 50 } });
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.1, w: 2.95, h: 0.52, rectRadius: 0.1,
      fill: { color: ph.color, transparency: 82 }, line: { color: ph.color, width: 0.5, transparency: 40 } });
    s.addText(ph.label, { x, y: 1.14, w: 2.95, h: 0.24,
      fontSize: 8, bold: true, color: ph.color, align: 'center', letterSpacingPt: 0.5 });
    s.addText(ph.title, { x, y: 1.38, w: 2.95, h: 0.24,
      fontSize: 10, bold: true, color: C.white, align: 'center', fontFace: 'Arial' });
    ph.items.forEach((it, j) => {
      s.addText(`• ${it}`, { x: x+0.14, y: 1.74 + j * 0.5, w: 2.65, h: 0.44,
        fontSize: 8.5, color: C.muted, fontFace: 'Arial' });
    });
  });

  // Ask box
  s.addShape(prs.ShapeType.roundRect, { x: 0.42, y: 4.2, w: 6.5, h: 0.95, rectRadius: 0.08,
    fill: { color: '0A1A28' }, line: { color: C.bright, width: 0.8, transparency: 40 } });
  s.addText('What I need to build this:', { x: 0.62, y: 4.28, w: 6.3, h: 0.26,
    fontSize: 9.5, bold: true, color: C.bright, fontFace: 'Arial' });
  s.addText('Subscriber usage data access  ·  Commercial/plan team alignment  ·  1 ML engineer  ·  iOS/Android feature slot', {
    x: 0.62, y: 4.56, w: 6.3, h: 0.46, fontSize: 9, color: '90B0C0', fontFace: 'Arial' });

  // Contact
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade', {
    x: 0.4, y: 5.3, w: 9.2, h: 0.25, fontSize: 8.5, color: '405060', align: 'center', fontFace: 'Arial' });
}

// ─── BUILD ─────────────────────────────────────────────────────────────
prs.writeFile({ fileName: 'assets/circles-smartplan-deck.pptx' })
  .then(() => console.log('✅  circles-smartplan-deck.pptx rebuilt'))
  .catch(e => { console.error(e); process.exit(1); });
