const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '060D1A',
  hero:   '0D1826',
  bright: '00C6FF',
  accent: '3B82F6',
  purple: '6366F1',
  orange: 'F97316',
  green:  '22C55E',
  white:  'FFFFFF',
  lgray:  'F1F5F9',
  muted:  '64748B',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });

// ─── SLIDE 1 — COVER ───────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  // Diagonal accent lines (decorative)
  for (let i = 0; i < 6; i++) {
    s.addShape(prs.ShapeType.line, {
      x: -0.5 + i * 2, y: 0, w: 4, h: 7.5,
      line: { color: C.bright, width: 0.5, dashType: 'solid' },
      rotate: 30,
      transparency: 88,
    });
  }

  // Top tag
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 0.35, w: 3.2, h: 0.32, rectRadius: 0.05,
    fill: { color: C.bright, transparency: 88 }, line: { color: C.bright, width: 0.5, transparency: 50 } });
  s.addText('CIRCLES — MENA DIGITAL BRAND', { x: 0.5, y: 0.35, w: 3.2, h: 0.32,
    fontSize: 8, bold: true, color: C.bright, align: 'center', letterSpacingPt: 1 });

  // Product name
  s.addText('SmartPlan', { x: 0.5, y: 1.1, w: 6, h: 1.4,
    fontSize: 72, bold: true, color: C.white, fontFace: 'Arial' });

  // Cyan underline
  s.addShape(prs.ShapeType.rect, { x: 0.5, y: 2.35, w: 2.2, h: 0.06, fill: { color: C.bright } });

  // Subtitle
  s.addText('AI-Powered Plan Personalisation for Digital Telco Subscribers', {
    x: 0.5, y: 2.6, w: 7, h: 0.5,
    fontSize: 16, color: 'A0C8E0', bold: false, fontFace: 'Arial',
  });

  // Presenter
  s.addText('Presented by Ajay Avaghade  ·  Product Manager — AI/ML & Consumer Growth', {
    x: 0.5, y: 3.2, w: 7, h: 0.3,
    fontSize: 11, color: '5A7A8A', fontFace: 'Arial',
  });

  // Big stat box (bottom right)
  s.addShape(prs.ShapeType.roundRect, { x: 6.8, y: 5.8, w: 2.8, h: 1.4, rectRadius: 0.1,
    fill: { color: C.hero }, line: { color: C.bright, width: 1, transparency: 50 } });
  s.addText('34%', { x: 6.8, y: 5.85, w: 2.8, h: 0.7,
    fontSize: 40, bold: true, color: C.orange, align: 'center', fontFace: 'Arial', fontFace: 'Courier New' });
  s.addText('of subscribers on the wrong plan', { x: 6.8, y: 6.55, w: 2.8, h: 0.5,
    fontSize: 9, color: '7090A0', align: 'center', fontFace: 'Arial' });
}

// ─── SLIDE 2 — THE PROBLEM ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE PROBLEM', { x: 0.5, y: 0.35, w: 9, h: 0.25,
    fontSize: 9, bold: true, color: C.bright, letterSpacingPt: 1.5 });
  s.addText('One-Size Plans Drive Churn Before You Can See It Coming', {
    x: 0.5, y: 0.65, w: 9, h: 0.7,
    fontSize: 28, bold: true, color: C.white, fontFace: 'Arial',
  });

  const cols3 = [
    { icon: '📊', stat: '34%', label: 'Plan Mismatch Rate', sub: 'Subscribers on the wrong plan in mature digital telco markets' },
    { icon: '📉', stat: '3.2×', label: 'Higher Churn Risk', sub: 'Mismatched users churn at 3.2× the rate of plan-fit subscribers' },
    { icon: '💸', stat: '28%', label: 'Preventable Churn', sub: 'Of all churn is addressable with right-fit plan detection at the right moment' },
  ];
  cols3.forEach((col, i) => {
    const x = 0.5 + i * 3.1;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.6, w: 2.9, h: 3.4, rectRadius: 0.12,
      fill: { color: C.hero }, line: { color: C.bright, width: 0.5, transparency: 70 } });
    s.addText(col.icon, { x, y: 1.75, w: 2.9, h: 0.5, fontSize: 24, align: 'center' });
    s.addText(col.stat, { x, y: 2.35, w: 2.9, h: 0.8,
      fontSize: 44, bold: true, color: C.bright, align: 'center', fontFace: 'Courier New' });
    s.addText(col.label, { x, y: 3.15, w: 2.9, h: 0.35,
      fontSize: 12, bold: true, color: C.white, align: 'center', fontFace: 'Arial' });
    s.addText(col.sub, { x: x + 0.15, y: 3.55, w: 2.6, h: 0.9,
      fontSize: 9.5, color: '7090A0', align: 'center', fontFace: 'Arial' });
  });

  // Insight box
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 5.25, w: 9.0, h: 0.95, rectRadius: 0.08,
    fill: { color: '0A1A28' }, line: { color: C.bright, width: 0.8, transparency: 40 } });
  s.addText('Root cause: Static plan portfolios treat all subscribers identically. Without usage-based intelligence, mismatched subscribers silently churn — and winback costs 5× more than prevention.', {
    x: 0.7, y: 5.3, w: 8.6, h: 0.85,
    fontSize: 10, color: '90B0C0', italic: true, fontFace: 'Arial',
  });
}

// ─── SLIDE 3 — THE INSIGHT ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE INSIGHT', { x: 0.5, y: 0.35, w: 9, h: 0.25,
    fontSize: 9, bold: true, color: C.accent, letterSpacingPt: 1.5 });
  s.addText("Churn Isn't a Price Problem — It's a Plan Fit Problem", {
    x: 0.5, y: 0.65, w: 9, h: 0.65,
    fontSize: 26, bold: true, color: C.dark, fontFace: 'Arial',
  });

  // Left column — current
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 1.55, w: 4.0, h: 3.6, rectRadius: 0.1,
    fill: { color: 'FEF2F2' }, line: { color: 'FECACA', width: 1 } });
  s.addText('❌  Static Plans (Current)', { x: 0.6, y: 1.7, w: 3.8, h: 0.4,
    fontSize: 13, bold: true, color: 'B91C1C', fontFace: 'Arial' });
  [
    'Same plan options for all users regardless of usage',
    'No intelligence surfaced until bill shock or cancellation',
    'Merchant pays 5× acquisition cost to win back churned user',
    'ARPU stays flat — upsell requires manual outreach',
    'PM has no visibility into plan mismatch cohorts',
  ].forEach((t, i) => {
    s.addText(`• ${t}`, { x: 0.7, y: 2.25 + i * 0.52, w: 3.6, h: 0.45,
      fontSize: 10, color: '7F1D1D', fontFace: 'Arial', bullet: false });
  });

  // VS divider
  s.addShape(prs.ShapeType.ellipse, { x: 4.38, y: 2.9, w: 0.54, h: 0.54,
    fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
  s.addText('VS', { x: 4.38, y: 2.9, w: 0.54, h: 0.54,
    fontSize: 9, bold: true, color: C.white, align: 'center', valign: 'middle' });

  // Right column — SmartPlan
  s.addShape(prs.ShapeType.roundRect, { x: 4.8, y: 1.55, w: 4.7, h: 3.6, rectRadius: 0.1,
    fill: { color: 'F0FAFB' }, line: { color: '67E8F9', width: 1 } });
  s.addText('✅  SmartPlan (Proposed)', { x: 4.9, y: 1.7, w: 4.5, h: 0.4,
    fontSize: 13, bold: true, color: '0E7490', fontFace: 'Arial' });
  [
    'ML scores each subscriber\'s usage pattern across 90 days',
    'Plan Fit Score (0–100) surfaces at 80% data usage & renewal',
    'Personalised recommendation with projected savings shown in-app',
    'Right-sized plan reduces churn risk 3.2× at zero winback cost',
    'Ops dashboard exposes plan mismatch cohorts for PM team',
  ].forEach((t, i) => {
    s.addText(`• ${t}`, { x: 5.0, y: 2.25 + i * 0.52, w: 4.3, h: 0.45,
      fontSize: 10, color: '0C4A6E', fontFace: 'Arial' });
  });

  // Callout
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 5.3, w: 9, h: 0.85, rectRadius: 0.08,
    fill: { color: 'E0F7FA' }, line: { color: C.bright, width: 1, transparency: 30 } });
  s.addText('Key insight: The switch only triggers after the subscriber\'s own usage data validates the recommendation — zero pressure sales, maximum personalisation, subscriber keeps control.', {
    x: 0.7, y: 5.35, w: 8.6, h: 0.75,
    fontSize: 10, color: '0E4A6E', italic: true, fontFace: 'Arial',
  });
}

// ─── SLIDE 4 — THE MECHANIC ────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE MECHANIC', { x: 0.5, y: 0.35, w: 9, h: 0.25,
    fontSize: 9, bold: true, color: C.bright, letterSpacingPt: 1.5 });
  s.addText('How SmartPlan Works: 5 Steps', {
    x: 0.5, y: 0.65, w: 9, h: 0.55,
    fontSize: 28, bold: true, color: C.white, fontFace: 'Arial',
  });

  const steps = [
    { n:'01', title:'Data Ingestion', body:'90 days: data consumption, peak-hour patterns, roaming frequency, voice usage, add-on utilisation rate' },
    { n:'02', title:'ML Scoring', body:'Propensity model outputs Plan Fit Score (0–100) + cohort tag: Over-Data, Under-Data, Roaming Candidate, Voice-Heavy' },
    { n:'03', title:'Recommendation Engine', body:'Maps user to optimal plan from current portfolio; calculates projected ARPU delta and personalised savings narrative' },
    { n:'04', title:'Smart Nudge Trigger', body:'Surfaces SmartPlan card at high-intent moments: 80% data used, billing renewal, post-travel session, bill shock event' },
    { n:'05', title:'Switch + Track', body:'One-tap plan switch; 90-day outcome tracked against Plan Fit improvement, ARPU delta, and churn rate vs control cohort' },
  ];

  steps.forEach((st, i) => {
    const x = 0.4 + i * 1.86;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.45, w: 1.75, h: 3.6, rectRadius: 0.1,
      fill: { color: C.hero }, line: { color: C.bright, width: 0.6, transparency: 65 } });
    s.addText(st.n, { x, y: 1.6, w: 1.75, h: 0.55,
      fontSize: 28, bold: true, color: C.bright, align: 'center', fontFace: 'Courier New' });
    s.addText(st.title, { x: x + 0.1, y: 2.25, w: 1.55, h: 0.5,
      fontSize: 11, bold: true, color: C.white, align: 'center', fontFace: 'Arial' });
    s.addText(st.body, { x: x + 0.08, y: 2.85, w: 1.6, h: 1.9,
      fontSize: 9, color: '7090A0', align: 'center', fontFace: 'Arial' });

    // Dashed arrow connector
    if (i < 4) {
      s.addShape(prs.ShapeType.line, { x: x + 1.75, y: 3.2, w: 0.11, h: 0,
        line: { color: C.bright, width: 1, dashType: 'dash', transparency: 50 } });
    }
  });

  // A/B note
  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 5.25, w: 9, h: 0.95, rectRadius: 0.08,
    fill: { color: '0A1A28' }, line: { color: C.orange, width: 0.8, transparency: 50 } });
  s.addText('A/B proof at PhonePe: Behaviour-linked incentive variant (Rent Waiver) vs flat discount → 60% improvement in 90-day device activation. Same trigger architecture as SmartPlan\'s nudge engine.', {
    x: 0.7, y: 5.3, w: 8.6, h: 0.85,
    fontSize: 10, color: 'FDBA74', italic: true, fontFace: 'Arial',
  });
}

// ─── SLIDE 5 — THE PRODUCT ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE PRODUCT', { x: 0.5, y: 0.35, w: 9, h: 0.25,
    fontSize: 9, bold: true, color: C.accent, letterSpacingPt: 1.5 });
  s.addText('4 Key Screen States — Designed for the Circles MENA App', {
    x: 0.5, y: 0.65, w: 9, h: 0.55,
    fontSize: 26, bold: true, color: C.dark, fontFace: 'Arial',
  });

  const screens = [
    {
      n: '01', title: 'Current Plan Home',
      desc: 'Data usage ring + Plan Fit Score (62/100) + SmartPlan banner surfaced when fit score < 75 or data > 80% used',
      mock: ['[circles.life]      🔔', '─────────────────────', '  Active Plan', '  Circles Base 20GB', '  ████████████░░ 18.2/20GB', '  Plan Fit: 62/100 ⚠', '─────────────────────', '  💡 Better plan found', '     Same price · +10GB', '  [View SmartPlan →]  '],
    },
    {
      n: '02', title: 'SmartPlan Analyzer',
      desc: '90-day usage breakdown with ML-scored recommendation card showing Plan Fit uplift from 62 → 94',
      mock: ['← SmartPlan Analysis', '─────────────────────', '  Usage (90 days)', '  Data:   ████████ 19.1GB', '  Roaming: █████░░ 3 ctry', '  Voice:  █░░░░░░  42min', '─────────────────────', '  ✦ Circles Smart 30GB', '  AED 89/mo · +10GB + Roam', '  Fit: 62 → 94 ↑↑↑  '],
    },
    {
      n: '03', title: 'Plan Switch Confirmed',
      desc: 'Confetti celebration + before/after comparison showing zero price change with 10GB uplift and free roaming',
      mock: ['  🎉 Plan Switched!    ', '─────────────────────', '  Circles Smart 30GB  ', '  activates midnight  ', '─────────────────────', '  Data    20GB → 30GB ', '  Roaming   No → 10ct ', '  Cost  AED 89 → 89 ✓ ', '  Fit Score 62 → 94 ✓ ', '  [← Back to Home]   '],
    },
    {
      n: '04', title: 'Ops Intelligence',
      desc: 'PM dashboard: plan mismatch cohorts, ARPU at risk (AED 2.1M), churn risk segments, one-click nudge launcher',
      mock: ['SmartPlan Intelligence', '─────────────────────', ' Mismatch: 34% ⚠ +4pp', ' Churn Risk: 3.2× ⚠ ', ' Avg Fit Score: 62   ', '─────────────────────', ' Data Overage   8,420', ' Unused Add-ons 12.3K', ' Travel Cands    3.8K', ' ARPU at Risk AED 2.1M'],
    },
  ];

  screens.forEach((sc, i) => {
    const x = 0.35 + i * 2.35;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.45, w: 2.2, h: 4.0, rectRadius: 0.12,
      fill: { color: C.dark }, line: { color: C.bright, width: 0.6, transparency: 55 } });
    s.addText(sc.n, { x, y: 1.58, w: 2.2, h: 0.35,
      fontSize: 11, bold: true, color: C.bright, align: 'center', letterSpacingPt: 1 });
    s.addText(sc.title, { x: x + 0.08, y: 1.95, w: 2.04, h: 0.35,
      fontSize: 10.5, bold: true, color: C.white, align: 'center', fontFace: 'Arial' });
    s.addShape(prs.ShapeType.roundRect, { x: x + 0.12, y: 2.38, w: 1.96, h: 1.85, rectRadius: 0.06,
      fill: { color: C.hero }, line: { color: '1A2840', width: 0.5 } });
    s.addText(sc.mock.join('\n'), { x: x + 0.14, y: 2.42, w: 1.92, h: 1.78,
      fontSize: 6.2, color: '90B8C8', fontFace: 'Courier New', valign: 'top' });
    s.addText(sc.desc, { x: x + 0.08, y: 4.3, w: 2.04, h: 1.0,
      fontSize: 8.5, color: '4A6A7A', align: 'center', fontFace: 'Arial' });
  });
}

// ─── SLIDE 6 — IMPACT & ROI ────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('IMPACT & ROI', { x: 0.5, y: 0.35, w: 9, h: 0.25,
    fontSize: 9, bold: true, color: C.bright, letterSpacingPt: 1.5 });
  s.addText('Projected Impact — Built on PhonePe Proof Points', {
    x: 0.5, y: 0.65, w: 9, h: 0.55,
    fontSize: 28, bold: true, color: C.white, fontFace: 'Arial',
  });

  // Left block
  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 1.45, w: 4.35, h: 0.4, rectRadius: 0.05,
    fill: { color: C.hero }, line: { color: C.bright, width: 0.5, transparency: 50 } });
  s.addText('SUBSCRIBER IMPACT', { x: 0.4, y: 1.45, w: 4.35, h: 0.4,
    fontSize: 9, bold: true, color: C.bright, align: 'center', letterSpacingPt: 1 });

  const leftMetrics = [
    { val: '↑ 28%', lbl: 'Right-fit plan adoption within 90 days of SmartPlan launch' },
    { val: '↓ 31%', lbl: 'Churn rate in mismatched subscriber cohort post-intervention' },
    { val: '94/100', lbl: 'Average Plan Fit Score after SmartPlan switch (vs 62 baseline)' },
    { val: '< 4 min', lbl: 'Time from SmartPlan recommendation surfaced to plan switch confirmed' },
  ];
  leftMetrics.forEach((m, i) => {
    s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 2.0 + i * 0.88, w: 4.35, h: 0.78, rectRadius: 0.08,
      fill: { color: C.hero }, line: { color: '1A3050', width: 0.5 } });
    s.addText(m.val, { x: 0.55, y: 2.05 + i * 0.88, w: 1.3, h: 0.68,
      fontSize: 22, bold: true, color: C.bright, align: 'center', fontFace: 'Courier New', valign: 'middle' });
    s.addText(m.lbl, { x: 2.0, y: 2.1 + i * 0.88, w: 2.6, h: 0.58,
      fontSize: 9.5, color: '7090A0', fontFace: 'Arial', valign: 'middle' });
  });

  // Right block
  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.45, w: 4.5, h: 0.4, rectRadius: 0.05,
    fill: { color: '0A1C14' }, line: { color: C.green, width: 0.5, transparency: 50 } });
  s.addText('BUSINESS ROI', { x: 5.1, y: 1.45, w: 4.5, h: 0.4,
    fontSize: 9, bold: true, color: C.green, align: 'center', letterSpacingPt: 1 });

  const rightMetrics = [
    { val: '+AED 8.40', lbl: 'ARPU uplift per switched subscriber per month via upsell to better-fit plans' },
    { val: '3.2× ↓', lbl: 'Churn risk reduction in targeted mismatched cohort vs non-intervened control' },
    { val: 'AED 2.1M+', lbl: 'Annual ARPU at risk rescued per 250K-subscriber MENA base' },
    { val: '↓ 22%', lbl: 'Reduction in data-cap-related customer support contacts within 60 days' },
  ];
  rightMetrics.forEach((m, i) => {
    s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 2.0 + i * 0.88, w: 4.5, h: 0.78, rectRadius: 0.08,
      fill: { color: '080F08' }, line: { color: '0F2A1A', width: 0.5 } });
    s.addText(m.val, { x: 5.15, y: 2.05 + i * 0.88, w: 1.55, h: 0.68,
      fontSize: 18, bold: true, color: C.green, align: 'center', fontFace: 'Courier New', valign: 'middle' });
    s.addText(m.lbl, { x: 6.8, y: 2.1 + i * 0.88, w: 2.65, h: 0.58,
      fontSize: 9.5, color: '4A7A5A', fontFace: 'Arial', valign: 'middle' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 5.52, w: 9.2, h: 0.68, rectRadius: 0.08,
    fill: { color: '0A1A28' }, line: { color: C.bright, width: 0.8, transparency: 40 } });
  s.addText('SmartPlan only recommends a switch after the subscriber\'s own usage data confirms the mismatch — making this zero-cost retention. The ARPU uplift is a by-product of fit, not aggressive upselling.', {
    x: 0.6, y: 5.56, w: 9.0, h: 0.6,
    fontSize: 9.5, color: '90B0C0', italic: true, fontFace: 'Arial',
  });
}

// ─── SLIDE 7 — PROOF OF WORK ───────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('PROOF OF WORK', { x: 0.5, y: 0.35, w: 9, h: 0.25,
    fontSize: 9, bold: true, color: C.accent, letterSpacingPt: 1.5 });
  s.addText("I've Built Every Piece of This. Here's the Evidence.", {
    x: 0.5, y: 0.65, w: 9, h: 0.55,
    fontSize: 24, bold: true, color: C.dark, fontFace: 'Arial',
  });

  // Left dark panel
  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 1.4, w: 4.4, h: 4.7, rectRadius: 0.12,
    fill: { color: C.dark }, line: { color: C.bright, width: 0.5, transparency: 60 } });
  s.addText('What I shipped at PhonePe', { x: 0.55, y: 1.55, w: 4.1, h: 0.35,
    fontSize: 12, bold: true, color: C.bright, fontFace: 'Arial' });

  const leftItems = [
    { t: 'ML Propensity-to-Transact models', d: 'Replaced ₹1,000+ Cr/yr of static targeting → 32% spend reduction, sustained GMV. Same scoring engine as Plan Fit Score.' },
    { t: 'Context-aware trigger engine', d: 'Cart value × user intent × time-of-day signals for dynamic incentivisation → 35% AOV uplift. Same architecture as SmartPlan nudge layer.' },
    { t: 'Ops intelligence dashboard', d: 'Replaced 6 analyst workflows → TAT 2 days to 4 hours. Same as SmartPlan Intelligence cohort dashboard.' },
    { t: 'Milestone A/B test framework', d: 'Behaviour-linked incentive vs flat discount → 60% device activation improvement. SmartPlan recommendation uses same BAT structure.' },
  ];
  leftItems.forEach((item, i) => {
    s.addText(`▸ ${item.t}`, { x: 0.6, y: 2.1 + i * 0.92, w: 4.0, h: 0.3,
      fontSize: 10.5, bold: true, color: C.white, fontFace: 'Arial' });
    s.addText(item.d, { x: 0.65, y: 2.42 + i * 0.92, w: 3.9, h: 0.45,
      fontSize: 8.5, color: '7090A0', fontFace: 'Arial' });
  });

  // Right light panel
  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.4, w: 4.5, h: 4.7, rectRadius: 0.12,
    fill: { color: 'F0F8FB' }, line: { color: C.accent, width: 0.5, transparency: 50 } });
  s.addText('How it maps to this JD', { x: 5.25, y: 1.55, w: 4.2, h: 0.35,
    fontSize: 12, bold: true, color: C.accent, fontFace: 'Arial' });

  const rightItems = [
    { t: 'User-level ML scoring', d: '→ Plan Fit Score engine (0–100) replacing static plan assignment for MENA subscribers' },
    { t: 'Funnel performance tracking', d: '→ KPI ownership across acquisition, plan fit, churn, and ARPU tracked in Ops Intelligence' },
    { t: 'Cross-functional Agile delivery', d: '→ Product, engineering, DS, commercial, and CS alignment for SmartPlan phased rollout' },
    { t: 'Data-to-recommendation loop', d: '→ 90-day usage data → Plan Fit Score → personalised nudge → switch outcome tracking' },
  ];
  rightItems.forEach((item, i) => {
    s.addText(`▸ ${item.t}`, { x: 5.25, y: 2.1 + i * 0.92, w: 4.2, h: 0.3,
      fontSize: 10.5, bold: true, color: C.dark, fontFace: 'Arial' });
    s.addText(item.d, { x: 5.3, y: 2.42 + i * 0.92, w: 4.1, h: 0.45,
      fontSize: 8.5, color: '3A5A6A', fontFace: 'Arial' });
  });
}

// ─── SLIDE 8 — ROLLOUT PLAN ────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('ROLLOUT PLAN', { x: 0.5, y: 0.35, w: 9, h: 0.25,
    fontSize: 9, bold: true, color: C.bright, letterSpacingPt: 1.5 });
  s.addText('Phased Delivery — MENA Digital Brand Launch', {
    x: 0.5, y: 0.65, w: 9, h: 0.55,
    fontSize: 28, bold: true, color: C.white, fontFace: 'Arial',
  });

  const phases = [
    {
      phase: 'Phase 1',
      period: 'Month 1–2',
      title: 'Pilot · 5K Subscribers',
      items: [
        'Instrument Plan Fit Score on 5,000 MENA beta users',
        'Baseline churn rate, ARPU, and support ticket volume',
        'A/B test SmartPlan banner vs control (no banner)',
        'Validate 90-day usage data pipeline with ML team',
      ],
      color: C.bright,
    },
    {
      phase: 'Phase 2',
      period: 'Month 3–4',
      title: 'Scale · 50K Users',
      items: [
        'Roll SmartPlan to 50K users; A/B: proactive nudge vs reactive trigger',
        'Tune Plan Fit Score thresholds based on pilot churn data',
        'Launch Ops Intelligence dashboard for PM and commercial team',
        'Instrument ARPU delta and 90-day churn outcomes',
      ],
      color: C.accent,
    },
    {
      phase: 'Phase 3',
      period: 'Month 5–6',
      title: 'Full Rollout',
      items: [
        'Deploy to full MENA subscriber base with winning nudge variant',
        'Integrate Plan Fit Score into onboarding flow for new subscribers',
        'Launch upsell module: SmartPlan as premium plan migration path',
        'Publish ARPU and churn impact report for board/investors',
      ],
      color: C.green,
    },
  ];

  phases.forEach((ph, i) => {
    const x = 0.4 + i * 3.12;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.45, w: 2.95, h: 3.45, rectRadius: 0.1,
      fill: { color: C.hero }, line: { color: ph.color, width: 1, transparency: 50 } });
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.45, w: 2.95, h: 0.62, rectRadius: 0.1,
      fill: { color: ph.color, transparency: 80 }, line: { color: ph.color, width: 0.5, transparency: 40 } });
    s.addText(`${ph.phase}  ·  ${ph.period}`, { x, y: 1.5, w: 2.95, h: 0.28,
      fontSize: 9, bold: true, color: ph.color, align: 'center', letterSpacingPt: 0.5 });
    s.addText(ph.title, { x, y: 1.78, w: 2.95, h: 0.28,
      fontSize: 11, bold: true, color: C.white, align: 'center', fontFace: 'Arial' });
    ph.items.forEach((it, j) => {
      s.addText(`• ${it}`, { x: x + 0.15, y: 2.2 + j * 0.58, w: 2.65, h: 0.5,
        fontSize: 9, color: '7090A0', fontFace: 'Arial' });
    });
  });

  // The Ask
  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 5.1, w: 6.3, h: 1.1, rectRadius: 0.08,
    fill: { color: '0A1A28' }, line: { color: C.bright, width: 0.8, transparency: 40 } });
  s.addText('What I need to build this:', { x: 0.6, y: 5.16, w: 6.1, h: 0.28,
    fontSize: 10, bold: true, color: C.bright, fontFace: 'Arial' });
  s.addText('Access to subscriber usage data  ·  Alignment with commercial/plan team  ·  1 ML engineer  ·  iOS/Android feature slot', {
    x: 0.6, y: 5.46, w: 6.1, h: 0.6,
    fontSize: 9.5, color: '90B0C0', fontFace: 'Arial',
  });

  // Contact footer
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade', {
    x: 0.4, y: 6.8, w: 9.2, h: 0.35,
    fontSize: 9, color: '405060', align: 'center', fontFace: 'Arial',
  });
}

// ─── BUILD ─────────────────────────────────────────────────────────────
prs.writeFile({ fileName: 'assets/circles-smartplan-deck.pptx' })
  .then(() => console.log('✅  circles-smartplan-deck.pptx created'))
  .catch(e => { console.error(e); process.exit(1); });
