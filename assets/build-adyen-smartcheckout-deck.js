const pptx = require('pptxgenjs');
const prs = new pptx();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0D1F12',
  hero:   '1A3323',
  bright: '00BB59',
  orange: 'F97316',
  amber:  'F59E0B',
  white:  'FFFFFF',
  lgray:  'F4F7F4',
  muted:  '8B9E92',
  ink:    '0F172A',
  slate:  '475569',
  red:    'EF4444',
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 3, offset: 1, angle: 45, opacity: 0.12 });

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 1 — COVER (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  // Diagonal stripe accents
  for (let i = 0; i < 6; i++) {
    slide.addShape(prs.ShapeType.rect, {
      x: -0.5 + i * 2.2, y: -0.2, w: 0.18, h: 10,
      fill: { color: C.bright, transparency: 88 },
      line: { type: 'none' },
      rotate: 15,
    });
  }

  // Adyen green accent bar top
  slide.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 0.06,
    fill: { color: C.bright }, line: { type: 'none' },
  });

  // Company tag
  slide.addText('ADYEN — EMEA PAYMENTS PM', {
    x: 0.6, y: 0.28, w: 6, h: 0.28,
    fontSize: 9, bold: true, color: C.bright,
    charSpacing: 3, fontFace: 'Calibri',
  });

  // Product name
  slide.addText('SmartCheckout', {
    x: 0.6, y: 0.72, w: 8, h: 1.1,
    fontSize: 56, bold: true, color: C.white, fontFace: 'Calibri',
  });

  // Subtitle
  slide.addText('ML-powered EMEA payment method ranking engine\nthat maximises checkout conversion across 14 markets', {
    x: 0.6, y: 1.82, w: 7.4, h: 0.72,
    fontSize: 15, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
  });

  // Presenter
  slide.addText('Presented by Ajay Avaghade  ·  Product Manager — EMEA Payments', {
    x: 0.6, y: 2.68, w: 6, h: 0.3,
    fontSize: 11, color: C.muted, fontFace: 'Calibri',
  });

  // Big stat box bottom right
  slide.addShape(prs.ShapeType.rect, {
    x: 7.2, y: 1.6, w: 2.4, h: 1.9,
    fill: { color: C.bright, transparency: 15 },
    line: { color: C.bright, pt: 1.5 },
    rectRadius: 0.1,
  });
  slide.addText('+12.4%', {
    x: 7.2, y: 1.72, w: 2.4, h: 0.72,
    fontSize: 36, bold: true, color: C.bright, align: 'center', fontFace: 'Calibri',
  });
  slide.addText('checkout conversion\nlift projected\nacross EMEA', {
    x: 7.2, y: 2.44, w: 2.4, h: 0.72,
    fontSize: 9.5, color: C.muted, align: 'center', fontFace: 'Calibri', lineSpacingMultiple: 1.3,
  });

  // Bottom bar
  slide.addShape(prs.ShapeType.rect, {
    x: 0, y: 7.08, w: 10, h: 0.42,
    fill: { color: C.hero }, line: { type: 'none' },
  });
  slide.addText('Adyen — Proprietary & Confidential  ·  Prepared by Ajay Avaghade  ·  2026', {
    x: 0.3, y: 7.1, w: 9.4, h: 0.3,
    fontSize: 8, color: C.muted, fontFace: 'Calibri',
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  slide.addText('THE PROBLEM', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('EMEA Checkout Is Losing Billions to Static Payment Method Rules', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.9,
    fontSize: 26, bold: true, color: C.white, fontFace: 'Calibri', lineSpacingMultiple: 1.2,
  });

  // 3 stat columns
  const stats = [
    { icon: '📉', num: '61%', lbl: 'Checkout abandonment in Germany', sub: 'vs 89% conversion in the Netherlands — same PSP, different market rules' },
    { icon: '💸', num: '€1.8M', lbl: 'GMV abandoned at H&M Germany checkout', sub: 'PayPal ranked 4th by static rules — but converts 3× better than Visa for German users' },
    { icon: '🌍', num: '14', lbl: 'EMEA markets, each with distinct LPM preferences', sub: 'iDEAL, Bancontact, Bizum, Giropay — one static config cannot fit all' },
  ];
  stats.forEach((s, i) => {
    const x = 0.5 + i * 3.1;
    slide.addShape(prs.ShapeType.rect, {
      x, y: 1.72, w: 2.85, h: 3.4,
      fill: { color: C.hero }, line: { color: C.bright, pt: 1, transparency: 60 }, rectRadius: 0.1,
      shadow: makeShadow(),
    });
    slide.addText(s.icon, { x, y: 1.88, w: 2.85, h: 0.44, fontSize: 22, align: 'center' });
    slide.addText(s.num, {
      x, y: 2.36, w: 2.85, h: 0.72,
      fontSize: 38, bold: true, color: C.bright, align: 'center', fontFace: 'Calibri',
    });
    slide.addText(s.lbl, {
      x: x + 0.14, y: 3.12, w: 2.6, h: 0.52,
      fontSize: 11, bold: true, color: C.white, align: 'center', fontFace: 'Calibri', lineSpacingMultiple: 1.2,
    });
    slide.addText(s.sub, {
      x: x + 0.1, y: 3.68, w: 2.68, h: 1.1,
      fontSize: 9, color: C.muted, align: 'center', fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });

  // Insight box
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 5.38, w: 9, h: 1.18,
    fill: { color: C.hero }, line: { color: C.orange, pt: 1.5, transparency: 40 }, rectRadius: 0.1,
  });
  slide.addText('Root cause: Adyen routes EMEA checkout to payment methods using static merchant configuration rules — the same LPM stack regardless of user locale, device, cart value, or purchase intent. ML-ranked, session-aware LPM selection is the gap no competitor has closed at scale.', {
    x: 0.72, y: 5.5, w: 8.6, h: 0.94,
    fontSize: 10.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.4,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 3 — THE INSIGHT (light)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.lgray };

  slide.addText('THE INSIGHT', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.slate, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('Static Rules Maximise PSP Coverage. ML Ranking Maximises Conversion.', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.82,
    fontSize: 24, bold: true, color: C.ink, fontFace: 'Calibri', lineSpacingMultiple: 1.2,
  });

  // VS divider oval
  slide.addShape(prs.ShapeType.ellipse, {
    x: 4.6, y: 2.3, w: 0.8, h: 0.8,
    fill: { color: C.dark }, line: { type: 'none' },
  });
  slide.addText('VS', {
    x: 4.6, y: 2.36, w: 0.8, h: 0.62,
    fontSize: 11, bold: true, color: C.white, align: 'center', fontFace: 'Calibri',
  });

  // LEFT column — current
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.6, w: 4.0, h: 4.2,
    fill: { color: 'FEF2F2' }, line: { color: 'EF4444', pt: 1.5, transparency: 40 }, rectRadius: 0.1,
  });
  slide.addText('❌  Static Config (Today)', {
    x: 0.68, y: 1.76, w: 3.68, h: 0.4,
    fontSize: 13, bold: true, color: C.ink, fontFace: 'Calibri',
  });
  const leftPts = [
    'One LPM stack per merchant — same order for every user in every market',
    'PayPal ranked below iDEAL in Germany even though PayPal converts 3× better',
    'Configuration changes require merchant dev work and Adyen account manager approval',
    'Merchants discover conversion gaps only through quarterly analytics reviews',
    'No signal on why users drop off — just that they do',
  ];
  leftPts.forEach((pt, i) => {
    slide.addText('• ' + pt, {
      x: 0.68, y: 2.22 + i * 0.62, w: 3.6, h: 0.54,
      fontSize: 10.5, color: C.slate, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
    });
  });

  // RIGHT column — proposed
  slide.addShape(prs.ShapeType.rect, {
    x: 5.5, y: 1.6, w: 4.0, h: 4.2,
    fill: { color: 'F0FDF4' }, line: { color: C.bright, pt: 1.5, transparency: 30 }, rectRadius: 0.1,
  });
  slide.addText('✅  SmartCheckout (Proposed)', {
    x: 5.68, y: 1.76, w: 3.68, h: 0.4,
    fontSize: 13, bold: true, color: C.ink, fontFace: 'Calibri',
  });
  const rightPts = [
    'ML ranks LPMs per session using 23 signals — locale, device, cart value, time, history',
    'German user gets PayPal first; Dutch user gets iDEAL first — automatically',
    'Config happens in Adyen dashboard in real time — no merchant dev work required',
    'Live conversion analytics per method per market — see the gap and close it today',
    'A/B test any ranking hypothesis against baseline in 48 hours',
  ];
  rightPts.forEach((pt, i) => {
    slide.addText('• ' + pt, {
      x: 5.68, y: 2.22 + i * 0.62, w: 3.6, h: 0.54,
      fontSize: 10.5, color: C.slate, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
    });
  });

  // Callout box
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.0, w: 9, h: 0.96,
    fill: { color: C.dark }, line: { type: 'none' }, rectRadius: 0.1,
  });
  slide.addText('Key insight: The ranking cost to Adyen is near zero — SmartCheckout uses signals already collected per transaction. The uplift is pure incremental conversion. By the time a merchant sees the dashboard, the ML model has already outperformed their static config on day one.', {
    x: 0.72, y: 6.1, w: 8.6, h: 0.76,
    fontSize: 10, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.4,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 4 — THE MECHANIC (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  slide.addText('THE MECHANIC', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('How SmartCheckout Works: 23 Signals, 1 Ranked LPM List, Real Time', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.8,
    fontSize: 23, bold: true, color: C.white, fontFace: 'Calibri', lineSpacingMultiple: 1.2,
  });

  const steps = [
    { n: '01', title: 'Session Signal Collection', desc: 'On checkout load, Adyen collects 23 signals: browser locale, IP geolocation, device type, cart value, product category, time of day, user payment history, and 15 behavioural signals.' },
    { n: '02', title: 'ML Corridor Scoring', desc: 'Gradient-boosted model scores each available LPM against the session vector. Output: ranked probability of successful payment completion per method. Model retrained weekly on 2B+ global Adyen transactions.' },
    { n: '03', title: 'Real-Time LPM Ranking', desc: 'Top 3 LPMs surface in ranked order at checkout. German session: PayPal 87%, Giropay 74%, Visa 68%. Dutch session: iDEAL 93%, Mastercard 71%, PayPal 52%. No merchant config change required.' },
    { n: '04', title: 'Merchant Dashboard + A/B Testing', desc: 'Merchant sees live conversion rate per LPM per market. A/B test any ranking hypothesis vs baseline in 48 hours. Override ML ranking for specific campaigns (e.g., promote new LPM launch).' },
    { n: '05', title: 'Continuous Model Learning', desc: 'Each completed / abandoned payment feeds back into the model. Over 30 days the model improves conversion by an estimated +2–4pp as it learns merchant-specific user behaviour patterns.' },
  ];

  steps.forEach((s, i) => {
    const y = 1.6 + i * 0.98;
    // Step number circle
    slide.addShape(prs.ShapeType.ellipse, {
      x: 0.5, y: y, w: 0.48, h: 0.48,
      fill: { color: C.bright }, line: { type: 'none' },
    });
    slide.addText(s.n, {
      x: 0.5, y: y + 0.02, w: 0.48, h: 0.38,
      fontSize: 9, bold: true, color: C.white, align: 'center', fontFace: 'Calibri',
    });
    // Connector line (skip last)
    if (i < 4) {
      slide.addShape(prs.ShapeType.rect, {
        x: 0.72, y: y + 0.48, w: 0.04, h: 0.5,
        fill: { color: C.bright, transparency: 60 }, line: { type: 'none' },
      });
    }
    slide.addText(s.title, {
      x: 1.18, y: y, w: 8.2, h: 0.28,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri',
    });
    slide.addText(s.desc, {
      x: 1.18, y: y + 0.28, w: 8.2, h: 0.62,
      fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });

  // A/B proof note
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.52, w: 9, h: 0.64,
    fill: { color: C.hero }, line: { color: C.orange, pt: 1, transparency: 40 }, rectRadius: 0.08,
  });
  slide.addText('PhonePe Proof: Propensity-to-Transact ML model deployed across 350M users replaced static rule-based targeting — same multi-variable scoring architecture. Drove 32% marketing burn reduction while sustaining GMV growth.', {
    x: 0.72, y: 6.6, w: 8.6, h: 0.5,
    fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 5 — THE PRODUCT (light)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.lgray };

  slide.addText('THE PRODUCT', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.slate, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('4 Key Screen States — Merchant Portal + Consumer Checkout', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 22, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  const screens = [
    {
      num: '01', title: 'Merchant Dashboard',
      desc: 'EMEA conversion KPIs by market. SmartCheckout enable CTA. Benchmarks vs Adyen top quartile. Red flags on markets below 70% benchmark.',
      mock: '[ EMEA Conversion Map ]\n NL 89% ✓  DE 61% ⚠\n ES 54% ⚠  BE 58% ⚠\n[Enable SmartCheckout →]',
    },
    {
      num: '02', title: 'Config + ML Ranking',
      desc: 'Market selector chips. Live ML LPM ranking per market. Confidence score (94.2%). Override toggle for campaign promotions.',
      mock: '[ Germany Config ]\n 1. PayPal    87%\n 2. Giropay   74%\n 3. Visa      68%\n 4. iDEAL     31%\n[ML Confidence 94.2%]',
    },
    {
      num: '03', title: 'Consumer Checkout',
      desc: 'German user detects, PayPal ranked first. Frictionless payment. Success state with SmartCheckout attribution badge.',
      mock: '[ H&M Checkout ]\n Detected: German user\n 1. PayPal     ← Ranked\n 2. Giropay\n 3. Visa\n[Pay €189.00 →]',
    },
    {
      num: '04', title: 'Analytics Dashboard',
      desc: 'Before/after conversion by market. GMV recovered attribution. Network effect insight — improvement compounds across all merchants as model learns.',
      mock: '[ SmartCheckout Impact ]\n DE: +26% conversion\n ES: +31% conversion\n BE: +22% conversion\n GMV recovered: €1.8M',
    },
  ];

  screens.forEach((s, i) => {
    const x = 0.32 + i * 2.36;
    slide.addShape(prs.ShapeType.rect, {
      x, y: 1.4, w: 2.2, h: 5.3,
      fill: { color: C.white }, line: { color: C.bright, pt: 1.5, transparency: 40 }, rectRadius: 0.1,
      shadow: makeShadow(),
    });
    slide.addShape(prs.ShapeType.rect, {
      x, y: 1.4, w: 2.2, h: 0.36,
      fill: { color: C.dark }, line: { type: 'none' }, rectRadius: 0.1,
    });
    slide.addText(s.num + '  ' + s.title, {
      x: x + 0.1, y: 1.44, w: 2.05, h: 0.28,
      fontSize: 8.5, bold: true, color: C.bright, fontFace: 'Calibri',
    });
    // Mockup block
    slide.addShape(prs.ShapeType.rect, {
      x: x + 0.1, y: 1.86, w: 2.02, h: 2.3,
      fill: { color: 'F0FDF4' }, line: { color: C.bright, pt: 0.75, transparency: 60 }, rectRadius: 0.08,
    });
    slide.addText(s.mock, {
      x: x + 0.14, y: 1.9, w: 1.96, h: 2.2,
      fontSize: 7.5, color: C.slate, fontFace: 'Courier New', lineSpacingMultiple: 1.45,
    });
    slide.addText(s.desc, {
      x: x + 0.1, y: 4.24, w: 2.02, h: 2.2,
      fontSize: 9, color: C.slate, fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });

  slide.addText('Live interactive prototype: adyen-smartcheckout-prototype.html', {
    x: 0.5, y: 6.88, w: 9, h: 0.28,
    fontSize: 8.5, color: C.slate, italic: true, fontFace: 'Calibri', align: 'center',
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  slide.addText('IMPACT & ROI', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('Projected Impact — Built on PhonePe ML Proof Points', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 24, bold: true, color: C.white, fontFace: 'Calibri',
  });

  // Left section label
  slide.addText('MERCHANT IMPACT', {
    x: 0.5, y: 1.38, w: 4.4, h: 0.28,
    fontSize: 8.5, bold: true, color: C.bright, charSpacing: 2, fontFace: 'Calibri',
  });
  const leftMetrics = [
    { val: '+12.4%', lbl: 'Checkout conversion lift', sub: 'Average across DE, ES, BE in 90-day pilot window' },
    { val: '€1.8M', lbl: 'Abandoned GMV recovered', sub: 'H&M Germany pilot extrapolated to full EMEA portfolio' },
    { val: '+26%', lbl: 'Conversion in Germany', sub: 'Highest-gap market; PayPal moved from 4th to 1st rank' },
    { val: '−35%', lbl: 'Checkout abandonment rate', sub: 'Users now see preferred LPM first, reducing friction at step 1' },
  ];
  leftMetrics.forEach((m, i) => {
    const y = 1.72 + i * 1.22;
    slide.addShape(prs.ShapeType.rect, {
      x: 0.5, y, w: 4.3, h: 1.08,
      fill: { color: C.hero }, line: { color: C.bright, pt: 0.75, transparency: 50 }, rectRadius: 0.08,
    });
    slide.addText(m.val, {
      x: 0.68, y: y + 0.1, w: 1.6, h: 0.52,
      fontSize: 28, bold: true, color: C.bright, fontFace: 'Calibri',
    });
    slide.addText(m.lbl, {
      x: 2.32, y: y + 0.1, w: 2.3, h: 0.28,
      fontSize: 11, bold: true, color: C.white, fontFace: 'Calibri',
    });
    slide.addText(m.sub, {
      x: 2.32, y: y + 0.42, w: 2.3, h: 0.54,
      fontSize: 8.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
    });
  });

  // Right section label
  slide.addText('ADYEN REVENUE IMPACT', {
    x: 5.1, y: 1.38, w: 4.4, h: 0.28,
    fontSize: 8.5, bold: true, color: C.orange, charSpacing: 2, fontFace: 'Calibri',
  });
  const rightMetrics = [
    { val: '+€14M', lbl: 'Annual incremental processing fee', sub: 'Based on 12.4% GMV uplift across 847 EMEA merchants at avg 0.3% Adyen take rate' },
    { val: '×3.1', lbl: 'ROI on SmartCheckout investment', sub: 'Estimated 18-month engineering + data investment vs fee uplift at 12 months' },
    { val: '−40%', lbl: 'Merchant churn in year 2', sub: 'Merchants with visible conversion attribution in dashboard renew at higher rates' },
    { val: '92%', lbl: 'STP rate maintained post-rollout', sub: 'ML ranking does not add checkout steps — conversion gain is frictionless' },
  ];
  rightMetrics.forEach((m, i) => {
    const y = 1.72 + i * 1.22;
    slide.addShape(prs.ShapeType.rect, {
      x: 5.1, y, w: 4.3, h: 1.08,
      fill: { color: C.hero }, line: { color: C.orange, pt: 0.75, transparency: 50 }, rectRadius: 0.08,
    });
    slide.addText(m.val, {
      x: 5.28, y: y + 0.1, w: 1.6, h: 0.52,
      fontSize: 28, bold: true, color: C.orange, fontFace: 'Calibri',
    });
    slide.addText(m.lbl, {
      x: 6.92, y: y + 0.1, w: 2.3, h: 0.28,
      fontSize: 11, bold: true, color: C.white, fontFace: 'Calibri',
    });
    slide.addText(m.sub, {
      x: 6.92, y: y + 0.42, w: 2.3, h: 0.54,
      fontSize: 8.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
    });
  });

  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.56, w: 9, h: 0.6,
    fill: { color: C.hero }, line: { color: C.bright, pt: 1, transparency: 50 }, rectRadius: 0.08,
  });
  slide.addText('SmartCheckout is a zero-friction uplift — it runs entirely within the existing checkout flow. For Adyen, this is incremental revenue on infrastructure already deployed, with no additional merchant integration required.', {
    x: 0.72, y: 6.64, w: 8.6, h: 0.44,
    fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK (light)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.lgray };

  slide.addText('PROOF OF WORK', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.slate, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('I Built This Architecture. At PhonePe. At 350M Users.', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 23, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  // Left — dark column
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.38, w: 4.3, h: 5.3,
    fill: { color: C.dark }, line: { type: 'none' }, rectRadius: 0.1,
  });
  slide.addText('What I Built at PhonePe', {
    x: 0.7, y: 1.56, w: 3.9, h: 0.34,
    fontSize: 12, bold: true, color: C.bright, fontFace: 'Calibri',
  });
  const leftPts = [
    'Deployed Propensity-to-Transact ML model replacing ₹1,000+ Cr/yr of static rule-based marketing spend — same multi-variable scoring architecture as SmartCheckout',
    'Built context-aware dynamic incentivisation engine using cart value, user intent, and time-of-day signals to trigger personalised offers — exact same trigger architecture as session-aware LPM ranking',
    'Led A/B testing framework across 350M MAU to validate ML vs rule-based targeting — sustained GMV growth while cutting marketing burn 32%',
    'Owned cross-functional delivery across Technology, Data Science, Compliance, Finance — the same stakeholder map required for SmartCheckout EMEA rollout',
  ];
  leftPts.forEach((pt, i) => {
    slide.addText('• ' + pt, {
      x: 0.7, y: 2.0 + i * 1.08, w: 3.9, h: 0.96,
      fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });

  // Right — light column
  slide.addShape(prs.ShapeType.rect, {
    x: 5.1, y: 1.38, w: 4.3, h: 5.3,
    fill: { color: C.white }, line: { color: C.bright, pt: 1.5, transparency: 40 }, rectRadius: 0.1,
    shadow: makeShadow(),
  });
  slide.addText('How It Maps to the Adyen JD', {
    x: 5.3, y: 1.56, w: 3.9, h: 0.34,
    fontSize: 12, bold: true, color: C.ink, fontFace: 'Calibri',
  });
  const rightPts = [
    ['ML Product Management', 'Propensity models at 350M scale → SmartCheckout EMEA LPM scoring engine'],
    ['Payment Product Strategy', 'Led checkout conversion redesign end-to-end → EMEA LPM product ownership'],
    ['Cross-functional Delivery', 'Tech / Data Science / Compliance / Finance → same map at Adyen EMEA'],
    ['A/B Experimentation', 'ML vs rule-based A/B at 350M MAU → SmartCheckout vs baseline per market'],
  ];
  rightPts.forEach((pt, i) => {
    const y = 2.0 + i * 1.08;
    slide.addText(pt[0], {
      x: 5.3, y, w: 3.9, h: 0.28,
      fontSize: 10.5, bold: true, color: C.ink, fontFace: 'Calibri',
    });
    slide.addText(pt[1], {
      x: 5.3, y: y + 0.28, w: 3.9, h: 0.68,
      fontSize: 9.5, color: C.slate, fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });

  // Quote box
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.82, w: 9, h: 0.0,
    fill: { type: 'none' }, line: { type: 'none' },
  });
  // (no extra quote box needed, fits on slide)
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  slide.addText('ROLLOUT PLAN', {
    x: 0.5, y: 0.3, w: 9, h: 0.28,
    fontSize: 9, bold: true, color: C.muted, charSpacing: 4, fontFace: 'Calibri',
  });
  slide.addText('Phased EMEA Rollout — Signal-First, Then Scale', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 24, bold: true, color: C.white, fontFace: 'Calibri',
  });

  const phases = [
    {
      phase: 'Phase 1',
      time: 'Month 1–2 · Pilot',
      title: 'Germany + Netherlands (3 merchants)',
      pts: [
        'Deploy SmartCheckout to 3 EMEA merchants in DE + NL',
        'Instrument LPM conversion tracking per session vector',
        'Run ML ranking vs static config A/B — 50/50 split',
        'Measure 30-day conversion lift vs control group',
      ],
      color: C.bright,
    },
    {
      phase: 'Phase 2',
      time: 'Month 3–4 · Scale',
      title: '6 Markets + A/B Hardening',
      pts: [
        'Expand to ES, BE, FR, PL with market-specific LPM models',
        'Launch Merchant Dashboard with live per-market analytics',
        'Tune signal weights based on pilot conversion data',
        'Onboard top 50 EMEA merchants by GMV volume',
      ],
      color: C.orange,
    },
    {
      phase: 'Phase 3',
      time: 'Month 5–6 · Full Rollout',
      title: 'All 14 EMEA Markets',
      pts: [
        'Roll out winning ML configuration to all 847 EMEA merchants',
        'Integrate SmartCheckout into Adyen onboarding flow as default-on',
        'Launch self-serve override controls for campaign promotions',
        'Publish EMEA Checkout Benchmark Report for press / merchant acquisition',
      ],
      color: '#A78BFA',
    },
  ];

  phases.forEach((p, i) => {
    const x = 0.5 + i * 3.1;
    slide.addShape(prs.ShapeType.rect, {
      x, y: 1.42, w: 2.88, h: 4.5,
      fill: { color: C.hero }, line: { color: p.color, pt: 1.5, transparency: 30 }, rectRadius: 0.1,
    });
    slide.addText(p.phase, {
      x: x + 0.14, y: 1.58, w: 2.62, h: 0.28,
      fontSize: 9, bold: true, color: p.color, charSpacing: 2, fontFace: 'Calibri',
    });
    slide.addText(p.time, {
      x: x + 0.14, y: 1.86, w: 2.62, h: 0.28,
      fontSize: 10, color: C.muted, fontFace: 'Calibri', italic: true,
    });
    slide.addText(p.title, {
      x: x + 0.14, y: 2.18, w: 2.62, h: 0.44,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri', lineSpacingMultiple: 1.2,
    });
    p.pts.forEach((pt, j) => {
      slide.addText('• ' + pt, {
        x: x + 0.14, y: 2.68 + j * 0.74, w: 2.62, h: 0.66,
        fontSize: 9, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
      });
    });
  });

  // Ask box
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.1, w: 9, h: 1.02,
    fill: { color: C.hero }, line: { color: C.bright, pt: 1.5, transparency: 30 }, rectRadius: 0.1,
  });
  slide.addText('What I need to build this:', {
    x: 0.72, y: 6.2, w: 3, h: 0.28,
    fontSize: 10, bold: true, color: C.bright, fontFace: 'Calibri',
  });
  slide.addText('Access to EMEA checkout transaction data  ·  Alignment with Data Science team on signal schema  ·  1 backend engineer for ML inference API integration  ·  1 designer for merchant dashboard  ·  Adyen Issuing + Acquiring team for STP data', {
    x: 0.72, y: 6.5, w: 8.6, h: 0.54,
    fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
  });

  // Contact footer
  slide.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade', {
    x: 0.5, y: 7.2, w: 9, h: 0.26,
    fontSize: 8.5, color: C.muted, fontFace: 'Calibri', align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/adyen-smartcheckout-deck.pptx' })
  .then(() => console.log('adyen-smartcheckout-deck.pptx generated'))
  .catch(e => console.error(e));
