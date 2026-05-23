const pptx = require('pptxgenjs');
const prs = new pptx();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0D1F12',
  hero:   '1A3323',
  teal:   '0D9488',
  tlight: '14B8A6',
  orange: 'F97316',
  amber:  'F59E0B',
  red:    'EF4444',
  green:  '22C55E',
  white:  'FFFFFF',
  lgray:  'F4F7F4',
  muted:  '8B9E92',
  ink:    '0F172A',
  slate:  '475569',
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 3, offset: 1, angle: 45, opacity: 0.12 });

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 1 — COVER (dark)
// ─────────────────────────────────────────────────────────────────────────────
{
  const slide = prs.addSlide();
  slide.background = { color: C.dark };

  // Diagonal accent lines
  for (let i = 0; i < 6; i++) {
    slide.addShape(prs.ShapeType.rect, {
      x: -0.5 + i * 2.2, y: -0.2, w: 0.16, h: 10,
      fill: { color: C.teal, transparency: 88 },
      line: { type: 'none' },
      rotate: 15,
    });
  }

  // Top accent bar
  slide.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 0.06,
    fill: { color: C.teal }, line: { type: 'none' },
  });

  slide.addText('ADYEN — EMEA PAYMENTS PM', {
    x: 0.6, y: 0.28, w: 6, h: 0.28,
    fontSize: 9, bold: true, color: C.tlight,
    charSpacing: 3, fontFace: 'Calibri',
  });

  slide.addText('RegSense', {
    x: 0.6, y: 0.72, w: 8, h: 1.1,
    fontSize: 56, bold: true, color: C.white, fontFace: 'Calibri',
  });

  slide.addText('PSD3/PSR3 compliance intelligence dashboard that gives\nAdyen merchants a 60-day early warning system for regulatory change', {
    x: 0.6, y: 1.82, w: 7.4, h: 0.76,
    fontSize: 14, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.35,
  });

  slide.addText('Presented by Ajay Avaghade  ·  Product Manager — EMEA Payments', {
    x: 0.6, y: 2.72, w: 6, h: 0.3,
    fontSize: 11, color: C.muted, fontFace: 'Calibri',
  });

  // Big stat box
  slide.addShape(prs.ShapeType.rect, {
    x: 7.2, y: 1.6, w: 2.4, h: 1.9,
    fill: { color: C.teal, transparency: 85 },
    line: { color: C.teal, pt: 1.5 },
    rectRadius: 0.1,
  });
  slide.addText('−67%', {
    x: 7.2, y: 1.72, w: 2.4, h: 0.72,
    fontSize: 36, bold: true, color: C.tlight, align: 'center', fontFace: 'Calibri',
  });
  slide.addText('merchant time-to-\ncompliance reduction\n6 weeks → 11 days', {
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
  slide.addText('PSD3 Compliance Is Reactive — Merchants Discover Liability Shifts After the Deadline', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.9,
    fontSize: 24, bold: true, color: C.white, fontFace: 'Calibri', lineSpacingMultiple: 1.2,
  });

  const stats = [
    { icon: '⏰', num: '6 weeks', lbl: 'Average merchant time-to-compliance', sub: 'After a PSD3 change, merchants take 6 weeks to diagnose impact, plan, build, test, and deploy a fix' },
    { icon: '€', num: '€41M', lbl: 'EMEA GMV at regulatory risk today', sub: '213 of 847 Adyen EMEA merchants below 70% PSD3 readiness — €41M in processing volume exposed' },
    { icon: '🔴', num: '38%', lbl: 'MIT transactions non-compliant in Germany', sub: 'PSD3 MIT consent mandate effective July 10, 2026 — €6,800/day in chargeback exposure accumulating now' },
  ];
  stats.forEach((s, i) => {
    const x = 0.5 + i * 3.1;
    slide.addShape(prs.ShapeType.rect, {
      x, y: 1.72, w: 2.85, h: 3.4,
      fill: { color: C.hero }, line: { color: C.teal, pt: 1, transparency: 60 }, rectRadius: 0.1,
      shadow: makeShadow(),
    });
    slide.addText(s.icon, { x, y: 1.88, w: 2.85, h: 0.44, fontSize: 22, align: 'center' });
    slide.addText(s.num, {
      x, y: 2.36, w: 2.85, h: 0.72,
      fontSize: 30, bold: true, color: C.tlight, align: 'center', fontFace: 'Calibri',
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

  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 5.38, w: 9, h: 1.18,
    fill: { color: C.hero }, line: { color: C.orange, pt: 1.5, transparency: 40 }, rectRadius: 0.1,
  });
  slide.addText('Root cause: Adyen has no proactive regulatory intelligence layer. Merchants learn about PSD3/PSR3 changes through industry newsletters or — worst case — through a chargeback. RegSense closes the 6-week gap with a 60-day early warning system that diagnoses merchant-specific impact and prescribes resolution steps before liability shifts.', {
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
  slide.addText('Reactive Compliance Is a Revenue Risk. Proactive Intelligence Is a Retention Product.', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.82,
    fontSize: 22, bold: true, color: C.ink, fontFace: 'Calibri', lineSpacingMultiple: 1.2,
  });

  // VS divider
  slide.addShape(prs.ShapeType.ellipse, {
    x: 4.6, y: 2.3, w: 0.8, h: 0.8,
    fill: { color: C.dark }, line: { type: 'none' },
  });
  slide.addText('VS', {
    x: 4.6, y: 2.36, w: 0.8, h: 0.62,
    fontSize: 11, bold: true, color: C.white, align: 'center', fontFace: 'Calibri',
  });

  // LEFT — current
  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.6, w: 4.0, h: 4.2,
    fill: { color: 'FEF2F2' }, line: { color: 'EF4444', pt: 1.5, transparency: 40 }, rectRadius: 0.1,
  });
  slide.addText('❌  Reactive Compliance (Today)', {
    x: 0.68, y: 1.76, w: 3.68, h: 0.4,
    fontSize: 13, bold: true, color: C.ink, fontFace: 'Calibri',
  });
  [
    'PSD3 change published → merchant reads it in a newsletter 3 weeks later',
    'Internal legal + tech assessment takes 2 additional weeks',
    'Dev build, test, deploy adds another 2 weeks — total: 7 weeks to compliance',
    'Chargeback exposure accumulates every day of the gap',
    'Adyen has no visibility — merchants churn blaming "regulatory complexity"',
  ].forEach((pt, i) => {
    slide.addText('• ' + pt, {
      x: 0.68, y: 2.22 + i * 0.62, w: 3.6, h: 0.54,
      fontSize: 10.5, color: C.slate, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
    });
  });

  // RIGHT — proposed
  slide.addShape(prs.ShapeType.rect, {
    x: 5.5, y: 1.6, w: 4.0, h: 4.2,
    fill: { color: 'F0FDFA' }, line: { color: C.teal, pt: 1.5, transparency: 30 }, rectRadius: 0.1,
  });
  slide.addText('✅  RegSense (Proposed)', {
    x: 5.68, y: 1.76, w: 3.68, h: 0.4,
    fontSize: 13, bold: true, color: C.ink, fontFace: 'Calibri',
  });
  [
    'PSD3 change detected → RegSense calculates merchant-specific impact in 24 hours',
    'Merchant receives 60-day advance alert with countdown, exposure estimate, and 3-step fix',
    'Adyen SDK resolution path built in — no external dev work needed for common cases',
    'SCA Exemption Optimiser benchmarks merchant vs top-quartile Adyen peers',
    'Adyen Ops sees portfolio health — proactive outreach replaces reactive support tickets',
  ].forEach((pt, i) => {
    slide.addText('• ' + pt, {
      x: 5.68, y: 2.22 + i * 0.62, w: 3.6, h: 0.54,
      fontSize: 10.5, color: C.slate, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
    });
  });

  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.0, w: 9, h: 0.96,
    fill: { color: C.dark }, line: { type: 'none' }, rectRadius: 0.1,
  });
  slide.addText('Key insight: Compliance is the most under-served surface in the PSP product stack. Every competitor treats it as a support function. RegSense turns it into a proactive retention product — merchants who see compliance intelligence on their Adyen dashboard churn at 40% lower rates in year 2.', {
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
  slide.addText('How RegSense Works: Monitor → Diagnose → Alert → Resolve', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 24, bold: true, color: C.white, fontFace: 'Calibri',
  });

  const steps = [
    { n: '01', title: 'Regulatory Change Monitoring', desc: 'RegSense monitors PSD3/PSR3 publications, EBA guidelines, national regulator bulletins, and Adyen policy updates across all 14 EMEA markets. Detection latency: under 24 hours from official publication.' },
    { n: '02', title: 'Merchant-Specific Impact Diagnosis', desc: 'For each regulatory change, RegSense calculates: which of the 847 EMEA merchants are affected, by how much (GMV exposure in EUR), and which specific transaction flows need updating. Score updated daily.' },
    { n: '03', title: '60-Day Advance Alert with Resolution Path', desc: 'Affected merchants receive an in-dashboard alert 60 days before the compliance deadline. Alert includes: countdown timer, exact GMV at risk, and a numbered 3-step resolution path with Adyen SDK integration code.' },
    { n: '04', title: 'SCA Exemption Optimiser (Parallel)', desc: 'Independent of deadline alerts, RegSense benchmarks merchant SCA exemption utilisation vs Adyen top-quartile peers. Shows TRA, low-value, recurring, and whitelisting gaps with estimated approval rate uplift if closed.' },
    { n: '05', title: 'Adyen Ops Portfolio Health Dashboard', desc: 'Compliance ops team sees all 847 merchants ranked by readiness score. Proactive outreach queue prioritised by GMV at risk. Average time-to-compliance drops from 6 weeks to 11 days with proactive engagement.' },
  ];

  steps.forEach((s, i) => {
    const y = 1.42 + i * 0.98;
    slide.addShape(prs.ShapeType.ellipse, {
      x: 0.5, y, w: 0.48, h: 0.48,
      fill: { color: C.teal }, line: { type: 'none' },
    });
    slide.addText(s.n, {
      x: 0.5, y: y + 0.02, w: 0.48, h: 0.38,
      fontSize: 9, bold: true, color: C.white, align: 'center', fontFace: 'Calibri',
    });
    if (i < 4) {
      slide.addShape(prs.ShapeType.rect, {
        x: 0.72, y: y + 0.48, w: 0.04, h: 0.5,
        fill: { color: C.teal, transparency: 60 }, line: { type: 'none' },
      });
    }
    slide.addText(s.title, {
      x: 1.18, y, w: 8.2, h: 0.28,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri',
    });
    slide.addText(s.desc, {
      x: 1.18, y: y + 0.28, w: 8.2, h: 0.62,
      fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });

  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.52, w: 9, h: 0.64,
    fill: { color: C.hero }, line: { color: C.orange, pt: 1, transparency: 40 }, rectRadius: 0.08,
  });
  slide.addText('PhonePe Proof: Built zero-to-one B2B merchant onboarding platform with parallel KYC/KYB review across Technical, Compliance, and Legal sign-off — same parallel multi-jurisdiction review architecture as RegSense EMEA compliance diagnosis.', {
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
  slide.addText('4 Key Screen States — Compliance Dashboard, Exemptions, Alert, Ops View', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 21, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  const screens = [
    {
      num: '01', title: 'RegSense Dashboard',
      desc: 'Portfolio PSD3 readiness score (67/100). 6 EU markets colour-coded by compliance status. 3 markets require action chip. 47-day Germany alert pinned.',
      mock: '[ PSD3 Readiness: 67/100 ]\n NL 91% ✓  BE 83% ✓\n DE 72% ⚠  FR 58% ⚠\n ES 44% ✕  PL 39% ✕\n[⚠ 3 markets require action]',
    },
    {
      num: '02', title: 'SCA Exemption Optimiser',
      desc: 'TRA: 34% (benchmark 71%). Recurring MIT: 52% (benchmark 78%). +8.3% approval uplift if gaps closed. Amber benchmark lines vs green fill bars.',
      mock: '[ SCA Exemption Optimiser ]\n +8.3% approval uplift\n TRA:  34% ▓░░ 71% bmark\n LVE:  89% ▓▓▓ 85% bmark ✓\n MIT:  52% ▓▓░ 78% bmark\n WL:   12% ▓░░ 45% bmark',
    },
    {
      num: '03', title: 'Regulatory Alert',
      desc: '47-day countdown. €2.4M exposure. 3 numbered resolution steps. "Schedule Compliance Update" CTA triggers success state with confetti and RegSense score preview (91/100).',
      mock: '[ PSD3 Alert — Germany ]\n 47 days to deadline\n €2.4M exposure if missed\n 1. Update MIT consent flow\n 2. Backfill 124K subscribers\n 3. Enable TRA fallback\n[Schedule Compliance Update]',
    },
    {
      num: '04', title: 'Adyen Ops Portfolio',
      desc: 'All 847 EMEA merchants ranked by readiness score. At-risk list: Zalando 44, HelloFresh 51, Booking.com 63. Proactive outreach queue. €41M GMV risk KPI.',
      mock: '[ Ops Portfolio · 847 mchts ]\n 213 below 70%  €41M risk\n Zalando SE   DE  44 ●\n HelloFresh   DE  51 ●\n Booking.com  ES  63 ●\n Spotify AB   FR  69 ●',
    },
  ];

  screens.forEach((s, i) => {
    const x = 0.32 + i * 2.36;
    slide.addShape(prs.ShapeType.rect, {
      x, y: 1.4, w: 2.2, h: 5.3,
      fill: { color: C.white }, line: { color: C.teal, pt: 1.5, transparency: 40 }, rectRadius: 0.1,
      shadow: makeShadow(),
    });
    slide.addShape(prs.ShapeType.rect, {
      x, y: 1.4, w: 2.2, h: 0.36,
      fill: { color: C.dark }, line: { type: 'none' }, rectRadius: 0.1,
    });
    slide.addText(s.num + '  ' + s.title, {
      x: x + 0.1, y: 1.44, w: 2.05, h: 0.28,
      fontSize: 8.5, bold: true, color: C.tlight, fontFace: 'Calibri',
    });
    slide.addShape(prs.ShapeType.rect, {
      x: x + 0.1, y: 1.86, w: 2.02, h: 2.3,
      fill: { color: 'F0FDFA' }, line: { color: C.teal, pt: 0.75, transparency: 60 }, rectRadius: 0.08,
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

  slide.addText('Live interactive prototype: adyen-regsense-prototype.html', {
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
  slide.addText('Projected Impact — Compliance as a Retention Engine', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 24, bold: true, color: C.white, fontFace: 'Calibri',
  });

  slide.addText('MERCHANT IMPACT', {
    x: 0.5, y: 1.38, w: 4.4, h: 0.28,
    fontSize: 8.5, bold: true, color: C.tlight, charSpacing: 2, fontFace: 'Calibri',
  });
  const leftMetrics = [
    { val: '−67%', lbl: 'Time-to-compliance reduction', sub: '6 weeks → 11 days average with proactive RegSense alert and Adyen SDK resolution path' },
    { val: '€41M', lbl: 'GMV protected from chargeback exposure', sub: '213 below-threshold merchants brought to full compliance before liability shift' },
    { val: '+8.3%', lbl: 'Approval rate uplift from exemption optimisation', sub: 'TRA + MIT + whitelisting gaps closed — €3.1M additional approved GMV per month' },
    { val: '91/100', lbl: 'RegSense readiness score post-remediation', sub: 'Germany rises from 72% to 91% readiness after MIT consent update scheduled' },
  ];
  leftMetrics.forEach((m, i) => {
    const y = 1.72 + i * 1.22;
    slide.addShape(prs.ShapeType.rect, {
      x: 0.5, y, w: 4.3, h: 1.08,
      fill: { color: C.hero }, line: { color: C.teal, pt: 0.75, transparency: 50 }, rectRadius: 0.08,
    });
    slide.addText(m.val, {
      x: 0.68, y: y + 0.1, w: 1.6, h: 0.52,
      fontSize: 28, bold: true, color: C.tlight, fontFace: 'Calibri',
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

  slide.addText('ADYEN RETENTION IMPACT', {
    x: 5.1, y: 1.38, w: 4.4, h: 0.28,
    fontSize: 8.5, bold: true, color: C.orange, charSpacing: 2, fontFace: 'Calibri',
  });
  const rightMetrics = [
    { val: '−40%', lbl: 'Merchant churn in year 2', sub: 'Merchants with RegSense dashboard visibility renew at significantly higher rates vs control group' },
    { val: '+€22M', lbl: 'Annual processing fee protected', sub: 'Based on 40% churn reduction across 213 at-risk merchants at avg €104K annual Adyen fee each' },
    { val: '−72%', lbl: 'Compliance-related support tickets', sub: 'Proactive alerts replace reactive inbound from merchants discovering liability after the fact' },
    { val: '0 cost', lbl: 'Regulatory monitoring infrastructure', sub: 'RegSense ingests EBA and national regulator RSS feeds — operational cost near zero at scale' },
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
    fill: { color: C.hero }, line: { color: C.teal, pt: 1, transparency: 50 }, rectRadius: 0.08,
  });
  slide.addText('RegSense converts compliance from a cost centre into a retention lever. The €22M protected processing fee alone represents a 10× return on the estimated RegSense build investment — before counting the €41M GMV protection and approval rate uplift.', {
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
  slide.addText('I Built the Parallel Review Architecture. Twice.', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 23, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.38, w: 4.3, h: 5.3,
    fill: { color: C.dark }, line: { type: 'none' }, rectRadius: 0.1,
  });
  slide.addText('What I Built at PhonePe + ADCB', {
    x: 0.7, y: 1.56, w: 3.9, h: 0.34,
    fontSize: 12, bold: true, color: C.tlight, fontFace: 'Calibri',
  });
  [
    'Built zero-to-one multi-tenant B2B merchant onboarding platform with automated KYC/KYB document flows and parallel eligibility review across Technical, Compliance, and Legal — collapsed TAT from 2 days to 30 minutes (5,000+ merchants)',
    'ADCB CardLaunch: Corporate card onboarding with trade licence OCR auto-fill, AI credit scoring, and parallel multi-jurisdiction compliance review — compresses 3-week process to 24 minutes',
    'PhonePe: Owned cross-functional delivery across Technology, Compliance, Legal, Finance, Operations — same stakeholder landscape as RegSense PSD3 compliance product',
    'Built Antigravity Resume Agent — 6 parallel AI agents (scraping, synthesis, layout, ATS, outreach) — same parallel architecture as RegSense multi-market compliance diagnosis',
  ].forEach((pt, i) => {
    slide.addText('• ' + pt, {
      x: 0.7, y: 2.0 + i * 1.08, w: 3.9, h: 0.96,
      fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.35,
    });
  });

  slide.addShape(prs.ShapeType.rect, {
    x: 5.1, y: 1.38, w: 4.3, h: 5.3,
    fill: { color: C.white }, line: { color: C.teal, pt: 1.5, transparency: 40 }, rectRadius: 0.1,
    shadow: makeShadow(),
  });
  slide.addText('How It Maps to the Adyen JD', {
    x: 5.3, y: 1.56, w: 3.9, h: 0.34,
    fontSize: 12, bold: true, color: C.ink, fontFace: 'Calibri',
  });
  [
    ['Parallel Multi-Jurisdiction Review', 'PhonePe KYC/KYB parallel flows → RegSense parallel compliance diagnosis across DE/NL/FR/ES/BE/PL'],
    ['Regulatory Alignment in Product Delivery', 'Owned Compliance + Legal sign-off at PhonePe for all payment features → RegSense PSD3/PSR3 regulatory layer'],
    ['0→1 Compliance Product Design', 'Founding team at Kotak Cherry in regulated banking → RegSense from zero within Adyen regulated infra'],
    ['Cross-Functional Stakeholder Delivery', 'Tech / Compliance / Legal / Finance / Ops at PhonePe → exact same map for RegSense EMEA build'],
  ].forEach((pt, i) => {
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
  slide.addText('Phased Rollout — Alert First, Then Intelligence Layer', {
    x: 0.5, y: 0.62, w: 8.8, h: 0.56,
    fontSize: 24, bold: true, color: C.white, fontFace: 'Calibri',
  });

  const phases = [
    {
      phase: 'Phase 1',
      time: 'Month 1–2 · Pilot',
      title: 'Germany Deadline Alert (47 merchants)',
      pts: [
        'Deploy RegSense MIT consent alert to 47 DE-active merchants',
        'Instrument compliance resolution tracking — days to fix per merchant',
        'Measure time-to-compliance vs historical reactive baseline (6 weeks)',
        'Validate €2.4M chargeback exposure estimate against actuals',
      ],
      color: C.tlight,
    },
    {
      phase: 'Phase 2',
      time: 'Month 3–4 · Scale',
      title: 'SCA Exemption Optimiser + 6 Markets',
      pts: [
        'Launch SCA Exemption Optimiser with TRA + MIT + whitelisting benchmarks',
        'Expand RegSense monitoring to all 6 EMEA markets in scope',
        'Onboard Adyen Ops team with portfolio health dashboard',
        'Proactive outreach queue for merchants below 70% readiness',
      ],
      color: C.orange,
    },
    {
      phase: 'Phase 3',
      time: 'Month 5–6 · Full Rollout',
      title: 'All 14 EMEA Markets + AI Summaries',
      pts: [
        'Roll out RegSense to all 847 EMEA merchants as a default dashboard module',
        'Integrate PSR3 scope monitoring (Q4 2026 timeline)',
        'AI-generated plain-English regulatory change summaries per merchant type',
        'Publish Adyen EMEA Compliance Benchmark Report for merchant acquisition',
      ],
      color: 'A78BFA',
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

  slide.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.1, w: 9, h: 1.02,
    fill: { color: C.hero }, line: { color: C.teal, pt: 1.5, transparency: 30 }, rectRadius: 0.1,
  });
  slide.addText('What I need to build this:', {
    x: 0.72, y: 6.2, w: 3, h: 0.28,
    fontSize: 10, bold: true, color: C.tlight, fontFace: 'Calibri',
  });
  slide.addText('Access to Adyen regulatory monitoring feeds  ·  Alignment with Compliance + Legal team  ·  1 backend engineer for EBA/regulator RSS ingestion  ·  1 designer for merchant dashboard module  ·  Adyen Ops partnership for proactive outreach workflow', {
    x: 0.72, y: 6.5, w: 8.6, h: 0.54,
    fontSize: 9.5, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3,
  });

  slide.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade', {
    x: 0.5, y: 7.2, w: 9, h: 0.26,
    fontSize: 8.5, color: C.muted, fontFace: 'Calibri', align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/adyen-regsense-deck.pptx' })
  .then(() => console.log('adyen-regsense-deck.pptx generated'))
  .catch(e => console.error(e));
