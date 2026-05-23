const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:  '0A1E3D',
  hero:  '0F2A54',
  gold:  'C9A227',
  teal:  '0D9488',
  lteal: '2DD4BF',
  white: 'FFFFFF',
  lgray: 'F1F5F9',
  muted: '64748B',
  blue:  '2563EB',
  red:   'DC2626',
  green: '16A34A',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });

// ─────────────────────────────────────────
// SLIDE 1 — COVER
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  // Diagonal motif
  for (let i = 0; i < 6; i++) {
    sl.addShape(prs.ShapeType.rect, {
      x: 7.2 + i * 0.45, y: -0.2, w: 0.18, h: 9.0,
      fill: { color: C.gold, transparency: 88 - i * 10 },
      line: { type: 'none' }, rotate: 20,
    });
  }

  // Gold accent left strip
  sl.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: 0.18, h: 7.5,
    fill: { color: C.gold }, line: { type: 'none' },
  });

  // Tag
  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 0.38, w: 3.5, h: 0.28,
    fill: { color: C.gold, transparency: 85 },
    line: { color: C.gold, transparency: 65, pt: 0.5 },
    rectRadius: 0.14,
  });
  sl.addText('J.P. MORGAN · CEEMEA CORRESPONDENT BANKING', {
    x: 0.5, y: 0.38, w: 3.5, h: 0.28,
    fontSize: 7, bold: true, color: C.gold,
    align: 'center', valign: 'middle', fontFace: 'Calibri',
  });

  // Product name
  sl.addText('ClearPath', {
    x: 0.48, y: 0.82, w: 7.0, h: 1.35,
    fontSize: 68, bold: true, color: C.white,
    fontFace: 'Calibri', charSpacing: -2,
  });

  // Subtitle
  sl.addText('CEEMEA Correspondent KYC Acceleration', {
    x: 0.48, y: 2.0, w: 6.5, h: 0.45,
    fontSize: 20, bold: false, color: C.lteal,
    fontFace: 'Calibri',
  });

  // Description
  sl.addText('Transforming KYC renewal from a 6-month email chain into a 38-day self-serve digital process', {
    x: 0.48, y: 2.52, w: 6.5, h: 0.32,
    fontSize: 12, color: 'A0AABA', fontFace: 'Calibri',
  });

  // Presenter
  sl.addText('Presented by Ajay Avaghade · Growth & Monetization PM', {
    x: 0.48, y: 3.0, w: 5.5, h: 0.28,
    fontSize: 10, color: '8899AA', fontFace: 'Calibri', italic: true,
  });

  // Big stat bottom-right
  sl.addShape(prs.ShapeType.rect, {
    x: 7.4, y: 4.8, w: 2.3, h: 1.9,
    fill: { color: C.teal, transparency: 82 },
    line: { color: C.teal, transparency: 55, pt: 1 },
    rectRadius: 0.14,
  });
  sl.addText('−81%', {
    x: 7.4, y: 4.82, w: 2.3, h: 0.65,
    fontSize: 34, bold: true, color: C.lteal,
    fontFace: 'Calibri', align: 'center',
  });
  sl.addText('KYC renewal time\n(7 months → 38 days)', {
    x: 7.4, y: 5.44, w: 2.3, h: 0.6,
    fontSize: 10, color: 'A0AABA', fontFace: 'Calibri', align: 'center', valign: 'top',
  });

  sl.addShape(prs.ShapeType.rect, {
    x: 0, y: 7.2, w: 10, h: 0.02,
    fill: { color: C.gold, transparency: 70 }, line: { type: 'none' },
  });
}

// ─────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.hero };

  sl.addText('THE PROBLEM', {
    x: 0.5, y: 0.38, w: 9, h: 0.24,
    fontSize: 9, bold: true, color: C.gold,
    fontFace: 'Calibri', charSpacing: 1.5,
  });
  sl.addText('De-risking Is Destroying CEEMEA Correspondent\nBanking — Not Because of Risk, But Because of KYC', {
    x: 0.5, y: 0.66, w: 9, h: 1.1,
    fontSize: 24, bold: true, color: C.white,
    fontFace: 'Calibri', charSpacing: -0.5,
  });

  const cols3 = [
    { icon: '📉', stat: '30%+', lbl: 'CEEMEA correspondent relationships lost to de-risking', sub: 'Banks citing "KYC burden too high" — not actual compliance risk — as reason for terminating relationships over the past decade' },
    { icon: '📅', stat: '6–9 mo', lbl: 'Average KYC renewal time for one correspondent', sub: 'Sequential manual review: document chase by email → UK compliance → UAE → US OFAC → country-level sign-off. Each stage waits for the previous.' },
    { icon: '💵', stat: '$240M+', lbl: 'Annual payment volume at risk per at-risk corridor', sub: 'Every terminated correspondent relationship removes a payment corridor entirely. CEEMEA market requires 1,500+ active relationships for full coverage.' },
  ];

  cols3.forEach((c, i) => {
    const x = 0.5 + i * 3.1;
    sl.addShape(prs.ShapeType.rect, {
      x, y: 2.0, w: 2.85, h: 3.4,
      fill: { color: 'FFFFFF', transparency: 93 },
      line: { color: C.gold, transparency: 65, pt: 0.75 },
      rectRadius: 0.12,
    });
    sl.addText(c.icon, { x, y: 2.12, w: 2.85, h: 0.45, fontSize: 22, align: 'center', fontFace: 'Calibri' });
    sl.addText(c.stat, {
      x, y: 2.62, w: 2.85, h: 0.72,
      fontSize: 40, bold: true, color: C.gold,
      fontFace: 'Calibri', align: 'center',
    });
    sl.addText(c.lbl, {
      x: x + 0.12, y: 3.38, w: 2.6, h: 0.48,
      fontSize: 11, bold: true, color: C.white,
      fontFace: 'Calibri', align: 'center',
    });
    sl.addText(c.sub, {
      x: x + 0.12, y: 3.9, w: 2.6, h: 0.85,
      fontSize: 9.5, color: 'A0AABA', fontFace: 'Calibri',
      align: 'center', valign: 'top',
    });
  });

  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 5.6, w: 9, h: 1.42,
    fill: { color: C.teal, transparency: 88 },
    line: { color: C.teal, transparency: 60, pt: 0.75 },
    rectRadius: 0.12,
  });
  sl.addText('Root cause:  ', {
    x: 0.72, y: 5.72, w: 8.6, h: 0.28,
    fontSize: 11, bold: true, color: C.lteal, fontFace: 'Calibri',
  });
  sl.addText('The KYC process is not automated, not parallel, and not self-serve. Correspondent banks submit documents by email. Each jurisdiction reviews sequentially. J.P. Morgan relationship managers spend 40% of renewal time on document chasing rather than relationship management. ClearPath converts this from a compliance bottleneck into a structured digital workflow.', {
    x: 0.72, y: 6.02, w: 8.6, h: 0.85,
    fontSize: 10.5, color: 'C8D8E8', fontFace: 'Calibri', valign: 'top',
  });
}

// ─────────────────────────────────────────
// SLIDE 3 — THE INSIGHT
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE INSIGHT', {
    x: 0.5, y: 0.38, w: 9, h: 0.24,
    fontSize: 9, bold: true, color: C.teal,
    fontFace: 'Calibri', charSpacing: 1.5,
  });
  sl.addText('KYC Renewal is a Revenue Protection Product,\nNot a Compliance Cost Centre', {
    x: 0.5, y: 0.66, w: 9, h: 1.1,
    fontSize: 26, bold: true, color: C.dark,
    fontFace: 'Calibri', charSpacing: -0.5,
  });

  // Left: current
  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.9, w: 4.1, h: 4.7,
    fill: { color: 'FEF2F2' }, line: { color: 'FCA5A5', pt: 0.75 },
    rectRadius: 0.12, shadow: makeShadow(),
  });
  sl.addText('❌  Current Process', {
    x: 0.7, y: 2.04, w: 3.7, h: 0.35,
    fontSize: 13, bold: true, color: 'DC2626', fontFace: 'Calibri',
  });
  [
    'Documents requested via email — 3–4 rounds per renewal',
    'Sequential compliance review: UK → UAE → US OFAC → local',
    'No shared tracker — both sides fly blind on status',
    '6–9 months average: relationship at risk the entire time',
    'RM spends 40% of time as document chaser, not banker',
    'At-risk window: payment volume paused during limbo',
  ].forEach((pt, i) => {
    sl.addText(`• ${pt}`, {
      x: 0.7, y: 2.5 + i * 0.55, w: 3.7, h: 0.48,
      fontSize: 10.5, color: '1A1A2E', fontFace: 'Calibri', valign: 'top',
    });
  });

  // VS divider
  sl.addShape(prs.ShapeType.ellipse, {
    x: 4.62, y: 3.85, w: 0.76, h: 0.76,
    fill: { color: C.dark }, line: { type: 'none' },
  });
  sl.addText('VS', {
    x: 4.62, y: 3.85, w: 0.76, h: 0.76,
    fontSize: 10, bold: true, color: C.white,
    fontFace: 'Calibri', align: 'center', valign: 'middle',
  });

  // Right: ClearPath
  sl.addShape(prs.ShapeType.rect, {
    x: 5.4, y: 1.9, w: 4.1, h: 4.7,
    fill: { color: 'F0FDFA' }, line: { color: C.teal, pt: 0.75 },
    rectRadius: 0.12, shadow: makeShadow(),
  });
  sl.addText('✅  ClearPath', {
    x: 5.6, y: 2.04, w: 3.7, h: 0.35,
    fontSize: 13, bold: true, color: C.teal, fontFace: 'Calibri',
  });
  [
    'Self-serve portal: correspondent bank submits documents directly',
    'OCR auto-fill from prior submissions + public registry APIs',
    'Parallel review across all jurisdictions simultaneously',
    '38 days average — relationship risk window cut by 81%',
    'Shared live tracker: both sides see status in real time',
    'RM focus shifts from document chasing to relationship growth',
  ].forEach((pt, i) => {
    sl.addText(`• ${pt}`, {
      x: 5.6, y: 2.5 + i * 0.55, w: 3.7, h: 0.48,
      fontSize: 10.5, color: '1A1A2E', fontFace: 'Calibri', valign: 'top',
    });
  });

  // Callout
  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.78, w: 9, h: 0.52,
    fill: { color: C.gold, transparency: 88 },
    line: { color: C.gold, transparency: 60, pt: 0.75 },
    rectRadius: 0.1,
  });
  sl.addText('Key insight:  Every prevented correspondent relationship termination protects payment volume. ClearPath converts KYC from pure cost into a revenue-protection product with a direct P&L line.', {
    x: 0.7, y: 6.8, w: 8.6, h: 0.46,
    fontSize: 10, color: C.dark, fontFace: 'Calibri', valign: 'middle',
  });
}

// ─────────────────────────────────────────
// SLIDE 4 — THE MECHANIC
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE MECHANIC', {
    x: 0.5, y: 0.38, w: 9, h: 0.24,
    fontSize: 9, bold: true, color: C.gold,
    fontFace: 'Calibri', charSpacing: 1.5,
  });
  sl.addText('How ClearPath Works: Self-Serve + Parallel Review', {
    x: 0.5, y: 0.66, w: 9, h: 0.7,
    fontSize: 26, bold: true, color: C.white,
    fontFace: 'Calibri', charSpacing: -0.5,
  });

  const steps = [
    { n: '01', title: 'Invite Sent', body: 'J.P. Morgan sends ClearPath invite 90 days before KYC expiry. Correspondent bank accesses secure self-serve portal via link + MFA.' },
    { n: '02', title: 'OCR Auto-fill', body: 'Documents uploaded. OCR extracts entity data, pre-populates form from prior submission + public registry APIs. 8 fields auto-filled on average.', hl: true },
    { n: '03', title: 'Submit & Parallel', body: 'One submission triggers parallel compliance review across all jurisdictions: UK FCA, UAE CBUAE, US OFAC, Egypt CBE — simultaneously, not sequentially.' },
    { n: '04', title: 'Live Tracker', body: 'Both the correspondent bank and J.P. Morgan RM see real-time status across all 4 compliance streams. No email chase. Exceptions surfaced with self-serve guidance.' },
    { n: '05', title: 'Activate', body: 'All jurisdictions cleared → relationship auto-renewed. Payment corridors remain live throughout. Volume protected. RM notified with next renewal date.' },
  ];

  steps.forEach((st, i) => {
    const x = 0.3 + i * 1.9;
    const highlighted = st.hl;
    sl.addShape(prs.ShapeType.rect, {
      x, y: 1.55, w: 1.7, h: 4.7,
      fill: { color: highlighted ? C.teal : 'FFFFFF', transparency: highlighted ? 0 : 93 },
      line: { color: highlighted ? C.teal : C.teal, transparency: highlighted ? 0 : 72, pt: 0.5 },
      rectRadius: 0.12,
    });
    sl.addText(st.n, {
      x, y: 1.7, w: 1.7, h: 0.42,
      fontSize: 24, bold: true,
      color: highlighted ? C.dark : C.lteal,
      fontFace: 'Calibri', align: 'center',
    });
    sl.addText(st.title, {
      x: x + 0.1, y: 2.18, w: 1.5, h: 0.55,
      fontSize: 11, bold: true,
      color: highlighted ? C.dark : C.white,
      fontFace: 'Calibri', align: 'center',
    });
    sl.addText(st.body, {
      x: x + 0.1, y: 2.78, w: 1.5, h: 2.5,
      fontSize: 9.5,
      color: highlighted ? '064E3B' : 'A0AABA',
      fontFace: 'Calibri', valign: 'top',
    });
    if (i < steps.length - 1) {
      sl.addText('→', {
        x: x + 1.72, y: 3.8, w: 0.18, h: 0.3,
        fontSize: 11, color: C.gold, fontFace: 'Calibri', align: 'center',
      });
    }
  });

  // Proof note
  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.46, w: 9, h: 0.78,
    fill: { color: C.gold, transparency: 88 },
    line: { color: C.gold, transparency: 60, pt: 0.5 },
    rectRadius: 0.1,
  });
  sl.addText('PhonePe proof:  Built zero-to-one B2B merchant self-serve onboarding platform at PhonePe — collapsed 2-day manual onboarding to 30 minutes via automated KYC/KYB document flows and parallel eligibility checks. The ADCB CardLaunch case study (3 weeks → 24 minutes) is the same architecture in a banking context. ClearPath applies the same parallel workflow to correspondent KYC.', {
    x: 0.7, y: 6.5, w: 8.6, h: 0.72,
    fontSize: 9.5, color: C.gold, fontFace: 'Calibri', valign: 'middle',
  });
}

// ─────────────────────────────────────────
// SLIDE 5 — THE PRODUCT
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE PRODUCT', {
    x: 0.5, y: 0.38, w: 9, h: 0.24,
    fontSize: 9, bold: true, color: C.teal,
    fontFace: 'Calibri', charSpacing: 1.5,
  });
  sl.addText('4 Key Screens — Built & Live as a Working Prototype', {
    x: 0.5, y: 0.66, w: 9, h: 0.7,
    fontSize: 26, bold: true, color: C.dark,
    fontFace: 'Calibri', charSpacing: -0.5,
  });

  const screens = [
    {
      n: '01', title: 'Correspondent Portal', color: C.hero,
      desc: 'Correspondent bank landing page: KYC renewal countdown, document submission status (uploaded / in review / required), progress bar across 4 phases, one-tap upload CTA.',
      ascii: 'ClearPath · KYC Portal \n─────────────────────── \nArab African Int Bank   \n🇪🇬 Cairo · ARAIEGCX  \n⚠️ KYC Renewal: 47d   \n─────────────────────── \n✓ Central Bank License \n✓ Audited Financials   \n⏳ AML Policy (review) \n! UBO Declaration req. \n─────────────────────── \n[ Upload UBO Decl. →  ]',
    },
    {
      n: '02', title: 'OCR Auto-fill Form', color: C.teal,
      desc: 'Document upload triggers OCR extraction. 8 fields pre-populated from prior submission + registry data. Missing fields highlighted in red. Parallel compliance status shown below.',
      ascii: 'UBO Declaration        \n[ ✓ ][ 2 ][ 3 ][ 4 ]   \n─────────────────────── \n✦ 8 fields pre-filled  \nEntity: AAIB S.A.E. ✦  \nReg: EG-CAI-1964  ✦   \nJurisdiction: Egypt ✦  \nUBO Name: [Required]   \nUBO %: [Required]      \n─────────────────────── \n🇬🇧 UK FCA    ✓ Done \n🇦🇪 UAE CBUAE ✓ Done \n🇪🇬 Egypt CBE ⏳ Live \n🇺🇸 US OFAC   ⌛ Queue',
    },
    {
      n: '03', title: 'Relationship Activated', color: C.green,
      desc: 'All jurisdictions cleared: confetti celebration, full clearance certificate, before/after time comparison (7 months → 38 days), relationship validity date, volume protection confirmed.',
      ascii: '✓ Relationship Activated\n─────────────────────── \nArab African Int Bank  \n🇪🇬 · $840M/yr flow   \nKYC Status: Cleared ✓ \nValid until: May 2028  \n─────────────────────── \nTime: 7 months → 38d  \n        ↓ 81% faster   \n4 jurisdictions: 4/4 ✓\n─────────────────────── \n[Download Cert][Dashboard]',
    },
    {
      n: '04', title: 'J.P. Morgan Ops View', color: C.dark,
      desc: 'Compliance team view: pipeline of 1,500+ correspondent renewals. Clearance time KPI, churn reduction, at-risk alerts, health score by corridor, de-risking early warning system.',
      ascii: 'ClearPath · Compliance \n─────────────────────── \n38d avg clearance time \n−40% relationship churn\n7 at-risk (action req) \n─────────────────────── \nArab African  100% ✓  \nFirst Natl NG  72% ⏳  \nBanque Misr    31% ⚠️  \n─────────────────────── \nHealth: Egypt 87 · NG 61\nDe-risk alert: PAK $240M',
    },
  ];

  screens.forEach((sc, i) => {
    const x = 0.3 + i * 2.42;
    sl.addShape(prs.ShapeType.rect, {
      x, y: 1.55, w: 2.22, h: 5.6,
      fill: { color: C.white }, line: { color: sc.color, pt: 1 },
      rectRadius: 0.12, shadow: makeShadow(),
    });
    sl.addShape(prs.ShapeType.rect, {
      x, y: 1.55, w: 2.22, h: 0.45,
      fill: { color: sc.color }, line: { type: 'none' }, rectRadius: 0.12,
    });
    sl.addText(`${sc.n} — ${sc.title}`, {
      x: x + 0.08, y: 1.55, w: 2.06, h: 0.45,
      fontSize: 9, bold: true, color: C.white,
      fontFace: 'Calibri', align: 'center', valign: 'middle',
    });
    sl.addText(sc.ascii, {
      x: x + 0.1, y: 2.08, w: 2.02, h: 2.55,
      fontSize: 7, color: '1E293B', fontFace: 'Courier New', valign: 'top',
    });
    sl.addText(sc.desc, {
      x: x + 0.1, y: 4.72, w: 2.02, h: 2.1,
      fontSize: 9.5, color: C.muted, fontFace: 'Calibri', valign: 'top',
    });
  });

  sl.addText('Prototype live at: ajay-avaghade.github.io/Portfolio/jpmorgan-clearpath-prototype.html', {
    x: 0.5, y: 7.2, w: 9, h: 0.24,
    fontSize: 9, color: C.teal, fontFace: 'Calibri', italic: true, align: 'center',
  });
}

// ─────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.hero };

  sl.addText('IMPACT & ROI', {
    x: 0.5, y: 0.38, w: 9, h: 0.24,
    fontSize: 9, bold: true, color: C.gold,
    fontFace: 'Calibri', charSpacing: 1.5,
  });
  sl.addText('Projected Impact — Revenue Protection at Scale', {
    x: 0.5, y: 0.66, w: 9, h: 0.7,
    fontSize: 26, bold: true, color: C.white,
    fontFace: 'Calibri', charSpacing: -0.5,
  });

  sl.addText('CORRESPONDENT BANK EXPERIENCE', {
    x: 0.5, y: 1.55, w: 4.4, h: 0.28,
    fontSize: 9, bold: true, color: C.lteal,
    fontFace: 'Calibri', charSpacing: 1.2,
  });
  const cbMetrics = [
    { val: '38 days', lbl: 'Average time to KYC clearance (vs 6–9 months)', color: C.lteal },
    { val: '−81%', lbl: 'Time in "at-risk" limbo window during renewal cycle', color: C.gold },
    { val: '8 fields', lbl: 'Pre-filled automatically via OCR + registry APIs per submission', color: C.lteal },
    { val: '4 / 4', lbl: 'Jurisdiction reviews completed in parallel (not sequential)', color: C.gold },
  ];
  cbMetrics.forEach((m, i) => {
    const row = i % 2; const col = Math.floor(i / 2);
    const x = 0.5 + col * 2.3; const y = 1.92 + row * 1.55;
    sl.addShape(prs.ShapeType.rect, {
      x, y, w: 2.1, h: 1.38,
      fill: { color: 'FFFFFF', transparency: 93 },
      line: { color: C.teal, transparency: 65, pt: 0.5 },
      rectRadius: 0.1,
    });
    sl.addText(m.val, {
      x: x + 0.1, y: y + 0.12, w: 1.9, h: 0.5,
      fontSize: 24, bold: true, color: m.color, fontFace: 'Calibri',
    });
    sl.addText(m.lbl, {
      x: x + 0.1, y: y + 0.64, w: 1.9, h: 0.62,
      fontSize: 9, color: 'A0AABA', fontFace: 'Calibri', valign: 'top',
    });
  });

  sl.addShape(prs.ShapeType.rect, {
    x: 5.05, y: 1.55, w: 0.02, h: 5.0,
    fill: { color: C.gold, transparency: 70 }, line: { type: 'none' },
  });

  sl.addText('J.P. MORGAN REVENUE PROTECTION', {
    x: 5.3, y: 1.55, w: 4.2, h: 0.28,
    fontSize: 9, bold: true, color: C.gold,
    fontFace: 'Calibri', charSpacing: 1.2,
  });
  const jpmMetrics = [
    { val: '−40%', lbl: 'Correspondent relationship churn in CEEMEA corridors', color: C.lteal },
    { val: '+$30M', lbl: 'Annual payment volume protected per 100 relationships saved', color: C.gold },
    { val: '+40%', lbl: 'RM capacity freed from document chasing for relationship growth', color: C.lteal },
    { val: '$240M', lbl: 'At-risk Pakistan corridor volume flagged 30 days early for action', color: C.gold },
  ];
  jpmMetrics.forEach((m, i) => {
    const row = i % 2; const col = Math.floor(i / 2);
    const x = 5.3 + col * 2.2; const y = 1.92 + row * 1.55;
    sl.addShape(prs.ShapeType.rect, {
      x, y, w: 2.0, h: 1.38,
      fill: { color: 'FFFFFF', transparency: 93 },
      line: { color: C.gold, transparency: 60, pt: 0.5 },
      rectRadius: 0.1,
    });
    sl.addText(m.val, {
      x: x + 0.1, y: y + 0.12, w: 1.8, h: 0.5,
      fontSize: 22, bold: true, color: m.color, fontFace: 'Calibri',
    });
    sl.addText(m.lbl, {
      x: x + 0.1, y: y + 0.64, w: 1.8, h: 0.62,
      fontSize: 9, color: 'A0AABA', fontFace: 'Calibri', valign: 'top',
    });
  });

  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 5.2, w: 9, h: 1.0,
    fill: { color: C.teal, transparency: 88 },
    line: { color: C.teal, transparency: 60, pt: 0.75 },
    rectRadius: 0.12,
  });
  sl.addText('The ClearPath investment thesis:  De-risking is destroying CEEMEA payment volume not because of risk — because the KYC process is too painful. ClearPath converts the compliance renewal cycle from a liability into a retention tool: every relationship saved is incremental payment volume at near-zero marginal cost.', {
    x: 0.7, y: 5.25, w: 8.6, h: 0.9,
    fontSize: 10.5, color: C.lteal, fontFace: 'Calibri', valign: 'middle',
  });
}

// ─────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('PROOF OF WORK', {
    x: 0.5, y: 0.38, w: 9, h: 0.24,
    fontSize: 9, bold: true, color: C.teal,
    fontFace: 'Calibri', charSpacing: 1.5,
  });
  sl.addText('I Built This Architecture at PhonePe and ADCB.\nHere is the Evidence.', {
    x: 0.5, y: 0.66, w: 9, h: 1.1,
    fontSize: 24, bold: true, color: C.dark,
    fontFace: 'Calibri', charSpacing: -0.5,
  });

  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.75, w: 4.3, h: 5.1,
    fill: { color: C.dark }, line: { type: 'none' }, rectRadius: 0.14,
  });
  sl.addText('What I built at PhonePe + ADCB', {
    x: 0.7, y: 1.9, w: 3.9, h: 0.34,
    fontSize: 12, bold: true, color: C.gold, fontFace: 'Calibri',
  });
  [
    'Zero-to-one B2B merchant self-serve PG onboarding: 2-day manual TAT → 30 minutes via automated KYC/KYB document flows + parallel eligibility checks across credit, compliance, and technical teams — 5,000+ merchants acquired',
    'ADCB CardLaunch (case study): corporate card programme onboarding digitisation from 3-week RM-led process to 24 minutes via OCR auto-fill, AI credit scoring, and parallel document review — same architecture as ClearPath',
    'B2B marketplace compliance engine: built eligibility and onboarding controls across Tech, Legal, Compliance, and Finance — 500+ brand partners onboarded with full audit trail across all jurisdictions',
    'Ops dashboard for compliance teams: real-time queue management, exception tracking, partner health scoring — maps directly to ClearPath Slide 5 Screen 4 ops view',
  ].forEach((pt, i) => {
    sl.addText(`${i + 1}.  ${pt}`, {
      x: 0.7, y: 2.34 + i * 1.06, w: 3.9, h: 0.9,
      fontSize: 9.5, color: 'A0AABA', fontFace: 'Calibri', valign: 'top',
    });
  });

  sl.addShape(prs.ShapeType.rect, {
    x: 5.1, y: 1.75, w: 4.4, h: 5.1,
    fill: { color: 'F0FDFA' }, line: { color: C.teal, pt: 0.75 }, rectRadius: 0.14,
    shadow: makeShadow(),
  });
  sl.addText('How it maps to this JD', {
    x: 5.3, y: 1.9, w: 4.0, h: 0.34,
    fontSize: 12, bold: true, color: C.teal, fontFace: 'Calibri',
  });
  [
    '"Design and implement solutions for improvement of internal processes" → PhonePe B2B onboarding digitisation (2d → 30min) + ADCB CardLaunch (3wk → 24min): same architectural pattern applied here',
    '"Anticipate internal/external factors impacting Legal Entities" → ClearPath health score flags de-risking risk 30–90 days before corridor termination — exactly the early warning the JD requires',
    '"Understand risks inherent in the product and drive programs to reduce" → ClearPath is inherently a risk product: it quantifies relationship risk, tracks compliance coverage, and surfaces exceptions before they become terminations',
    '"Collaborate with Sales, Service, Operations and Technology" → Built PhonePe rewards marketplace with full cross-functional alignment; ADCB CardLaunch required Legal + Compliance + Finance + Tech sign-off across 3 markets',
  ].forEach((pt, i) => {
    sl.addText(`• ${pt}`, {
      x: 5.3, y: 2.34 + i * 1.06, w: 3.9, h: 0.9,
      fontSize: 9.5, color: '1A1A2E', fontFace: 'Calibri', valign: 'top',
    });
  });

  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 7.0, w: 9, h: 0.28,
    fill: { color: C.teal, transparency: 90 },
    line: { color: C.teal, transparency: 60, pt: 0.5 },
    rectRadius: 0.08,
  });
  sl.addText('"ClearPath is the ADCB CardLaunch architecture applied to correspondent banking KYC. I have shipped both the onboarding digitisation and the parallel compliance workflow. The domain is different; the product mechanics are the same."', {
    x: 0.7, y: 7.0, w: 8.6, h: 0.28,
    fontSize: 9, italic: true, color: C.dark, fontFace: 'Calibri', valign: 'middle',
  });
}

// ─────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('ROLLOUT PLAN', {
    x: 0.5, y: 0.38, w: 9, h: 0.24,
    fontSize: 9, bold: true, color: C.gold,
    fontFace: 'Calibri', charSpacing: 1.5,
  });
  sl.addText('Phased Rollout — UAE + Egypt Pilot to CEEMEA Scale', {
    x: 0.5, y: 0.66, w: 9, h: 0.7,
    fontSize: 26, bold: true, color: C.white,
    fontFace: 'Calibri', charSpacing: -0.5,
  });

  const phases = [
    {
      ph: 'Phase 1', period: 'Months 1–2', title: 'UAE + Egypt Pilot',
      body: 'Deploy ClearPath for 50 correspondent renewals across UAE and Egypt corridors — highest renewal volume, known KYC document requirements. Measure time-to-clearance vs email baseline. Validate OCR accuracy on Arabic and English documents.',
      color: C.teal,
    },
    {
      ph: 'Phase 2', period: 'Months 3–5', title: 'CEEMEA Parallel Review Engine',
      body: 'Expand parallel compliance review to all 4 primary jurisdictions (UK, UAE, US, local). Add West Africa (Nigeria, Ghana), East Africa (Kenya, Tanzania) corridors. Build RM dashboard with health score and at-risk alerting.',
      color: C.gold,
    },
    {
      ph: 'Phase 3', period: 'Months 6–12', title: 'Full CEEMEA + De-risking Firewall',
      body: 'All 1,500+ CEEMEA correspondent renewals managed via ClearPath. Health score model detects de-risking risk 90 days early. Integrate with J.P. Morgan KYC utility (KYC.com). Expand to new relationship onboarding (not just renewals).',
      color: C.blue,
    },
  ];

  phases.forEach((ph, i) => {
    const x = 0.5 + i * 3.18;
    sl.addShape(prs.ShapeType.rect, {
      x, y: 1.55, w: 2.95, h: 4.3,
      fill: { color: 'FFFFFF', transparency: 93 },
      line: { color: ph.color, transparency: 55, pt: 1 },
      rectRadius: 0.14,
    });
    sl.addShape(prs.ShapeType.rect, {
      x, y: 1.55, w: 2.95, h: 0.42,
      fill: { color: ph.color }, line: { type: 'none' }, rectRadius: 0.14,
    });
    sl.addText(`${ph.ph}  ·  ${ph.period}`, {
      x: x + 0.1, y: 1.55, w: 2.75, h: 0.42,
      fontSize: 9.5, bold: true,
      color: ph.color === C.gold ? C.dark : C.white,
      fontFace: 'Calibri', align: 'center', valign: 'middle',
    });
    sl.addText(ph.title, {
      x: x + 0.1, y: 2.08, w: 2.75, h: 0.44,
      fontSize: 14, bold: true, color: C.white, fontFace: 'Calibri',
    });
    sl.addText(ph.body, {
      x: x + 0.1, y: 2.6, w: 2.75, h: 2.7,
      fontSize: 10, color: 'A0AABA', fontFace: 'Calibri', valign: 'top',
    });
  });

  sl.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 6.02, w: 9, h: 0.88,
    fill: { color: C.teal, transparency: 88 },
    line: { color: C.teal, transparency: 55, pt: 0.75 },
    rectRadius: 0.12,
  });
  sl.addText('What I need to build this', {
    x: 0.7, y: 6.08, w: 8.6, h: 0.26,
    fontSize: 10, bold: true, color: C.gold, fontFace: 'Calibri',
  });
  sl.addText('Access to KYC renewal pipeline data  ·  Alignment with Compliance and Legal teams  ·  1 engineer for portal + OCR  ·  50 pilot correspondent bank nominations  ·  J.P. Morgan KYC utility (KYC.com) API access', {
    x: 0.7, y: 6.36, w: 8.6, h: 0.44,
    fontSize: 9.5, color: 'C8D8E8', fontFace: 'Calibri', valign: 'top',
  });

  sl.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  ajay-avaghade.github.io/Portfolio', {
    x: 0.5, y: 7.12, w: 9, h: 0.24,
    fontSize: 8.5, color: '6B7E9A', fontFace: 'Calibri', align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/jpmorgan-clearpath-deck.pptx' }).then(() => {
  console.log('jpmorgan-clearpath-deck.pptx done');
}).catch(e => console.error(e));
