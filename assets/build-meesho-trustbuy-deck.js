const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '353543',  // Meesho Dark Slate
  hero:   '1E1A26',  // Deep Violet/Dark
  pink:   'F43397',  // Meesho Brand Pink
  accent: 'E7E0FC',  // Lavender Accent
  white:  'FFFFFF',
  lgray:  'F8F8F9',  // Light Gray Surface
  green:  '038543',  // Meesho Green
  orange: 'FF7043',  // Alert Orange
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 3, offset: 1, angle: 45, opacity: 0.12 });

// ─────────────────────────────────────────
// SLIDE 1 — COVER
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  for (let i = 0; i < 8; i++) {
    sl.addShape(prs.ShapeType.line, { x: -0.5 + i * 1.5, y: 0, w: 4, h: 7.5, line: { color: C.pink, width: 0.4, transparency: 88 } });
  }

  sl.addText('MEESHO SUPPLIER QUALITY & RECOMMENDATION MECHANICS', { x: 0.5, y: 0.32, w: 9, h: 0.28, fontSize: 9, color: C.pink, charSpacing: 2 });
  sl.addText('TrustBuy', { x: 0.5, y: 0.75, w: 9, h: 1.4, fontSize: 64, color: C.white, bold: true });
  sl.addText('Supplier Quality Incentives Integrated Directly into the PRISM AI Feed', {
    x: 0.5, y: 2.0, w: 7.5, h: 0.6, fontSize: 22, color: C.accent, bold: false,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.5, y: 2.75, w: 2.8, h: 0.04, fill: { color: C.pink } });

  sl.addText('Presented by Ajay Avaghade · Product Manager 2 Candidate', {
    x: 0.5, y: 2.95, w: 7, h: 0.35, fontSize: 13, color: 'FFD5E5',
  });

  // Big stat
  sl.addShape(prs.ShapeType.rect, { x: 7.2, y: 5.2, w: 2.5, h: 1.9, fill: { color: C.pink }, shadow: makeShadow() });
  sl.addText('4.5★', { x: 7.2, y: 5.3, w: 2.5, h: 0.85, fontSize: 52, color: C.white, bold: true, align: 'center' });
  sl.addText('minimum rating required\nto get organic search boost', { x: 7.2, y: 6.1, w: 2.5, h: 0.8, fontSize: 11, color: 'FFE6F0', align: 'center' });
}

// ─────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE PROBLEM', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.pink, bold: true, charSpacing: 2 });
  sl.addText('Inconsistent Supplier Quality Causes High Return Rates and Drains Customer Lifetime Value.', {
    x: 0.5, y: 0.62, w: 9, h: 0.9, fontSize: 21, color: C.white, bold: true,
  });

  const cols = [
    { icon: '📉', stat: '18%', lbl: 'Impulsive product returns', sub: 'Fabric running thin or stitching defects drive immediate consumer returns. Wastes logistics margin.' },
    { icon: '⚠️', stat: '30%', lbl: 'Supplier metadata errors', sub: 'Incorrect color labeling and outdated measurement dimensions create expectation mismatches.' },
    { icon: '🔄', stat: '2.5×', lbl: 'CAC leak from churn', sub: 'Low catalog trust prevents repeat purchases. Users shop once and leave, inflating CAC.' },
  ];
  cols.forEach((c, i) => {
    const x = 0.45 + i * 3.12;
    sl.addShape(prs.ShapeType.rect, { x, y: 1.72, w: 2.9, h: 3.1, fill: { color: C.hero }, shadow: makeShadow() });
    sl.addText(c.icon, { x, y: 1.88, w: 2.9, h: 0.55, fontSize: 26, align: 'center' });
    sl.addText(c.stat, { x, y: 2.52, w: 2.9, h: 0.72, fontSize: 40, color: C.pink, bold: true, align: 'center' });
    sl.addText(c.lbl, { x, y: 3.22, w: 2.9, h: 0.38, fontSize: 12.5, color: C.white, bold: true, align: 'center' });
    sl.addText(c.sub, { x: x + 0.1, y: 3.62, w: 2.7, h: 0.9, fontSize: 10, color: 'FFD5E5', align: 'center' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.0, w: 9.1, h: 0.82, fill: { color: C.pink } });
  sl.addText('Root cause: Meesho\'s recommendation engine (PRISM) prioritizes conversion velocity over delivery success. Suppliers face no organic penalties for high returns. We must align search visibility with delivery and quality outcomes.', {
    x: 0.62, y: 5.06, w: 8.76, h: 0.7, fontSize: 10.5, color: C.white,
  });
}

// ─────────────────────────────────────────
// SLIDE 3 — THE INSIGHT
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE INSIGHT', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.pink, bold: true, charSpacing: 2 });
  sl.addText('Incentivize Quality: Link Supplier Search Visibility (PRISM) Directly to Return Rates.', {
    x: 0.5, y: 0.62, w: 9, h: 0.9, fontSize: 22, color: C.dark, bold: true,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 1.72, w: 4.1, h: 3.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('❌  Legacy Recommendation Feed', { x: 0.62, y: 1.85, w: 3.76, h: 0.38, fontSize: 12, color: '444444', bold: true });
  const left = [
    'Feeds ranked purely on conversion rate and price',
    'High-return products get same visibility as good ones',
    'No consumer trust badges for verified supplier quality',
    'Suppliers receive monthly return logs without action points',
    'Category managers must manually blacklist bad suppliers',
  ];
  left.forEach((t, i) => sl.addText(`• ${t}`, { x: 0.72, y: 2.35 + i * 0.42, w: 3.56, h: 0.36, fontSize: 11, color: '555555' }));

  sl.addShape(prs.ShapeType.ellipse, { x: 4.62, y: 2.9, w: 0.72, h: 0.72, fill: { color: C.pink } });
  sl.addText('VS', { x: 4.62, y: 2.9, w: 0.72, h: 0.72, fontSize: 11, color: C.white, bold: true, align: 'center', valign: 'middle' });

  sl.addShape(prs.ShapeType.rect, { x: 5.42, y: 1.72, w: 4.1, h: 3.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('✅  TrustBuy Loop (Proposed)', { x: 5.58, y: 1.85, w: 3.76, h: 0.38, fontSize: 12, color: C.pink, bold: true });
  const right = [
    'Feeds rank products based on successful delivery rate',
    'PRISM algorithm applies search boosts to verified sellers',
    'TrustBuy Badges surface supplier rating on PDP',
    'Action Center sends automated catalog warnings to sellers',
    'Admin console automatically alerts high-return sellers',
  ];
  right.forEach((t, i) => sl.addText(`• ${t}`, { x: 5.58, y: 2.35 + i * 0.42, w: 3.72, h: 0.36, fontSize: 10.5, color: '333333' }));

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.08, w: 9.1, h: 0.72, fill: { color: C.pink } });
  sl.addText('Key insight: Aligning incentives is better than policing. Suppliers will willingly correct descriptions, size charts, and fabric parameters when their daily search visibility is on the line.', {
    x: 0.62, y: 5.12, w: 8.76, h: 0.62, fontSize: 10, color: C.white,
  });
}

// ─────────────────────────────────────────
// SLIDE 4 — THE MECHANIC
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE MECHANIC', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.pink, bold: true, charSpacing: 2 });
  sl.addText('TrustBuy System: Badge, Score, Alert, and Re-rank', {
    x: 0.5, y: 0.62, w: 9, h: 0.65, fontSize: 22, color: C.white, bold: true,
  });

  const steps = [
    { n: '01', title: 'TrustBuy Badge', body: 'Displays supplier\'s quality validation score (e.g. 4.5★) directly on PDP. Builds immediate buyer confidence.' },
    { n: '02', title: 'Supplier Quality Dashboard', body: 'Supplier sees their Quality Score, return rates, and organic PRISM search boost status (+15% visibility).' },
    { n: '03', title: 'Catalog Action Center', body: 'Pinpoints specific SKUs causing return spikes and triggers suggestions: "Update fabric details to pre-shrunk to save ₹18K."' },
    { n: '04', title: 'PRISM Algorithmic Link', body: 'Products with return rates above category averages face automatic organic de-boosting in PRISM recommendations.' },
    { n: '05', title: 'Category Console Nudge', body: 'Allows Meesho category managers to track overall returns, audit listings, and send automated warning alerts to low-scoring suppliers.' },
  ];

  sl.addShape(prs.ShapeType.line, { x: 0.88, y: 1.82, w: 0, h: 4.2, line: { color: C.pink, width: 1.5, dashType: 'dash' } });

  steps.forEach((s, i) => {
    const y = 1.65 + i * 0.86;
    sl.addShape(prs.ShapeType.ellipse, { x: 0.6, y, w: 0.55, h: 0.55, fill: { color: C.pink } });
    sl.addText(s.n, { x: 0.6, y, w: 0.55, h: 0.55, fontSize: 10, color: C.white, bold: true, align: 'center', valign: 'middle' });
    sl.addText(s.title, { x: 1.3, y, w: 2.6, h: 0.3, fontSize: 11.5, color: C.pink, bold: true });
    sl.addText(s.body, { x: 1.3, y: y + 0.3, w: 8.1, h: 0.44, fontSize: 9.5, color: 'FFD5E5' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 6.18, w: 9.1, h: 0.62, fill: { color: C.hero } });
  sl.addText('PhonePe Parallel: Scaling supplier-side tools and incentive dashboards maps to rebuilding the rewards engine. Reorganizing onboarding paths mirrors catalog correction pipelines.', {
    x: 0.62, y: 6.22, w: 8.76, h: 0.52, fontSize: 9.5, color: C.accent,
  });
}

// ─────────────────────────────────────────
// SLIDE 5 — THE PRODUCT
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE PRODUCT', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.pink, bold: true, charSpacing: 2 });
  sl.addText('Interactive Prototypes — PDP Badging, Quality Scores, and Admin Controls', {
    x: 0.5, y: 0.62, w: 9, h: 0.55, fontSize: 22, color: C.dark, bold: true,
  });

  const screens = [
    { n: '01', title: 'PDP Trust Badge', desc: 'Displays TrustBuy Verification status, rating, and sizing validation index to consumers.' },
    { n: '02', title: 'Quality Scorecard', desc: 'Supplier Quality Dashboard showing score (92.5/100) and corresponding PRISM visibility boost.' },
    { n: '03', title: 'Catalog Actions', desc: 'Highlights SKUs with sizing or fabric issues and suggests 1-click description and size adjustments.' },
    { n: '04', title: 'Category Admin', desc: 'Meesho Category Manager dashboard to audit high-return sellers and send automated nudges.' },
  ];

  screens.forEach((s, i) => {
    const x = 0.45 + i * 2.38;
    sl.addShape(prs.ShapeType.rect, { x, y: 1.38, w: 2.18, h: 3.5, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(prs.ShapeType.rect, { x, y: 1.38, w: 2.18, h: 0.32, fill: { color: C.pink } });
    sl.addText(`${s.n} — ${s.title}`, { x, y: 1.38, w: 2.18, h: 0.32, fontSize: 9.5, color: C.white, bold: true, align: 'center', valign: 'middle' });

    const mockup = i === 0
      ? '👕 Cotton T-Shirts (3) \n─────────────────\nPrice: ₹348\n\n🛡️ TrustBuy Verified\n96% of buyers verified size\nfit & fabric quality\n\nRatings:\n[ 4.5★ ] [ 98% Size ] [ 94% Col ]'
      : i === 1
      ? '📊 Quality Scorecard\n─────────────────\nKuber Apparels (Level 4)\nScore: 92.5 / 100\n[██████████░] Top 8%\n\nPRISM Visibility Boost:\n+15% organic views\nReturn Rate: 11.2%'
      : i === 2
      ? '⚠️ SKU: COTTON-TSHIRT-L\n─────────────────\nAlert: Shrinkage Risk\n• Return Rate: 18.2%\n• Complaint: Shrinks in wash\n\nAI Fix: Update description:\n"100% Cotton. Buy 1 size up"\n[ Update Description ]'
      : '🤖 Category: Apparel\n─────────────────\nHigh Return Sellers:\n1. Shivay Creations (21.4%)\n   [ Nudge ]\n2. Kuber Apparels   (11.2%)\n   [ Verified ✓ ]\n3. Balaji Textiles  (16.8%)';

    sl.addText(mockup, { x: x + 0.08, y: 1.82, w: 2.02, h: 2.6, fontSize: 7.5, color: '333333', fontFace: 'Courier New' });
    sl.addText(s.desc, { x: x + 0.06, y: 4.94, w: 2.06, h: 0.8, fontSize: 9, color: '555555' });
  });

  sl.addText('Live prototype: meesho-trustbuy-prototype.html', { x: 0.45, y: 6.88, w: 9.1, h: 0.28, fontSize: 10, color: C.pink, align: 'center' });
}

// ─────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('IMPACT & ROI', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.pink, bold: true, charSpacing: 2 });
  sl.addText('Supplier Quality & E-Commerce Margin Improvement', {
    x: 0.5, y: 0.62, w: 9, h: 0.7, fontSize: 21, color: C.white, bold: true,
  });

  sl.addText('CONSUMER & SUPPLIER IMPACT', { x: 0.45, y: 1.5, w: 4.5, h: 0.3, fontSize: 9.5, color: C.pink, bold: true, charSpacing: 1 });
  sl.addText('BUSINESS & PLATFORM ROI', { x: 5.05, y: 1.5, w: 4.5, h: 0.3, fontSize: 9.5, color: C.pink, bold: true, charSpacing: 1 });

  const left = [
    { v: '↑ 15%', l: 'Consumer catalog trust index' },
    { v: '↓ 32%', l: 'Wrong-item/defective return rate' },
    { v: '↑ 22%', l: 'Repeat purchases (12-month)' },
    { v: '↑ 85%', l: 'Supplier satisfaction score' },
  ];
  const right = [
    { v: '↓ 18%', l: 'Logistics return processing costs' },
    { v: '↑ 12%', l: 'Organic catalog conversion rate' },
    { v: '↓ 45%', l: 'Manual blacklisting operations' },
    { v: '↑ 94%', l: 'Correct catalog measurements' },
  ];

  left.forEach((m, i) => {
    const y = 1.92 + i * 0.84;
    sl.addShape(prs.ShapeType.rect, { x: 0.45, y, w: 4.38, h: 0.72, fill: { color: C.hero }, shadow: makeShadow() });
    sl.addText(m.v, { x: 0.6, y: y + 0.05, w: 1.5, h: 0.62, fontSize: 26, color: C.pink, bold: true, valign: 'middle' });
    sl.addText(m.l, { x: 2.18, y: y + 0.05, w: 2.6, h: 0.62, fontSize: 12, color: C.white, valign: 'middle' });
  });
  right.forEach((m, i) => {
    const y = 1.92 + i * 0.84;
    sl.addShape(prs.ShapeType.rect, { x: 5.05, y, w: 4.48, h: 0.72, fill: { color: C.hero }, shadow: makeShadow() });
    sl.addText(m.v, { x: 5.2, y: y + 0.05, w: 1.5, h: 0.62, fontSize: 26, color: C.pink, bold: true, valign: 'middle' });
    sl.addText(m.l, { x: 6.78, y: y + 0.05, w: 2.7, h: 0.62, fontSize: 12, color: C.white, valign: 'middle' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.44, w: 9.1, h: 0.82, fill: { color: C.pink } });
  sl.addText('Aligning search visibility with quality metrics leverages Meesho\'s platform strengths. Higher quality listings prevent reverse logistical waste, increasing overall profitability.', {
    x: 0.62, y: 5.5, w: 8.76, h: 0.68, fontSize: 10.5, color: C.white,
  });
}

// ─────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('PROOF OF WORK', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.pink, bold: true, charSpacing: 2 });
  sl.addText('Leveraging PhonePe Partner Platforms to Scale TrustBuy', { x: 0.5, y: 0.62, w: 9, h: 0.68, fontSize: 26, color: C.dark, bold: true });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 1.42, w: 4.2, h: 4.08, fill: { color: C.dark }, shadow: makeShadow() });
  sl.addText('PhonePe Experience', { x: 0.62, y: 1.55, w: 3.86, h: 0.35, fontSize: 12.5, color: C.pink, bold: true });
  const proof = [
    'Rebuilt third-party rewards platform into an ML-driven self-serve marketplace onboarding 500+ brands',
    'Drove B2B onboarding optimization that cut offer operations TAT from 2 days to 30 minutes',
    'Collaborated with business teams to structure revenue and monetization frameworks for partners',
    'A/B tested campaign conversions, adjusting merchant-facing visibility variables dynamically',
    'Monitored B2B platform adoption metrics, using feedback loops to prioritize product roadmap',
  ];
  proof.forEach((t, i) => sl.addText(`• ${t}`, { x: 0.72, y: 2.0 + i * 0.65, w: 3.76, h: 0.56, fontSize: 10, color: C.white }));

  sl.addShape(prs.ShapeType.rect, { x: 4.82, y: 1.42, w: 4.7, h: 4.08, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('Meesho PM2 Relevance', { x: 4.98, y: 1.55, w: 4.38, h: 0.35, fontSize: 12.5, color: C.pink, bold: true });
  const maps = [
    ['ML-driven self-serve marketplace', '→ TrustBuy supplier dashboard metrics'],
    ['Offer operations TAT reduction', '→ Supplier Hub Action Center recommendations'],
    ['Partner visibility A/B testing', '→ Dynamic PRISM search boosts'],
    ['B2B adoption metric scaling', '→ Admin category console auditing'],
    ['Stakeholder alignment', '→ Aligning supplier actions with margins'],
  ];
  maps.forEach(([f, m], i) => {
    sl.addText(f, { x: 4.98, y: 2.0 + i * 0.65, w: 2.2, h: 0.56, fontSize: 10, color: '444444', bold: true });
    sl.addText(m, { x: 7.2, y: 2.0 + i * 0.65, w: 2.2, h: 0.56, fontSize: 10, color: '333333' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.62, w: 9.1, h: 0.65, fill: { color: C.pink } });
  sl.addText('"Building self-serve B2B tools is about empowering merchants. Surfacing returns data as catalog recommendations gives suppliers the direct power to correct and grow."', {
    x: 0.62, y: 5.66, w: 8.76, h: 0.56, fontSize: 10.5, color: C.white, italic: true,
  });
}

// ─────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('ROLLOUT PLAN', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.pink, bold: true, charSpacing: 2 });
  sl.addText('Scorecard Pilot → PRISM Re-ranking → Category Console Scaling', {
    x: 0.5, y: 0.62, w: 9, h: 0.65, fontSize: 24, color: C.white, bold: true,
  });

  const phases = [
    {
      phase: 'Phase 1', time: 'Month 1–2: Scorecard Launch',
      body: 'Release Quality Scorecard to fashion suppliers. Surfaced via Supplier Hub. Measure scorecard views and baseline catalog updates. Target: 10% description updates.',
    },
    {
      phase: 'Phase 2', time: 'Month 3–4: PRISM Ranking link',
      body: 'Integrate Quality Score into PRISM recommendation algorithm for 10% of search queries. Validate conversion rate and return rate shift. Target: 15% return reduction.',
    },
    {
      phase: 'Phase 3', time: 'Month 5–6: Full scaling & badges',
      body: 'Roll out TrustBuy badges to consumers. Scale console nudges to all category managers. Target: 30% wrong-item return reduction.',
    },
  ];

  phases.forEach((p, i) => {
    const x = 0.45 + i * 3.12;
    sl.addShape(prs.ShapeType.rect, { x, y: 1.5, w: 2.92, h: 3.4, fill: { color: C.hero }, shadow: makeShadow() });
    sl.addShape(prs.ShapeType.rect, { x, y: 1.5, w: 2.92, h: 0.52, fill: { color: C.pink } });
    sl.addText(p.phase, { x, y: 1.5, w: 2.92, h: 0.52, fontSize: 15, color: C.white, bold: true, align: 'center', valign: 'middle' });
    sl.addText(p.time, { x: x + 0.1, y: 2.12, w: 2.72, h: 0.35, fontSize: 11, color: C.accent, bold: true });
    sl.addText(p.body, { x: x + 0.1, y: 2.5, w: 2.72, h: 2.2, fontSize: 10, color: 'FFD5E5' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.06, w: 9.1, h: 0.72, fill: { color: C.hero } });
  sl.addText('Requirements: Integration with PRISM algorithm, Supplier Hub UI updates, 1 backend engineer, 1 search engineer, and 1 UI/UX designer.', {
    x: 0.62, y: 5.1, w: 8.76, h: 0.62, fontSize: 10.5, color: C.accent,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.9, w: 9.1, h: 0.62, fill: { color: C.pink } });
  sl.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  ajay-avaghade.github.io/Portfolio', {
    x: 0.62, y: 5.94, w: 8.76, h: 0.52, fontSize: 10.5, color: C.white, align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/meesho-trustbuy-deck.pptx' })
  .then(() => console.log('✅ assets/meesho-trustbuy-deck.pptx written'))
  .catch(e => { console.error(e); process.exit(1); });
