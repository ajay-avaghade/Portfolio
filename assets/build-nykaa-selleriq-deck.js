const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0F0A00',
  hero:   '1E1400',
  gold:   'F59E0B',
  amber:  'FCD34D',
  white:  'FFFFFF',
  lgray:  'FFFBEB',
  ink:    '1A1000',
  muted:  '92704A',
};
const makeShadow = () => ({ type: 'outer', color: '000000', blur: 3, offset: 1, angle: 45, opacity: 0.12 });

// ─────────────────────────────────────────
// SLIDE 1 — COVER
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  for (let i = 0; i < 8; i++) {
    sl.addShape(prs.ShapeType.line, { x: -0.5 + i * 1.5, y: 0, w: 4, h: 7.5, line: { color: 'F59E0B', width: 0.4, transparency: 88 } });
  }

  sl.addText('NYKAA SUPERSTORE — SELLER INTELLIGENCE', { x: 0.5, y: 0.32, w: 9, h: 0.28, fontSize: 9, color: C.amber, charSpacing: 2 });
  sl.addText('SellerIQ', { x: 0.5, y: 0.75, w: 9, h: 1.4, fontSize: 64, color: C.white, bold: true });
  sl.addText('AI-Powered Insights & Predictive Alerts That Tell Sellers What to Do Next', {
    x: 0.5, y: 2.0, w: 7.5, h: 0.6, fontSize: 22, color: C.amber, bold: false,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.5, y: 2.75, w: 2.8, h: 0.04, fill: { color: C.gold } });

  sl.addText('Presented by Ajay Avaghade · Growth & Monetization PM', {
    x: 0.5, y: 2.95, w: 7, h: 0.35, fontSize: 13, color: 'D4A855',
  });

  // Big stat
  sl.addShape(prs.ShapeType.rect, { x: 7.2, y: 5.2, w: 2.5, h: 1.9, fill: { color: C.gold }, shadow: makeShadow() });
  sl.addText('78%', { x: 7.2, y: 5.3, w: 2.5, h: 0.85, fontSize: 52, color: C.white, bold: true, align: 'center' });
  sl.addText('of sellers don\'t know\nwhy their GMV dropped', { x: 7.2, y: 6.1, w: 2.5, h: 0.8, fontSize: 11, color: 'FFF0CC', align: 'center' });
}

// ─────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE PROBLEM', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.amber, bold: true, charSpacing: 2 });
  sl.addText('Sellers See GMV Drop. They Don\'t Know Why. By the Time They Investigate, the Sale Window Has Passed.', {
    x: 0.5, y: 0.62, w: 9, h: 0.9, fontSize: 21, color: C.white, bold: true,
  });

  const cols = [
    { icon: '📊', stat: '78%', lbl: 'Can\'t explain GMV drops', sub: 'Raw metrics (impressions, CVR, returns) don\'t connect. Sellers see numbers, not causes.' },
    { icon: '⏳', stat: '3 days', lbl: 'Average investigation lag', sub: 'By the time a seller identifies a stockout risk or return spike, it\'s already costing them rank.' },
    { icon: '🏆', stat: '2.3×', lbl: 'Top-10% vs median GMV gap', sub: 'Top sellers act on data weekly. Median sellers react monthly. The gap is operational, not product.' },
  ];
  cols.forEach((c, i) => {
    const x = 0.45 + i * 3.12;
    sl.addShape(prs.ShapeType.rect, { x, y: 1.72, w: 2.9, h: 3.1, fill: { color: C.hero }, shadow: makeShadow() });
    sl.addText(c.icon, { x, y: 1.88, w: 2.9, h: 0.55, fontSize: 26, align: 'center' });
    sl.addText(c.stat, { x, y: 2.52, w: 2.9, h: 0.72, fontSize: 40, color: C.gold, bold: true, align: 'center' });
    sl.addText(c.lbl, { x, y: 3.22, w: 2.9, h: 0.38, fontSize: 12.5, color: C.white, bold: true, align: 'center' });
    sl.addText(c.sub, { x: x + 0.1, y: 3.62, w: 2.7, h: 0.9, fontSize: 10, color: 'D4A855', align: 'center' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.0, w: 9.1, h: 0.82, fill: { color: C.gold } });
  sl.addText('Root cause: Nykaa\'s seller dashboard shows raw data — GMV, impressions, returns. It does not show root causes, competitive context, or what action to take next. Sellers need a daily intelligence brief, not a spreadsheet.', {
    x: 0.62, y: 5.06, w: 8.76, h: 0.7, fontSize: 10.5, color: C.white,
  });
}

// ─────────────────────────────────────────
// SLIDE 3 — THE INSIGHT
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE INSIGHT', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.gold, bold: true, charSpacing: 2 });
  sl.addText('Sellers Don\'t Need More Data. They Need an AI That Reads the Data and Tells Them What\'s Wrong.', {
    x: 0.5, y: 0.62, w: 9, h: 0.9, fontSize: 22, color: C.ink, bold: true,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 1.72, w: 4.1, h: 3.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('❌  Current Dashboard', { x: 0.62, y: 1.85, w: 3.76, h: 0.38, fontSize: 12, color: '444444', bold: true });
  const left = ['Raw GMV, impressions, returns — no context', 'No root cause linking (why did CVR drop?)', 'No competitive benchmarking by category', 'Seller must manually correlate metrics across tabs', 'No proactive alerts — stockout discovered after rank drops'];
  left.forEach((t, i) => sl.addText(`• ${t}`, { x: 0.72, y: 2.35 + i * 0.42, w: 3.56, h: 0.36, fontSize: 11, color: '555555' }));

  sl.addShape(prs.ShapeType.ellipse, { x: 4.62, y: 2.9, w: 0.72, h: 0.72, fill: { color: C.gold } });
  sl.addText('VS', { x: 4.62, y: 2.9, w: 0.72, h: 0.72, fontSize: 11, color: C.white, bold: true, align: 'center', valign: 'middle' });

  sl.addShape(prs.ShapeType.rect, { x: 5.42, y: 1.72, w: 4.1, h: 3.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('✅  SellerIQ (Proposed)', { x: 5.58, y: 1.85, w: 3.76, h: 0.38, fontSize: 12, color: C.gold, bold: true });
  const right = ['Daily AI brief: "GMV down 12% — here are the 3 causes"', 'Natural language Q&A: "Why did my Vitamin C Serum drop?"', 'Predictive alerts: "Stockout in 3 days — 144 units left"', 'Competitive benchmarking: "Your CVR 2.3% vs. 4.1% category avg"', 'Weekly Action Cards: ranked by estimated GMV impact'];
  right.forEach((t, i) => sl.addText(`• ${t}`, { x: 5.58, y: 2.35 + i * 0.42, w: 3.72, h: 0.36, fontSize: 10.5, color: '333333' }));

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.08, w: 9.1, h: 0.72, fill: { color: C.gold } });
  sl.addText('Key insight: The data to answer every seller question already exists on Nykaa — search rank history, return rates, competitor pricing, stockout signals. SellerIQ surfaces it proactively, in plain language, before the seller even knows to ask.', {
    x: 0.62, y: 5.12, w: 8.76, h: 0.62, fontSize: 10, color: C.white,
  });
}

// ─────────────────────────────────────────
// SLIDE 4 — THE MECHANIC
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE MECHANIC', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.amber, bold: true, charSpacing: 2 });
  sl.addText('Four Intelligence Layers: Brief → Q&A → Alerts → Benchmark', {
    x: 0.5, y: 0.62, w: 9, h: 0.65, fontSize: 22, color: C.white, bold: true,
  });

  const steps = [
    { n: '01', title: 'Daily AI Brief', body: 'Every morning: GMV delta vs. last week, top 3 root causes (ranked by impact), and a one-line action for each. No charts to interpret — just: "Your CVR dropped because your search rank fell from page 1 to page 3."' },
    { n: '02', title: 'Natural Language Q&A', body: 'Seller asks: "Why did my Vitamin C Serum drop this week?" AI responds with 3 root causes, each linked to a data source: rank change, return spike, competitor price cut. Follow-up: "What should I do first?"' },
    { n: '03', title: 'Predictive Alerts', body: 'ML model monitors 6 signals: stockout risk, return rate spike, rank velocity, competitor pricing, seasonal demand surge, ad budget exhaustion. Alert 48–72 hours before the event becomes a problem.' },
    { n: '04', title: 'Competitive Benchmarking', body: 'Category-level CVR, rating, return rate, and pricing benchmarks. "Your CVR is 2.3% vs. 4.1% category average — sellers above 4% share these 3 listing attributes." Actionable gap, not raw data.' },
    { n: '05', title: 'Weekly Action Cards', body: 'Every Sunday: 3–5 Action Cards ranked by estimated GMV impact. Each card: problem + root cause + recommended action + estimated uplift. One-tap to act (restock / run ad / fix listing).' },
  ];

  sl.addShape(prs.ShapeType.line, { x: 0.88, y: 1.82, w: 0, h: 4.2, line: { color: C.gold, width: 1.5, dashType: 'dash' } });

  steps.forEach((s, i) => {
    const y = 1.65 + i * 0.86;
    sl.addShape(prs.ShapeType.ellipse, { x: 0.6, y, w: 0.55, h: 0.55, fill: { color: C.gold } });
    sl.addText(s.n, { x: 0.6, y, w: 0.55, h: 0.55, fontSize: 10, color: C.white, bold: true, align: 'center', valign: 'middle' });
    sl.addText(s.title, { x: 1.3, y, w: 2.6, h: 0.3, fontSize: 11.5, color: C.gold, bold: true });
    sl.addText(s.body, { x: 1.3, y: y + 0.3, w: 8.1, h: 0.44, fontSize: 9.5, color: 'D4A855' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 6.18, w: 9.1, h: 0.62, fill: { color: C.hero } });
  sl.addText('PhonePe Proof: Built Propensity-to-Transact CDP that replaced static rules with real-time user-level signals — same architecture: data signals → ML root cause → automated action recommendation (32% marketing burn reduction, GMV maintained).', {
    x: 0.62, y: 6.22, w: 8.76, h: 0.52, fontSize: 9.5, color: C.amber,
  });
}

// ─────────────────────────────────────────
// SLIDE 5 — THE PRODUCT
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE PRODUCT', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.gold, bold: true, charSpacing: 2 });
  sl.addText('4 Key Screen States — Intelligence to Action in One Flow', {
    x: 0.5, y: 0.62, w: 9, h: 0.55, fontSize: 22, color: C.ink, bold: true,
  });

  const screens = [
    { n: '01', title: 'Insights Home', desc: 'AI summary of week\'s performance, 3 KPI mini-cards, and ranked action cards. From landing to first action: under 60 seconds.' },
    { n: '02', title: 'AI Q&A', desc: 'Natural language: "Why did my serum drop?" → 3 root causes (rank, return rate, competitor pricing) with source data linked.' },
    { n: '03', title: 'Alert Detail', desc: 'Stockout risk alert: 144 units, 3 days left, velocity chart (Mon–Sun trending). 3-step resolution flow. "Raise Restock Order" CTA.' },
    { n: '04', title: 'Benchmark View', desc: 'CVR 2.3% vs. 4.1% category vs. 6.4% top-10%. Bar charts. Gap-closing action cards: "Sellers above 4% share 3 listing attributes."' },
  ];

  screens.forEach((s, i) => {
    const x = 0.45 + i * 2.38;
    sl.addShape(prs.ShapeType.rect, { x, y: 1.38, w: 2.18, h: 3.5, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(prs.ShapeType.rect, { x, y: 1.38, w: 2.18, h: 0.32, fill: { color: C.gold } });
    sl.addText(`${s.n} — ${s.title}`, { x, y: 1.38, w: 2.18, h: 0.32, fontSize: 9.5, color: C.white, bold: true, align: 'center', valign: 'middle' });

    const mockup = i === 0
      ? '🤖 GMV ↓₹18,400 (−12%)\n─────────────────\n🔴 Stockout risk · High\n🟡 Image quality · Med\n🟢 Price competitive\n─────────────────\n📦 Action: Restock Vit C\n🖼️ Action: Add 3 images\n[ View All Actions ]'
      : i === 1
      ? '💬 "Why did my Vit C\n     Serum drop this week?"\n─────────────────\n🤖 3 root causes:\n1. Rank: Page1 → Page3\n   (−62% impressions)\n2. Return rate +6%\n   (4 reviews flagged)\n3. Competitor: −15%\n─────────────────\n[ What should I do? ]'
      : i === 2
      ? '⚠️ Stockout Risk · HIGH\nVit C Serum\n─────────────────\n144 units · 3 days left\n📈 Velocity trending up\n Mon ██ 28\n Thu ████ 42\n Sun █████ 52 /day\n─────────────────\n🟢 [ Raise Restock Order ]'
      : '📊 CVR Benchmarks\n─────────────────\n   You  ██ 2.3%\n  Cat  ████ 4.1%\n Top  ██████ 6.4%\n─────────────────\n⭐ Ratings\n   You: 4.2★ vs 3.8★\n─────────────────\n[ See Gap-Closing Tips ]';

    sl.addText(mockup, { x: x + 0.08, y: 1.82, w: 2.02, h: 2.6, fontSize: 7.5, color: '333333', fontFace: 'Courier New' });
    sl.addText(s.desc, { x: x + 0.06, y: 4.94, w: 2.06, h: 0.8, fontSize: 9, color: '555555' });
  });

  sl.addText('Live prototype: nykaa-selleriq-prototype.html', { x: 0.45, y: 6.88, w: 9.1, h: 0.28, fontSize: 10, color: C.gold, align: 'center' });
}

// ─────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('IMPACT & ROI', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.amber, bold: true, charSpacing: 2 });
  sl.addText('Projected Impact — Grounded in PhonePe CDP & Suppression Platform Proof Points', {
    x: 0.5, y: 0.62, w: 9, h: 0.7, fontSize: 21, color: C.white, bold: true,
  });

  sl.addText('SELLER IMPACT', { x: 0.45, y: 1.5, w: 4.5, h: 0.3, fontSize: 9.5, color: C.gold, bold: true, charSpacing: 1 });
  sl.addText('NYKAA PLATFORM ROI', { x: 5.05, y: 1.5, w: 4.5, h: 0.3, fontSize: 9.5, color: C.gold, bold: true, charSpacing: 1 });

  const left = [
    { v: '↑ 28%', l: 'GMV per seller (alert-active)' },
    { v: '↓ 65%', l: 'Stockout-related rank drops' },
    { v: '↑ 18%', l: 'CVR via listing improvements' },
    { v: '↑ 3.2×', l: 'Seller dashboard DAU' },
  ];
  const right = [
    { v: '↓ 40%', l: 'Return rate (root cause closed faster)' },
    { v: '₹0', l: 'Incremental seller support ops cost' },
    { v: '↑ 22%', l: 'Seller retention (12-month)' },
    { v: '↑ 15%', l: 'Catalog health score platform-wide' },
  ];

  left.forEach((m, i) => {
    const y = 1.92 + i * 0.84;
    sl.addShape(prs.ShapeType.rect, { x: 0.45, y, w: 4.38, h: 0.72, fill: { color: C.hero }, shadow: makeShadow() });
    sl.addText(m.v, { x: 0.6, y: y + 0.05, w: 1.5, h: 0.62, fontSize: 26, color: C.gold, bold: true, valign: 'middle' });
    sl.addText(m.l, { x: 2.18, y: y + 0.05, w: 2.6, h: 0.62, fontSize: 12, color: C.white, valign: 'middle' });
  });
  right.forEach((m, i) => {
    const y = 1.92 + i * 0.84;
    sl.addShape(prs.ShapeType.rect, { x: 5.05, y, w: 4.48, h: 0.72, fill: { color: C.hero }, shadow: makeShadow() });
    sl.addText(m.v, { x: 5.2, y: y + 0.05, w: 1.5, h: 0.62, fontSize: 26, color: C.gold, bold: true, valign: 'middle' });
    sl.addText(m.l, { x: 6.78, y: y + 0.05, w: 2.7, h: 0.62, fontSize: 12, color: C.white, valign: 'middle' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.44, w: 9.1, h: 0.82, fill: { color: C.gold } });
  sl.addText('Every GMV point recovered by a seller is platform GMV. SellerIQ is a zero-subsidy lever: it makes existing inventory and listings perform better — no discounts, no new traffic acquisition, just execution intelligence.', {
    x: 0.62, y: 5.5, w: 8.76, h: 0.68, fontSize: 10.5, color: C.white,
  });
}

// ─────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('PROOF OF WORK', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.gold, bold: true, charSpacing: 2 });
  sl.addText('I Built This. Here\'s the Proof.', { x: 0.5, y: 0.62, w: 9, h: 0.68, fontSize: 26, color: C.ink, bold: true });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 1.42, w: 4.2, h: 4.08, fill: { color: C.dark }, shadow: makeShadow() });
  sl.addText('What I Built at PhonePe', { x: 0.62, y: 1.55, w: 3.86, h: 0.35, fontSize: 12.5, color: C.gold, bold: true });
  const proof = [
    'Built Propensity-to-Transact CDP: real-time user-level scoring replacing ₹1,000+ Cr/yr static rules — same root-cause signal architecture as SellerIQ alerts',
    'Diagnosed 60%+ cart abandonment: identified 3 root causes (cart value, intent, time signals) → built targeted intervention engine (35% AOV uplift)',
    'Owned full DS collaboration cycle: feature engineering → model → production → monitoring with clear P&L accountability',
    'Built offer eligibility surfacing across 350M MAU with real-time signal-to-UI pipeline — same data-to-insight layer as SellerIQ',
    'A/B tested 6 campaign variants; built performance dashboards used by 15-person ops team weekly',
  ];
  proof.forEach((t, i) => sl.addText(`• ${t}`, { x: 0.72, y: 2.0 + i * 0.65, w: 3.76, h: 0.56, fontSize: 10, color: C.white }));

  sl.addShape(prs.ShapeType.rect, { x: 4.82, y: 1.42, w: 4.7, h: 4.08, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('How It Maps to This JD', { x: 4.98, y: 1.55, w: 4.38, h: 0.35, fontSize: 12.5, color: C.gold, bold: true });
  const maps = [
    ['Propensity-to-Transact CDP', '→ SellerIQ predictive alert engine'],
    ['Root cause diagnosis (cart abandon)', '→ SellerIQ daily AI brief + Q&A'],
    ['Signal-to-UI pipeline (350M MAU)', '→ SellerIQ real-time insights layer'],
    ['DS collaboration (ML in production)', '→ SellerIQ ML model scoping'],
    ['Performance dashboards for ops', '→ SellerIQ benchmark + action cards'],
  ];
  maps.forEach(([f, m], i) => {
    sl.addText(f, { x: 4.98, y: 2.0 + i * 0.65, w: 2.2, h: 0.56, fontSize: 10, color: '444444', bold: true });
    sl.addText(m, { x: 7.2, y: 2.0 + i * 0.65, w: 2.2, h: 0.56, fontSize: 10, color: '333333' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.62, w: 9.1, h: 0.65, fill: { color: C.gold } });
  sl.addText('"SellerIQ is the seller-facing version of the same data pipeline I shipped at PhonePe — where signal becomes recommendation becomes action, automatically."', {
    x: 0.62, y: 5.66, w: 8.76, h: 0.56, fontSize: 10.5, color: C.white, italic: true,
  });
}

// ─────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('ROLLOUT PLAN', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.amber, bold: true, charSpacing: 2 });
  sl.addText('Phased Rollout — Alert Layer → Q&A → Benchmark Intelligence', {
    x: 0.5, y: 0.62, w: 9, h: 0.65, fontSize: 24, color: C.white, bold: true,
  });

  const phases = [
    {
      phase: 'Phase 1', time: 'Month 1–2: Stockout Alerts',
      body: 'Launch predictive stockout alert to 1,000 sellers. Instrument: alert open rate, restock action rate, rank recovery vs. control group. Target: 60% stockout prevented.',
    },
    {
      phase: 'Phase 2', time: 'Month 3–4: Daily Brief + Q&A',
      body: 'Activate AI daily brief (3 root causes per seller). Launch natural language Q&A. Measure: dashboard DAU, seller action rate on recommendations, GMV delta (alert-active vs. passive).',
    },
    {
      phase: 'Phase 3', time: 'Month 5–6: Benchmarks + Scale',
      body: 'Launch competitive benchmarking by category. Roll SellerIQ to all sellers. Integrate Action Cards into Seller App home. Target: 28% GMV uplift for SellerIQ-active sellers.',
    },
  ];

  phases.forEach((p, i) => {
    const x = 0.45 + i * 3.12;
    sl.addShape(prs.ShapeType.rect, { x, y: 1.5, w: 2.92, h: 3.4, fill: { color: C.hero }, shadow: makeShadow() });
    sl.addShape(prs.ShapeType.rect, { x, y: 1.5, w: 2.92, h: 0.52, fill: { color: C.gold } });
    sl.addText(p.phase, { x, y: 1.5, w: 2.92, h: 0.52, fontSize: 15, color: C.white, bold: true, align: 'center', valign: 'middle' });
    sl.addText(p.time, { x: x + 0.1, y: 2.12, w: 2.72, h: 0.35, fontSize: 11, color: C.amber, bold: true });
    sl.addText(p.body, { x: x + 0.1, y: 2.5, w: 2.72, h: 2.2, fontSize: 10, color: 'D4A855' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.06, w: 9.1, h: 0.72, fill: { color: C.hero } });
  sl.addText('What I need to build this: Seller transaction + inventory data feed · 1 ML engineer (stockout predictor + root cause model) · 1 NLP engineer (Q&A layer) · Seller UX designer for brief + action card flows', {
    x: 0.62, y: 5.1, w: 8.76, h: 0.62, fontSize: 10.5, color: C.amber,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.9, w: 9.1, h: 0.62, fill: { color: C.gold } });
  sl.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  ajay-avaghade.github.io/Portfolio', {
    x: 0.62, y: 5.94, w: 8.76, h: 0.52, fontSize: 10.5, color: C.white, align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/nykaa-selleriq-deck.pptx' })
  .then(() => console.log('✅ nykaa-selleriq-deck.pptx written'))
  .catch(e => { console.error(e); process.exit(1); });
