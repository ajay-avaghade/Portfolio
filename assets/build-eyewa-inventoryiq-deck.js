// build-eyewa-inventoryiq-deck.js
// eyewa InventoryIQ — ML Demand Forecasting & Smart Reorder
// Run: node assets/build-eyewa-inventoryiq-deck.js

const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0A0714',
  hero:   '0D0B1A',
  card:   '130F22',
  purple: '8B5CF6',
  amber:  'F59E0B',
  green:  '22C55E',
  red:    'EF4444',
  white:  'FFFFFF',
  muted:  '94A3B8',
  lgray:  'F1F5F9',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.18 });

// ─── SLIDE 1: COVER ───────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  for (let i = 0; i < 8; i++) {
    s.addShape(prs.ShapeType.line, { x: -1 + i * 1.6, y: 0, w: 5, h: 7.5, line: { color: C.purple, width: 0.4, transparency: 88 }, rotate: 28 });
  }

  s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.purple } });
  s.addText('EYEWA · SUPPLY CHAIN INTELLIGENCE', { x: 0.6, y: 0.55, w: 7, h: 0.25, fontSize: 9, color: C.purple, bold: true, charSpacing: 3 });
  s.addText('InventoryIQ', { x: 0.6, y: 0.88, w: 9, h: 1.35, fontSize: 66, color: C.white, bold: true });
  s.addText('ML Demand Forecasting & Smart Reorder', { x: 0.6, y: 2.1, w: 7, h: 0.5, fontSize: 22, color: C.muted });
  s.addText('Replacing gut-feel reorder decisions with ML-powered demand forecasting — so eyewa always has the right stock in the right market at the right time, from UAE to KSA and beyond.', {
    x: 0.6, y: 2.75, w: 6.2, h: 0.9, fontSize: 13, color: C.muted, lineSpacingMultiple: 1.4,
  });
  s.addText('Presented by Ajay Avaghade · Growth & Operations PM', { x: 0.6, y: 3.82, w: 6, h: 0.28, fontSize: 11, color: '475569' });

  s.addShape(prs.ShapeType.rect, { x: 7.4, y: 1.8, w: 2.2, h: 2.2, fill: { color: C.purple, transparency: 90 }, line: { color: C.purple, width: 1 }, shadow: makeShadow() });
  s.addText('6,066', { x: 7.4, y: 2.0, w: 2.2, h: 0.9, fontSize: 42, color: C.purple, bold: true, align: 'center' });
  s.addText('active SKUs\nmanaged by\nrule-based logic', { x: 7.4, y: 2.85, w: 2.2, h: 0.8, fontSize: 10, color: C.muted, align: 'center', lineSpacingMultiple: 1.35 });
}

// ─── SLIDE 2: THE PROBLEM ─────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.hero };

  s.addText('THE PROBLEM', { x: 0.55, y: 0.35, w: 9, h: 0.22, fontSize: 9, color: C.purple, bold: true, charSpacing: 3 });
  s.addText('6,066 SKUs. 4 Markets. Zero ML. Pure Guesswork.', {
    x: 0.55, y: 0.62, w: 9, h: 0.6, fontSize: 26, color: C.white, bold: true,
  });

  const stats = [
    { icon: '📦', num: '12', lbl: 'SKUs in stockout today', sub: 'Lost revenue on bestsellers while slow-movers collect dust in the warehouse' },
    { icon: '📉', num: '10%', lbl: 'Dead stock risk', sub: '608 SKUs with no projected reorder need — capital locked in unsellable inventory' },
    { icon: '🌍', num: '0', lbl: 'Data in new markets', sub: 'KSA, Bahrain, Kuwait launches run on analogies and gut feel — no historical signal' },
  ];

  stats.forEach((st, i) => {
    const x = 0.55 + i * 3.15;
    s.addShape(prs.ShapeType.rect, { x, y: 1.62, w: 2.9, h: 2.85, fill: { color: C.card }, line: { color: '1e1635', width: 0.8 }, shadow: makeShadow() });
    s.addText(st.icon, { x, y: 1.78, w: 2.9, h: 0.45, fontSize: 22, align: 'center' });
    s.addText(st.num, { x, y: 2.22, w: 2.9, h: 0.75, fontSize: 38, color: C.red, bold: true, align: 'center' });
    s.addText(st.lbl, { x, y: 2.95, w: 2.9, h: 0.38, fontSize: 11, color: C.white, bold: true, align: 'center' });
    s.addText(st.sub, { x: x + 0.1, y: 3.36, w: 2.7, h: 0.88, fontSize: 9.5, color: C.muted, align: 'center', lineSpacingMultiple: 1.4 });
  });

  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 4.62, w: 9.0, h: 0.88, fill: { color: '0a0414' }, line: { color: C.amber, width: 0.8 } });
  s.addText('Root cause: eyewa reorder decisions are driven by fixed MIN/MAX thresholds — the same logic regardless of brand, season, market maturity, or upcoming events. A 6,066-SKU catalogue at international scale demands ML.', {
    x: 0.75, y: 4.72, w: 8.6, h: 0.68, fontSize: 11, color: C.amber, lineSpacingMultiple: 1.4,
  });
}

// ─── SLIDE 3: THE INSIGHT ─────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.white };

  s.addText('THE INSIGHT', { x: 0.55, y: 0.35, w: 9, h: 0.22, fontSize: 9, color: C.purple, bold: true, charSpacing: 3 });
  s.addText('Fixed Reorder Rules vs. SKU-Level ML Demand Intelligence', {
    x: 0.55, y: 0.62, w: 9, h: 0.6, fontSize: 26, color: '0f172a', bold: true,
  });

  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 1.42, w: 4.2, h: 3.4, fill: { color: 'fef2f2' }, line: { color: 'fca5a5', width: 0.8 } });
  s.addText('CURRENT STATE', { x: 0.75, y: 1.58, w: 3.8, h: 0.25, fontSize: 9, color: C.red, bold: true, charSpacing: 2 });
  s.addText('Rule-Based Reorder Logic', { x: 0.75, y: 1.88, w: 3.8, h: 0.35, fontSize: 15, color: '1e293b', bold: true });
  [
    'Same MIN stock threshold for every SKU regardless of demand velocity',
    'Seasonal spikes (Eid, Dubai Shopping Festival) missed until stockout',
    'New market launches seeded with UAE mix — wrong for local preferences',
    'Buyer reviews 6,000+ SKUs manually every reorder cycle',
    'Dead stock only identified after cash has already been tied up',
  ].forEach((t, i) => s.addText('✗  ' + t, { x: 0.75, y: 2.32 + i * 0.42, w: 3.8, h: 0.38, fontSize: 10.5, color: '64748b', lineSpacingMultiple: 1.3 }));

  s.addShape(prs.ShapeType.ellipse, { x: 4.52, y: 2.55, w: 0.96, h: 0.96, fill: { color: C.purple }, shadow: makeShadow() });
  s.addText('VS', { x: 4.52, y: 2.55, w: 0.96, h: 0.96, fontSize: 13, color: C.white, bold: true, align: 'center', valign: 'middle' });

  s.addShape(prs.ShapeType.rect, { x: 5.25, y: 1.42, w: 4.2, h: 3.4, fill: { color: 'faf5ff' }, line: { color: 'c4b5fd', width: 0.8 } });
  s.addText('INVENTORYIQ', { x: 5.45, y: 1.58, w: 3.8, h: 0.25, fontSize: 9, color: C.purple, bold: true, charSpacing: 2 });
  s.addText('Per-SKU ML Demand Intelligence', { x: 5.45, y: 1.88, w: 3.8, h: 0.35, fontSize: 15, color: '1e293b', bold: true });
  [
    'ML predicts demand per SKU per week incorporating seasonality + trend',
    'Eid / Dubai Shopping Festival uplift baked in 6 weeks in advance',
    'New market seeds calibrated from analogous-market launch curves',
    'Buyer reviews AI-drafted POs — approves rather than calculates',
    'Dead stock flagged 8–12 weeks before cash is over-committed',
  ].forEach((t, i) => s.addText('✓  ' + t, { x: 5.45, y: 2.32 + i * 0.42, w: 3.8, h: 0.38, fontSize: 10.5, color: '5b21b6', lineSpacingMultiple: 1.3 }));

  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 5.0, w: 8.9, h: 0.72, fill: { color: 'faf5ff' }, line: { color: C.purple, width: 0.8 } });
  s.addText('Key insight: The bottleneck is not buying budget — it is information latency. InventoryIQ compresses the signal-to-decision cycle from weeks to hours, so eyewa can buy confidently rather than reactively.', {
    x: 0.75, y: 5.1, w: 8.5, h: 0.52, fontSize: 10.5, color: '5b21b6', lineSpacingMultiple: 1.4,
  });
}

// ─── SLIDE 4: THE MECHANIC ────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE MECHANIC', { x: 0.55, y: 0.35, w: 9, h: 0.22, fontSize: 9, color: C.purple, bold: true, charSpacing: 3 });
  s.addText('Five Intelligence Layers From SKU Signal to Purchase Order', {
    x: 0.55, y: 0.62, w: 9, h: 0.55, fontSize: 26, color: C.white, bold: true,
  });

  const steps = [
    { n: '01', title: 'Data Ingestion', body: 'Sales velocity, returns, page views, search data, and promo calendar pulled from Shopify, ERP, and analytics — per SKU, per market, daily.' },
    { n: '02', title: 'ML Forecasting', body: 'Time-series ML model (LSTM + gradient boosting ensemble) predicts 30/60/90-day demand per SKU with confidence intervals. Seasonality, trend, and event uplift built in.' },
    { n: '03', title: 'Smart PO Draft', body: 'System auto-calculates reorder quantity (demand forecast minus in-transit stock minus safety buffer), drafts PO with supplier details, MOQ validation, and AI reasoning.' },
    { n: '04', title: 'Buyer Approval', body: 'Buyer reviews AI-drafted POs in a single queue — approve, adjust quantity, or reject with reason. One-click send to supplier. No manual calculation required.' },
    { n: '05', title: 'Market Playbook', body: 'For new market launches, ML borrows demand curves from analogous existing markets and calibrates for population size, category preference, and seasonality delta.' },
  ];

  steps.forEach((st, i) => {
    const x = 0.3 + i * 1.88;
    s.addShape(prs.ShapeType.rect, { x, y: 1.45, w: 1.68, h: 3.55, fill: { color: C.card }, line: { color: '1e1635', width: 0.8 }, shadow: makeShadow() });
    s.addShape(prs.ShapeType.ellipse, { x: x + 0.44, y: 1.52, w: 0.8, h: 0.8, fill: { color: C.purple, transparency: 85 }, line: { color: C.purple, width: 0.8 } });
    s.addText(st.n, { x, y: 1.52, w: 1.68, h: 0.8, fontSize: 16, color: C.purple, bold: true, align: 'center', valign: 'middle' });
    s.addText(st.title, { x: x + 0.08, y: 2.42, w: 1.52, h: 0.44, fontSize: 11, color: C.white, bold: true, align: 'center', lineSpacingMultiple: 1.2 });
    s.addText(st.body, { x: x + 0.1, y: 2.9, w: 1.48, h: 1.85, fontSize: 8.8, color: C.muted, align: 'center', lineSpacingMultiple: 1.38 });
    if (i < 4) s.addShape(prs.ShapeType.line, { x: x + 1.68, y: 2.22, w: 0.2, h: 0, line: { color: C.purple, width: 1, dashType: 'dash' } });
  });

  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 5.12, w: 9.0, h: 0.62, fill: { color: '07051A' }, line: { color: C.purple, width: 0.6 } });
  s.addText('PhonePe proof point: ML-driven marketplace with audience segmentation replaced static rules across a 350M+ MAU platform — same signal-to-decision compression applied to inventory replenishment.', {
    x: 0.75, y: 5.22, w: 8.6, h: 0.42, fontSize: 10, color: C.purple, lineSpacingMultiple: 1.3,
  });
}

// ─── SLIDE 5: THE PRODUCT ─────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.white };

  s.addText('THE PRODUCT', { x: 0.55, y: 0.35, w: 9, h: 0.22, fontSize: 9, color: C.purple, bold: true, charSpacing: 3 });
  s.addText('4 Console Views — From SKU Health to Market Launch', {
    x: 0.55, y: 0.62, w: 9, h: 0.5, fontSize: 26, color: '0f172a', bold: true,
  });

  const screens = [
    { n: '01', title: 'SKU Health Dashboard', desc: 'Traffic-light status across all 6,066 SKUs. Filters by category, status, reorder urgency. Stockout and low-stock SKUs surfaced at the top with action buttons.', mock: ['Healthy    4,064 🟢', 'Low Stock  1,394 🟡', 'Stockouts     12 🔴', 'Dead Risk    608', '──────────────', 'Oakley Holbrook 🔴', 'Ray-Ban RB3025 ⚠'] },
    { n: '02', title: 'Demand Forecast', desc: 'Per-SKU demand chart (actual + ML projected) with 30/60/90-day horizons and confidence intervals. Seasonal events flagged on the chart. Reorder urgency calculated automatically.', mock: ['RB3025 Aviator', '30d: 47 ±8 units', '60d: 89 ±12 units', '90d: 134 ±18 units', '──────────────', 'Eid uplift: +34%', 'Reorder in: 8 days ⚠'] },
    { n: '03', title: 'Smart PO Generator', desc: 'AI-drafted purchase orders in a review queue. Each PO shows AI reasoning, quantity recommendation, supplier, MOQ validation, and total value. One-click approve or adjust.', mock: ['3 POs Pending', 'Total: AED 38,400', '──────────────', 'Oakley Critical ↑', 'Ray-Ban High ↑', 'Varilux High ↑', '[ Approve All 3 ]'] },
    { n: '04', title: 'New Market Playbook', desc: 'For new market launches, shows demand curve projected from analogous markets. Recommends seed inventory mix calibrated for local category preferences and market size.', mock: ['KSA · Riyadh Launch', 'Analogy: UAE +20%', '──────────────', 'Frames: 45% seed', 'Sunglasses: 31%', 'Contact L: 24%', 'Break-even: M4'] },
  ];

  screens.forEach((sc, i) => {
    const x = 0.35 + i * 2.38;
    s.addShape(prs.ShapeType.rect, { x, y: 1.3, w: 2.2, h: 4.4, fill: { color: 'f8fafc' }, line: { color: 'd1d5db', width: 0.6 }, shadow: makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y: 1.3, w: 2.2, h: 0.32, fill: { color: C.purple } });
    s.addText(sc.n + ' — ' + sc.title, { x: x + 0.08, y: 1.3, w: 2.04, h: 0.32, fontSize: 9, color: C.white, bold: true, valign: 'middle' });
    s.addText(sc.desc, { x: x + 0.1, y: 1.68, w: 2.0, h: 1.1, fontSize: 8.8, color: '475569', lineSpacingMultiple: 1.35 });
    s.addShape(prs.ShapeType.rect, { x: x + 0.1, y: 2.82, w: 2.0, h: 1.75, fill: { color: '07051A' }, line: { color: '1e1635', width: 0.5 } });
    s.addText(sc.mock.join('\n'), { x: x + 0.14, y: 2.88, w: 1.92, h: 1.65, fontSize: 7.8, color: '94a3b8', fontFace: 'Courier New', lineSpacingMultiple: 1.45 });
  });
}

// ─── SLIDE 6: IMPACT & ROI ────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.hero };

  s.addText('IMPACT & ROI', { x: 0.55, y: 0.35, w: 9, h: 0.22, fontSize: 9, color: C.purple, bold: true, charSpacing: 3 });
  s.addText('From Reactive Buying to Precision Capital Deployment', {
    x: 0.55, y: 0.62, w: 9, h: 0.55, fontSize: 26, color: C.white, bold: true,
  });

  const leftMetrics = [
    { val: '↓ 75%', lbl: 'Stockout frequency', sub: 'From 12+ SKUs/day to under 3' },
    { val: '↓ 40%', lbl: 'Dead stock tied-up capital', sub: 'ML flags slow-movers 8 weeks early' },
    { val: '−60%', lbl: 'Manual buyer review time', sub: 'Approve AI drafts vs build from scratch' },
    { val: '91%', lbl: 'Forecast accuracy at 30 days', sub: 'vs ~55% with current rule-based approach' },
  ];
  const rightMetrics = [
    { val: 'AED 2.1M', lbl: 'KSA seed PO confidence', sub: 'AI-calibrated from UAE launch analogy' },
    { val: '↑ 18%', lbl: 'Gross margin on replenishment', sub: 'Right quantity = less discount to clear' },
    { val: '< 10 wks', lbl: 'Time to first ROI', sub: 'Stockout reduction visible within weeks' },
    { val: 'AED 4.8M+', lbl: 'Annual inventory efficiency gain', sub: 'Stockouts averted + dead stock reduction' },
  ];

  s.addText('INVENTORY EFFICIENCY', { x: 0.55, y: 1.4, w: 4.2, h: 0.22, fontSize: 8.5, color: C.purple, bold: true, charSpacing: 2 });
  leftMetrics.forEach((m, i) => {
    s.addShape(prs.ShapeType.rect, { x: 0.55, y: 1.68 + i * 0.88, w: 4.2, h: 0.78, fill: { color: C.card }, line: { color: '1e1635', width: 0.6 } });
    s.addText(m.val, { x: 0.7, y: 1.73 + i * 0.88, w: 1.6, h: 0.68, fontSize: 24, color: C.purple, bold: true, valign: 'middle' });
    s.addText(m.lbl + '\n' + m.sub, { x: 2.35, y: 1.73 + i * 0.88, w: 2.25, h: 0.68, fontSize: 9.5, color: C.muted, lineSpacingMultiple: 1.35, valign: 'middle' });
  });

  s.addText('BUSINESS ROI', { x: 5.25, y: 1.4, w: 4.2, h: 0.22, fontSize: 8.5, color: C.amber, bold: true, charSpacing: 2 });
  rightMetrics.forEach((m, i) => {
    s.addShape(prs.ShapeType.rect, { x: 5.25, y: 1.68 + i * 0.88, w: 4.2, h: 0.78, fill: { color: C.card }, line: { color: '1e1635', width: 0.6 } });
    s.addText(m.val, { x: 5.4, y: 1.73 + i * 0.88, w: 1.8, h: 0.68, fontSize: 22, color: C.amber, bold: true, valign: 'middle' });
    s.addText(m.lbl + '\n' + m.sub, { x: 7.25, y: 1.73 + i * 0.88, w: 2.05, h: 0.68, fontSize: 9.5, color: C.muted, lineSpacingMultiple: 1.35, valign: 'middle' });
  });

  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 5.28, w: 8.9, h: 0.72, fill: { color: '050312' }, line: { color: C.purple, width: 0.8 } });
  s.addText('At PhonePe: replacing static rules with ML-driven decisions cut marketing burn by 32% while sustaining GMV growth — the same capital efficiency gain InventoryIQ delivers on the supply side.', {
    x: 0.75, y: 5.38, w: 8.5, h: 0.52, fontSize: 10.5, color: C.purple, lineSpacingMultiple: 1.4,
  });
}

// ─── SLIDE 7: PROOF OF WORK ───────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.white };

  s.addText('PROOF OF WORK', { x: 0.55, y: 0.35, w: 9, h: 0.22, fontSize: 9, color: C.purple, bold: true, charSpacing: 3 });
  s.addText('I Replaced Static Rules with ML at Scale. Here Is What I Shipped.', {
    x: 0.55, y: 0.62, w: 9, h: 0.55, fontSize: 26, color: '0f172a', bold: true,
  });

  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 1.38, w: 4.3, h: 3.9, fill: { color: C.dark }, line: { color: '1e1635', width: 0.8 } });
  s.addText('WHAT I BUILT AT PHONEPE', { x: 0.75, y: 1.55, w: 3.9, h: 0.22, fontSize: 8.5, color: C.purple, bold: true, charSpacing: 2 });
  [
    ['Propensity-to-Transact ML model', 'Real-time per-user scoring replacing static rule-based cohorts. 32% marketing burn reduction on a INR 1000+ Cr/yr budget — same precision-vs-rules principle.'],
    ['ML-driven brand partner marketplace', '500+ brand partners on a self-serve system. Audience segmentation replaced fixed targeting. 11% YoY revenue growth, 26% engagement increase.'],
    ['Dynamic cart incentivisation engine', 'Context-aware triggers (cart value x user intent x time) replacing flat rules. 35% AOV uplift and 20% improvement in 30-day retention.'],
    ['Antigravity Resume Agent (side project)', 'Multi-agent ML pipeline auto-generating ATS-optimised outputs from input signals — same feature engineering mindset as demand forecasting.'],
  ].forEach(([title, desc], i) => {
    s.addText('■  ' + title, { x: 0.75, y: 1.88 + i * 0.84, w: 3.9, h: 0.24, fontSize: 10.5, color: C.white, bold: true });
    s.addText(desc, { x: 0.95, y: 2.14 + i * 0.84, w: 3.7, h: 0.52, fontSize: 9.5, color: C.muted, lineSpacingMultiple: 1.35 });
  });

  s.addShape(prs.ShapeType.rect, { x: 5.15, y: 1.38, w: 4.3, h: 3.9, fill: { color: 'faf5ff' }, line: { color: 'c4b5fd', width: 0.8 } });
  s.addText('HOW IT MAPS TO EYEWA', { x: 5.35, y: 1.55, w: 3.9, h: 0.22, fontSize: 8.5, color: '5b21b6', bold: true, charSpacing: 2 });
  [
    ['AI and automation in ops', 'InventoryIQ is ML applied to a high-volume, high-variance ops decision — the exact use case eyewa is hiring for.'],
    ['Data-driven product decisions', 'Model selection, feature engineering, confidence thresholds, and A/B design — all owned end-to-end at PhonePe.'],
    ['Strong analytical mindset', 'Translating ambiguous demand signals into clear purchase decisions is the core of this product and my background.'],
    ['Stakeholder management', 'Aligned buyers, finance, engineering, and data science across competing inventory and margin priorities.'],
  ].forEach(([title, desc], i) => {
    s.addText('✓  ' + title, { x: 5.35, y: 1.88 + i * 0.84, w: 3.9, h: 0.24, fontSize: 10.5, color: '5b21b6', bold: true });
    s.addText(desc, { x: 5.55, y: 2.14 + i * 0.84, w: 3.7, h: 0.52, fontSize: 9.5, color: '475569', lineSpacingMultiple: 1.35 });
  });

  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 5.42, w: 8.9, h: 0.72, fill: { color: 'faf5ff' }, line: { color: C.purple, width: 0.8 } });
  s.addText('"Replacing static rules with ML is not a technical challenge — it is a product clarity challenge. You need to know exactly what signal matters, what decision it drives, and what the cost of a wrong prediction is. I have done this at scale."', {
    x: 0.75, y: 5.52, w: 8.5, h: 0.52, fontSize: 10, color: '5b21b6', italic: true, lineSpacingMultiple: 1.35,
  });
}

// ─── SLIDE 8: ROLLOUT ─────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('ROLLOUT PLAN', { x: 0.55, y: 0.35, w: 9, h: 0.22, fontSize: 9, color: C.purple, bold: true, charSpacing: 3 });
  s.addText('Three Phases — First Forecast Accuracy Measurement in 8 Weeks', {
    x: 0.55, y: 0.62, w: 9, h: 0.55, fontSize: 26, color: C.white, bold: true,
  });

  const phases = [
    { phase: 'Phase 1', period: 'Weeks 1–8', title: 'Forecast Foundation', color: C.purple, items: ['Connect Shopify + ERP data feeds', 'Train baseline ML model on 24-month sales history', 'Run forecast in shadow mode — compare vs actuals', 'Measure accuracy: target 85%+ at 30-day horizon'] },
    { phase: 'Phase 2', period: 'Weeks 9–16', title: 'Smart PO Generator', color: C.amber, items: ['Launch AI-drafted PO queue for top 500 SKUs', 'Buyer approves AI recommendations vs building manually', 'A/B measure: AI-drafted vs buyer-calculated stockout rate', 'Expand to full 6,066 SKU catalogue after bake-in'] },
    { phase: 'Phase 3', period: 'Weeks 17–24', title: 'New Market Playbook', color: C.green, items: ['Deploy analogy-market model for KSA Riyadh launch', 'Calibrate seed inventory mix from UAE launch curve data', 'Build market-specific seasonality and event calendars', 'Expand playbook to Bahrain, Kuwait on next launch cycle'] },
  ];

  phases.forEach((ph, i) => {
    const x = 0.45 + i * 3.18;
    s.addShape(prs.ShapeType.rect, { x, y: 1.42, w: 2.95, h: 3.5, fill: { color: C.card }, line: { color: ph.color, width: 1.2 }, shadow: makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y: 1.42, w: 2.95, h: 0.36, fill: { color: ph.color, transparency: 88 }, line: { color: ph.color, width: 0.6 } });
    s.addText(ph.phase + '  ·  ' + ph.period, { x, y: 1.42, w: 2.95, h: 0.36, fontSize: 9, color: ph.color, bold: true, align: 'center', valign: 'middle' });
    s.addText(ph.title, { x: x + 0.12, y: 1.88, w: 2.7, h: 0.38, fontSize: 13.5, color: C.white, bold: true });
    ph.items.forEach((item, j) => {
      s.addText('→  ' + item, { x: x + 0.12, y: 2.34 + j * 0.56, w: 2.7, h: 0.48, fontSize: 9.8, color: C.muted, lineSpacingMultiple: 1.3 });
    });
  });

  s.addShape(prs.ShapeType.rect, { x: 0.55, y: 5.08, w: 9.0, h: 0.9, fill: { color: C.card }, line: { color: C.purple, width: 0.8 } });
  s.addText('WHAT I NEED TO BUILD THIS', { x: 0.75, y: 5.18, w: 3.5, h: 0.2, fontSize: 8.5, color: C.purple, bold: true, charSpacing: 2 });
  s.addText('24-month sales history export from ERP  ·  Shopify + warehouse data access  ·  1 data engineer  ·  Buyer team as pilot users (Phase 2)  ·  Supplier lead time database', {
    x: 0.75, y: 5.4, w: 8.6, h: 0.45, fontSize: 10.5, color: C.muted, lineSpacingMultiple: 1.35,
  });

  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade', {
    x: 0.55, y: 6.9, w: 9.0, h: 0.25, fontSize: 9, color: '475569', align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/eyewa-inventoryiq-deck.pptx' })
  .then(() => console.log('✅ eyewa-inventoryiq-deck.pptx written'))
  .catch(e => console.error(e));
