const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '1A0A14',
  hero:   '2D1220',
  pink:   'FC2779',
  rose:   'FF6BAD',
  white:  'FFFFFF',
  lgray:  'FFF5F8',
  ink:    '1A0A14',
  muted:  '8B5070',
};
const makeShadow = () => ({ type: 'outer', color: '000000', blur: 3, offset: 1, angle: 45, opacity: 0.12 });

// ─────────────────────────────────────────
// SLIDE 1 — COVER
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  // Diagonal accent lines
  for (let i = 0; i < 8; i++) {
    sl.addShape(prs.ShapeType.line, { x: -0.5 + i * 1.5, y: 0, w: 4, h: 7.5, line: { color: 'FC2779', width: 0.4, transparency: 88 } });
  }

  // Top tag
  sl.addText('NYKAA SUPERSTORE — B2B PRODUCT', { x: 0.5, y: 0.32, w: 9, h: 0.28, fontSize: 9, color: C.rose, bold: false, charSpacing: 2 });

  // Product name
  sl.addText('QuickStore', { x: 0.5, y: 0.75, w: 9, h: 1.4, fontSize: 64, color: C.white, bold: true });

  // Subtitle
  sl.addText('Self-Serve Guided Ordering for Non-Tech Bharat Retailers', {
    x: 0.5, y: 2.0, w: 7.5, h: 0.6, fontSize: 22, color: C.rose, bold: false,
  });

  // Divider
  sl.addShape(prs.ShapeType.rect, { x: 0.5, y: 2.75, w: 2.8, h: 0.04, fill: { color: C.pink } });

  // Presenter
  sl.addText('Presented by Ajay Avaghade · Growth & Monetization PM', {
    x: 0.5, y: 2.95, w: 7, h: 0.35, fontSize: 13, color: 'CCAABB',
  });

  // Big stat bottom-right
  sl.addShape(prs.ShapeType.rect, { x: 7.2, y: 5.2, w: 2.5, h: 1.9, fill: { color: C.pink }, shadow: makeShadow() });
  sl.addText('70%', { x: 7.2, y: 5.3, w: 2.5, h: 0.85, fontSize: 52, color: C.white, bold: true, align: 'center' });
  sl.addText('of B2B orders still\nrequire BDE calls', { x: 7.2, y: 6.1, w: 2.5, h: 0.8, fontSize: 11, color: 'FFD0E8', align: 'center' });
}

// ─────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE PROBLEM', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.rose, bold: true, charSpacing: 2 });
  sl.addText('Bharat Retailers Can\'t Self-Serve — And Nykaa Is Paying for Every Order with BDE Time', {
    x: 0.5, y: 0.62, w: 9, h: 0.9, fontSize: 22, color: C.white, bold: true,
  });

  const cols = [
    { icon: '📱', stat: '70%', lbl: 'Orders via BDE calls', sub: 'Non-tech retailers can\'t navigate a 10K+ SKU catalog without hand-holding' },
    { icon: '⏱️', stat: '48 hrs', lbl: 'Average order cycle time', sub: 'WhatsApp → BDE → quote → confirm → payment: a 4-step relay before first unit ships' },
    { icon: '📉', stat: '40%', lbl: 'Reorder drop-off', sub: 'Retailers who ordered once don\'t reorder — no smart reminder, no frictionless refill' },
  ];
  cols.forEach((c, i) => {
    const x = 0.45 + i * 3.12;
    sl.addShape(prs.ShapeType.rect, { x, y: 1.72, w: 2.9, h: 3.1, fill: { color: C.hero }, shadow: makeShadow() });
    sl.addText(c.icon, { x, y: 1.88, w: 2.9, h: 0.55, fontSize: 26, align: 'center' });
    sl.addText(c.stat, { x, y: 2.52, w: 2.9, h: 0.72, fontSize: 40, color: C.pink, bold: true, align: 'center' });
    sl.addText(c.lbl, { x, y: 3.22, w: 2.9, h: 0.38, fontSize: 12.5, color: C.white, bold: true, align: 'center' });
    sl.addText(c.sub, { x: x + 0.1, y: 3.62, w: 2.7, h: 0.9, fontSize: 10, color: 'CCAABB', align: 'center' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.0, w: 9.1, h: 0.82, fill: { color: C.pink } });
  sl.addText('Root cause: Non-tech Bharat retailers — salons, clinics, boutiques — are systematically excluded from self-serve because the catalog is too large, jargon-heavy, and doesn\'t know what type of business they run.', {
    x: 0.62, y: 5.06, w: 8.76, h: 0.72, fontSize: 10.5, color: C.white,
  });
}

// ─────────────────────────────────────────
// SLIDE 3 — THE INSIGHT
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE INSIGHT', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.pink, bold: true, charSpacing: 2 });
  sl.addText('Retailers Don\'t Need a Search Bar. They Need a Shelf Curated for Their Business Type.', {
    x: 0.5, y: 0.62, w: 9, h: 0.9, fontSize: 22, color: C.ink, bold: true,
  });

  // Left box — current
  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 1.72, w: 4.1, h: 3.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('❌  Current Approach', { x: 0.62, y: 1.85, w: 3.76, h: 0.38, fontSize: 13, color: '444444', bold: true });
  const left = ['Open catalog — 10,000+ SKUs', 'No store-type filter or context', 'No peer-group data or social proof', 'Retailer calls BDE for recommendation', 'Order eventually placed after 2+ days'];
  left.forEach((t, i) => sl.addText(`• ${t}`, { x: 0.72, y: 2.35 + i * 0.42, w: 3.56, h: 0.36, fontSize: 11, color: '555555' }));

  // VS divider
  sl.addShape(prs.ShapeType.ellipse, { x: 4.62, y: 2.9, w: 0.72, h: 0.72, fill: { color: C.pink } });
  sl.addText('VS', { x: 4.62, y: 2.9, w: 0.72, h: 0.72, fontSize: 11, color: C.white, bold: true, align: 'center', valign: 'middle' });

  // Right box — proposed
  sl.addShape(prs.ShapeType.rect, { x: 5.42, y: 1.72, w: 4.1, h: 3.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('✅  QuickStore', { x: 5.58, y: 1.85, w: 3.76, h: 0.38, fontSize: 13, color: C.pink, bold: true });
  const right = ['Store-type wizard — 6 categories, 30 sec', 'Peer-group smart shelf: "840 salons ordered this"', 'Reorder cart pre-filled from last purchase', 'One-tap repeat for regular SKUs', 'BDE freed for acquisition, not hand-holding'];
  right.forEach((t, i) => sl.addText(`• ${t}`, { x: 5.58, y: 2.35 + i * 0.42, w: 3.72, h: 0.36, fontSize: 11, color: '333333' }));

  // Bottom callout
  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.08, w: 9.1, h: 0.72, fill: { color: C.pink } });
  sl.addText('Key insight: Nykaa already has the peer-group data — which SKUs salons vs. clinics vs. boutiques in each city actually buy. QuickStore surfaces this social proof at the moment of decision, making the catalog feel curated rather than overwhelming.', {
    x: 0.62, y: 5.12, w: 8.76, h: 0.62, fontSize: 10, color: C.white,
  });
}

// ─────────────────────────────────────────
// SLIDE 4 — THE MECHANIC
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE MECHANIC', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.rose, bold: true, charSpacing: 2 });
  sl.addText('Three Layers: Know the Retailer → Show the Right Shelf → Make Reorder Effortless', {
    x: 0.5, y: 0.62, w: 9, h: 0.75, fontSize: 22, color: C.white, bold: true,
  });

  const steps = [
    { n: '01', title: 'Store-Type Onboarding', body: 'Retailer selects store type (Salon / Clinic / Boutique / Pharmacy / Parlour / General) + city in under 30 seconds. No search, no scroll.' },
    { n: '02', title: 'Peer-Group Smart Shelf', body: '"840 salons in Jaipur ordered these this week" — top 12 SKUs for your store type, sorted by popularity among peers. Social proof replaces BDE recommendation.' },
    { n: '03', title: 'Pre-Filled Reorder Cart', body: 'From the 2nd visit, cart is pre-filled with last order quantities. Retailer adjusts 1–2 items and checks out in 90 seconds.' },
    { n: '04', title: 'Smart Reorder Nudge', body: 'ML model predicts stockout date based on order frequency. Push notification 3 days before: "Your Wella Color stock runs out in 3 days — reorder now."' },
    { n: '05', title: 'BDE Graduation Dashboard', body: 'Ops screen tracks self-serve rate by retailer. At 3 consecutive self-serve orders, retailer "graduates" — BDE effort reallocated to new acquisition.' },
  ];

  // Dashed connecting line
  sl.addShape(prs.ShapeType.line, { x: 0.88, y: 1.82, w: 0, h: 4.2, line: { color: C.pink, width: 1.5, dashType: 'dash' } });

  steps.forEach((s, i) => {
    const y = 1.65 + i * 0.86;
    sl.addShape(prs.ShapeType.ellipse, { x: 0.6, y: y, w: 0.55, h: 0.55, fill: { color: C.pink } });
    sl.addText(s.n, { x: 0.6, y: y, w: 0.55, h: 0.55, fontSize: 10, color: C.white, bold: true, align: 'center', valign: 'middle' });
    sl.addText(s.title, { x: 1.3, y: y, w: 2.6, h: 0.3, fontSize: 11.5, color: C.pink, bold: true });
    sl.addText(s.body, { x: 1.3, y: y + 0.3, w: 8.1, h: 0.44, fontSize: 9.5, color: 'CCAABB' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 6.18, w: 9.1, h: 0.62, fill: { color: C.hero } });
  sl.addText('PhonePe Proof: Built the same BDE graduation construct at PhonePe — "merchant self-serve rate" tracked weekly; BDEs re-deployed when merchants hit 3 consecutive self-serve transactions → 23% lower CAC.', {
    x: 0.62, y: 6.22, w: 8.76, h: 0.52, fontSize: 9.5, color: C.rose,
  });
}

// ─────────────────────────────────────────
// SLIDE 5 — THE PRODUCT
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE PRODUCT', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.pink, bold: true, charSpacing: 2 });
  sl.addText('4 Key Screen States — Built & Tested at PhonePe Scale', {
    x: 0.5, y: 0.62, w: 9, h: 0.55, fontSize: 22, color: C.ink, bold: true,
  });

  const screens = [
    { n: '01', title: 'Store Wizard', desc: 'Store-type picker (6 types) + city selection. Peer insight banner loads before catalog opens. 30-second context capture.' },
    { n: '02', title: 'Smart Shelf', desc: '"840 salons ordered this" social proof badge on every SKU. Relevance ranked. B2B pricing visible upfront.' },
    { n: '03', title: 'Reorder Cart', desc: 'Pre-filled from last order. Edit quantities inline. B2B discount applied. One-tap checkout. 90-second reorder.' },
    { n: '04', title: 'BDE Dashboard', desc: 'Ops screen: self-serve rate, "at risk" retailers, graduation progress. BDE effort directed by data, not guesswork.' },
  ];

  screens.forEach((s, i) => {
    const x = 0.45 + i * 2.38;
    sl.addShape(prs.ShapeType.rect, { x, y: 1.38, w: 2.18, h: 3.5, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(prs.ShapeType.rect, { x, y: 1.38, w: 2.18, h: 0.32, fill: { color: C.pink } });
    sl.addText(`${s.n} — ${s.title}`, { x, y: 1.38, w: 2.18, h: 0.32, fontSize: 9.5, color: C.white, bold: true, align: 'center', valign: 'middle' });

    // ASCII mockup
    const mockup = i === 0
      ? '[ ✂ Salon  ] [ 💊 Clinic ]\n[ 👗 Boutique ] [ 💅 Parlour ]\n─────────────────\n📍 City: Jaipur\n[ Continue → ]'
      : i === 1
      ? '🔥 840 salons ordered\n┌─────────────────┐\n│ Wella Color Care │\n│ ⭐ 4.7  B2B ₹280 │\n└─────────────────┘\n[ Add to Cart ]'
      : i === 2
      ? 'Last order: 22 days ago\n─────────────────\nWella Color × 3 ₹840\nStreax Pro × 2  ₹560\nOPI Top Coat × 1₹380\n─────────────────\nTotal: ₹5,032\n[ Place Order ]'
      : 'Self-Serve Rate: 62%\n█████████░ 18 graduated\n─────────────────\n⚠️ Priya Salon – At risk\n🔵 Meena Boutique – Good\n✅ Rekha Clinic – Done';

    sl.addText(mockup, { x: x + 0.08, y: 1.82, w: 2.02, h: 2.6, fontSize: 7.5, color: '333333', fontFace: 'Courier New' });
    sl.addText(s.desc, { x: x + 0.06, y: 4.94, w: 2.06, h: 0.8, fontSize: 9, color: '555555' });
  });

  sl.addText('Live prototype: nykaa-quickstore-prototype.html', { x: 0.45, y: 6.88, w: 9.1, h: 0.28, fontSize: 10, color: C.pink, align: 'center' });
}

// ─────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('IMPACT & ROI', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.rose, bold: true, charSpacing: 2 });
  sl.addText('Projected Impact — Anchored in PhonePe Merchant Self-Serve Proof Points', {
    x: 0.5, y: 0.62, w: 9, h: 0.7, fontSize: 22, color: C.white, bold: true,
  });

  sl.addText('RETAILER IMPACT', { x: 0.45, y: 1.5, w: 4.5, h: 0.3, fontSize: 9.5, color: C.pink, bold: true, charSpacing: 1 });
  sl.addText('NYKAA PLATFORM ROI', { x: 5.05, y: 1.5, w: 4.5, h: 0.3, fontSize: 9.5, color: C.pink, bold: true, charSpacing: 1 });

  const left = [
    { v: '↓ 60%', l: 'BDE calls per order' },
    { v: '90 sec', l: 'Reorder time (vs 48 hrs)' },
    { v: '↑ 2.4×', l: 'Reorder frequency' },
    { v: '↑ 35%', l: 'Basket size via smart shelf' },
  ];
  const right = [
    { v: '↓ 23%', l: 'BDE cost per merchant' },
    { v: '↑ 40%', l: 'GMV per active retailer' },
    { v: '62%', l: 'Self-serve target at 90 days' },
    { v: '₹0', l: 'Incremental tech cost for reorder' },
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
  sl.addText('BDE reallocation alone funds the QuickStore build: if 60% of BDE calls are eliminated, the equivalent headcount can be redeployed to net-new acquisition — the feature pays for itself in the first quarter.', {
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
  sl.addText('I Built This. Here\'s the Proof.', { x: 0.5, y: 0.62, w: 9, h: 0.68, fontSize: 26, color: C.ink, bold: true });

  // Left dark panel
  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 1.42, w: 4.2, h: 4.08, fill: { color: C.dark }, shadow: makeShadow() });
  sl.addText('What I Built at PhonePe', { x: 0.62, y: 1.55, w: 3.86, h: 0.35, fontSize: 12.5, color: C.pink, bold: true });
  const proof = [
    'Rebuilt merchant onboarding from 2-day manual TAT to 30-min self-serve PG integration',
    'Designed multi-tenant Instant Discount & EMI subvention platform — zero-to-one',
    'Built BDE graduation tracking — self-serve rate dashboard used by 15-person ops team',
    'Acquired 5,000+ B2B merchants — 23% lower CAC vs BDE-led acquisition',
    'Reduced offer ops TAT from 2 days to 30 minutes via self-serve config workflow',
  ];
  proof.forEach((t, i) => sl.addText(`• ${t}`, { x: 0.72, y: 2.0 + i * 0.65, w: 3.76, h: 0.56, fontSize: 10.5, color: C.white }));

  // Right light panel
  sl.addShape(prs.ShapeType.rect, { x: 4.82, y: 1.42, w: 4.7, h: 4.08, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('How It Maps to This JD', { x: 4.98, y: 1.55, w: 4.38, h: 0.35, fontSize: 12.5, color: C.pink, bold: true });
  const maps = [
    ['Self-serve PG integration', '→ QuickStore self-serve ordering flow'],
    ['30-min offer ops TAT', '→ 90-sec reorder experience design'],
    ['BDE graduation tracking', '→ Retailer self-serve rate dashboard'],
    ['Merchant growth platform', '→ Peer-group smart shelf & nudge engine'],
    ['B2B go-to-market execution', '→ QuickStore phased retailer rollout'],
  ];
  maps.forEach(([f, m], i) => {
    sl.addText(f, { x: 4.98, y: 2.0 + i * 0.65, w: 2.2, h: 0.56, fontSize: 10, color: '444444', bold: true });
    sl.addText(m, { x: 7.2, y: 2.0 + i * 0.65, w: 2.2, h: 0.56, fontSize: 10, color: '333333' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.62, w: 9.1, h: 0.65, fill: { color: C.pink } });
  sl.addText('"QuickStore isn\'t a pitch exercise. It\'s the same self-serve + graduation architecture I shipped at PhonePe — adapted for Nykaa\'s Bharat retailer problem."', {
    x: 0.62, y: 5.66, w: 8.76, h: 0.56, fontSize: 10.5, color: C.white, italic: true,
  });
}

// ─────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('ROLLOUT PLAN', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.rose, bold: true, charSpacing: 2 });
  sl.addText('Phased Rollout — Pilot → Validate → Scale', {
    x: 0.5, y: 0.62, w: 9, h: 0.65, fontSize: 24, color: C.white, bold: true,
  });

  const phases = [
    {
      phase: 'Phase 1', time: 'Month 1–2: Pilot',
      body: 'Deploy QuickStore to 200 non-tech retailers across 3 cities (Jaipur, Lucknow, Nagpur). Measure self-serve rate, reorder frequency, BDE call reduction vs. control group.',
    },
    {
      phase: 'Phase 2', time: 'Month 3–4: Smart Shelf',
      body: 'Activate peer-group shelf using Nykaa purchase data. A/B test: peer-ranked shelf vs. catalog-ranked shelf. Instrument basket size, session time, add-to-cart rate.',
    },
    {
      phase: 'Phase 3', time: 'Month 5–6: Nationwide',
      body: 'Roll winning variant to all new B2B onboarding. Integrate reorder nudge (ML stockout predictor). Launch BDE graduation dashboard to all city ops teams.',
    },
  ];

  phases.forEach((p, i) => {
    const x = 0.45 + i * 3.12;
    sl.addShape(prs.ShapeType.rect, { x, y: 1.5, w: 2.92, h: 3.4, fill: { color: C.hero }, shadow: makeShadow() });
    sl.addShape(prs.ShapeType.rect, { x, y: 1.5, w: 2.92, h: 0.52, fill: { color: C.pink } });
    sl.addText(p.phase, { x, y: 1.5, w: 2.92, h: 0.52, fontSize: 15, color: C.white, bold: true, align: 'center', valign: 'middle' });
    sl.addText(p.time, { x: x + 0.1, y: 2.12, w: 2.72, h: 0.35, fontSize: 11, color: C.rose, bold: true });
    sl.addText(p.body, { x: x + 0.1, y: 2.5, w: 2.72, h: 2.2, fontSize: 10, color: 'CCAABB' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.06, w: 9.1, h: 0.72, fill: { color: C.hero } });
  sl.addText('What I need to build this: City-level purchase data by store type · 1 engineer for reorder cart API · 1 ML engineer for stockout predictor · Ops alignment on BDE graduation SOP', {
    x: 0.62, y: 5.1, w: 8.76, h: 0.62, fontSize: 10.5, color: C.rose,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.9, w: 9.1, h: 0.62, fill: { color: C.pink } });
  sl.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  ajay-avaghade.github.io/Portfolio', {
    x: 0.62, y: 5.94, w: 8.76, h: 0.52, fontSize: 10.5, color: C.white, align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/nykaa-quickstore-deck.pptx' })
  .then(() => console.log('✅ nykaa-quickstore-deck.pptx written'))
  .catch(e => { console.error(e); process.exit(1); });
