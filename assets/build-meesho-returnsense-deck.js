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

  sl.addText('MEESHO GROWTH & LOSS MITIGATION STRATEGY', { x: 0.5, y: 0.32, w: 9, h: 0.28, fontSize: 9, color: C.pink, charSpacing: 2 });
  sl.addText('ReturnSense', { x: 0.5, y: 0.75, w: 9, h: 1.4, fontSize: 64, color: C.white, bold: true });
  sl.addText('Reducing Returns and RTO via Sizing Intelligence and Checkout Verification', {
    x: 0.5, y: 2.0, w: 7.5, h: 0.6, fontSize: 22, color: C.accent, bold: false,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.5, y: 2.75, w: 2.8, h: 0.04, fill: { color: C.pink } });

  sl.addText('Presented by Ajay Avaghade · Product Manager 2 Candidate', {
    x: 0.5, y: 2.95, w: 7, h: 0.35, fontSize: 13, color: 'FFD5E5',
  });

  // Big stat
  sl.addShape(prs.ShapeType.rect, { x: 7.2, y: 5.2, w: 2.5, h: 1.9, fill: { color: C.pink }, shadow: makeShadow() });
  sl.addText('30%', { x: 7.2, y: 5.3, w: 2.5, h: 0.85, fontSize: 52, color: C.white, bold: true, align: 'center' });
  sl.addText('average return rate in\nIndian fashion e-commerce', { x: 7.2, y: 6.1, w: 2.5, h: 0.8, fontSize: 11, color: 'FFE6F0', align: 'center' });
}

// ─────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE PROBLEM', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.pink, bold: true, charSpacing: 2 });
  sl.addText('Returns and RTO Drain E-Commerce Margins. COD Purchases in Tier 2/3/4 Towns Suffer Most.', {
    x: 0.5, y: 0.62, w: 9, h: 0.9, fontSize: 21, color: C.white, bold: true,
  });

  const cols = [
    { icon: '🧥', stat: '40-60%', lbl: 'Size-related returns', sub: 'Inconsistent supplier size charts cause major fit issues. Users default to guessing sizes.' },
    { icon: '🚚', stat: '25-35%', lbl: 'RTO on COD orders', sub: 'Vague address data and lack of buyer intent at delivery lead to high rejection rates.' },
    { icon: '💸', stat: '₹120+', lbl: 'Cost per returned unit', sub: 'Forward and reverse logistics costs quickly wipe out the low-margin benefits of Tier 2+ e-commerce.' },
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
  sl.addText('Root cause: Traditional e-commerce relies on users reading complex standard charts in English. For the next 500M users, sizing needs visual, vernacular, and crowdsourced references, backed by checkout validation of delivery intent.', {
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
  sl.addText('Prevent Returns by Bridging Sizing Gaps & Adding Friction to High-Risk COD Checkouts.', {
    x: 0.5, y: 0.62, w: 9, h: 0.9, fontSize: 22, color: C.dark, bold: true,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 1.72, w: 4.1, h: 3.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('❌  Current Checkout Flow', { x: 0.62, y: 1.85, w: 3.76, h: 0.38, fontSize: 12, color: '444444', bold: true });
  const left = [
    'Static English size charts with zero user context',
    'No sizing warnings even if product is frequently returned',
    'Frictionless COD allows impulse orders with low intent',
    'High RTO areas treated same as low RTO areas at checkout',
    'Suppliers receive return data without suggestions on fixes',
  ];
  left.forEach((t, i) => sl.addText(`• ${t}`, { x: 0.72, y: 2.35 + i * 0.42, w: 3.56, h: 0.36, fontSize: 11, color: '555555' }));

  sl.addShape(prs.ShapeType.ellipse, { x: 4.62, y: 2.9, w: 0.72, h: 0.72, fill: { color: C.pink } });
  sl.addText('VS', { x: 4.62, y: 2.9, w: 0.72, h: 0.72, fontSize: 11, color: C.white, bold: true, align: 'center', valign: 'middle' });

  sl.addShape(prs.ShapeType.rect, { x: 5.42, y: 1.72, w: 4.1, h: 3.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('✅  ReturnSense (Proposed)', { x: 5.58, y: 1.85, w: 3.76, h: 0.38, fontSize: 12, color: C.pink, bold: true });
  const right = [
    'AI Size Intelligence surfaces personalized sizing matches',
    'Return Risk warning badges call out sizing discrepancies',
    'Vernacular Size Guides show real photos of similar buyers',
    'COD verification checks high-risk orders with WhatsApp bots',
    'Actionable supplier alerts highlight specific size chart edits',
  ];
  right.forEach((t, i) => sl.addText(`• ${t}`, { x: 5.58, y: 2.35 + i * 0.42, w: 3.72, h: 0.36, fontSize: 10.5, color: '333333' }));

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.08, w: 9.1, h: 0.72, fill: { color: C.pink } });
  sl.addText('Key insight: Sizing returns are catalog errors, and RTO is an intent problem. Bridging size confidence through visual crowdsourcing and validating delivery intent using automated voice calls solves both at the root.', {
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
  sl.addText('ReturnSense Stack: Surf, Show, Validate, and Action', {
    x: 0.5, y: 0.62, w: 9, h: 0.65, fontSize: 22, color: C.white, bold: true,
  });

  const steps = [
    { n: '01', title: 'PDP Size Surfacing', body: 'ML engine cross-references buyer\'s purchase history to recommend best size (e.g. Medium). Surfaced via Return Risk warning badge on high-return SKUs.' },
    { n: '02', title: 'Vernacular Visual Guide', body: 'Buyers see real photo-fit comparisons of other buyers with similar height/weight. High-density Hindi/regional labels eliminate guesswork.' },
    { n: '03', title: 'Risk Checkout Intercept', body: 'COD orders from high-risk pin codes or return profiles trigger bottom-sheet warning. Encourages prepaid via ₹15 discount or forces voice validation.' },
    { n: '04', title: 'WhatsApp Voice Validation', body: 'Uses GenAI voice assistant (Vaani) to call and confirm: "Your order is COD ₹482. Confirm by saying Yes." Removes low-intent orders.' },
    { n: '05', title: 'Supplier Action Center', body: 'Supplier dashboard identifies SKUs causing high returns and outlines edits: "Kurta sleeves run tight. Update size chart by 1.5 inches to save ₹42K."' },
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
  sl.addText('PhonePe Parallel: Surfacing eligibility and checks at transaction points matches theoffers/redemption stack. ML propensity-based intercept mirrors dynamic checkout discount suppression.', {
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
  sl.addText('Interactive Prototypes — Size Intelligence to Dynamic Intercepts', {
    x: 0.5, y: 0.62, w: 9, h: 0.55, fontSize: 22, color: C.dark, bold: true,
  });

  const screens = [
    { n: '01', title: 'PDP Sizing Alert', desc: 'Displays Return Risk warning. Custom Size Confidence Meter calculates match rating (84%) using historical data.' },
    { n: '02', title: 'Visual Photo Guide', desc: 'Crowdsourced photo guide in Hindi shows how Medium fits people of similar heights/weights. Self-selected size verification.' },
    { n: '03', title: 'COD Intercept', desc: 'RTO intercept sheet offers online prepay discount (₹15 off) or WhatsApp voice confirmation. Converts or cancels early.' },
    { n: '04', title: 'Supplier Dashboard', desc: 'Alerts supplier about sizing defects on specific SKUs and provides 1-click catalog modifications.' },
  ];

  screens.forEach((s, i) => {
    const x = 0.45 + i * 2.38;
    sl.addShape(prs.ShapeType.rect, { x, y: 1.38, w: 2.18, h: 3.5, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(prs.ShapeType.rect, { x, y: 1.38, w: 2.18, h: 0.32, fill: { color: C.pink } });
    sl.addText(`${s.n} — ${s.title}`, { x, y: 1.38, w: 2.18, h: 0.32, fontSize: 9.5, color: C.white, bold: true, align: 'center', valign: 'middle' });

    const mockup = i === 0
      ? '⚠️ Sizing Alert: Returns\n─────────────────\n18% returned for tight chest\n\n✦ AI Intelligence\nSuggested Size: M (84% Match)\n[████████░░░] 84%\n\nSelect Size:\n[ S ] [ M* ] [ L ] [ XL ]'
      : i === 1
      ? '👩 फोटो गाइड (Medium)\n─────────────────\nकद: 5\'3" | वजन: 56kg\nसंतुष्टि: परफेक्ट फिट (सही माप)\n\n👩‍🦰 कद: 5\'5" | वजन: 59kg\nसंतुष्टि: छाती में टाइट\n\n[ Select M & Proceed ]'
      : i === 2
      ? '⚠️ Important Verification\n─────────────────\nYour pin code has high RTO\n\n🎁 Switch to Online:\nSave ₹15 instant discount!\n\n[ Switch to Online & Save ]\n[ Verify via WhatsApp Call ]'
      : '📊 Supplier Hub\n─────────────────\nReturn + RTO: 21.4% (Critical)\n\nAlert SKU: ANARKALI-SILK-M\n• Reason: Chest runs small\n• AI Fix: Adjust chart chest\n  from 38" to 39.5"';

    sl.addText(mockup, { x: x + 0.08, y: 1.82, w: 2.02, h: 2.6, fontSize: 7.5, color: '333333', fontFace: 'Courier New' });
    sl.addText(s.desc, { x: x + 0.06, y: 4.94, w: 2.06, h: 0.8, fontSize: 9, color: '555555' });
  });

  sl.addText('Live prototype: meesho-returnsense-prototype.html', { x: 0.45, y: 6.88, w: 9.1, h: 0.28, fontSize: 10, color: C.pink, align: 'center' });
}

// ─────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('IMPACT & ROI', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.pink, bold: true, charSpacing: 2 });
  sl.addText('Loss Mitigation & Customer Retention Metrics', {
    x: 0.5, y: 0.62, w: 9, h: 0.7, fontSize: 21, color: C.white, bold: true,
  });

  sl.addText('CUSTOMER & SUPPLIER VALUE', { x: 0.45, y: 1.5, w: 4.5, h: 0.3, fontSize: 9.5, color: C.pink, bold: true, charSpacing: 1 });
  sl.addText('BUSINESS & FINANCIAL ROI', { x: 5.05, y: 1.5, w: 4.5, h: 0.3, fontSize: 9.5, color: C.pink, bold: true, charSpacing: 1 });

  const left = [
    { v: '↓ 40%', l: 'Sizing-related return rate' },
    { v: '↑ 18%', l: 'Customer LTV (apparel repeat)' },
    { v: '↓ 50%', l: 'Supplier disputes regarding sizes' },
    { v: '↑ 92%', l: 'Size chart data accuracy index' },
  ];
  const right = [
    { v: '↓ 25%', l: 'RTO (Return to Origin) rate' },
    { v: '₹12 Cr', l: 'Annual savings in shipping costs' },
    { v: '↑ 84%', l: 'COD validation success rate' },
    { v: '↓ 18%', l: 'Forward and reverse logistics burn' },
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
  sl.addText('Suppressed returns translate to instant margin recovery. By reducing logistics waste, Meesho retains more margin and passes savings to consumers in Tier 2/3/4 towns, maintaining a low-cost structure.', {
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
  sl.addText('Leveraging PhonePe Product Experience to Solve Meesho Returns', { x: 0.5, y: 0.62, w: 9, h: 0.68, fontSize: 26, color: C.dark, bold: true });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 1.42, w: 4.2, h: 4.08, fill: { color: C.dark }, shadow: makeShadow() });
  sl.addText('PhonePe Experience', { x: 0.62, y: 1.55, w: 3.86, h: 0.35, fontSize: 12.5, color: C.pink, bold: true });
  const proof = [
    'Managed user conversion and eligibility pipelines for 350M+ MAU checkout conversion flow (22% lift)',
    'Deployed ML propensity models that replaced rule-based targets, reducing marketing burn by 32%',
    'Redesigned transaction surfacing and checkout flows, resolving core conversion bottlenecks',
    'Collaborated with engineering to set up real-time analytics dashboards for tracking operational SLAs',
    'Designed merchant-facing dashboards for tracking campaign metrics and catalog optimizations',
  ];
  proof.forEach((t, i) => sl.addText(`• ${t}`, { x: 0.72, y: 2.0 + i * 0.65, w: 3.76, h: 0.56, fontSize: 10, color: C.white }));

  sl.addShape(prs.ShapeType.rect, { x: 4.82, y: 1.42, w: 4.7, h: 4.08, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('Meesho PM2 Relevance', { x: 4.98, y: 1.55, w: 4.38, h: 0.35, fontSize: 12.5, color: C.pink, bold: true });
  const maps = [
    ['Propensity ML suppression', '→ Risk checkout intercept model'],
    ['Checkout conversion UX', '→ Sizing warning & meter PDP entry'],
    ['Surfacing rules (350M MAU)', '→ Real-time size matching calculations'],
    ['Ops dashboards', '→ Supplier Hub return insights panel'],
    ['A/B testing methodologies', '→ Funnel tracking on returns mitigation'],
  ];
  maps.forEach(([f, m], i) => {
    sl.addText(f, { x: 4.98, y: 2.0 + i * 0.65, w: 2.2, h: 0.56, fontSize: 10, color: '444444', bold: true });
    sl.addText(m, { x: 7.2, y: 2.0 + i * 0.65, w: 2.2, h: 0.56, fontSize: 10, color: '333333' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.62, w: 9.1, h: 0.65, fill: { color: C.pink } });
  sl.addText('"Dynamic checkout friction is the digital equivalent of merchant validation. Applying this to high-risk COD orders directly mitigates logistical waste and returns."', {
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
  sl.addText('Sizing Pilot → COD Intercept → Supplier Platform Scaling', {
    x: 0.5, y: 0.62, w: 9, h: 0.65, fontSize: 24, color: C.white, bold: true,
  });

  const phases = [
    {
      phase: 'Phase 1', time: 'Month 1–2: Size Alert Pilot',
      body: 'Launch size warnings & matching widget for top 50 fashion suppliers. Benchmark returns rate against control. Target: 15% return reduction.',
    },
    {
      phase: 'Phase 2', time: 'Month 3–4: COD Intercept & WhatsApp',
      body: 'Introduce high-risk COD overlay and WhatsApp voice bot verification. A/B test discount levels (₹10 vs ₹15). Target: 20% RTO reduction.',
    },
    {
      phase: 'Phase 3', time: 'Month 5–6: Supplier Integration',
      body: 'Roll out size chart recommendations in Supplier Hub. Penalize visibility of unresponsive suppliers via PRISM feed. Target: 40% overall return drop.',
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
  sl.addText('Requirements: 1 ML engineer (RTO model), 1 Data analyst, integration with WhatsApp voice API (Vaani team), and Supplier catalog API access.', {
    x: 0.62, y: 5.1, w: 8.76, h: 0.62, fontSize: 10.5, color: C.accent,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.9, w: 9.1, h: 0.62, fill: { color: C.pink } });
  sl.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  ajay-avaghade.github.io/Portfolio', {
    x: 0.62, y: 5.94, w: 8.76, h: 0.52, fontSize: 10.5, color: C.white, align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/meesho-returnsense-deck.pptx' })
  .then(() => console.log('✅ assets/meesho-returnsense-deck.pptx written'))
  .catch(e => { console.error(e); process.exit(1); });
