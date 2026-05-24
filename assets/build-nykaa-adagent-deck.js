const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0D0A1A',
  hero:   '1A1430',
  purple: '8B5CF6',
  violet: 'A78BFA',
  white:  'FFFFFF',
  lgray:  'F5F3FF',
  ink:    '1A0A30',
  muted:  '8B7AAA',
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
    sl.addShape(prs.ShapeType.line, { x: -0.5 + i * 1.5, y: 0, w: 4, h: 7.5, line: { color: '8B5CF6', width: 0.4, transparency: 88 } });
  }

  sl.addText('NYKAA SUPERSTORE — B2B SELLER TOOLS', { x: 0.5, y: 0.32, w: 9, h: 0.28, fontSize: 9, color: C.violet, charSpacing: 2 });
  sl.addText('AdAgent', { x: 0.5, y: 0.75, w: 9, h: 1.4, fontSize: 64, color: C.white, bold: true });
  sl.addText('Agentic AI That Plans, Launches & Optimises Seller Campaigns — End to End', {
    x: 0.5, y: 2.0, w: 7.5, h: 0.6, fontSize: 22, color: C.violet, bold: false,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.5, y: 2.75, w: 2.8, h: 0.04, fill: { color: C.purple } });

  sl.addText('Presented by Ajay Avaghade · Growth & Monetization PM', {
    x: 0.5, y: 2.95, w: 7, h: 0.35, fontSize: 13, color: 'BBAAD0',
  });

  // Big stat
  sl.addShape(prs.ShapeType.rect, { x: 7.2, y: 5.2, w: 2.5, h: 1.9, fill: { color: C.purple }, shadow: makeShadow() });
  sl.addText('60%', { x: 7.2, y: 5.3, w: 2.5, h: 0.85, fontSize: 52, color: C.white, bold: true, align: 'center' });
  sl.addText('of sellers have\nnever run an ad', { x: 7.2, y: 6.1, w: 2.5, h: 0.8, fontSize: 11, color: 'D0C0F0', align: 'center' });
}

// ─────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE PROBLEM', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.violet, bold: true, charSpacing: 2 });
  sl.addText('Most Nykaa Sellers Leave Ad Revenue on the Table — Because Running Ads Requires Expertise They Don\'t Have', {
    x: 0.5, y: 0.62, w: 9, h: 0.9, fontSize: 21, color: C.white, bold: true,
  });

  const cols = [
    { icon: '🧠', stat: '60%', lbl: 'Sellers never ran an ad', sub: 'Ad Manager requires keyword research, bid calibration, creative testing — skills most small sellers don\'t have' },
    { icon: '⏰', stat: '4+ hrs', lbl: 'Manual ad management weekly', sub: 'Sellers who do run ads spend 4+ hours/week on campaign tweaks, report reading, and budget adjustments' },
    { icon: '💸', stat: '1.8×', lbl: 'ROAS gap: experts vs. novices', sub: 'Expert-managed campaigns outperform self-run by 1.8× — same budget, same catalog, just better execution' },
  ];
  cols.forEach((c, i) => {
    const x = 0.45 + i * 3.12;
    sl.addShape(prs.ShapeType.rect, { x, y: 1.72, w: 2.9, h: 3.1, fill: { color: C.hero }, shadow: makeShadow() });
    sl.addText(c.icon, { x, y: 1.88, w: 2.9, h: 0.55, fontSize: 26, align: 'center' });
    sl.addText(c.stat, { x, y: 2.52, w: 2.9, h: 0.72, fontSize: 40, color: C.purple, bold: true, align: 'center' });
    sl.addText(c.lbl, { x, y: 3.22, w: 2.9, h: 0.38, fontSize: 12.5, color: C.white, bold: true, align: 'center' });
    sl.addText(c.sub, { x: x + 0.1, y: 3.62, w: 2.7, h: 0.9, fontSize: 10, color: 'BBAAD0', align: 'center' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.0, w: 9.1, h: 0.82, fill: { color: C.purple } });
  sl.addText('Root cause: The gap isn\'t budget — it\'s execution capability. A seller with ₹5,000/week can achieve 3.5× ROAS if the campaign is set up correctly. Most never get there because nobody guides the setup.', {
    x: 0.62, y: 5.06, w: 8.76, h: 0.7, fontSize: 10.5, color: C.white,
  });
}

// ─────────────────────────────────────────
// SLIDE 3 — THE INSIGHT
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE INSIGHT', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.purple, bold: true, charSpacing: 2 });
  sl.addText('Sellers Don\'t Want to Learn Ad Management. They Want Sales. Give Them an Agent That Does It.', {
    x: 0.5, y: 0.62, w: 9, h: 0.9, fontSize: 22, color: C.ink, bold: true,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 1.72, w: 4.1, h: 3.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('❌  Self-Serve Ad Manager (Current)', { x: 0.62, y: 1.85, w: 3.76, h: 0.38, fontSize: 12, color: '444444', bold: true });
  const left = ['Keyword research — 1 hour', 'Bid calibration — guesswork', 'Creative selection — manual', 'Campaign monitoring — daily', 'Result: 60% never start, 30% give up in week 2'];
  left.forEach((t, i) => sl.addText(`• ${t}`, { x: 0.72, y: 2.35 + i * 0.42, w: 3.56, h: 0.36, fontSize: 11, color: '555555' }));

  sl.addShape(prs.ShapeType.ellipse, { x: 4.62, y: 2.9, w: 0.72, h: 0.72, fill: { color: C.purple } });
  sl.addText('VS', { x: 4.62, y: 2.9, w: 0.72, h: 0.72, fontSize: 11, color: C.white, bold: true, align: 'center', valign: 'middle' });

  sl.addShape(prs.ShapeType.rect, { x: 5.42, y: 1.72, w: 4.1, h: 3.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('✅  AdAgent (Proposed)', { x: 5.58, y: 1.85, w: 3.76, h: 0.38, fontSize: 12, color: C.purple, bold: true });
  const right = ['Seller states goal: "Maximize sales for Vitamin C Serum"', 'Agent analyzes catalog + competition in 90 seconds', 'Agent selects keywords, generates creatives, sets bids', 'Campaign live in 3 minutes, optimised every 6 hours', 'Seller gets weekly: "You earned ₹74,400. Here\'s what worked."'];
  right.forEach((t, i) => sl.addText(`• ${t}`, { x: 5.58, y: 2.35 + i * 0.42, w: 3.72, h: 0.36, fontSize: 10.5, color: '333333' }));

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.08, w: 9.1, h: 0.72, fill: { color: C.purple } });
  sl.addText('Key insight: The bottleneck is not knowledge — it\'s time and confidence. An agentic layer that handles the entire execution chain (with the seller only setting the goal) unlocks the 60% who have never run a single ad.', {
    x: 0.62, y: 5.12, w: 8.76, h: 0.62, fontSize: 10, color: C.white,
  });
}

// ─────────────────────────────────────────
// SLIDE 4 — THE MECHANIC
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE MECHANIC', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.violet, bold: true, charSpacing: 2 });
  sl.addText('Goal → Agent → Campaign → Optimize → Report: The Full Agentic Loop', {
    x: 0.5, y: 0.62, w: 9, h: 0.65, fontSize: 22, color: C.white, bold: true,
  });

  const steps = [
    { n: '01', title: 'Goal Input', body: 'Seller selects SKU + campaign goal (Maximize Sales / Maximize Visibility / Defend Rank). Sets weekly budget. One screen, 60 seconds.' },
    { n: '02', title: 'Catalog & Competitor Analysis', body: 'Agent reads SKU metadata, category search trends, top competitor keywords, and historical brand performance. Builds a targeting brief automatically.' },
    { n: '03', title: 'Keyword + Creative + Bid', body: 'Agent selects 15–30 keywords, generates 3 ad creative variants, sets opening bids using category benchmarks. Presents a plan for seller review (optional).' },
    { n: '04', title: 'Live Optimization Loop', body: 'Every 6 hours: pause underperforming keywords, shift budget to top performers, rotate creatives. All decisions logged in plain language for seller transparency.' },
    { n: '05', title: 'Weekly Plain-Language Report', body: '"Your Vitamin C Serum earned ₹74,400 this week. Best-performing creative was the before/after image. Suggestion: increase budget 20% — I\'ll handle the rest."' },
  ];

  sl.addShape(prs.ShapeType.line, { x: 0.88, y: 1.82, w: 0, h: 4.2, line: { color: C.purple, width: 1.5, dashType: 'dash' } });

  steps.forEach((s, i) => {
    const y = 1.65 + i * 0.86;
    sl.addShape(prs.ShapeType.ellipse, { x: 0.6, y, w: 0.55, h: 0.55, fill: { color: C.purple } });
    sl.addText(s.n, { x: 0.6, y, w: 0.55, h: 0.55, fontSize: 10, color: C.white, bold: true, align: 'center', valign: 'middle' });
    sl.addText(s.title, { x: 1.3, y, w: 2.6, h: 0.3, fontSize: 11.5, color: C.purple, bold: true });
    sl.addText(s.body, { x: 1.3, y: y + 0.3, w: 8.1, h: 0.44, fontSize: 9.5, color: 'BBAAD0' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 6.18, w: 9.1, h: 0.62, fill: { color: C.hero } });
  sl.addText('PhonePe Proof: Led deployment of Propensity-to-Transact ML models replacing ₹1,000+ Cr/yr of static targeting — same agentic pattern: input (user signal) → autonomous execution (scoring + campaign selection) → outcome (32% burn reduction).', {
    x: 0.62, y: 6.22, w: 8.76, h: 0.52, fontSize: 9.5, color: C.violet,
  });
}

// ─────────────────────────────────────────
// SLIDE 5 — THE PRODUCT
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE PRODUCT', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.purple, bold: true, charSpacing: 2 });
  sl.addText('4 Key Screen States — The Full Agent Experience', {
    x: 0.5, y: 0.62, w: 9, h: 0.55, fontSize: 22, color: C.ink, bold: true,
  });

  const screens = [
    { n: '01', title: 'Goal Setup', desc: 'SKU selector, goal picker, budget slider. AI note previews what the agent will do. One screen, 60 seconds to launch.' },
    { n: '02', title: 'Agent Running', desc: '5-step progress ring: catalog ✓ keywords ✓ creatives ✓ bids 🔄 launch ⬜. Real-time transparency into agent execution.' },
    { n: '03', title: 'Live Campaign', desc: '₹42,800 attributed sales, 3.8× ROAS, agent optimization log showing decisions made in the last 6 hours.' },
    { n: '04', title: 'Weekly Report', desc: '₹74,400 sales summary in plain language. 2 agent recommendations with estimated impact. One-tap "Accept All".' },
  ];

  screens.forEach((s, i) => {
    const x = 0.45 + i * 2.38;
    sl.addShape(prs.ShapeType.rect, { x, y: 1.38, w: 2.18, h: 3.5, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(prs.ShapeType.rect, { x, y: 1.38, w: 2.18, h: 0.32, fill: { color: C.purple } });
    sl.addText(`${s.n} — ${s.title}`, { x, y: 1.38, w: 2.18, h: 0.32, fontSize: 9.5, color: C.white, bold: true, align: 'center', valign: 'middle' });

    const mockup = i === 0
      ? '[ Mamaearth Vit C ]\nGoal: [Maximize Sales]\nBudget: ₹5,500/wk\n─────────────────\n🤖 "I will select top\nkeywords, generate 3\ncreative variants &\noptimise every 6 hrs"\n[ Launch AdAgent ]'
      : i === 1
      ? '   ◉ 80% Complete\n─────────────────\n✓ Catalog analyzed\n✓ Keywords selected\n✓ Creatives generated\n🔄 Setting bids...\n⬜ Launching...\n─────────────────\nEst. live in 45 sec'
      : i === 2
      ? '₹42,800 attributed\n3.8× ROAS · 48K impr\n─────────────────\n🤖 Agent log:\n↑ Bid +12% on "vit c"\n⏸ Paused "face serum"\n🔄 Rotated to creative B\n─────────────────\nSpend: ₹11,200'
      : '📊 Week 3 Summary\n₹74,400 total sales\n148 orders · 3.6× ROAS\n─────────────────\n💡 Recommendations:\n• Budget +20% → +₹14K\n• New keyword: "glow"\n─────────────────\n[ Accept All ]';

    sl.addText(mockup, { x: x + 0.08, y: 1.82, w: 2.02, h: 2.6, fontSize: 7.5, color: '333333', fontFace: 'Courier New' });
    sl.addText(s.desc, { x: x + 0.06, y: 4.94, w: 2.06, h: 0.8, fontSize: 9, color: '555555' });
  });

  sl.addText('Live prototype: nykaa-adagent-prototype.html', { x: 0.45, y: 6.88, w: 9.1, h: 0.28, fontSize: 10, color: C.purple, align: 'center' });
}

// ─────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('IMPACT & ROI', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.violet, bold: true, charSpacing: 2 });
  sl.addText('Projected Impact — Anchored in PhonePe ML Targeting & Campaign Management Proof', {
    x: 0.5, y: 0.62, w: 9, h: 0.7, fontSize: 21, color: C.white, bold: true,
  });

  sl.addText('SELLER IMPACT', { x: 0.45, y: 1.5, w: 4.5, h: 0.3, fontSize: 9.5, color: C.purple, bold: true, charSpacing: 1 });
  sl.addText('NYKAA PLATFORM ROI', { x: 5.05, y: 1.5, w: 4.5, h: 0.3, fontSize: 9.5, color: C.purple, bold: true, charSpacing: 1 });

  const left = [
    { v: '↑ 2.1×', l: 'ROAS for first-time advertisers' },
    { v: '↓ 97%', l: 'Setup time (3 min vs 4 hrs)' },
    { v: '+60%', l: 'Sellers activating ads for first time' },
    { v: '↑ 34%', l: 'GMV attributed to ads (per seller)' },
  ];
  const right = [
    { v: '↑ 40%', l: 'Ad revenue per active seller' },
    { v: '₹0', l: 'Incremental human ops cost' },
    { v: '3.8×', l: 'Platform ROAS on managed campaigns' },
    { v: '↑ 25%', l: 'Ad platform DAU (sellers logging in)' },
  ];

  left.forEach((m, i) => {
    const y = 1.92 + i * 0.84;
    sl.addShape(prs.ShapeType.rect, { x: 0.45, y, w: 4.38, h: 0.72, fill: { color: C.hero }, shadow: makeShadow() });
    sl.addText(m.v, { x: 0.6, y: y + 0.05, w: 1.5, h: 0.62, fontSize: 26, color: C.purple, bold: true, valign: 'middle' });
    sl.addText(m.l, { x: 2.18, y: y + 0.05, w: 2.6, h: 0.62, fontSize: 12, color: C.white, valign: 'middle' });
  });
  right.forEach((m, i) => {
    const y = 1.92 + i * 0.84;
    sl.addShape(prs.ShapeType.rect, { x: 5.05, y, w: 4.48, h: 0.72, fill: { color: C.hero }, shadow: makeShadow() });
    sl.addText(m.v, { x: 5.2, y: y + 0.05, w: 1.5, h: 0.62, fontSize: 26, color: C.purple, bold: true, valign: 'middle' });
    sl.addText(m.l, { x: 6.78, y: y + 0.05, w: 2.7, h: 0.62, fontSize: 12, color: C.white, valign: 'middle' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.44, w: 9.1, h: 0.82, fill: { color: C.purple } });
  sl.addText('Each new seller that AdAgent activates on ads is incremental ad revenue at near-zero marginal cost — the agent replaces a human account manager. At 10,000 newly activated sellers spending ₹5,000/week: ₹50 Cr/week in managed ad spend.', {
    x: 0.62, y: 5.5, w: 8.76, h: 0.68, fontSize: 10.5, color: C.white,
  });
}

// ─────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('PROOF OF WORK', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.purple, bold: true, charSpacing: 2 });
  sl.addText('I Built This. Here\'s the Proof.', { x: 0.5, y: 0.62, w: 9, h: 0.68, fontSize: 26, color: C.ink, bold: true });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 1.42, w: 4.2, h: 4.08, fill: { color: C.dark }, shadow: makeShadow() });
  sl.addText('What I Built at PhonePe', { x: 0.62, y: 1.55, w: 3.86, h: 0.35, fontSize: 12.5, color: C.purple, bold: true });
  const proof = [
    'Led Propensity-to-Transact ML deployment replacing ₹1,000+ Cr/yr static targeting — same autonomous execution pattern as AdAgent',
    'Rebuilt ₹100 Cr/yr rewards portfolio into ML-driven marketplace with audience segmentation + self-serve brand campaign tooling',
    'Owned full DS collaboration cycle: feature engineering → model → production rollout → monitoring (32% marketing burn reduction)',
    'Built dynamic cart incentivization engine using context signals (cart value × intent × time) → 35% AOV uplift',
    'A/B tested 6 campaign variants in parallel — same experimentation discipline as AdAgent creative rotation logic',
  ];
  proof.forEach((t, i) => sl.addText(`• ${t}`, { x: 0.72, y: 2.0 + i * 0.65, w: 3.76, h: 0.56, fontSize: 10, color: C.white }));

  sl.addShape(prs.ShapeType.rect, { x: 4.82, y: 1.42, w: 4.7, h: 4.08, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('How It Maps to This JD', { x: 4.98, y: 1.55, w: 4.38, h: 0.35, fontSize: 12.5, color: C.purple, bold: true });
  const maps = [
    ['Propensity-to-Transact ML', '→ AdAgent keyword + bid selection engine'],
    ['ML-driven offer marketplace', '→ AdAgent creative optimisation loop'],
    ['DS product collaboration', '→ AdAgent ML model scoping & production'],
    ['A/B experimentation', '→ AdAgent creative variant rotation'],
    ['Self-serve brand tooling', '→ AdAgent goal-input + report UX'],
  ];
  maps.forEach(([f, m], i) => {
    sl.addText(f, { x: 4.98, y: 2.0 + i * 0.65, w: 2.2, h: 0.56, fontSize: 10, color: '444444', bold: true });
    sl.addText(m, { x: 7.2, y: 2.0 + i * 0.65, w: 2.2, h: 0.56, fontSize: 10, color: '333333' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.62, w: 9.1, h: 0.65, fill: { color: C.purple } });
  sl.addText('"AdAgent is the same autonomous signal-to-action loop I shipped at PhonePe at 350M MAU — adapted for Nykaa\'s seller activation problem."', {
    x: 0.62, y: 5.66, w: 8.76, h: 0.56, fontSize: 10.5, color: C.white, italic: true,
  });
}

// ─────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('ROLLOUT PLAN', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.violet, bold: true, charSpacing: 2 });
  sl.addText('Phased Rollout — Closed Beta → Optimize → Scale', {
    x: 0.5, y: 0.62, w: 9, h: 0.65, fontSize: 24, color: C.white, bold: true,
  });

  const phases = [
    {
      phase: 'Phase 1', time: 'Month 1–2: Closed Beta',
      body: 'Deploy AdAgent to 500 sellers who have never run an ad. Measure: campaign activation rate, ROAS vs. self-managed control, seller session time reduction. Target: 40% activation.',
    },
    {
      phase: 'Phase 2', time: 'Month 3–4: Optimization Loop',
      body: 'Activate 6-hour optimization cycle. A/B test 3 creative strategies per SKU category. Tune keyword selection model on first 500-seller dataset. Launch weekly report UX.',
    },
    {
      phase: 'Phase 3', time: 'Month 5–6: Full Rollout',
      body: 'Roll out to all sellers. Integrate AdAgent entry point into onboarding flow (day 7 nudge). Launch "AdAgent Managed" badge for top-performing sellers on PDP.',
    },
  ];

  phases.forEach((p, i) => {
    const x = 0.45 + i * 3.12;
    sl.addShape(prs.ShapeType.rect, { x, y: 1.5, w: 2.92, h: 3.4, fill: { color: C.hero }, shadow: makeShadow() });
    sl.addShape(prs.ShapeType.rect, { x, y: 1.5, w: 2.92, h: 0.52, fill: { color: C.purple } });
    sl.addText(p.phase, { x, y: 1.5, w: 2.92, h: 0.52, fontSize: 15, color: C.white, bold: true, align: 'center', valign: 'middle' });
    sl.addText(p.time, { x: x + 0.1, y: 2.12, w: 2.72, h: 0.35, fontSize: 11, color: C.violet, bold: true });
    sl.addText(p.body, { x: x + 0.1, y: 2.5, w: 2.72, h: 2.2, fontSize: 10, color: 'BBAAD0' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.06, w: 9.1, h: 0.72, fill: { color: C.hero } });
  sl.addText('What I need to build this: Category keyword index + historical campaign data · 2 ML engineers (keyword model + bid optimizer) · 1 NLP engineer (creative generation) · Seller UX designer', {
    x: 0.62, y: 5.1, w: 8.76, h: 0.62, fontSize: 10.5, color: C.violet,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.9, w: 9.1, h: 0.62, fill: { color: C.purple } });
  sl.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  ajay-avaghade.github.io/Portfolio', {
    x: 0.62, y: 5.94, w: 8.76, h: 0.52, fontSize: 10.5, color: C.white, align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/nykaa-adagent-deck.pptx' })
  .then(() => console.log('✅ nykaa-adagent-deck.pptx written'))
  .catch(e => { console.error(e); process.exit(1); });
