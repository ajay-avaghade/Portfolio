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
  purple: '7D58BA',  // Vaani Purple
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

  sl.addText('MEESHO VOICE COMMERCE & USER EXPERIENCE', { x: 0.5, y: 0.32, w: 9, h: 0.28, fontSize: 9, color: C.pink, charSpacing: 2 });
  sl.addText('Vaani-Cart', { x: 0.5, y: 0.75, w: 9, h: 1.4, fontSize: 64, color: C.white, bold: true });
  sl.addText('Voice-First Conversational Checkout for the Next 500 Million Users', {
    x: 0.5, y: 2.0, w: 7.5, h: 0.6, fontSize: 22, color: C.accent, bold: false,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.5, y: 2.75, w: 2.8, h: 0.04, fill: { color: C.pink } });

  sl.addText('Presented by Ajay Avaghade · Product Manager 2 Candidate', {
    x: 0.5, y: 2.95, w: 7, h: 0.35, fontSize: 13, color: 'FFD5E5',
  });

  // Big stat
  sl.addShape(prs.ShapeType.rect, { x: 7.2, y: 5.2, w: 2.5, h: 1.9, fill: { color: C.pink }, shadow: makeShadow() });
  sl.addText('60%', { x: 7.2, y: 5.3, w: 2.5, h: 0.85, fontSize: 52, color: C.white, bold: true, align: 'center' });
  sl.addText('cart abandonment in\nnon-metro regions of India', { x: 7.2, y: 6.1, w: 2.5, h: 0.8, fontSize: 11, color: 'FFE6F0', align: 'center' });
}

// ─────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE PROBLEM', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.pink, bold: true, charSpacing: 2 });
  sl.addText('First-Time Internet Shoppers Face Significant Friction During Checkout Form Fields.', {
    x: 0.5, y: 0.62, w: 9, h: 0.9, fontSize: 21, color: C.white, bold: true,
  });

  const cols = [
    { icon: '⌨️', stat: '50-60%', lbl: 'Cart drop-offs at address', sub: 'Typing full addresses in English and pinpointing locations is difficult for non-literate buyers.' },
    { icon: '❌', stat: '40%', lbl: 'Invalid delivery address RTO', sub: 'Typing vague landmarks like "near school" without street details leads to delivery failures.' },
    { icon: '📞', stat: '30%', lbl: 'Support tickets from checkout', sub: 'Sellers and users face confusion during manual payment and address input on mobile screens.' },
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
  sl.addText('Root cause: Mobile checkout forms are optimized for desktop/Western paradigms. For India\'s Tier 2/3/4 towns, voice is the natural medium. Text-based inputs block e-commerce adoption for the next 500 million internet users.', {
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
  sl.addText('Shift from Static Typing to Voice-First Conversational Checkout in Local Languages.', {
    x: 0.5, y: 0.62, w: 9, h: 0.9, fontSize: 22, color: C.dark, bold: true,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 1.72, w: 4.1, h: 3.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('❌  Text-Based Checkout', { x: 0.62, y: 1.85, w: 3.76, h: 0.38, fontSize: 12, color: '444444', bold: true });
  const left = [
    'Complex text boxes requiring manual keyboard input',
    'English-only fields for Hindi/vernacular consumers',
    'No validation of addresses until delivery attempt',
    'Manual entry of landmarks is confusing on mobile screens',
    'Static confirmation screen with zero support assistance',
  ];
  left.forEach((t, i) => sl.addText(`• ${t}`, { x: 0.72, y: 2.35 + i * 0.42, w: 3.56, h: 0.36, fontSize: 11, color: '555555' }));

  sl.addShape(prs.ShapeType.ellipse, { x: 4.62, y: 2.9, w: 0.72, h: 0.72, fill: { color: C.pink } });
  sl.addText('VS', { x: 4.62, y: 2.9, w: 0.72, h: 0.72, fontSize: 11, color: C.white, bold: true, align: 'center', valign: 'middle' });

  sl.addShape(prs.ShapeType.rect, { x: 5.42, y: 1.72, w: 4.1, h: 3.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('✅  Vaani-Cart (Proposed)', { x: 5.58, y: 1.85, w: 3.76, h: 0.38, fontSize: 12, color: C.pink, bold: true });
  const right = [
    '1-click microphone to dictate details instantly',
    'Full vernacular integration supporting regional dialects',
    'Speech-to-Address parser auto-fills structured forms',
    'Auto-validation maps spoken landmarks to pin locations',
    'Generative audio bot confirms order details aloud',
  ];
  right.forEach((t, i) => sl.addText(`• ${t}`, { x: 5.58, y: 2.35 + i * 0.42, w: 3.72, h: 0.36, fontSize: 10.5, color: '333333' }));

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.08, w: 9.1, h: 0.72, fill: { color: C.pink } });
  sl.addText('Key insight: Voice inputs eliminate the digital literacy gap. Translating regional dialects to structured form values at checkout makes transactions feel like talking to a local shopkeeper.', {
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
  sl.addText('Vaani-Cart Stack: Capture, Parse, Validate, and Speak', {
    x: 0.5, y: 0.62, w: 9, h: 0.65, fontSize: 22, color: C.white, bold: true,
  });

  const steps = [
    { n: '01', title: 'Voice Checkout Banner', body: 'Surfaces a glowing purple microphone widget directly in the shopping cart. "बोलकर आर्डर करें" CTA prompts user engagement.' },
    { n: '02', title: 'Speech Capture & Wave', body: 'Edge speech processing starts on microphone tap. User describes address details in Hindi, Bengali, or local dialects.' },
    { n: '03', title: 'Real-Time Address Parsing', body: 'AI parser breaks down speech input into structured parameters: City, Locality, Landmark, House Number, and Pincode.' },
    { n: '04', title: 'Landmark Verification', body: 'Resolves spoken landmarks against geographic databases to assign precise GPS pins, preventing delivery failures.' },
    { n: '05', title: 'Generative Audio Confirmation', body: 'Vaani speaks out the confirmation summary: "Your order is confirmed for Gandhi Nagar. Total is ₹482 COD." Users confirm via simple voice response.' },
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
  sl.addText('PhonePe Parallel: Surfacing cart-level triggers to incentivize checkouts maps directly to cart optimization. Surfacing real-time parameters dynamically mirrors offers eligibility engines.', {
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
  sl.addText('Interactive Prototypes — Voice Checkout Funnel and Analytics', {
    x: 0.5, y: 0.62, w: 9, h: 0.55, fontSize: 22, color: C.dark, bold: true,
  });

  const screens = [
    { n: '01', title: 'Cart Voice Banner', desc: 'Prominent Vaani Voice Checkout card in shopping cart. Microphone tap activates recording wave.' },
    { n: '02', title: 'Speech Form Fill', desc: 'Displays spoken address parsed in real-time. Structured fields (City, Landmark, Pincode) auto-populate.' },
    { n: '03', title: 'Audio Confirmation', desc: 'GenAI voice reciting confirmation: "Order for Gandhi Nagar. Say Yes to confirm." Simple voice agreement.' },
    { n: '04', title: 'Funnel Analytics', desc: 'Ops dashboard mapping voice completion vs. text checkouts. Visualizes language distribution.' },
  ];

  screens.forEach((s, i) => {
    const x = 0.45 + i * 2.38;
    sl.addShape(prs.ShapeType.rect, { x, y: 1.38, w: 2.18, h: 3.5, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(prs.ShapeType.rect, { x, y: 1.38, w: 2.18, h: 0.32, fill: { color: C.pink } });
    sl.addText(`${s.n} — ${s.title}`, { x, y: 1.38, w: 2.18, h: 0.32, fontSize: 9.5, color: C.white, bold: true, align: 'center', valign: 'middle' });

    const mockup = i === 0
      ? '🛒 My Cart (1 Item)\n─────────────────\nAnarkali Kurta Set (M)\nPrice: ₹482\n\n✨ MEESHO VAANI AI\nबोलकर आर्डर करें (Voice)\n[ 🎤 Tap to Speak Address ]\n  | | | | | Wave active'
      : i === 1
      ? '🗣️ Address Parser\n─────────────────\n"राम मंदिर के पीछे, गांधी नगर,\n पटना..."\n\nParsed Details:\n• City: Patna\n• Locality: Gandhi Nagar\n• Landmark: Behind Temple'
      : i === 2
      ? '🎙️ Confirmation\n─────────────────\n🔊 "आर्डर पटना के लिए है।\nकुल ₹482 COD. आर्डर करने\nके लिए हाँ बोलें"\n\nSay: "हाँ आर्डर कर दो"\n[ Confirm Order ✓ ]'
      : '📊 Vaani-Cart Metrics\n─────────────────\n• Voice Share: 34.6%\n• Completion: 82.4%\n\nFunnel completion:\nVoice:  ████████ 82.4%\nEnglish:██████ 68.2%\nHindi:  █████ 58.5%';

    sl.addText(mockup, { x: x + 0.08, y: 1.82, w: 2.02, h: 2.6, fontSize: 7.5, color: '333333', fontFace: 'Courier New' });
    sl.addText(s.desc, { x: x + 0.06, y: 4.94, w: 2.06, h: 0.8, fontSize: 9, color: '555555' });
  });

  sl.addText('Live prototype: meesho-vaanicart-prototype.html', { x: 0.45, y: 6.88, w: 9.1, h: 0.28, fontSize: 10, color: C.pink, align: 'center' });
}

// ─────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('IMPACT & ROI', { x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 9, color: C.pink, bold: true, charSpacing: 2 });
  sl.addText('Funnel Lift & Address Validation Metrics', {
    x: 0.5, y: 0.62, w: 9, h: 0.7, fontSize: 21, color: C.white, bold: true,
  });

  sl.addText('USER ENGAGEMENT & CONVERSION', { x: 0.45, y: 1.5, w: 4.5, h: 0.3, fontSize: 9.5, color: C.pink, bold: true, charSpacing: 1 });
  sl.addText('OPERATIONAL & LOGISTICAL ROI', { x: 5.05, y: 1.5, w: 4.5, h: 0.3, fontSize: 9.5, color: C.pink, bold: true, charSpacing: 1 });

  const left = [
    { v: '↑ 14%', l: 'Checkout conversion rate (lift)' },
    { v: '↑ 34%', l: 'Share of checkouts via voice' },
    { v: '↓ 50%', l: 'Support tickets regarding checkouts' },
    { v: '82%', l: 'Voice-checkout user retention' },
  ];
  const right = [
    { v: '↓ 40%', l: 'Invalid address delivery failures' },
    { v: '↑ 28%', l: 'Address parsing match accuracy' },
    { v: '↓ 18%', l: 'RTO due to incorrect address data' },
    { v: '₹8 Cr', l: 'Annual savings in delivery re-routing' },
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
  sl.addText('Voice-driven cart conversion opens transactional doors to non-tech-savvy users. Eliminating keyboard barriers makes ordering simple, driving brand penetration in rural markets.', {
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
  sl.addText('Leveraging PhonePe Cart & UX Optimization to Scale Vaani-Cart', { x: 0.5, y: 0.62, w: 9, h: 0.68, fontSize: 26, color: C.dark, bold: true });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 1.42, w: 4.2, h: 4.08, fill: { color: C.dark }, shadow: makeShadow() });
  sl.addText('PhonePe Experience', { x: 0.62, y: 1.55, w: 3.86, h: 0.35, fontSize: 12.5, color: C.pink, bold: true });
  const proof = [
    'Built dynamic cart incentivization engine for quick commerce, reducing cart abandonment by 60%',
    'Diagnosed checkout drops and designed contextual prompts that raised average order value (35% lift)',
    'Deployed ML propensity models for surfacing transaction-level offers across large scales',
    'Maintained and optimized payment confirmation scripts across complex checkout corridoors',
    'Set up performance analytics pipelines to track funnel conversions at the city level',
  ];
  proof.forEach((t, i) => sl.addText(`• ${t}`, { x: 0.72, y: 2.0 + i * 0.65, w: 3.76, h: 0.56, fontSize: 10, color: C.white }));

  sl.addShape(prs.ShapeType.rect, { x: 4.82, y: 1.42, w: 4.7, h: 4.08, fill: { color: C.white }, shadow: makeShadow() });
  sl.addText('Meesho PM2 Relevance', { x: 4.98, y: 1.55, w: 4.38, h: 0.35, fontSize: 12.5, color: C.pink, bold: true });
  const maps = [
    ['Cart abandonment diagnostics', '→ Vaani-Cart address validation focus'],
    ['Contextual checkout prompts', '→ Voice mic surfacing & audio brief'],
    ['Transaction eligibility engine', '→ Speech-to-Address auto form fill'],
    ['Funnel conversion metrics', '→ Checkout drop-off dashboard metrics'],
    ['P&L impact ownership', '→ Shipping cost recovery tracking'],
  ];
  maps.forEach(([f, m], i) => {
    sl.addText(f, { x: 4.98, y: 2.0 + i * 0.65, w: 2.2, h: 0.56, fontSize: 10, color: '444444', bold: true });
    sl.addText(m, { x: 7.2, y: 2.0 + i * 0.65, w: 2.2, h: 0.56, fontSize: 10, color: '333333' });
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.62, w: 9.1, h: 0.65, fill: { color: C.pink } });
  sl.addText('"Voice inputs are checkout boosters. Applying conversational flows to address entry resolves the single biggest drop-off point in Tier-2+ Indian e-commerce."', {
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
  sl.addText('Cart Integration → Address Parsing Test → Platform Rollout', {
    x: 0.5, y: 0.62, w: 9, h: 0.65, fontSize: 24, color: C.white, bold: true,
  });

  const phases = [
    {
      phase: 'Phase 1', time: 'Month 1–2: Cart Voice Banner',
      body: 'Integrate Vaani mic widget into cart page for 10% of Hindi-speaking users. Measure mic tap rate and basic completion metrics. Target: 10% drop-off reduction.',
    },
    {
      phase: 'Phase 2', time: 'Month 3–4: Address Parser Scaling',
      body: 'Introduce landmark-mapping API and speech-to-form fills. A/B test voice confirmation voice notes in 3 languages (Hindi, Bengali, Telugu). Target: 15% checkout lift.',
    },
    {
      phase: 'Phase 3', time: 'Month 5–6: Full Platform Scaling',
      body: 'Roll out voice checkout platform-wide. Link Vaani voice help to delivery agent coordinate maps to ensure seamless last-mile drops. Target: 40% invalid address reduction.',
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
  sl.addText('Requirements: Speech-to-text API integration, landmark resolution database, 1 UI/UX designer, 1 Android engineer, and analytics support.', {
    x: 0.62, y: 5.1, w: 8.76, h: 0.62, fontSize: 10.5, color: C.accent,
  });

  sl.addShape(prs.ShapeType.rect, { x: 0.45, y: 5.9, w: 9.1, h: 0.62, fill: { color: C.pink } });
  sl.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  ajay-avaghade.github.io/Portfolio', {
    x: 0.62, y: 5.94, w: 8.76, h: 0.52, fontSize: 10.5, color: C.white, align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/meesho-vaanicart-deck.pptx' })
  .then(() => console.log('✅ assets/meesho-vaanicart-deck.pptx written'))
  .catch(e => { console.error(e); process.exit(1); });
