const PptxGen = require('pptxgenjs');

const prs = new PptxGen();
prs.defineLayout({ name: 'LAYOUT_16x9', width: 10, height: 7.5 });
prs.layout = 'LAYOUT_16x9';

const C = {
  dark: '0F0E0E',
  hero: '2A2A2A',
  bright: 'FF0000',
  blue: '0078D4',
  gold: 'FFB81C',
  white: 'FFFFFF',
  lgray: 'F7F7F7',
  muted: '757575',
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 3, offset: 1, angle: 45, opacity: 0.12 });

// ===== SLIDE 1: COVER =====
let slide = prs.addSlide();
slide.background = { color: C.dark };

slide.addShape('rect', { x: 7.5, y: 0, w: 2.5, h: 7.5, fill: { color: C.blue, transparency: 20 }, line: { type: 'none' } });
slide.addShape('rect', { x: 8, y: 0, w: 2, h: 7.5, fill: { color: C.gold, transparency: 25 }, line: { type: 'none' } });

slide.addText('ADOBE DSP · PRODUCT CASE STUDY', {
  x: 0.3, y: 0.4, w: 4, h: 0.3,
  fontSize: 9, bold: true, color: C.gold, fontFace: 'Arial',
});

slide.addText('Audience Codec', {
  x: 0.3, y: 1.8, w: 7, h: 1.2,
  fontSize: 44, bold: true, color: C.white, fontFace: 'Arial',
});

slide.addText('Conversational Audience Assembly via Agent', {
  x: 0.3, y: 3, w: 7, h: 0.4,
  fontSize: 18, color: C.gold, fontFace: 'Arial',
});

slide.addText('Ajay Avaghade · Product Manager', {
  x: 0.3, y: 3.6, w: 7, h: 0.3,
  fontSize: 12, color: C.muted, fontFace: 'Arial',
});

slide.addShape('rect', { x: 6.5, y: 5.8, w: 3, h: 1.3, fill: { color: C.hero }, line: { type: 'none' } });
slide.addText('5–7 days → 8 mins', {
  x: 6.5, y: 5.95, w: 3, h: 0.4,
  fontSize: 32, bold: true, color: C.blue, fontFace: 'Arial', align: 'center',
});
slide.addText('Audience creation time', {
  x: 6.5, y: 6.4, w: 3, h: 0.35,
  fontSize: 10, color: C.muted, fontFace: 'Arial', align: 'center',
});

// ===== SLIDE 2: THE PROBLEM =====
slide = prs.addSlide();
slide.background = { color: C.dark };

slide.addText('THE PROBLEM', {
  x: 0.3, y: 0.3, w: 9, h: 0.25,
  fontSize: 9, bold: true, color: C.gold, fontFace: 'Arial',
});

slide.addText('60% of Advertisers Stuck at Basic Targeting—Lost Revenue from Untapped Tier-2 Sophistication', {
  x: 0.3, y: 0.8, w: 9, h: 0.8,
  fontSize: 26, bold: true, color: C.white, fontFace: 'Arial',
});

const statCols2 = [
  { val: '60%', label: 'Basic Segment Only', desc: 'No sequential or lookalike' },
  { val: '5–7 days', label: 'Ops Review TAT', desc: 'Per new audience request' },
  { val: '16%', label: 'ROAS Uplift Lost', desc: 'From advanced targeting' },
];

statCols2.forEach((stat, i) => {
  const x = 0.3 + i * 3.2;
  slide.addText(stat.val, {
    x, y: 2.2, w: 3, h: 0.5,
    fontSize: 28, bold: true, color: C.bright, fontFace: 'Arial', align: 'center',
  });
  slide.addText(stat.label, {
    x, y: 2.8, w: 3, h: 0.3,
    fontSize: 12, bold: true, color: C.white, fontFace: 'Arial', align: 'center',
  });
  slide.addText(stat.desc, {
    x, y: 3.15, w: 3, h: 0.6,
    fontSize: 10, color: C.muted, fontFace: 'Arial', align: 'center',
  });
});

slide.addShape('rect', { x: 0.3, y: 4.2, w: 9.2, h: 2.8, fill: { color: C.hero }, line: { color: C.bright, width: 2 }, shadow: makeShadow() });
slide.addText('Root Cause', {
  x: 0.5, y: 4.35, w: 8.8, h: 0.2,
  fontSize: 10, bold: true, color: C.gold, fontFace: 'Arial',
});
slide.addText('DSP audience building requires SQL expertise or ops overhead. Sequential audiences, lookalike seeds, and frequency caps are hidden behind technical configuration. Mid-market advertisers lack expertise; tier-1 teams hoard knowledge. Result: 60% stay at tier 1, leaving money on the table.', {
  x: 0.5, y: 4.65, w: 8.8, h: 2.1,
  fontSize: 12, color: C.white, fontFace: 'Arial',
});

// ===== SLIDE 3: THE INSIGHT =====
slide = prs.addSlide();
slide.background = { color: C.white };

slide.addText('THE INSIGHT', {
  x: 0.3, y: 0.3, w: 9, h: 0.25,
  fontSize: 9, bold: true, color: C.bright, fontFace: 'Arial',
});

slide.addText('Let Users Talk. Let Agents Translate Intent to Rules.', {
  x: 0.3, y: 0.8, w: 9, h: 0.7,
  fontSize: 26, bold: true, color: C.dark, fontFace: 'Arial',
});

slide.addShape('rect', { x: 0.3, y: 1.7, w: 4.3, h: 5, fill: { color: C.lgray }, line: { type: 'none' } });
slide.addText('❌ Status Quo', {
  x: 0.4, y: 1.85, w: 4.1, h: 0.3,
  fontSize: 13, bold: true, color: C.dark, fontFace: 'Arial',
});
const oldPoints2 = [
  'Email ops team: "Build me X audience"',
  '5–7 day turnaround',
  'Limited to pre-built segment library',
  'No ability to iterate or A/B variants',
];
oldPoints2.forEach((pt, i) => {
  slide.addText('• ' + pt, {
    x: 0.45, y: 2.35 + i * 0.65, w: 4, h: 0.6,
    fontSize: 11, color: C.dark, fontFace: 'Arial',
  });
});

slide.addShape('rect', { x: 5.4, y: 1.7, w: 4.3, h: 5, fill: { color: C.lgray }, line: { type: 'none' } });
slide.addText('✓ Audience Codec', {
  x: 5.5, y: 1.85, w: 4.1, h: 0.3,
  fontSize: 13, bold: true, color: C.blue, fontFace: 'Arial',
});
const newPoints2 = [
  'Describe audience in plain language',
  '8 minutes end-to-end (agent + preview)',
  'Agent recommends rules + reach forecast',
  'Save and A/B test variants instantly',
];
newPoints2.forEach((pt, i) => {
  slide.addText('• ' + pt, {
    x: 5.55, y: 2.35 + i * 0.65, w: 4, h: 0.6,
    fontSize: 11, color: C.dark, fontFace: 'Arial',
  });
});

// ===== SLIDE 4: THE MECHANIC =====
slide = prs.addSlide();
slide.background = { color: C.dark };

slide.addText('THE MECHANIC', {
  x: 0.3, y: 0.3, w: 9, h: 0.25,
  fontSize: 9, bold: true, color: C.gold, fontFace: 'Arial',
});

slide.addText('Five-Step Conversational Audience Assembly', {
  x: 0.3, y: 0.8, w: 9, h: 0.5,
  fontSize: 24, bold: true, color: C.white, fontFace: 'Arial',
});

const steps2 = [
  { n: '1', t: 'NL Parsing', d: 'Agent extracts intent' },
  { n: '2', t: 'Rule Assembly', d: 'Recommends segment logic' },
  { n: '3', t: 'Reach Forecast', d: 'CPM + conversion estimate' },
  { n: '4', t: 'Save Audience', d: 'Zero engineering required' },
  { n: '5', t: 'Deploy & A/B', d: 'Apply to campaigns instantly' },
];

steps2.forEach((step, i) => {
  const x = 0.3 + i * 1.85;
  slide.addText(step.n, {
    x, y: 1.6, w: 1.7, h: 0.4,
    fontSize: 16, bold: true, color: C.gold, fontFace: 'Arial', align: 'center',
  });
  slide.addText(step.t, {
    x, y: 2.1, w: 1.7, h: 0.3,
    fontSize: 11, bold: true, color: C.white, fontFace: 'Arial', align: 'center',
  });
  slide.addText(step.d, {
    x, y: 2.45, w: 1.7, h: 0.6,
    fontSize: 9, color: C.muted, fontFace: 'Arial', align: 'center',
  });
  if (i < steps2.length - 1) {
    slide.addShape('line', { x: x + 1.85, y: 1.8, w: 0.3, h: 0, line: { color: C.blue, width: 2, dashType: 'dash' } });
  }
});

slide.addShape('rect', { x: 0.3, y: 3.8, w: 9.2, h: 2.9, fill: { color: C.hero }, line: { color: C.blue, width: 2 }, shadow: makeShadow() });
slide.addText('PhonePe Proof Point', {
  x: 0.5, y: 3.95, w: 8.8, h: 0.2,
  fontSize: 10, bold: true, color: C.gold, fontFace: 'Arial',
});
slide.addText('Built a zero-to-one self-serve merchant acquisition platform where 5,000+ merchants assembled their own campaigns (Instant Discount + EMI subvention) without engineering overhead. TAT: 6 weeks → 2 days. Audience Codec applies the same "merchant becomes power user" pattern to advertiser audience building.', {
  x: 0.5, y: 4.25, w: 8.8, h: 2.3,
  fontSize: 11, color: C.white, fontFace: 'Arial',
});

// ===== SLIDE 5: THE PRODUCT =====
slide = prs.addSlide();
slide.background = { color: C.white };

slide.addText('THE PRODUCT', {
  x: 0.3, y: 0.3, w: 9, h: 0.25,
  fontSize: 9, bold: true, color: C.bright, fontFace: 'Arial',
});

slide.addText('Four Screens — Library to Confirmation', {
  x: 0.3, y: 0.8, w: 9, h: 0.4,
  fontSize: 22, bold: true, color: C.dark, fontFace: 'Arial',
});

const screens2 = [
  { n: '01', t: 'Library', d: 'Saved audiences + performance sparkline' },
  { n: '02', t: 'Agent Conversation', d: 'Multi-turn dialog + live reach forecast' },
  { n: '03', t: 'Preview', d: 'Segment breakdown + CPM band + save' },
  { n: '04', t: 'Saved', d: 'Ready to deploy to any campaign' },
];

screens2.forEach((screen, i) => {
  const x = 0.3 + (i % 2) * 4.8;
  const y = 1.5 + Math.floor(i / 2) * 2.9;
  slide.addShape('rect', { x, y, w: 4.5, h: 2.7, fill: { color: C.lgray }, line: { color: C.blue, width: 1 }, shadow: makeShadow() });
  slide.addText(screen.n + ' · ' + screen.t, {
    x: x + 0.2, y: y + 0.2, w: 4.1, h: 0.35,
    fontSize: 12, bold: true, color: C.dark, fontFace: 'Arial',
  });
  slide.addText(screen.d, {
    x: x + 0.2, y: y + 0.65, w: 4.1, h: 1.8,
    fontSize: 10, color: C.muted, fontFace: 'Arial',
  });
});

// ===== SLIDE 6: IMPACT & ROI =====
slide = prs.addSlide();
slide.background = { color: C.dark };

slide.addText('IMPACT & ROI', {
  x: 0.3, y: 0.3, w: 9, h: 0.25,
  fontSize: 9, bold: true, color: C.gold, fontFace: 'Arial',
});

slide.addText('Proven on PhonePe—Adjusted for Adobe Scale', {
  x: 0.3, y: 0.8, w: 9, h: 0.3,
  fontSize: 18, bold: true, color: C.white, fontFace: 'Arial',
});

slide.addText('Advertiser Impact', {
  x: 0.3, y: 1.4, w: 4.5, h: 0.25,
  fontSize: 12, bold: true, color: C.blue, fontFace: 'Arial',
});

const userMetrics2 = [
  { val: '5–7d → 8m', lbl: 'Audience TAT' },
  { val: '+60%', lbl: 'Unlock Tier-2 Targeting' },
  { val: '+22%', lbl: 'Audience Reuse Rate' },
  { val: '+16%', lbl: 'ROAS from Adv. Targeting' },
];

userMetrics2.forEach((m, i) => {
  slide.addText(m.val, {
    x: 0.3, y: 1.85 + i * 0.9, w: 2, h: 0.3,
    fontSize: 16, bold: true, color: C.blue, fontFace: 'Arial',
  });
  slide.addText(m.lbl, {
    x: 2.4, y: 1.85 + i * 0.9, w: 2.4, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: 'Arial',
  });
});

slide.addText('Adobe Platform Impact', {
  x: 5.2, y: 1.4, w: 4.5, h: 0.25,
  fontSize: 12, bold: true, color: C.gold, fontFace: 'Arial',
});

const adobeMetrics2 = [
  { val: '+28%', lbl: 'Platform Stickiness (NRR)' },
  { val: '−31%', lbl: 'Customer Churn' },
  { val: '+19%', lbl: 'Feature Adoption' },
  { val: '+$38M', lbl: 'Annual Incremental ARR' },
];

adobeMetrics2.forEach((m, i) => {
  slide.addText(m.val, {
    x: 5.2, y: 1.85 + i * 0.9, w: 2, h: 0.3,
    fontSize: 16, bold: true, color: C.gold, fontFace: 'Arial',
  });
  slide.addText(m.lbl, {
    x: 7.3, y: 1.85 + i * 0.9, w: 2.4, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: 'Arial',
  });
});

slide.addShape('rect', { x: 0.3, y: 5.4, w: 9.2, h: 1.7, fill: { color: C.hero }, line: { color: C.gold, width: 2 }, shadow: makeShadow() });
slide.addText('Key Insight: Codec unlocks TAT reduction (5–7 days → 8 mins), freeing ops team to focus on strategic audience strategy instead of manual rule building. Lower CAC per activated advertiser.', {
  x: 0.5, y: 5.55, w: 8.8, h: 1.4,
  fontSize: 11, color: C.white, fontFace: 'Arial',
});

// ===== SLIDE 7: PROOF OF WORK =====
slide = prs.addSlide();
slide.background = { color: C.white };

slide.addText('PROOF OF WORK', {
  x: 0.3, y: 0.3, w: 9, h: 0.25,
  fontSize: 9, bold: true, color: C.bright, fontFace: 'Arial',
});

slide.addText('I Built This at PhonePe. Here\'s the Map.', {
  x: 0.3, y: 0.8, w: 9, h: 0.5,
  fontSize: 20, bold: true, color: C.dark, fontFace: 'Arial',
});

slide.addShape('rect', { x: 0.3, y: 1.5, w: 4.5, h: 5.5, fill: { color: C.dark }, line: { type: 'none' }, shadow: makeShadow() });
slide.addText('PhonePe (APM, Apr 2022–Apr 2023)', {
  x: 0.4, y: 1.65, w: 4.3, h: 0.25,
  fontSize: 11, bold: true, color: C.blue, fontFace: 'Arial',
});

const phonePeWork2 = [
  'Self-serve PG integration (5K merchants, 0 eng)',
  'TAT: 6 weeks → 2 days via wizard + smart defaults',
  'Merchant power-user in <30 min with no ops',
  'Zero-to-one merchant acquisition platform',
  'Reward platform with brand self-serve tooling',
];

phonePeWork2.forEach((w, i) => {
  slide.addText('• ' + w, {
    x: 0.45, y: 2.05 + i * 0.8, w: 4.2, h: 0.75,
    fontSize: 10, color: C.white, fontFace: 'Arial',
  });
});

slide.addShape('rect', { x: 5.2, y: 1.5, w: 4.5, h: 5.5, fill: { color: C.lgray }, line: { type: 'none' }, shadow: makeShadow() });
slide.addText('→ Maps to Audience Codec', {
  x: 5.3, y: 1.65, w: 4.3, h: 0.25,
  fontSize: 11, bold: true, color: C.blue, fontFace: 'Arial',
});

const mappings2 = [
  'Conversational onboarding (merchant → advertiser)',
  'Frictionless TAT (2 days → 8 mins)',
  'Power-user enabling (brand → advertiser)',
  'Zero-dependency platform scaling',
  'Self-serve + agent collaboration (same arch)',
];

mappings2.forEach((m, i) => {
  slide.addText('• ' + m, {
    x: 5.35, y: 2.05 + i * 0.8, w: 4.2, h: 0.75,
    fontSize: 10, color: C.dark, fontFace: 'Arial',
  });
});

// ===== SLIDE 8: ROLLOUT PLAN =====
slide = prs.addSlide();
slide.background = { color: C.dark };

slide.addText('ROLLOUT PLAN', {
  x: 0.3, y: 0.3, w: 9, h: 0.25,
  fontSize: 9, bold: true, color: C.gold, fontFace: 'Arial',
});

slide.addText('Three Phases: Alpha to Tier-2 Unlock', {
  x: 0.3, y: 0.8, w: 9, h: 0.3,
  fontSize: 20, bold: true, color: C.white, fontFace: 'Arial',
});

const phases2 = [
  { t: 'Month 1–2: Closed Beta', desc: '30 power users, teach agent patterns, gather feedback, tune NL understanding.' },
  { t: 'Month 3–4: Public Beta', desc: '500 advertisers, measure TAT reduction + ROAS lift, scale NL model.' },
  { t: 'Month 5–6: GA + Suite Upsell', desc: 'Platform feature. Bundle with Campaign Autopilot. Target 40% opt-in within 90 days.' },
];

phases2.forEach((phase, i) => {
  const x = 0.3 + i * 3.2;
  slide.addShape('rect', { x, y: 1.4, w: 3, h: 4.5, fill: { color: C.hero }, line: { color: C.blue, width: 2 }, shadow: makeShadow() });
  slide.addText(phase.t, {
    x: x + 0.1, y: 1.6, w: 2.8, h: 0.35,
    fontSize: 12, bold: true, color: C.blue, fontFace: 'Arial', align: 'center',
  });
  slide.addText(phase.desc, {
    x: x + 0.1, y: 2.1, w: 2.8, h: 3.5,
    fontSize: 10, color: C.white, fontFace: 'Arial', align: 'left',
  });
});

slide.addShape('rect', { x: 0.3, y: 6.2, w: 9.2, h: 1, fill: { color: C.hero }, line: { color: C.gold, width: 2 } });
slide.addText('What I need: NL/Semantic understanding API · Audience rules engine access · 1 ML/NLP Engineer + 1 FE + 1 Designer · 7-week sprint', {
  x: 0.45, y: 6.35, w: 9, h: 0.7,
  fontSize: 11, color: C.white, fontFace: 'Arial',
});

slide.addText('Ajay Avaghade · +91 9561558439 · avaghadeajay009@gmail.com', {
  x: 0.3, y: 7.1, w: 9, h: 0.3,
  fontSize: 9, color: C.muted, fontFace: 'Arial', align: 'center',
});

prs.writeFile({ fileName: 'assets/adobe-codec-deck.pptx' }).then(()=>console.log('✓ adobe-codec-deck.pptx built'));
