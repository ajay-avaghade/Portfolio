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

slide.addShape('rect', { x: 8.5, y: 0, w: 1.5, h: 7.5, fill: { color: C.bright, transparency: 15 }, line: { type: 'none' } });
slide.addShape('rect', { x: 7.8, y: 0, w: 2.2, h: 7.5, fill: { color: C.gold, transparency: 22 }, line: { type: 'none' } });

slide.addText('ADOBE DSP · PRODUCT CASE STUDY', {
  x: 0.3, y: 0.4, w: 4, h: 0.3,
  fontSize: 9, bold: true, color: C.gold, fontFace: 'Arial',
});

slide.addText('PaceIQ', {
  x: 0.3, y: 1.8, w: 7, h: 1.2,
  fontSize: 44, bold: true, color: C.white, fontFace: 'Arial',
});

slide.addText('Real-Time Budget Pacing & Impression Diagnostics', {
  x: 0.3, y: 3, w: 7, h: 0.4,
  fontSize: 18, color: C.gold, fontFace: 'Arial',
});

slide.addText('Ajay Avaghade · Product Manager', {
  x: 0.3, y: 3.6, w: 7, h: 0.3,
  fontSize: 12, color: C.muted, fontFace: 'Arial',
});

slide.addShape('rect', { x: 6.5, y: 5.8, w: 3, h: 1.3, fill: { color: C.hero }, line: { type: 'none' } });
slide.addText('+19% Budget Recovery', {
  x: 6.5, y: 5.95, w: 3, h: 0.4,
  fontSize: 32, bold: true, color: C.gold, fontFace: 'Arial', align: 'center',
});
slide.addText('From lost impression diagnostics', {
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

slide.addText('27% of Daily Budget Evaporates Invisibly—Advertisers Have No Idea Why', {
  x: 0.3, y: 0.8, w: 9, h: 0.8,
  fontSize: 26, bold: true, color: C.white, fontFace: 'Arial',
});

const statCols3 = [
  { val: '27%', label: 'Unspent / Day', desc: 'No transparency into why' },
  { val: '42%', label: 'Blind Budget Waste', desc: 'vs diagnosed spend' },
  { val: '−27%', label: 'Churn to Competitors', desc: 'From perceived DSP failure' },
];

statCols3.forEach((stat, i) => {
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
slide.addText('When a campaign underspends, advertisers see a flat "Budget Used: 73% of $10K" with no context. Hidden losses: auction loss (didn\'t win price), frequency cap (saw too many times), budget constraint (hit daily limit early), low intent (below quality threshold). Without diagnostics, advertisers assume the DSP is broken and move budget elsewhere.', {
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

slide.addText('Transparency Builds Confidence. Diagnostics Drive Optimization.', {
  x: 0.3, y: 0.8, w: 9, h: 0.7,
  fontSize: 26, bold: true, color: C.dark, fontFace: 'Arial',
});

slide.addShape('rect', { x: 0.3, y: 1.7, w: 4.3, h: 5, fill: { color: C.lgray }, line: { type: 'none' } });
slide.addText('❌ Status Quo', {
  x: 0.4, y: 1.85, w: 4.1, h: 0.3,
  fontSize: 13, bold: true, color: C.dark, fontFace: 'Arial',
});
const oldPoints3 = [
  '"Budget Used: 73%"',
  'Black box — no root cause',
  'Assume DSP is broken',
  'Move budget to competitor',
];
oldPoints3.forEach((pt, i) => {
  slide.addText('• ' + pt, {
    x: 0.45, y: 2.35 + i * 0.65, w: 4, h: 0.6,
    fontSize: 11, color: C.dark, fontFace: 'Arial',
  });
});

slide.addShape('rect', { x: 5.4, y: 1.7, w: 4.3, h: 5, fill: { color: C.lgray }, line: { type: 'none' } });
slide.addText('✓ PaceIQ', {
  x: 5.5, y: 1.85, w: 4.1, h: 0.3,
  fontSize: 13, bold: true, color: C.gold, fontFace: 'Arial',
});
const newPoints3 = [
  '24% auction loss, 31% freq cap, ...',
  'Categorized loss breakdown',
  'Real-time agent recommendations',
  'Advertiser adjusts and profits',
];
newPoints3.forEach((pt, i) => {
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

slide.addText('Five-Step Real-Time Diagnostics & Nudge Loop', {
  x: 0.3, y: 0.8, w: 9, h: 0.5,
  fontSize: 24, bold: true, color: C.white, fontFace: 'Arial',
});

const steps3 = [
  { n: '1', t: 'Budget Tracking', d: 'Impression-by-impression spend' },
  { n: '2', t: 'Loss Attribution', d: 'Categorize unspent (auction/freq/cap)' },
  { n: '3', t: 'Intent Monitoring', d: 'Heatmap vs peer baseline' },
  { n: '4', t: 'Agent Diagnostics', d: 'Top 3 issues + recommended fixes' },
  { n: '5', t: 'One-Tap Nudge', d: 'Adjust bid/cap without leaving dashboard' },
];

steps3.forEach((step, i) => {
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
  if (i < steps3.length - 1) {
    slide.addShape('line', { x: x + 1.85, y: 1.8, w: 0.3, h: 0, line: { color: C.gold, width: 2, dashType: 'dash' } });
  }
});

slide.addShape('rect', { x: 0.3, y: 3.8, w: 9.2, h: 2.9, fill: { color: C.hero }, line: { color: C.blue, width: 2 }, shadow: makeShadow() });
slide.addText('PhonePe Proof Point', {
  x: 0.5, y: 3.95, w: 8.8, h: 0.2,
  fontSize: 10, bold: true, color: C.gold, fontFace: 'Arial',
});
slide.addText('Rebuilt 6 siloed payment ops workflows into 1 unified ops intelligence layer with auto-categorization and routing. Result: TAT 2 days → 4 hours, −68% escalation tickets. Freed ops capacity from firefighting to strategy. PaceIQ brings the same ops intelligence to DSP impression diagnostics.', {
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

slide.addText('Four Screens — Dashboard to Ops Console', {
  x: 0.3, y: 0.8, w: 9, h: 0.4,
  fontSize: 22, bold: true, color: C.dark, fontFace: 'Arial',
});

const screens3 = [
  { n: '01', t: 'Dashboard', d: 'Health score + budget pacing bar' },
  { n: '02', t: 'Diagnostics', d: 'Spend breakdown by loss reason + heatmap' },
  { n: '03', t: 'Quality Deep-Dive', d: '7-day intent trend + recommendations' },
  { n: '04', t: 'Ops Console', d: 'All campaigns health grid + bulk action' },
];

screens3.forEach((screen, i) => {
  const x = 0.3 + (i % 2) * 4.8;
  const y = 1.5 + Math.floor(i / 2) * 2.9;
  slide.addShape('rect', { x, y, w: 4.5, h: 2.7, fill: { color: C.lgray }, line: { color: C.gold, width: 1 }, shadow: makeShadow() });
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
  fontSize: 12, bold: true, color: C.gold, fontFace: 'Arial',
});

const userMetrics3 = [
  { val: '+19%', lbl: 'Budget Recovery (via diagnostics)' },
  { val: '−42%', lbl: 'Blind Spend Waste' },
  { val: '−31%', lbl: 'Troubleshooting Time' },
  { val: '+$1,890', lbl: 'Monthly GMV recovered' },
];

userMetrics3.forEach((m, i) => {
  slide.addText(m.val, {
    x: 0.3, y: 1.85 + i * 0.9, w: 2, h: 0.3,
    fontSize: 16, bold: true, color: C.gold, fontFace: 'Arial',
  });
  slide.addText(m.lbl, {
    x: 2.4, y: 1.85 + i * 0.9, w: 2.4, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: 'Arial',
  });
});

slide.addText('Adobe Platform Impact', {
  x: 5.2, y: 1.4, w: 4.5, h: 0.25,
  fontSize: 12, bold: true, color: C.bright, fontFace: 'Arial',
});

const adobeMetrics3 = [
  { val: '−27%', lbl: 'Churn to Competitors' },
  { val: '+34%', lbl: 'NRR (from transparency feature)' },
  { val: '+12%', lbl: 'Feature Adoption' },
  { val: '+$52M', lbl: 'Annual Incremental Revenue' },
];

adobeMetrics3.forEach((m, i) => {
  slide.addText(m.val, {
    x: 5.2, y: 1.85 + i * 0.9, w: 2, h: 0.3,
    fontSize: 16, bold: true, color: C.bright, fontFace: 'Arial',
  });
  slide.addText(m.lbl, {
    x: 7.3, y: 1.85 + i * 0.9, w: 2.4, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: 'Arial',
  });
});

slide.addShape('rect', { x: 0.3, y: 5.4, w: 9.2, h: 1.7, fill: { color: C.hero }, line: { color: C.gold, width: 2 }, shadow: makeShadow() });
slide.addText('Key Insight: Transparency is the strongest retention lever. Advertisers who understand *why* they\'re underspending stay; those who don\'t, leave. PaceIQ converts ad spend uncertainty into actionable optimization opportunities.', {
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
slide.addText('PhonePe (PM & APM, 2022–Now)', {
  x: 0.4, y: 1.65, w: 4.3, h: 0.25,
  fontSize: 11, bold: true, color: C.blue, fontFace: 'Arial',
});

const phonePeWork3 = [
  '6 siloed ops workflows unified into 1 dashboard',
  'TAT: 2 days → 4 hours via auto-categorization',
  '−68% escalation tickets (via diagnosis)',
  'Freed ops to strategic work instead of fire-fighting',
  'Real-time categorization of payment failures',
];

phonePeWork3.forEach((w, i) => {
  slide.addText('• ' + w, {
    x: 0.45, y: 2.05 + i * 0.8, w: 4.2, h: 0.75,
    fontSize: 10, color: C.white, fontFace: 'Arial',
  });
});

slide.addShape('rect', { x: 5.2, y: 1.5, w: 4.5, h: 5.5, fill: { color: C.lgray }, line: { type: 'none' }, shadow: makeShadow() });
slide.addText('→ Maps to PaceIQ', {
  x: 5.3, y: 1.65, w: 4.3, h: 0.25,
  fontSize: 11, bold: true, color: C.blue, fontFace: 'Arial',
});

const mappings3 = [
  'Unified impression loss categorization',
  'Real-time diagnostics (impression-level)',
  'Auto-categorization (auction/freq/cap/intent)',
  'Frees media team from manual debugging',
  'Impression-level failure taxonomy',
];

mappings3.forEach((m, i) => {
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

slide.addText('Three Phases: Diagnostics to Predictive', {
  x: 0.3, y: 0.8, w: 9, h: 0.3,
  fontSize: 20, bold: true, color: C.white, fontFace: 'Arial',
});

const phases3 = [
  { t: 'Month 1–2: Closed Beta', desc: '30 campaigns, validate loss categorization accuracy, build confidence.' },
  { t: 'Month 3–4: Public Beta', desc: 'Scale to 300 campaigns, measure advertiser confidence lift + churn reduction.' },
  { t: 'Month 5–6: GA + Predictive', desc: 'Predict pacing shortfalls 24h ahead. Auto-recommend bid adjustments. Platform default.' },
];

phases3.forEach((phase, i) => {
  const x = 0.3 + i * 3.2;
  slide.addShape('rect', { x, y: 1.4, w: 3, h: 4.5, fill: { color: C.hero }, line: { color: C.gold, width: 2 }, shadow: makeShadow() });
  slide.addText(phase.t, {
    x: x + 0.1, y: 1.6, w: 2.8, h: 0.35,
    fontSize: 12, bold: true, color: C.gold, fontFace: 'Arial', align: 'center',
  });
  slide.addText(phase.desc, {
    x: x + 0.1, y: 2.1, w: 2.8, h: 3.5,
    fontSize: 10, color: C.white, fontFace: 'Arial', align: 'left',
  });
});

slide.addShape('rect', { x: 0.3, y: 6.2, w: 9.2, h: 1, fill: { color: C.hero }, line: { color: C.gold, width: 2 } });
slide.addText('What I need: Impression event log access (auction loss signals) · Ops intelligence infrastructure · 1 Data Engineer + 1 Analyst + 1 FE + 1 Designer · 6-week sprint', {
  x: 0.45, y: 6.35, w: 9, h: 0.7,
  fontSize: 11, color: C.white, fontFace: 'Arial',
});

slide.addText('Ajay Avaghade · +91 9561558439 · avaghadeajay009@gmail.com', {
  x: 0.3, y: 7.1, w: 9, h: 0.3,
  fontSize: 9, color: C.muted, fontFace: 'Arial', align: 'center',
});

prs.writeFile({ fileName: 'assets/adobe-paceiq-deck.pptx' }).then(()=>console.log('✓ adobe-paceiq-deck.pptx built'));
