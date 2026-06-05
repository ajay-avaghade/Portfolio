const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '1F1F1F',
  hero:   '2A2A2A',
  bright: 'E57CD8',
  orange: 'FF9B82',
  white:  'FFFFFF',
  lgray:  'F1F5F9',
  muted:  '8B92A5',
};

const makeShadow = () => ({
  type: 'outer', color: '000000', blur: 3, offset: 1, angle: 45, opacity: 0.12
});

const defaultText = { fontFace: 'Helvetica', color: C.white };

// SLIDE 1: Cover
let slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: '25%', h: '100%', fill: { color: C.bright } });
slide.addText('TOGGL FOCUS', { x: 0.5, y: 0.5, w: 2.5, fontSize: 14, color: C.dark, bold: true, align: 'center' });
slide.addText('CapacityPulse Engine', { x: 3, y: 2.5, w: 6, fontSize: 32, bold: true, color: C.white });
slide.addText('Real-time team utilization & burnout risk prediction', { x: 3, y: 3.2, w: 6, fontSize: 16, color: C.muted });
slide.addText('By Ajay Avaghade', { x: 3, y: 6.5, w: 4, fontSize: 12, color: C.muted });

// SLIDE 2: The Problem
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('THE PROBLEM', { x: 0.5, y: 0.5, w: 9, fontSize: 10, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Managers lack real-time visibility into burnout.', { x: 0.5, y: 1, w: 9, fontSize: 28, bold: true, color: C.white });

slide.addText('42%', { x: 0.5, y: 3, w: 2.5, fontSize: 40, bold: true, color: C.bright });
slide.addText('of SaaS employee churn is linked to sustained burnout.', { x: 0.5, y: 4, w: 2.5, fontSize: 14, color: C.muted });

slide.addText('Lagging', { x: 3.5, y: 3, w: 2.5, fontSize: 40, bold: true, color: C.bright });
slide.addText('Current capacity tools rely on lagging HR data or pulse surveys.', { x: 3.5, y: 4, w: 2.5, fontSize: 14, color: C.muted });

slide.addText('Data ROI', { x: 6.5, y: 3, w: 2.5, fontSize: 40, bold: true, color: C.bright });
slide.addText('Teams do not trust time tracking if it does not protect them.', { x: 6.5, y: 4, w: 2.5, fontSize: 14, color: C.muted });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 5.5, w: 9, h: 1.2, fill: { color: C.hero }, rectRadius: 0.1 });
slide.addText('Insight: Transforming time tracking from a passive reporting tool into an active manager intelligence layer creates a "protect the team" use case for Toggl Focus.', { x: 0.8, y: 5.6, w: 8.4, h: 1, fontSize: 14, color: C.white });

// SLIDE 3: The Insight
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addText('THE INSIGHT', { x: 0.5, y: 0.5, w: 9, fontSize: 10, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('From Reporting to Intervention', { x: 0.5, y: 1, w: 9, fontSize: 28, bold: true, color: C.dark });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2.5, w: 4, h: 4, fill: { color: C.white }, shadow: makeShadow(), rectRadius: 0.1 });
slide.addText('❌ Legacy Time Tracking', { x: 1, y: 3, w: 3, fontSize: 16, bold: true, color: C.dark });
slide.addText('• Backward-looking data\n• Used primarily for billing/payroll\n• Manual capacity calculation\n• Employees feel surveilled', { x: 1, y: 3.5, w: 3, fontSize: 13, color: C.muted, bullet: true, lineSpacing: 24 });

slide.addShape(prs.ShapeType.rect, { x: 5.5, y: 2.5, w: 4, h: 4, fill: { color: C.dark }, shadow: makeShadow(), rectRadius: 0.1 });
slide.addText('✅ Toggl CapacityPulse', { x: 6, y: 3, w: 3, fontSize: 16, bold: true, color: C.bright });
slide.addText('• Forward-looking predictive ML\n• Identifies context-switching strain\n• 1-click workload redistribution\n• Employees feel protected', { x: 6, y: 3.5, w: 3, fontSize: 13, color: C.lgray, bullet: true, lineSpacing: 24 });

slide.addShape(prs.ShapeType.ellipse, { x: 4.6, y: 4.1, w: 0.8, h: 0.8, fill: { color: C.bright }, shadow: makeShadow() });
slide.addText('VS', { x: 4.6, y: 4.1, w: 0.8, h: 0.8, fontSize: 12, bold: true, color: C.dark, align: 'center' });

// SLIDE 4: The Mechanic
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('THE MECHANIC', { x: 0.5, y: 0.5, w: 9, fontSize: 10, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('How CapacityPulse Intervenes', { x: 0.5, y: 1, w: 9, fontSize: 28, bold: true, color: C.white });

const steps = ['Track Time', 'AI Analysis', 'Risk Flagged', 'Reassign', 'Prevent Burnout'];
slide.addShape(prs.ShapeType.rect, { x: 1, y: 4, w: 8, h: 0.05, fill: { color: C.muted } });

steps.forEach((step, i) => {
  slide.addShape(prs.ShapeType.ellipse, { x: 0.8 + i*2, y: 3.8, w: 0.4, h: 0.4, fill: { color: C.bright } });
  slide.addText(String(i+1), { x: 0.8 + i*2, y: 3.8, w: 0.4, h: 0.4, fontSize: 12, bold: true, color: C.dark, align: 'center' });
  slide.addText(step, { x: 0 + i*2, y: 4.5, w: 2, fontSize: 14, bold: true, color: C.white, align: 'center' });
});

// SLIDE 5: The Product
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addText('THE PRODUCT', { x: 0.5, y: 0.5, w: 9, fontSize: 10, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Built for the Manager Persona', { x: 0.5, y: 1, w: 9, fontSize: 28, bold: true, color: C.dark });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2, w: 4.3, h: 2.2, fill: { color: C.white }, shadow: makeShadow(), rectRadius: 0.1 });
slide.addText('Real-time Risk Dashboard', { x: 0.8, y: 2.3, w: 3.7, fontSize: 14, bold: true, color: C.dark });
slide.addText('Aggregates team velocity and highlights critical overload instances instantly.', { x: 0.8, y: 2.7, w: 3.7, fontSize: 12, color: C.muted });

slide.addShape(prs.ShapeType.rect, { x: 5.2, y: 2, w: 4.3, h: 2.2, fill: { color: C.white }, shadow: makeShadow(), rectRadius: 0.1 });
slide.addText('Contextual AI Analysis', { x: 5.5, y: 2.3, w: 3.7, fontSize: 14, bold: true, color: C.dark });
slide.addText('Analyzes non-obvious strain metrics like context-switching and unplanned bug-fix tags.', { x: 5.5, y: 2.7, w: 3.7, fontSize: 12, color: C.muted });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 4.5, w: 4.3, h: 2.2, fill: { color: C.white }, shadow: makeShadow(), rectRadius: 0.1 });
slide.addText('1-Click Reassignment', { x: 0.8, y: 4.8, w: 3.7, fontSize: 14, bold: true, color: C.dark });
slide.addText('Surfaces available capacity across the team to effortlessly redistribute ticket volume.', { x: 0.8, y: 5.2, w: 3.7, fontSize: 12, color: C.muted });

slide.addShape(prs.ShapeType.rect, { x: 5.2, y: 4.5, w: 4.3, h: 2.2, fill: { color: C.white }, shadow: makeShadow(), rectRadius: 0.1 });
slide.addText('Forward Projections', { x: 5.5, y: 4.8, w: 3.7, fontSize: 14, bold: true, color: C.dark });
slide.addText('Projects future burnout risk, empowering PMs to push back on unrealistic roadmap scopes.', { x: 5.5, y: 5.2, w: 3.7, fontSize: 12, color: C.muted });

// SLIDE 6: Impact & ROI
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('IMPACT & ROI', { x: 0.5, y: 0.5, w: 9, fontSize: 10, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Strategic Value to Toggl Focus', { x: 0.5, y: 1, w: 9, fontSize: 28, bold: true, color: C.white });

const metrics = [
  { val: '+18%', lbl: 'Expansion MRR', d: 'Drives seat expansion for management users.' },
  { val: '+2.5x', lbl: 'Platform Stickiness', d: 'Deeply embeds into core HR & Ops workflows.' },
  { val: '-30%', lbl: 'SaaS Churn', d: 'Managers rely on it as a critical retention tool.' },
  { val: 'Top 3', lbl: 'Differentiation', d: 'Unique moat vs basic time-tracking clones.' }
];

metrics.forEach((m, i) => {
  let cx = 0.5 + (i%2)*4.5;
  let cy = 2.5 + Math.floor(i/2)*2.2;
  slide.addShape(prs.ShapeType.rect, { x: cx, y: cy, w: 4.2, h: 1.8, fill: { color: C.hero }, rectRadius: 0.1 });
  slide.addText(m.val, { x: cx+0.2, y: cy+0.3, w: 1.5, fontSize: 24, bold: true, color: C.bright });
  slide.addText(m.lbl, { x: cx+1.8, y: cy+0.3, w: 2.2, fontSize: 12, bold: true, color: C.white });
  slide.addText(m.d, { x: cx+0.2, y: cy+1, w: 3.8, fontSize: 11, color: C.muted });
});

// SLIDE 7: Proof of Work
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 5, y: 0, w: 5, h: 7.5, fill: { color: C.lgray } });

slide.addText('MY EXPERIENCE', { x: 0.5, y: 0.5, w: 4, fontSize: 10, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('PhonePe ML Scale', { x: 0.5, y: 1, w: 4, fontSize: 24, bold: true, color: C.white });
slide.addText('• Led ML-driven Propensity-to-Transact modeling for 350M+ users.\n• Abstracted complex ML logic into clear, actionable dashboards.\n• Optimized checkout workflows to drive +22% conversion lift.', { x: 0.5, y: 2, w: 4, fontSize: 14, color: C.muted, bullet: true, lineSpacing: 24 });

slide.addText('TOGGL ALIGNMENT', { x: 5.5, y: 0.5, w: 4, fontSize: 10, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Why I fit this role', { x: 5.5, y: 1, w: 4, fontSize: 24, bold: true, color: C.dark });
slide.addText('• Proven ability to leverage data & ML for intelligent products.\n• Embedded highly with engineering and design to execute complex architectures.\n• Deeply problem-first, obsessed with uncovering user utility.', { x: 5.5, y: 2, w: 4, fontSize: 14, color: C.dark, bullet: true, lineSpacing: 24 });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 5.5, w: 9, h: 1.5, fill: { color: C.bright }, rectRadius: 0.1 });
slide.addText('"The best PMs do not just organize tickets. They use ML and data logic to fundamentally shift how users perceive value and scale product-led growth."', { x: 1, y: 5.6, w: 8, h: 1.3, fontSize: 16, bold: true, color: C.dark, align: 'center', italic: true });

// SLIDE 8: Rollout Plan
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('ROLLOUT PLAN', { x: 0.5, y: 0.5, w: 9, fontSize: 10, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Execution Milestones', { x: 0.5, y: 1, w: 9, fontSize: 28, bold: true, color: C.white });

const phases = [
  { t: 'Phase 1: Discovery', d: 'Analyze historical time-tracking data for predictive burn signatures.' },
  { t: 'Phase 2: Alpha', d: 'Internal dogfooding with engineering managers to tune ML risk thresholds.' },
  { t: 'Phase 3: Scale', d: 'Launch CapacityPulse tier as a premium add-on for Enterprise customers.' }
];

phases.forEach((p, i) => {
  slide.addShape(prs.ShapeType.rect, { x: 0.5 + i*3.1, y: 2.5, w: 2.8, h: 2, fill: { color: C.hero }, rectRadius: 0.1, line: { color: C.bright, width: 2 } });
  slide.addText(p.t, { x: 0.7 + i*3.1, y: 2.8, w: 2.4, fontSize: 14, bold: true, color: C.white });
  slide.addText(p.d, { x: 0.7 + i*3.1, y: 3.5, w: 2.4, fontSize: 12, color: C.muted });
});

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 5, w: 9, h: 1.2, fill: { color: C.hero }, rectRadius: 0.1 });
slide.addText('What I Need: A 20-minute chat to discuss Toggl Focus AI aspirations.', { x: 0.8, y: 5.3, w: 8.4, fontSize: 14, color: C.bright, bold: true });
slide.addText('Ajay Avaghade | avaghadeajay009@gmail.com', { x: 0.5, y: 6.8, w: 9, fontSize: 11, color: C.muted, align: 'center' });

prs.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/toggl-capacitypulse-deck.pptx' }).then(() => {
  console.log('CapacityPulse Deck generated!');
});
