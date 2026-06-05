const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '1F1F1F',
  hero:   '2A2A2A',
  bright: '10B981',
  orange: 'F59E0B',
  white:  'FFFFFF',
  lgray:  'F1F5F9',
  muted:  '8B92A5',
};

const makeShadow = () => ({
  type: 'outer', color: '000000', blur: 3, offset: 1, angle: 45, opacity: 0.12
});

// SLIDE 1: Cover
let slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: '25%', h: '100%', fill: { color: C.bright } });
slide.addText('TOGGL FOCUS', { x: 0.5, y: 0.5, w: 2.5, fontSize: 14, color: C.dark, bold: true, align: 'center' });
slide.addText('SmartScope AI', { x: 3, y: 2.5, w: 6, fontSize: 32, bold: true, color: C.white });
slide.addText('Predictive duration scoping using historical time data', { x: 3, y: 3.2, w: 6, fontSize: 16, color: C.muted });
slide.addText('By Ajay Avaghade', { x: 3, y: 6.5, w: 4, fontSize: 12, color: C.muted });

// SLIDE 2: The Problem
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('THE PROBLEM', { x: 0.5, y: 0.5, w: 9, fontSize: 10, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Project planning is based on gut feeling, not data.', { x: 0.5, y: 1, w: 9, fontSize: 28, bold: true, color: C.white });

slide.addText('68%', { x: 0.5, y: 3, w: 2.5, fontSize: 40, bold: true, color: C.bright });
slide.addText('of IT projects miss their original delivery estimate.', { x: 0.5, y: 4, w: 2.5, fontSize: 14, color: C.muted });

slide.addText('Black Box', { x: 3.5, y: 3, w: 2.5, fontSize: 40, bold: true, color: C.bright });
slide.addText('Teams do not use their own historical data to plan better.', { x: 3.5, y: 4, w: 2.5, fontSize: 14, color: C.muted });

slide.addText('Bias', { x: 6.5, y: 3, w: 2.5, fontSize: 40, bold: true, color: C.bright });
slide.addText('Humans are inherently optimistic when estimating complex tasks.', { x: 6.5, y: 4, w: 2.5, fontSize: 14, color: C.muted });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 5.5, w: 9, h: 1.2, fill: { color: C.hero }, rectRadius: 0.1 });
slide.addText('Insight: Time tracking is only valuable if it improves future performance. SmartScope turns passive logs into active estimation intelligence.', { x: 0.8, y: 5.6, w: 8.4, h: 1, fontSize: 14, color: C.white });

// SLIDE 3: The Insight
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addText('THE INSIGHT', { x: 0.5, y: 0.5, w: 9, fontSize: 10, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Data-Driven Realism', { x: 0.5, y: 1, w: 9, fontSize: 28, bold: true, color: C.dark });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2.5, w: 4, h: 4, fill: { color: C.white }, shadow: makeShadow(), rectRadius: 0.1 });
slide.addText('❌ Manual Estimation', { x: 1, y: 3, w: 3, fontSize: 16, bold: true, color: C.dark });
slide.addText('• Relies on PM gut feeling\n• Optimism bias uncorrected\n• Missed deadlines\n• Data goes unused', { x: 1, y: 3.5, w: 3, fontSize: 13, color: C.muted, bullet: true, lineSpacing: 24 });

slide.addShape(prs.ShapeType.rect, { x: 5.5, y: 2.5, w: 4, h: 4, fill: { color: C.dark }, shadow: makeShadow(), rectRadius: 0.1 });
slide.addText('✅ Toggl SmartScope', { x: 6, y: 3, w: 3, fontSize: 16, bold: true, color: C.bright });
slide.addText('• AI matches similar projects\n• Applies realism buffers\n• Confidence scoring\n• Pacing alerts', { x: 6, y: 3.5, w: 3, fontSize: 13, color: C.lgray, bullet: true, lineSpacing: 24 });

slide.addShape(prs.ShapeType.ellipse, { x: 4.6, y: 4.1, w: 0.8, h: 0.8, fill: { color: C.bright }, shadow: makeShadow() });
slide.addText('VS', { x: 4.6, y: 4.1, w: 0.8, h: 0.8, fontSize: 12, bold: true, color: C.dark, align: 'center' });

// SLIDE 4: The Mechanic
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('THE MECHANIC', { x: 0.5, y: 0.5, w: 9, fontSize: 10, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('How SmartScope Learns', { x: 0.5, y: 1, w: 9, fontSize: 28, bold: true, color: C.white });

const steps = ['Input Tags', 'Data Match', 'AI Buffer', 'Lock Scope', 'Track Pacing'];
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
slide.addText('AI Built for Planners', { x: 0.5, y: 1, w: 9, fontSize: 28, bold: true, color: C.dark });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2, w: 4.3, h: 2.2, fill: { color: C.white }, shadow: makeShadow(), rectRadius: 0.1 });
slide.addText('Parameter Input', { x: 0.8, y: 2.3, w: 3.7, fontSize: 14, bold: true, color: C.dark });
slide.addText('Uses standard project tags to semantically match against thousands of historical time entries.', { x: 0.8, y: 2.7, w: 3.7, fontSize: 12, color: C.muted });

slide.addShape(prs.ShapeType.rect, { x: 5.2, y: 2, w: 4.3, h: 2.2, fill: { color: C.white }, shadow: makeShadow(), rectRadius: 0.1 });
slide.addText('Confidence Scoring', { x: 5.5, y: 2.3, w: 3.7, fontSize: 14, bold: true, color: C.dark });
slide.addText('Outputs a raw estimate plus an AI-applied "realism buffer" with a % confidence score.', { x: 5.5, y: 2.7, w: 3.7, fontSize: 12, color: C.muted });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 4.5, w: 4.3, h: 2.2, fill: { color: C.white }, shadow: makeShadow(), rectRadius: 0.1 });
slide.addText('Transparent Distribution', { x: 0.8, y: 4.8, w: 3.7, fontSize: 14, bold: true, color: C.dark });
slide.addText('Shows the raw historical spread chart to build PM trust in the AI recommendation.', { x: 0.8, y: 5.2, w: 3.7, fontSize: 12, color: C.muted });

slide.addShape(prs.ShapeType.rect, { x: 5.2, y: 4.5, w: 4.3, h: 2.2, fill: { color: C.white }, shadow: makeShadow(), rectRadius: 0.1 });
slide.addText('Predictive Tracking', { x: 5.5, y: 4.8, w: 3.7, fontSize: 14, bold: true, color: C.dark });
slide.addText('Compares real-time burn against historical pacing, alerting PMs if early consumption spikes.', { x: 5.5, y: 5.2, w: 3.7, fontSize: 12, color: C.muted });

// SLIDE 6: Impact & ROI
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('IMPACT & ROI', { x: 0.5, y: 0.5, w: 9, fontSize: 10, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Strategic Value to Toggl Focus', { x: 0.5, y: 1, w: 9, fontSize: 28, bold: true, color: C.white });

const metrics = [
  { val: '+45%', lbl: 'Estimate Accuracy', d: 'Significant drop in delayed deliverables.' },
  { val: '2.8x', lbl: 'Value Perception', d: 'Users actively use Toggl to plan, not just log.' },
  { val: '-20%', lbl: 'Churn Rate', d: 'Indispensable tool for enterprise PMOs.' },
  { val: 'AI', lbl: 'Category Leader', d: 'Proves the "Time Intelligence" mission statement.' }
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
slide.addText('PhonePe Ops Data', { x: 0.5, y: 1, w: 4, fontSize: 24, bold: true, color: C.white });
slide.addText('• Reduced offer operations TAT from 2 days to 30 mins.\n• Used complex data ingestion to build self-serve merchant toolsets.\n• Focused entirely on leveraging existing data to optimize workflows.', { x: 0.5, y: 2, w: 4, fontSize: 14, color: C.muted, bullet: true, lineSpacing: 24 });

slide.addText('TOGGL ALIGNMENT', { x: 5.5, y: 0.5, w: 4, fontSize: 10, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Why I fit this role', { x: 5.5, y: 1, w: 4, fontSize: 24, bold: true, color: C.dark });
slide.addText('• Strong outcome-driven product management.\n• Heavily utilizes AI in execution and planning (as requested in JD).\n• Problem-first thinker focused on B2B SaaS retention metrics.', { x: 5.5, y: 2, w: 4, fontSize: 14, color: C.dark, bullet: true, lineSpacing: 24 });

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 5.5, w: 9, h: 1.5, fill: { color: C.bright }, rectRadius: 0.1 });
slide.addText('"Time tracking data is a goldmine that most tools ignore. Building intelligence on top of it creates unshakeable B2B stickiness."', { x: 1, y: 5.6, w: 8, h: 1.3, fontSize: 16, bold: true, color: C.dark, align: 'center', italic: true });

// SLIDE 8: Rollout Plan
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('ROLLOUT PLAN', { x: 0.5, y: 0.5, w: 9, fontSize: 10, bold: true, color: C.muted, letterSpacing: 0.08 });
slide.addText('Execution Milestones', { x: 0.5, y: 1, w: 9, fontSize: 28, bold: true, color: C.white });

const phases = [
  { t: 'Phase 1: Analysis', d: 'Analyze large-scale semantic overlap between client task tags and tracked hours.' },
  { t: 'Phase 2: Alpha', d: 'Roll out "Buffer Recommendation" UI to power users for implicit feedback.' },
  { t: 'Phase 3: Scale', d: 'Launch full SmartScope as a premium "Time Intelligence" feature.' }
];

phases.forEach((p, i) => {
  slide.addShape(prs.ShapeType.rect, { x: 0.5 + i*3.1, y: 2.5, w: 2.8, h: 2, fill: { color: C.hero }, rectRadius: 0.1, line: { color: C.bright, width: 2 } });
  slide.addText(p.t, { x: 0.7 + i*3.1, y: 2.8, w: 2.4, fontSize: 14, bold: true, color: C.white });
  slide.addText(p.d, { x: 0.7 + i*3.1, y: 3.5, w: 2.4, fontSize: 12, color: C.muted });
});

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 5, w: 9, h: 1.2, fill: { color: C.hero }, rectRadius: 0.1 });
slide.addText('What I Need: A 20-minute chat to discuss Toggl Focus AI aspirations.', { x: 0.8, y: 5.3, w: 8.4, fontSize: 14, color: C.bright, bold: true });
slide.addText('Ajay Avaghade | avaghadeajay009@gmail.com', { x: 0.5, y: 6.8, w: 9, fontSize: 11, color: C.muted, align: 'center' });

prs.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/toggl-smartscope-deck.pptx' }).then(() => {
  console.log('SmartScope Deck generated!');
});
