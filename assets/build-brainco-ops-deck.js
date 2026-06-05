const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0B1120',
  hero:   '1E293B',
  bright: '0EA5E9',
  orange: 'F59E0B',
  white:  'F8FAFC',
  lgray:  'F1F5F9',
  green:  '10B981',
};

const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });

// SLIDE 1: Cover
let s1 = pptx.addSlide();
s1.background = { color: C.dark };
s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.2, fill: { color: C.bright } });
s1.addText('BrainOps', { x: 0.5, y: 2.2, w: 9, fontSize: 44, bold: true, color: C.white, fontFace: 'Arial' });
s1.addText('Permissioned AI Deployment & Shadow Testing Pipeline', { x: 0.5, y: 3.2, w: 9, fontSize: 24, color: C.bright, fontFace: 'Arial' });
s1.addText('Ajay Avaghade · Product Manager', { x: 0.5, y: 6.5, w: 5, fontSize: 14, color: C.white, fontFace: 'Arial' });
s1.addText('Brain Co. Case Study 3', { x: 0.5, y: 6.8, w: 5, fontSize: 12, color: C.bright, fontFace: 'Arial', bold: true });
s1.addShape(pptx.ShapeType.ellipse, { x: 6.5, y: 1.5, w: 3, h: 3, fill: { color: C.hero } });
s1.addText('Risk-Free', { x: 6.5, y: 2.2, w: 3, fontSize: 32, bold: true, color: C.green, align: 'center', fontFace: 'Arial' });
s1.addText('Production Cutover', { x: 6.5, y: 3.2, w: 3, fontSize: 14, color: C.white, align: 'center', fontFace: 'Arial' });

// SLIDE 2: The Problem
let s2 = pptx.addSlide();
s2.background = { color: C.dark };
s2.addText('THE PROBLEM', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s2.addText('Deployment Anxiety in Mission-Critical Systems', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.white });
s2.addText('Clients are terrified to "flip the switch" on AI because model drift or hallucination can have severe legal consequences.', { x: 0.5, y: 1.8, w: 9, fontSize: 16, color: C.white });

s2.addText('No Testing Ground', { x: 0.5, y: 3.0, w: 3.0, fontSize: 20, bold: true, color: C.orange });
s2.addText('Clients don\'t trust AI models trained purely in lab environments.', { x: 0.5, y: 3.6, w: 2.5, fontSize: 13, color: '94A3B8' });

s2.addText('All-or-Nothing Risk', { x: 3.5, y: 3.0, w: 3.0, fontSize: 20, bold: true, color: C.orange });
s2.addText('Standard CI/CD lacks granular traffic routing for operational users.', { x: 3.5, y: 3.6, w: 2.5, fontSize: 13, color: '94A3B8' });

s2.addText('Lack of Accountability', { x: 6.5, y: 3.0, w: 3.0, fontSize: 20, bold: true, color: C.orange });
s2.addText('Government deployments require cryptographic executive sign-offs.', { x: 6.5, y: 3.6, w: 2.5, fontSize: 13, color: '94A3B8' });

// SLIDE 3: The Insight
let s3 = pptx.addSlide();
s3.background = { color: C.lgray };
s3.addText('THE INSIGHT', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s3.addText('Prove AI Efficacy Silently, Then Roll Out Safely', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.dark });

s3.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2.5, w: 4.0, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
s3.addText('1. Shadow Mode', { x: 0.8, y: 3.0, w: 3.4, fontSize: 16, bold: true, color: C.dark });
s3.addText('Run the AI on live production data, but hide decisions. Compare AI performance vs Human Baseline in real-time.', { x: 0.8, y: 3.5, w: 3.4, fontSize: 13, color: '64748B' });

s3.addShape(pptx.ShapeType.rect, { x: 5.0, y: 2.5, w: 4.0, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
s3.addText('2. Canary Rollouts & Sign-Offs', { x: 5.3, y: 3.0, w: 3.4, fontSize: 16, bold: true, color: C.dark });
s3.addText('Gradually dial up AI traffic. Force explicit, audited executive approval before a 100% production cutover.', { x: 5.3, y: 3.5, w: 3.4, fontSize: 13, color: '64748B' });

// SLIDE 4: Proof of Work
let s4 = pptx.addSlide();
s4.background = { color: C.white };
s4.addText('PROOF OF WORK', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s4.addText('Why Ajay Is Equipped to Build This', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.dark });

s4.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2.0, w: 4.3, h: 4.0, fill: { color: C.dark } });
s4.addText('A/B Testing Infrastructure at PhonePe', { x: 0.8, y: 2.5, w: 3.7, fontSize: 16, bold: true, color: C.white });
s4.addText('• Designed and executed 25+ A/B tests to validate checkout logic.\n• Led agile delivery cycles for ML model rollouts without breaking core CX.\n• Deeply familiar with phased rollouts and traffic routing.', { x: 0.8, y: 3.2, w: 3.7, fontSize: 13, color: '94A3B8', lineSpacing: 24 });

s4.addShape(pptx.ShapeType.rect, { x: 5.2, y: 2.0, w: 4.3, h: 4.0, fill: { color: C.lgray } });
s4.addText('Mapping to Brain Co.', { x: 5.5, y: 2.5, w: 3.7, fontSize: 16, bold: true, color: C.dark });
s4.addText('• Can translate complex MLOps constraints into a simple UX.\n• Understands how to build trust with enterprise stakeholders.\n• Proven ability to mitigate deployment risk through structured testing.', { x: 5.5, y: 3.2, w: 3.7, fontSize: 13, color: C.dark, lineSpacing: 24 });

// SLIDE 5: Contact
let s5 = pptx.addSlide();
s5.background = { color: C.dark };
s5.addText('Brain Co. Case Study 3: BrainOps', { x: 0.5, y: 2.5, w: 9, fontSize: 24, bold: true, color: C.white, align: 'center' });
s5.addShape(pptx.ShapeType.rect, { x: 0.5, y: 4.5, w: 9, h: 1.2, fill: { color: C.white } });
s5.addText('Contact: Ajay Avaghade | avaghadeajay009@gmail.com | linkedin.com/in/ajay-avaghade', { x: 0.5, y: 4.9, w: 9, fontSize: 14, bold: true, color: C.dark, align: 'center' });

pptx.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/brainco-ops-deck.pptx' });
