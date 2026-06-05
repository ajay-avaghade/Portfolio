const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_16x9';

const C = {
  dark:   '1E1B4B',
  hero:   '312E81',
  bright: '4F46E5',
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
s1.addText('FinnFlow', { x: 0.5, y: 2.2, w: 9, fontSize: 44, bold: true, color: C.white, fontFace: 'Arial' });
s1.addText('Progressive KYC & Open Banking Onboarding', { x: 0.5, y: 3.2, w: 9, fontSize: 24, color: C.bright, fontFace: 'Arial' });
s1.addText('Ajay Avaghade · Senior Product Manager Candidate', { x: 0.5, y: 6.5, w: 5, fontSize: 14, color: C.white, fontFace: 'Arial' });
s1.addText('FINN Case Study 1', { x: 0.5, y: 6.8, w: 5, fontSize: 12, color: C.bright, fontFace: 'Arial', bold: true });
s1.addShape(pptx.ShapeType.ellipse, { x: 6.5, y: 1.5, w: 3, h: 3, fill: { color: C.hero } });
s1.addText('+25%', { x: 6.5, y: 2.2, w: 3, fontSize: 44, bold: true, color: C.green, align: 'center', fontFace: 'Arial' });
s1.addText('Activation Rate', { x: 6.5, y: 3.2, w: 3, fontSize: 14, color: C.white, align: 'center', fontFace: 'Arial' });

// SLIDE 2: The Problem
let s2 = pptx.addSlide();
s2.background = { color: C.dark };
s2.addText('THE PROBLEM', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s2.addText('High Drop-offs in Fintech Onboarding', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.white });
s2.addText('Underbanked users face immense friction when greeted with 15-field KYC forms before seeing any value.', { x: 0.5, y: 1.8, w: 9, fontSize: 16, color: C.white });

s2.addText('Friction', { x: 0.5, y: 3.0, w: 2.5, fontSize: 40, bold: true, color: C.orange });
s2.addText('Manual Payslips', { x: 0.5, y: 3.8, w: 2.5, fontSize: 16, bold: true, color: C.white });
s2.addText('Uploading documents creates days of delay.', { x: 0.5, y: 4.2, w: 2.5, fontSize: 13, color: '94A3B8' });

s2.addText('Trust', { x: 3.5, y: 3.0, w: 2.5, fontSize: 40, bold: true, color: C.orange });
s2.addText('Anxiety', { x: 3.5, y: 3.8, w: 2.5, fontSize: 16, bold: true, color: C.white });
s2.addText('Connecting bank accounts requires high trust upfront.', { x: 3.5, y: 4.2, w: 2.5, fontSize: 13, color: '94A3B8' });

s2.addText('Value', { x: 6.5, y: 3.0, w: 2.5, fontSize: 40, bold: true, color: C.orange });
s2.addText('Delayed', { x: 6.5, y: 3.8, w: 2.5, fontSize: 16, bold: true, color: C.white });
s2.addText('Users churn before seeing their earned wages.', { x: 6.5, y: 4.2, w: 2.5, fontSize: 13, color: '94A3B8' });

// SLIDE 3: The Insight
let s3 = pptx.addSlide();
s3.background = { color: C.lgray };
s3.addText('THE INSIGHT', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s3.addText('Value Before Effort: Reverse the Funnel', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.dark });
s3.addText('By utilizing Open Banking early and delaying hard KYC until withdrawal, we leverage the "sunk cost fallacy" to boost conversion.', { x: 0.5, y: 1.8, w: 9, fontSize: 14, color: C.dark });

s3.addShape(pptx.ShapeType.rect, { x: 0.5, y: 3.0, w: 4.0, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
s3.addText('1. The Instant Hook', { x: 0.8, y: 3.5, w: 3.4, fontSize: 16, bold: true, color: C.dark });
s3.addText('Show the user their earned wages immediately via payroll API before asking for KYC.', { x: 0.8, y: 4.0, w: 3.4, fontSize: 13, color: '64748B' });

s3.addShape(pptx.ShapeType.rect, { x: 5.0, y: 3.0, w: 4.0, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
s3.addText('2. Progressive Profiling', { x: 5.3, y: 3.5, w: 3.4, fontSize: 16, bold: true, color: C.dark });
s3.addText('Only prompt for mandatory identity checks right before the first withdrawal.', { x: 5.3, y: 4.0, w: 3.4, fontSize: 13, color: '64748B' });

// SLIDE 4: Proof of Work
let s4 = pptx.addSlide();
s4.background = { color: C.white };
s4.addText('PROOF OF WORK', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s4.addText('Why Ajay Is Equipped to Build This', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.dark });

s4.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2.0, w: 4.3, h: 4.0, fill: { color: C.dark } });
s4.addText('A/B Testing Funnels at PhonePe', { x: 0.8, y: 2.5, w: 3.7, fontSize: 16, bold: true, color: C.white });
s4.addText('• Designed and executed 25+ A/B tests to optimize checkout conversion.\n• Redesigned end-to-end flows resulting in a 22% conversion lift.\n• Deep expertise in identifying and removing friction points.', { x: 0.8, y: 3.2, w: 3.7, fontSize: 13, color: '94A3B8', lineSpacing: 24 });

s4.addShape(pptx.ShapeType.rect, { x: 5.2, y: 2.0, w: 4.3, h: 4.0, fill: { color: C.lgray } });
s4.addText('Mapping to FINN', { x: 5.5, y: 2.5, w: 3.7, fontSize: 16, bold: true, color: C.dark });
s4.addText('• Proven ability to use data (SQL, Mixpanel) to find drop-offs.\n• Understands how to integrate complex third-party APIs (Open Banking) cleanly.\n• Balances growth hacking with strict regulatory compliance.', { x: 5.5, y: 3.2, w: 3.7, fontSize: 13, color: C.dark, lineSpacing: 24 });

// SLIDE 5: Contact
let s5 = pptx.addSlide();
s5.background = { color: C.dark };
s5.addText('FINN Case Study 1: FinnFlow', { x: 0.5, y: 2.5, w: 9, fontSize: 24, bold: true, color: C.white, align: 'center' });
s5.addShape(pptx.ShapeType.rect, { x: 0.5, y: 4.5, w: 9, h: 1.2, fill: { color: C.white } });
s5.addText('Contact: Ajay Avaghade | avaghadeajay009@gmail.com | linkedin.com/in/ajay-avaghade', { x: 0.5, y: 4.9, w: 9, fontSize: 14, bold: true, color: C.dark, align: 'center' });

pptx.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/finn-flow-deck.pptx' });
