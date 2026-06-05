const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_16x9';

const C = {
  dark:   '1E1B4B',
  bright: '6D28D9',
  orange: 'F97316',
  white:  'F8FAFC',
  lgray:  'F1F5F9',
  green:  '10B981',
};

const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });

// SLIDE 1: Cover
let s1 = pptx.addSlide();
s1.background = { color: C.dark };
s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.2, fill: { color: C.orange } });
s1.addText('EatSure: SurePass', { x: 0.5, y: 2.2, w: 9, fontSize: 44, bold: true, color: C.white, fontFace: 'Arial' });
s1.addText('Predictive Loyalty & ML Discounting', { x: 0.5, y: 3.2, w: 9, fontSize: 24, color: C.orange, fontFace: 'Arial' });
s1.addText('Ajay Avaghade · Senior Product Manager Candidate', { x: 0.5, y: 6.5, w: 5, fontSize: 14, color: C.white, fontFace: 'Arial' });
s1.addText('Rebel Foods Case Study 2', { x: 0.5, y: 6.8, w: 5, fontSize: 12, color: C.orange, fontFace: 'Arial', bold: true });
s1.addShape(pptx.ShapeType.ellipse, { x: 6.5, y: 1.5, w: 3, h: 3, fill: { color: C.bright } });
s1.addText('-32%', { x: 6.5, y: 2.2, w: 3, fontSize: 44, bold: true, color: C.green, align: 'center', fontFace: 'Arial' });
s1.addText('Marketing Burn', { x: 6.5, y: 3.2, w: 3, fontSize: 14, color: C.white, align: 'center', fontFace: 'Arial' });

// SLIDE 2: The Problem
let s2 = pptx.addSlide();
s2.background = { color: C.dark };
s2.addText('THE PROBLEM', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s2.addText('Blanket Discounting Destroys Unit Economics', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.white });
s2.addText('Food tech platforms bleed money by giving 20% off coupons to users who were going to order anyway.', { x: 0.5, y: 1.8, w: 9, fontSize: 16, color: C.white });

s2.addText('Margin Loss', { x: 0.5, y: 3.0, w: 2.5, fontSize: 40, bold: true, color: C.orange });
s2.addText('Habitual Users', { x: 0.5, y: 3.8, w: 2.5, fontSize: 16, bold: true, color: C.white });
s2.addText('They value speed and reliability over discounts.', { x: 0.5, y: 4.2, w: 2.5, fontSize: 13, color: '94A3B8' });

s2.addText('High CAC', { x: 3.5, y: 3.0, w: 2.5, fontSize: 40, bold: true, color: C.orange });
s2.addText('Churn Risk', { x: 3.5, y: 3.8, w: 2.5, fontSize: 16, bold: true, color: C.white });
s2.addText('Dormant users need aggressive offers to build habit.', { x: 3.5, y: 4.2, w: 2.5, fontSize: 13, color: '94A3B8' });

// SLIDE 3: The Insight
let s3 = pptx.addSlide();
s3.background = { color: C.lgray };
s3.addText('THE INSIGHT', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s3.addText('Propensity-to-Transact ML Models', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.dark });

s3.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2.5, w: 4.0, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
s3.addText('1. Withhold from High-Propensity', { x: 0.8, y: 3.0, w: 3.4, fontSize: 16, bold: true, color: C.dark });
s3.addText('If the ML model scores a user at 80%+ likelihood to order, withhold the discount and offer a 1-click reorder.', { x: 0.8, y: 3.5, w: 3.4, fontSize: 13, color: '64748B' });

s3.addShape(pptx.ShapeType.rect, { x: 5.0, y: 2.5, w: 4.0, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
s3.addText('2. Inject for Low-Propensity', { x: 5.3, y: 3.0, w: 3.4, fontSize: 16, bold: true, color: C.dark });
s3.addText('If the user drops below a 20% score (churn risk), inject a high-value offer to reactivate their ordering habit.', { x: 5.3, y: 3.5, w: 3.4, fontSize: 13, color: '64748B' });

// SLIDE 4: Proof of Work
let s4 = pptx.addSlide();
s4.background = { color: C.white };
s4.addText('PROOF OF WORK', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s4.addText('Why Ajay Is Equipped to Build This', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.dark });

s4.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2.0, w: 4.3, h: 4.0, fill: { color: C.dark } });
s4.addText('ML Deployment at PhonePe', { x: 0.8, y: 2.5, w: 3.7, fontSize: 16, bold: true, color: C.white });
s4.addText('• Led the deployment of Propensity-to-Transact ML models.\n• Replaced static, manual cohort targeting with real-time AI.\n• Achieved a 32% reduction in marketing burn while sustaining GMV.', { x: 0.8, y: 3.2, w: 3.7, fontSize: 13, color: '94A3B8', lineSpacing: 24 });

s4.addShape(pptx.ShapeType.rect, { x: 5.2, y: 2.0, w: 4.3, h: 4.0, fill: { color: C.lgray } });
s4.addText('Mapping to EatSure', { x: 5.5, y: 2.5, w: 3.7, fontSize: 16, bold: true, color: C.dark });
s4.addText('• Can translate complex Data Science outputs into front-end lifecycle journeys.\n• Deep expertise in loyalty programs and retention cohorts.\n• Prioritizes sustainable unit economics.', { x: 5.5, y: 3.2, w: 3.7, fontSize: 13, color: C.dark, lineSpacing: 24 });

// SLIDE 5: Contact
let s5 = pptx.addSlide();
s5.background = { color: C.dark };
s5.addText('Rebel Foods Case Study: SurePass', { x: 0.5, y: 2.5, w: 9, fontSize: 24, bold: true, color: C.white, align: 'center' });
s5.addShape(pptx.ShapeType.rect, { x: 0.5, y: 4.5, w: 9, h: 1.2, fill: { color: C.white } });
s5.addText('Contact: Ajay Avaghade | avaghadeajay009@gmail.com | linkedin.com/in/ajay-avaghade', { x: 0.5, y: 4.9, w: 9, fontSize: 14, bold: true, color: C.dark, align: 'center' });

pptx.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/rebelfoods-pass-deck.pptx' });
