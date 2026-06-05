const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_16x9';

const C = {
  dark:   '020617',
  hero:   '0F172A',
  bright: '6366F1',
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
s1.addText('FinnEarn', { x: 0.5, y: 2.2, w: 9, fontSize: 44, bold: true, color: C.white, fontFace: 'Arial' });
s1.addText('Gig-Income Marketplace & Interstitial Activation', { x: 0.5, y: 3.2, w: 9, fontSize: 24, color: C.bright, fontFace: 'Arial' });
s1.addText('Ajay Avaghade · Senior Product Manager Candidate', { x: 0.5, y: 6.5, w: 5, fontSize: 14, color: C.white, fontFace: 'Arial' });
s1.addText('FINN Case Study 3', { x: 0.5, y: 6.8, w: 5, fontSize: 12, color: C.bright, fontFace: 'Arial', bold: true });
s1.addShape(pptx.ShapeType.ellipse, { x: 6.5, y: 1.5, w: 3, h: 3, fill: { color: C.hero } });
s1.addText('-60%', { x: 6.5, y: 2.2, w: 3, fontSize: 44, bold: true, color: C.green, align: 'center', fontFace: 'Arial' });
s1.addText('Abandonment', { x: 6.5, y: 3.2, w: 3, fontSize: 14, color: C.white, align: 'center', fontFace: 'Arial' });

// SLIDE 2: The Problem
let s2 = pptx.addSlide();
s2.background = { color: C.dark };
s2.addText('THE PROBLEM', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s2.addText('Low Engagement with Secondary Features', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.white });
s2.addText('Users open the app to withdraw wages. They ignore the "Gig Marketplace" hidden in a side menu because they lack immediate motivation.', { x: 0.5, y: 1.8, w: 9, fontSize: 16, color: C.white });

s2.addText('Blind Spot', { x: 0.5, y: 3.0, w: 3.0, fontSize: 20, bold: true, color: C.orange });
s2.addText('Side menus get ignored.', { x: 0.5, y: 3.6, w: 2.5, fontSize: 13, color: '94A3B8' });

s2.addText('Loss Aversion', { x: 3.5, y: 3.0, w: 3.0, fontSize: 20, bold: true, color: C.orange });
s2.addText('Withdrawal fees feel like a penalty.', { x: 3.5, y: 3.6, w: 2.5, fontSize: 13, color: '94A3B8' });

s2.addText('Delayed Gratification', { x: 6.5, y: 3.0, w: 3.0, fontSize: 20, bold: true, color: C.orange });
s2.addText('Gigs rarely pay instantly.', { x: 6.5, y: 3.6, w: 2.5, fontSize: 13, color: '94A3B8' });

// SLIDE 3: The Insight
let s3 = pptx.addSlide();
s3.background = { color: C.lgray };
s3.addText('THE INSIGHT', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s3.addText('Dynamic Interstitial Incentivization', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.dark });

s3.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2.5, w: 4.0, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
s3.addText('1. The Interstitial Offer', { x: 0.8, y: 3.0, w: 3.4, fontSize: 16, bold: true, color: C.dark });
s3.addText('When users hit "withdraw" and see a fee, immediately offer a micro-task (like a survey) priced exactly to cover that fee.', { x: 0.8, y: 3.5, w: 3.4, fontSize: 13, color: '64748B' });

s3.addShape(pptx.ShapeType.rect, { x: 5.0, y: 2.5, w: 4.0, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
s3.addText('2. Instant Crediting', { x: 5.3, y: 3.0, w: 3.4, fontSize: 16, bold: true, color: C.dark });
s3.addText('Fulfill FINN\'s promise of speed. Credit the survey earnings instantly, creating a powerful reward loop.', { x: 5.3, y: 3.5, w: 3.4, fontSize: 13, color: '64748B' });

// SLIDE 4: Proof of Work
let s4 = pptx.addSlide();
s4.background = { color: C.white };
s4.addText('PROOF OF WORK', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s4.addText('Why Ajay Is Equipped to Build This', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.dark });

s4.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2.0, w: 4.3, h: 4.0, fill: { color: C.dark } });
s4.addText('Cart Incentivization at PhonePe', { x: 0.8, y: 2.5, w: 3.7, fontSize: 16, bold: true, color: C.white });
s4.addText('• Built a dynamic cart incentivization engine for quick commerce.\n• Reduced cart abandonment by 60% by offering the right incentive at the right time.\n• Rebuilt a ₹100Cr rewards portfolio into an ML marketplace.', { x: 0.8, y: 3.2, w: 3.7, fontSize: 13, color: '94A3B8', lineSpacing: 24 });

s4.addShape(pptx.ShapeType.rect, { x: 5.2, y: 2.0, w: 4.3, h: 4.0, fill: { color: C.lgray } });
s4.addText('Mapping to FINN', { x: 5.5, y: 2.5, w: 3.7, fontSize: 16, bold: true, color: C.dark });
s4.addText('• Deep understanding of behavioral psychology in checkout/withdrawal funnels.\n• Proven ability to match B2B supply (brands) with B2C demand (users).\n• Drives engagement by removing friction.', { x: 5.5, y: 3.2, w: 3.7, fontSize: 13, color: C.dark, lineSpacing: 24 });

// SLIDE 5: Contact
let s5 = pptx.addSlide();
s5.background = { color: C.dark };
s5.addText('FINN Case Study 3: FinnEarn', { x: 0.5, y: 2.5, w: 9, fontSize: 24, bold: true, color: C.white, align: 'center' });
s5.addShape(pptx.ShapeType.rect, { x: 0.5, y: 4.5, w: 9, h: 1.2, fill: { color: C.white } });
s5.addText('Contact: Ajay Avaghade | avaghadeajay009@gmail.com | linkedin.com/in/ajay-avaghade', { x: 0.5, y: 4.9, w: 9, fontSize: 14, bold: true, color: C.dark, align: 'center' });

pptx.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/finn-earn-deck.pptx' });
