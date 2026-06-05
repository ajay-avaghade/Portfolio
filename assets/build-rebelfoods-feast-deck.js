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
s1.addText('EatSure: SureFeast', { x: 0.5, y: 2.2, w: 9, fontSize: 44, bold: true, color: C.white, fontFace: 'Arial' });
s1.addText('AI-Curated Occasion Ordering', { x: 0.5, y: 3.2, w: 9, fontSize: 24, color: C.orange, fontFace: 'Arial' });
s1.addText('Ajay Avaghade · Senior Product Manager Candidate', { x: 0.5, y: 6.5, w: 5, fontSize: 14, color: C.white, fontFace: 'Arial' });
s1.addText('Rebel Foods Case Study 3', { x: 0.5, y: 6.8, w: 5, fontSize: 12, color: C.orange, fontFace: 'Arial', bold: true });
s1.addShape(pptx.ShapeType.ellipse, { x: 6.5, y: 1.5, w: 3, h: 3, fill: { color: C.bright } });
s1.addText('4x', { x: 6.5, y: 2.2, w: 3, fontSize: 44, bold: true, color: C.orange, align: 'center', fontFace: 'Arial' });
s1.addText('AOV Growth', { x: 6.5, y: 3.2, w: 3, fontSize: 14, color: C.white, align: 'center', fontFace: 'Arial' });

// SLIDE 2: The Problem
let s2 = pptx.addSlide();
s2.background = { color: C.dark };
s2.addText('THE PROBLEM', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s2.addText('Large Orders Carry Massive Cognitive Load', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.white });
s2.addText('Ordering for 6 people with different dietary preferences across multiple brands takes 15 minutes of scrolling, causing drop-offs.', { x: 0.5, y: 1.8, w: 9, fontSize: 16, color: C.white });

s2.addText('High AOV', { x: 0.5, y: 3.0, w: 2.5, fontSize: 40, bold: true, color: C.orange });
s2.addText('High Friction', { x: 0.5, y: 3.8, w: 2.5, fontSize: 16, bold: true, color: C.white });
s2.addText('Occasion baskets are highly profitable but UX is lacking.', { x: 0.5, y: 4.2, w: 2.5, fontSize: 13, color: '94A3B8' });

s2.addText('The Moat', { x: 3.5, y: 3.0, w: 2.5, fontSize: 40, bold: true, color: C.orange });
s2.addText('Cloud Kitchens', { x: 3.5, y: 3.8, w: 2.5, fontSize: 16, bold: true, color: C.white });
s2.addText('Rebel Foods is uniquely positioned to fulfill multi-brand orders from a single location.', { x: 3.5, y: 4.2, w: 2.5, fontSize: 13, color: '94A3B8' });

// SLIDE 3: The Insight
let s3 = pptx.addSlide();
s3.background = { color: C.lgray };
s3.addText('THE INSIGHT', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s3.addText('AI-Curated Multi-Brand Carts', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.dark });

s3.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2.5, w: 4.0, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
s3.addText('1. Prompt-to-Cart Flow', { x: 0.8, y: 3.0, w: 3.4, fontSize: 16, bold: true, color: C.dark });
s3.addText('Users input their intent: "Dinner for 6, mostly Biryani but 2 vegans". The AI translates intent directly into a complete, balanced cart.', { x: 0.8, y: 3.5, w: 3.4, fontSize: 13, color: '64748B' });

s3.addShape(pptx.ShapeType.rect, { x: 5.0, y: 2.5, w: 4.0, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
s3.addText('2. 1-Click Checkout', { x: 5.3, y: 3.0, w: 3.4, fontSize: 16, bold: true, color: C.dark });
s3.addText('Skip the menu browsing phase entirely. Funnel high-intent "party" users straight to review and checkout, maximizing conversion.', { x: 5.3, y: 3.5, w: 3.4, fontSize: 13, color: '64748B' });

// SLIDE 4: Proof of Work
let s4 = pptx.addSlide();
s4.background = { color: C.white };
s4.addText('PROOF OF WORK', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s4.addText('Why Ajay Is Equipped to Build This', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.dark });

s4.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2.0, w: 4.3, h: 4.0, fill: { color: C.dark } });
s4.addText('Discovery Transformation at PhonePe', { x: 0.8, y: 2.5, w: 3.7, fontSize: 16, bold: true, color: C.white });
s4.addText('• Rebuilt ₹100 Cr/yr legacy systems into ML-driven experiences.\n• Transformed static discovery workflows into dynamic curation.\n• Increased user engagement by 26% and revenue by 11%.', { x: 0.8, y: 3.2, w: 3.7, fontSize: 13, color: '94A3B8', lineSpacing: 24 });

s4.addShape(pptx.ShapeType.rect, { x: 5.2, y: 2.0, w: 4.3, h: 4.0, fill: { color: C.lgray } });
s4.addText('Mapping to EatSure', { x: 5.5, y: 2.5, w: 3.7, fontSize: 16, bold: true, color: C.dark });
s4.addText('• Capable of designing complex AI-to-UX flows.\n• Focuses on leveraging unique business advantages (Cloud Kitchen logistics).\n• Proven impact on topline GMV growth.', { x: 5.5, y: 3.2, w: 3.7, fontSize: 13, color: C.dark, lineSpacing: 24 });

// SLIDE 5: Contact
let s5 = pptx.addSlide();
s5.background = { color: C.dark };
s5.addText('Rebel Foods Case Study: SureFeast', { x: 0.5, y: 2.5, w: 9, fontSize: 24, bold: true, color: C.white, align: 'center' });
s5.addShape(pptx.ShapeType.rect, { x: 0.5, y: 4.5, w: 9, h: 1.2, fill: { color: C.white } });
s5.addText('Contact: Ajay Avaghade | avaghadeajay009@gmail.com | linkedin.com/in/ajay-avaghade', { x: 0.5, y: 4.9, w: 9, fontSize: 14, bold: true, color: C.dark, align: 'center' });

pptx.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/rebelfoods-feast-deck.pptx' });
