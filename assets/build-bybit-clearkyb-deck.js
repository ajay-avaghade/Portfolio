const pptxgen = require('pptxgenjs');
const prs = new pptxgen();

// Bybit Brand Colors
const C = {
  bg: '0B0E11',
  panel: '181A20',
  bright: 'F7A600', // Bybit Yellow
  white: 'EAECEF',
  lgray: '848E9C',
  dark: '000000',
  success: '0ECB81'
};

prs.layout = 'LAYOUT_16x9';
prs.defineSlideMaster({
  title: 'BYBIT_MASTER',
  background: { color: C.bg },
  objects: [
    { rect: { x: 0, y: 0, w: '100%', h: 0.1, fill: { color: C.bright } } },
    { text: { text: 'BYBIT EMI PRODUCT STRATEGY // CONFIDENTIAL', options: { x: 0.5, y: 5.2, w: 5, fontSize: 8, color: C.lgray } } }
  ]
});

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 3, offset: 2, opacity: 0.8 });

// SLIDE 1: Title
let slide = prs.addSlide({ masterName: 'BYBIT_MASTER' });
slide.addText('ClearKYB:', { x: 1, y: 1.8, w: 8, fontSize: 36, bold: true, color: C.bright, shadow: makeShadow() });
slide.addText('Progressive Institutional Onboarding', { x: 1, y: 2.5, w: 8, fontSize: 24, bold: true, color: C.white });
slide.addText('Solving KYC drop-offs for European entities', { x: 1, y: 3.1, w: 8, fontSize: 14, color: C.lgray });
slide.addText('Ajay Avaghade', { x: 1, y: 4.5, w: 4, fontSize: 12, bold: true, color: C.white });

// SLIDE 2: The Problem
slide = prs.addSlide({ masterName: 'BYBIT_MASTER' });
slide.addText('The Friction', { x: 0.5, y: 0.5, w: 4, fontSize: 20, bold: true, color: C.bright });
slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 1.5, w: 4, h: 2.5, fill: { color: C.panel } });
slide.addText('Current State', { x: 0.7, y: 1.7, w: 3.6, fontSize: 14, bold: true, color: C.white });
slide.addText('• 100% upfront verification required\n• Heavy document uploads at step 1\n• 40% drop-off in institutional funnels\n• Delayed time-to-first-trade', { x: 0.7, y: 2.1, w: 3.6, fontSize: 12, color: C.lgray, bullet: true, lineSpacing: 20 });

slide.addShape(prs.ShapeType.rect, { x: 5.5, y: 1.5, w: 4, h: 2.5, fill: { color: C.panel }, line: { color: C.bright, width: 1 } });
slide.addText('The Business Impact', { x: 5.7, y: 1.7, w: 3.6, fontSize: 14, bold: true, color: C.white });
slide.addText('• Lost AUM from high-net-worth clients\n• Overwhelmed compliance ops teams\n• Poor UX vs local EEA competitors', { x: 5.7, y: 2.1, w: 3.6, fontSize: 12, color: C.lgray, bullet: true, lineSpacing: 20 });

// SLIDE 3: The Solution
slide = prs.addSlide({ masterName: 'BYBIT_MASTER' });
slide.addText('ClearKYB Architecture', { x: 0.5, y: 0.5, w: 5, fontSize: 20, bold: true, color: C.bright });

slide.addText('❌ Legacy Verification', { x: 0.5, y: 1.5, w: 3, fontSize: 16, bold: true, color: C.lgray });
slide.addText('• Monolithic forms\n• Synchronous checks\n• High cognitive load\n• Binary access (all/none)', { x: 0.5, y: 2, w: 3, fontSize: 13, color: C.lgray, bullet: true, lineSpacing: 24 });

slide.addText('✅ ClearKYB Flow', { x: 6, y: 1.5, w: 3, fontSize: 16, bold: true, color: C.bright });
slide.addText('• Progressive profiling\n• Tiered API access limits\n• Async document review\n• Immediate Crypto-only access', { x: 6, y: 2, w: 3, fontSize: 13, color: C.white, bullet: true, lineSpacing: 24 });

slide.addShape(prs.ShapeType.ellipse, { x: 4.6, y: 2.1, w: 0.8, h: 0.8, fill: { color: C.bright }, shadow: makeShadow() });
slide.addText('VS', { x: 4.6, y: 2.1, w: 0.8, h: 0.8, fontSize: 12, bold: true, color: C.dark, align: 'center' });

// SLIDE 4: The Mechanic
slide = prs.addSlide({ masterName: 'BYBIT_MASTER' });
slide.addText('How it works', { x: 0.5, y: 0.5, w: 4, fontSize: 20, bold: true, color: C.bright });

const steps = ['Basic Reg', 'Tier 1 Access', 'Doc Upload', 'Full SEPA'];
slide.addShape(prs.ShapeType.rect, { x: 1, y: 2.5, w: 8, h: 0.05, fill: { color: C.lgray } });

steps.forEach((step, i) => {
  slide.addShape(prs.ShapeType.ellipse, { x: 1.3 + i*2, y: 2.3, w: 0.4, h: 0.4, fill: { color: C.bright } });
  slide.addText(String(i+1), { x: 1.3 + i*2, y: 2.3, w: 0.4, h: 0.4, fontSize: 12, bold: true, color: C.dark, align: 'center' });
  slide.addText(step, { x: 0.5 + i*2, y: 3.0, w: 2, fontSize: 14, bold: true, color: C.white, align: 'center' });
});

// SLIDE 5: Metrics
slide = prs.addSlide({ masterName: 'BYBIT_MASTER' });
slide.addText('Impact & KPIs', { x: 0.5, y: 0.5, w: 4, fontSize: 20, bold: true, color: C.bright });

const metrics = [
  { val: '+35%', lbl: 'Onboarding Conv.' },
  { val: '-40%', lbl: 'Manual Review Time' },
  { val: 'T+0', lbl: 'Crypto Access' }
];

metrics.forEach((m, i) => {
  slide.addShape(prs.ShapeType.rect, { x: 0.5 + i*3.2, y: 1.8, w: 2.8, h: 2, fill: { color: C.panel }, line: { color: C.bright, width: 1 } });
  slide.addText(m.val, { x: 0.5 + i*3.2, y: 2.2, w: 2.8, fontSize: 36, bold: true, color: C.success, align: 'center' });
  slide.addText(m.lbl, { x: 0.5 + i*3.2, y: 3.0, w: 2.8, fontSize: 14, color: C.white, align: 'center' });
});

prs.writeFile({ fileName: 'assets/bybit-clearkyb-deck.pptx' }).then(() => {
  console.log('ClearKYB Deck generated!');
});
