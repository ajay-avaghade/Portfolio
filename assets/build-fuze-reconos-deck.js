const pptxgen = require('pptxgenjs');
const prs = new pptxgen();

// Fuze Brand Colors
const C = {
  bg: '0A0E17',
  panel: '111827',
  accent: '06B6D4', // Cyan
  accent2: '3B82F6', // Blue
  white: 'F3F4F6',
  lgray: '9CA3AF',
  dark: '000000',
  success: '10B981'
};

prs.layout = 'LAYOUT_16x9';
prs.defineSlideMaster({
  title: 'FUZE_MASTER',
  background: { color: C.bg },
  objects: [
    { rect: { x: 0, y: 0, w: '100%', h: 0.1, fill: { color: C.accent } } },
    { text: { text: 'FUZE API INFRASTRUCTURE // CONFIDENTIAL', options: { x: 0.5, y: 5.2, w: 5, fontSize: 8, color: C.lgray } } }
  ]
});

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 3, offset: 2, opacity: 0.8 });

// SLIDE 1: Title
let slide = prs.addSlide({ masterName: 'FUZE_MASTER' });
slide.addText('Fuze ReconOS:', { x: 1, y: 1.8, w: 8, fontSize: 36, bold: true, color: C.accent, shadow: makeShadow() });
slide.addText('Automated Multi-Currency Ledger', { x: 1, y: 2.5, w: 8, fontSize: 24, bold: true, color: C.white });
slide.addText('Auto-matching embedded crypto payouts with fiat bank settlements for B2B partners', { x: 1, y: 3.1, w: 8, fontSize: 14, color: C.lgray });
slide.addText('Ajay Avaghade', { x: 1, y: 4.5, w: 4, fontSize: 12, bold: true, color: C.white });

// SLIDE 2: The Problem
slide = prs.addSlide({ masterName: 'FUZE_MASTER' });
slide.addText('The Friction', { x: 0.5, y: 0.5, w: 4, fontSize: 20, bold: true, color: C.accent });
slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 1.5, w: 4, h: 2.5, fill: { color: C.panel } });
slide.addText('Current State', { x: 0.7, y: 1.7, w: 3.6, fontSize: 14, bold: true, color: C.white });
slide.addText('• Manual spreadsheet reconciliation\n• Disconnect between on-chain & off-chain data\n• High operational overhead\n• Regulatory compliance risk', { x: 0.7, y: 2.1, w: 3.6, fontSize: 12, color: C.lgray, bullet: true, lineSpacing: 20 });

slide.addShape(prs.ShapeType.rect, { x: 5.5, y: 1.5, w: 4, h: 2.5, fill: { color: C.panel }, line: { color: C.accent, width: 1 } });
slide.addText('The Business Impact', { x: 5.7, y: 1.7, w: 3.6, fontSize: 14, bold: true, color: C.white });
slide.addText('• Stalled B2B partner scaling\n• Delayed merchant payouts\n• Costly compliance audits', { x: 5.7, y: 2.1, w: 3.6, fontSize: 12, color: C.lgray, bullet: true, lineSpacing: 20 });

// SLIDE 3: The Solution
slide = prs.addSlide({ masterName: 'FUZE_MASTER' });
slide.addText('ReconOS Architecture', { x: 0.5, y: 0.5, w: 5, fontSize: 20, bold: true, color: C.accent });

slide.addText('❌ Legacy Operations', { x: 0.5, y: 1.5, w: 3, fontSize: 16, bold: true, color: C.lgray });
slide.addText('• Siloed databases\n• T+1 batch matching\n• High human error rate', { x: 0.5, y: 2, w: 3, fontSize: 13, color: C.lgray, bullet: true, lineSpacing: 24 });

slide.addText('✅ Fuze ReconOS', { x: 6, y: 1.5, w: 3, fontSize: 16, bold: true, color: C.accent });
slide.addText('• Unified data layer\n• Auto-matching engine\n• Proactive risk flagging\n• 100% automated settlement', { x: 6, y: 2, w: 3, fontSize: 13, color: C.white, bullet: true, lineSpacing: 24 });

slide.addShape(prs.ShapeType.ellipse, { x: 4.6, y: 2.1, w: 0.8, h: 0.8, fill: { color: C.accent }, shadow: makeShadow() });
slide.addText('VS', { x: 4.6, y: 2.1, w: 0.8, h: 0.8, fontSize: 12, bold: true, color: C.dark, align: 'center' });

// SLIDE 4: The Mechanic
slide = prs.addSlide({ masterName: 'FUZE_MASTER' });
slide.addText('How it works', { x: 0.5, y: 0.5, w: 4, fontSize: 20, bold: true, color: C.accent });

const steps = ['Data Ingest', 'Index Flows', 'Auto Match', 'Settle Fiat'];
slide.addShape(prs.ShapeType.rect, { x: 1, y: 2.5, w: 8, h: 0.05, fill: { color: C.lgray } });

steps.forEach((step, i) => {
  slide.addShape(prs.ShapeType.ellipse, { x: 1.3 + i*2, y: 2.3, w: 0.4, h: 0.4, fill: { color: C.accent } });
  slide.addText(String(i+1), { x: 1.3 + i*2, y: 2.3, w: 0.4, h: 0.4, fontSize: 12, bold: true, color: C.dark, align: 'center' });
  slide.addText(step, { x: 0.5 + i*2, y: 3.0, w: 2, fontSize: 14, bold: true, color: C.white, align: 'center' });
});

// SLIDE 5: Metrics
slide = prs.addSlide({ masterName: 'FUZE_MASTER' });
slide.addText('Impact & KPIs', { x: 0.5, y: 0.5, w: 4, fontSize: 20, bold: true, color: C.accent });

const metrics = [
  { val: '-60%', lbl: 'Ops Time' },
  { val: '100%', lbl: 'Automation' },
  { val: 'Zero', lbl: 'Compliance Drops' }
];

metrics.forEach((m, i) => {
  slide.addShape(prs.ShapeType.rect, { x: 0.5 + i*3.2, y: 1.8, w: 2.8, h: 2, fill: { color: C.panel }, line: { color: C.accent, width: 1 } });
  slide.addText(m.val, { x: 0.5 + i*3.2, y: 2.2, w: 2.8, fontSize: 36, bold: true, color: C.success, align: 'center' });
  slide.addText(m.lbl, { x: 0.5 + i*3.2, y: 3.0, w: 2.8, fontSize: 14, color: C.white, align: 'center' });
});

prs.writeFile({ fileName: 'assets/fuze-reconos-deck.pptx' }).then(() => {
  console.log('Fuze ReconOS Deck generated!');
});
