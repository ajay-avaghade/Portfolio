const pptxgen = require('pptxgenjs');

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';

const C = {
    dark: '0F172A',
    hero: '1E293B',
    bright: '3B82F6',
    accent: '60A5FA',
    orange: 'F59E0B',
    red: 'EF4444',
    bg: 'F8FAFC',
    text: '0F172A',
    muted: '64748B',
    green: '10B981'
};

// Title Slide
let slide1 = pres.addSlide();
slide1.background = { color: C.dark };
slide1.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.dark } });
slide1.addText('ReconHub Engine', { x: 0.5, y: 2.0, w: 9, fontSize: 54, bold: true, color: 'FFFFFF', fontFace: 'Helvetica' });
slide1.addText('Multi-Channel Payout & Dispute Reconciliation', { x: 0.5, y: 3.0, w: 9, fontSize: 24, color: C.accent, fontFace: 'Helvetica' });
slide1.addText('Prepared for STYLI | MaaS Initiative', { x: 0.5, y: 4.8, w: 9, fontSize: 16, color: C.muted, fontFace: 'Helvetica' });

// 2. The Bottleneck
let slide2 = pres.addSlide();
slide2.background = { color: C.bg };
slide2.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 0.2, h: '100%', fill: { color: C.red } });
slide2.addText('The Reconciliation Nightmare', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: C.dark, fontFace: 'Helvetica' });
slide2.addText('Sellers lose significant revenue to untracked channel fees:', { x: 0.5, y: 1.2, w: 9, fontSize: 18, color: C.muted, fontFace: 'Helvetica' });
slide2.addText([
    { text: 'Complex Fee Structures: ', options: { bold: true } },
    { text: 'Amazon and Noon have different category fees, making manual audits impossible.' }
], { x: 0.5, y: 2.0, w: 8, fontSize: 20, color: C.text, bullet: true, fontFace: 'Helvetica' });
slide2.addText([
    { text: 'Lost Disputes: ', options: { bold: true } },
    { text: 'Missing the 14-day window to contest incorrect dimension fees or damaged returns.' }
], { x: 0.5, y: 2.8, w: 8, fontSize: 20, color: C.text, bullet: true, fontFace: 'Helvetica' });
slide2.addText([
    { text: 'Blind Payouts: ', options: { bold: true } },
    { text: 'Finance teams cannot accurately attribute STYLI net revenue per external channel.' }
], { x: 0.5, y: 3.6, w: 8, fontSize: 20, color: C.text, bullet: true, fontFace: 'Helvetica' });

// 3. The Solution
let slide3 = pres.addSlide();
slide3.background = { color: C.hero };
slide3.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 0.2, h: '100%', fill: { color: C.bright } });
slide3.addText('The Solution: ReconHub', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: 'FFFFFF', fontFace: 'Helvetica' });
slide3.addText('A unified financial ledger that automatically audits marketplace payouts.', { x: 0.5, y: 1.2, w: 9, fontSize: 18, color: C.accent, fontFace: 'Helvetica' });

slide3.addShape(pres.ShapeType.roundRect, { x: 0.5, y: 2.0, w: 2.5, h: 2, fill: { color: C.dark }, line: { color: C.bright, width: 2 } });
slide3.addText('1. Ingest', { x: 0.5, y: 2.3, w: 2.5, fontSize: 20, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });
slide3.addText('API fetch of Amazon/Noon settlement reports.', { x: 0.7, y: 2.8, w: 2.1, fontSize: 14, color: C.muted, align: 'center', fontFace: 'Helvetica' });

slide3.addShape(pres.ShapeType.rightArrow, { x: 3.2, y: 2.8, w: 0.6, h: 0.4, fill: { color: C.accent } });

slide3.addShape(pres.ShapeType.roundRect, { x: 4.0, y: 2.0, w: 2.5, h: 2, fill: { color: C.dark }, line: { color: C.bright, width: 2 } });
slide3.addText('2. Parse & Match', { x: 4.0, y: 2.3, w: 2.5, fontSize: 20, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });
slide3.addText('Compare actual deductions vs STYLI expected fees.', { x: 4.2, y: 2.8, w: 2.1, fontSize: 14, color: C.muted, align: 'center', fontFace: 'Helvetica' });

slide3.addShape(pres.ShapeType.rightArrow, { x: 6.7, y: 2.8, w: 0.6, h: 0.4, fill: { color: C.accent } });

slide3.addShape(pres.ShapeType.roundRect, { x: 7.5, y: 2.0, w: 2.5, h: 2, fill: { color: C.dark }, line: { color: C.bright, width: 2 } });
slide3.addText('3. Dispute', { x: 7.5, y: 2.3, w: 2.5, fontSize: 20, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });
slide3.addText('Auto-generate evidence for fee mismatch disputes.', { x: 7.7, y: 2.8, w: 2.1, fontSize: 14, color: C.muted, align: 'center', fontFace: 'Helvetica' });

// 4. Feature Highlights
let slide4 = pres.addSlide();
slide4.background = { color: C.bg };
slide4.addText('Key Capabilities', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: C.dark, fontFace: 'Helvetica' });

slide4.addText('Automated Discrepancy Audits', { x: 0.5, y: 1.5, w: 4, fontSize: 20, bold: true, color: C.bright, fontFace: 'Helvetica' });
slide4.addText('Instantly flags when a channel charges a 12% "Jewelry" fee on an 8% "Apparel" item.', { x: 0.5, y: 2.0, w: 4, fontSize: 16, color: C.text, fontFace: 'Helvetica' });

slide4.addText('One-Click Disputes', { x: 5.5, y: 1.5, w: 4, fontSize: 20, bold: true, color: C.orange, fontFace: 'Helvetica' });
slide4.addText('Aggregates PIM category logs as evidence to auto-file disputes via Seller API.', { x: 5.5, y: 2.0, w: 4, fontSize: 16, color: C.text, fontFace: 'Helvetica' });

slide4.addText('SLA Monitoring', { x: 0.5, y: 3.5, w: 4, fontSize: 20, bold: true, color: C.bright, fontFace: 'Helvetica' });
slide4.addText('Tracks external marketplace payout delays and flags breaches in terms of service.', { x: 0.5, y: 4.0, w: 4, fontSize: 16, color: C.text, fontFace: 'Helvetica' });

slide4.addText('Unified Ledger', { x: 5.5, y: 3.5, w: 4, fontSize: 20, bold: true, color: C.bright, fontFace: 'Helvetica' });
slide4.addText('Provides STYLI Finance with a consolidated view of gross GMV vs Net Channel Payout.', { x: 5.5, y: 4.0, w: 4, fontSize: 16, color: C.text, fontFace: 'Helvetica' });

// 5. Seller Impact (PhonePe Parallel)
let slide5 = pres.addSlide();
slide5.background = { color: C.dark };
slide5.addText('Proven Parallel: PhonePe B2B Reconciliation', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: 'FFFFFF', fontFace: 'Helvetica' });
slide5.addText('Managing ₹100 Cr/yr third-party reward budgets:', { x: 0.5, y: 1.2, w: 9, fontSize: 18, color: C.accent, fontFace: 'Helvetica' });

slide5.addShape(pres.ShapeType.rect, { x: 0.5, y: 2.0, w: 4.2, h: 2.5, fill: { color: C.hero } });
slide5.addText('The Challenge', { x: 0.7, y: 2.2, w: 3.8, fontSize: 20, bold: true, color: C.orange, fontFace: 'Helvetica' });
slide5.addText('Reconciling complex marketing budgets, commission payouts, and discount attribution across 500+ brand partners.', { x: 0.7, y: 2.8, w: 3.8, fontSize: 16, color: 'FFFFFF', fontFace: 'Helvetica' });

slide5.addShape(pres.ShapeType.rect, { x: 5.3, y: 2.0, w: 4.2, h: 2.5, fill: { color: C.hero } });
slide5.addText('The Impact', { x: 5.5, y: 2.2, w: 3.8, fontSize: 20, bold: true, color: C.green, fontFace: 'Helvetica' });
slide5.addText('Engineered a self-serve financial dashboard for brands, ensuring 100% accurate dispute resolution and zero revenue leakage.', { x: 5.5, y: 2.8, w: 3.8, fontSize: 16, color: 'FFFFFF', fontFace: 'Helvetica' });

// 6. Success Metrics
let slide6 = pres.addSlide();
slide6.background = { color: C.hero };
slide6.addText('Success Metrics (MaaS OKRs)', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: 'FFFFFF', fontFace: 'Helvetica' });

slide6.addText('100%', { x: 1.0, y: 2.0, w: 3, fontSize: 64, bold: true, color: C.green, fontFace: 'Helvetica' });
slide6.addText('Payout Reconciliation Accuracy', { x: 1.0, y: 3.2, w: 3, fontSize: 18, color: C.muted, fontFace: 'Helvetica' });

slide6.addText('+12%', { x: 4.0, y: 2.0, w: 3, fontSize: 64, bold: true, color: C.bright, fontFace: 'Helvetica' });
slide6.addText('Dispute Recovery Win Rate', { x: 4.0, y: 3.2, w: 3, fontSize: 18, color: C.muted, fontFace: 'Helvetica' });

slide6.addText('2 Hrs', { x: 7.0, y: 2.0, w: 3, fontSize: 64, bold: true, color: C.accent, fontFace: 'Helvetica' });
slide6.addText('Saved Daily per Ops Resource', { x: 7.0, y: 3.2, w: 3, fontSize: 18, color: C.muted, fontFace: 'Helvetica' });

// 7. Closing
let slide7 = pres.addSlide();
slide7.background = { color: C.dark };
slide7.addText('Thank You', { x: 0.5, y: 2.0, w: 9, fontSize: 54, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });
slide7.addText('Ajay Avaghade | Product Manager', { x: 0.5, y: 3.0, w: 9, fontSize: 24, color: C.accent, align: 'center', fontFace: 'Helvetica' });
slide7.addText('avaghadeajay009@gmail.com | linkedin.com/in/ajay-avaghade', { x: 0.5, y: 3.6, w: 9, fontSize: 16, color: C.muted, align: 'center', fontFace: 'Helvetica' });

pres.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/styli-reconhub-deck.pptx' });
