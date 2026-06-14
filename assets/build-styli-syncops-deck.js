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
slide1.addText('SyncOps Engine', { x: 0.5, y: 2.0, w: 9, fontSize: 54, bold: true, color: 'FFFFFF', fontFace: 'Helvetica' });
slide1.addText('Real-Time Inventory Orchestration & SLA Protection', { x: 0.5, y: 3.0, w: 9, fontSize: 24, color: C.accent, fontFace: 'Helvetica' });
slide1.addText('Prepared for STYLI | MaaS Initiative', { x: 0.5, y: 4.8, w: 9, fontSize: 16, color: C.muted, fontFace: 'Helvetica' });

// 2. The Bottleneck
let slide2 = pres.addSlide();
slide2.background = { color: C.bg };
slide2.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 0.2, h: '100%', fill: { color: C.red } });
slide2.addText('The Oversell Risk', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: C.dark, fontFace: 'Helvetica' });
slide2.addText('Multi-channel scale breaks manual inventory syncing:', { x: 0.5, y: 1.2, w: 9, fontSize: 18, color: C.muted, fontFace: 'Helvetica' });
slide2.addText([
    { text: 'High Oversell Rates: ', options: { bold: true } },
    { text: 'A single item bought on Noon might still show as available on Amazon.' }
], { x: 0.5, y: 2.0, w: 8, fontSize: 20, color: C.text, bullet: true, fontFace: 'Helvetica' });
slide2.addText([
    { text: 'SLA Penalties: ', options: { bold: true } },
    { text: 'Cancellations lead to account strikes on external marketplaces.' }
], { x: 0.5, y: 2.8, w: 8, fontSize: 20, color: C.text, bullet: true, fontFace: 'Helvetica' });
slide2.addText([
    { text: 'Manual Buffering: ', options: { bold: true } },
    { text: 'Sellers artificially limit stock to avoid overselling, hurting GMV.' }
], { x: 0.5, y: 3.6, w: 8, fontSize: 20, color: C.text, bullet: true, fontFace: 'Helvetica' });

// 3. The Solution
let slide3 = pres.addSlide();
slide3.background = { color: C.hero };
slide3.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 0.2, h: '100%', fill: { color: C.bright } });
slide3.addText('The Solution: SyncOps Engine', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: 'FFFFFF', fontFace: 'Helvetica' });
slide3.addText('A centralized buffer that reserves global stock in sub-second latency.', { x: 0.5, y: 1.2, w: 9, fontSize: 18, color: C.accent, fontFace: 'Helvetica' });

slide3.addShape(pres.ShapeType.roundRect, { x: 0.5, y: 2.0, w: 2.5, h: 2, fill: { color: C.dark }, line: { color: C.bright, width: 2 } });
slide3.addText('1. Ingest', { x: 0.5, y: 2.3, w: 2.5, fontSize: 20, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });
slide3.addText('Order drops on Amazon API.', { x: 0.7, y: 2.8, w: 2.1, fontSize: 14, color: C.muted, align: 'center', fontFace: 'Helvetica' });

slide3.addShape(pres.ShapeType.rightArrow, { x: 3.2, y: 2.8, w: 0.6, h: 0.4, fill: { color: C.accent } });

slide3.addShape(pres.ShapeType.roundRect, { x: 4.0, y: 2.0, w: 2.5, h: 2, fill: { color: C.dark }, line: { color: C.bright, width: 2 } });
slide3.addText('2. Reserve', { x: 4.0, y: 2.3, w: 2.5, fontSize: 20, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });
slide3.addText('SyncOps instantly decrements global master stock.', { x: 4.2, y: 2.8, w: 2.1, fontSize: 14, color: C.muted, align: 'center', fontFace: 'Helvetica' });

slide3.addShape(pres.ShapeType.rightArrow, { x: 6.7, y: 2.8, w: 0.6, h: 0.4, fill: { color: C.accent } });

slide3.addShape(pres.ShapeType.roundRect, { x: 7.5, y: 2.0, w: 2.5, h: 2, fill: { color: C.dark }, line: { color: C.bright, width: 2 } });
slide3.addText('3. Broadcast', { x: 7.5, y: 2.3, w: 2.5, fontSize: 20, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });
slide3.addText('New stock level pushed to STYLI & Noon.', { x: 7.7, y: 2.8, w: 2.1, fontSize: 14, color: C.muted, align: 'center', fontFace: 'Helvetica' });

// 4. Feature Highlights
let slide4 = pres.addSlide();
slide4.background = { color: C.bg };
slide4.addText('Key Capabilities', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: C.dark, fontFace: 'Helvetica' });

slide4.addText('Sub-Second Orchestration', { x: 0.5, y: 1.5, w: 4, fontSize: 20, bold: true, color: C.bright, fontFace: 'Helvetica' });
slide4.addText('Webhooks and polling ensure <800ms latency between an order drop and global stock decrement.', { x: 0.5, y: 2.0, w: 4, fontSize: 16, color: C.text, fontFace: 'Helvetica' });

slide4.addText('Safety Buffers', { x: 5.5, y: 1.5, w: 4, fontSize: 20, bold: true, color: C.orange, fontFace: 'Helvetica' });
slide4.addText('Sellers can set dynamic buffers (e.g. pause Amazon when stock hits 2) to prevent overlapping checkouts.', { x: 5.5, y: 2.0, w: 4, fontSize: 16, color: C.text, fontFace: 'Helvetica' });

slide4.addText('Exception Routing', { x: 0.5, y: 3.5, w: 4, fontSize: 20, bold: true, color: C.bright, fontFace: 'Helvetica' });
slide4.addText('If an item is oversold, the system automatically routes to secondary warehouses before cancelling.', { x: 0.5, y: 4.0, w: 4, fontSize: 16, color: C.text, fontFace: 'Helvetica' });

slide4.addText('Ops Central Dashboard', { x: 5.5, y: 3.5, w: 4, fontSize: 20, bold: true, color: C.bright, fontFace: 'Helvetica' });
slide4.addText('Real-time visibility into channel performance, SLA breaches, and order sync states.', { x: 5.5, y: 4.0, w: 4, fontSize: 16, color: C.text, fontFace: 'Helvetica' });

// 5. Seller Impact (PhonePe Parallel)
let slide5 = pres.addSlide();
slide5.background = { color: C.dark };
slide5.addText('Proven Parallel: PhonePe High-Frequency Scale', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: 'FFFFFF', fontFace: 'Helvetica' });
slide5.addText('Managing concurrent transaction orchestration for 350M+ MAU:', { x: 0.5, y: 1.2, w: 9, fontSize: 18, color: C.accent, fontFace: 'Helvetica' });

slide5.addShape(pres.ShapeType.rect, { x: 0.5, y: 2.0, w: 4.2, h: 2.5, fill: { color: C.hero } });
slide5.addText('The Challenge', { x: 0.7, y: 2.2, w: 3.8, fontSize: 20, bold: true, color: C.orange, fontFace: 'Helvetica' });
slide5.addText('Handling high-frequency payment gateway checkouts without transaction race conditions.', { x: 0.7, y: 2.8, w: 3.8, fontSize: 16, color: 'FFFFFF', fontFace: 'Helvetica' });

slide5.addShape(pres.ShapeType.rect, { x: 5.3, y: 2.0, w: 4.2, h: 2.5, fill: { color: C.hero } });
slide5.addText('The Impact', { x: 5.5, y: 2.2, w: 3.8, fontSize: 20, bold: true, color: C.green, fontFace: 'Helvetica' });
slide5.addText('Built idempotent state management systems ensuring 0% duplicate charges, perfectly paralleling cross-channel inventory locking.', { x: 5.5, y: 2.8, w: 3.8, fontSize: 16, color: 'FFFFFF', fontFace: 'Helvetica' });

// 6. Success Metrics
let slide6 = pres.addSlide();
slide6.background = { color: C.hero };
slide6.addText('Success Metrics (MaaS OKRs)', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: 'FFFFFF', fontFace: 'Helvetica' });

slide6.addText('0%', { x: 1.0, y: 2.0, w: 3, fontSize: 64, bold: true, color: C.green, fontFace: 'Helvetica' });
slide6.addText('Oversell Rate', { x: 1.0, y: 3.2, w: 3, fontSize: 18, color: C.muted, fontFace: 'Helvetica' });

slide6.addText('< 1s', { x: 4.0, y: 2.0, w: 3, fontSize: 64, bold: true, color: C.bright, fontFace: 'Helvetica' });
slide6.addText('Cross-Channel Stock Sync', { x: 4.0, y: 3.2, w: 3, fontSize: 18, color: C.muted, fontFace: 'Helvetica' });

slide6.addText('100%', { x: 7.0, y: 2.0, w: 3, fontSize: 64, bold: true, color: C.accent, fontFace: 'Helvetica' });
slide6.addText('SLA Protection', { x: 7.0, y: 3.2, w: 3, fontSize: 18, color: C.muted, fontFace: 'Helvetica' });

// 7. Closing
let slide7 = pres.addSlide();
slide7.background = { color: C.dark };
slide7.addText('Thank You', { x: 0.5, y: 2.0, w: 9, fontSize: 54, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });
slide7.addText('Ajay Avaghade | Product Manager', { x: 0.5, y: 3.0, w: 9, fontSize: 24, color: C.accent, align: 'center', fontFace: 'Helvetica' });
slide7.addText('avaghadeajay009@gmail.com | linkedin.com/in/ajay-avaghade', { x: 0.5, y: 3.6, w: 9, fontSize: 16, color: C.muted, align: 'center', fontFace: 'Helvetica' });

pres.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/styli-syncops-deck.pptx' });
