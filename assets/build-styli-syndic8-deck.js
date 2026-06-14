const pptxgen = require('pptxgenjs');

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';

// Styli MaaS Palette
const C = {
    dark: '0F172A',
    hero: '1E293B',
    bright: '3B82F6',
    accent: '60A5FA',
    orange: 'F59E0B',
    bg: 'F8FAFC',
    text: '0F172A',
    muted: '64748B'
};

// Title Slide
let slide1 = pres.addSlide();
slide1.background = { color: C.dark };
slide1.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.dark } });
slide1.addText('Syndic8 Engine', { x: 0.5, y: 2.0, w: 9, fontSize: 54, bold: true, color: 'FFFFFF', fontFace: 'Helvetica' });
slide1.addText('Cross-Channel Listing & SKU Syndication Platform', { x: 0.5, y: 3.0, w: 9, fontSize: 24, color: C.accent, fontFace: 'Helvetica' });
slide1.addText('Prepared for STYLI | MaaS Initiative', { x: 0.5, y: 4.8, w: 9, fontSize: 16, color: C.muted, fontFace: 'Helvetica' });

// 2. The Bottleneck
let slide2 = pres.addSlide();
slide2.background = { color: C.bg };
slide2.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 0.2, h: '100%', fill: { color: C.orange } });
slide2.addText('The Syndication Bottleneck', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: C.dark, fontFace: 'Helvetica' });
slide2.addText('Current seller experience limits STYLI\'s multi-channel velocity:', { x: 0.5, y: 1.2, w: 9, fontSize: 18, color: C.muted, fontFace: 'Helvetica' });
slide2.addText([
    { text: 'Manual Taxonomy Mapping: ', options: { bold: true } },
    { text: 'Sellers must manually translate STYLI attributes to Noon & Amazon schemas.' }
], { x: 0.5, y: 2.0, w: 8, fontSize: 20, color: C.text, bullet: true, fontFace: 'Helvetica' });
slide2.addText([
    { text: 'High Rejection Rates: ', options: { bold: true } },
    { text: 'Missing mandatory fields (e.g., HSN codes) cause metadata failures.' }
], { x: 0.5, y: 2.8, w: 8, fontSize: 20, color: C.text, bullet: true, fontFace: 'Helvetica' });
slide2.addText([
    { text: 'Slow Time-to-Live (TTL): ', options: { bold: true } },
    { text: 'Takes 48+ hours to launch a new collection across all external channels.' }
], { x: 0.5, y: 3.6, w: 8, fontSize: 20, color: C.text, bullet: true, fontFace: 'Helvetica' });

// 3. The Solution
let slide3 = pres.addSlide();
slide3.background = { color: C.hero };
slide3.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 0.2, h: '100%', fill: { color: C.bright } });
slide3.addText('The Solution: Syndic8 Engine', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: 'FFFFFF', fontFace: 'Helvetica' });
slide3.addText('A centralized API hub that maps, transforms, and syndicates listings autonomously.', { x: 0.5, y: 1.2, w: 9, fontSize: 18, color: C.accent, fontFace: 'Helvetica' });

slide3.addShape(pres.ShapeType.roundRect, { x: 0.5, y: 2.0, w: 2.5, h: 2, fill: { color: C.dark }, line: { color: C.bright, width: 2 } });
slide3.addText('1. Connect', { x: 0.5, y: 2.3, w: 2.5, fontSize: 20, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });
slide3.addText('Seller inputs STYLI SKU data once.', { x: 0.7, y: 2.8, w: 2.1, fontSize: 14, color: C.muted, align: 'center', fontFace: 'Helvetica' });

slide3.addShape(pres.ShapeType.rightArrow, { x: 3.2, y: 2.8, w: 0.6, h: 0.4, fill: { color: C.accent } });

slide3.addShape(pres.ShapeType.roundRect, { x: 4.0, y: 2.0, w: 2.5, h: 2, fill: { color: C.dark }, line: { color: C.bright, width: 2 } });
slide3.addText('2. AI Transform', { x: 4.0, y: 2.3, w: 2.5, fontSize: 20, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });
slide3.addText('LLM maps STYLI attributes to channel taxonomies.', { x: 4.2, y: 2.8, w: 2.1, fontSize: 14, color: C.muted, align: 'center', fontFace: 'Helvetica' });

slide3.addShape(pres.ShapeType.rightArrow, { x: 6.7, y: 2.8, w: 0.6, h: 0.4, fill: { color: C.accent } });

slide3.addShape(pres.ShapeType.roundRect, { x: 7.5, y: 2.0, w: 2.5, h: 2, fill: { color: C.dark }, line: { color: C.bright, width: 2 } });
slide3.addText('3. Syndicate', { x: 7.5, y: 2.3, w: 2.5, fontSize: 20, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });
slide3.addText('Parallel API push to Noon, Amazon & Namshi.', { x: 7.7, y: 2.8, w: 2.1, fontSize: 14, color: C.muted, align: 'center', fontFace: 'Helvetica' });

// 4. Feature Highlights
let slide4 = pres.addSlide();
slide4.background = { color: C.bg };
slide4.addText('Key Capabilities', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: C.dark, fontFace: 'Helvetica' });

slide4.addText('Smart Taxonomy Mapping', { x: 0.5, y: 1.5, w: 4, fontSize: 20, bold: true, color: C.bright, fontFace: 'Helvetica' });
slide4.addText('Automatically translates STYLI "Material" to Amazon "FabricType". Zero manual entry required.', { x: 0.5, y: 2.0, w: 4, fontSize: 16, color: C.text, fontFace: 'Helvetica' });

slide4.addText('Idempotent Retries', { x: 5.5, y: 1.5, w: 4, fontSize: 20, bold: true, color: C.bright, fontFace: 'Helvetica' });
slide4.addText('If an API rate limit is hit, the engine queues and safely retries the push without duplicating listings.', { x: 5.5, y: 2.0, w: 4, fontSize: 16, color: C.text, fontFace: 'Helvetica' });

slide4.addText('Pre-flight Validation', { x: 0.5, y: 3.5, w: 4, fontSize: 20, bold: true, color: C.bright, fontFace: 'Helvetica' });
slide4.addText('Checks for missing mandatory fields before pushing to Noon, dropping rejection rates to <1%.', { x: 0.5, y: 4.0, w: 4, fontSize: 16, color: C.text, fontFace: 'Helvetica' });

slide4.addText('Unified Channel Health', { x: 5.5, y: 3.5, w: 4, fontSize: 20, bold: true, color: C.bright, fontFace: 'Helvetica' });
slide4.addText('A single dashboard showing API uptime, live listings, and pending QC per channel.', { x: 5.5, y: 4.0, w: 4, fontSize: 16, color: C.text, fontFace: 'Helvetica' });

// 5. Seller Impact (PhonePe Parallel)
let slide5 = pres.addSlide();
slide5.background = { color: C.dark };
slide5.addText('Proven Parallel: PhonePe B2B Growth', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: 'FFFFFF', fontFace: 'Helvetica' });
slide5.addText('How I solved a similar operational bottleneck for 5,000+ merchants:', { x: 0.5, y: 1.2, w: 9, fontSize: 18, color: C.accent, fontFace: 'Helvetica' });

slide5.addShape(pres.ShapeType.rect, { x: 0.5, y: 2.0, w: 4.2, h: 2.5, fill: { color: C.hero } });
slide5.addText('The Problem', { x: 0.7, y: 2.2, w: 3.8, fontSize: 20, bold: true, color: C.orange, fontFace: 'Helvetica' });
slide5.addText('Merchant discount operations required manual backend mapping, causing a 48-hour Time-to-Market for promotions.', { x: 0.7, y: 2.8, w: 3.8, fontSize: 16, color: 'FFFFFF', fontFace: 'Helvetica' });

slide5.addShape(pres.ShapeType.rect, { x: 5.3, y: 2.0, w: 4.2, h: 2.5, fill: { color: C.hero } });
slide5.addText('The Impact', { x: 5.5, y: 2.2, w: 3.8, fontSize: 20, bold: true, color: C.bright, fontFace: 'Helvetica' });
slide5.addText('Rebuilt the rules engine to automate backend validation. Reduced offer Turnaround Time from 2 days to just 30 minutes.', { x: 5.5, y: 2.8, w: 3.8, fontSize: 16, color: 'FFFFFF', fontFace: 'Helvetica' });

// 6. Architecture
let slide6 = pres.addSlide();
slide6.background = { color: C.bg };
slide6.addText('High-Level Syndication Flow', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: C.dark, fontFace: 'Helvetica' });

slide6.addShape(pres.ShapeType.rect, { x: 0.5, y: 2.0, w: 2, h: 1.5, fill: { color: C.hero } });
slide6.addText('STYLI PIM', { x: 0.5, y: 2.5, w: 2, fontSize: 18, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });

slide6.addShape(pres.ShapeType.rightArrow, { x: 2.7, y: 2.5, w: 0.8, h: 0.4, fill: { color: C.muted } });

slide6.addShape(pres.ShapeType.roundRect, { x: 3.7, y: 1.5, w: 2.6, h: 2.5, fill: { color: C.bright } });
slide6.addText('Syndic8 Hub\n- Transformer\n- Validator\n- Feed Queue', { x: 3.7, y: 2.2, w: 2.6, fontSize: 16, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });

slide6.addShape(pres.ShapeType.rightArrow, { x: 6.5, y: 1.5, w: 0.8, h: 0.2, fill: { color: C.muted } });
slide6.addShape(pres.ShapeType.rightArrow, { x: 6.5, y: 2.65, w: 0.8, h: 0.2, fill: { color: C.muted } });
slide6.addShape(pres.ShapeType.rightArrow, { x: 6.5, y: 3.8, w: 0.8, h: 0.2, fill: { color: C.muted } });

slide6.addShape(pres.ShapeType.rect, { x: 7.5, y: 1.1, w: 2, h: 1, fill: { color: C.dark } });
slide6.addText('Amazon API', { x: 7.5, y: 1.4, w: 2, fontSize: 16, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });

slide6.addShape(pres.ShapeType.rect, { x: 7.5, y: 2.3, w: 2, h: 1, fill: { color: C.dark } });
slide6.addText('Noon API', { x: 7.5, y: 2.6, w: 2, fontSize: 16, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });

slide6.addShape(pres.ShapeType.rect, { x: 7.5, y: 3.5, w: 2, h: 1, fill: { color: C.dark } });
slide6.addText('Namshi Feed', { x: 7.5, y: 3.8, w: 2, fontSize: 16, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });

// 7. Success Metrics
let slide7 = pres.addSlide();
slide7.background = { color: C.hero };
slide7.addText('Success Metrics (MaaS OKRs)', { x: 0.5, y: 0.5, w: 9, fontSize: 32, bold: true, color: 'FFFFFF', fontFace: 'Helvetica' });

slide7.addText('80%', { x: 1.0, y: 2.0, w: 3, fontSize: 64, bold: true, color: C.bright, fontFace: 'Helvetica' });
slide7.addText('Reduction in Time-to-List (TTL)', { x: 1.0, y: 3.2, w: 3, fontSize: 18, color: C.muted, fontFace: 'Helvetica' });

slide7.addText('< 1%', { x: 4.0, y: 2.0, w: 3, fontSize: 64, bold: true, color: C.orange, fontFace: 'Helvetica' });
slide7.addText('Rejection / Defect Rate', { x: 4.0, y: 3.2, w: 3, fontSize: 18, color: C.muted, fontFace: 'Helvetica' });

slide7.addText('100%', { x: 7.0, y: 2.0, w: 3, fontSize: 64, bold: true, color: C.accent, fontFace: 'Helvetica' });
slide7.addText('Automated Multi-Channel Sync', { x: 7.0, y: 3.2, w: 3, fontSize: 18, color: C.muted, fontFace: 'Helvetica' });

// 8. Closing
let slide8 = pres.addSlide();
slide8.background = { color: C.dark };
slide8.addText('Thank You', { x: 0.5, y: 2.0, w: 9, fontSize: 54, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Helvetica' });
slide8.addText('Ajay Avaghade | Product Manager', { x: 0.5, y: 3.0, w: 9, fontSize: 24, color: C.accent, align: 'center', fontFace: 'Helvetica' });
slide8.addText('avaghadeajay009@gmail.com | linkedin.com/in/ajay-avaghade', { x: 0.5, y: 3.6, w: 9, fontSize: 16, color: C.muted, align: 'center', fontFace: 'Helvetica' });

pres.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/styli-syndic8-deck.pptx' });
