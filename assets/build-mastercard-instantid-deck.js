const PptxGenJS = require('pptxgenjs');

const prs = new PptxGenJS();
prs.defineLayout({ name: 'LAYOUT_16x9', width: 10, height: 7.5 });
prs.layout = 'LAYOUT_16x9';

const C = {
  dark: '0A1F2E',
  hero: '152A3D',
  red: 'FF5F00',
  gold: 'D4AF37',
  white: 'FFFFFF',
  lgray: 'F5F5F5',
  muted: '8B92A5',
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 3, offset: 1, angle: 45, opacity: 0.12 });

// SLIDE 1: COVER
let slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 1.2, fill: { color: C.hero } });
slide.addText('MASTERCARD · EEMEA EXPANSION', { x: 0.5, y: 0.25, w: 9, h: 0.7, fontSize: 9, bold: true, color: C.red, align: 'center', charSpacing: 2 });
slide.addText('InstantID', { x: 1, y: 2, w: 8, h: 1.2, fontSize: 52, bold: true, color: C.white, align: 'center', fontFace: 'Sora' });
slide.addText('SME Self-Serve KYC Platform', { x: 1, y: 3.3, w: 8, h: 0.5, fontSize: 24, color: C.red, align: 'center' });
slide.addText('Ajay Avaghade · Product Manager', { x: 1, y: 4, w: 8, h: 0.4, fontSize: 13, color: C.muted, align: 'center', italic: true });
slide.addShape(prs.ShapeType.rect, { x: 6.5, y: 5.8, w: 3, h: 1.2, fill: { color: C.hero }, shadow: makeShadow() });
slide.addText('TAT: 5 days\n→ 4 hours', { x: 6.5, y: 5.85, w: 3, h: 0.5, fontSize: 28, bold: true, color: C.red, align: 'center' });
slide.addText('98% reduction', { x: 6.5, y: 6.35, w: 3, h: 0.4, fontSize: 11, color: C.muted, align: 'center' });

// SLIDE 2: THE PROBLEM
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.hero } });
slide.addText('THE PROBLEM', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.red, align: 'left', charSpacing: 2 });
slide.addText('SME KYC Takes 5–7 Days & Costs Issuers 35% CAC', { x: 0.5, y: 1.1, w: 9, h: 0.8, fontSize: 30, bold: true, color: C.white });

// Stat columns
const statX = [0.5, 3.5, 6.5];
const statLabels = ['📋\n\n5–7 Days', '💰\n\n35% CAC', '🏢\n\n70% Staff'];
const statDescs = ['Onboarding TAT\nper SME', 'Issuer cost\nper acquisition', 'Staff time on\ndocument collection'];

statX.forEach((x, i) => {
  slide.addShape(prs.ShapeType.rect, { x, y: 2.2, w: 2.8, h: 2, fill: { color: C.hero }, border: { pt: 1, color: C.red } });
  slide.addText(statLabels[i], { x: x + 0.2, y: 2.4, w: 2.4, h: 1, fontSize: 20, bold: true, color: C.red, align: 'center' });
  slide.addText(statDescs[i], { x: x + 0.2, y: 3.6, w: 2.4, h: 0.5, fontSize: 10, color: C.muted, align: 'center' });
});

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 4.5, w: 9, h: 1.8, fill: { color: C.hero }, border: { pt: 1, color: C.gold } });
slide.addText('Root Cause: Manual KYC process puts burden on issuer staff. SMEs abandon midway (40–60% drop-off). Issuers deprioritize SME segment due to unclear ROI.', { x: 0.8, y: 4.7, w: 8.4, h: 1.4, fontSize: 12, color: C.white, align: 'left' });

// SLIDE 3: THE INSIGHT
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.dark } });
slide.addText('THE INSIGHT', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.red, align: 'left', charSpacing: 2 });
slide.addText('Shift KYC Burden from Issuer Staff to AI', { x: 0.5, y: 1.1, w: 9, h: 0.8, fontSize: 28, bold: true, color: C.dark });

// Left side (Status Quo)
slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2.1, w: 4.2, h: 4.8, fill: { color: C.white }, border: { pt: 2, color: C.dark } });
slide.addText('❌ Status Quo', { x: 0.7, y: 2.3, w: 3.8, h: 0.4, fontSize: 14, bold: true, color: C.dark });
const statusQuo = ['Manual document review\nby issuer staff', 'Back-and-forth emails\nfor clarifications', 'Slow, human-error prone\nKYC verification', 'SME abandons mid-process'];
let yPos = 2.8;
statusQuo.forEach(item => {
  slide.addText(item, { x: 0.7, y: yPos, w: 3.8, h: 0.6, fontSize: 11, color: C.dark });
  yPos += 0.7;
});

// Right side (Proposed)
slide.addShape(prs.ShapeType.rect, { x: 5.3, y: 2.1, w: 4.2, h: 4.8, fill: { color: C.hero }, border: { pt: 2, color: C.red } });
slide.addText('✅ InstantID', { x: 5.5, y: 2.3, w: 3.8, h: 0.4, fontSize: 14, bold: true, color: C.white });
const proposed = ['AI-powered video KYC\n& OCR verification', 'Real-time document\nvalidation', 'Instant approval decision\nin 2–4 hours', 'Zero abandonment,\nhigh conversion'];
yPos = 2.8;
proposed.forEach(item => {
  slide.addText(item, { x: 5.5, y: yPos, w: 3.8, h: 0.6, fontSize: 11, color: C.white });
  yPos += 0.7;
});

// SLIDE 4: THE MECHANIC
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.hero } });
slide.addText('THE MECHANIC', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.gold, align: 'left', charSpacing: 2 });
slide.addText('Video + AI OCR → Instant Approval → Issuer Exception Handler', { x: 0.5, y: 1.1, w: 9, h: 0.8, fontSize: 26, bold: true, color: C.white });

// Timeline steps
const steps = [
  { num: '1', title: 'SME Initiation', desc: 'Downloads app, selects card type, starts KYC' },
  { num: '2', title: 'Video + Docs', desc: 'Records selfie video, uploads invoices/tax docs' },
  { num: '3', title: 'AI Verification', desc: 'OCR scans docs, validates against registries in real-time' },
  { num: '4', title: 'Instant Decision', desc: 'Approval given in 2–4 hours, virtual card issued' },
  { num: '5', title: 'Issuer Dashboard', desc: 'Exception handling only, staff handles edge cases' }
];

const stepX = [0.5, 2.3, 4.1, 5.9, 7.7];
const stepW = 1.7;
steps.forEach((step, i) => {
  slide.addShape(prs.ShapeType.rect, { x: stepX[i], y: 2.2, w: stepW, h: 4.2, fill: { color: C.hero }, border: { pt: 1, color: C.gold } });
  slide.addText(step.num, { x: stepX[i], y: 2.3, w: stepW, h: 0.35, fontSize: 20, bold: true, color: C.gold, align: 'center' });
  slide.addText(step.title, { x: stepX[i] + 0.1, y: 2.8, w: stepW - 0.2, h: 0.6, fontSize: 11, bold: true, color: C.white, align: 'center' });
  slide.addText(step.desc, { x: stepX[i] + 0.1, y: 3.5, w: stepW - 0.2, h: 2, fontSize: 9, color: C.muted, align: 'center' });
  if (i < steps.length - 1) {
    slide.addShape(prs.ShapeType.line, { x: stepX[i] + stepW + 0.05, y: 4.3, w: 0, h: 0, line: { color: C.gold, dashType: 'dash' } });
  }
});

slide.addText('A/B tested at PhonePe: Similar shift (self-serve PG platform) drove 6-week → 2-day TAT', { x: 0.5, y: 6.8, w: 9, h: 0.5, fontSize: 11, color: C.gold, italic: true });

// SLIDE 5: THE PRODUCT
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.dark } });
slide.addText('THE PRODUCT', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.red, align: 'left', charSpacing: 2 });
slide.addText('4 Screen States · End-to-End SME Journey', { x: 0.5, y: 1.1, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.dark });

// 4 product cards
const screens = [
  { num: '01', title: 'Card Selection', desc: 'Choose product (Debit/Credit), view benefits, initiate KYC' },
  { num: '02', title: 'KYC & Verification', desc: 'Video selfie + document uploads; AI validates in real-time' },
  { num: '03', title: 'Instant Approval', desc: 'Approval decision, virtual card issued, copy card number' },
  { num: '04', title: 'Issuer Dashboard', desc: 'Issuer pipeline view, compliance status, exception queue' }
];

const cardX = [0.5, 2.7, 4.9, 7.1];
const cardW = 2.1;
screens.forEach((s, i) => {
  slide.addShape(prs.ShapeType.rect, { x: cardX[i], y: 2.1, w: cardW, h: 4.8, fill: { color: C.white }, border: { pt: 2, color: C.red } });
  slide.addText(s.num, { x: cardX[i], y: 2.3, w: cardW, h: 0.35, fontSize: 16, bold: true, color: C.red, align: 'center' });
  slide.addText(s.title, { x: cardX[i] + 0.15, y: 2.8, w: cardW - 0.3, h: 0.5, fontSize: 12, bold: true, color: C.dark, align: 'center' });
  slide.addText(s.desc, { x: cardX[i] + 0.15, y: 3.5, w: cardW - 0.3, h: 3.2, fontSize: 9, color: C.dark, align: 'center' });
});

slide.addText('Live prototype: mastercard-instantid-prototype.html', { x: 0.5, y: 7, w: 9, h: 0.35, fontSize: 10, color: C.muted, italic: true });

// SLIDE 6: IMPACT & ROI
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.hero } });
slide.addText('IMPACT & ROI', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.gold, align: 'left', charSpacing: 2 });
slide.addText('Built on PhonePe Proof Points — Adjusted for Mastercard Scale', { x: 0.5, y: 1.1, w: 9, h: 0.6, fontSize: 22, bold: true, color: C.white });

// User/Issuer Impact
slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2, w: 4.3, h: 4.6, fill: { color: C.hero }, border: { pt: 1, color: C.red } });
slide.addText('SME & Issuer Impact', { x: 0.7, y: 2.15, w: 3.9, h: 0.35, fontSize: 11, bold: true, color: C.red });
const impacts = ['↓ 98% TAT reduction\n(5d → 4h)', '↓ 35% CAC reduction\nper SME', '↑ 65% activation rate\nat 30 days', '↑ 70% staff capacity freed'];
let impactY = 2.65;
impacts.forEach(item => {
  slide.addText(item, { x: 0.7, y: impactY, w: 3.9, h: 0.8, fontSize: 11, bold: true, color: C.white });
  impactY += 1;
});

// Mastercard ROI
slide.addShape(prs.ShapeType.rect, { x: 5.2, y: 2, w: 4.3, h: 4.6, fill: { color: C.hero }, border: { pt: 1, color: C.gold } });
slide.addText('Mastercard ROI', { x: 5.4, y: 2.15, w: 3.9, h: 0.35, fontSize: 11, bold: true, color: C.gold });
const rois = ['10K SMEs onboarded\nby year-end', 'AED 0 platform cost\n(issuer-funded)', 'Network differentiation\nin SME segment', 'Replicable across\nall 47 EEMEA issuers'];
let roiY = 2.65;
rois.forEach(item => {
  slide.addText(item, { x: 5.4, y: roiY, w: 3.9, h: 0.8, fontSize: 11, bold: true, color: C.white });
  roiY += 1;
});

// SLIDE 7: PROOF OF WORK
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.dark } });
slide.addText('PROOF OF WORK', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.red, align: 'left', charSpacing: 2 });
slide.addText("I Built This at PhonePe. Here's the Direct Map.", { x: 0.5, y: 1.1, w: 9, h: 0.6, fontSize: 22, bold: true, color: C.dark });

// Left: What I built
slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2, w: 4.5, h: 4.8, fill: { color: C.dark }, border: { pt: 1, color: C.red } });
slide.addText('What I Built at PhonePe', { x: 0.7, y: 2.15, w: 4.1, h: 0.35, fontSize: 12, bold: true, color: C.red });
const builtItems = [
  '5K merchant self-serve onboarding platform',
  '6 weeks → 2 days TAT',
  '23% lower merchant CAC',
  'Zero engineering overhead per merchant'
];
let builtY = 2.65;
builtItems.forEach(item => {
  slide.addText(item, { x: 0.7, y: builtY, w: 4.1, h: 0.8, fontSize: 10, color: C.white });
  builtY += 1;
});

// Right: Maps to InstantID
slide.addShape(prs.ShapeType.rect, { x: 5.5, y: 2, w: 4.3, h: 4.8, fill: { color: C.red }, border: { pt: 1, color: C.red } });
slide.addText('→ InstantID Architecture', { x: 5.7, y: 2.15, w: 3.9, h: 0.35, fontSize: 12, bold: true, color: C.white });
const mapsItems = [
  'Self-serve SME KYC platform',
  '5 days → 4 hours TAT',
  '35% lower SME CAC',
  'Zero issuer staff overhead'
];
let mapY = 2.65;
mapsItems.forEach((item, i) => {
  slide.addText(item, { x: 5.7, y: mapY, w: 3.9, h: 0.8, fontSize: 10, color: C.white });
  mapY += 1;
});

// SLIDE 8: ROLLOUT PLAN
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.hero } });
slide.addText('ROLLOUT PLAN', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.gold, align: 'left', charSpacing: 2 });
slide.addText('Three Phases — Zero Risk Pilot to Full EEMEA Rollout', { x: 0.5, y: 1.1, w: 9, h: 0.6, fontSize: 22, bold: true, color: C.white });

// Phase cards
const phases = [
  { title: 'Month 1–2: Pilot', desc: '500 SMEs, 1 issuer, 1 geographic market. Instrument health score. Measure TAT, activation, CAC vs. control.' },
  { title: 'Month 3–4: Scale + Learn', desc: '5 issuers, 5K SMEs. Run TAT/activation/CAC A/A test. Tune AI thresholds based on pilot data.' },
  { title: 'Month 5–6: EEMEA Rollout', desc: 'All 47 EEMEA issuers. Roll out to new SME cohorts across all markets. Scale to 25K+ SMEs.' }
];

const phaseX = [0.5, 3.4, 6.3];
const phaseW = 2.8;
phases.forEach((phase, i) => {
  slide.addShape(prs.ShapeType.rect, { x: phaseX[i], y: 2.1, w: phaseW, h: 4, fill: { color: C.hero }, border: { pt: 1, color: C.gold } });
  slide.addText(phase.title, { x: phaseX[i] + 0.15, y: 2.3, w: phaseW - 0.3, h: 0.5, fontSize: 11, bold: true, color: C.gold });
  slide.addText(phase.desc, { x: phaseX[i] + 0.15, y: 2.95, w: phaseW - 0.3, h: 3.1, fontSize: 9, color: C.white });
});

// Ask box
slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 6.3, w: 9, h: 0.9, fill: { color: C.hero }, border: { pt: 2, color: C.gold } });
slide.addText('What I Need: Issuer transaction data access · 1–2 engineers (KYC API + AI integration) · Alignment with EEMEA regional delivery & issuer stakeholders', { x: 0.7, y: 6.4, w: 8.6, h: 0.7, fontSize: 11, bold: true, color: C.gold });

// Footer
slide.addText('Ajay Avaghade | avaghadeajay009@gmail.com | +91 9561558439', { x: 0.5, y: 7.3, w: 9, h: 0.2, fontSize: 9, color: C.muted, align: 'center' });

prs.writeFile({ fileName: 'assets/mastercard-instantid-deck.pptx' }).then(() => {
  console.log('✓ mastercard-instantid-deck.pptx generated');
});
