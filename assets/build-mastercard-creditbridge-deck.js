const PptxGenJS = require('pptxgenjs');

const prs = new PptxGenJS();
prs.defineLayout({ name: 'LAYOUT_16x9', width: 10, height: 7.5 });
prs.layout = 'LAYOUT_16x9';

const C = {
  dark: '0A1F2E',
  hero: '152A3D',
  blue: '1E90FF',
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
slide.addText('MASTERCARD · SME ENGAGEMENT + CREDIT', { x: 0.5, y: 0.25, w: 9, h: 0.7, fontSize: 9, bold: true, color: C.blue, align: 'center', charSpacing: 2 });
slide.addText('CreditBridge', { x: 1, y: 2, w: 8, h: 1.2, fontSize: 52, bold: true, color: C.white, align: 'center', fontFace: 'Sora' });
slide.addText('Milestone Credit + Spend-Based Lending', { x: 1, y: 3.3, w: 8, h: 0.5, fontSize: 24, color: C.blue, align: 'center' });
slide.addText('Ajay Avaghade · Product Manager', { x: 1, y: 4, w: 8, h: 0.4, fontSize: 13, color: C.muted, align: 'center', italic: true });
slide.addShape(prs.ShapeType.rect, { x: 6.5, y: 5.8, w: 3, h: 1.2, fill: { color: C.hero }, shadow: makeShadow() });
slide.addText('12K SMEs\nAccess Credit', { x: 6.5, y: 5.85, w: 3, h: 0.5, fontSize: 24, bold: true, color: C.blue, align: 'center' });
slide.addText('AED 6.5M origination revenue', { x: 6.5, y: 6.35, w: 3, h: 0.4, fontSize: 11, color: C.muted, align: 'center' });

// SLIDE 2: THE PROBLEM
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.hero } });
slide.addText('THE PROBLEM', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.blue, align: 'left', charSpacing: 2 });
slide.addText("SMEs Don't Activate Repeat Spending; Can't Access Working Capital", { x: 0.5, y: 1.1, w: 9, h: 0.8, fontSize: 30, bold: true, color: C.white });

// Stat columns
const statX = [0.5, 3.5, 6.5];
const statLabels = ['📊\n\n40%', '💳\n\n0%', '💰\n\n0'];
const statDescs = ['30-day SME\ncard activation rate', 'SMEs with access to\nworking capital', 'Mastercard SME\ncredit revenue'];

statX.forEach((x, i) => {
  slide.addShape(prs.ShapeType.rect, { x, y: 2.2, w: 2.8, h: 2, fill: { color: C.hero }, border: { pt: 1, color: C.blue } });
  slide.addText(statLabels[i], { x: x + 0.2, y: 2.4, w: 2.4, h: 1, fontSize: 20, bold: true, color: C.blue, align: 'center' });
  slide.addText(statDescs[i], { x: x + 0.2, y: 3.6, w: 2.4, h: 0.5, fontSize: 10, color: C.muted, align: 'center' });
});

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 4.5, w: 9, h: 1.8, fill: { color: C.hero }, border: { pt: 1, color: C.gold } });
slide.addText("Root Cause: SMEs don't spend repeatedly (low activation) because no incentive. SMEs need working capital but can't prove creditworthiness without formal history.", { x: 0.8, y: 4.7, w: 8.4, h: 1.4, fontSize: 12, color: C.white, align: 'left' });

// SLIDE 3: THE INSIGHT
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.dark } });
slide.addText('THE INSIGHT', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.blue, align: 'left', charSpacing: 2 });
slide.addText('Milestone Credit is Zero-Cost When Spend-Triggered', { x: 0.5, y: 1.1, w: 9, h: 0.8, fontSize: 28, bold: true, color: C.dark });

// Left side
slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2.1, w: 4.2, h: 4.8, fill: { color: C.white }, border: { pt: 2, color: C.dark } });
slide.addText('The Problem', { x: 0.7, y: 2.3, w: 3.8, h: 0.4, fontSize: 14, bold: true, color: C.dark });
const statusQuo = ['Issuers want to drive\ncard spend but discounting\nerodes margins', 'SMEs need working\ncapital but have\nno credit history', 'Traditional lending\nneeds formal credit\nhistory (lacking)'];
let yPos = 2.8;
statusQuo.forEach(item => {
  slide.addText(item, { x: 0.7, y: yPos, w: 3.8, h: 0.7, fontSize: 11, color: C.dark });
  yPos += 0.85;
});

// Right side
slide.addShape(prs.ShapeType.rect, { x: 5.3, y: 2.1, w: 4.2, h: 4.8, fill: { color: C.hero }, border: { pt: 2, color: C.blue } });
slide.addText('✅ CreditBridge Magic', { x: 5.5, y: 2.3, w: 3.8, h: 0.4, fontSize: 14, bold: true, color: C.white });
const proposed = ['Milestone credit only\nfires AFTER SME\nhits spend goal', 'Issuer margin generated\nby SME spend pays\nfor the credit', 'SME card spend data =\nproof of creditworthiness\nfor lenders'];
yPos = 2.8;
proposed.forEach(item => {
  slide.addText(item, { x: 5.5, y: yPos, w: 3.8, h: 0.7, fontSize: 11, color: C.white });
  yPos += 0.85;
});

// SLIDE 4: THE MECHANIC
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.hero } });
slide.addText('THE MECHANIC', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.gold, align: 'left', charSpacing: 2 });
slide.addText('Milestone (Drive Spend) → Credit Readiness (90d) → MasterCredit Lending → Lender Origination', { x: 0.5, y: 1.1, w: 9, h: 0.8, fontSize: 24, bold: true, color: C.white });

// Timeline steps
const steps = [
  { num: '1', title: 'Enrollment', desc: 'SME activates CreditBridge. Milestone window opens: hit AED 500 spend in 90 days to unlock AED 50 credit.' },
  { num: '2', title: 'Milestone Drive', desc: 'Progress bar visible in dashboard. Nudges at 50%, 80%, 95% to goal. Issuer margin covers credit cost.' },
  { num: '3', title: 'Credit Readiness', desc: 'After 90d of spend history, SME becomes credit-eligible. ML score: 72/100 (acceptable lending risk).' },
  { num: '4', title: 'Pre-Qualification', desc: 'Mastercard shares spend data with lenders. Pre-qual amount: AED 5K–50K based on velocity/seasonality.' },
  { num: '5', title: 'Lender Origination', desc: 'SME applies. Lender approves in 24h. Credit deployed. Mastercard earns origination fee.' }
];

const stepX = [0.5, 2.3, 4.1, 5.9, 7.7];
const stepW = 1.7;
steps.forEach((step, i) => {
  slide.addShape(prs.ShapeType.rect, { x: stepX[i], y: 2.2, w: stepW, h: 4.2, fill: { color: C.hero }, border: { pt: 1, color: C.blue } });
  slide.addText(step.num, { x: stepX[i], y: 2.3, w: stepW, h: 0.35, fontSize: 20, bold: true, color: C.blue, align: 'center' });
  slide.addText(step.title, { x: stepX[i] + 0.1, y: 2.8, w: stepW - 0.2, h: 0.6, fontSize: 11, bold: true, color: C.white, align: 'center' });
  slide.addText(step.desc, { x: stepX[i] + 0.1, y: 3.5, w: stepW - 0.2, h: 2, fontSize: 8, color: C.muted, align: 'center' });
});

slide.addText('PhonePe proof: Milestone rental waiver (60% activation, zero net cost) + propensity ML for lending', { x: 0.5, y: 6.8, w: 9, h: 0.5, fontSize: 11, color: C.gold, italic: true });

// SLIDE 5: THE PRODUCT
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.dark } });
slide.addText('THE PRODUCT', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.blue, align: 'left', charSpacing: 2 });
slide.addText('4 Screen States · SME Dashboard to Ops Console', { x: 0.5, y: 1.1, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.dark });

// 4 product cards
const screens = [
  { num: '01', title: 'Card Dashboard', desc: 'Milestone progress bars, next reward unlock, available balance' },
  { num: '02', title: 'Milestone Tiers', desc: 'Bronze (AED 500), Silver (AED 2K), Gold (AED 5K), Platinum (AED 10K)' },
  { num: '03', title: 'Credit Readiness', desc: 'Spend score (72/100), eligibility status, pre-qual amount, apply button' },
  { num: '04', title: 'Ops Dashboard', desc: 'Issuer + Lender views: activation rate, spend lift, loan originations' }
];

const cardX = [0.5, 2.7, 4.9, 7.1];
const cardW = 2.1;
screens.forEach((s, i) => {
  slide.addShape(prs.ShapeType.rect, { x: cardX[i], y: 2.1, w: cardW, h: 4.8, fill: { color: C.white }, border: { pt: 2, color: C.blue } });
  slide.addText(s.num, { x: cardX[i], y: 2.3, w: cardW, h: 0.35, fontSize: 16, bold: true, color: C.blue, align: 'center' });
  slide.addText(s.title, { x: cardX[i] + 0.15, y: 2.8, w: cardW - 0.3, h: 0.5, fontSize: 12, bold: true, color: C.dark, align: 'center' });
  slide.addText(s.desc, { x: cardX[i] + 0.15, y: 3.5, w: cardW - 0.3, h: 3.2, fontSize: 9, color: C.dark, align: 'center' });
});

slide.addText('Live prototype: mastercard-creditbridge-prototype.html', { x: 0.5, y: 7, w: 9, h: 0.35, fontSize: 10, color: C.muted, italic: true });

// SLIDE 6: IMPACT & ROI
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.hero } });
slide.addText('IMPACT & ROI', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.gold, align: 'left', charSpacing: 2 });
slide.addText('SME Activation + Credit Access + Multiple Stakeholder Value', { x: 0.5, y: 1.1, w: 9, h: 0.6, fontSize: 22, bold: true, color: C.white });

// SME Impact
slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2, w: 2.9, h: 4.6, fill: { color: C.hero }, border: { pt: 1, color: C.blue } });
slide.addText('SME Impact', { x: 0.7, y: 2.15, w: 2.5, h: 0.35, fontSize: 11, bold: true, color: C.blue });
const smeinpacts = ['↑ 65% activation\nat 30 days', '↑ 70% avg spend\nper SME', '↑ 12K SMEs with\ncredit access'];
let smeY = 2.65;
smeinpacts.forEach(item => {
  slide.addText(item, { x: 0.7, y: smeY, w: 2.5, h: 0.95, fontSize: 10, bold: true, color: C.white });
  smeY += 1.2;
});

// Issuer Impact
slide.addShape(prs.ShapeType.rect, { x: 3.5, y: 2, w: 2.9, h: 4.6, fill: { color: C.hero }, border: { pt: 1, color: C.blue } });
slide.addText('Issuer Impact', { x: 3.7, y: 2.15, w: 2.5, h: 0.35, fontSize: 11, bold: true, color: C.blue });
const issuerimpacts = ['↑ Revenue per SME\nvia milestone spend', '↓ CAC payback\nin 90 days', 'Zero upfront cost\n(margin covers)'];
let issuerY = 2.65;
issuerimpacts.forEach(item => {
  slide.addText(item, { x: 3.7, y: issuerY, w: 2.5, h: 0.95, fontSize: 10, bold: true, color: C.white });
  issuerY += 1.2;
});

// Mastercard ROI
slide.addShape(prs.ShapeType.rect, { x: 6.6, y: 2, w: 2.9, h: 4.6, fill: { color: C.hero }, border: { pt: 1, color: C.gold } });
slide.addText('Mastercard ROI', { x: 6.8, y: 2.15, w: 2.5, h: 0.35, fontSize: 11, bold: true, color: C.gold });
const mcimpacts = ['AED 6.5M\norigination fee/Q1', 'New revenue from\nlender partnerships', 'SME lock-in via\nfinancial products'];
let mcY = 2.65;
mcimpacts.forEach(item => {
  slide.addText(item, { x: 6.8, y: mcY, w: 2.5, h: 0.95, fontSize: 10, bold: true, color: C.white });
  mcY += 1.2;
});

// SLIDE 7: PROOF OF WORK
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.dark } });
slide.addText('PROOF OF WORK', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.blue, align: 'left', charSpacing: 2 });
slide.addText("I Built This at PhonePe. Here's the Direct Map.", { x: 0.5, y: 1.1, w: 9, h: 0.6, fontSize: 22, bold: true, color: C.dark });

// Left: PhonePe
slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2, w: 4.5, h: 4.8, fill: { color: C.dark }, border: { pt: 1, color: C.blue } });
slide.addText('PhonePe Experience', { x: 0.7, y: 2.15, w: 4.1, h: 0.35, fontSize: 12, bold: true, color: C.blue });
const builtItems = [
  'Milestone rental waiver for merchant activation',
  '60% improvement in 90-day activation',
  'Zero net cost to issuer (margin pays for it)',
  'Propensity ML for merchant credit scoring'
];
let builtY = 2.65;
builtItems.forEach(item => {
  slide.addText(item, { x: 0.7, y: builtY, w: 4.1, h: 0.8, fontSize: 9, color: C.white });
  builtY += 1;
});

// Right: CreditBridge
slide.addShape(prs.ShapeType.rect, { x: 5.5, y: 2, w: 4.3, h: 4.8, fill: { color: C.blue }, border: { pt: 1, color: C.blue } });
slide.addText('→ CreditBridge Architecture', { x: 5.7, y: 2.15, w: 3.9, h: 0.35, fontSize: 12, bold: true, color: C.dark });
const mapsItems = [
  'Milestone credit for SME card spend',
  '65% activation at 30 days',
  'Issuer margin generates credit credit cost',
  'Spend-based credit eligibility for lending'
];
let mapY = 2.65;
mapsItems.forEach((item, i) => {
  slide.addText(item, { x: 5.7, y: mapY, w: 3.9, h: 0.8, fontSize: 9, color: C.dark });
  mapY += 1;
});

// SLIDE 8: ROLLOUT PLAN
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.hero } });
slide.addText('ROLLOUT PLAN', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.gold, align: 'left', charSpacing: 2 });
slide.addText('Three Phases — Milestone Pilot to Full Credit Ecosystem', { x: 0.5, y: 1.1, w: 9, h: 0.6, fontSize: 22, bold: true, color: C.white });

// Phase cards
const phases = [
  { title: 'Month 1–2: Milestone Pilot', desc: '1K SMEs, 1 issuer. Instrument milestone tracking, measure activation rate vs. control. Baseline: 40% hit rate.' },
  { title: 'Month 3–4: Add MasterCredit', desc: '3 lenders pre-vetted. SMEs with 90d history eligible. Run A/B test: Milestone → Credit vs. control. Measure CAC payback.' },
  { title: 'Month 5–6: EEMEA Scale', desc: 'Roll out to all issuers, all lender partners. Scale to 50K+ SMEs. Milestone + MasterCredit as standard SME product.' }
];

const phaseX = [0.5, 3.4, 6.3];
const phaseW = 2.8;
phases.forEach((phase, i) => {
  slide.addShape(prs.ShapeType.rect, { x: phaseX[i], y: 2.1, w: phaseW, h: 4, fill: { color: C.hero }, border: { pt: 1, color: C.blue } });
  slide.addText(phase.title, { x: phaseX[i] + 0.15, y: 2.3, w: phaseW - 0.3, h: 0.5, fontSize: 11, bold: true, color: C.blue });
  slide.addText(phase.desc, { x: phaseX[i] + 0.15, y: 2.95, w: phaseW - 0.3, h: 3.1, fontSize: 9, color: C.white });
});

// Ask box
slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 6.3, w: 9, h: 0.9, fill: { color: C.hero }, border: { pt: 2, color: C.blue } });
slide.addText('What I Need: Lender partnership team (3–4 people for relationship management) · ML engineering for credit scoring · Alignment with Risk & Compliance, Issuers, Lenders', { x: 0.7, y: 6.4, w: 8.6, h: 0.7, fontSize: 11, bold: true, color: C.blue });

// Footer
slide.addText('Ajay Avaghade | avaghadeajay009@gmail.com | +91 9561558439', { x: 0.5, y: 7.3, w: 9, h: 0.2, fontSize: 9, color: C.muted, align: 'center' });

prs.writeFile({ fileName: 'assets/mastercard-creditbridge-deck.pptx' }).then(() => {
  console.log('✓ mastercard-creditbridge-deck.pptx generated');
});
