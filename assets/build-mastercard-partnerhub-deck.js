const PptxGenJS = require('pptxgenjs');

const prs = new PptxGenJS();
prs.defineLayout({ name: 'LAYOUT_16x9', width: 10, height: 7.5 });
prs.layout = 'LAYOUT_16x9';

const C = {
  dark: '0A1F2E',
  hero: '152A3D',
  teal: '00D9A3',
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
slide.addText('MASTERCARD · SME ECOSYSTEM PLAY', { x: 0.5, y: 0.25, w: 9, h: 0.7, fontSize: 9, bold: true, color: C.teal, align: 'center', charSpacing: 2 });
slide.addText('PartnerHub', { x: 1, y: 2, w: 8, h: 1.2, fontSize: 52, bold: true, color: C.white, align: 'center', fontFace: 'Sora' });
slide.addText('SME Fintech Partnership Marketplace', { x: 1, y: 3.3, w: 8, h: 0.5, fontSize: 24, color: C.teal, align: 'center' });
slide.addText('Ajay Avaghade · Product Manager', { x: 1, y: 4, w: 8, h: 0.4, fontSize: 13, color: C.muted, align: 'center', italic: true });
slide.addShape(prs.ShapeType.rect, { x: 6.5, y: 5.8, w: 3, h: 1.2, fill: { color: C.hero }, shadow: makeShadow() });
slide.addText('Partner CAC\n↓ 50%', { x: 6.5, y: 5.85, w: 3, h: 0.5, fontSize: 28, bold: true, color: C.teal, align: 'center' });
slide.addText('Fintech distribution unlock', { x: 6.5, y: 6.35, w: 3, h: 0.4, fontSize: 11, color: C.muted, align: 'center' });

// SLIDE 2: THE PROBLEM
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.hero } });
slide.addText('THE PROBLEM', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.teal, align: 'left', charSpacing: 2 });
slide.addText("70% SMEs Use No Tools; Fintech Partners Can't Reach Them", { x: 0.5, y: 1.1, w: 9, h: 0.8, fontSize: 30, bold: true, color: C.white });

// Stat columns
const statX = [0.5, 3.5, 6.5];
const statLabels = ['📊\n\n70%', '💼\n\n40-60%', '💰\n\n0'];
const statDescs = ['SMEs lack accounting,\nPOS, lending tools', 'Fintech partner CAC\nfrom direct acquisition', 'Mastercard revenue\nfrom partnerships'];

statX.forEach((x, i) => {
  slide.addShape(prs.ShapeType.rect, { x, y: 2.2, w: 2.8, h: 2, fill: { color: C.hero }, border: { pt: 1, color: C.teal } });
  slide.addText(statLabels[i], { x: x + 0.2, y: 2.4, w: 2.4, h: 1, fontSize: 20, bold: true, color: C.teal, align: 'center' });
  slide.addText(statDescs[i], { x: x + 0.2, y: 3.6, w: 2.4, h: 0.5, fontSize: 10, color: C.muted, align: 'center' });
});

slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 4.5, w: 9, h: 1.8, fill: { color: C.hero }, border: { pt: 1, color: C.gold } });
slide.addText("The Gap: SMEs don't know tools exist. Fintech partners can't efficiently reach SMEs. Mastercard sits in the middle with 2M+ SME cardholders but has no ecosystem leverage.", { x: 0.8, y: 4.7, w: 8.4, h: 1.4, fontSize: 12, color: C.white, align: 'left' });

// SLIDE 3: THE INSIGHT
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.dark } });
slide.addText('THE INSIGHT', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.teal, align: 'left', charSpacing: 2 });
slide.addText('Mastercard as SME Fintech Distribution Layer', { x: 0.5, y: 1.1, w: 9, h: 0.8, fontSize: 28, bold: true, color: C.dark });

// Left side
slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2.1, w: 4.2, h: 4.8, fill: { color: C.white }, border: { pt: 2, color: C.dark } });
slide.addText('❌ Today', { x: 0.7, y: 2.3, w: 3.8, h: 0.4, fontSize: 14, bold: true, color: C.dark });
const statusQuo = ["SMEs don't know\nwhich tools exist", "Fintech partners\nhave no SME channel", "High CAC prevents\nfintechs from reaching SMEs", "Mastercard misses\npartnership revenue"];
let yPos = 2.8;
statusQuo.forEach(item => {
  slide.addText(item, { x: 0.7, y: yPos, w: 3.8, h: 0.6, fontSize: 11, color: C.dark });
  yPos += 0.7;
});

// Right side
slide.addShape(prs.ShapeType.rect, { x: 5.3, y: 2.1, w: 4.2, h: 4.8, fill: { color: C.hero }, border: { pt: 2, color: C.teal } });
slide.addText('✅ PartnerHub', { x: 5.5, y: 2.3, w: 3.8, h: 0.4, fontSize: 14, bold: true, color: C.white });
const proposed = ['SME discovers partners\nby use case', 'Mastercard becomes\nfintech distribution channel', 'Partner CAC drops 50%\nvia Mastercard pipeline', 'Mastercard earns\npartnership fees'];
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
slide.addText('Browse → Request Integration → Mastercard Handoff → Usage Tracked', { x: 0.5, y: 1.1, w: 9, h: 0.8, fontSize: 26, bold: true, color: C.white });

// Timeline steps
const steps = [
  { num: '1', title: 'SME Discovers', desc: 'Browse partners by category (Accounting, POS, Lending, etc.)' },
  { num: '2', title: 'Partner Comparison', desc: 'Rate, reviews, pricing, integrations side-by-side' },
  { num: '3', title: 'Integration Request', desc: 'SME clicks to request integration; 1-click handoff' },
  { num: '4', title: 'Mastercard Enables', desc: 'SSO/API integration managed by Mastercard, not partner' },
  { num: '5', title: 'Track & Monetize', desc: 'Mastercard tracks usage, earns partnership fee' }
];

const stepX = [0.5, 2.3, 4.1, 5.9, 7.7];
const stepW = 1.7;
steps.forEach((step, i) => {
  slide.addShape(prs.ShapeType.rect, { x: stepX[i], y: 2.2, w: stepW, h: 4.2, fill: { color: C.hero }, border: { pt: 1, color: C.teal } });
  slide.addText(step.num, { x: stepX[i], y: 2.3, w: stepW, h: 0.35, fontSize: 20, bold: true, color: C.teal, align: 'center' });
  slide.addText(step.title, { x: stepX[i] + 0.1, y: 2.8, w: stepW - 0.2, h: 0.6, fontSize: 11, bold: true, color: C.white, align: 'center' });
  slide.addText(step.desc, { x: stepX[i] + 0.1, y: 3.5, w: stepW - 0.2, h: 2, fontSize: 9, color: C.muted, align: 'center' });
});

slide.addText('PhonePe proof: 500+ brand partner marketplace (11% YoY growth) using same adoption curve', { x: 0.5, y: 6.8, w: 9, h: 0.5, fontSize: 11, color: C.gold, italic: true });

// SLIDE 5: THE PRODUCT
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.dark } });
slide.addText('THE PRODUCT', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.teal, align: 'left', charSpacing: 2 });
slide.addText('4 Screen States · SME to Partner Lifecycle', { x: 0.5, y: 1.1, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.dark });

// 4 product cards
const screens = [
  { num: '01', title: 'SME Home', desc: 'Personalized recommendations based on business profile & spend' },
  { num: '02', title: 'Partner Browse', desc: 'Filter by category, sort by rating/price, compare integrations' },
  { num: '03', title: 'Partner Detail', desc: 'Full description, pricing, integration list, 1-click request' },
  { num: '04', title: 'Partner Ops', desc: 'Partner dashboard: onboarded SMEs, integration success, revenue' }
];

const cardX = [0.5, 2.7, 4.9, 7.1];
const cardW = 2.1;
screens.forEach((s, i) => {
  slide.addShape(prs.ShapeType.rect, { x: cardX[i], y: 2.1, w: cardW, h: 4.8, fill: { color: C.white }, border: { pt: 2, color: C.teal } });
  slide.addText(s.num, { x: cardX[i], y: 2.3, w: cardW, h: 0.35, fontSize: 16, bold: true, color: C.teal, align: 'center' });
  slide.addText(s.title, { x: cardX[i] + 0.15, y: 2.8, w: cardW - 0.3, h: 0.5, fontSize: 12, bold: true, color: C.dark, align: 'center' });
  slide.addText(s.desc, { x: cardX[i] + 0.15, y: 3.5, w: cardW - 0.3, h: 3.2, fontSize: 9, color: C.dark, align: 'center' });
});

slide.addText('Live prototype: mastercard-partnerhub-prototype.html', { x: 0.5, y: 7, w: 9, h: 0.35, fontSize: 10, color: C.muted, italic: true });

// SLIDE 6: IMPACT & ROI
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.hero } });
slide.addText('IMPACT & ROI', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.gold, align: 'left', charSpacing: 2 });
slide.addText('SME Adoption + Partner Economics + Mastercard Revenue', { x: 0.5, y: 1.1, w: 9, h: 0.6, fontSize: 22, bold: true, color: C.white });

// SME Adoption
slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2, w: 4.3, h: 4.6, fill: { color: C.hero }, border: { pt: 1, color: C.teal } });
slide.addText('SME Adoption', { x: 0.7, y: 2.15, w: 3.9, h: 0.35, fontSize: 11, bold: true, color: C.teal });
const adoption = ['↑ 12% → 45%\nPartner adoption rate', '↑ 150+ partners\nin marketplace', '↑ 5,000+ integrations\nper month'];
let adoptY = 2.65;
adoption.forEach(item => {
  slide.addText(item, { x: 0.7, y: adoptY, w: 3.9, h: 1, fontSize: 11, bold: true, color: C.white });
  adoptY += 1.3;
});

// Partner Economics
slide.addShape(prs.ShapeType.rect, { x: 5.2, y: 2, w: 4.3, h: 4.6, fill: { color: C.hero }, border: { pt: 1, color: C.gold } });
slide.addText('Partner Economics', { x: 5.4, y: 2.15, w: 3.9, h: 0.35, fontSize: 11, bold: true, color: C.gold });
const partner = ['↓ 50% fintech CAC\nvia Mastercard channel', '↑ 3–5x user acquisition\nspeed', '↑ 80K+ SME customers\nfirst year'];
let partnerY = 2.65;
partner.forEach(item => {
  slide.addText(item, { x: 5.4, y: partnerY, w: 3.9, h: 1, fontSize: 11, bold: true, color: C.white });
  partnerY += 1.3;
});

// Revenue note
slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 7, w: 9, h: 0.4, fill: { color: C.hero }, border: { pt: 1, color: C.gold } });
slide.addText('Mastercard Licensing Revenue: AED 2.4M annually (150 partners × 500 SMEs × AED 32 annual fee)', { x: 0.7, y: 7.05, w: 8.6, h: 0.3, fontSize: 10, bold: true, color: C.gold });

// SLIDE 7: PROOF OF WORK
slide = prs.addSlide();
slide.background = { color: C.lgray };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.dark } });
slide.addText('PROOF OF WORK', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.teal, align: 'left', charSpacing: 2 });
slide.addText('PhonePe Brand Partner Marketplace — Direct Parallel', { x: 0.5, y: 1.1, w: 9, h: 0.6, fontSize: 22, bold: true, color: C.dark });

// Left: PhonePe
slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 2, w: 4.5, h: 4.8, fill: { color: C.dark }, border: { pt: 1, color: C.teal } });
slide.addText('PhonePe Rewards Marketplace', { x: 0.7, y: 2.15, w: 4.1, h: 0.35, fontSize: 12, bold: true, color: C.teal });
const phonepeItems = [
  '500+ brand partners',
  '11% YoY revenue growth',
  '26% user engagement lift',
  'Self-serve partner tooling'
];
let phoneY = 2.65;
phonepeItems.forEach(item => {
  slide.addText(item, { x: 0.7, y: phoneY, w: 4.1, h: 0.8, fontSize: 10, color: C.white });
  phoneY += 1;
});

// Right: PartnerHub
slide.addShape(prs.ShapeType.rect, { x: 5.5, y: 2, w: 4.3, h: 4.8, fill: { color: C.teal }, border: { pt: 1, color: C.teal } });
slide.addText('→ PartnerHub Application', { x: 5.7, y: 2.15, w: 3.9, h: 0.35, fontSize: 12, bold: true, color: C.dark });
const partnerhubItems = [
  '150+ fintech partners',
  'AED 2.4M annual revenue',
  '45% SME adoption rate',
  'Mastercard-enabled integrations'
];
let partnerHubY = 2.65;
partnerhubItems.forEach((item, i) => {
  slide.addText(item, { x: 5.7, y: partnerHubY, w: 3.9, h: 0.8, fontSize: 10, color: C.dark });
  partnerHubY += 1;
});

// SLIDE 8: ROLLOUT PLAN
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: C.hero } });
slide.addText('ROLLOUT PLAN', { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 10, bold: true, color: C.gold, align: 'left', charSpacing: 2 });
slide.addText('Three Phases — Vetting to Scale', { x: 0.5, y: 1.1, w: 9, h: 0.6, fontSize: 22, bold: true, color: C.white });

// Phase cards
const phases = [
  { title: 'Month 1–2: Partner Vetting', desc: 'Identify, vet, onboard 50 fintech/ISV partners. Soft-launch marketplace to 5K SMEs.' },
  { title: 'Month 3–4: Optimization', desc: 'Expand to 150 partners. Measure SME adoption, partner CAC, integration success rates.' },
  { title: 'Month 5–6: EEMEA Scale', desc: 'Launch PartnerHub to all 2M+ Mastercard SME cardholders across EEMEA.' }
];

const phaseX = [0.5, 3.4, 6.3];
const phaseW = 2.8;
phases.forEach((phase, i) => {
  slide.addShape(prs.ShapeType.rect, { x: phaseX[i], y: 2.1, w: phaseW, h: 4, fill: { color: C.hero }, border: { pt: 1, color: C.teal } });
  slide.addText(phase.title, { x: phaseX[i] + 0.15, y: 2.3, w: phaseW - 0.3, h: 0.5, fontSize: 11, bold: true, color: C.teal });
  slide.addText(phase.desc, { x: phaseX[i] + 0.15, y: 2.95, w: phaseW - 0.3, h: 3.1, fontSize: 9, color: C.white });
});

// Ask box
slide.addShape(prs.ShapeType.rect, { x: 0.5, y: 6.3, w: 9, h: 0.9, fill: { color: C.hero }, border: { pt: 2, color: C.teal } });
slide.addText('What I Need: Partner relationship management team (2–3 people) · Technical integration capability · Alignment with Business Development & Partnerships', { x: 0.7, y: 6.4, w: 8.6, h: 0.7, fontSize: 11, bold: true, color: C.teal });

// Footer
slide.addText('Ajay Avaghade | avaghadeajay009@gmail.com | +91 9561558439', { x: 0.5, y: 7.3, w: 9, h: 0.2, fontSize: 9, color: C.muted, align: 'center' });

prs.writeFile({ fileName: 'assets/mastercard-partnerhub-deck.pptx' }).then(() => {
  console.log('✓ mastercard-partnerhub-deck.pptx generated');
});
