const pptxgen = require('pptxgenjs');

const C = {
  dark:   '0C0C0C',
  hero:   '141414',
  bright: '7D00D4',
  orange: 'F97316',
  white:  'FFFFFF',
  lgray:  'F1F5F9',
};

const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';

const phases = [
  {
    title: "Cover",
    bg: C.dark,
    fn: (s) => {
      s.addText('QLUB', { x: 0.5, y: 0.5, w: 2, h: 0.5, fontSize: 18, bold: true, color: C.white, letterSpacing: 2 });
      s.addText('SmartSplit', { x: 0.5, y: 2.5, w: 8, h: 1.2, fontSize: 44, bold: true, color: C.white });
      s.addText('Zero-friction bill splitting for large groups to increase table turnover.', { x: 0.5, y: 3.7, w: 7, h: 0.6, fontSize: 18, color: C.lgray });
      s.addText('Ajay Avaghade · Product Manager', { x: 0.5, y: 6.5, w: 5, h: 0.5, fontSize: 14, color: C.lgray });
      
      s.addShape(pres.ShapeType.rect, { x: 7, y: 4.5, w: 2.5, h: 2, fill: { color: C.bright } });
      s.addText('40%', { x: 7.2, y: 4.8, w: 2, h: 1, fontSize: 48, bold: true, color: C.white });
      s.addText('Reduction in table blocking time', { x: 7.2, y: 5.7, w: 2.2, h: 0.5, fontSize: 12, color: C.white });
    }
  },
  {
    title: "The Problem",
    bg: C.dark,
    fn: (s) => {
      s.addText('THE PROBLEM', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.lgray, letterSpacing: 1 });
      s.addText('Bill splitting is the highest friction moment in dining.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white });
      
      const cols = [
        { x: 0.5, stat: '12 min', lbl: 'Average time spent', desc: 'calculating and collecting group payments' },
        { x: 3.8, stat: '2.5x', lbl: 'Higher churn risk', desc: 'during peak hours due to waitlist delays' },
        { x: 7.1, stat: '15%', lbl: 'Lost capacity', desc: 'due to table blocking post-meal' }
      ];
      cols.forEach(c => {
        s.addText(c.stat, { x: c.x, y: 2.5, w: 2.5, h: 1, fontSize: 40, bold: true, color: C.bright });
        s.addText(c.lbl, { x: c.x, y: 3.5, w: 2.8, h: 0.4, fontSize: 14, bold: true, color: C.white });
        s.addText(c.desc, { x: c.x, y: 3.9, w: 2.8, h: 0.6, fontSize: 13, color: C.lgray });
      });
      
      s.addShape(pres.ShapeType.rect, { x: 0.5, y: 5.5, w: 9, h: 1.2, fill: { color: C.hero } });
      s.addText('Insight: Restaurants don\'t just lose time; they lose the next table\'s revenue.', { x: 0.8, y: 5.8, w: 8.4, h: 0.6, fontSize: 16, bold: true, color: C.white });
    }
  },
  {
    title: "The Insight",
    bg: C.white,
    fn: (s) => {
      s.addText('THE INSIGHT', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.dark, letterSpacing: 1 });
      s.addText('Move the math from the table to the phone.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.dark });
      
      s.addShape(pres.ShapeType.rect, { x: 0.5, y: 2.5, w: 4.2, h: 2.5, fill: { color: C.lgray } });
      s.addText('CURRENT: Manual Math', { x: 0.8, y: 2.8, w: 3.6, h: 0.4, fontSize: 14, bold: true, color: C.dark });
      s.addText('❌ One person pays, chases others\n❌ Waiter does complex POS splits\n❌ Physical cards passed around', { x: 0.8, y: 3.4, w: 3.6, h: 1.2, fontSize: 13, color: C.dark });
      
      s.addShape(pres.ShapeType.ellipse, { x: 4.6, y: 3.4, w: 0.8, h: 0.8, fill: { color: C.dark } });
      s.addText('VS', { x: 4.6, y: 3.4, w: 0.8, h: 0.8, fontSize: 12, bold: true, color: C.white, align: 'center' });
      
      s.addShape(pres.ShapeType.rect, { x: 5.3, y: 2.5, w: 4.2, h: 2.5, fill: { color: C.bright }, shadow: makeShadow() });
      s.addText('PROPOSED: Qlub SmartSplit', { x: 5.6, y: 2.8, w: 3.6, h: 0.4, fontSize: 14, bold: true, color: C.white });
      s.addText('✅ Select your items on your phone\n✅ Pay via Apple/Google Pay instantly\n✅ Waiter sees live table status', { x: 5.6, y: 3.4, w: 3.6, h: 1.2, fontSize: 13, color: C.white });
    }
  },
  {
    title: "The Mechanic",
    bg: C.dark,
    fn: (s) => {
      s.addText('THE MECHANIC', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.lgray, letterSpacing: 1 });
      s.addText('How SmartSplit works in 5 steps.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white });
      
      const steps = ['Scan QR', 'View Receipt', 'Select Items', '1-Tap Pay', 'Table Cleared'];
      steps.forEach((step, i) => {
        const x = 0.5 + (i * 1.8);
        s.addShape(pres.ShapeType.ellipse, { x: x, y: 3, w: 0.6, h: 0.6, fill: { color: C.bright } });
        s.addText(String(i+1), { x: x, y: 3, w: 0.6, h: 0.6, fontSize: 14, bold: true, color: C.white, align: 'center' });
        s.addText(step, { x: x - 0.2, y: 3.8, w: 1, h: 0.5, fontSize: 12, bold: true, color: C.white, align: 'center' });
      });
      s.addShape(pres.ShapeType.line, { x: 1, y: 3.3, w: 7, h: 0, line: { color: C.lgray, dashType: 'dash' } });
      
      s.addText('* Tested successfully at PhonePe: reducing friction in multi-party flows yields highest conversion lifts.', { x: 0.5, y: 6, w: 9, h: 0.5, fontSize: 11, color: C.lgray });
    }
  },
  {
    title: "The Product",
    bg: C.white,
    fn: (s) => {
      s.addText('THE PRODUCT', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.dark, letterSpacing: 1 });
      s.addText('End-to-end split payment flow.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.dark });
      
      const screens = [
        { x: 0.5, num: '01', title: 'Digital Receipt', desc: 'Scan to view live POS data' },
        { x: 2.8, num: '02', title: 'Split Mode', desc: 'Select items or split evenly' },
        { x: 5.1, num: '03', title: 'Success Status', desc: 'See who else paid at table' },
        { x: 7.4, num: '04', title: 'Ops Dashboard', desc: 'Waiters track live progress' }
      ];
      
      screens.forEach(sc => {
        s.addShape(pres.ShapeType.rect, { x: sc.x, y: 2.5, w: 2.1, h: 4, fill: { color: C.lgray } });
        s.addText(sc.num, { x: sc.x + 0.2, y: 2.7, w: 1.7, h: 0.3, fontSize: 10, bold: true, color: C.bright });
        s.addText(sc.title, { x: sc.x + 0.2, y: 3.1, w: 1.7, h: 0.4, fontSize: 14, bold: true, color: C.dark });
        s.addText(sc.desc, { x: sc.x + 0.2, y: 3.5, w: 1.7, h: 0.8, fontSize: 11, color: C.dark });
        s.addText('[ UI MOCKUP ]', { x: sc.x, y: 4.5, w: 2.1, h: 1, fontSize: 12, color: '#94a3b8', align: 'center' });
      });
    }
  },
  {
    title: "Impact & ROI",
    bg: C.dark,
    fn: (s) => {
      s.addText('IMPACT & ROI', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.lgray, letterSpacing: 1 });
      s.addText('Value created for both sides of the marketplace.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white });
      
      s.addText('For Restaurants', { x: 0.5, y: 2.2, w: 4, h: 0.4, fontSize: 16, bold: true, color: C.bright });
      s.addText('For Diners', { x: 5.5, y: 2.2, w: 4, h: 0.4, fontSize: 16, bold: true, color: C.bright });
      
      s.addText('15%\nCapacity Increase', { x: 0.5, y: 3, w: 4, h: 0.8, fontSize: 18, bold: true, color: C.white });
      s.addText('0\nAwkward Math', { x: 5.5, y: 3, w: 4, h: 0.8, fontSize: 18, bold: true, color: C.white });
      
      s.addText('100%\nPOS Reconciliation', { x: 0.5, y: 4.2, w: 4, h: 0.8, fontSize: 18, bold: true, color: C.white });
      s.addText('< 30s\nTime to Leave', { x: 5.5, y: 4.2, w: 4, h: 0.8, fontSize: 18, bold: true, color: C.white });
      
      s.addShape(pres.ShapeType.rect, { x: 0.5, y: 6, w: 9, h: 1, fill: { color: C.hero } });
      s.addText('ROI: Faster table turns = higher revenue per square foot.', { x: 0.8, y: 6.2, w: 8.4, h: 0.6, fontSize: 14, color: C.lgray });
    }
  },
  {
    title: "Proof of Work",
    bg: C.white,
    fn: (s) => {
      s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 5, h: 7.5, fill: { color: C.dark } });
      s.addText('PROOF OF WORK', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.bright, letterSpacing: 1 });
      s.addText('PhonePe Checkout', { x: 0.5, y: 1.5, w: 4, h: 0.6, fontSize: 24, bold: true, color: C.white });
      s.addText('Redesigned the end-to-end payment and offers flow across 350M MAU, removing friction and automating logic.', { x: 0.5, y: 2.5, w: 4, h: 1.5, fontSize: 14, color: C.lgray });
      s.addText('→ 22% Conversion Lift\n→ 32% Burn Reduction', { x: 0.5, y: 4.5, w: 4, h: 1, fontSize: 16, bold: true, color: C.bright });
      
      s.addText('Why this matters for Qlub', { x: 5.5, y: 1.5, w: 4, h: 0.6, fontSize: 24, bold: true, color: C.dark });
      s.addText('1. I know how to build zero-friction payment UI at scale.\n2. I understand complex ledger/split math.\n3. I focus obsessively on merchant ROI.', { x: 5.5, y: 2.5, w: 4, h: 2, fontSize: 14, color: C.dark });
    }
  },
  {
    title: "Rollout Plan",
    bg: C.dark,
    fn: (s) => {
      s.addText('ROLLOUT PLAN', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.lgray, letterSpacing: 1 });
      s.addText('From prototype to launch.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white });
      
      const p = [
        { x: 0.5, title: 'Phase 1: POS Integration', desc: 'Map item-level data from major POS systems into Qlub ledger.' },
        { x: 3.6, title: 'Phase 2: UI Beta', desc: 'Deploy split UI to 5 high-volume restaurants. Measure table turn time.' },
        { x: 6.7, title: 'Phase 3: Scale', desc: 'Roll out globally. Introduce gamified "Pay for a friend" features.' }
      ];
      
      p.forEach(ph => {
        s.addShape(pres.ShapeType.rect, { x: ph.x, y: 2.5, w: 2.8, h: 1.8, fill: { color: C.hero } });
        s.addText(ph.title, { x: ph.x + 0.2, y: 2.8, w: 2.4, h: 0.4, fontSize: 14, bold: true, color: C.bright });
        s.addText(ph.desc, { x: ph.x + 0.2, y: 3.3, w: 2.4, h: 0.8, fontSize: 12, color: C.lgray });
      });
      
      s.addShape(pres.ShapeType.rect, { x: 0.5, y: 5, w: 9, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
      s.addText('Let\'s build this together.', { x: 0.8, y: 5.3, w: 8.4, h: 0.5, fontSize: 18, bold: true, color: C.dark });
      s.addText('ajay-avaghade.github.io/Portfolio · +91 9561558439', { x: 0.8, y: 5.9, w: 8.4, h: 0.4, fontSize: 13, color: '#64748b' });
    }
  }
];

phases.forEach(p => {
  let slide = pres.addSlide();
  slide.background = { color: p.bg };
  p.fn(slide);
});

pres.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/qlub-smartsplit-deck.pptx' });
