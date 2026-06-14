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
      s.addText('Upsell Engine', { x: 0.5, y: 2.5, w: 8, h: 1.2, fontSize: 44, bold: true, color: C.white });
      s.addText('Contextual post-payment upsells via ML propensity scoring.', { x: 0.5, y: 3.7, w: 7, h: 0.6, fontSize: 18, color: C.lgray });
      s.addText('Ajay Avaghade · Product Manager', { x: 0.5, y: 6.5, w: 5, h: 0.5, fontSize: 14, color: C.lgray });
      
      s.addShape(pres.ShapeType.rect, { x: 7, y: 4.5, w: 2.5, h: 2, fill: { color: C.bright } });
      s.addText('18%', { x: 7.2, y: 4.8, w: 2, h: 1, fontSize: 48, bold: true, color: C.white });
      s.addText('Increase in Avg Order Value', { x: 7.2, y: 5.7, w: 2.2, h: 0.5, fontSize: 12, color: C.white });
    }
  },
  {
    title: "The Problem",
    bg: C.dark,
    fn: (s) => {
      s.addText('THE PROBLEM', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.lgray, letterSpacing: 1 });
      s.addText('Restaurants lose the guest relationship at payment.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white });
      
      const cols = [
        { x: 0.5, stat: '80%', lbl: 'Guests leave', desc: 'immediately after paying the bill' },
        { x: 3.8, stat: '0', lbl: 'Contextual offers', desc: 'during the highest intent moment (payment)' },
        { x: 7.1, stat: 'Missed', lbl: 'Dessert & drinks', desc: 'revenue from impulse purchases' }
      ];
      cols.forEach(c => {
        s.addText(c.stat, { x: c.x, y: 2.5, w: 2.5, h: 1, fontSize: 40, bold: true, color: C.bright });
        s.addText(c.lbl, { x: c.x, y: 3.5, w: 2.8, h: 0.4, fontSize: 14, bold: true, color: C.white });
        s.addText(c.desc, { x: c.x, y: 3.9, w: 2.8, h: 0.6, fontSize: 13, color: C.lgray });
      });
      
      s.addShape(pres.ShapeType.rect, { x: 0.5, y: 5.5, w: 9, h: 1.2, fill: { color: C.hero } });
      s.addText('Insight: The digital payment screen is an under-utilized billboard.', { x: 0.8, y: 5.8, w: 8.4, h: 0.6, fontSize: 16, bold: true, color: C.white });
    }
  },
  {
    title: "The Insight",
    bg: C.white,
    fn: (s) => {
      s.addText('THE INSIGHT', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.dark, letterSpacing: 1 });
      s.addText('Turn payment into a point of sale.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.dark });
      
      s.addShape(pres.ShapeType.rect, { x: 0.5, y: 2.5, w: 4.2, h: 2.5, fill: { color: C.lgray } });
      s.addText('CURRENT: Just Payment', { x: 0.8, y: 2.8, w: 3.6, h: 0.4, fontSize: 14, bold: true, color: C.dark });
      s.addText('❌ Scan QR\n❌ Pay exact bill\n❌ Leave', { x: 0.8, y: 3.4, w: 3.6, h: 1.2, fontSize: 13, color: C.dark });
      
      s.addShape(pres.ShapeType.ellipse, { x: 4.6, y: 3.4, w: 0.8, h: 0.8, fill: { color: C.dark } });
      s.addText('VS', { x: 4.6, y: 3.4, w: 0.8, h: 0.8, fontSize: 12, bold: true, color: C.white, align: 'center' });
      
      s.addShape(pres.ShapeType.rect, { x: 5.3, y: 2.5, w: 4.2, h: 2.5, fill: { color: C.bright }, shadow: makeShadow() });
      s.addText('PROPOSED: Upsell Engine', { x: 5.6, y: 2.8, w: 3.6, h: 0.4, fontSize: 14, bold: true, color: C.white });
      s.addText('✅ Scan QR\n✅ See AI-targeted dessert offer\n✅ Add to bill + Pay in 1 tap', { x: 5.6, y: 3.4, w: 3.6, h: 1.2, fontSize: 13, color: C.white });
    }
  },
  {
    title: "The Mechanic",
    bg: C.dark,
    fn: (s) => {
      s.addText('THE MECHANIC', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.lgray, letterSpacing: 1 });
      s.addText('How ML Upsells work.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white });
      
      const steps = ['Read POS Cart', 'Run ML Model', 'Surface Offer', '1-Tap Add', 'Kitchen Notified'];
      steps.forEach((step, i) => {
        const x = 0.5 + (i * 1.8);
        s.addShape(pres.ShapeType.ellipse, { x: x, y: 3, w: 0.6, h: 0.6, fill: { color: C.bright } });
        s.addText(String(i+1), { x: x, y: 3, w: 0.6, h: 0.6, fontSize: 14, bold: true, color: C.white, align: 'center' });
        s.addText(step, { x: x - 0.2, y: 3.8, w: 1, h: 0.5, fontSize: 12, bold: true, color: C.white, align: 'center' });
      });
      s.addShape(pres.ShapeType.line, { x: 1, y: 3.3, w: 7, h: 0, line: { color: C.lgray, dashType: 'dash' } });
      
      s.addText('* Tested successfully at PhonePe: dynamic cart incentivization yields 35% AOV uplift.', { x: 0.5, y: 6, w: 9, h: 0.5, fontSize: 11, color: C.lgray });
    }
  },
  {
    title: "The Product",
    bg: C.white,
    fn: (s) => {
      s.addText('THE PRODUCT', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.dark, letterSpacing: 1 });
      s.addText('Seamless upsell experience.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.dark });
      
      const screens = [
        { x: 0.5, num: '01', title: 'Checkout Upsell', desc: 'Targeted offer on payment screen' },
        { x: 2.8, num: '02', title: 'Offer Detail', desc: 'Visual appeal & new total' },
        { x: 5.1, num: '03', title: 'Success', desc: 'Order confirmed & paid' },
        { x: 7.4, num: '04', title: 'Ops Dashboard', desc: 'Conversion metrics for restaurant' }
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
      s.addText('Pure margin expansion for restaurants.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white });
      
      s.addText('For Restaurants', { x: 0.5, y: 2.2, w: 4, h: 0.4, fontSize: 16, bold: true, color: C.bright });
      s.addText('For Qlub', { x: 5.5, y: 2.2, w: 4, h: 0.4, fontSize: 16, bold: true, color: C.bright });
      
      s.addText('+18%\nAverage Order Value', { x: 0.5, y: 3, w: 4, h: 0.8, fontSize: 18, bold: true, color: C.white });
      s.addText('Higher GMV\nprocessed per table', { x: 5.5, y: 3, w: 4, h: 0.8, fontSize: 18, bold: true, color: C.white });
      
      s.addText('High Margin\nDessert/drinks focus', { x: 0.5, y: 4.2, w: 4, h: 0.8, fontSize: 18, bold: true, color: C.white });
      s.addText('New Revenue\nUpsell commission model', { x: 5.5, y: 4.2, w: 4, h: 0.8, fontSize: 18, bold: true, color: C.white });
      
      s.addShape(pres.ShapeType.rect, { x: 0.5, y: 6, w: 9, h: 1, fill: { color: C.hero } });
      s.addText('ROI: Qlub evolves from a cost center (payment) to a revenue center.', { x: 0.8, y: 6.2, w: 8.4, h: 0.6, fontSize: 14, color: C.lgray });
    }
  },
  {
    title: "Proof of Work",
    bg: C.white,
    fn: (s) => {
      s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 5, h: 7.5, fill: { color: C.dark } });
      s.addText('PROOF OF WORK', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.bright, letterSpacing: 1 });
      s.addText('PhonePe Offers', { x: 0.5, y: 1.5, w: 4, h: 0.6, fontSize: 24, bold: true, color: C.white });
      s.addText('Led Propensity-to-Transact ML models for cart incentivization on Pincode, replacing static offers with dynamic intent triggers.', { x: 0.5, y: 2.5, w: 4, h: 1.5, fontSize: 14, color: C.lgray });
      s.addText('→ 35% AOV Uplift\n→ 60% Drop in Cart Abandonment', { x: 0.5, y: 4.5, w: 4, h: 1, fontSize: 16, bold: true, color: C.bright });
      
      s.addText('Why this matters for Qlub', { x: 5.5, y: 1.5, w: 4, h: 0.6, fontSize: 24, bold: true, color: C.dark });
      s.addText('1. I know how to build contextual offer engines.\n2. I have hands-on experience with ML targeting.\n3. I understand merchant economics & margins.', { x: 5.5, y: 2.5, w: 4, h: 2, fontSize: 14, color: C.dark });
    }
  },
  {
    title: "Rollout Plan",
    bg: C.dark,
    fn: (s) => {
      s.addText('ROLLOUT PLAN', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.lgray, letterSpacing: 1 });
      s.addText('From prototype to launch.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white });
      
      const p = [
        { x: 0.5, title: 'Phase 1: Rule-Based', desc: 'Simple mapping: "If Main Course ordered -> Show Dessert".' },
        { x: 3.6, title: 'Phase 2: ML Engine', desc: 'Collaborative filtering based on what similar tables order.' },
        { x: 6.7, title: 'Phase 3: Self-Serve', desc: 'Restaurants set their own inventory-based upsells.' }
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

pres.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/qlub-upsell-deck.pptx' });
