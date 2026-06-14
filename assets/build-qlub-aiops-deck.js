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
      s.addText('AI-Ops Hub', { x: 0.5, y: 2.5, w: 8, h: 1.2, fontSize: 44, bold: true, color: C.white });
      s.addText('Real-time table turnover prediction & accounting reconciliation.', { x: 0.5, y: 3.7, w: 7, h: 0.6, fontSize: 18, color: C.lgray });
      s.addText('Ajay Avaghade · Product Manager', { x: 0.5, y: 6.5, w: 5, h: 0.5, fontSize: 14, color: C.lgray });
      
      s.addShape(pres.ShapeType.rect, { x: 7, y: 4.5, w: 2.5, h: 2, fill: { color: C.bright } });
      s.addText('100%', { x: 7.2, y: 4.8, w: 2, h: 1, fontSize: 48, bold: true, color: C.white });
      s.addText('Automated POS Sync & Alerts', { x: 7.2, y: 5.7, w: 2.2, h: 0.5, fontSize: 12, color: C.white });
    }
  },
  {
    title: "The Problem",
    bg: C.dark,
    fn: (s) => {
      s.addText('THE PROBLEM', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.lgray, letterSpacing: 1 });
      s.addText('Payment data sits in silos, unused for operations.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white });
      
      const cols = [
        { x: 0.5, stat: 'Blind', lbl: 'Waitstaff', desc: 'don\'t know when a table is ready to pay' },
        { x: 3.8, stat: '2 Hrs', lbl: 'Manual Sync', desc: 'spent daily reconciling POS and payment gateways' },
        { x: 7.1, stat: 'Lost', lbl: 'Table Turns', desc: 'due to operational latency' }
      ];
      cols.forEach(c => {
        s.addText(c.stat, { x: c.x, y: 2.5, w: 2.5, h: 1, fontSize: 40, bold: true, color: C.bright });
        s.addText(c.lbl, { x: c.x, y: 3.5, w: 2.8, h: 0.4, fontSize: 14, bold: true, color: C.white });
        s.addText(c.desc, { x: c.x, y: 3.9, w: 2.8, h: 0.6, fontSize: 13, color: C.lgray });
      });
      
      s.addShape(pres.ShapeType.rect, { x: 0.5, y: 5.5, w: 9, h: 1.2, fill: { color: C.hero } });
      s.addText('Insight: Qlub has the ultimate operational trigger — the payment state.', { x: 0.8, y: 5.8, w: 8.4, h: 0.6, fontSize: 16, bold: true, color: C.white });
    }
  },
  {
    title: "The Insight",
    bg: C.white,
    fn: (s) => {
      s.addText('THE INSIGHT', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.dark, letterSpacing: 1 });
      s.addText('From static analytics to Agentic Ops.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.dark });
      
      s.addShape(pres.ShapeType.rect, { x: 0.5, y: 2.5, w: 4.2, h: 2.5, fill: { color: C.lgray } });
      s.addText('CURRENT: Static Dashboards', { x: 0.8, y: 2.8, w: 3.6, h: 0.4, fontSize: 14, bold: true, color: C.dark });
      s.addText('❌ Managers pull reports end of day\n❌ Waitstaff check tables visually\n❌ Manual ERP sync', { x: 0.8, y: 3.4, w: 3.6, h: 1.2, fontSize: 13, color: C.dark });
      
      s.addShape(pres.ShapeType.ellipse, { x: 4.6, y: 3.4, w: 0.8, h: 0.8, fill: { color: C.dark } });
      s.addText('VS', { x: 4.6, y: 3.4, w: 0.8, h: 0.8, fontSize: 12, bold: true, color: C.white, align: 'center' });
      
      s.addShape(pres.ShapeType.rect, { x: 5.3, y: 2.5, w: 4.2, h: 2.5, fill: { color: C.bright }, shadow: makeShadow() });
      s.addText('PROPOSED: AI-Ops Hub', { x: 5.6, y: 2.8, w: 3.6, h: 0.4, fontSize: 14, bold: true, color: C.white });
      s.addText('✅ AI predicts payment intent\n✅ Nudges waiters proactively\n✅ Auto-reconciles POS nightly', { x: 5.6, y: 3.4, w: 3.6, h: 1.2, fontSize: 13, color: C.white });
    }
  },
  {
    title: "The Mechanic",
    bg: C.dark,
    fn: (s) => {
      s.addText('THE MECHANIC', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.lgray, letterSpacing: 1 });
      s.addText('How Predictive Ops work.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white });
      
      const steps = ['Monitor POS', 'Run Turnover ML', 'Alert Waiter', 'Auto-Drop Bill', 'Nightly Sync'];
      steps.forEach((step, i) => {
        const x = 0.5 + (i * 1.8);
        s.addShape(pres.ShapeType.ellipse, { x: x, y: 3, w: 0.6, h: 0.6, fill: { color: C.bright } });
        s.addText(String(i+1), { x: x, y: 3, w: 0.6, h: 0.6, fontSize: 14, bold: true, color: C.white, align: 'center' });
        s.addText(step, { x: x - 0.2, y: 3.8, w: 1, h: 0.5, fontSize: 12, bold: true, color: C.white, align: 'center' });
      });
      s.addShape(pres.ShapeType.line, { x: 1, y: 3.3, w: 7, h: 0, line: { color: C.lgray, dashType: 'dash' } });
      
      s.addText('* Tested successfully at PhonePe: B2B dashboards with actionable insights drive 500% higher engagement.', { x: 0.5, y: 6, w: 9, h: 0.5, fontSize: 11, color: C.lgray });
    }
  },
  {
    title: "The Product",
    bg: C.white,
    fn: (s) => {
      s.addText('THE PRODUCT', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.dark, letterSpacing: 1 });
      s.addText('The merchant control plane.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.dark });
      
      const screens = [
        { x: 0.5, num: '01', title: 'AI Dashboard', desc: 'Live floor & urgent alerts' },
        { x: 2.8, num: '02', title: 'Table Detail', desc: 'AI prediction score & POS data' },
        { x: 5.1, num: '03', title: 'Auto-Sync', desc: 'Zero discrepancy accounting' },
        { x: 7.4, num: '04', title: 'Automation Settings', desc: 'Configurable agent rules' }
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
      s.addText('Unmatched B2B stickiness.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white });
      
      s.addText('For Restaurants', { x: 0.5, y: 2.2, w: 4, h: 0.4, fontSize: 16, bold: true, color: C.bright });
      s.addText('For Qlub', { x: 5.5, y: 2.2, w: 4, h: 0.4, fontSize: 16, bold: true, color: C.bright });
      
      s.addText('2 Hrs Saved\nDaily in accounting', { x: 0.5, y: 3, w: 4, h: 0.8, fontSize: 18, bold: true, color: C.white });
      s.addText('Lower Churn\nIndispensable tool', { x: 5.5, y: 3, w: 4, h: 0.8, fontSize: 18, bold: true, color: C.white });
      
      s.addText('Faster Turns\nVia proactive nudges', { x: 0.5, y: 4.2, w: 4, h: 0.8, fontSize: 18, bold: true, color: C.white });
      s.addText('Moat Creation\nBeyond just payments', { x: 5.5, y: 4.2, w: 4, h: 0.8, fontSize: 18, bold: true, color: C.white });
      
      s.addShape(pres.ShapeType.rect, { x: 0.5, y: 6, w: 9, h: 1, fill: { color: C.hero } });
      s.addText('ROI: Qlub becomes the operating system for the restaurant floor.', { x: 0.8, y: 6.2, w: 8.4, h: 0.6, fontSize: 14, color: C.lgray });
    }
  },
  {
    title: "Proof of Work",
    bg: C.white,
    fn: (s) => {
      s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 5, h: 7.5, fill: { color: C.dark } });
      s.addText('PROOF OF WORK', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.bright, letterSpacing: 1 });
      s.addText('PhonePe B2B', { x: 0.5, y: 1.5, w: 4, h: 0.6, fontSize: 24, bold: true, color: C.white });
      s.addText('Built zero-to-one self-serve platform for 5,000+ merchants. Focused heavily on operational tooling and reconciliation APIs.', { x: 0.5, y: 2.5, w: 4, h: 1.5, fontSize: 14, color: C.lgray });
      s.addText('→ 5,000+ Merchants\n→ 30-min Ops TAT', { x: 0.5, y: 4.5, w: 4, h: 1, fontSize: 16, bold: true, color: C.bright });
      
      s.addText('Why this matters for Qlub', { x: 5.5, y: 1.5, w: 4, h: 0.6, fontSize: 24, bold: true, color: C.dark });
      s.addText('1. I know how to build B2B tools that merchants actually use.\n2. I understand complex ledger/reconciliation logic.\n3. I use AI to automate operational drag.', { x: 5.5, y: 2.5, w: 4, h: 2, fontSize: 14, color: C.dark });
    }
  },
  {
    title: "Rollout Plan",
    bg: C.dark,
    fn: (s) => {
      s.addText('ROLLOUT PLAN', { x: 0.5, y: 0.5, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.lgray, letterSpacing: 1 });
      s.addText('From prototype to launch.', { x: 0.5, y: 1, w: 9, h: 0.6, fontSize: 28, bold: true, color: C.white });
      
      const p = [
        { x: 0.5, title: 'Phase 1: Basic Dash', desc: 'Surface live POS table data securely to waitstaff iPads.' },
        { x: 3.6, title: 'Phase 2: Predictive Alerts', desc: 'ML models train on wait times to start sending nudges.' },
        { x: 6.7, title: 'Phase 3: ERP Integration', desc: 'Direct webhook sync for zero-click accounting.' }
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

pres.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/qlub-aiops-deck.pptx' });
