(async () => {
const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();

prs.defineLayout({ name: 'LAYOUT_16x9', width: 10, height: 7.5 });
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0D1F14',
  hero:   '0F2A1A',
  bright: '7C3AED',
  accent: '00C853',
  orange: 'FF6B35',
  white:  'FFFFFF',
  lgray:  'F0F4F8',
  muted:  '8b92a5',
};

const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });
const fs = require('fs');

// SLIDE 1: COVER
const s1 = prs.addSlide();
s1.background = { color: C.dark };

s1.addText('REVOLUT · NEW VERTICAL ACTIVATION', {
  x: 0.3, y: 0.4, w: 9.4, h: 0.3,
  fontSize: 9, bold: true, color: C.muted, align: 'center'
});

s1.addText('VerticalShift', {
  x: 0.5, y: 1.2, w: 9, h: 1.5,
  fontSize: 52, bold: true, color: C.white, align: 'center'
});

s1.addText('Milestone-Driven Activation Engine for New Fintech Verticals', {
  x: 0.5, y: 2.8, w: 9, h: 0.8,
  fontSize: 20, color: C.bright, align: 'center'
});

s1.addText('Ajay Avaghade · Product Manager', {
  x: 0.5, y: 3.8, w: 9, h: 0.4,
  fontSize: 12, color: C.muted, align: 'center'
});

s1.addShape('rect', {
  x: 7.5, y: 5.2, w: 2, h: 0.1,
  fill: { color: C.bright }, line: { type: 'none' }
});

s1.addText('+110% Day-30 activation', {
  x: 8.2, y: 5.4, w: 1.8, h: 0.8,
  fontSize: 18, bold: true, color: C.bright, align: 'right'
});

// SLIDE 2: PROBLEM
const s2 = prs.addSlide();
s2.background = { color: C.dark };

s2.addText('THE PROBLEM', {
  x: 0.3, y: 0.3, w: 9.4, h: 0.25,
  fontSize: 9, bold: true, color: C.muted, align: 'left'
});

s2.addText('New Revolut Vertical Day-30 Activation: 18% vs Core Payments 40% — Same Users, Same App', {
  x: 0.5, y: 0.8, w: 9, h: 1.2,
  fontSize: 28, bold: true, color: C.white
});

const statCol = [
  { num: '18%', label: 'new vertical activation', sub: 'Wealth, Lending, Cards (lagging)' },
  { num: '40%', label: 'core payments activation', sub: 'established, habit-formed' },
  { num: '0', label: 'milestone or habit mechanic', sub: 'for new verticals today' }
];

statCol.forEach((stat, i) => {
  const x = 0.5 + i * 3.1;
  s2.addText(stat.num, {
    x: x, y: 2.3, w: 2.8, h: 0.6,
    fontSize: 32, bold: true, color: C.bright, align: 'center'
  });
  s2.addText(stat.label, {
    x: x, y: 3.0, w: 2.8, h: 0.4,
    fontSize: 12, bold: true, color: C.white, align: 'center'
  });
  s2.addText(stat.sub, {
    x: x, y: 3.5, w: 2.8, h: 0.8,
    fontSize: 10, color: C.muted, align: 'center'
  });
});

s2.addShape('rect', {
  x: 0.5, y: 5.2, w: 9, h: 1.8,
  fill: { color: C.hero }, line: { color: C.bright, width: 1 }
});

s2.addText('Root cause: New verticals lack onboarding momentum. Users lack a clear goal or sense of progress. Without milestone + reward mechanic, new vertical feels optional, gets buried in app, user churns by day 30.', {
  x: 0.7, y: 5.4, w: 8.6, h: 1.4,
  fontSize: 11, color: C.white, align: 'left'
});

// SLIDE 3: INSIGHT
const s3 = prs.addSlide();
s3.background = { color: C.lgray };

s3.addText('THE INSIGHT', {
  x: 0.3, y: 0.3, w: 9.4, h: 0.25,
  fontSize: 9, bold: true, color: C.muted, align: 'left'
});

s3.addText('Habit Formation Requires: Goal-Setting + Milestone Visibility + Reward', {
  x: 0.5, y: 0.8, w: 9, h: 1.0,
  fontSize: 28, bold: true, color: C.dark
});

s3.addShape('rect', {
  x: 0.5, y: 2.0, w: 4.2, h: 3.8,
  fill: { color: 'FFFFFF' }, line: { color: C.muted, width: 1 }
});

s3.addText('❌ Today: No Milestone', {
  x: 0.7, y: 2.2, w: 3.8, h: 0.4,
  fontSize: 12, bold: true, color: C.orange
});

const statusQuo = [
  'User onboards to Wealth vertical',
  'No clear next step or goal',
  'No progress visibility',
  'No reward for activation',
  '82% abandon by day 30'
];

statusQuo.forEach((item, i) => {
  s3.addText(item, {
    x: 0.9, y: 2.8 + i * 0.6, w: 3.6, h: 0.55,
    fontSize: 10, color: C.dark, align: 'left'
  });
});

s3.addShape('rect', {
  x: 5.3, y: 2.0, w: 4.2, h: 3.8,
  fill: { color: C.hero }, line: { color: C.bright, width: 1 }
});

s3.addText('✅ VerticalShift: Milestone', {
  x: 5.5, y: 2.2, w: 3.8, h: 0.4,
  fontSize: 12, bold: true, color: C.bright
});

const verticalShift = [
  'Set milestone: "Open investment account"',
  '3 sub-goals with progress bar',
  'Nudges at 33%, 66%, completion',
  'Reward at finish: 500 RevPoints',
  '38% activate by day 30 (+110%)'
];

verticalShift.forEach((item, i) => {
  s3.addText(item, {
    x: 5.7, y: 2.8 + i * 0.6, w: 3.6, h: 0.55,
    fontSize: 10, color: C.white, align: 'left'
  });
});

s3.addShape('ellipse', {
  x: 4.6, y: 3.5, w: 0.8, h: 0.8,
  fill: { color: C.dark }, line: { color: C.bright, width: 2 }
});

s3.addText('VS', {
  x: 4.6, y: 3.55, w: 0.8, h: 0.7,
  fontSize: 14, bold: true, color: C.bright, align: 'center'
});

// SLIDE 4: MECHANIC
const s4 = prs.addSlide();
s4.background = { color: C.dark };

s4.addText('THE MECHANIC', {
  x: 0.3, y: 0.3, w: 9.4, h: 0.25,
  fontSize: 9, bold: true, color: C.muted, align: 'left'
});

s4.addText('Milestone → Sub-Goals → Progress Bar → Reward → Habit Loop', {
  x: 0.5, y: 0.8, w: 9, h: 0.5,
  fontSize: 24, bold: true, color: C.white
});

const steps = [
  { num: '1', title: 'Milestone Set', desc: 'Vertical onboarding: "Open account" + reward signal' },
  { num: '2', title: 'Sub-Goals', desc: '3–5 micro-behaviors (KYC, Fund, Trade) with progress visibility' },
  { num: '3', title: 'Nudges', desc: 'Push at 33%, 66%, completion. Kept in vertical, not app-wide' },
  { num: '4', title: 'Reward', desc: 'Non-dilutable (RevPoints + feature unlock). High psychological value.' },
  { num: '5', title: 'Loop', desc: 'Completion → next milestone chain. Habit formation by day 60+' }
];

steps.forEach((step, i) => {
  const y = 1.6 + i * 0.85;

  s4.addShape('ellipse', {
    x: 0.5, y: y, w: 0.4, h: 0.4,
    fill: { color: C.bright }, line: { type: 'none' }
  });

  s4.addText(step.num, {
    x: 0.5, y: y, w: 0.4, h: 0.4,
    fontSize: 16, bold: true, color: C.dark, align: 'center'
  });

  s4.addText(step.title, {
    x: 1.1, y: y, w: 2.3, h: 0.22,
    fontSize: 10, bold: true, color: C.white
  });

  s4.addText(step.desc, {
    x: 1.1, y: y + 0.25, w: 2.3, h: 0.25,
    fontSize: 8, color: C.muted
  });

  if (i < 4) {
    s4.addShape('line', {
      x: 0.7, y: y + 0.45, w: 0, h: 0.25,
      line: { color: C.bright, width: 2, dashType: 'dash' }
    });
  }
});

s4.addShape('rect', {
  x: 3.8, y: 1.6, w: 5.7, h: 4.3,
  fill: { color: C.hero }, line: { color: C.bright, width: 1 }
});

s4.addText('Proof: PhonePe Quick Commerce Milestone Waiver (Pincode)', {
  x: 4.0, y: 1.8, w: 5.3, h: 0.4,
  fontSize: 10, bold: true, color: C.bright
});

const proof = [
  'Mechanic: Spend ₹2K in 30 days → unlock ₹300 fee waiver',
  'Same ₹300 value as flat discount, but milestone-framed',
  'A/B test: Milestone vs discount-only control group',
  'Result: 60% improvement in Day-30 activation (milestone group)',
  'Insight: Goal-setting + progress visibility > generic discount'
];

proof.forEach((line, i) => {
  s4.addText(line, {
    x: 4.2, y: 2.4 + i * 0.65, w: 5.1, h: 0.6,
    fontSize: 9, color: C.white, align: 'left'
  });
});

// SLIDE 5: PRODUCT
const s5 = prs.addSlide();
s5.background = { color: C.lgray };

s5.addText('THE PRODUCT', {
  x: 0.3, y: 0.3, w: 9.4, h: 0.25,
  fontSize: 9, bold: true, color: C.muted, align: 'left'
});

s5.addText('4 Screens — Milestone Lifecycle End-to-End', {
  x: 0.5, y: 0.8, w: 9, h: 0.5,
  fontSize: 24, bold: true, color: C.dark
});

const screens = [
  { num: '01', name: 'Welcome', desc: 'Vertical onboarding: "Complete this milestone to unlock 500 RevPoints." Clear, achievable goal visible immediately.' },
  { num: '02', name: 'Progress', desc: 'In-vertical: 3 sub-goals (KYC, Fund, Trade) with progress bar. Nudge card: "You\'re 33% done."' },
  { num: '03', name: 'Success', desc: 'Confetti, reward display, breakdown (100 + 200 + 200 RevPoints), unlock message, next challenge teaser.' },
  { num: '04', name: 'Ops Console', desc: 'Milestone acceptance %, completion rate, Day-30 activation lift, A/B test variant performance, winning rule.' }
];

screens.forEach((scr, i) => {
  const y = 1.6 + i * 1.35;

  s5.addShape('rect', {
    x: 0.5, y: y, w: 9, h: 1.2,
    fill: { color: 'FFFFFF' }, line: { color: C.muted, width: 1 }
  });

  s5.addText(`${scr.num} — ${scr.name}`, {
    x: 0.7, y: y + 0.15, w: 3.2, h: 0.3,
    fontSize: 11, bold: true, color: C.dark
  });

  s5.addText(scr.desc, {
    x: 0.7, y: y + 0.5, w: 8.6, h: 0.55,
    fontSize: 9, color: C.dark, align: 'left'
  });
});

// SLIDE 6: IMPACT & ROI
const s6 = prs.addSlide();
s6.background = { color: C.dark };

s6.addText('IMPACT & ROI', {
  x: 0.3, y: 0.3, w: 9.4, h: 0.25,
  fontSize: 9, bold: true, color: C.muted, align: 'left'
});

s6.addText('New Vertical Activation & Retention Metrics', {
  x: 0.5, y: 0.8, w: 9, h: 0.5,
  fontSize: 22, bold: true, color: C.white
});

const metrics = [
  { icon: '🚀', val: '+110%', label: 'Day-30 Activation', sub: '18% → 38% (milestone cohort)' },
  { icon: '✅', val: '42%', label: 'Milestone Completion', sub: 'by day 60 (vs 18% control)' },
  { icon: '🔄', val: '56%', label: 'Repeat Usage (Month 2)', sub: 'completers vs 22% non-completers' },
  { icon: '💰', val: '~0', label: 'Net Cost', sub: 'RevPoints pre-budgeted; features free' }
];

metrics.forEach((m, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.5 + col * 4.5;
  const y = 1.6 + row * 2;

  s6.addShape('rect', {
    x: x, y: y, w: 4, h: 1.8,
    fill: { color: C.hero }, line: { color: C.bright, width: 1 }
  });

  s6.addText(m.icon, {
    x: x + 0.2, y: y + 0.15, w: 3.6, h: 0.4,
    fontSize: 24, align: 'left'
  });

  s6.addText(m.val, {
    x: x + 0.2, y: y + 0.55, w: 3.6, h: 0.4,
    fontSize: 20, bold: true, color: C.bright, align: 'left'
  });

  s6.addText(m.label, {
    x: x + 0.2, y: y + 1.0, w: 3.6, h: 0.3,
    fontSize: 10, bold: true, color: C.white, align: 'left'
  });

  s6.addText(m.sub, {
    x: x + 0.2, y: y + 1.35, w: 3.6, h: 0.35,
    fontSize: 8, color: C.muted, align: 'left'
  });
});

// SLIDE 7: PROOF OF WORK
const s7 = prs.addSlide();
s7.background = { color: C.lgray };

s7.addText('PROOF OF WORK', {
  x: 0.3, y: 0.3, w: 9.4, h: 0.25,
  fontSize: 9, bold: true, color: C.muted, align: 'left'
});

s7.addText('PhonePe Pincode Milestone Waiver: 60% Day-30 Activation Lift', {
  x: 0.5, y: 0.8, w: 9, h: 0.6,
  fontSize: 24, bold: true, color: C.dark
});

s7.addShape('rect', {
  x: 0.5, y: 1.6, w: 4.3, h: 4.4,
  fill: { color: C.dark }, line: { color: C.bright, width: 1 }
});

s7.addText('PhonePe: What I Shipped', {
  x: 0.7, y: 1.8, w: 3.9, h: 0.3,
  fontSize: 11, bold: true, color: C.bright
});

const phonePe = [
  'Quick commerce (Pincode) Day-30 activation: 40% baseline',
  'Milestone: "Spend ₹2K → unlock ₹300 fee waiver"',
  'Mechanic: goal-setting + progress bar + reward',
  'A/B test: Milestone vs cashback (same ₹300 value)',
  'Result: +60% activation lift in milestone cohort'
];

phonePe.forEach((item, i) => {
  s7.addText(item, {
    x: 0.9, y: 2.3 + i * 0.6, w: 3.7, h: 0.55,
    fontSize: 9, color: C.white
  });
});

s7.addShape('rect', {
  x: 5.2, y: 1.6, w: 4.3, h: 4.4,
  fill: { color: 'FFFFFF' }, line: { color: C.bright, width: 1 }
});

s7.addText('VerticalShift: Application', {
  x: 5.4, y: 1.8, w: 3.9, h: 0.3,
  fontSize: 11, bold: true, color: C.bright
});

const application = [
  'New vertical (Wealth) baseline: 18% Day-30 activation',
  'Milestone: "Open investment account" + 500 RevPoints',
  'Same goal-setting + progress + reward mechanics',
  'Expected lift: +110% (18% → 38%, aligned to PhonePe proof)',
  'Flywheel: milestone → repeat usage → habit'
];

application.forEach((item, i) => {
  s7.addText(item, {
    x: 5.4, y: 2.3 + i * 0.6, w: 3.9, h: 0.55,
    fontSize: 9, color: C.dark
  });
});

// SLIDE 8: ROLLOUT
const s8 = prs.addSlide();
s8.background = { color: C.dark };

s8.addText('ROLLOUT PLAN', {
  x: 0.3, y: 0.3, w: 9.4, h: 0.25,
  fontSize: 9, bold: true, color: C.muted, align: 'left'
});

s8.addText('Three Phases — Pilot to Full Vertical Coverage', {
  x: 0.5, y: 0.8, w: 9, h: 0.5,
  fontSize: 22, bold: true, color: C.white
});

const phases = [
  {
    title: 'Phase 1: Pilot (Month 1–2)',
    bullets: [
      'Deploy VerticalShift to Wealth vertical (Poland)',
      'Milestone: "Open investment account" + 500 RevPoints',
      'A/B: milestone vs control (no milestone) onboarding',
      'Measure: acceptance %, completion rate, Day-30 activation'
    ]
  },
  {
    title: 'Phase 2: Vertical Expansion (Month 3–4)',
    bullets: [
      'Roll out to Lending + Cards verticals',
      'Customize milestones per vertical (credit-specific for Lending)',
      'Run A/B test on milestone copy ("Your first trade" vs "Your first investment")',
      'Scale from 1 vertical → 3 verticals, monitor adoption'
    ]
  },
  {
    title: 'Phase 3: Full Scale (Month 5–6)',
    bullets: [
      'VerticalShift default for all new verticals (future launches)',
      'Milestone template library: ops can spin up new milestone in 30 min',
      'Reward type testing: RevPoints vs feature unlock vs discounts',
      'Ops console tracks all metrics; winning variations auto-promoted'
    ]
  }
];

phases.forEach((phase, i) => {
  const y = 1.5 + i * 1.75;

  s8.addShape('rect', {
    x: 0.5, y: y, w: 9, h: 1.6,
    fill: { color: C.hero }, line: { color: C.bright, width: 1 }
  });

  s8.addText(phase.title, {
    x: 0.7, y: y + 0.15, w: 8.6, h: 0.3,
    fontSize: 11, bold: true, color: C.bright
  });

  phase.bullets.slice(0, 2).forEach((bullet, bi) => {
    s8.addText(`• ${bullet}`, {
      x: 0.9, y: y + 0.55 + bi * 0.3, w: 8.2, h: 0.25,
      fontSize: 8, color: C.white
    });
  });
});

s8.addShape('rect', {
  x: 0.5, y: 6.5, w: 9, h: 0.8,
  fill: { color: C.hero }, line: { color: C.orange, width: 2 }
});

s8.addText('What I Need: 1 engineer (milestone tracking + reward fulfillment) · Product research (copy A/B testing) · Analytics (cohort analysis, repeat engagement tracking)', {
  x: 0.7, y: 6.6, w: 8.6, h: 0.6,
  fontSize: 10, bold: true, color: C.orange
});

s8.addText('Ajay Avaghade | avaghadeajay009@gmail.com | +91 9561558439', {
  x: 0.5, y: 7.15, w: 9, h: 0.25,
  fontSize: 9, color: C.muted, align: 'center'
});

console.log('✓ revolut-verticalshift-deck.pptx created');


  await prs.writeFile({ fileName: "assets/revolut-verticalshift-deck.pptx" });
})();