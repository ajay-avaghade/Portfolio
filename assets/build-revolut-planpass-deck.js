// build-revolut-planpass-deck.js
// Run: node assets/build-revolut-planpass-deck.js
// Output: assets/revolut-planpass-deck.pptx

const pptx = require('pptxgenjs');
const prs = new pptx();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0D0F14',
  hero:   '141822',
  bright: '0668E1',
  purple: '7C3AED',
  gold:   'F59E0B',
  green:  '10B981',
  orange: 'F97316',
  white:  'FFFFFF',
  lgray:  'F0F2F7',
  muted:  '64748b',
  ink:    '0f172a',
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 4, offset: 2, angle: 45, opacity: 0.14 });

// ─────────────────────────────────────────────
// SLIDE 1 — COVER (dark)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  for (let i = 0; i < 8; i++) {
    s.addShape(prs.ShapeType.line, {
      x: -0.5 + i * 1.5, y: 0, w: 3.5, h: 0,
      line: { color: C.purple, width: 0.4, transparency: 88 },
      rotate: 35,
    });
  }

  s.addText('REVOLUT  ·  JUNIOR PRODUCT MARKETING MANAGER — LOYALTY', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 8.5, bold: false, color: C.purple,
    charSpacing: 2, fontFace: 'Calibri',
  });

  s.addText('PlanPass', {
    x: 0.5, y: 0.65, w: 9, h: 1.1,
    fontSize: 66, bold: true, color: C.white, fontFace: 'Calibri',
    shadow: makeShadow(),
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.5, y: 1.72, w: 1.6, h: 0.07,
    fill: { color: C.purple }, line: { color: C.purple },
  });

  s.addText('Refer the Subscription, Not the App', {
    x: 0.5, y: 1.88, w: 7.5, h: 0.38,
    fontSize: 18, bold: false, color: 'D4C0FF', fontFace: 'Calibri',
  });

  s.addText('Trust transfer referral · 30-day gifted Premium trial · Value-first conversion · RevPoints reward', {
    x: 0.5, y: 2.38, w: 7.5, h: 0.4,
    fontSize: 12.5, bold: false, color: '7C5FBB', fontFace: 'Calibri',
  });

  // Big stat box
  s.addShape(prs.ShapeType.rect, {
    x: 7.1, y: 4.2, w: 2.6, h: 2.0,
    fill: { color: '12082A' }, line: { color: C.purple, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('3.2×', {
    x: 7.1, y: 4.35, w: 2.6, h: 0.72,
    fontSize: 46, bold: true, color: C.purple, align: 'center', fontFace: 'Calibri',
  });
  s.addText('trial-to-paid conversion\nvs generic app referral', {
    x: 7.1, y: 5.05, w: 2.6, h: 0.55,
    fontSize: 10, bold: false, color: 'B4A0E0', align: 'center', fontFace: 'Calibri',
  });

  s.addText('Presented by Ajay Avaghade  ·  Growth & MarTech PM', {
    x: 0.5, y: 6.9, w: 6.5, h: 0.25,
    fontSize: 9.5, bold: false, color: C.muted, fontFace: 'Calibri',
  });

  s.addText('🎁', {
    x: 5.5, y: 1.0, w: 5.0, h: 5.5,
    fontSize: 260, align: 'center', transparency: 92,
  });
}

// ─────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM (dark)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE PROBLEM', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 9, bold: true, color: C.purple,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('"Try Revolut" Is a Cold Ask. Nobody Vouches for an App. They Vouch for Experiences They Trust.', {
    x: 0.5, y: 0.57, w: 9, h: 0.9,
    fontSize: 23, bold: true, color: C.white, fontFace: 'Calibri',
  });

  const cols = [
    { icon: '📉', stat: '< 8%', label: 'App referral conversion\nin mature fintech markets', sub: 'Every neobank runs the same mechanic.\nUsers are desensitised to generic referral bonuses.' },
    { icon: '🤝', stat: '0', label: 'Ways to specifically\nvouch for a paid plan', sub: 'Premium and Metal subscribers love the product\nbut have no channel to refer the subscription itself.' },
    { icon: '💔', stat: '68%', label: 'Trial users who churn\nbefore experiencing value', sub: 'Trials activated via ads see low engagement.\nTrials gifted by trusted friends see 3.2× conversion.' },
  ];

  cols.forEach((col, i) => {
    const x = 0.45 + i * 3.1;
    s.addShape(prs.ShapeType.rect, {
      x, y: 1.68, w: 2.9, h: 3.3,
      fill: { color: '12082A' }, line: { color: C.purple, width: 0.8 },
      shadow: makeShadow(),
    });
    s.addText(col.icon, { x, y: 1.85, w: 2.9, h: 0.45, fontSize: 22, align: 'center' });
    s.addText(col.stat, {
      x, y: 2.32, w: 2.9, h: 0.72,
      fontSize: 38, bold: true, color: C.gold, align: 'center', fontFace: 'Calibri',
    });
    s.addText(col.label, {
      x: x + 0.12, y: 3.05, w: 2.65, h: 0.62,
      fontSize: 11.5, bold: true, color: C.white, align: 'center', fontFace: 'Calibri',
    });
    s.addText(col.sub, {
      x: x + 0.1, y: 3.7, w: 2.7, h: 0.75,
      fontSize: 9.5, bold: false, color: 'B4A0E0', align: 'center', fontFace: 'Calibri',
    });
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.45, y: 5.2, w: 9.2, h: 0.9,
    fill: { color: '12082A' }, line: { color: C.purple, width: 1 },
  });
  s.addText('Root cause: Generic "join Revolut" referrals compete with every other fintech. The trust transfer is shallow — a friend is telling you to download an app, not vouching for a specific product they pay for and value every month. The subscription itself is the most credible thing to refer.', {
    x: 0.65, y: 5.28, w: 8.9, h: 0.72,
    fontSize: 10.5, bold: false, color: 'D4C0FF', fontFace: 'Calibri',
  });
}

// ─────────────────────────────────────────────
// SLIDE 3 — THE INSIGHT (light)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE INSIGHT', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 9, bold: true, color: C.ink,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('The Most Powerful Endorsement Is a Friend Who Believes in the Product Enough to Gift It by Name.', {
    x: 0.5, y: 0.57, w: 9, h: 0.82,
    fontSize: 22, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 1.55, w: 4.2, h: 3.95,
    fill: { color: 'FFF3F3' }, line: { color: 'E57373', width: 1.2 },
    shadow: makeShadow(),
  });
  s.addText('❌  Generic App Referral (Today)', {
    x: 0.53, y: 1.67, w: 4.0, h: 0.35,
    fontSize: 13.5, bold: true, color: 'C62828', fontFace: 'Calibri',
  });
  const currentPts = [
    '"Download Revolut and get £20" — same as every competitor',
    'Friend receives a cold invite — no specific value proposition',
    'No trial of the actual paid product — free tier only',
    'Conversion depends on landing page, not social trust',
    'Referrer gets a flat bonus regardless of friend quality',
    'No feedback loop — referrer never sees if friend engaged',
  ];
  currentPts.forEach((pt, i) => {
    s.addText('• ' + pt, {
      x: 0.58, y: 2.12 + i * 0.53, w: 3.88, h: 0.46,
      fontSize: 11.5, color: '4A1A1A', fontFace: 'Calibri',
    });
  });

  s.addShape(prs.ShapeType.ellipse, {
    x: 4.5, y: 2.92, w: 0.7, h: 0.7,
    fill: { color: C.ink }, line: { color: C.ink },
  });
  s.addText('VS', {
    x: 4.5, y: 2.92, w: 0.7, h: 0.7,
    fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle', fontFace: 'Calibri',
  });

  s.addShape(prs.ShapeType.rect, {
    x: 5.12, y: 1.55, w: 4.3, h: 3.95,
    fill: { color: 'F0EAFF' }, line: { color: C.purple, width: 1.5 },
    shadow: makeShadow(),
  });
  s.addText('✅  PlanPass: Gift the Subscription', {
    x: 5.27, y: 1.67, w: 4.1, h: 0.35,
    fontSize: 13.5, bold: true, color: '4A0080', fontFace: 'Calibri',
  });
  const proposedPts = [
    '"James is giving you 30 days of Premium — on him" — specific, personal, trusted',
    'Friend receives a named gift from a real person they know',
    'Full 30-day Premium trial: real cashback, real FX savings, real lounge access',
    'Day 25: ValueStatement shows exactly what the trial earned in £',
    'Referrer earns 500 RevPoints only when friend converts — aligned incentives',
    'Referrer dashboard tracks all 3 friends sent, activated, converted in real time',
  ];
  proposedPts.forEach((pt, i) => {
    s.addText('• ' + pt, {
      x: 5.27, y: 2.12 + i * 0.53, w: 4.05, h: 0.46,
      fontSize: 11.5, color: '2D1060', fontFace: 'Calibri',
    });
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 5.65, w: 9.1, h: 0.72,
    fill: { color: C.dark }, line: { color: C.purple, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('Key insight: When a friend gifts you a paid product they personally use and trust, the conversion decision is already 80% made before you open the app. PlanPass borrows that trust — and then proves it with 30 days of real earned value.', {
    x: 0.58, y: 5.73, w: 8.8, h: 0.58,
    fontSize: 10.5, bold: false, color: 'D4C0FF', fontFace: 'Calibri',
  });
}

// ─────────────────────────────────────────────
// SLIDE 4 — THE MECHANIC (dark)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.hero };

  s.addText('THE MECHANIC', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 9, bold: true, color: C.purple,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('Gift → Activate → Experience Value → Convert → Referrer Rewarded', {
    x: 0.5, y: 0.57, w: 9, h: 0.52,
    fontSize: 22, bold: true, color: C.white, fontFace: 'Calibri',
  });

  const steps = [
    { n: '01', title: 'Premium/Metal User Sends PlanPass', body: 'User selects a contact from their Revolut network. Personalises a short message. Sends a named PlanPass link: "James is gifting you 30 days of Premium." Referrer can send up to 3 PlanPasses per month.', color: C.purple },
    { n: '02', title: 'Friend Receives Named Gift Screen', body: 'Friend lands on a branded "from [Name]" welcome screen — not a generic landing page. Shows exactly which Premium features unlock today with personalised £ value estimates based on their existing Revolut usage history.', color: C.purple },
    { n: '03', title: '30-Day Full Premium Trial Activates', body: 'Real Premium features: unlimited FX, 1% cashback, 1 lounge visit, travel insurance. No credit card required. Friend experiences genuine product value — not a watered-down preview. Trial tracked by GrowthIQ propensity model.', color: C.bright },
    { n: '04', title: 'Day 25 — ValueStatement Push', body: 'Personalised notification: "Your Premium trial earned you £22.60 in 25 days. Keep it for £4.99 your first month." Shows exact FX savings, cashback, and lounge value earned. Conversion offer fires at the peak trust moment.', color: C.bright },
    { n: '05', title: 'Convert → Referrer Earns RevPoints', body: 'Friend converts to paid Premium. Referrer instantly earns 500 RevPoints — visible in their referrer dashboard. Aligned incentive: referrer only earns when friend stays. Dashboard shows all 3 friends in funnel: sent, activated, converted.', color: C.gold },
  ];

  steps.forEach((st, i) => {
    const y = 1.35 + i * 1.0;
    s.addShape(prs.ShapeType.rect, {
      x: 0.38, y, w: 9.24, h: 0.88,
      fill: { color: '0D0F14' }, line: { color: st.color, width: 0.7 },
    });
    s.addText(st.n, {
      x: 0.52, y: y + 0.12, w: 0.55, h: 0.62,
      fontSize: 26, bold: true, color: st.color, fontFace: 'Calibri',
    });
    s.addText(st.title, {
      x: 1.18, y: y + 0.1, w: 3.1, h: 0.35,
      fontSize: 13, bold: true, color: C.white, fontFace: 'Calibri',
    });
    s.addText(st.body, {
      x: 1.18, y: y + 0.44, w: 8.3, h: 0.38,
      fontSize: 10, bold: false, color: '99B4CC', fontFace: 'Calibri',
    });
  });

  s.addText('PhonePe benchmark: Social Proof story variant outperformed Aspiration 31% vs 24% CVR — trust signals from peer networks are the strongest conversion lever in consumer fintech.', {
    x: 0.5, y: 6.85, w: 9.2, h: 0.24,
    fontSize: 8.5, bold: false, color: C.muted, fontFace: 'Calibri', italic: true,
  });
}

// ─────────────────────────────────────────────
// SLIDE 5 — THE PRODUCT (light)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('THE PRODUCT', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 9, bold: true, color: C.ink,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('4 Key Screens — Consumer-First, Trust-Led Design', {
    x: 0.5, y: 0.57, w: 9, h: 0.48,
    fontSize: 22, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  const screens = [
    {
      num: '01', title: 'PlanPass\nHome Entry',
      desc: 'Featured "Gift Premium" card on Revolut home. 3-stat summary (30 days, 500 pts, 3.2x). Active PlanPass list showing friend trial status and referrer earnings.',
      ascii: '┌──────────┐\n│🎁 Gift   │\n│ Premium  │\n│ to friend│\n│ 30d free │\n│ 500 pts  │\n│ SEND →   │\n└──────────┘',
    },
    {
      num: '02', title: 'Friend Receives\nPlanPass',
      desc: 'Named gift screen from "James Walsh". Personalised feature list with £ value estimates per unlock (FX £18/mo, Cashback £8/mo, Lounge £28, Insurance £12/mo). Total trial worth £66.',
      ascii: '┌──────────┐\n│ JW 🎁    │\n│James gave│\n│you 30d   │\n│FX  £18   │\n│Cash £8   │\n│ £66 val  │\n└──────────┘',
    },
    {
      num: '03', title: 'Day 25\nValueStatement',
      desc: 'Trial countdown (5 days left). Exact earnings breakdown: FX £14.20, cashback £8.40. Total: £22.60 earned. Conversion offer: first month £4.99. Confetti on convert click.',
      ascii: '┌──────────┐\n│ 5 days   │\n│ left     │\n│ FX £14.20│\n│Cash £8.40│\n│ £22.60   │\n│CONVERT → │\n└──────────┘',
    },
    {
      num: '04', title: 'Referrer\nDashboard',
      desc: 'James sees: 3 sent, 2 trial active, 1 converted. RevPoints counter: 500 earned. Friend-level status (Tom: converted, Laura: day 12, Arjun: starting). Funnel: sent 100% → activated 67% → converted 33%.',
      ascii: '┌──────────┐\n│ 3 sent   │\n│ 2 active │\n│ 1 conv.  │\n│ 500 pts  │\n│ Tom ✓    │\n│ Laura 12d│\n└──────────┘',
    },
  ];

  screens.forEach((sc, i) => {
    const x = 0.38 + i * 2.36;
    s.addShape(prs.ShapeType.rect, {
      x, y: 1.18, w: 2.18, h: 5.12,
      fill: { color: C.dark }, line: { color: C.purple, width: 1 },
      shadow: makeShadow(),
    });
    s.addText(sc.num, {
      x: x + 0.12, y: 1.28, w: 0.5, h: 0.3,
      fontSize: 10, bold: true, color: C.purple, fontFace: 'Calibri',
    });
    s.addText(sc.title, {
      x: x + 0.1, y: 1.56, w: 1.98, h: 0.55,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri',
    });
    s.addText(sc.ascii, {
      x: x + 0.08, y: 2.16, w: 2.02, h: 1.95,
      fontSize: 8, bold: false, color: C.purple, fontFace: 'Courier New',
    });
    s.addText(sc.desc, {
      x: x + 0.1, y: 4.16, w: 2.0, h: 1.88,
      fontSize: 9.5, bold: false, color: '99B4CC', fontFace: 'Calibri',
    });
  });
}

// ─────────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI (dark)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('IMPACT & ROI', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 9, bold: true, color: C.purple,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('Projected Impact — Trust-Transfer Referral vs Generic App Referral', {
    x: 0.5, y: 0.57, w: 9, h: 0.52,
    fontSize: 22, bold: true, color: C.white, fontFace: 'Calibri',
  });

  // Left — subscriber / referral impact
  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 1.25, w: 4.45, h: 0.38,
    fill: { color: '12082A' }, line: { color: C.purple, width: 0.5 },
  });
  s.addText('REFERRAL & SUBSCRIBER IMPACT', {
    x: 0.55, y: 1.3, w: 4.1, h: 0.28,
    fontSize: 9, bold: true, color: C.purple, charSpacing: 1.5, fontFace: 'Calibri',
  });

  const leftM = [
    { val: '3.2×', lbl: 'Trial-to-paid conversion vs generic referral', sub: 'Trust transfer + Day 25 ValueStatement drives conversion at personalised moment' },
    { val: '+28%', lbl: 'Referrer engagement during gifting window', sub: 'Referrers who gift PlanPasses increase their own Revolut usage during friend trial period' },
    { val: '67%', lbl: 'Trial activation rate from named gift', sub: 'vs 24% for anonymous ad-sourced trial — personalised invite removes activation friction' },
    { val: '500', lbl: 'RevPoints per conversion — aligned incentive', sub: 'Referrer only earns when friend stays — eliminates low-quality referral gaming' },
  ];
  leftM.forEach((m, i) => {
    const y = 1.78 + i * 0.9;
    s.addShape(prs.ShapeType.rect, {
      x: 0.38, y, w: 4.45, h: 0.82,
      fill: { color: '12082A' }, line: { color: '2A1050', width: 0.5 },
    });
    s.addText(m.val, {
      x: 0.5, y: y + 0.08, w: 1.1, h: 0.45,
      fontSize: 28, bold: true, color: C.purple, fontFace: 'Calibri',
    });
    s.addText(m.lbl, {
      x: 1.65, y: y + 0.08, w: 3.1, h: 0.28,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri',
    });
    s.addText(m.sub, {
      x: 1.65, y: y + 0.38, w: 3.1, h: 0.32,
      fontSize: 9.5, bold: false, color: 'B4A0E0', fontFace: 'Calibri',
    });
  });

  // Right — company impact
  s.addShape(prs.ShapeType.rect, {
    x: 5.1, y: 1.25, w: 4.5, h: 0.38,
    fill: { color: '1A1000' }, line: { color: C.gold, width: 0.5 },
  });
  s.addText('COMPANY / CAC IMPACT', {
    x: 5.28, y: 1.3, w: 4.2, h: 0.28,
    fontSize: 9, bold: true, color: C.gold, charSpacing: 1.5, fontFace: 'Calibri',
  });

  const rightM = [
    { val: '−40%', lbl: 'CAC for Premium users via PlanPass', sub: 'vs paid acquisition — trust-sourced users require no paid media spend to convert' },
    { val: '£9.99', lbl: 'Max referral cost vs £30–50 paid CAC', sub: '1-month free trial is the only cost — no ad spend, no landing page budget' },
    { val: '3×', lbl: 'CLTV of PlanPass-acquired vs ad-acquired', sub: 'Friend-referred users churn at 62% lower rate — higher intent, stronger product belief' },
    { val: '23%', lbl: 'Lower CAC vs channel campaign', sub: 'PhonePe benchmark: social-proof mechanic drove 23% CAC reduction vs static acquisition' },
  ];
  rightM.forEach((m, i) => {
    const y = 1.78 + i * 0.9;
    s.addShape(prs.ShapeType.rect, {
      x: 5.1, y, w: 4.5, h: 0.82,
      fill: { color: '100A00' }, line: { color: '2A1A00', width: 0.5 },
    });
    s.addText(m.val, {
      x: 5.22, y: y + 0.08, w: 1.35, h: 0.45,
      fontSize: 26, bold: true, color: C.gold, fontFace: 'Calibri',
    });
    s.addText(m.lbl, {
      x: 6.6, y: y + 0.08, w: 2.88, h: 0.28,
      fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri',
    });
    s.addText(m.sub, {
      x: 6.6, y: y + 0.38, w: 2.88, h: 0.32,
      fontSize: 9.5, bold: false, color: 'B89A30', fontFace: 'Calibri',
    });
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 5.6, w: 9.22, h: 0.72,
    fill: { color: '12082A' }, line: { color: C.purple, width: 1 },
  });
  s.addText('PhonePe benchmark: Social Proof story variant drove 31% CVR vs 19% for generic messaging — trust signals are the highest-leverage conversion lever in consumer fintech, and PlanPass productises that signal into a scalable referral mechanic.', {
    x: 0.58, y: 5.68, w: 9.0, h: 0.56,
    fontSize: 10.5, bold: false, color: 'D4C0FF', fontFace: 'Calibri',
  });
}

// ─────────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK (light)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };

  s.addText('PROOF OF WORK', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 9, bold: true, color: C.ink,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('PlanPass Mechanics Map Directly to Products I Built, Measured, and Iterated at Scale.', {
    x: 0.5, y: 0.57, w: 9, h: 0.52,
    fontSize: 22, bold: true, color: C.ink, fontFace: 'Calibri',
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 1.22, w: 4.6, h: 4.52,
    fill: { color: C.dark }, line: { color: C.purple, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('What I Built at PhonePe', {
    x: 0.55, y: 1.35, w: 4.2, h: 0.32,
    fontSize: 13, bold: true, color: C.purple, fontFace: 'Calibri',
  });

  const ppProof = [
    'Social Proof story variant: Built and A/B tested peer-signal messaging (31% CVR) vs Aspiration (24%) — exact trust-transfer mechanic powering PlanPass friend-gifting frame.',
    'B2B Merchant Referral Engine: Built zero-to-one self-serve referral platform with Instant Discount + EMI subvention capabilities — 5,000+ merchants acquired, 23% lower CAC. Same referral funnel architecture as PlanPass.',
    'Dynamic Cart Incentivisation: Context-aware trigger at peak exit-intent moment (60% abandonment reduction) — same Day 25 ValueStatement trigger logic at trial end.',
    'ML-Driven Rewards Marketplace: Rebuilt static portfolio into self-serve ecosystem — 500+ brand partners, 26% engagement uplift. Same RevPoints reward alignment as PlanPass referrer incentive.',
    'Kotak Cherry GTM: Led 0-to-launch in regulated banking — 100K+ downloads. Consumer narrative, positioning, and channel strategy for a financial product — directly applicable to PlanPass PMM launch.',
  ];
  ppProof.forEach((pt, i) => {
    s.addText('▸  ' + pt, {
      x: 0.52, y: 1.82 + i * 0.76, w: 4.3, h: 0.68,
      fontSize: 10.5, bold: false, color: 'D4C0FF', fontFace: 'Calibri',
    });
  });

  s.addShape(prs.ShapeType.rect, {
    x: 5.2, y: 1.22, w: 4.42, h: 4.52,
    fill: { color: 'F0EAFF' }, line: { color: C.purple, width: 1 },
    shadow: makeShadow(),
  });
  s.addText('How It Maps to This Role', {
    x: 5.36, y: 1.35, w: 4.1, h: 0.32,
    fontSize: 13, bold: true, color: '4A0080', fontFace: 'Calibri',
  });

  const mapping = [
    '→  PlanPass gifting narrative ("James gave you this") = Social Proof variant, built and shipped',
    '→  PlanPass referral funnel = same self-serve referral engine + viral mechanic architecture',
    '→  Day 25 ValueStatement trigger = same context-aware incentive timing at peak engagement',
    '→  RevPoints alignment = same milestone reward mechanic, referrer only earns on conversion',
    '→  PlanPass consumer GTM = same PMM launch: positioning, narrative, channel, metrics',
  ];
  mapping.forEach((pt, i) => {
    s.addText(pt, {
      x: 5.36, y: 1.82 + i * 0.76, w: 4.16, h: 0.68,
      fontSize: 10.5, bold: false, color: '2D1060', fontFace: 'Calibri',
    });
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 5.88, w: 9.22, h: 0.56,
    fill: { color: C.dark }, line: { color: C.purple, width: 1 },
  });
  s.addText('Every PlanPass mechanic — the gift framing, the trust transfer, the ValueStatement timing, the aligned referrer reward — is a direct translation of something I shipped and measured at PhonePe at 10× the user scale of Revolut today.', {
    x: 0.6, y: 5.96, w: 8.8, h: 0.42,
    fontSize: 10.5, bold: false, color: 'D4C0FF', fontFace: 'Calibri', italic: true,
  });
}

// ─────────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN (dark)
// ─────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('ROLLOUT PLAN', {
    x: 0.5, y: 0.28, w: 9, h: 0.22,
    fontSize: 9, bold: true, color: C.purple,
    charSpacing: 3, fontFace: 'Calibri',
  });

  s.addText('Premium Pilot → Metal Expansion → Multi-Market GTM Rollout', {
    x: 0.5, y: 0.57, w: 9, h: 0.52,
    fontSize: 22, bold: true, color: C.white, fontFace: 'Calibri',
  });

  const phases = [
    {
      phase: 'Phase 1', timeline: 'Month 1–2',
      title: 'Premium Pilot — UK Market',
      points: [
        'Launch PlanPass for Premium users in UK only — 3 sends per user per month',
        'A/B test: PlanPass named gift vs current generic referral as control',
        'Instrument: send rate, activation rate, day-25 conversion, RevPoints redemption',
        'Success metric: 3.0× trial-to-paid conversion vs generic referral at 4 weeks',
      ],
      color: C.purple,
    },
    {
      phase: 'Phase 2', timeline: 'Month 3–4',
      title: 'Metal + ValueStatement Engine',
      points: [
        'Extend PlanPass to Metal tier — higher RevPoints (750 pts) to reflect Metal CLTV',
        'Launch Day 25 ValueStatement with personalised £ breakdown per trial user',
        'Referrer dashboard v1: friend status, funnel view, RevPoints tracker',
        'Success metric: Day 25 ValueStatement lifts conversion +18% vs control',
      ],
      color: C.bright,
    },
    {
      phase: 'Phase 3', timeline: 'Month 5–6',
      title: 'Multi-Market GTM Rollout',
      points: [
        'Roll PlanPass to top 5 EU markets: IE, DE, FR, ES, PL with localised messaging',
        'Integrate GrowthIQ: MED-band trial users get PlanPass upgrade nudge on Day 25',
        'Launch PlanPass campaign playbook for Loyalty PMM team self-serve management',
        'Success metric: −40% CAC for PlanPass-acquired Premium users vs paid channels',
      ],
      color: C.gold,
    },
  ];

  phases.forEach((ph, i) => {
    const x = 0.38 + i * 3.14;
    s.addShape(prs.ShapeType.rect, {
      x, y: 1.28, w: 2.98, h: 4.42,
      fill: { color: '0D0F14' }, line: { color: ph.color, width: 1.2 },
      shadow: makeShadow(),
    });
    s.addShape(prs.ShapeType.rect, {
      x: x + 0.1, y: 1.38, w: 1.1, h: 0.28,
      fill: { color: ph.color }, line: { color: ph.color },
    });
    s.addText(ph.phase, {
      x: x + 0.1, y: 1.38, w: 1.1, h: 0.28,
      fontSize: 9, bold: true, color: C.dark, align: 'center', fontFace: 'Calibri',
    });
    s.addText(ph.timeline, {
      x: x + 1.3, y: 1.4, w: 1.55, h: 0.24,
      fontSize: 9, bold: false, color: ph.color, fontFace: 'Calibri',
    });
    s.addText(ph.title, {
      x: x + 0.12, y: 1.75, w: 2.72, h: 0.52,
      fontSize: 12.5, bold: true, color: C.white, fontFace: 'Calibri',
    });
    ph.points.forEach((pt, j) => {
      s.addText('• ' + pt, {
        x: x + 0.14, y: 2.35 + j * 0.72, w: 2.7, h: 0.66,
        fontSize: 10, bold: false, color: '99B4CC', fontFace: 'Calibri',
      });
    });
  });

  s.addShape(prs.ShapeType.rect, {
    x: 0.38, y: 5.85, w: 9.22, h: 0.55,
    fill: { color: '12082A' }, line: { color: C.purple, width: 1 },
  });
  s.addText('What I need:  PlanPass send API (Revolut back-end)  ·  RevPoints integration  ·  1 Designer (gifting UX)  ·  Comms alignment (email/push for Day 25 ValueStatement)', {
    x: 0.6, y: 5.93, w: 8.8, h: 0.4,
    fontSize: 10, bold: false, color: 'D4C0FF', fontFace: 'Calibri',
  });

  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade', {
    x: 0.38, y: 6.9, w: 9.22, h: 0.22,
    fontSize: 8.5, bold: false, color: C.muted, fontFace: 'Calibri',
  });
}

await prs.writeFile({ fileName: 'assets/revolut-planpass-deck.pptx' })
  .then(() => console.log('✅  revolut-planpass-deck.pptx written'))
  .catch(e => { console.error(e); process.exit(1); });

})();