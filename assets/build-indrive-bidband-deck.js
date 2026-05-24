const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0B0E13',
  hero:   '141810',
  bright: 'B5E033',
  accent: '8BC220',
  orange: 'FF6B35',
  red:    'EF4444',
  amber:  'F59E0B',
  green:  '22C55E',
  white:  'FFFFFF',
  lgray:  'F1F5F9',
  muted:  '94A3B8',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.15 });

// ─────────────────────────────────────────────
// SLIDE 1 — COVER
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  // diagonal lines watermark
  for (let i = 0; i < 14; i++) {
    sl.addShape(prs.ShapeType.line, {
      x: -1 + i * 0.9, y: 0, w: 4, h: 7.5,
      line: { color: C.bright, width: 0.4, transparency: 88 },
    });
  }

  // large watermark character
  sl.addText('⟷', {
    x: 6.2, y: 0.4, w: 3.5, h: 3.5,
    fontSize: 200, color: C.bright, transparency: 93,
    fontFace: 'Sora', bold: true,
  });

  // eyebrow
  sl.addText('INDRIVE  ·  SENIOR PRODUCT MANAGER — GROWTH & PERSONALIZATION', {
    x: 0.5, y: 0.5, w: 7, h: 0.3,
    fontSize: 9, color: C.bright, bold: true,
    charSpacing: 2,
  });

  // title
  sl.addText('inDrive BidBand', {
    x: 0.5, y: 1.1, w: 7.5, h: 1.0,
    fontSize: 52, color: C.white, bold: true, fontFace: 'Arial',
  });

  // subtitle
  sl.addText('Fixing the worst UX in ride-hailing:\nshow 4 fare bands with match probabilities\nbefore the passenger negotiates — not after failure.', {
    x: 0.5, y: 2.2, w: 7.2, h: 1.3,
    fontSize: 17, color: C.muted, lineSpacingMultiple: 1.35,
  });

  // presenter line
  sl.addText('Presented by Ajay Avaghade  ·  Growth & Personalization PM', {
    x: 0.5, y: 3.7, w: 6, h: 0.3,
    fontSize: 11, color: C.accent,
  });

  // big stat block bottom-right
  sl.addShape(prs.ShapeType.roundRect, {
    x: 7.2, y: 5.2, w: 2.5, h: 2.0,
    fill: { color: C.bright, transparency: 10 },
    line: { color: C.bright, width: 1 },
    rectRadius: 0.12,
  });
  sl.addText('+28%', {
    x: 7.2, y: 5.35, w: 2.5, h: 0.8,
    fontSize: 44, color: '0B0E13', bold: true, align: 'center', fontFace: 'Arial',
  });
  sl.addText('match success rate\nvs single-price baseline', {
    x: 7.2, y: 6.15, w: 2.5, h: 0.6,
    fontSize: 10, color: '1A2A00', align: 'center', lineSpacingMultiple: 1.25,
  });
}

// ─────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE PROBLEM', {
    x: 0.5, y: 0.4, w: 9, h: 0.3,
    fontSize: 9, color: C.bright, bold: true, charSpacing: 2,
  });
  sl.addText('A Single Suggested Price Destroys Trust\nAt The Moment That Matters Most', {
    x: 0.5, y: 0.75, w: 9, h: 1.0,
    fontSize: 28, color: C.white, bold: true, lineSpacingMultiple: 1.25,
  });

  // 3 stat columns
  const cols3 = [
    { icon:'✗', val:'~35%', lbl:'of fare offers\nget no match', sub:'User hits a dead end with\nzero useful information', color: C.red },
    { icon:'⏱', val:'60 sec', lbl:'then a jarring new\nprice recommendation', sub:'Gap feels manipulative — trust\nbreak at the critical moment', color: C.amber },
    { icon:'📉', val:'High churn', lbl:'users who fail once\noften never return', sub:'Competitor apps benefit from\nevery inDrive no-match event', color: C.orange },
  ];
  cols3.forEach((item, i) => {
    const x = 0.5 + i * 3.2;
    sl.addShape(prs.ShapeType.roundRect, {
      x, y: 1.95, w: 3.0, h: 3.0,
      fill: { color: '1A1A1A', transparency: 0 },
      line: { color: item.color, width: 1.2 },
      rectRadius: 0.1,
      shadow: makeShadow(),
    });
    sl.addText(item.icon, { x, y: 2.05, w: 3.0, h: 0.5, fontSize: 22, align:'center', color: item.color });
    sl.addText(item.val, { x, y: 2.55, w: 3.0, h: 0.7, fontSize: 36, color: item.color, bold:true, align:'center', fontFace:'Arial' });
    sl.addText(item.lbl, { x, y: 3.25, w: 3.0, h: 0.55, fontSize: 12, color: C.white, bold:true, align:'center', lineSpacingMultiple:1.2 });
    sl.addText(item.sub, { x, y: 3.8, w: 3.0, h: 0.7, fontSize: 10, color: C.muted, align:'center', lineSpacingMultiple:1.3 });
  });

  // Insight box
  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.5, y: 5.2, w: 9.0, h: 0.9,
    fill: { color: '1A1200', transparency: 0 },
    line: { color: C.amber, width: 1 },
    rectRadius: 0.08,
  });
  sl.addText('Root cause: The negotiation model — inDrive\'s biggest differentiator — becomes its biggest trust destroyer when it fails silently. The app gives passengers no information to make a better decision. The fix is not a higher default price. The fix is transparency before the bid.', {
    x: 0.7, y: 5.3, w: 8.6, h: 0.7,
    fontSize: 10.5, color: C.white, lineSpacingMultiple: 1.35,
  });
}

// ─────────────────────────────────────────────
// SLIDE 3 — THE INSIGHT
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE INSIGHT', {
    x: 0.5, y: 0.4, w: 9, h: 0.3,
    fontSize: 9, color: C.accent, bold: true, charSpacing: 2,
  });
  sl.addText('The Problem Is Not The Price — It\'s The Missing Information', {
    x: 0.5, y: 0.75, w: 9, h: 0.7,
    fontSize: 26, color: C.dark, bold: true, lineSpacingMultiple: 1.2,
  });

  // Left: Current
  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.5, y: 1.65, w: 4.2, h: 3.6,
    fill: { color: 'FFF0EE', transparency: 0 },
    line: { color: C.red, width: 1.5 },
    rectRadius: 0.1,
  });
  sl.addText('❌  Single Price (Current)', {
    x: 0.7, y: 1.8, w: 3.8, h: 0.4,
    fontSize: 13, color: C.red, bold: true,
  });
  const leftPoints = [
    'One suggested fare shown with no context',
    'No indication of match probability before submitting',
    'No-match = dead end UX with zero information',
    'After 60s a much higher price appears — feels like manipulation',
    'User churns to competitor app — lost GMV, lost trust',
  ];
  leftPoints.forEach((pt, i) => {
    sl.addText(`• ${pt}`, {
      x: 0.7, y: 2.3 + i * 0.52, w: 3.8, h: 0.45,
      fontSize: 11, color: '444444', lineSpacingMultiple: 1.25,
    });
  });

  // VS divider
  sl.addShape(prs.ShapeType.ellipse, {
    x: 4.55, y: 3.0, w: 0.9, h: 0.9,
    fill: { color: C.dark },
    line: { color: C.dark, width: 0 },
  });
  sl.addText('VS', { x: 4.55, y: 3.0, w: 0.9, h: 0.9, fontSize: 12, color: C.white, bold: true, align:'center', valign:'middle' });

  // Right: Proposed
  sl.addShape(prs.ShapeType.roundRect, {
    x: 5.3, y: 1.65, w: 4.2, h: 3.6,
    fill: { color: 'F0FFF4', transparency: 0 },
    line: { color: C.accent, width: 1.5 },
    rectRadius: 0.1,
  });
  sl.addText('✅  BidBand (Proposed)', {
    x: 5.5, y: 1.8, w: 3.8, h: 0.4,
    fontSize: 13, color: C.accent, bold: true,
  });
  const rightPoints = [
    '4 fare bands shown before the passenger commits',
    'Each band shows match probability (20% / 45% / 72% / 91%)',
    'User self-selects their risk tolerance — informed choice',
    'No-match triggers a one-tap nudge to the next band — not a jarring new suggestion',
    'Recovers 43% of Band 1 failures via the progressive nudge flow',
  ];
  rightPoints.forEach((pt, i) => {
    sl.addText(`• ${pt}`, {
      x: 5.5, y: 2.3 + i * 0.52, w: 3.8, h: 0.45,
      fontSize: 11, color: '1A3A00', lineSpacingMultiple: 1.25,
    });
  });

  // Callout
  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.5, y: 5.55, w: 9.0, h: 0.8,
    fill: { color: '0B0E13', transparency: 0 },
    line: { color: C.bright, width: 1 },
    rectRadius: 0.08,
  });
  sl.addText('Key insight: This is only possible at inDrive. Uber and Bolt cannot build BidBand because they have no negotiation layer. inDrive\'s unique counteroffer model is the moat — and BidBand converts that moat into a transparency advantage.', {
    x: 0.7, y: 5.65, w: 8.6, h: 0.6,
    fontSize: 10.5, color: C.white, lineSpacingMultiple: 1.3,
  });
}

// ─────────────────────────────────────────────
// SLIDE 4 — THE MECHANIC
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('THE MECHANIC', {
    x: 0.5, y: 0.4, w: 9, h: 0.3,
    fontSize: 9, color: C.bright, bold: true, charSpacing: 2,
  });
  sl.addText('Four Bands, One Informed Decision — Then Progressive Recovery', {
    x: 0.5, y: 0.75, w: 9, h: 0.6,
    fontSize: 24, color: C.white, bold: true,
  });

  const steps = [
    { n:'01', title:'Route Submitted', body:'Passenger enters pickup + destination. BidBand ML queries real-time driver density, time of day, historical acceptance rates, and weather for this corridor.', color: C.red },
    { n:'02', title:'4 Bands Rendered', body:'Fare entry screen shows 4 bands with probability bars: Try Your Luck (20%), Fair Offer (45%), Good Chance (72%), Fast Match (91%). User taps to select.', color: C.amber },
    { n:'03', title:'Offer Submitted', body:'Selected band fare submitted to driver pool. Live feed shows how many drivers viewed, countered, or declined — full transparency throughout the wait.', color: C.bright },
    { n:'04', title:'No-Match Nudge', body:'If no driver accepts within 3 min at Band N, a one-tap nudge appears: "Move to Band N+1 for X% better odds." Progressive, never jarring. No surprise price jump.', color: C.orange },
    { n:'05', title:'Match + Insight', body:'Driver accepts. Confirmation screen shows agreed fare, band used, time to match, and how much was saved vs Band 4 — reinforcing the value of the informed choice.', color: C.green },
  ];

  steps.forEach((s, i) => {
    const x = 0.4 + i * 1.84;
    sl.addShape(prs.ShapeType.roundRect, {
      x, y: 1.55, w: 1.6, h: 3.9,
      fill: { color: '141810', transparency: 0 },
      line: { color: s.color, width: 1.2 },
      rectRadius: 0.1,
      shadow: makeShadow(),
    });
    sl.addText(s.n, { x, y: 1.65, w: 1.6, h: 0.45, fontSize: 22, color: s.color, bold: true, align:'center', fontFace:'Arial' });
    sl.addText(s.title, { x: x+0.08, y: 2.1, w: 1.44, h: 0.45, fontSize: 10.5, color: C.white, bold: true, align:'center', lineSpacingMultiple:1.2 });
    sl.addText(s.body, { x: x+0.08, y: 2.6, w: 1.44, h: 2.5, fontSize: 9, color: C.muted, lineSpacingMultiple:1.35, align:'center' });
    // connector line (not last)
    if (i < 4) {
      sl.addShape(prs.ShapeType.line, {
        x: x + 1.6, y: 2.0, w: 0.24, h: 0,
        line: { color: C.muted, width: 0.8, dashType: 'dash' },
      });
    }
  });

  // Proof note
  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.4, y: 5.6, w: 9.2, h: 0.75,
    fill: { color: '1A2A00', transparency: 0 },
    line: { color: C.bright, width: 1 },
    rectRadius: 0.08,
  });
  sl.addText('A/B proof at PhonePe: Dynamic offer eligibility surface (show band / suppress) drove 22% checkout conversion lift at 350M MAU — same real-time signal-to-UI architecture as BidBand\'s probability rendering layer.', {
    x: 0.6, y: 5.7, w: 8.8, h: 0.55,
    fontSize: 10, color: C.white, lineSpacingMultiple: 1.3,
  });
}

// ─────────────────────────────────────────────
// SLIDE 5 — THE PRODUCT
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('THE PRODUCT', {
    x: 0.5, y: 0.4, w: 9, h: 0.3,
    fontSize: 9, color: C.accent, bold: true, charSpacing: 2,
  });
  sl.addText('4 Screen States — Built for inDrive\'s Negotiation Model', {
    x: 0.5, y: 0.75, w: 9, h: 0.5,
    fontSize: 24, color: C.dark, bold: true,
  });

  const screens = [
    { n:'01', title:'Fare Entry', desc:'4 BidBand tiles with probability bars. User taps their risk band. Fair Offer pre-selected. Single submit CTA. No cognitive overload.', mock: ['[Band 1: Try Luck 20%]','[Band 2: Fair Offer 45%] ◀','[Band 3: Good Chance 72%]','[Band 4: Fast Match 91%]','[Request Ride]'], color: C.red },
    { n:'02', title:'Waiting State', desc:'Live driver activity feed: viewed, countered, declined. Progressive nudge card at 3 min — never a jarring new price, always a gentle one-step upgrade.', mock: ['Finding driver…  $11.20','14 drivers viewed  3 countered','[========= 42%]','Still waiting? Try Band 3','[Move to Good Chance $12.60]'], color: C.amber },
    { n:'03', title:'Match Found', desc:'Driver accepted — shows which band succeeded, time to match, savings vs Band 4. Insight reinforces that the informed bid worked.', mock: ['✓ Driver Accepted!','Band 2 matched in 2m 31s','[Driver: Mikhail K. 4.9★]','Fare: $11.20  ETA: 3:20','Saved $2.60 vs Fast Match'], color: C.bright },
    { n:'04', title:'Ops Analytics', desc:'Match rate per band, no-match churn rate, nudge upgrade rate. Ops team monitors band distribution to tune probability thresholds city by city.', mock: ['BidBand Performance Today','Band 1:  21%  2,140 reqs','Band 2:  46%  3,812 reqs','Band 3:  73%  5,940 reqs','Band 4:  92%  1,604 reqs'], color: C.green },
  ];

  screens.forEach((sc, i) => {
    const x = 0.35 + i * 2.35;
    sl.addShape(prs.ShapeType.roundRect, {
      x, y: 1.45, w: 2.1, h: 4.95,
      fill: { color: C.dark, transparency: 0 },
      line: { color: sc.color, width: 2 },
      rectRadius: 0.12,
      shadow: makeShadow(),
    });
    sl.addText(sc.n, { x, y: 1.55, w: 2.1, h: 0.3, fontSize: 9, color: sc.color, bold:true, align:'center', charSpacing:1 });
    sl.addText(sc.title, { x, y: 1.85, w: 2.1, h: 0.35, fontSize: 12, color: C.white, bold:true, align:'center' });
    sl.addText(sc.desc, { x: x+0.1, y: 2.25, w: 1.9, h: 1.1, fontSize: 9, color: C.muted, align:'center', lineSpacingMultiple:1.3 });
    // mock block
    sl.addShape(prs.ShapeType.roundRect, {
      x: x+0.12, y: 3.4, w: 1.86, h: 1.7,
      fill: { color: '1A1F1A', transparency: 0 },
      line: { color: sc.color, width: 0.6, transparency:50 },
      rectRadius: 0.07,
    });
    sc.mock.forEach((line, j) => {
      sl.addText(line, {
        x: x+0.17, y: 3.48 + j * 0.31, w: 1.76, h: 0.28,
        fontSize: 7.5, color: sc.color, fontFace:'Courier New', align:'center',
      });
    });
  });

  sl.addText('Prototype: indrive-bidband-prototype.html  —  interactive live demo of all 4 screen states', {
    x: 0.5, y: 6.7, w: 9.0, h: 0.3,
    fontSize: 9.5, color: C.accent, italic: true, align: 'center',
  });
}

// ─────────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('IMPACT & ROI', {
    x: 0.5, y: 0.4, w: 9, h: 0.3,
    fontSize: 9, color: C.bright, bold: true, charSpacing: 2,
  });
  sl.addText('Projected Impact — Built on PhonePe Proof Points', {
    x: 0.5, y: 0.75, w: 9, h: 0.5,
    fontSize: 24, color: C.white, bold: true,
  });

  // Left column: User Impact
  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.4, y: 1.45, w: 0.12, h: 3.5,
    fill: { color: C.bright }, line: { color: C.bright, width: 0 }, rectRadius: 0.06,
  });
  sl.addText('USER IMPACT', { x: 0.7, y: 1.45, w: 4.2, h: 0.3, fontSize: 9, color: C.bright, bold:true, charSpacing:1.5 });
  const leftM = [
    { val:'+28%', lbl:'match success rate', sub:'vs single-price baseline' },
    { val:'−55%', lbl:'post-no-match churn', sub:'riders retained at failure moment' },
    { val:'43%', lbl:'Band 1 → Band 3 via nudge', sub:'rides recovered that would be lost' },
    { val:'+18%', lbl:'GMV per session', sub:'informed users complete more rides' },
  ];
  leftM.forEach((m, i) => {
    sl.addText(m.val, { x: 0.7, y: 1.85 + i * 0.85, w: 2.2, h: 0.55, fontSize: 32, color: C.bright, bold:true, fontFace:'Arial' });
    sl.addText(m.lbl, { x: 2.9, y: 1.9 + i * 0.85, w: 2.0, h: 0.25, fontSize: 10.5, color: C.white, bold:true });
    sl.addText(m.sub, { x: 2.9, y: 2.17 + i * 0.85, w: 2.0, h: 0.25, fontSize: 9, color: C.muted });
  });

  // Right column: Platform ROI
  sl.addShape(prs.ShapeType.roundRect, {
    x: 5.3, y: 1.45, w: 0.12, h: 3.5,
    fill: { color: C.orange }, line: { color: C.orange, width: 0 }, rectRadius: 0.06,
  });
  sl.addText('PLATFORM ROI', { x: 5.6, y: 1.45, w: 4.0, h: 0.3, fontSize: 9, color: C.orange, bold:true, charSpacing:1.5 });
  const rightM = [
    { val:'Zero', lbl:'incremental subsidy cost', sub:'BidBand is UI + ML, no discount spend' },
    { val:'+X%', lbl:'ride completion rate', sub:'directly improves driver utilization' },
    { val:'High', lbl:'defensibility', sub:'Uber/Bolt cannot build this — no negotiation layer' },
    { val:'Day 1', lbl:'payback period', sub:'no CAPEX; API + frontend deployment only' },
  ];
  rightM.forEach((m, i) => {
    sl.addText(m.val, { x: 5.6, y: 1.85 + i * 0.85, w: 2.2, h: 0.55, fontSize: 32, color: C.orange, bold:true, fontFace:'Arial' });
    sl.addText(m.lbl, { x: 7.8, y: 1.9 + i * 0.85, w: 2.0, h: 0.25, fontSize: 10.5, color: C.white, bold:true });
    sl.addText(m.sub, { x: 7.8, y: 2.17 + i * 0.85, w: 2.0, h: 0.25, fontSize: 9, color: C.muted });
  });

  // Insight
  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.4, y: 5.55, w: 9.2, h: 0.75,
    fill: { color: '1A2A00', transparency: 0 },
    line: { color: C.bright, width: 1 },
    rectRadius: 0.08,
  });
  sl.addText('This is inDrive\'s only competitive moat deepened — not a cost center. Every competitor\'s surge model creates resentment. BidBand turns inDrive\'s negotiation into a transparency advantage that compounds with scale.', {
    x: 0.6, y: 5.65, w: 8.8, h: 0.55,
    fontSize: 10.5, color: C.white, lineSpacingMultiple: 1.3,
  });
}

// ─────────────────────────────────────────────
// SLIDE 7 — PROOF OF WORK
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.lgray };

  sl.addText('PROOF OF WORK', {
    x: 0.5, y: 0.4, w: 9, h: 0.3,
    fontSize: 9, color: C.accent, bold: true, charSpacing: 2,
  });
  sl.addText('I Built This Architecture. Here\'s The Evidence.', {
    x: 0.5, y: 0.75, w: 9, h: 0.5,
    fontSize: 24, color: C.dark, bold: true,
  });

  // Left: PhonePe
  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.4, y: 1.45, w: 4.4, h: 4.6,
    fill: { color: C.dark, transparency: 0 },
    line: { color: C.bright, width: 1 },
    rectRadius: 0.1,
    shadow: makeShadow(),
  });
  sl.addText('What I built at PhonePe', {
    x: 0.6, y: 1.6, w: 4.0, h: 0.35,
    fontSize: 12, color: C.bright, bold: true,
  });
  const leftPts = [
    'Real-time offer eligibility surface at 350M MAU — same signal-to-UI pipeline as BidBand probability rendering',
    'ML propensity scoring replacing ₹1,000+ Cr/yr static rules — identical band suppression architecture (show offer only when it changes behaviour)',
    'Dynamic cart incentivization on Pincode — detect abandonment signal → trigger contextual intervention; BidBand\'s nudge flow follows the exact same pattern',
    'A/B tested offer show vs suppress — 22% checkout lift; same experimentation framework for band threshold tuning',
    '500+ brand partners onboarded to self-serve campaign tooling — same ops analytics thinking as BidBand band performance dashboard',
  ];
  leftPts.forEach((pt, i) => {
    sl.addText(`• ${pt}`, {
      x: 0.6, y: 2.05 + i * 0.72, w: 4.0, h: 0.65,
      fontSize: 9.5, color: C.muted, lineSpacingMultiple: 1.3,
    });
  });

  // Right: JD mapping
  sl.addShape(prs.ShapeType.roundRect, {
    x: 5.2, y: 1.45, w: 4.4, h: 4.6,
    fill: { color: 'F8F9FA', transparency: 0 },
    line: { color: C.accent, width: 1 },
    rectRadius: 0.1,
    shadow: makeShadow(),
  });
  sl.addText('Maps to inDrive JD requirements', {
    x: 5.4, y: 1.6, w: 4.0, h: 0.35,
    fontSize: 12, color: C.accent, bold: true,
  });
  const rightPts = [
    '"Translate customer behavior insights into segmented programmable journeys" → BidBand segments on intent signal (route + time + supply) before the bid',
    '"Partner with ML/DS to test hybrid AI stacks" → BidBand probability model: scoring + real-time retrieval; PhonePe shipped this in production',
    '"Own roadmap for engagement and monetization loops" → BidBand closes the activation loop (submit → match → next ride) that currently breaks at no-match',
    '"Drive growth through experimentation" → Band threshold A/B testing framework; same as PhonePe offer eligibility experiments',
    '"Scale interventions across engagement loops" → Nudge engine recovers 43% of failures; same trigger-to-action architecture as Pincode cart recovery',
  ];
  rightPts.forEach((pt, i) => {
    sl.addText(`• ${pt}`, {
      x: 5.4, y: 2.05 + i * 0.72, w: 4.0, h: 0.65,
      fontSize: 9.5, color: '2D3748', lineSpacingMultiple: 1.3,
    });
  });

  // Quote
  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.4, y: 6.15, w: 9.2, h: 0.65,
    fill: { color: C.dark, transparency: 0 },
    line: { color: C.bright, width: 1 },
    rectRadius: 0.07,
  });
  sl.addText('"The propensity scoring, the UI suppression, the progressive nudge — this is not a whiteboard idea. I shipped these primitives at 350M MAU."', {
    x: 0.6, y: 6.23, w: 8.8, h: 0.5,
    fontSize: 10.5, color: C.white, italic: true, lineSpacingMultiple: 1.3,
  });
}

// ─────────────────────────────────────────────
// SLIDE 8 — ROLLOUT PLAN
// ─────────────────────────────────────────────
{
  const sl = prs.addSlide();
  sl.background = { color: C.dark };

  sl.addText('ROLLOUT PLAN', {
    x: 0.5, y: 0.4, w: 9, h: 0.3,
    fontSize: 9, color: C.bright, bold: true, charSpacing: 2,
  });
  sl.addText('Phased Launch — 2 Cities, 90 Days to Full Proof', {
    x: 0.5, y: 0.75, w: 9, h: 0.5,
    fontSize: 24, color: C.white, bold: true,
  });

  const phases = [
    { ph:'Phase 1', time:'Month 1–2', title:'Pilot — 2 Cities', bullets:['Deploy BidBand to 100% of new ride requests in 2 cities','A/B test: BidBand flow vs control (single price)','Instrument: match rate, time-to-match, churn at no-match, band distribution','Tune probability thresholds using live acceptance data'], color: C.bright },
    { ph:'Phase 2', time:'Month 3–4', title:'Nudge Engine + Analytics', bullets:['Launch progressive nudge flow (Band N → N+1 at 3 min threshold)','Build ops analytics dashboard: band performance by city + time','Test nudge timing (2 min vs 3 min vs 5 min)','Expand to 5 cities based on pilot learnings'], color: C.orange },
    { ph:'Phase 3', time:'Month 5–6', title:'Full Rollout + Drivers', bullets:['Roll out to all markets globally','Adapt probability bands per market (pricing dynamics vary by city tier)','Launch driver-facing BidBand: accept counter-band offers in one tap','Integration: BidBand probability feeds into driver dispatching model'], color: C.green },
  ];

  phases.forEach((ph, i) => {
    const x = 0.4 + i * 3.2;
    sl.addShape(prs.ShapeType.roundRect, {
      x, y: 1.45, w: 2.9, h: 3.9,
      fill: { color: '141810', transparency: 0 },
      line: { color: ph.color, width: 1.5 },
      rectRadius: 0.1,
      shadow: makeShadow(),
    });
    sl.addText(ph.ph, { x, y: 1.55, w: 2.9, h: 0.3, fontSize: 9, color: ph.color, bold:true, align:'center', charSpacing:1 });
    sl.addText(ph.time, { x, y: 1.85, w: 2.9, h: 0.25, fontSize: 9, color: C.muted, align:'center' });
    sl.addText(ph.title, { x: x+0.12, y: 2.15, w: 2.66, h: 0.35, fontSize: 12, color: C.white, bold:true, align:'center' });
    ph.bullets.forEach((b, j) => {
      sl.addText(`• ${b}`, { x: x+0.15, y: 2.6 + j * 0.62, w: 2.6, h: 0.55, fontSize: 9, color: C.muted, lineSpacingMultiple:1.3 });
    });
  });

  // Ask box
  sl.addShape(prs.ShapeType.roundRect, {
    x: 0.4, y: 5.55, w: 9.2, h: 1.1,
    fill: { color: '0D1500', transparency: 0 },
    line: { color: C.bright, width: 1.2 },
    rectRadius: 0.1,
  });
  sl.addText('What I need to build this:', {
    x: 0.6, y: 5.65, w: 9.0, h: 0.28,
    fontSize: 11, color: C.bright, bold: true,
  });
  sl.addText('Access to ride acceptance rate data by corridor + time  ·  DS partner for probability model calibration  ·  1 frontend eng for BidBand UI  ·  City ops alignment on threshold tuning', {
    x: 0.6, y: 5.95, w: 8.8, h: 0.55,
    fontSize: 10, color: C.white, lineSpacingMultiple: 1.35,
  });

  // Contact
  sl.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  ajay-avaghade.github.io/Portfolio', {
    x: 0.5, y: 7.1, w: 9.0, h: 0.25,
    fontSize: 9, color: C.muted, align: 'center',
  });
}

prs.writeFile({ fileName: 'assets/indrive-bidband-deck.pptx' })
  .then(() => console.log('indrive-bidband-deck.pptx created'))
  .catch(e => console.error(e));
