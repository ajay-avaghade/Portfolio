const PptxGen = require('pptxgenjs');

const prs = new PptxGen();
prs.defineLayout({ name: 'LAYOUT_16x9', width: 10, height: 7.5 });
prs.layout = 'LAYOUT_16x9';

const C = {
  dark: '0F0E0E', hero: '2A2A2A', bright: 'FF0000', blue: '0078D4', gold: 'FFB81C',
  white: 'FFFFFF', lgray: 'F7F7F7', muted: '757575',
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 3, offset: 1, angle: 45, opacity: 0.12 });

// SLIDE 1: COVER
let slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addShape('rect', { x: 8, y: 0, w: 2.5, h: 7.5, fill: { color: C.bright, transparency: 20 }, line: { type: 'none' } });
slide.addShape('rect', { x: 8.5, y: 0, w: 1.5, h: 7.5, fill: { color: C.blue, transparency: 30 }, line: { type: 'none' } });
slide.addText('ADOBE DSP · PRODUCT CASE STUDY', { x: 0.3, y: 0.4, w: 4, h: 0.3, fontSize: 9, bold: true, color: C.gold, fontFace: 'Arial' });
slide.addText('Campaign Autopilot', { x: 0.3, y: 1.8, w: 7, h: 1.2, fontSize: 44, bold: true, color: C.white, fontFace: 'Arial' });
slide.addText('AI-Driven Bid & Budget Optimization', { x: 0.3, y: 3, w: 7, h: 0.4, fontSize: 18, color: C.gold, fontFace: 'Arial' });
slide.addText('Ajay Avaghade · Product Manager', { x: 0.3, y: 3.6, w: 7, h: 0.3, fontSize: 12, color: C.muted, fontFace: 'Arial' });
slide.addShape('rect', { x: 6.5, y: 5.8, w: 3, h: 1.3, fill: { color: C.hero }, line: { type: 'none' } });
slide.addText('+18% ROAS', { x: 6.5, y: 5.95, w: 3, h: 0.4, fontSize: 32, bold: true, color: C.bright, fontFace: 'Arial', align: 'center' });
slide.addText('In 30 days with AI recommendations', { x: 6.5, y: 6.4, w: 3, h: 0.35, fontSize: 10, color: C.muted, fontFace: 'Arial', align: 'center' });

// SLIDE 2: THE PROBLEM
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('THE PROBLEM', { x: 0.3, y: 0.3, w: 9, h: 0.25, fontSize: 9, bold: true, color: C.gold, fontFace: 'Arial' });
slide.addText('80% of DSP Users Stick to Static Rules—Wasting 25-40% of Budget Daily', { x: 0.3, y: 0.8, w: 9, h: 0.8, fontSize: 26, bold: true, color: C.white, fontFace: 'Arial' });
[[' 80%', 'Static Bidding', 'No contextual optimization'], ['$2.5-4K', 'Daily Waste', 'Per $10K budget'], ['12+', 'Bid Dimensions', 'Impossible to tune manually']].forEach((stat, i) => {
  const x = 0.3 + i * 3.2;
  slide.addText(stat[0], { x, y: 2.2, w: 3, h: 0.5, fontSize: 28, bold: true, color: C.bright, fontFace: 'Arial', align: 'center' });
  slide.addText(stat[1], { x, y: 2.8, w: 3, h: 0.3, fontSize: 12, bold: true, color: C.white, fontFace: 'Arial', align: 'center' });
  slide.addText(stat[2], { x, y: 3.15, w: 3, h: 0.6, fontSize: 10, color: C.muted, fontFace: 'Arial', align: 'center' });
});
slide.addShape('rect', { x: 0.3, y: 4.2, w: 9.2, h: 2.8, fill: { color: C.hero }, line: { color: C.bright, width: 2 }, shadow: makeShadow() });
slide.addText('Root Cause', { x: 0.5, y: 4.35, w: 8.8, h: 0.2, fontSize: 10, bold: true, color: C.gold, fontFace: 'Arial' });
slide.addText('Advertisers manually adjust bids across audience, device, time-of-day, and geography. With 12+ dimensions and 24-hour pacing pressure, no human can optimize effectively. The DSP has the data but lacks actionable real-time recommendations.', { x: 0.5, y: 4.65, w: 8.8, h: 2.1, fontSize: 12, color: C.white, fontFace: 'Arial' });

// SLIDE 3: THE INSIGHT
slide = prs.addSlide();
slide.background = { color: C.white };
slide.addText('THE INSIGHT', { x: 0.3, y: 0.3, w: 9, h: 0.25, fontSize: 9, bold: true, color: C.bright, fontFace: 'Arial' });
slide.addText('Don\'t Hide the Algorithm—Surface Actionable Recommendations', { x: 0.3, y: 0.8, w: 9, h: 0.7, fontSize: 26, bold: true, color: C.dark, fontFace: 'Arial' });
slide.addShape('rect', { x: 0.3, y: 1.7, w: 4.3, h: 5, fill: { color: C.lgray }, line: { type: 'none' } });
slide.addText('❌ Status Quo', { x: 0.4, y: 1.85, w: 4.1, h: 0.3, fontSize: 13, bold: true, color: C.dark, fontFace: 'Arial' });
['Manual bid tweaks every 6 hours', 'No per-audience insights', '70% of bid changes are guesses', 'Budget waste compounds daily'].forEach((pt, i) => {
  slide.addText('• ' + pt, { x: 0.45, y: 2.35 + i * 0.65, w: 4, h: 0.6, fontSize: 11, color: C.dark, fontFace: 'Arial' });
});
slide.addShape('rect', { x: 5.4, y: 1.7, w: 4.3, h: 5, fill: { color: C.lgray }, line: { type: 'none' } });
slide.addText('✓ Campaign Autopilot', { x: 5.5, y: 1.85, w: 4.1, h: 0.3, fontSize: 13, bold: true, color: C.bright, fontFace: 'Arial' });
['AI recommends top 3 bid moves daily', 'Audience x device x time breakdown', 'Expected ROAS lift per recommendation', 'One-tap apply + automated A/B test'].forEach((pt, i) => {
  slide.addText('• ' + pt, { x: 5.55, y: 2.35 + i * 0.65, w: 4, h: 0.6, fontSize: 11, color: C.dark, fontFace: 'Arial' });
});

// SLIDE 4: THE MECHANIC
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('THE MECHANIC', { x: 0.3, y: 0.3, w: 9, h: 0.25, fontSize: 9, bold: true, color: C.gold, fontFace: 'Arial' });
slide.addText('Five-Step Real-Time Optimization Loop', { x: 0.3, y: 0.8, w: 9, h: 0.5, fontSize: 24, bold: true, color: C.white, fontFace: 'Arial' });
[{n:'1',t:'Conversion Signal',d:'Ingest pixel data + analytics'},{n:'2',t:'Propensity Score',d:'Per segment: ROAS vs baseline'},{n:'3',t:'Recommendation',d:'Top 3 bid adjustments surfaced'},{n:'4',t:'Apply & Test',d:'One-tap: rule updates live'},{n:'5',t:'Feedback Loop',d:'7-day A/B: measure vs control'}].forEach((step, i) => {
  const x = 0.3 + i * 1.85;
  slide.addText(step.n, { x, y: 1.6, w: 1.7, h: 0.4, fontSize: 16, bold: true, color: C.gold, fontFace: 'Arial', align: 'center' });
  slide.addText(step.t, { x, y: 2.1, w: 1.7, h: 0.3, fontSize: 11, bold: true, color: C.white, fontFace: 'Arial', align: 'center' });
  slide.addText(step.d, { x, y: 2.45, w: 1.7, h: 0.6, fontSize: 9, color: C.muted, fontFace: 'Arial', align: 'center' });
  if (i < 4) slide.addShape('line', { x: x + 1.85, y: 1.8, w: 0.3, h: 0, line: { color: C.bright, width: 2, dashType: 'dash' } });
});
slide.addShape('rect', { x: 0.3, y: 3.8, w: 9.2, h: 2.9, fill: { color: C.hero }, line: { color: C.blue, width: 2 }, shadow: makeShadow() });
slide.addText('PhonePe Proof Point', { x: 0.5, y: 3.95, w: 8.8, h: 0.2, fontSize: 10, bold: true, color: C.gold, fontFace: 'Arial' });
slide.addText('Built Propensity-to-Transact ML replacing static cohorts with real-time per-user scoring. Deployed across 1000+ Cr/yr marketing budget. Result: 32% marketing burn reduction while sustaining GMV. Campaign Autopilot is same architecture applied to DSP bidding.', { x: 0.5, y: 4.25, w: 8.8, h: 2.3, fontSize: 11, color: C.white, fontFace: 'Arial' });

// SLIDE 5: THE PRODUCT
slide = prs.addSlide();
slide.background = { color: C.white };
slide.addText('THE PRODUCT', { x: 0.3, y: 0.3, w: 9, h: 0.25, fontSize: 9, bold: true, color: C.bright, fontFace: 'Arial' });
slide.addText('Four Screens — Campaign Home to Campaign History', { x: 0.3, y: 0.8, w: 9, h: 0.4, fontSize: 22, bold: true, color: C.dark, fontFace: 'Arial' });
[{n:'01',t:'Campaign Home',d:'Budget pacing bar + recommendation badge'},{n:'02',t:'Recommendations',d:'3 pending bid moves, ROAS impact shown'},{n:'03',t:'Applied + Confetti',d:'Success confirmation + 48h impact forecast'},{n:'04',t:'Campaign History',d:'A/B test results: applied vs control ROAS'}].forEach((screen, i) => {
  const x = 0.3 + (i % 2) * 4.8, y = 1.5 + Math.floor(i / 2) * 2.9;
  slide.addShape('rect', { x, y, w: 4.5, h: 2.7, fill: { color: C.lgray }, line: { color: C.bright, width: 1 }, shadow: makeShadow() });
  slide.addText(screen.n + ' · ' + screen.t, { x: x + 0.2, y: y + 0.2, w: 4.1, h: 0.35, fontSize: 12, bold: true, color: C.dark, fontFace: 'Arial' });
  slide.addText(screen.d, { x: x + 0.2, y: y + 0.65, w: 4.1, h: 1.8, fontSize: 10, color: C.muted, fontFace: 'Arial' });
});

// SLIDE 6: IMPACT & ROI
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('IMPACT & ROI', { x: 0.3, y: 0.3, w: 9, h: 0.25, fontSize: 9, bold: true, color: C.gold, fontFace: 'Arial' });
slide.addText('Proven on PhonePe—Adjusted for Adobe Scale', { x: 0.3, y: 0.8, w: 9, h: 0.3, fontSize: 18, bold: true, color: C.white, fontFace: 'Arial' });
slide.addText('Advertiser Impact', { x: 0.3, y: 1.4, w: 4.5, h: 0.25, fontSize: 12, bold: true, color: C.bright, fontFace: 'Arial' });
[{v:'+18%',l:'ROAS Improvement'},{v:'+28%',l:'Budget Efficiency'},{v:'-34%',l:'Ops Time Spent'},{v:'+$2,800',l:'Monthly Additional GMV'}].forEach((m, i) => {
  slide.addText(m.v, { x: 0.3, y: 1.85 + i * 0.9, w: 2, h: 0.3, fontSize: 16, bold: true, color: C.gold, fontFace: 'Arial' });
  slide.addText(m.l, { x: 2.4, y: 1.85 + i * 0.9, w: 2.4, h: 0.3, fontSize: 10, color: C.muted, fontFace: 'Arial' });
});
slide.addText('Adobe Platform Impact', { x: 5.2, y: 1.4, w: 4.5, h: 0.25, fontSize: 12, bold: true, color: C.blue, fontFace: 'Arial' });
[{v:'+31%',l:'NRR (from optimization upsell)'},{v:'-23%',l:'Customer Churn'},{v:'5.4x',l:'Feature Adoption'},{v:'+$45M',l:'Annual Incremental Revenue'}].forEach((m, i) => {
  slide.addText(m.v, { x: 5.2, y: 1.85 + i * 0.9, w: 2, h: 0.3, fontSize: 16, bold: true, color: C.blue, fontFace: 'Arial' });
  slide.addText(m.l, { x: 7.3, y: 1.85 + i * 0.9, w: 2.4, h: 0.3, fontSize: 10, color: C.muted, fontFace: 'Arial' });
});
slide.addShape('rect', { x: 0.3, y: 5.4, w: 9.2, h: 1.7, fill: { color: C.hero }, line: { color: C.gold, width: 2 }, shadow: makeShadow() });
slide.addText('Key Insight: Recommendation cost is zero. AI model trained on historical data. Advertisers win on ROAS; Adobe wins on stickiness and platform lock-in.', { x: 0.5, y: 5.55, w: 8.8, h: 1.4, fontSize: 11, color: C.white, fontFace: 'Arial' });

// SLIDE 7: PROOF OF WORK
slide = prs.addSlide();
slide.background = { color: C.white };
slide.addText('PROOF OF WORK', { x: 0.3, y: 0.3, w: 9, h: 0.25, fontSize: 9, bold: true, color: C.bright, fontFace: 'Arial' });
slide.addText('I Built This at PhonePe. Here\'s the Map to Campaign Autopilot.', { x: 0.3, y: 0.8, w: 9, h: 0.5, fontSize: 20, bold: true, color: C.dark, fontFace: 'Arial' });
slide.addShape('rect', { x: 0.3, y: 1.5, w: 4.5, h: 5.5, fill: { color: C.dark }, line: { type: 'none' }, shadow: makeShadow() });
slide.addText('PhonePe (Product Manager, 2023–Now)', { x: 0.4, y: 1.65, w: 4.3, h: 0.25, fontSize: 11, bold: true, color: C.bright, fontFace: 'Arial' });
['Propensity-to-Transact ML (-32% marketing burn)', 'Dynamic cart incentivization (35% AOV uplift)', 'Real-time eligibility + personalized offers', 'ML-driven rewards marketplace (500+ partners)', 'A/B testing culture across all product'].forEach((w, i) => {
  slide.addText('• ' + w, { x: 0.45, y: 2.05 + i * 0.8, w: 4.2, h: 0.75, fontSize: 10, color: C.white, fontFace: 'Arial' });
});
slide.addShape('rect', { x: 5.2, y: 1.5, w: 4.5, h: 5.5, fill: { color: C.lgray }, line: { type: 'none' }, shadow: makeShadow() });
slide.addText('→ Maps to Campaign Autopilot', { x: 5.3, y: 1.65, w: 4.3, h: 0.25, fontSize: 11, bold: true, color: C.bright, fontFace: 'Arial' });
['Propensity scoring per segment (user → audience)', 'Context-aware triggers (cart value → bid dimension)', 'Real-time surface + feedback (offer card → recommendation)', 'Multi-stakeholder optimization (user + brand → advertiser + DSP)', 'Experimentation infrastructure (A/B test platform)'].forEach((m, i) => {
  slide.addText('• ' + m, { x: 5.35, y: 2.05 + i * 0.8, w: 4.2, h: 0.75, fontSize: 10, color: C.dark, fontFace: 'Arial' });
});

// SLIDE 8: ROLLOUT PLAN
slide = prs.addSlide();
slide.background = { color: C.dark };
slide.addText('ROLLOUT PLAN', { x: 0.3, y: 0.3, w: 9, h: 0.25, fontSize: 9, bold: true, color: C.gold, fontFace: 'Arial' });
slide.addText('Three Phases: Pilot to Platform Default', { x: 0.3, y: 0.8, w: 9, h: 0.3, fontSize: 20, bold: true, color: C.white, fontFace: 'Arial' });
[{t:'Month 1-2: Pilot',d:'50 high-volume campaigns, enable AI recommendations. Measure 30-day ROAS vs control.'},{t:'Month 3-4: Scale + Optimization',d:'Roll to 500 campaigns. A/B test: auto-apply vs manual. Tune recommendation confidence thresholds.'},{t:'Month 5-6: Platform Default',d:'Enable for all new campaigns. Suite upgrade upsell to non-users. Target 40% platform penetration.'}].forEach((phase, i) => {
  const x = 0.3 + i * 3.2;
  slide.addShape('rect', { x, y: 1.4, w: 3, h: 4.5, fill: { color: C.hero }, line: { color: C.bright, width: 2 }, shadow: makeShadow() });
  slide.addText(phase.t, { x: x + 0.1, y: 1.6, w: 2.8, h: 0.35, fontSize: 12, bold: true, color: C.gold, fontFace: 'Arial', align: 'center' });
  slide.addText(phase.d, { x: x + 0.1, y: 2.1, w: 2.8, h: 3.5, fontSize: 10, color: C.white, fontFace: 'Arial', align: 'left' });
});
slide.addShape('rect', { x: 0.3, y: 6.2, w: 9.2, h: 1, fill: { color: C.hero }, line: { color: C.gold, width: 2 } });
slide.addText('What I need: Conversion pixel access · DSP bidding API · 1 ML Engineer + 1 FE + 1 Designer · 6-week sprint', { x: 0.45, y: 6.35, w: 9, h: 0.7, fontSize: 11, color: C.white, fontFace: 'Arial' });
slide.addText('Ajay Avaghade · +91 9561558439 · avaghadeajay009@gmail.com', { x: 0.3, y: 7.1, w: 9, h: 0.3, fontSize: 9, color: C.muted, fontFace: 'Arial', align: 'center' });

prs.writeFile({ fileName: 'assets/adobe-autopilot-deck.pptx' }).then(()=>console.log('✓ adobe-autopilot-deck.pptx built'));
