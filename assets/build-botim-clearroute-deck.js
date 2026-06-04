(async () => {
const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.defineLayout({ name:'LAYOUT_16x9', width:10, height:7.5 });
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:  '0A1628', hero:  '0F1E3A',
  teal:  '00BFA5', blue:  '0668E1',
  orange:'FF6B35', white: 'FFFFFF',
  lgray: 'F0F4F8', muted: '8b92a5',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });

// S1 COVER
const s1 = prs.addSlide(); s1.background = { color: C.dark };
s1.addText('BOTIM FOR BUSINESS · CROSS-BORDER PAYMENTS', { x:.3, y:.4, w:9.4, h:.3, fontSize:9, bold:true, color:C.muted, align:'center' });
s1.addText('ClearRoute', { x:.5, y:1.2, w:9, h:1.5, fontSize:52, bold:true, color:C.white, align:'center' });
s1.addText('Intelligent Cross-Border Rail Selector — SWIFT vs Local Rails', { x:.5, y:2.8, w:9, h:.8, fontSize:20, color:C.teal, align:'center' });
s1.addText('Ajay Avaghade · Product Manager', { x:.5, y:3.8, w:9, h:.4, fontSize:12, color:C.muted, align:'center' });
s1.addShape('rect', { x:7.5, y:5.2, w:2, h:.1, fill:{ color:C.teal }, line:{ type:'none' } });
s1.addText('AED 180 saved per transaction', { x:7.2, y:5.4, w:2.6, h:.8, fontSize:16, bold:true, color:C.teal, align:'right' });

// S2 PROBLEM
const s2 = prs.addSlide(); s2.background = { color: C.dark };
s2.addText('THE PROBLEM', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s2.addText('SMEs Default to SWIFT for All International Payments — Overpaying 40–60% on Routes With Cheaper Local Rail Alternatives', { x:.5, y:.8, w:9, h:1.3, fontSize:26, bold:true, color:C.white });
[{ num:'AED 225', label:'Avg SWIFT cost', sub:'flat fee + FX spread per transaction' },
 { num:'2–3 days', label:'SWIFT settlement TAT', sub:'vs 2–4 hours via local rail (IMPS/Raast)' },
 { num:'40–60%', label:'overpayment rate', sub:'on corridors with cheaper local rail options' }
].forEach((s, i) => {
  const x = .5 + i * 3.1;
  s2.addText(s.num, { x, y:2.3, w:2.8, h:.6, fontSize:30, bold:true, color:C.teal, align:'center' });
  s2.addText(s.label, { x, y:3.0, w:2.8, h:.4, fontSize:12, bold:true, color:C.white, align:'center' });
  s2.addText(s.sub, { x, y:3.5, w:2.8, h:.8, fontSize:10, color:C.muted, align:'center' });
});
s2.addShape('rect', { x:.5, y:5.2, w:9, h:1.7, fill:{ color:C.hero }, line:{ color:C.teal, width:1 } });
s2.addText('Root cause: No intelligence layer exists between the SME and the payment rail. SWIFT is the default because it is the only option the business knows. Local rails (India IMPS, Pakistan Raast, Philippines InstaPay, UAE AANI) are cheaper and faster — but invisible to the user.', { x:.7, y:5.35, w:8.6, h:1.4, fontSize:10, color:C.white });

// S3 INSIGHT
const s3 = prs.addSlide(); s3.background = { color: C.lgray };
s3.addText('THE INSIGHT', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s3.addText('Rail Selection Is a Product Problem, Not a Treasury Problem', { x:.5, y:.8, w:9, h:.9, fontSize:28, bold:true, color:C.dark });
s3.addShape('rect', { x:.5, y:2.0, w:4.2, h:3.8, fill:{ color:'FFFFFF' }, line:{ color:C.muted, width:1 } });
s3.addText('❌ Today: SWIFT Default', { x:.7, y:2.2, w:3.8, h:.4, fontSize:12, bold:true, color:C.orange });
['SME initiates AED→INR payment','Default route: SWIFT','Fee: AED 225 flat + spread','Settlement: 2–3 business days','User has no visibility into alternatives'].forEach((t,i) => s3.addText(t, { x:.9, y:2.8+i*.58, w:3.6, h:.5, fontSize:10, color:C.dark }));
s3.addShape('rect', { x:5.3, y:2.0, w:4.2, h:3.8, fill:{ color:C.hero }, line:{ color:C.teal, width:1 } });
s3.addText('✅ ClearRoute: Intelligent', { x:5.5, y:2.2, w:3.8, h:.4, fontSize:12, bold:true, color:C.teal });
['ML scores SWIFT vs IMPS vs Correspondent','Recommends IMPS: AED 45, 2–4 hrs','Side-by-side comparison in-app','One-tap confirm','"You saved AED 180" moment at success'].forEach((t,i) => s3.addText(t, { x:5.7, y:2.8+i*.58, w:3.6, h:.5, fontSize:10, color:C.white }));
s3.addShape('ellipse', { x:4.6, y:3.5, w:.8, h:.8, fill:{ color:C.dark }, line:{ color:C.teal, width:2 } });
s3.addText('VS', { x:4.6, y:3.55, w:.8, h:.7, fontSize:14, bold:true, color:C.teal, align:'center' });

// S4 MECHANIC
const s4 = prs.addSlide(); s4.background = { color: C.dark };
s4.addText('THE MECHANIC', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s4.addText('5-Signal ML Model → Optimal Rail → Transparent Choice → Confirm', { x:.5, y:.8, w:9, h:.5, fontSize:22, bold:true, color:C.white });
[{ n:'1', t:'Signal Capture', d:'Corridor + amount + urgency + currency + business type captured' },
 { n:'2', t:'Rail Scoring', d:'ML scores SWIFT / local rail / correspondent on cost, TAT, reliability' },
 { n:'3', t:'Recommendation', d:'Best route shown with cost, TAT, success rate. Side-by-side comparison' },
 { n:'4', t:'One-Tap Confirm', d:'Business confirms. FX rate locked. Payment initiated on optimal rail' },
 { n:'5', t:'Feedback Loop', d:'Success/failure data retrains model. Rail health alerts trigger automatic rerouting' },
].forEach((s, i) => {
  const y = 1.6 + i * .85;
  s4.addShape('ellipse', { x:.5, y, w:.4, h:.4, fill:{ color:C.teal }, line:{ type:'none' } });
  s4.addText(s.n, { x:.5, y, w:.4, h:.4, fontSize:14, bold:true, color:C.dark, align:'center' });
  s4.addText(s.t, { x:1.1, y, w:2.3, h:.22, fontSize:10, bold:true, color:C.white });
  s4.addText(s.d, { x:1.1, y:y+.25, w:2.3, h:.28, fontSize:8, color:C.muted });
  if (i < 4) s4.addShape('line', { x:.7, y:y+.45, w:0, h:.25, line:{ color:C.teal, width:2, dashType:'dash' } });
});
s4.addShape('rect', { x:3.8, y:1.6, w:5.7, h:4.3, fill:{ color:C.hero }, line:{ color:C.teal, width:1 } });
s4.addText('Proof: PhonePe Propensity-to-Transact ML (350M MAU)', { x:4.0, y:1.8, w:5.3, h:.4, fontSize:10, bold:true, color:C.teal });
['ML replaced manual rule-based cohort targeting at PhonePe','Real-time user-level scoring → optimal offer selected per user','Result: 32% marketing burn reduction, sustained GMV growth','ClearRoute: same ML-for-routing architecture, different domain','Signal inputs differ; model architecture and business logic identical'].forEach((t,i) => s4.addText(t, { x:4.2, y:2.4+i*.62, w:5.1, h:.56, fontSize:9, color:C.white }));

// S5 PRODUCT
const s5 = prs.addSlide(); s5.background = { color: C.lgray };
s5.addText('THE PRODUCT', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s5.addText('4 Screens — Initiation to Ops Intelligence', { x:.5, y:.8, w:9, h:.5, fontSize:24, bold:true, color:C.dark });
[{ n:'01', t:'Payment Initiation', d:'ClearRoute recommendation card. Best route (IMPS) vs default SWIFT. Cost, TAT, success rate side-by-side. AED 180 saving highlighted.' },
 { n:'02', t:'Route Comparison', d:'Full 3-route comparison table: CBUAE IMPS / SWIFT / Correspondent. Fee, FX rate, settlement time, success rate, partner bank.' },
 { n:'03', t:'Payment Confirmed', d:'Confetti success state. Amount sent, FX rate locked, expected arrival. Saving highlighted: "You saved AED 180 vs default SWIFT."' },
 { n:'04', t:'Ops Analytics', d:'Routing console: avg saving/txn, success rates, SWIFT→local rail shift %, corridor-level rail usage, real-time alert on rail degradation.' },
].forEach((s, i) => {
  const y = 1.6 + i * 1.35;
  s5.addShape('rect', { x:.5, y, w:9, h:1.2, fill:{ color:'FFFFFF' }, line:{ color:C.muted, width:1 } });
  s5.addText(`${s.n} — ${s.t}`, { x:.7, y:y+.15, w:3.5, h:.3, fontSize:11, bold:true, color:C.dark });
  s5.addText(s.d, { x:.7, y:y+.5, w:8.6, h:.55, fontSize:9, color:C.dark });
});

// S6 ROI
const s6 = prs.addSlide(); s6.background = { color: C.dark };
s6.addText('IMPACT & ROI', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s6.addText('Cost Savings, Speed & Volume Growth — Built on PhonePe ML Proof Points', { x:.5, y:.8, w:9, h:.5, fontSize:22, bold:true, color:C.white });
[{ ico:'💰', v:'AED 180', l:'Avg saving per txn', s:'vs default SWIFT (40–60% cost reduction)' },
 { ico:'⚡', v:'3.2 hrs', l:'Avg settlement TAT', s:'vs 2–3 business days SWIFT (10× faster)' },
 { ico:'📈', v:'72%', l:'Local rail shift', s:'of SWIFT transactions migrated to local rails' },
 { ico:'✅', v:'98.6%', l:'Route success rate', s:'with automatic rerouting on rail degradation' },
].forEach((m, i) => {
  const x = .5 + (i%2)*4.6, y = 1.6 + Math.floor(i/2)*2;
  s6.addShape('rect', { x, y, w:4, h:1.8, fill:{ color:C.hero }, line:{ color:C.teal, width:1 } });
  s6.addText(m.ico, { x:x+.2, y:y+.1, w:3.6, h:.4, fontSize:22, align:'left' });
  s6.addText(m.v, { x:x+.2, y:y+.52, w:3.6, h:.38, fontSize:20, bold:true, color:C.teal, align:'left' });
  s6.addText(m.l, { x:x+.2, y:y+.96, w:3.6, h:.28, fontSize:10, bold:true, color:C.white, align:'left' });
  s6.addText(m.s, { x:x+.2, y:y+1.28, w:3.6, h:.38, fontSize:8, color:C.muted, align:'left' });
});

// S7 PROOF
const s7 = prs.addSlide(); s7.background = { color: C.lgray };
s7.addText('PROOF OF WORK', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s7.addText('PhonePe Propensity ML: Same Architecture, Different Domain', { x:.5, y:.8, w:9, h:.6, fontSize:24, bold:true, color:C.dark });
s7.addShape('rect', { x:.5, y:1.6, w:4.3, h:4.4, fill:{ color:C.dark }, line:{ color:C.teal, width:1 } });
s7.addText('PhonePe: What I Built', { x:.7, y:1.8, w:3.9, h:.3, fontSize:11, bold:true, color:C.teal });
['Inherited rule-based offer targeting across 350M users','Replaced with Propensity-to-Transact ML model','Real-time user scoring: signal capture → model → optimal action','Result: 32% marketing burn reduction, sustained GMV','Managed ₹1,000Cr+ annual marketing budget optimisation'].forEach((t,i) => s7.addText(t, { x:.9, y:2.3+i*.58, w:3.7, h:.52, fontSize:9, color:C.white }));
s7.addShape('rect', { x:5.2, y:1.6, w:4.3, h:4.4, fill:{ color:'FFFFFF' }, line:{ color:C.teal, width:1 } });
s7.addText('ClearRoute: Application', { x:5.4, y:1.8, w:3.9, h:.3, fontSize:11, bold:true, color:C.teal });
['Input: corridor + amount + urgency + currency','ML scores 3 rails: cost, TAT, reliability','Output: optimal rail recommendation + comparison','User saves AED 180 per transaction on average','Model self-improves from success/failure feedback'].forEach((t,i) => s7.addText(t, { x:5.4, y:2.3+i*.58, w:3.9, h:.52, fontSize:9, color:C.dark }));

// S8 ROLLOUT
const s8 = prs.addSlide(); s8.background = { color: C.dark };
s8.addText('ROLLOUT PLAN', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s8.addText('Three Phases — India/Pakistan Pilot to Full Multi-Corridor Coverage', { x:.5, y:.8, w:9, h:.5, fontSize:22, bold:true, color:C.white });
[{ t:'Phase 1 — Pilot (Month 1–2)', b:'Deploy ClearRoute on India (IMPS) + Pakistan (Raast). A/B: ClearRoute vs control (SWIFT default). Measure: saving per txn, TAT reduction, rail success rate.' },
 { t:'Phase 2 — Corridors + Rerouting (Month 3–4)', b:'Add Philippines (InstaPay), UK (FPS). Auto-rerouting logic on rail degradation. Ops console live for payment ops team.' },
 { t:'Phase 3 — Full Scale (Month 5–6)', b:'All live corridors under ClearRoute. Model retraining loop active. Expand to Africa (M-Pesa), Europe (SEPA). Cross-sell FX rate lock (FXGuard).' },
].forEach((p, i) => {
  const y = 1.5 + i * 1.75;
  s8.addShape('rect', { x:.5, y, w:9, h:1.6, fill:{ color:C.hero }, line:{ color:C.teal, width:1 } });
  s8.addText(p.t, { x:.7, y:y+.15, w:8.6, h:.3, fontSize:11, bold:true, color:C.teal });
  s8.addText(p.b, { x:.9, y:y+.52, w:8.2, h:.85, fontSize:9, color:C.white });
});
s8.addShape('rect', { x:.5, y:6.5, w:9, h:.8, fill:{ color:C.hero }, line:{ color:C.orange, width:2 } });
s8.addText('What I Need: Banking API credentials (YES Bank, BDO) · ML engineer (rail scoring model) · Payments ops team alignment on rerouting logic', { x:.7, y:6.6, w:8.6, h:.6, fontSize:10, bold:true, color:C.orange });
s8.addText('Ajay Avaghade | avaghadeajay009@gmail.com | +91 9561558439', { x:.5, y:7.15, w:9, h:.25, fontSize:9, color:C.muted, align:'center' });

await prs.writeFile({ fileName: 'assets/botim-clearroute-deck.pptx' });
console.log('✓ botim-clearroute-deck.pptx created');
})();
