(async () => {
const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.defineLayout({ name:'LAYOUT_16x9', width:10, height:7.5 });
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:  '0A1628', hero:  '0F1E3A',
  amber: 'F59E0B', gold:  'FCD34D',
  teal:  '00BFA5', orange:'FF6B35',
  green: '00C853', white: 'FFFFFF',
  lgray: 'F0F4F8', muted: '8b92a5',
};

// S1 COVER
const s1 = prs.addSlide(); s1.background = { color: C.dark };
s1.addText('BOTIM FOR BUSINESS · CORRIDOR EXPANSION STRATEGY', { x:.3, y:.4, w:9.4, h:.3, fontSize:9, bold:true, color:C.muted, align:'center' });
s1.addText('CorridorIQ', { x:.5, y:1.2, w:9, h:1.5, fontSize:52, bold:true, color:C.white, align:'center' });
s1.addText('Data-Driven Corridor Expansion Intelligence — From Gut Feel to Evidence', { x:.5, y:2.8, w:9, h:.8, fontSize:20, color:C.amber, align:'center' });
s1.addText('Ajay Avaghade · Product Manager', { x:.5, y:3.8, w:9, h:.4, fontSize:12, color:C.muted, align:'center' });
s1.addShape('rect', { x:7.5, y:5.2, w:2, h:.1, fill:{ color:C.amber }, line:{ type:'none' } });
s1.addText('6 new corridors in 12 months', { x:6.8, y:5.4, w:2.9, h:.8, fontSize:16, bold:true, color:C.amber, align:'right' });

// S2 PROBLEM
const s2 = prs.addSlide(); s2.background = { color: C.dark };
s2.addText('THE PROBLEM', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s2.addText('Botim Wants 6 New Corridors in 12 Months — But Corridor Decisions Are Made on Gut Feel, Not Data', { x:.5, y:.8, w:9, h:1.2, fontSize:26, bold:true, color:C.white });
[{ num:'6', label:'new corridors planned', sub:'India, Pakistan, Philippines, Africa, UK, Europe' },
 { num:'Ad hoc', label:'current decision process', sub:'country manager memo → committee decision → 3mo delay' },
 { num:'$3.2B', label:'UAE→Philippines annual remittance', sub:'corridor Botim is NOT in yet — missed revenue' },
].forEach((s, i) => {
  const x = .5 + i * 3.1;
  s2.addText(s.num, { x, y:2.3, w:2.8, h:.6, fontSize:28, bold:true, color:C.amber, align:'center' });
  s2.addText(s.label, { x, y:3.0, w:2.8, h:.4, fontSize:12, bold:true, color:C.white, align:'center' });
  s2.addText(s.sub, { x, y:3.5, w:2.8, h:.8, fontSize:10, color:C.muted, align:'center' });
});
s2.addShape('rect', { x:.5, y:5.2, w:9, h:1.7, fill:{ color:C.hero }, line:{ color:C.amber, width:1 } });
s2.addText('Root cause: No standardised framework for evaluating corridor opportunity. Each corridor decision requires weeks of manual research: regulatory complexity, SME demand signals, banking partnerships, FX benchmarks, competitive saturation. Result: wrong corridors prioritised, delayed launches, underperforming volumes.', { x:.7, y:5.35, w:8.6, h:1.4, fontSize:10, color:C.white });

// S3 INSIGHT
const s3 = prs.addSlide(); s3.background = { color: C.lgray };
s3.addText('THE INSIGHT', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s3.addText("Corridor Expansion Is a Data Problem, Not a Geography Problem", { x:.5, y:.8, w:9, h:1.0, fontSize:28, bold:true, color:C.dark });
s3.addShape('rect', { x:.5, y:2.0, w:4.2, h:3.8, fill:{ color:'FFFFFF' }, line:{ color:C.muted, width:1 } });
s3.addText('❌ Today: Ad-Hoc Decisions', { x:.7, y:2.2, w:3.8, h:.4, fontSize:12, bold:true, color:C.orange });
['Country manager writes proposal','Manual research (2–4 weeks per corridor)','Committee vote based on intuition','Wrong corridor prioritised (high cost, low vol)','Philippines missed while India over-indexed'].forEach((t,i) => s3.addText(t, { x:.9, y:2.8+i*.58, w:3.6, h:.5, fontSize:10, color:C.dark }));
s3.addShape('rect', { x:5.3, y:2.0, w:4.2, h:3.8, fill:{ color:C.hero }, line:{ color:C.amber, width:1 } });
s3.addText('✅ CorridorIQ: Evidence-Based', { x:5.5, y:2.2, w:3.8, h:.4, fontSize:12, bold:true, color:C.amber });
['4 scored dimensions per corridor','Demand signal, regulatory, banking, competition','Ranked output: top corridors by 12-mo volume potential','Launch plan auto-generated with go/no-go gates','Philippines scores 87/100 → launch Q3, not Q1 2025'].forEach((t,i) => s3.addText(t, { x:5.7, y:2.8+i*.58, w:3.6, h:.5, fontSize:10, color:C.white }));
s3.addShape('ellipse', { x:4.6, y:3.5, w:.8, h:.8, fill:{ color:C.dark }, line:{ color:C.amber, width:2 } });
s3.addText('VS', { x:4.6, y:3.55, w:.8, h:.7, fontSize:14, bold:true, color:C.amber, align:'center' });

// S4 MECHANIC
const s4 = prs.addSlide(); s4.background = { color: C.dark };
s4.addText('THE MECHANIC', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s4.addText('Demand Signals + Regulatory + Banking + Competition → Ranked Corridor Score → Launch Plan', { x:.5, y:.8, w:9, h:.5, fontSize:19, bold:true, color:C.white });
[{ n:'1', t:'Signal Aggregation', d:'OFW remittance data, SME transaction requests, expat nationality split from Botim user base' },
 { n:'2', t:'Regulatory Scoring', d:'Central bank licensing complexity, FATF status, estimated approval timeline per target market' },
 { n:'3', t:'Banking Assessment', d:'Available partner banks, API readiness, FX spread benchmarks vs Wise/Paysend/Western Union' },
 { n:'4', t:'Corridor Score', d:'Weighted 0–100. High score = high demand + low regulation + banking ready. Top 3 corridors ranked.' },
 { n:'5', t:'Launch Plan Auto-Gen', d:'Phase checklist, go/no-go gates, KPI targets, cross-sell hooks (WPS payroll, acquiring) per corridor' },
].forEach((s, i) => {
  const y = 1.6 + i * .84;
  s4.addShape('ellipse', { x:.5, y, w:.4, h:.4, fill:{ color:C.amber }, line:{ type:'none' } });
  s4.addText(s.n, { x:.5, y, w:.4, h:.4, fontSize:14, bold:true, color:C.dark, align:'center' });
  s4.addText(s.t, { x:1.1, y, w:2.3, h:.22, fontSize:10, bold:true, color:C.white });
  s4.addText(s.d, { x:1.1, y:y+.25, w:2.3, h:.28, fontSize:8, color:C.muted });
  if (i < 4) s4.addShape('line', { x:.7, y:y+.45, w:0, h:.24, line:{ color:C.amber, width:2, dashType:'dash' } });
});
s4.addShape('rect', { x:3.8, y:1.6, w:5.7, h:4.3, fill:{ color:C.hero }, line:{ color:C.amber, width:1 } });
s4.addText('Proof: PhonePe Merchant Platform (1 city → 25 cities, 12 months)', { x:4.0, y:1.8, w:5.3, h:.4, fontSize:10, bold:true, color:C.amber });
['Scaled merchant GTM from Bangalore to 25+ Indian cities','Each city had different merchant profiles → different GTM','Bangalore = API-first startups; Tier-2 = phone support + training','Hand-crafted playbook per city, then standardised and replicated','CorridorIQ automates the "what does this city/corridor need?" question'].forEach((t,i) => s4.addText(t, { x:4.2, y:2.4+i*.62, w:5.1, h:.56, fontSize:9, color:C.white }));

// S5 PRODUCT
const s5 = prs.addSlide(); s5.background = { color: C.lgray };
s5.addText('THE PRODUCT', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s5.addText('4 Screens — Dashboard to Post-Launch Intelligence', { x:.5, y:.8, w:9, h:.5, fontSize:24, bold:true, color:C.dark });
[{ n:'01', t:'Corridor Dashboard', d:'Live corridors (vol, success rate, TAT) + pipeline corridors (score, est. 12mo vol, partner bank). One view across all geographies.' },
 { n:'02', t:'Corridor Analysis', d:'4-dimension scoring (demand, regulatory, banking, competition). Market intelligence: remittance data, competitor pricing, regulatory timeline.' },
 { n:'03', t:'Launch Plan', d:'Auto-generated 4-phase plan: regulatory, integration, soft launch, full launch. Each phase has go/no-go gates. Cross-sell hooks per corridor.' },
 { n:'04', t:'Performance View', d:'Post-launch KPIs: volume, success rate, cost per txn, TAT. Corridor health summary. Insight card: unit economics vs target.' },
].forEach((s, i) => {
  const y = 1.6 + i * 1.35;
  s5.addShape('rect', { x:.5, y, w:9, h:1.2, fill:{ color:'FFFFFF' }, line:{ color:C.muted, width:1 } });
  s5.addText(`${s.n} — ${s.t}`, { x:.7, y:y+.15, w:3.5, h:.3, fontSize:11, bold:true, color:C.dark });
  s5.addText(s.d, { x:.7, y:y+.5, w:8.6, h:.55, fontSize:9, color:C.dark });
});

// S6 ROI
const s6 = prs.addSlide(); s6.background = { color: C.dark };
s6.addText('IMPACT & ROI', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s6.addText('Corridor Velocity, Revenue & Decision Quality', { x:.5, y:.8, w:9, h:.5, fontSize:22, bold:true, color:C.white });
[{ ico:'🚀', v:'6 corridors', l:'12-month launch velocity', s:'vs current 2 corridors/yr (3× faster expansion)' },
 { ico:'💰', v:'AED 2.1M', l:'Philippines projected vol/mo', s:'corridor currently missed — available Q3 2024' },
 { ico:'📊', v:'2 weeks', l:'Corridor research time', s:'vs 4–6 weeks manual (75% faster decision cycle)' },
 { ico:'🎯', v:'87/100', l:'Philippines corridor score', s:'top-ranked next launch — data confirms intuition' },
].forEach((m, i) => {
  const x = .5 + (i%2)*4.6, y = 1.6 + Math.floor(i/2)*2;
  s6.addShape('rect', { x, y, w:4, h:1.8, fill:{ color:C.hero }, line:{ color:C.amber, width:1 } });
  s6.addText(m.ico, { x:x+.2, y:y+.1, w:3.6, h:.4, fontSize:22, align:'left' });
  s6.addText(m.v, { x:x+.2, y:y+.52, w:3.6, h:.38, fontSize:18, bold:true, color:C.amber, align:'left' });
  s6.addText(m.l, { x:x+.2, y:y+.96, w:3.6, h:.28, fontSize:10, bold:true, color:C.white, align:'left' });
  s6.addText(m.s, { x:x+.2, y:y+1.28, w:3.6, h:.38, fontSize:8, color:C.muted, align:'left' });
});

// S7 PROOF
const s7 = prs.addSlide(); s7.background = { color: C.lgray };
s7.addText('PROOF OF WORK', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s7.addText('PhonePe Merchant Platform: 1 City to 25 in 12 Months', { x:.5, y:.8, w:9, h:.6, fontSize:24, bold:true, color:C.dark });
s7.addShape('rect', { x:.5, y:1.6, w:4.3, h:4.4, fill:{ color:C.dark }, line:{ color:C.amber, width:1 } });
s7.addText('PhonePe: What I Built', { x:.7, y:1.8, w:3.9, h:.3, fontSize:11, bold:true, color:C.amber });
['Scaled self-serve merchant platform 1 → 25+ cities in 12 months','Each city: different merchant type, different GTM required','Bangalore: API docs + self-serve; Tier-2: phone support + training','Built standardised city playbook, replicated with local tweaks','5K merchants, −23% CAC — by optimising GTM per segment'].forEach((t,i) => s7.addText(t, { x:.9, y:2.3+i*.58, w:3.7, h:.52, fontSize:9, color:C.white }));
s7.addShape('rect', { x:5.2, y:1.6, w:4.3, h:4.4, fill:{ color:'FFFFFF' }, line:{ color:C.amber, width:1 } });
s7.addText('CorridorIQ: Application', { x:5.4, y:1.8, w:3.9, h:.3, fontSize:11, bold:true, color:C.amber });
['Same "where do we go next, and in what order?" problem','Each corridor: different regulatory, banking, demand profile','CorridorIQ scores 4 dimensions → ranked priority list','Auto-generates phased launch plan with go/no-go gates','Flywheel: post-launch KPIs train the scoring model'].forEach((t,i) => s7.addText(t, { x:5.4, y:2.3+i*.58, w:3.9, h:.52, fontSize:9, color:C.dark }));

// S8 ROLLOUT
const s8 = prs.addSlide(); s8.background = { color: C.dark };
s8.addText('ROLLOUT PLAN', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s8.addText('Three Phases — Philippines Pilot to Full 6-Corridor Coverage', { x:.5, y:.8, w:9, h:.5, fontSize:22, bold:true, color:C.white });
[{ t:'Phase 1 — Philippines Analysis & Launch (Month 1–3)', b:'CorridorIQ confirms Philippines as top-ranked next corridor (score: 87/100). Execute BDO banking agreement + BSP regulatory filing. Soft launch with go/no-go gate: >96% success, <4h TAT.' },
 { t:'Phase 2 — Africa + UK Scoring (Month 3–4)', b:'Run CorridorIQ analysis on Africa (Kenya M-Pesa, Nigeria NIBSS) and UK (FPS). Rank by combined score. Begin regulatory engagement for highest-ranked. A/B: CorridorIQ vs manual analysis team.' },
 { t:'Phase 3 — Continuous Corridor Intelligence (Month 5–6)', b:'CorridorIQ standard for all future corridor decisions. Post-launch KPIs feed back into corridor scoring model. Cross-sell hooks (WPS payroll, SME lending) activated per corridor at full launch.' },
].forEach((p, i) => {
  const y = 1.5 + i * 1.75;
  s8.addShape('rect', { x:.5, y, w:9, h:1.6, fill:{ color:C.hero }, line:{ color:C.amber, width:1 } });
  s8.addText(p.t, { x:.7, y:y+.15, w:8.6, h:.3, fontSize:11, bold:true, color:C.amber });
  s8.addText(p.b, { x:.9, y:y+.52, w:8.2, h:.85, fontSize:9, color:C.white });
});
s8.addShape('rect', { x:.5, y:6.5, w:9, h:.8, fill:{ color:C.hero }, line:{ color:C.orange, width:2 } });
s8.addText('What I Need: Remittance data access (OFW volumes by corridor) · Regulatory contacts per market · Banking partnership pipeline · 1 data engineer (scoring model)', { x:.7, y:6.6, w:8.6, h:.6, fontSize:10, bold:true, color:C.orange });
s8.addText('Ajay Avaghade | avaghadeajay009@gmail.com | +91 9561558439', { x:.5, y:7.15, w:9, h:.25, fontSize:9, color:C.muted, align:'center' });

await prs.writeFile({ fileName: 'assets/botim-corridoriq-deck.pptx' });
console.log('✓ botim-corridoriq-deck.pptx created');
})();
