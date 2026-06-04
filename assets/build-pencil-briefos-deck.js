(async () => {
const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.defineLayout({ name:'LAYOUT_16x9', width:10, height:7.5 });
prs.layout = 'LAYOUT_16x9';
const C = { dark:'0F0A1E', hero:'1A1030', bright:'EC4899', accent:'F9A8D4', indigo:'6366F1', green:'10B981', orange:'F97316', white:'FFFFFF', lgray:'F8F9FC', muted:'8b92a5' };

// S1 COVER
const s1 = prs.addSlide(); s1.background = { color:C.dark };
s1.addText('PENCIL · EDITOR SUITE · CASE STUDY 02', { x:.3, y:.4, w:9.4, h:.3, fontSize:9, bold:true, color:C.muted, align:'center' });
s1.addText('BriefOS', { x:.5, y:1.1, w:9, h:1.5, fontSize:54, bold:true, color:C.white, align:'center' });
s1.addText('Structured Brief → AI Creative Suite with Transparent Rationale', { x:.5, y:2.75, w:9, h:.8, fontSize:20, color:C.accent, align:'center' });
s1.addText('Ajay Avaghade · Product Manager', { x:.5, y:3.75, w:9, h:.4, fontSize:12, color:C.muted, align:'center' });
s1.addShape('rect', { x:7.5, y:5.1, w:2, h:.1, fill:{ color:C.bright }, line:{ type:'none' } });
s1.addText('−68% regeneration rate', { x:7.0, y:5.3, w:2.8, h:.8, fontSize:16, bold:true, color:C.accent, align:'right' });

// S2 PROBLEM
const s2 = prs.addSlide(); s2.background = { color:C.dark };
s2.addText('THE PROBLEM', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s2.addText('AI Generation Is a Black Box — 80% of Brief Context Is Lost in Translation', { x:.5, y:.8, w:9, h:1.1, fontSize:28, bold:true, color:C.white });
[{ num:'80%', label:'of brief context lost', sub:'audience insight, tone, hierarchy intent all disappear' },
 { num:'6×', label:'avg regeneration cycles', sub:'users prompt, get output, re-prompt — low trust in AI' },
 { num:'40%', label:'abandon to Canva', sub:'users leave Pencil for manual tools after repeated regenerations' },
].forEach((s,i) => {
  const x=.5+i*3.1;
  s2.addText(s.num, { x, y:2.2, w:2.8, h:.6, fontSize:32, bold:true, color:C.accent, align:'center' });
  s2.addText(s.label, { x, y:2.9, w:2.8, h:.4, fontSize:12, bold:true, color:C.white, align:'center' });
  s2.addText(s.sub, { x, y:3.4, w:2.8, h:.8, fontSize:10, color:C.muted, align:'center' });
});
s2.addShape('rect', { x:.5, y:5.1, w:9, h:1.7, fill:{ color:C.hero }, line:{ color:C.bright, width:1 } });
s2.addText('Root cause: "Summer sale 30% off women 25-45" is treated as a search query, not a creative brief. The AI makes decisions about colour, hierarchy, CTA placement — but never explains why. When the output misses, the user has no idea what to change. Regeneration rate climbs. Confidence in AI drops.', { x:.7, y:5.3, w:8.6, h:1.3, fontSize:10, color:C.white });

// S3 INSIGHT
const s3 = prs.addSlide(); s3.background = { color:C.lgray };
s3.addText('THE INSIGHT', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s3.addText("Explaining Why Is More Valuable Than Changing What", { x:.5, y:.8, w:9, h:.9, fontSize:28, bold:true, color:C.dark });
s3.addShape('rect', { x:.5, y:2.0, w:4.2, h:3.8, fill:{ color:'FFFFFF' }, line:{ color:C.muted, width:1 } });
s3.addText('❌ Today: Black Box AI', { x:.7, y:2.2, w:3.8, h:.4, fontSize:12, bold:true, color:C.orange });
['User types a short prompt','AI generates 4 creatives','User: "Why is the background pink?"','No answer. User re-prompts.','6 rounds of regeneration','Still not right. User gives up.'].forEach((t,i) => s3.addText(t, { x:.9, y:2.8+i*.48, w:3.6, h:.44, fontSize:10, color:C.dark }));
s3.addShape('rect', { x:5.3, y:2.0, w:4.2, h:3.8, fill:{ color:C.hero }, line:{ color:C.bright, width:1 } });
s3.addText('✅ BriefOS: Transparent AI', { x:5.5, y:2.2, w:3.8, h:.4, fontSize:12, bold:true, color:C.accent });
['Structured brief captures intent','AI generates with inline rationale cards','"Pink gradient: summer warmth signal"','User changes tone → only colour regenerates','Output is right on first try','-68% regeneration. Trust builds.'].forEach((t,i) => s3.addText(t, { x:5.7, y:2.8+i*.48, w:3.6, h:.44, fontSize:10, color:C.white }));
s3.addShape('ellipse', { x:4.6, y:3.5, w:.8, h:.8, fill:{ color:C.dark }, line:{ color:C.bright, width:2 } });
s3.addText('VS', { x:4.6, y:3.55, w:.8, h:.7, fontSize:14, bold:true, color:C.accent, align:'center' });

// S4 MECHANIC
const s4 = prs.addSlide(); s4.background = { color:C.dark };
s4.addText('THE MECHANIC', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s4.addText('Structured Brief → AI Decisions → Rationale Cards → Edit-a-Decision → Regenerate Only That', { x:.5, y:.8, w:9, h:.5, fontSize:18, bold:true, color:C.white });
[{ n:'1', t:'Structured Brief', d:'Audience, tone, offer, CTA intent, format families, brand constraints. Each field is a signal.' },
 { n:'2', t:'AI Generation', d:'Multi-format creative suite generated. Every AI decision tagged: colour, layout, copy length, CTA position.' },
 { n:'3', t:'Rationale Cards', d:'Each tagged decision has a visible rationale: "CTA bottom-center — thumb zone for mobile feed." User understands why.' },
 { n:'4', t:'Edit a Decision', d:'User taps a rationale card to change it. Sees alternative options with performance context. One tap to apply.' },
 { n:'5', t:'Surgical Regeneration', d:'Only the affected elements regenerate. Colour, copy, and layout that were not changed stay intact.' },
].forEach((s,i) => {
  const y=1.6+i*.84;
  s4.addShape('ellipse', { x:.5, y, w:.4, h:.4, fill:{ color:C.bright }, line:{ type:'none' } });
  s4.addText(s.n, { x:.5, y, w:.4, h:.4, fontSize:14, bold:true, color:C.white, align:'center' });
  s4.addText(s.t, { x:1.1, y, w:2.3, h:.22, fontSize:10, bold:true, color:C.white });
  s4.addText(s.d, { x:1.1, y:y+.25, w:2.3, h:.28, fontSize:8, color:C.muted });
  if(i<4) s4.addShape('line', { x:.7, y:y+.45, w:0, h:.24, line:{ color:C.bright, width:2, dashType:'dash' } });
});
s4.addShape('rect', { x:3.8, y:1.6, w:5.7, h:4.3, fill:{ color:C.hero }, line:{ color:C.bright, width:1 } });
s4.addText('Proof: Antigravity Resume Agent — Ajay Built This Architecture', { x:4.0, y:1.8, w:5.3, h:.4, fontSize:10, bold:true, color:C.accent });
['Built multi-agent AI pipeline: structured JD input → 6 parallel agents','Each agent handles one synthesis task (scraping, layout, ATS scoring, outreach)','Output: multiple resume variants, each with traceable decisions','User edits one section → only that agent re-runs','BriefOS is the same architecture: structured input → AI multivariant → edit a decision'].forEach((t,i) => s4.addText(t, { x:4.2, y:2.4+i*.62, w:5.1, h:.56, fontSize:9, color:C.white }));

// S5 PRODUCT
const s5 = prs.addSlide(); s5.background = { color:C.lgray };
s5.addText('THE PRODUCT', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s5.addText('4 Screens — Brief to Approval', { x:.5, y:.8, w:9, h:.5, fontSize:26, bold:true, color:C.dark });
[{ n:'01', t:'Creative Brief', d:'Structured form: audience, offer, tone tags, CTA intent, format families, brand constraints. Not a search bar — a brief.' },
 { n:'02', t:'AI Creative Suite + Rationale', d:'Multi-format output with inline rationale cards per AI decision. Colour, layout, copy, CTA each explained.' },
 { n:'03', t:'Edit-a-Decision', d:'Tap any rationale card to see alternatives with context. Apply change → surgical regeneration of only affected elements.' },
 { n:'04', t:'Approval Dashboard', d:'All formats with approval status. Brief-to-approval time. Regeneration count. Export package ready.' },
].forEach((s,i) => {
  const y=1.6+i*1.35;
  s5.addShape('rect', { x:.5, y, w:9, h:1.2, fill:{ color:'FFFFFF' }, line:{ color:C.muted, width:1 } });
  s5.addText(`${s.n} — ${s.t}`, { x:.7, y:y+.15, w:3.5, h:.3, fontSize:11, bold:true, color:C.dark });
  s5.addText(s.d, { x:.7, y:y+.5, w:8.6, h:.55, fontSize:9, color:C.dark });
});

// S6 ROI
const s6 = prs.addSlide(); s6.background = { color:C.dark };
s6.addText('IMPACT & ROI', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s6.addText('Trust, Velocity & Creative Quality', { x:.5, y:.8, w:9, h:.5, fontSize:22, bold:true, color:C.white });
[{ ico:'🔄', v:'−68%', l:'Regeneration rate', s:'from 6× avg cycles to 1.9× (user knows what to change)' },
 { ico:'⏱', v:'22 min', l:'Brief to first approval', s:'vs 90+ minutes current (structured brief → AI generates correctly 1st time)' },
 { ico:'❤️', v:'+47%', l:'AI trust score', s:'users who see rationale cards rate AI output higher and abandon less' },
 { ico:'📤', v:'5.2×', l:'Creative output per user', s:'same user produces 5× more approved formats in same time' },
].forEach((m,i) => {
  const x=.5+(i%2)*4.6, y=1.6+Math.floor(i/2)*2;
  s6.addShape('rect', { x, y, w:4, h:1.8, fill:{ color:C.hero }, line:{ color:C.bright, width:1 } });
  s6.addText(m.ico, { x:x+.2, y:y+.1, w:3.6, h:.4, fontSize:22, align:'left' });
  s6.addText(m.v, { x:x+.2, y:y+.52, w:3.6, h:.38, fontSize:20, bold:true, color:C.accent, align:'left' });
  s6.addText(m.l, { x:x+.2, y:y+.96, w:3.6, h:.28, fontSize:10, bold:true, color:C.white, align:'left' });
  s6.addText(m.s, { x:x+.2, y:y+1.28, w:3.6, h:.38, fontSize:8, color:C.muted, align:'left' });
});

// S7 PROOF
const s7 = prs.addSlide(); s7.background = { color:C.lgray };
s7.addText('PROOF OF WORK', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s7.addText('Antigravity Resume Agent: I Built This Architecture', { x:.5, y:.8, w:9, h:.6, fontSize:24, bold:true, color:C.dark });
s7.addShape('rect', { x:.5, y:1.6, w:4.3, h:4.4, fill:{ color:C.dark }, line:{ color:C.bright, width:1 } });
s7.addText('Antigravity: What I Built', { x:.7, y:1.8, w:3.9, h:.3, fontSize:11, bold:true, color:C.accent });
['Multi-agent LLM pipeline: structured JD URL input','6 agents in parallel: scraping, synthesis, layout, ATS scoring, outreach','Structured input → explainable AI output per task','User can edit one section (experience bullet) → only that agent re-runs','90+ ATS-scored resume variants generated from one structured input'].forEach((t,i) => s7.addText(t, { x:.9, y:2.3+i*.58, w:3.7, h:.52, fontSize:9, color:C.white }));
s7.addShape('rect', { x:5.2, y:1.6, w:4.3, h:4.4, fill:{ color:'FFFFFF' }, line:{ color:C.bright, width:1 } });
s7.addText('BriefOS: Application', { x:5.4, y:1.8, w:3.9, h:.3, fontSize:11, bold:true, color:C.bright });
['Structured brief (not a search bar) captures all creative signals','Multi-format creative suite generated with tagged AI decisions','Rationale card per decision = transparency (not a black box)','User edits one rationale → only that element regenerates','Trust in AI builds through transparency, not luck'].forEach((t,i) => s7.addText(t, { x:5.4, y:2.3+i*.58, w:3.9, h:.52, fontSize:9, color:C.dark }));

// S8 ROLLOUT
const s8 = prs.addSlide(); s8.background = { color:C.dark };
s8.addText('ROLLOUT PLAN', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s8.addText('Three Phases — Display Pilot to Full Brief→Suite Workflow', { x:.5, y:.8, w:9, h:.5, fontSize:22, bold:true, color:C.white });
[{ t:'Phase 1 — Brief Form + Rationale Cards (Month 1–2)', b:'Replace freeform text prompt with structured brief form. Add rationale card layer to existing AI output. A/B test: structured brief vs current text prompt. Measure: regeneration rate, time to approval, NPS.' },
 { t:'Phase 2 — Edit-a-Decision (Month 3–4)', b:'Build surgical regeneration: user taps rationale card → selects alternative → only affected elements regenerate. Instrument: which rationale cards are most edited (informs default improvements).' },
 { t:'Phase 3 — Brief Templates + Learning Loop (Month 5–6)', b:'Save brief templates per campaign type (Sale, Product Launch, Brand). AI learns from accepted/rejected rationale cards per user. Defaults improve over time. Regeneration rate target: <2×.' },
].forEach((p,i) => {
  const y=1.5+i*1.75;
  s8.addShape('rect', { x:.5, y, w:9, h:1.6, fill:{ color:C.hero }, line:{ color:C.bright, width:1 } });
  s8.addText(p.t, { x:.7, y:y+.15, w:8.6, h:.3, fontSize:11, bold:true, color:C.accent });
  s8.addText(p.b, { x:.9, y:y+.52, w:8.2, h:.85, fontSize:9, color:C.white });
});
s8.addShape('rect', { x:.5, y:6.5, w:9, h:.8, fill:{ color:C.hero }, line:{ color:C.orange, width:2 } });
s8.addText('What I Need: Brief schema design (with Design) · AI decision tagging layer (with Eng) · Instrumentation for rationale card edits · User research on current regeneration reasons', { x:.7, y:6.6, w:8.6, h:.6, fontSize:10, bold:true, color:C.orange });
s8.addText('Ajay Avaghade | avaghadeajay009@gmail.com | +91 9561558439', { x:.5, y:7.15, w:9, h:.25, fontSize:9, color:C.muted, align:'center' });

await prs.writeFile({ fileName:'assets/pencil-briefos-deck.pptx' });
console.log('✓ pencil-briefos-deck.pptx created');
})();
