(async () => {
const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.defineLayout({ name:'LAYOUT_16x9', width:10, height:7.5 });
prs.layout = 'LAYOUT_16x9';
const C = { dark:'0F0A1E', hero:'1A1030', bright:'6366F1', accent:'818CF8', pink:'EC4899', green:'10B981', orange:'F97316', white:'FFFFFF', lgray:'F8F9FC', muted:'8b92a5' };

// S1 COVER
const s1 = prs.addSlide(); s1.background = { color:C.dark };
s1.addText('PENCIL · EDITOR SUITE · CASE STUDY 01', { x:.3, y:.4, w:9.4, h:.3, fontSize:9, bold:true, color:C.muted, align:'center' });
s1.addText('AdCanvas', { x:.5, y:1.1, w:9, h:1.5, fontSize:54, bold:true, color:C.white, align:'center' });
s1.addText('AI-Powered Format Adaptation — Master Once, Deploy Everywhere', { x:.5, y:2.75, w:9, h:.8, fontSize:20, color:C.accent, align:'center' });
s1.addText('Ajay Avaghade · Product Manager', { x:.5, y:3.75, w:9, h:.4, fontSize:12, color:C.muted, align:'center' });
s1.addShape('rect', { x:7.5, y:5.1, w:2, h:.1, fill:{ color:C.pink }, line:{ type:'none' } });
s1.addText('Saved 4.2 hrs per campaign', { x:6.9, y:5.3, w:2.8, h:.8, fontSize:16, bold:true, color:C.accent, align:'right' });

// S2 PROBLEM
const s2 = prs.addSlide(); s2.background = { color:C.dark };
s2.addText('THE PROBLEM', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s2.addText('60% of Creative Time in the Editor Is Mechanical Adaptation — Not Creative Decisions', { x:.5, y:.8, w:9, h:1.1, fontSize:28, bold:true, color:C.white });
[{ num:'12+', label:'format variants per campaign', sub:'Facebook, Display, Email, Print, Mobile, Leaderboard…' },
 { num:'60%', label:'time on mechanical work', sub:'resize, re-crop, reflow — not creative decisions' },
 { num:'0', label:'AI in the adaptation layer', sub:'AI helps generate; adaptation is still 100% manual' },
].forEach((s,i) => {
  const x=.5+i*3.1;
  s2.addText(s.num, { x, y:2.2, w:2.8, h:.6, fontSize:32, bold:true, color:C.accent, align:'center' });
  s2.addText(s.label, { x, y:2.9, w:2.8, h:.4, fontSize:12, bold:true, color:C.white, align:'center' });
  s2.addText(s.sub, { x, y:3.4, w:2.8, h:.8, fontSize:10, color:C.muted, align:'center' });
});
s2.addShape('rect', { x:.5, y:5.1, w:9, h:1.7, fill:{ color:C.hero }, line:{ color:C.bright, width:1 } });
s2.addText('Root cause: Every format is treated as a new creative. The editor has no concept of "master creative + format derivatives." AI generates; humans adapt — 12 times per campaign. This is the most painful, most avoidable bottleneck in the editor workflow.', { x:.7, y:5.3, w:8.6, h:1.3, fontSize:10, color:C.white });

// S3 INSIGHT
const s3 = prs.addSlide(); s3.background = { color:C.lgray };
s3.addText('THE INSIGHT', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s3.addText('AI Should Absorb Format Complexity — Not the User', { x:.5, y:.8, w:9, h:.9, fontSize:30, bold:true, color:C.dark });
s3.addShape('rect', { x:.5, y:2.0, w:4.2, h:3.8, fill:{ color:'FFFFFF' }, line:{ color:C.muted, width:1 } });
s3.addText('❌ Today: Format-First', { x:.7, y:2.2, w:3.8, h:.4, fontSize:12, bold:true, color:C.orange });
['Create Facebook master creative','Open new canvas — Display 300×250','Manually resize, re-crop, reflow copy','Open new canvas — Email template','Repeat 10 more times','4+ hours. Nothing new was created.'].forEach((t,i) => s3.addText(t, { x:.9, y:2.8+i*.48, w:3.6, h:.44, fontSize:10, color:C.dark }));
s3.addShape('rect', { x:5.3, y:2.0, w:4.2, h:3.8, fill:{ color:C.hero }, line:{ color:C.bright, width:1 } });
s3.addText('✅ AdCanvas: Master-First', { x:5.5, y:2.2, w:3.8, h:.4, fontSize:12, bold:true, color:C.accent });
['Create one master creative','Select 6 format families','AI auto-adapts all variants','Review conflict flags (2 of 6)','Resolve in 2 taps','18 minutes. 6 formats. Done.'].forEach((t,i) => s3.addText(t, { x:5.7, y:2.8+i*.48, w:3.6, h:.44, fontSize:10, color:C.white }));
s3.addShape('ellipse', { x:4.6, y:3.5, w:.8, h:.8, fill:{ color:C.dark }, line:{ color:C.bright, width:2 } });
s3.addText('VS', { x:4.6, y:3.55, w:.8, h:.7, fontSize:14, bold:true, color:C.accent, align:'center' });

// S4 MECHANIC
const s4 = prs.addSlide(); s4.background = { color:C.dark };
s4.addText('THE MECHANIC', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s4.addText('Master → AI Adapts All Variants → User Resolves Conflicts Only', { x:.5, y:.8, w:9, h:.5, fontSize:22, bold:true, color:C.white });
[{ n:'1', t:'Master Creative', d:'User builds one master canvas at highest quality. Brand constraints set once here.' },
 { n:'2', t:'Format Selection', d:'User selects target format families (Display, Email, Social, Print). Dimensions auto-loaded.' },
 { n:'3', t:'AI Adaptation', d:'AI applies crop zone rules, text reflow, hierarchy scaling, brand constraint checks per format.' },
 { n:'4', t:'Conflict Detection', d:'Formats where AI cannot auto-resolve (e.g. CTA clips) are flagged. Others auto-approve.' },
 { n:'5', t:'Conflict Resolution', d:'User sees AI suggestion for each conflict. One tap to apply. Edits propagate to that format only.' },
].forEach((s,i) => {
  const y=1.6+i*.84;
  s4.addShape('ellipse', { x:.5, y, w:.4, h:.4, fill:{ color:C.bright }, line:{ type:'none' } });
  s4.addText(s.n, { x:.5, y, w:.4, h:.4, fontSize:14, bold:true, color:C.white, align:'center' });
  s4.addText(s.t, { x:1.1, y, w:2.3, h:.22, fontSize:10, bold:true, color:C.white });
  s4.addText(s.d, { x:1.1, y:y+.25, w:2.3, h:.28, fontSize:8, color:C.muted });
  if(i<4) s4.addShape('line', { x:.7, y:y+.45, w:0, h:.24, line:{ color:C.bright, width:2, dashType:'dash' } });
});
s4.addShape('rect', { x:3.8, y:1.6, w:5.7, h:4.3, fill:{ color:C.hero }, line:{ color:C.bright, width:1 } });
s4.addText('Proof: PhonePe Offer Operations — Define Once, Render Everywhere', { x:4.0, y:1.8, w:5.3, h:.4, fontSize:10, bold:true, color:C.accent });
['Rebuilt ₹100Cr rewards portfolio with reusable offer card primitives','500+ brand partners launched independently using shared templates','Master offer template → configurable variables → rendered per audience','TAT: 2 days → 30 minutes (same "render everywhere" architecture)','AdCanvas: same "define once, adapt to format" pattern for creatives'].forEach((t,i) => s4.addText(t, { x:4.2, y:2.4+i*.62, w:5.1, h:.56, fontSize:9, color:C.white }));

// S5 PRODUCT
const s5 = prs.addSlide(); s5.background = { color:C.lgray };
s5.addText('THE PRODUCT', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s5.addText('4 Screens — Master to Export', { x:.5, y:.8, w:9, h:.5, fontSize:26, bold:true, color:C.dark });
[{ n:'01', t:'Master Canvas', d:'Hero editor view. Master creative with live format variant panel. Format chips show adapted/conflict status. Conflict count visible at a glance.' },
 { n:'02', t:'Format Variants', d:'All 6 variants shown side-by-side. AI adapts each from master. Conflicts highlighted in amber. Non-conflicting formats auto-approved.' },
 { n:'03', t:'All Formats Ready', d:'Success state: "Saved 4.2 hours." 6 formats adapted, 2 conflicts resolved. Export package ready individually or as a set.' },
 { n:'04', t:'Template Library', d:'Saved master templates with variant count and performance data. CTR of AdCanvas campaigns vs manual adaptation. Insight card per template.' },
].forEach((s,i) => {
  const y=1.6+i*1.35;
  s5.addShape('rect', { x:.5, y, w:9, h:1.2, fill:{ color:'FFFFFF' }, line:{ color:C.muted, width:1 } });
  s5.addText(`${s.n} — ${s.t}`, { x:.7, y:y+.15, w:3.5, h:.3, fontSize:11, bold:true, color:C.dark });
  s5.addText(s.d, { x:.7, y:y+.5, w:8.6, h:.55, fontSize:9, color:C.dark });
});

// S6 ROI
const s6 = prs.addSlide(); s6.background = { color:C.dark };
s6.addText('IMPACT & ROI', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s6.addText('Creative Velocity, Consistency & Campaign Performance', { x:.5, y:.8, w:9, h:.5, fontSize:22, bold:true, color:C.white });
[{ ico:'⏱', v:'18 min', l:'Format adaptation time', s:'vs 4+ hours manual per campaign (93% reduction)' },
 { ico:'🎯', v:'+31%', l:'CTR lift', s:'AdCanvas campaigns vs manually adapted variants' },
 { ico:'✅', v:'94%', l:'Auto-resolve rate', s:'conflicts resolved by AI without user input' },
 { ico:'📐', v:'12× formats', l:'from one master', s:'per campaign — consistent brand application guaranteed' },
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
s7.addText('PhonePe Offer Templates: Define Once, Render Everywhere', { x:.5, y:.8, w:9, h:.6, fontSize:24, bold:true, color:C.dark });
s7.addShape('rect', { x:.5, y:1.6, w:4.3, h:4.4, fill:{ color:C.dark }, line:{ color:C.bright, width:1 } });
s7.addText('PhonePe: What I Built', { x:.7, y:1.8, w:3.9, h:.3, fontSize:11, bold:true, color:C.accent });
['Rebuilt ₹100Cr static offer portfolio into reusable template system','500+ brand partners each created offers using shared primitives','Offer card: define once (headline, image, CTA, audience) → rendered per segment','Operations TAT: 2 days → 30 minutes through template-first architecture','Same creative defined once; rendered to 10+ audience segments automatically'].forEach((t,i) => s7.addText(t, { x:.9, y:2.3+i*.58, w:3.7, h:.52, fontSize:9, color:C.white }));
s7.addShape('rect', { x:5.2, y:1.6, w:4.3, h:4.4, fill:{ color:'FFFFFF' }, line:{ color:C.bright, width:1 } });
s7.addText('AdCanvas: Application', { x:5.4, y:1.8, w:3.9, h:.3, fontSize:11, bold:true, color:C.bright });
['Creative defined once in master canvas','Rendered to 12+ ad format variants automatically','Format constraints handled by AI, not the creator','Conflicts (CTA clips, crop zones) flagged for human resolution only','Same "define once, render everywhere" architecture — different medium'].forEach((t,i) => s7.addText(t, { x:5.4, y:2.3+i*.58, w:3.9, h:.52, fontSize:9, color:C.dark }));

// S8 ROLLOUT
const s8 = prs.addSlide(); s8.background = { color:C.dark };
s8.addText('ROLLOUT PLAN', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s8.addText('Three Phases — Display Pilot to Full Multi-Format Coverage', { x:.5, y:.8, w:9, h:.5, fontSize:22, bold:true, color:C.white });
[{ t:'Phase 1 — Display Pilot (Month 1–2)', b:'Launch AdCanvas for Display formats (300×250, 728×90, 160×600, 320×50). A/B: AdCanvas vs manual adaptation. Measure: adaptation time, conflict rate, user satisfaction score.' },
 { t:'Phase 2 — Social + Email (Month 3–4)', b:'Extend master→variant model to Social (Facebook, Instagram) and Email (hero banner, thumbnail). Add email-specific constraint rules (44px touch targets, mobile reflow).' },
 { t:'Phase 3 — Print + Full Suite (Month 5–6)', b:'Add Print format family (PDF/SVG output). Master creative covers all output families. Template library live with CTR performance benchmarks per saved master.' },
].forEach((p,i) => {
  const y=1.5+i*1.75;
  s8.addShape('rect', { x:.5, y, w:9, h:1.6, fill:{ color:C.hero }, line:{ color:C.bright, width:1 } });
  s8.addText(p.t, { x:.7, y:y+.15, w:8.6, h:.3, fontSize:11, bold:true, color:C.accent });
  s8.addText(p.b, { x:.9, y:y+.52, w:8.2, h:.85, fontSize:9, color:C.white });
});
s8.addShape('rect', { x:.5, y:6.5, w:9, h:.8, fill:{ color:C.hero }, line:{ color:C.orange, width:2 } });
s8.addText('What I Need: Format constraint rules library · AI crop/reflow model · Design team sign-off on master template spec · Eng sprint for conflict-flagging logic', { x:.7, y:6.6, w:8.6, h:.6, fontSize:10, bold:true, color:C.orange });
s8.addText('Ajay Avaghade | avaghadeajay009@gmail.com | +91 9561558439', { x:.5, y:7.15, w:9, h:.25, fontSize:9, color:C.muted, align:'center' });

await prs.writeFile({ fileName:'assets/pencil-adcanvas-deck.pptx' });
console.log('✓ pencil-adcanvas-deck.pptx created');
})();
